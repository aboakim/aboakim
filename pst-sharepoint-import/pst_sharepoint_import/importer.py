"""Walk a PST and create one SharePoint list item per message."""

from __future__ import annotations

import logging
import os
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

import requests

from .auth import GraphAuthenticator
from .config import ConfigError, Settings
from .pst import EmailRecord, ReadFailure, iter_records
from .sharepoint import SharePointListClient
from .state import ImportState

logger = logging.getLogger("pst_sharepoint_import")


@dataclass
class ImportSummary:
    seen: int = 0
    imported: int = 0
    skipped_duplicates: int = 0
    failed: int = 0
    dry_run: int = 0
    read_errors: int = 0

    def exit_code(self) -> int:
        if self.failed or self.read_errors:
            return 1
        return 0


def run_import(settings: Settings, *, records=None, client=None, state: ImportState | None = None) -> ImportSummary:
    summary = ImportSummary()
    lock_path = Path(settings.state_file).with_suffix(".lock")
    graph = client
    owns_client = False
    with import_lock(lock_path):
        if state is None:
            state = ImportState(Path(settings.state_file))
            state.load()
        if graph is None and settings.needs_graph():
            graph = build_client(settings)
            owns_client = True
        try:
            summary = _import_records(settings, summary, graph, state, records)
        finally:
            if owns_client and graph is not None:
                graph.session.close()
    return summary


def _import_records(settings, summary, graph, state, records):
    if settings.ensure_columns:
        if graph is None:
            raise ConfigError("ensure_columns needs SharePoint site_id, list_id, and auth.")
        graph.ensure_columns()
    if graph is not None and not settings.dry_run:
        missing = graph.missing_columns()
        if missing:
            raise ConfigError(
                "These columns are not on the SharePoint list: "
                + ", ".join(missing)
                + ". Run with --ensure-columns, or clear those field_map entries."
            )
    if settings.sync_existing:
        if graph is None:
            raise ConfigError("sync_existing needs SharePoint site_id, list_id, and auth.")
        column_name = settings.field_map["email_id"]
        logger.info("Loading existing %s values from SharePoint", column_name)
        remote_ids = 0
        for email_id in graph.iter_existing_email_ids(column_name):
            if state.add(email_id):
                remote_ids += 1
        logger.info("Recorded %s ids already on the list", remote_ids)

    if records is None:
        source = iter_records(settings.pst_file, settings.folder_prefixes)
    else:
        source = iter(records)

    verbose = settings.log_each_message or (settings.limit is not None and settings.limit <= 100)
    for item in source:
        if settings.limit is not None and summary.seen >= settings.limit:
            break
        summary.seen += 1
        if isinstance(item, ReadFailure):
            summary.read_errors += 1
            logger.error("Read error in %s message %s: %s", item.folder, item.index, item.error)
            continue
        if not isinstance(item, EmailRecord):
            summary.read_errors += 1
            logger.error("Unexpected record type: %s", type(item).__name__)
            continue
        if item.email_id in state:
            summary.skipped_duplicates += 1
            if verbose:
                logger.info("Skip duplicate %s | %s", item.date, item.subject)
            continue
        if settings.dry_run:
            summary.dry_run += 1
            if verbose:
                logger.info("Dry run %s | %s | %s", item.date, item.folder, item.subject)
        else:
            try:
                if graph is None:
                    raise ConfigError("SharePoint client is not configured.")
                graph.create_item(item)
            except Exception as exc:
                summary.failed += 1
                logger.error("Import failed in %s for %s: %s", item.folder, item.subject, exc)
            else:
                state.add(item.email_id)
                summary.imported += 1
                if verbose:
                    logger.info("Imported %s | %s", item.date, item.subject)
                if settings.request_delay_seconds > 0:
                    time.sleep(settings.request_delay_seconds)
        if summary.seen % 100 == 0:
            logger.info(
                "Progress seen=%s imported=%s skipped=%s failed=%s read_errors=%s",
                summary.seen,
                summary.imported,
                summary.skipped_duplicates,
                summary.failed,
                summary.read_errors,
            )
    return summary


def build_client(settings: Settings) -> SharePointListClient:
    session = requests.Session()
    session.headers["User-Agent"] = "aetechron-pst-import/0.1"
    return SharePointListClient(settings, session, GraphAuthenticator(settings))


@contextmanager
def import_lock(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        raise ConfigError(
            f"Another import appears to be running ({path}). "
            "If that process is gone, delete the lock file and run again."
        )
    path.write_text(str(os.getpid()), encoding="utf-8")
    try:
        yield
    finally:
        try:
            path.unlink()
        except FileNotFoundError:
            pass
