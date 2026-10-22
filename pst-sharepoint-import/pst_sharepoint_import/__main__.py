"""Command line entry: python -m pst_sharepoint_import --config config.json"""

from __future__ import annotations

import argparse
import sys

from .config import ConfigError, load_settings
from .importer import run_import
from .logging_setup import configure_logging
from .sharepoint import GraphError


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Import an Outlook PST archive into a SharePoint list."
    )
    parser.add_argument("--config", required=True, help="Path to config.json")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Read the PST and log rows without creating SharePoint items",
    )
    parser.add_argument("--limit", type=int, help="Stop after this many messages")
    parser.add_argument(
        "--sync-existing",
        action="store_true",
        help="Page the list and skip EmailID values already stored there",
    )
    parser.add_argument(
        "--no-sync-existing",
        action="store_true",
        help="Do not read the list first. A local state file is still used.",
    )
    parser.add_argument(
        "--ensure-columns",
        action="store_true",
        help="Create any mapped columns that are missing from the list",
    )
    parser.add_argument("--log-level", help="DEBUG, INFO, WARNING, or ERROR")
    args = parser.parse_args(argv)
    if args.sync_existing and args.no_sync_existing:
        parser.error("Use only one of --sync-existing and --no-sync-existing.")
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        settings = load_settings(args.config)
    except ConfigError as exc:
        print(f"Configuration error: {exc}", file=sys.stderr)
        return 2
    if args.dry_run:
        settings.dry_run = True
    if args.limit is not None:
        settings.limit = args.limit
    if args.sync_existing:
        settings.sync_existing = True
    if args.no_sync_existing:
        settings.sync_existing = False
    if args.ensure_columns:
        settings.ensure_columns = True
    if args.log_level:
        settings.log_level = args.log_level

    logger = configure_logging(settings.log_dir, settings.log_level)
    try:
        settings.validate()
        logger.info("PST file: %s", settings.pst_file)
        if settings.dry_run:
            logger.info("Dry run. SharePoint items will not be created.")
        summary = run_import(settings)
    except ConfigError as exc:
        logger.error("Configuration error: %s", exc)
        return 2
    except GraphError as exc:
        logger.error("%s", exc)
        return 1
    logger.info(
        "Finished seen=%s imported=%s skipped_duplicates=%s failed=%s read_errors=%s dry_run=%s",
        summary.seen,
        summary.imported,
        summary.skipped_duplicates,
        summary.failed,
        summary.read_errors,
        summary.dry_run,
    )
    return summary.exit_code()


if __name__ == "__main__":
    raise SystemExit(main())
