"""Read messages from a PST file through pypff / libpff."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from .config import ConfigError
from .textutil import (
    create_email_id,
    format_recipients,
    html_to_text,
    parse_transport_headers,
    safe_decode,
    split_mailbox,
)

logger = logging.getLogger("pst_sharepoint_import")

# MAPI attachment file name properties.
ATTACH_LONG_FILENAME = 0x3707
ATTACH_FILENAME = 0x3704


@dataclass
class EmailRecord:
    email_id: str
    subject: str
    sender: str
    sender_name: str
    recipients: str
    date: str
    folder: str
    body: str
    body_source: str
    attachments: str
    has_attachments: bool
    internet_message_id: str
    pst_identifier: str
    source_file: str


@dataclass
class ReadFailure:
    folder: str
    index: int
    error: str


def iter_records(pst_file: str, folder_prefixes: list[str] | None = None):
    """Yield EmailRecord or ReadFailure for every message in the PST."""
    pypff = _load_pypff()
    source_file = Path(pst_file).name
    pst = pypff.file()
    pst.open(pst_file)
    try:
        root = pst.get_root_folder()
        yield from _walk_folder(root, "", source_file, folder_prefixes or [])
    finally:
        pst.close()


def _load_pypff():
    try:
        import pypff
    except ImportError as exc:
        raise ConfigError(
            "pypff is not installed. Install libpff and the pypff Python bindings "
            "on the machine that holds the PST, then rerun. See the README."
        ) from exc
    return pypff


def _walk_folder(folder, parent_path: str, source_file: str, folder_prefixes: list[str]):
    folder_name = safe_decode(_read_value(folder, "name")).strip() or "Root"
    folder_path = f"{parent_path}/{folder_name}" if parent_path else folder_name
    selected = folder_selected(folder_path, folder_prefixes)
    message_count = _read_count(folder, "number_of_sub_messages")
    if selected:
        logger.info("Reading %s (%s messages)", folder_path, message_count)
    for index in range(message_count):
        if not selected:
            continue
        try:
            message = folder.get_sub_message(index)
            yield extract_record(message, folder_path, source_file)
        except Exception as exc:
            yield ReadFailure(folder=folder_path, index=index, error=str(exc))
    child_count = _read_count(folder, "number_of_sub_folders")
    for index in range(child_count):
        try:
            child = folder.get_sub_folder(index)
        except Exception as exc:
            logger.error("Could not open subfolder %s of %s: %s", index, folder_path, exc)
            continue
        yield from _walk_folder(child, folder_path, source_file, folder_prefixes)


def folder_selected(folder_path: str, prefixes: list[str]) -> bool:
    """True when prefixes is empty or any prefix matches a folder path segment."""
    cleaned = [prefix.replace("\\", "/").strip("/") for prefix in prefixes if prefix and prefix.strip()]
    if not cleaned:
        return True
    normalized = folder_path.replace("\\", "/").strip("/").lower()
    for prefix in cleaned:
        pref = prefix.lower()
        if normalized == pref or normalized.startswith(pref + "/"):
            return True
        if normalized.endswith("/" + pref) or ("/" + pref + "/") in ("/" + normalized + "/"):
            return True
    return False


def extract_record(message, folder_path: str, source_file: str) -> EmailRecord:
    """Turn one pypff message into a record. Property access is isolated per field."""
    headers = parse_transport_headers(_read_value(message, "transport_headers"))
    header_name, header_address = split_mailbox(headers.get("from", ""))
    sender_name = header_name or safe_decode(_read_value(message, "sender_name")).strip()
    sender = header_address
    subject = safe_decode(_read_value(message, "subject")).strip()
    if not subject:
        subject = safe_decode(_read_value(message, "conversation_topic")).strip()
    body, body_source = _read_body(message)
    attachment_list = _attachment_names(message)
    submitted = _read_value(message, "client_submit_time") or _read_value(message, "delivery_time")
    if not submitted:
        submitted = _read_value(message, "creation_time")
    date = format_graph_datetime(submitted)
    pst_identifier = safe_decode(_read_value(message, "identifier") or _read_value(message, "get_identifier")).strip()
    internet_message_id = headers.get("message_id", "")
    recipients = format_recipients(headers.get("to", ""), headers.get("cc", ""), headers.get("bcc", ""))
    email_id = create_email_id(
        source_file=source_file,
        folder=folder_path,
        subject=subject,
        sender=sender,
        date=date,
        internet_message_id=internet_message_id,
        pst_identifier=pst_identifier,
        body=body,
    )
    return EmailRecord(
        email_id=email_id,
        subject=subject,
        sender=sender,
        sender_name=sender_name,
        recipients=recipients,
        date=date,
        folder=folder_path,
        body=body,
        body_source=body_source,
        attachments="; ".join(attachment_list),
        has_attachments=bool(attachment_list),
        internet_message_id=internet_message_id,
        pst_identifier=pst_identifier,
        source_file=source_file,
    )


def format_graph_datetime(value: object) -> str:
    """Format a PST timestamp as UTC ISO-8601. Empty when the message has no date."""
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.strftime("%Y-%m-%dT%H:%M:%SZ")
        return value.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return safe_decode(value).strip()


def _read_body(message) -> tuple[str, str]:
    plain = safe_decode(_read_value(message, "plain_text_body")).strip()
    if plain:
        return plain, "plain"
    html_body = _read_value(message, "html_body")
    if safe_decode(html_body).strip():
        text = html_to_text(html_body).strip()
        if text:
            return text, "html"
    rtf_body = safe_decode(_read_value(message, "rtf_body")).strip()
    if rtf_body:
        return "[RTF body was present and was not converted to text]", "rtf"
    return "", "empty"


def _attachment_names(message) -> list[str]:
    count = _read_count(message, "number_of_attachments")
    names: list[str] = []
    for index in range(count):
        try:
            attachment = message.get_attachment(index)
        except Exception as exc:
            logger.warning("Unreadable attachment %s: %s", index, exc)
            names.append(f"(unreadable-{index})")
            continue
        name = _attachment_name(attachment).strip() or f"(unnamed-{index})"
        names.append(name)
    return names


def _attachment_name(attachment) -> str:
    for attribute in ("name", "get_name"):
        value = safe_decode(_read_value(attachment, attribute)).strip()
        if value:
            return value
    return _name_from_record_sets(attachment)


def _name_from_record_sets(attachment) -> str:
    """Read PR_ATTACH_LONG_FILENAME (0x3707), then PR_ATTACH_FILENAME (0x3704)."""
    count = _read_count(attachment, "number_of_record_sets")
    getter = getattr(attachment, "get_record_set", None)
    if not callable(getter):
        return ""
    found: list[tuple[int, str]] = []
    for index in range(count):
        try:
            record_set = getter(index)
        except Exception:
            continue
        entry_count = _read_count(record_set, "number_of_entries")
        get_entry = getattr(record_set, "get_entry", None)
        if not callable(get_entry):
            continue
        for entry_index in range(entry_count):
            try:
                entry = get_entry(entry_index)
            except Exception:
                continue
            entry_type = _read_value(entry, "get_entry_type")
            try:
                entry_type_int = int(entry_type) if entry_type is not None else None
            except (TypeError, ValueError):
                entry_type_int = None
            if entry_type_int not in (ATTACH_LONG_FILENAME, ATTACH_FILENAME):
                continue
            data = safe_decode(_read_value(entry, "get_data_as_string")).strip()
            if data:
                rank = 0 if entry_type_int == ATTACH_LONG_FILENAME else 1
                found.append((rank, data))
    if not found:
        return ""
    found.sort()
    return found[0][1]


def _read_value(obj, name: str):
    """Read a pypff attribute or no-arg method. Missing and failing fields return None."""
    try:
        value = getattr(obj, name)
    except Exception:
        return None
    if callable(value):
        try:
            return value()
        except Exception:
            return None
    return value


def _read_count(obj, name: str) -> int:
    value = _read_value(obj, name)
    try:
        return max(0, int(value))
    except (TypeError, ValueError):
        return 0
