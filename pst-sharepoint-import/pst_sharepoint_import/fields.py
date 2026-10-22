"""Map an extracted message onto SharePoint list columns."""

from __future__ import annotations

from typing import TYPE_CHECKING

from .textutil import limit_text

if TYPE_CHECKING:
    from .config import Settings
    from .pst import EmailRecord

# Logical name -> default SharePoint internal name.
# Title already exists on every list. Other names are created by --ensure-columns.
DEFAULT_FIELD_MAP: dict[str, str] = {
    "subject": "Title",
    "email_id": "EmailID",
    "internet_message_id": "InternetMessageId",
    "pst_identifier": "PstIdentifier",
    "date": "EmailDate",
    "sender": "Sender",
    "sender_name": "SenderName",
    "recipients": "Recipients",
    "folder": "Folder",
    "body": "EmailBody",
    "body_source": "BodySource",
    "attachments": "AttachmentNames",
    "has_attachments": "HasAttachments",
    "source_file": "SourceFile",
    "technical_relevant": "TechnicalRelevant",
    "technical_score": "TechnicalScore",
    "customer": "Customer",
    "product": "Product",
    "application": "Application",
    "problem": "Problem",
    "resolution": "Resolution",
    "keywords": "Keywords",
    "ai_processed": "AIProcessed",
}

SINGLE_LINE_FIELDS = {
    "subject",
    "email_id",
    "internet_message_id",
    "pst_identifier",
    "sender",
    "sender_name",
    "body_source",
    "source_file",
    "customer",
    "product",
    "application",
}

AI_DEFAULTS: dict[str, object] = {
    "technical_relevant": False,
    "technical_score": 0,
    "customer": "",
    "product": "",
    "application": "",
    "problem": "",
    "resolution": "",
    "keywords": "",
    "ai_processed": False,
}

COLUMN_DESCRIPTIONS: dict[str, str] = {
    "email_id": "Stable hash used to skip messages already imported from the PST.",
    "internet_message_id": "RFC Message-ID from the message headers, when the PST has one.",
    "pst_identifier": "libpff item identifier inside the source PST.",
    "date": "Client submit time, or delivery time when submit time is missing.",
    "sender": "Sender SMTP address from the transport headers.",
    "sender_name": "Sender display name.",
    "recipients": "To, Cc, and Bcc display values.",
    "folder": "Folder path inside the PST.",
    "body": "Plain-text body kept for later classification. HTML is converted to text.",
    "body_source": "Which body was stored: plain, html, rtf, or empty.",
    "attachments": "Attachment file names. File bytes are not uploaded.",
    "has_attachments": "True when the message has one or more attachments.",
    "source_file": "PST file name this row was read from.",
    "technical_relevant": "Reserved for later AI classification.",
    "technical_score": "Reserved for later AI classification.",
    "customer": "Reserved for later AI classification.",
    "product": "Reserved for later AI classification.",
    "application": "Reserved for later AI classification.",
    "problem": "Reserved for later AI classification.",
    "resolution": "Reserved for later AI classification.",
    "keywords": "Reserved for later AI classification.",
    "ai_processed": "Set true by a later classification job. This import leaves it false.",
}


def column_definition(logical_name: str, sharepoint_name: str) -> dict | None:
    """Graph column create body. Title already exists, so subject returns None."""
    if not sharepoint_name or sharepoint_name.lower() == "title":
        return None

    definition: dict = {
        "name": sharepoint_name,
        "displayName": sharepoint_name,
        "description": COLUMN_DESCRIPTIONS.get(logical_name, ""),
    }
    if logical_name == "date":
        definition["dateTime"] = {"displayAs": "default", "format": "dateTime"}
    elif logical_name in ("has_attachments", "technical_relevant", "ai_processed"):
        definition["boolean"] = {}
    elif logical_name == "technical_score":
        definition["number"] = {"decimalPlaces": "none", "displayAs": "number"}
    elif logical_name in ("body", "recipients", "folder", "attachments", "problem", "resolution", "keywords"):
        definition["text"] = {
            "allowMultipleLines": True,
            "appendChangesToExistingText": False,
            "linesForEditing": 6,
            "textType": "plain",
        }
    else:
        definition["text"] = {"allowMultipleLines": False, "maxLength": 255}
    if logical_name == "email_id":
        definition["indexed"] = True
    return definition


def build_fields(record: EmailRecord, settings: Settings) -> dict[str, object]:
    """Build the Graph `fields` object for one list item."""
    limits = settings.limits
    subject = limit_text(record.subject, limits.subject, single_line=True) or "(no subject)"
    values: dict[str, object] = {
        "subject": subject,
        "email_id": limit_text(record.email_id, limits.email_id, single_line=True),
        "internet_message_id": limit_text(
            record.internet_message_id,
            limits.internet_message_id,
            single_line=True,
        ),
        "pst_identifier": limit_text(record.pst_identifier, limits.pst_identifier, single_line=True),
        "date": record.date,
        "sender": limit_text(record.sender, limits.sender, single_line=True),
        "sender_name": limit_text(record.sender_name, limits.sender_name, single_line=True),
        "recipients": limit_text(record.recipients, limits.recipients, single_line=False),
        "folder": limit_text(record.folder, limits.folder, single_line=False),
        "body": limit_text(record.body, limits.body, single_line=False),
        "body_source": limit_text(record.body_source, limits.body_source, single_line=True),
        "attachments": limit_text(record.attachments, limits.attachments, single_line=False),
        "has_attachments": bool(record.has_attachments),
        "source_file": limit_text(record.source_file, limits.source_file, single_line=True),
        **AI_DEFAULTS,
    }

    fields: dict[str, object] = {}
    for logical_name, column_name in settings.field_map.items():
        if not column_name or logical_name not in values:
            continue
        value = values[logical_name]
        if logical_name == "date" and not value:
            continue
        fields[column_name] = value
    return fields
