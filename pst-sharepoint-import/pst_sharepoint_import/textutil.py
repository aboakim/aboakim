"""Text helpers for PST values, HTML bodies, and SharePoint field limits."""

from __future__ import annotations

import hashlib
import html
import re
from email import message_from_string
from email.utils import getaddresses, parseaddr
from html.parser import HTMLParser


def safe_decode(value: object) -> str:
    """Convert a PST value to text. Bytes are UTF-8 with replacement characters."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        text = value.decode("utf-8", errors="replace")
    else:
        text = str(value)
    return text.replace("\x00", "")


def clean_text(value: str, *, keep_newlines: bool) -> str:
    """Drop control characters SharePoint rejects. Keep newlines in bodies."""
    kept: list[str] = []
    for char in value:
        if char in ("\r", "\n", "\t"):
            if keep_newlines and char == "\n":
                kept.append("\n")
            elif keep_newlines and char == "\t":
                kept.append("\t")
            elif not keep_newlines:
                kept.append(" ")
            continue
        if ord(char) < 32:
            continue
        kept.append(char)
    return "".join(kept)


def one_line(value: str) -> str:
    return " ".join(value.split())


def limit_text(value: str, limit: int, *, single_line: bool) -> str:
    """Trim text to a SharePoint column size."""
    value = clean_text(value, keep_newlines=not single_line)
    if single_line:
        value = one_line(value)
    if limit <= 0 or len(value) <= limit:
        return value
    if single_line:
        return value[:limit]
    marker = " [truncated]"
    keep = max(0, limit - len(marker))
    return value[:keep] + marker


class _HTMLText(HTMLParser):
    """Collect visible text and skip script and style blocks."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in ("script", "style"):
            self._skip_depth += 1
        if tag in ("p", "br", "div", "tr", "li", "h1", "h2", "h3", "h4"):
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in ("script", "style") and self._skip_depth:
            self._skip_depth -= 1
        if tag in ("p", "div", "tr", "li"):
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self._skip_depth:
            self.parts.append(data)


def html_to_text(value: object) -> str:
    """Convert an HTML body to plain text for later classification."""
    raw = safe_decode(value)
    try:
        parser = _HTMLText()
        parser.feed(raw)
        parser.close()
        text = "".join(parser.parts)
    except Exception:
        without_hidden = re.sub(
            r"(?is)<(script|style)\b[^>]*>.*?</\1>",
            " ",
            raw,
        )
        text = re.sub(r"(?s)<[^>]+>", " ", without_hidden)
        text = html.unescape(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"[ \t]*\n[ \t]*", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def parse_transport_headers(raw: object) -> dict[str, str]:
    """Read From, To, Cc, Bcc, and Message-ID from a PST transport header block."""
    text = safe_decode(raw).strip()
    if not text:
        return {}
    message = message_from_string(text + "\n\n")
    return {
        "from": message.get("From", "") or "",
        "to": message.get("To", "") or "",
        "cc": message.get("Cc", "") or "",
        "bcc": message.get("Bcc", "") or "",
        "message_id": (message.get("Message-ID", "") or "").strip(),
    }


def split_mailbox(value: str) -> tuple[str, str]:
    """Return (display name, email). A bare name is not treated as an address."""
    name, address = parseaddr(value or "")
    name = name.strip().strip('"')
    address = address.strip()
    if address and "@" not in address:
        if not name:
            name = address
        address = ""
    return name, address


def format_recipients(*header_values: str) -> str:
    """Join To, Cc, and Bcc into one readable recipient string."""
    formatted: list[str] = []
    for value in header_values:
        if not value or not value.strip():
            continue
        addresses = [
            (name.strip(), addr.strip())
            for name, addr in getaddresses([value])
            if name.strip() or addr.strip()
        ]
        if not addresses:
            collapsed = one_line(value)
            if collapsed:
                formatted.append(collapsed)
            continue
        for name, address in addresses:
            if address and "@" not in address and not name:
                formatted.append(address)
            elif name and address:
                formatted.append(f"{name} <{address}>")
            elif address or name:
                formatted.append(address or name)
    return "; ".join(formatted)


def create_email_id(
    *,
    source_file: str,
    folder: str,
    subject: str,
    sender: str,
    date: str,
    internet_message_id: str,
    pst_identifier: str,
    body: str,
) -> str:
    """Stable id for one message in one folder.

    A later run of the same PST produces the same id, so the importer can
    skip rows it already created. The body fingerprint separates two
    messages that share a subject line when the PST has no message id.
    """
    body_fingerprint = hashlib.sha256(body.encode("utf-8")).hexdigest()[:16]
    source = "|".join(
        [
            source_file,
            folder,
            internet_message_id.strip().lower(),
            pst_identifier,
            subject,
            sender,
            date,
            body_fingerprint,
        ]
    )
    return hashlib.sha256(source.encode("utf-8")).hexdigest()
