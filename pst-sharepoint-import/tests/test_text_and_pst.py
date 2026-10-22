import logging
import sys
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pst_sharepoint_import.pst import (  # noqa: E402
    ATTACH_FILENAME,
    ATTACH_LONG_FILENAME,
    ReadFailure,
    _walk_folder,
    extract_record,
    folder_selected,
    format_graph_datetime,
)
from pst_sharepoint_import.textutil import (  # noqa: E402
    create_email_id,
    html_to_text,
    safe_decode,
)


class _Entry:
    def __init__(self, entry_type, data):
        self.entry_type = entry_type
        self.data = data

    def get_entry_type(self):
        return self.entry_type

    def get_data_as_string(self):
        return self.data


class _RecordSet:
    def __init__(self, entries):
        self.entries = entries
        self.number_of_entries = len(entries)

    def get_entry(self, index):
        return self.entries[index]


class _Attachment:
    def __init__(self, name=None, entries=None):
        self.name = name
        self.entries = entries or []
        self.number_of_record_sets = 1 if self.entries else 0

    def get_record_set(self, index):
        return _RecordSet(self.entries)


class _Message:
    def __init__(self, **kwargs):
        self.subject = kwargs.get("subject", "Amplifier trip")
        self.sender_name = kwargs.get("sender_name", "PST Name")
        self.plain_text_body = kwargs.get("plain_text_body", b"The unit tripped.")
        self.html_body = kwargs.get("html_body")
        self.rtf_body = kwargs.get("rtf_body")
        self.transport_headers = kwargs.get("transport_headers", "")
        self.client_submit_time = kwargs.get("client_submit_time")
        self.delivery_time = kwargs.get("delivery_time")
        self.creation_time = kwargs.get("creation_time")
        self.conversation_topic = kwargs.get("conversation_topic", "")
        self.identifier_value = kwargs.get("identifier", 42)
        self.attachments = kwargs.get("attachments", [])
        self.number_of_attachments = len(self.attachments)
        self.fail_attachment = kwargs.get("fail_attachment", False)

    def get_identifier(self):
        return self.identifier_value

    def get_attachment(self, index):
        if self.fail_attachment:
            raise OSError("bad attachment")
        return self.attachments[index]


class _Folder:
    def __init__(self, name, messages, children=None):
        self.name = name
        self.messages = messages
        self.children = children or []
        self.number_of_sub_messages = len(messages)
        self.number_of_sub_folders = len(self.children)

    def get_sub_message(self, index):
        message = self.messages[index]
        if isinstance(message, Exception):
            raise message
        return message

    def get_sub_folder(self, index):
        return self.children[index]


HEADERS = (
    "From: Larry Shank <larry@aetechron.com>\r\n"
    "To: Ann Ops <ann@aetechron.com>\r\n"
    "Cc: Bo Plant <bo@aetechron.com>\r\n"
    "Message-ID: <trip-1@aetechron.com>\r\n"
)


class TextAndPstTests(unittest.TestCase):
    def test_safe_decode_replaces_invalid_bytes_and_nuls(self):
        self.assertEqual(safe_decode(None), "")
        self.assertEqual(safe_decode(b"A\xff\x00B"), "A\ufffdB")
        self.assertEqual(safe_decode("ok"), "ok")

    def test_html_keeps_visible_text_and_drops_script(self):
        text = html_to_text("<p>A &amp; B</p><script>secret()</script><div>Keep</div>")
        self.assertIn("A & B", text)
        self.assertIn("Keep", text)
        self.assertNotIn("secret", text)

    def test_message_uses_headers_plain_body_and_stable_id(self):
        message = _Message(
            transport_headers=HEADERS,
            plain_text_body=b"Hello\x00 world",
            client_submit_time=datetime(2020, 5, 1, 15, 30, 0),
            attachments=[_Attachment(name="scope.png")],
        )
        record = extract_record(message, "Outlook/Inbox", "Larry_Shank.pst")
        self.assertEqual(record.sender, "larry@aetechron.com")
        self.assertEqual(record.sender_name, "Larry Shank")
        self.assertEqual(record.recipients, "Ann Ops <ann@aetechron.com>; Bo Plant <bo@aetechron.com>")
        self.assertEqual(record.date, "2020-05-01T15:30:00Z")
        self.assertEqual(record.body, "Hello world")
        self.assertEqual(record.body_source, "plain")
        self.assertEqual(record.attachments, "scope.png")
        self.assertTrue(record.has_attachments)
        self.assertEqual(record.internet_message_id, "<trip-1@aetechron.com>")
        self.assertEqual(record.pst_identifier, "42")
        again = extract_record(message, "Outlook/Inbox", "Larry_Shank.pst")
        self.assertEqual(record.email_id, again.email_id)

        message.plain_text_body = b"Different body"
        changed = extract_record(message, "Outlook/Inbox", "Larry_Shank.pst")
        self.assertNotEqual(record.email_id, changed.email_id)

    def test_html_and_rtf_fallbacks(self):
        html_message = _Message(
            plain_text_body=None,
            html_body="<html><body><p>Trip</p><style>.x{color:red}</style></body></html>",
        )
        html_record = extract_record(html_message, "Inbox", "Larry_Shank.pst")
        self.assertEqual(html_record.body_source, "html")
        self.assertIn("Trip", html_record.body)
        self.assertNotIn("color", html_record.body)

        rtf_message = _Message(plain_text_body=b"  ", html_body=None, rtf_body=b"{\\rtf1 hi}")
        rtf_record = extract_record(rtf_message, "Inbox", "Larry_Shank.pst")
        self.assertEqual(rtf_record.body_source, "rtf")
        self.assertIn("RTF body", rtf_record.body)

    def test_attachment_long_filename_beats_short_name(self):
        attachment = _Attachment(
            entries=[
                _Entry(ATTACH_FILENAME, "short.txt"),
                _Entry(ATTACH_LONG_FILENAME, "oscilloscope capture.txt"),
            ]
        )
        record = extract_record(
            _Message(plain_text_body=b"body", attachments=[attachment]),
            "Inbox",
            "Larry_Shank.pst",
        )
        self.assertEqual(record.attachments, "oscilloscope capture.txt")

    def test_unreadable_attachment_is_named_and_kept(self):
        record = extract_record(
            _Message(number_of_attachments=1, fail_attachment=True, attachments=[object()]),
            "Inbox",
            "Larry_Shank.pst",
        )
        self.assertEqual(record.attachments, "(unreadable-0)")
        self.assertTrue(record.has_attachments)

    def test_blank_subject_uses_conversation_topic(self):
        record = extract_record(
            _Message(subject="  ", conversation_topic="Power supply"),
            "Inbox",
            "Larry_Shank.pst",
        )
        self.assertEqual(record.subject, "Power supply")

    def test_folder_filter_and_read_errors(self):
        inbox = _Folder(
            "Inbox",
            [_Message(subject="Keep me"), OSError("corrupt message")],
        )
        sent = _Folder("Sent Items", [_Message(subject="Skip me")])
        root = _Folder("Outlook", [], [inbox, sent])
        items = list(_walk_folder(root, "", "Larry_Shank.pst", ["Inbox"]))
        self.assertEqual(len(items), 2)
        self.assertEqual(items[0].subject, "Keep me")
        self.assertIsInstance(items[1], ReadFailure)
        self.assertIn("corrupt", items[1].error)
        self.assertTrue(folder_selected("Outlook/Inbox/Customers", ["Inbox"]))
        self.assertFalse(folder_selected("Outlook/InboxArchive", ["Inbox"]))

    def test_datetime_formats_as_utc(self):
        central = timezone(timedelta(hours=-5))
        self.assertEqual(
            format_graph_datetime(datetime(2020, 5, 1, 10, 30, tzinfo=central)),
            "2020-05-01T15:30:00Z",
        )
        self.assertEqual(format_graph_datetime(None), "")

    def test_email_id_changes_with_folder(self):
        first = create_email_id(
            source_file="Larry_Shank.pst",
            folder="Inbox",
            subject="Hi",
            sender="a@aetechron.com",
            date="2020-01-01T00:00:00Z",
            internet_message_id="<id@aetechron.com>",
            pst_identifier="7",
            body="same",
        )
        second = create_email_id(
            source_file="Larry_Shank.pst",
            folder="Inbox/Customers",
            subject="Hi",
            sender="a@aetechron.com",
            date="2020-01-01T00:00:00Z",
            internet_message_id="<id@aetechron.com>",
            pst_identifier="7",
            body="same",
        )
        self.assertNotEqual(first, second)


if __name__ == "__main__":
    logging.disable(logging.CRITICAL)
    unittest.main()
