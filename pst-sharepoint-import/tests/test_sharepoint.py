import json
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pst_sharepoint_import.config import Settings  # noqa: E402
from pst_sharepoint_import.fields import DEFAULT_FIELD_MAP, build_fields, column_definition  # noqa: E402
from pst_sharepoint_import.pst import EmailRecord  # noqa: E402
from pst_sharepoint_import.sharepoint import GraphError, SharePointListClient  # noqa: E402


class FakeResponse:
    def __init__(self, status, payload=None, headers=None):
        self.status_code = status
        self._payload = {} if payload is None else payload
        self.headers = headers or {}
        self.text = json.dumps(self._payload)

    def json(self):
        return self._payload


class FakeSession:
    def __init__(self, responses=None):
        self.responses = list(responses or [])
        self.calls = []

    def request(self, method, url, headers=None, json=None, timeout=None):
        self.calls.append({"method": method, "url": url, "json": json, "headers": headers})
        if not self.responses:
            raise AssertionError(f"unexpected {method} {url}")
        outcome = self.responses.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome


class FakeAuth:
    def __init__(self):
        self.invalidations = 0
        self.token = "token-1"

    def get_token(self):
        return self.token

    def invalidate(self):
        self.invalidations += 1
        self.token = "token-2"


def sample_record(**overrides):
    values = dict(
        email_id="a" * 64,
        subject="Amplifier trip",
        sender="larry@aetechron.com",
        sender_name="Larry Shank",
        recipients="ann@aetechron.com",
        date="2020-05-01T15:30:00Z",
        folder="Outlook/Inbox",
        body="The unit tripped.",
        body_source="plain",
        attachments="scope.png",
        has_attachments=True,
        internet_message_id="<trip-1@aetechron.com>",
        pst_identifier="42",
        source_file="Larry_Shank.pst",
    )
    values.update(overrides)
    return EmailRecord(**values)


def settings(**overrides):
    value = Settings(
        pst_file="Larry_Shank.pst",
        site_id="site",
        list_id="list",
        auth_mode="client_credentials",
        tenant_id="tenant",
        client_id="client",
        client_secret="secret",
        max_retries=3,
        request_delay_seconds=0,
    )
    for key, item in overrides.items():
        setattr(value, key, item)
    return value


class FieldTests(unittest.TestCase):
    def test_payload_omits_blank_date_and_keeps_ai_defaults(self):
        record = sample_record(date="", subject="  \n  ")
        fields = build_fields(record, settings())
        self.assertEqual(fields["Title"], "(no subject)")
        self.assertNotIn("EmailDate", fields)
        self.assertIs(fields["TechnicalRelevant"], False)
        self.assertEqual(fields["TechnicalScore"], 0)
        self.assertIs(fields["AIProcessed"], False)
        self.assertIs(fields["HasAttachments"], True)
        self.assertEqual(fields["Customer"], "")

    def test_disabled_column_is_omitted_and_subject_is_capped(self):
        chosen = settings()
        chosen.field_map = dict(chosen.field_map)
        chosen.field_map["keywords"] = ""
        chosen.limits.subject = 10
        fields = build_fields(sample_record(subject="ABCDEFGHIJextra"), chosen)
        self.assertNotIn("Keywords", fields)
        self.assertEqual(fields["Title"], "ABCDEFGHIJ")
        self.assertEqual(
            build_fields(sample_record(subject="Amp\nlifier"), chosen)["Title"],
            "Amp lifier",
        )

    def test_column_types(self):
        self.assertIsNone(column_definition("subject", "Title"))
        body = column_definition("body", "EmailBody")
        self.assertTrue(body["text"]["allowMultipleLines"])
        self.assertIn("dateTime", column_definition("date", "EmailDate"))
        self.assertIn("boolean", column_definition("has_attachments", "HasAttachments"))
        self.assertIn("number", column_definition("technical_score", "TechnicalScore"))
        email_id = column_definition("email_id", "EmailID")
        self.assertTrue(email_id["indexed"])
        self.assertFalse(email_id["text"]["allowMultipleLines"])


class GraphClientTests(unittest.TestCase):
    def setUp(self):
        self.auth = FakeAuth()

    def client(self, session, **setting_overrides):
        return SharePointListClient(settings(**setting_overrides), session, self.auth)

    @patch("pst_sharepoint_import.sharepoint.time.sleep")
    def test_rate_limit_waits_then_succeeds(self, sleep):
        session = FakeSession(
            [
                FakeResponse(429, {"error": {"message": "slow down"}}, headers={"Retry-After": "7"}),
                FakeResponse(201, {"id": "99"}),
            ]
        )
        payload = self.client(session).create_item(sample_record())
        self.assertEqual(payload["id"], "99")
        sleep.assert_called_once_with(7.0)
        self.assertEqual(len(session.calls), 2)
        self.assertEqual(session.calls[0]["json"]["fields"]["Sender"], "larry@aetechron.com")

    @patch("pst_sharepoint_import.sharepoint.time.sleep")
    def test_rate_limit_retries_are_bounded(self, sleep):
        session = FakeSession([FakeResponse(429, headers={"Retry-After": "0"}) for _ in range(5)])
        with self.assertRaises(GraphError):
            self.client(session).create_item(sample_record())
        self.assertEqual(len(session.calls), 3)
        self.assertEqual(sleep.call_count, 2)

    def test_expired_token_is_refreshed_once(self):
        session = FakeSession(
            [
                FakeResponse(401, {"error": {"message": "expired"}}),
                FakeResponse(201, {"id": "1"}),
            ]
        )
        self.client(session).create_item(sample_record())
        self.assertEqual(self.auth.invalidations, 1)
        self.assertEqual(session.calls[0]["headers"]["Authorization"], "Bearer token-1")
        self.assertEqual(session.calls[1]["headers"]["Authorization"], "Bearer token-2")

    def test_static_access_token_is_not_refreshed(self):
        session = FakeSession([FakeResponse(401, {"error": {"message": "expired"}})])
        with self.assertRaises(GraphError):
            self.client(session, auth_mode="access_token", access_token="static").create_item(sample_record())
        self.assertEqual(self.auth.invalidations, 0)
        self.assertEqual(len(session.calls), 1)

    def test_existing_ids_follow_next_link(self):
        session = FakeSession(
            [
                FakeResponse(
                    200,
                    {
                        "value": [{"fields": {"EmailID": "aaa"}}],
                        "@odata.nextLink": "https://graph.microsoft.com/next",
                    },
                ),
                FakeResponse(200, {"value": [{"fields": {"EmailID": "bbb"}}]}),
            ]
        )
        ids = list(self.client(session).iter_existing_email_ids("EmailID"))
        self.assertEqual(ids, ["aaa", "bbb"])
        self.assertEqual(session.calls[1]["url"], "https://graph.microsoft.com/next")

    def test_ensure_columns_skips_title_and_existing_names(self):
        class DirectorySession:
            def __init__(self):
                self.posts = []

            def request(self, method, url, headers=None, json=None, timeout=None):
                if method == "GET":
                    return FakeResponse(200, {"value": [{"name": "Title"}, {"name": "EmailID"}]})
                self.posts.append(json)
                if json["name"] == "Sender":
                    return FakeResponse(409, {"error": {"message": "exists"}})
                return FakeResponse(201, {"name": json["name"]})

        session = DirectorySession()
        self.client(session).ensure_columns()
        posted = [body["name"] for body in session.posts]
        self.assertNotIn("Title", posted)
        self.assertNotIn("EmailID", posted)
        self.assertIn("Sender", posted)
        expected = [
            name
            for name in DEFAULT_FIELD_MAP.values()
            if name.lower() not in ("title", "emailid")
        ]
        self.assertCountEqual(posted, expected)
        body = next(item for item in session.posts if item["name"] == "EmailBody")
        self.assertTrue(body["text"]["allowMultipleLines"])


if __name__ == "__main__":
    unittest.main()
