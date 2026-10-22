import logging
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pst_sharepoint_import.config import ConfigError, Settings  # noqa: E402
from pst_sharepoint_import.importer import run_import  # noqa: E402
from pst_sharepoint_import.pst import EmailRecord, ReadFailure  # noqa: E402
from pst_sharepoint_import.sharepoint import GraphError  # noqa: E402
from pst_sharepoint_import.state import ImportState  # noqa: E402


def record(email_id, subject="Hello"):
    return EmailRecord(
        email_id=email_id,
        subject=subject,
        sender="larry@aetechron.com",
        sender_name="Larry Shank",
        recipients="ann@aetechron.com",
        date="2020-05-01T15:30:00Z",
        folder="Outlook/Inbox",
        body="The unit tripped.",
        body_source="plain",
        attachments="",
        has_attachments=False,
        internet_message_id="",
        pst_identifier="1",
        source_file="Larry_Shank.pst",
    )


class FakeGraph:
    def __init__(self, fail_ids=None, existing=None, missing=None):
        self.created = []
        self.fail_ids = set(fail_ids or [])
        self.existing = list(existing or [])
        self.missing = list(missing or [])
        self.ensured = False

    def ensure_columns(self):
        self.ensured = True
        self.missing = []

    def missing_columns(self):
        return list(self.missing)

    def iter_existing_email_ids(self, column_name):
        return iter(self.existing)

    def create_item(self, item):
        if item.email_id in self.fail_ids:
            raise GraphError(400, "bad item")
        self.created.append(item.email_id)
        return {"id": item.email_id}


def make_settings(tmp):
    return Settings(
        pst_file="Larry_Shank.pst",
        site_id="site",
        list_id="list",
        auth_mode="access_token",
        access_token="token",
        sync_existing=False,
        ensure_columns=False,
        dry_run=False,
        request_delay_seconds=0,
        state_file=str(Path(tmp) / "state" / "imported.jsonl"),
        log_dir=str(Path(tmp) / "logs"),
    )


class ImporterTests(unittest.TestCase):
    def test_imports_skips_duplicates_and_retries_failures(self):
        with tempfile.TemporaryDirectory() as tmp:
            chosen = make_settings(tmp)
            graph = FakeGraph(fail_ids={"b"})
            summary = run_import(
                chosen,
                records=[record("a"), record("a", "copy"), record("b"), ReadFailure("Inbox", 3, "bad")],
                client=graph,
            )
            self.assertEqual(summary.imported, 1)
            self.assertEqual(summary.skipped_duplicates, 1)
            self.assertEqual(summary.failed, 1)
            self.assertEqual(summary.read_errors, 1)
            self.assertEqual(summary.exit_code(), 1)
            self.assertEqual(graph.created, ["a"])

            graph.fail_ids.clear()
            again = run_import(chosen, records=[record("a"), record("b")], client=graph)
            self.assertEqual(again.skipped_duplicates, 1)
            self.assertEqual(again.imported, 1)
            self.assertEqual(again.exit_code(), 0)
            self.assertFalse(Path(chosen.state_file).with_suffix(".lock").exists())

    def test_dry_run_does_not_write_state(self):
        with tempfile.TemporaryDirectory() as tmp:
            chosen = make_settings(tmp)
            chosen.dry_run = True
            graph = FakeGraph()
            summary = run_import(chosen, records=[record("a"), record("b")], client=graph)
            self.assertEqual(summary.dry_run, 2)
            self.assertEqual(graph.created, [])
            self.assertFalse(Path(chosen.state_file).exists())

    def test_limit_and_remote_duplicates(self):
        with tempfile.TemporaryDirectory() as tmp:
            chosen = make_settings(tmp)
            chosen.sync_existing = True
            chosen.limit = 2
            graph = FakeGraph(existing=["a"])
            summary = run_import(
                chosen,
                records=[record("a"), record("b"), record("c")],
                client=graph,
            )
            self.assertEqual(summary.seen, 2)
            self.assertEqual(summary.skipped_duplicates, 1)
            self.assertEqual(summary.imported, 1)
            self.assertEqual(graph.created, ["b"])
            state = ImportState(Path(chosen.state_file))
            state.load()
            self.assertIn("a", state)
            self.assertIn("b", state)

    def test_missing_columns_stop_the_run(self):
        with tempfile.TemporaryDirectory() as tmp:
            chosen = make_settings(tmp)
            graph = FakeGraph(missing=["EmailBody"])
            with self.assertRaises(ConfigError):
                run_import(chosen, records=[record("a")], client=graph)
            self.assertEqual(graph.created, [])

    def test_ensure_columns_clears_the_gap(self):
        with tempfile.TemporaryDirectory() as tmp:
            chosen = make_settings(tmp)
            chosen.ensure_columns = True
            graph = FakeGraph(missing=["EmailBody"])
            summary = run_import(chosen, records=[record("a")], client=graph)
            self.assertTrue(graph.ensured)
            self.assertEqual(summary.imported, 1)

    def test_lock_blocks_a_second_run(self):
        with tempfile.TemporaryDirectory() as tmp:
            chosen = make_settings(tmp)
            lock = Path(chosen.state_file).with_suffix(".lock")
            lock.parent.mkdir(parents=True)
            lock.write_text("999", encoding="utf-8")
            with self.assertRaises(ConfigError):
                run_import(chosen, records=[record("a")], client=FakeGraph())


class StateTests(unittest.TestCase):
    def test_round_trip_ignores_corrupt_lines(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "imported.jsonl"
            state = ImportState(path)
            self.assertTrue(state.add("abc"))
            self.assertFalse(state.add("abc"))
            with path.open("a", encoding="utf-8") as handle:
                handle.write("not-json\n")
            reloaded = ImportState(path)
            reloaded.load()
            self.assertIn("abc", reloaded)
            self.assertNotIn("other", reloaded)


if __name__ == "__main__":
    logging.disable(logging.CRITICAL)
    unittest.main()
