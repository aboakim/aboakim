import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pst_sharepoint_import.__main__ import main, parse_args  # noqa: E402
from pst_sharepoint_import.auth import GraphAuthenticator, _has_access_token  # noqa: E402
from pst_sharepoint_import.config import (  # noqa: E402
    ENV_OVERRIDES,
    ConfigError,
    load_settings,
)


class ConfigTests(unittest.TestCase):
    def setUp(self):
        self.saved = {key: os.environ.get(key) for key in ENV_OVERRIDES}
        for key in ENV_OVERRIDES:
            os.environ.pop(key, None)

    def tearDown(self):
        for key, value in self.saved.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value

    def test_env_and_dotenv_override_json_and_paths_follow_the_config(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / ".env").write_text("SHAREPOINT_LIST_ID=from-dotenv\n", encoding="utf-8")
            config_path = root / "config.json"
            config_path.write_text(
                json.dumps(
                    {
                        "pst_file": "C:/EmailArchive/Larry_Shank.pst",
                        "site_id": "from-file",
                        "list_id": "from-file",
                        "state_file": "state/imported.jsonl",
                        "field_map": {"keywords": ""},
                    }
                ),
                encoding="utf-8",
            )
            os.environ["SHAREPOINT_SITE_ID"] = "from-env"
            settings = load_settings(config_path)
            self.assertEqual(settings.site_id, "from-env")
            self.assertEqual(settings.list_id, "from-dotenv")
            self.assertEqual(settings.field_map["keywords"], "")
            self.assertEqual(settings.field_map["email_id"], "EmailID")
            self.assertEqual(settings.state_file, str((root / "state" / "imported.jsonl").resolve()))

    def test_validate_requires_the_pst_and_auth(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            pst = root / "Larry_Shank.pst"
            pst.write_bytes(b"placeholder")
            config_path = root / "config.json"
            config_path.write_text(
                json.dumps(
                    {
                        "pst_file": str(pst),
                        "dry_run": True,
                        "sync_existing": False,
                        "auth_mode": "device_code",
                    }
                ),
                encoding="utf-8",
            )
            settings = load_settings(config_path)
            settings.validate()

            settings.sync_existing = True
            with self.assertRaises(ConfigError):
                settings.validate()

            settings.sync_existing = False
            settings.dry_run = False
            settings.site_id = "site"
            settings.list_id = "list"
            with self.assertRaises(ConfigError):
                settings.validate()

            settings.tenant_id = "tenant"
            settings.client_id = "client"
            settings.field_map = dict(settings.field_map)
            settings.field_map["sender"] = "Sender Address"
            with self.assertRaises(ConfigError):
                settings.validate()

    def test_access_token_auth_returns_the_configured_token(self):
        with tempfile.TemporaryDirectory() as tmp:
            pst = Path(tmp) / "Larry_Shank.pst"
            pst.write_bytes(b"placeholder")
            config_path = Path(tmp) / "config.json"
            config_path.write_text(
                json.dumps(
                    {
                        "pst_file": str(pst),
                        "site_id": "site",
                        "list_id": "list",
                        "auth_mode": "access_token",
                    }
                ),
                encoding="utf-8",
            )
            os.environ["GRAPH_ACCESS_TOKEN"] = "secret-token"
            settings = load_settings(config_path)
            settings.validate()
            auth = GraphAuthenticator(settings)
            self.assertEqual(auth.get_token(), "secret-token")
            auth.invalidate()
            self.assertEqual(auth.get_token(), "secret-token")

    def test_main_reports_a_missing_config_or_pst(self):
        with tempfile.TemporaryDirectory() as tmp:
            missing_config = main(["--config", str(Path(tmp) / "missing.json")])
            self.assertEqual(missing_config, 2)
            config_path = Path(tmp) / "config.json"
            config_path.write_text(
                json.dumps({"pst_file": str(Path(tmp) / "missing.pst"), "dry_run": True, "sync_existing": False}),
                encoding="utf-8",
            )
            missing_pst = main(["--config", str(config_path), "--dry-run", "--no-sync-existing"])
            self.assertEqual(missing_pst, 2)

    def test_cached_sign_in_without_a_token_is_not_treated_as_success(self):
        self.assertFalse(_has_access_token(None))
        self.assertFalse(_has_access_token({"error": "interaction_required"}))
        self.assertTrue(_has_access_token({"access_token": "abc"}))

    def test_conflicting_sync_flags_exit(self):
        with self.assertRaises(SystemExit):
            parse_args(["--config", "config.json", "--sync-existing", "--no-sync-existing"])


if __name__ == "__main__":
    unittest.main()
