"""Append-only record of SharePoint items this importer has already created."""

from __future__ import annotations

import json
import os
from pathlib import Path


class ImportState:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.ids: set[str] = set()

    def load(self) -> None:
        self.ids.clear()
        if not self.path.is_file():
            return
        for line in self.path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                payload = json.loads(line)
            except json.JSONDecodeError:
                continue
            email_id = payload.get("email_id") if isinstance(payload, dict) else None
            if email_id:
                self.ids.add(str(email_id))

    def __contains__(self, email_id: str) -> bool:
        return email_id in self.ids

    def add(self, email_id: str) -> bool:
        """Remember an id. Returns False when it was already recorded."""
        if not email_id or email_id in self.ids:
            return False
        self.ids.add(email_id)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps({"email_id": email_id}) + "\n")
            handle.flush()
            os.fsync(handle.fileno())
        return True
