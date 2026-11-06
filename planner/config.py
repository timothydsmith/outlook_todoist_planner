"""Settings, saved as JSON in the user's profile."""
from __future__ import annotations

import json
import os
import sys
import threading
from pathlib import Path

DEFAULTS = {
    "todoist_token": "",
    "todoist_filter": "",
    "calendar_source": "ics",          # "ics" | "outlook_desktop" | "none"
    "ics_url": "",
    "week_start": 1,                   # 1 = Monday, 0 = Sunday
    "show_weekends": True,
    "day_start_hour": 7,
    "default_task_minutes": 30,
    "refresh_minutes": 5,
}


def config_path() -> Path:
    override = os.environ.get("PLANNER_CONFIG")
    if override:
        return Path(override)
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA") or (Path.home() / "AppData" / "Roaming"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or (Path.home() / ".config"))
    return base / "OutlookTodoistPlanner" / "settings.json"


class Config:
    def __init__(self, path: Path, memory_only: bool = False):
        self.path = path
        self.memory_only = memory_only
        self.lock = threading.Lock()
        self.data = dict(DEFAULTS)
        if not memory_only and path.exists():
            try:
                loaded = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    for k in DEFAULTS:
                        if k in loaded:
                            self.data[k] = loaded[k]
            except (OSError, ValueError) as exc:
                print(f"Couldn't read settings from {path}: {exc}. Starting with defaults.")

    def get(self, key):
        with self.lock:
            return self.data.get(key, DEFAULTS.get(key))

    def snapshot(self) -> dict:
        with self.lock:
            return dict(self.data)

    def public(self) -> dict:
        """Settings for the UI. The token itself never goes back to the page."""
        d = self.snapshot()
        tok = d.pop("todoist_token", "") or ""
        d["has_token"] = bool(tok)
        d["token_hint"] = ("…" + tok[-4:]) if len(tok) >= 8 else ""
        d["platform"] = sys.platform
        return d

    def update(self, changes: dict) -> None:
        clean = {}
        for k, v in changes.items():
            if k not in DEFAULTS:
                continue
            if k == "todoist_token":
                v = str(v or "").strip()
                if not v:
                    continue  # blank means "keep the current token"
                if v == "__clear__":
                    v = ""
            elif k in ("todoist_filter", "ics_url"):
                v = str(v or "").strip()
            elif k == "calendar_source":
                v = v if v in ("ics", "outlook_desktop", "none") else "ics"
            elif k == "show_weekends":
                v = bool(v)
            elif k == "week_start":
                v = 0 if int(v) == 0 else 1
            elif k == "day_start_hour":
                v = min(12, max(0, int(v)))
            elif k == "default_task_minutes":
                v = min(480, max(5, int(v)))
            elif k == "refresh_minutes":
                v = min(60, max(1, int(v)))
            clean[k] = v
        with self.lock:
            self.data.update(clean)
            data = dict(self.data)
        if self.memory_only:
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
        try:
            os.chmod(tmp, 0o600)
        except OSError:
            pass
        os.replace(tmp, self.path)
