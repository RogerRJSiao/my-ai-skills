"""Audit log: one JSON object per tool call (JSON Lines).

Only file *names* are recorded, never file contents, summaries or notes.
Student names still appear inside file names (部門代號-工號-姓名.xlsx);
masking them is on the roadmap (core/masking.py) before external use.
"""
from __future__ import annotations

import datetime as dt
import json
import threading
from pathlib import Path, PurePath

_PATH_KEYS = ("path", "folder", "file")


def _name(v) -> str:
    return PurePath(str(v)).name


def summarize_args(kwargs: dict) -> dict:
    """Keep what is useful for an audit trail and drop the rest."""
    out: dict = {}
    for k, v in kwargs.items():
        if v is None:
            continue
        if any(t in k for t in _PATH_KEYS) and isinstance(v, str):
            out[k] = _name(v)
        elif k == "students" and isinstance(v, list):
            out[k] = [
                _name(getattr(s, "submission_path", None) or (s.get("submission_path") if isinstance(s, dict) else s))
                for s in v
            ]
        elif isinstance(v, (str, int, float, bool)) and k in ("skill", "on_duplicate", "overwrite"):
            out[k] = v
    return out


class AuditLogger:
    def __init__(self, path: Path | None, server: str):
        self.path = path
        self.server = server
        self._lock = threading.Lock()

    def record(self, *, tool: str, user: str, status: str, duration_ms: int,
               args: dict, error: str | None = None) -> None:
        if self.path is None:
            return
        entry = {
            "ts": dt.datetime.now(dt.timezone.utc).astimezone().isoformat(timespec="seconds"),
            "server": self.server,
            "tool": tool,
            "user": user,
            "status": status,
            "duration_ms": duration_ms,
            "args": summarize_args(args),
        }
        if error:
            entry["error"] = error[:500]
        line = json.dumps(entry, ensure_ascii=False, default=str)
        with self._lock:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with self.path.open("a", encoding="utf-8") as f:
                f.write(line + "\n")
