from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any


LOG_FILE = Path("uploads/action_log.jsonl")


def log_action(action: str, status: str = "info", payload: dict[str, Any] | None = None) -> None:
    """Append one action record as JSONL. Never raise to caller."""
    try:
        LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        record = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "action": action,
            "status": status,
            "payload": payload or {},
        }
        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception:
        return


def read_action_logs(limit: int = 50) -> list[dict[str, Any]]:
    """Read latest action records from JSONL file."""
    if not LOG_FILE.exists():
        return []

    rows: list[dict[str, Any]] = []
    try:
        with LOG_FILE.open("r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rows.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    except Exception:
        return []

    if limit <= 0:
        return rows
    return rows[-limit:]