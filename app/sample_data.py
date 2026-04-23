from __future__ import annotations

from typing import Any

from sqlalchemy.engine import Engine

try:
    from app.db import insert_rows
except ModuleNotFoundError:
    from db import insert_rows


def seed_reference_data(engine: Engine, table_name: str, rows: list[dict[str, Any]]) -> int:
    inserted, _ = insert_rows(engine, table_name, rows, conflict_strategy="skip")
    return inserted
