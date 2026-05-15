from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from sqlalchemy import Column, MetaData, Table, create_engine, inspect, insert, select
from sqlalchemy.dialects.mysql import insert as mysql_insert
from sqlalchemy.dialects.sqlite import insert as sqlite_insert
from sqlalchemy.engine import Engine
from sqlalchemy.sql.sqltypes import (
    Boolean,
    Date,
    DateTime,
    Float,
    Integer,
    String,
    Text,
)


TYPE_MAP = {
    "integer": Integer,
    "float": Float,
    "string": String,
    "text": Text,
    "boolean": Boolean,
    "date": Date,
    "datetime": DateTime,
}


@dataclass
class ColumnSpec:
    name: str
    type: str = "string"
    nullable: bool = True
    primary_key: bool = False
    unique: bool = False
    length: int | None = None


def get_engine(db_url: str) -> Engine:
    import sys
    # On Windows, AF_UNIX is unavailable — MySQL via localhost uses Unix socket
    # which fails. Force TCP by replacing localhost → 127.0.0.1 and stripping
    # unix_socket query param.
    if sys.platform == "win32" and db_url.startswith("mysql"):
        import re
        db_url = re.sub(r"@localhost", "@127.0.0.1", db_url)
        db_url = re.sub(r"[?&]unix_socket=[^&]*", "", db_url)
    return create_engine(db_url, future=True)


def _build_column(spec: ColumnSpec) -> Column:
    col_type = TYPE_MAP.get(spec.type.lower(), String)
    if col_type is String and spec.length:
        sql_type = String(spec.length)
    else:
        sql_type = col_type()

    return Column(
        spec.name,
        sql_type,
        nullable=spec.nullable,
        primary_key=spec.primary_key,
        unique=spec.unique,
    )


def create_table_from_schema(engine: Engine, table_name: str, columns: list[ColumnSpec]) -> None:
    metadata = MetaData()
    Table(table_name, metadata, *[_build_column(c) for c in columns], extend_existing=True)
    metadata.create_all(engine)


def list_tables(engine: Engine) -> list[str]:
    return inspect(engine).get_table_names()


def get_table_columns(engine: Engine, table_name: str) -> list[str]:
    return [c["name"] for c in inspect(engine).get_columns(table_name)]


def load_lookup_map(
    engine: Engine,
    master_table: str,
    match_col: str,
    return_col: str,
) -> dict[str, Any]:
    """
    Load entire master table into memory as a lookup dict.
    Returns {match_value_str: return_value} for fast O(1) lookup per row.
    """
    metadata = MetaData()
    table = Table(master_table, metadata, autoload_with=engine)
    with engine.connect() as conn:
        rows = conn.execute(select(table.c[match_col], table.c[return_col])).fetchall()
    return {str(row[0]): row[1] for row in rows}


def insert_rows(
    engine: Engine,
    table_name: str,
    rows: list[dict[str, Any]],
    conflict_strategy: str = "error",  # "error" | "skip" | "replace"
) -> tuple[int, int]:
    """Insert rows and return (inserted, skipped) counts."""
    if not rows:
        return 0, 0

    metadata = MetaData()
    table = Table(table_name, metadata, autoload_with=engine)
    dialect = engine.dialect.name

    inserted_count = 0
    skipped_count = 0

    if conflict_strategy == "error":
        with engine.begin() as conn:
            conn.execute(insert(table), rows)
        return len(rows), 0

    if conflict_strategy == "skip":
        if dialect == "sqlite":
            stmt = sqlite_insert(table).prefix_with("OR IGNORE")
            with engine.begin() as conn:
                result = conn.execute(stmt, rows)
                inserted_count = result.rowcount if result.rowcount >= 0 else len(rows)
            skipped_count = len(rows) - inserted_count
            return inserted_count, skipped_count
        elif dialect == "mysql":
            stmt = mysql_insert(table).prefix_with("IGNORE")
            with engine.begin() as conn:
                result = conn.execute(stmt, rows)
                inserted_count = result.rowcount if result.rowcount >= 0 else len(rows)
            skipped_count = len(rows) - inserted_count
            return inserted_count, skipped_count
        else:
            # Generic: insert one by one, skip on error
            for row in rows:
                try:
                    with engine.begin() as conn:
                        conn.execute(insert(table), [row])
                    inserted_count += 1
                except Exception:
                    skipped_count += 1
            return inserted_count, skipped_count

    if conflict_strategy == "replace":
        if dialect == "sqlite":
            stmt = sqlite_insert(table).prefix_with("OR REPLACE")
            with engine.begin() as conn:
                conn.execute(stmt, rows)
            return len(rows), 0
        elif dialect == "mysql":
            base = mysql_insert(table)
            inserted_cols = set(rows[0].keys())
            update_cols = {
                c.name: base.inserted[c.name]
                for c in table.columns
                if not c.primary_key and c.name in inserted_cols
            }
            stmt = base.on_duplicate_key_update(**update_cols) if update_cols else base.prefix_with("IGNORE")
            with engine.begin() as conn:
                conn.execute(stmt, rows)
            return len(rows), 0
        else:
            for row in rows:
                try:
                    with engine.begin() as conn:
                        conn.execute(insert(table), [row])
                    inserted_count += 1
                except Exception:
                    skipped_count += 1
            return inserted_count, skipped_count

    raise ValueError(f"Unknown conflict_strategy: {conflict_strategy}")
