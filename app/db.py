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


def bom_replace_by_join(
    engine: Engine,
    table_name: str,
    rows: list[dict[str, Any]],
    join_col: str,
    carry_cols: list[str],
    delete_col: str | None = None,
) -> tuple[int, int, list[str]]:
    """
    Replace BOM rows in three atomic steps:

    1. Look up carry_cols (e.g. id_produk, created_at) from DB keyed by join_col.
       Only products FOUND in DB are processed; not-found are skipped safely.
    2. DELETE existing rows using delete_col (e.g. id_produk) if provided,
       otherwise by join_col (e.g. nama_produk). Only found products are deleted.
    3. INSERT fresh rows from Excel with carry_cols merged in.

    Returns (inserted, skipped, not_found_vals).
    """
    if not rows:
        return 0, 0, []

    from sqlalchemy import delete as sa_delete

    metadata = MetaData()
    tbl = Table(table_name, metadata, autoload_with=engine)

    # ── 1. Collect unique join values from Excel rows ───────────
    unique_vals = list({str(r[join_col]) for r in rows if r.get(join_col) is not None})

    # ── 2. Fetch carry_cols + delete_col from DB ────────────────
    valid_carry = [c for c in carry_cols if c in tbl.c.keys()]
    fetch_extra = list(dict.fromkeys(
        valid_carry +
        ([delete_col] if delete_col and delete_col in tbl.c.keys() and delete_col not in valid_carry else [])
    ))

    carry_map: dict[str, dict[str, Any]] = {}
    # Always query DB to confirm which join values actually exist,
    # even when fetch_extra is empty (no carry cols selected).
    fetch_cols = [tbl.c[join_col]] + [tbl.c[c] for c in fetch_extra]
    with engine.connect() as conn:
        result = conn.execute(
            select(*fetch_cols).where(tbl.c[join_col].in_(unique_vals))
        ).fetchall()
    for r in result:
        key = str(r[0])
        if key not in carry_map:
            carry_map[key] = {fetch_extra[i]: r[i + 1] for i in range(len(fetch_extra))}

    not_found_vals = [v for v in unique_vals if v not in carry_map]
    found_vals = [v for v in unique_vals if v in carry_map]

    # ── 3. DELETE only found products ───────────────────────────
    if found_vals:
        if delete_col and delete_col in tbl.c.keys():
            delete_ids = list({
                carry_map[v][delete_col]
                for v in found_vals
                if carry_map[v].get(delete_col) is not None
            })
            if delete_ids:
                with engine.begin() as conn:
                    conn.execute(sa_delete(tbl).where(tbl.c[delete_col].in_(delete_ids)))
        else:
            with engine.begin() as conn:
                conn.execute(sa_delete(tbl).where(tbl.c[join_col].in_(found_vals)))

    # ── 4. INSERT — skip rows whose join_col not found in DB ────
    not_found_set = set(not_found_vals)
    insert_rows_list: list[dict[str, Any]] = []
    for row in rows:
        jv = str(row[join_col]) if row.get(join_col) is not None else None
        if jv is None or jv in not_found_set:
            continue
        merged = dict(row)
        for cc, cv in carry_map.get(jv, {}).items():
            if cc not in merged or merged[cc] is None:
                merged[cc] = cv
        insert_rows_list.append(merged)

    # Uniform keyset for SQLAlchemy bulk INSERT
    if insert_rows_list:
        all_keys: set[str] = set()
        for r in insert_rows_list:
            all_keys.update(r.keys())
        insert_rows_list = [{k: r.get(k) for k in all_keys} for r in insert_rows_list]
        with engine.begin() as conn:
            conn.execute(insert(tbl), insert_rows_list)

    skipped = sum(1 for r in rows if str(r.get(join_col, "")) in not_found_set)
    return len(insert_rows_list), skipped, not_found_vals


def bulk_update_by_join(
    engine: Engine,
    table_name: str,
    rows: list[dict[str, Any]],
    join_col: str,
    update_cols: list[str],
) -> tuple[int, int]:
    """
    Bulk UPDATE rows in *table_name* by matching on *join_col*.

    For each row in *rows*:
      UPDATE table_name
         SET col1 = :col1, col2 = :col2, …
       WHERE join_col = :join_col_val

    Returns (updated, not_found) counts.
    """
    if not rows:
        return 0, 0

    metadata = MetaData()
    tbl = Table(table_name, metadata, autoload_with=engine)

    updated = 0
    not_found = 0

    set_clause = {c: tbl.c[c] for c in update_cols if c in tbl.c}
    if not set_clause:
        raise ValueError(f"None of update_cols {update_cols} exist in table {table_name}")

    from sqlalchemy import update as sa_update

    with engine.begin() as conn:
        for row in rows:
            join_val = row.get(join_col)
            if join_val is None:
                not_found += 1
                continue
            stmt = (
                sa_update(tbl)
                .where(tbl.c[join_col] == join_val)
                .values({c: row.get(c) for c in update_cols})
            )
            result = conn.execute(stmt)
            if result.rowcount and result.rowcount > 0:
                updated += result.rowcount
            else:
                not_found += 1

    return updated, not_found
