from __future__ import annotations

import re
from dataclasses import dataclass


class SQLParseError(ValueError):
    pass


@dataclass
class ParsedInsert:
    table_name: str
    columns: list[str]
    sample_rows: list[dict]


def parse_insert_sql(sql: str) -> ParsedInsert:
    """
    Parse one or more INSERT INTO statements.
    Returns table name, column list, and sample rows (if values present).

    Supports:
        INSERT INTO table (col1, col2) VALUES (v1, v2), (v3, v4);
        INSERT INTO `table` (`col1`, `col2`) VALUES ...
    """
    sql = sql.strip().rstrip(";")

    header_match = re.search(
        r"INSERT\s+INTO\s+[`'\"]?(\w+)[`'\"]?\s*\(([^)]+)\)",
        sql,
        re.IGNORECASE,
    )
    if not header_match:
        raise SQLParseError(
            "Format tidak dikenali. Gunakan format:\n"
            "INSERT INTO table_name (col1, col2) VALUES (val1, val2);"
        )

    table_name = header_match.group(1)
    raw_cols = header_match.group(2)
    columns = [c.strip().strip("`'\"") for c in raw_cols.split(",")]

    sample_rows: list[dict] = []
    values_match = re.search(r"VALUES\s+(.+)$", sql, re.IGNORECASE | re.DOTALL)
    if values_match:
        values_raw = values_match.group(1).strip()
        row_pattern = re.findall(r"\(([^)]+)\)", values_raw)
        for row_str in row_pattern:
            parts = [p.strip().strip("'\"") for p in row_str.split(",")]
            if len(parts) == len(columns):
                sample_rows.append(dict(zip(columns, parts)))

    return ParsedInsert(table_name=table_name, columns=columns, sample_rows=sample_rows)
