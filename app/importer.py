from __future__ import annotations

from typing import Any

import pandas as pd


class ImportErrorValue(ValueError):
    pass


def read_excel(file, sheet_name: str | int | None = None) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=sheet_name)


def build_rows(df: pd.DataFrame, mapping: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    for _, source_row in df.iterrows():
        target_row: dict[str, Any] = {}
        for target_col, rule in mapping.items():
            if isinstance(rule, str):
                if rule not in df.columns:
                    raise ImportErrorValue(f"Kolom spreadsheet '{rule}' tidak ditemukan.")
                value = source_row[rule]
            elif isinstance(rule, dict):
                if "source" in rule:
                    source_col = rule["source"]
                    if source_col not in df.columns:
                        raise ImportErrorValue(
                            f"Kolom spreadsheet '{source_col}' tidak ditemukan."
                        )
                    value = source_row[source_col]
                elif "value" in rule:
                    value = rule["value"]
                else:
                    raise ImportErrorValue(
                        f"Rule untuk '{target_col}' harus punya 'source' atau 'value'."
                    )
            else:
                raise ImportErrorValue(
                    f"Rule mapping '{target_col}' harus string atau object."
                )

            if pd.isna(value):
                value = None

            target_row[target_col] = value

        rows.append(target_row)

    return rows
