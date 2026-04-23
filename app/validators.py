from __future__ import annotations

import json
from typing import Any

try:
    from app.db import ColumnSpec
except ModuleNotFoundError:
    from db import ColumnSpec


class ValidationError(ValueError):
    pass


def parse_schema_json(raw: str) -> list[ColumnSpec]:
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValidationError(f"Schema JSON tidak valid: {exc}") from exc

    if not isinstance(data, list) or not data:
        raise ValidationError("Schema harus berupa array JSON dan tidak boleh kosong.")

    columns: list[ColumnSpec] = []
    for index, item in enumerate(data):
        if not isinstance(item, dict):
            raise ValidationError(f"Item schema ke-{index + 1} harus object.")

        name = item.get("name")
        if not name or not isinstance(name, str):
            raise ValidationError(f"Kolom ke-{index + 1} harus punya 'name' string.")

        columns.append(
            ColumnSpec(
                name=name,
                type=str(item.get("type", "string")),
                nullable=bool(item.get("nullable", True)),
                primary_key=bool(item.get("primary_key", False)),
                unique=bool(item.get("unique", False)),
                length=int(item["length"]) if item.get("length") is not None else None,
            )
        )

    return columns


def parse_mapping_json(raw: str) -> dict[str, Any]:
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValidationError(f"Mapping JSON tidak valid: {exc}") from exc

    if not isinstance(data, dict) or not data:
        raise ValidationError("Mapping harus object JSON dan tidak boleh kosong.")

    return data


def parse_rows_json(raw: str) -> list[dict[str, Any]]:
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValidationError(f"Data dummy JSON tidak valid: {exc}") from exc

    if not isinstance(data, list):
        raise ValidationError("Data dummy harus array JSON.")

    normalized: list[dict[str, Any]] = []
    for index, item in enumerate(data):
        if not isinstance(item, dict):
            raise ValidationError(f"Baris dummy ke-{index + 1} harus object.")
        normalized.append(item)

    return normalized
