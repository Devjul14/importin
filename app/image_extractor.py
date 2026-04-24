from __future__ import annotations

import base64
import io
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage


@dataclass
class ExtractedImage:
    row: int          # 1-based row number in Excel (header = row 1, data starts row 2)
    col: int          # 1-based col number
    ext: str          # file extension: png, jpg, etc.
    data: bytes       # raw image bytes

    def to_base64(self) -> str:
        return base64.b64encode(self.data).decode("utf-8")

    def to_data_url(self) -> str:
        mime = "image/jpeg" if self.ext.lower() in ("jpg", "jpeg") else f"image/{self.ext.lower()}"
        return f"data:{mime};base64,{self.to_base64()}"

    def save(self, folder: str | Path, stem: str) -> Path:
        dest = Path(folder) / f"{stem}.{self.ext}"
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(self.data)
        return dest


def extract_images_from_sheet(
    file: Any,
    sheet_name: str | int | None = 0,
) -> list[ExtractedImage]:
    """
    Extract all embedded images from an Excel sheet.
    Returns list of ExtractedImage, sorted by anchor row.

    Anchor row mapping:
      - openpyxl TwoCellAnchor / OneCellAnchor → _from.row (0-based) + 1 = 1-based row
      - AbsoluteAnchor → not anchored to row, skipped with warning
    """
    if hasattr(file, "read"):
        raw = file.read()
        file.seek(0)
    else:
        raw = Path(file).read_bytes()

    wb = load_workbook(io.BytesIO(raw), data_only=True)

    if isinstance(sheet_name, int):
        ws = wb.worksheets[sheet_name]
    else:
        ws = wb[sheet_name] if sheet_name else wb.active

    images: list[ExtractedImage] = []

    for img in getattr(ws, "_images", []):
        try:
            anchor = img.anchor
            # TwoCellAnchor and OneCellAnchor have ._from
            from_cell = getattr(anchor, "_from", None) or getattr(anchor, "cell", None)
            if from_cell is None:
                continue  # AbsoluteAnchor — skip

            # openpyxl anchor rows are 0-based
            anchor_row = from_cell.row + 1   # convert to 1-based (same as Excel row numbers)
            anchor_col = from_cell.col + 1

            img_data: bytes
            if hasattr(img, "ref"):
                # openpyxl stores image bytes in img.ref for in-memory loading
                img_data = img.ref.getvalue() if hasattr(img.ref, "getvalue") else bytes(img.ref)
            elif hasattr(img, "_data"):
                img_data = img._data() if callable(img._data) else img._data
            else:
                continue

            # Detect extension from image header bytes
            ext = _detect_ext(img_data)

            images.append(ExtractedImage(row=anchor_row, col=anchor_col, ext=ext, data=img_data))
        except Exception:
            continue

    images.sort(key=lambda x: (x.row, x.col))
    return images


def map_images_to_rows(
    images: list[ExtractedImage],
    header_row: int = 1,
) -> dict[int, ExtractedImage]:
    """
    Build a dict: {data_row_index (0-based, matching DataFrame index)} -> ExtractedImage
    Excel row 2 (first data row after header) → DataFrame index 0
    """
    result: dict[int, ExtractedImage] = {}
    for img in images:
        df_index = img.row - header_row - 1  # header_row is 1-based
        if df_index >= 0:
            if df_index not in result:  # first image per row wins
                result[df_index] = img
    return result


def _detect_ext(data: bytes) -> str:
    if data[:4] == b"\x89PNG":
        return "png"
    if data[:2] in (b"\xff\xd8",):
        return "jpg"
    if data[:4] == b"GIF8":
        return "gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    return "png"  # default fallback
