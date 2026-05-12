from __future__ import annotations

import socket
import sys
import textwrap
import uuid
import json
from datetime import datetime
from typing import Any

# Windows does not have AF_UNIX — patch to prevent import errors from PyMySQL
if sys.platform == "win32" and not hasattr(socket, "AF_UNIX"):
    socket.AF_UNIX = None  # type: ignore[attr-defined]

import pandas as pd
import streamlit as st
from sqlalchemy import text

try:
    from app.action_log import log_action, read_action_logs
    from app.db import get_engine, insert_rows, list_tables, load_lookup_map
    from app.image_extractor import ExtractedImage, extract_images_from_sheet, map_images_to_rows
    from app.sql_parser import SQLParseError, parse_insert_sql
except ModuleNotFoundError:
    from action_log import log_action, read_action_logs  # type: ignore
    from db import get_engine, insert_rows, list_tables, load_lookup_map  # type: ignore
    from image_extractor import ExtractedImage, extract_images_from_sheet, map_images_to_rows  # type: ignore
    from sql_parser import SQLParseError, parse_insert_sql  # type: ignore


# ──────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────
SAMPLE_INSERT = textwrap.dedent("""\
    INSERT INTO customers (id, name, email, phone, city)
    VALUES
      (1, 'Alice', 'alice@example.com', '08123456789', 'Jakarta'),
      (2, 'Bob',   'bob@example.com',   '08987654321', 'Bandung');
""")

NONE_LABEL = "— skip —"


# ──────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────
@st.cache_resource
def get_cached_engine(db_url: str):
    return get_engine(db_url)


def _safe_df(df: pd.DataFrame) -> pd.DataFrame:
    """Convert all columns to string for safe Arrow/Streamlit display."""
    return df.astype(str)


def resolve_autogen(pattern: str, row_index: int, ts: datetime) -> str:
    """
    Resolve auto-generate pattern ke nilai string.

    Token yang didukung:
      {Y}       → tahun 4 digit          2026
      {m}       → bulan 2 digit          04
      {d}       → hari 2 digit           23
      {H}       → jam 2 digit            14
      {i}       → menit 2 digit          30
      {s}       → detik 2 digit          05
      {YmdHis}  → YYYYmmddHHMMSS        20260423143005
      {uuid}    → UUID4 tanpa dash       a1b2c3d4...
      {n}       → nomor baris (1-based)  1, 2, 3 ...
      {n:03}    → nomor baris padded     001, 002, 003 ...
    """
    tokens = {
        "{Y}":       ts.strftime("%Y"),
        "{m}":       ts.strftime("%m"),
        "{d}":       ts.strftime("%d"),
        "{H}":       ts.strftime("%H"),
        "{i}":       ts.strftime("%M"),
        "{s}":       ts.strftime("%S"),
        "{YmdHis}": ts.strftime("%Y%m%d%H%M%S"),
        "{uuid}":    uuid.uuid4().hex,
        "{n}":       str(row_index),
    }
    result = pattern
    for token, value in tokens.items():
        result = result.replace(token, value)

    # Handle {n:XX} zero-pad format e.g. {n:03}
    import re
    result = re.sub(
        r"\{n:(\d+)\}",
        lambda m: str(row_index).zfill(int(m.group(1))),
        result,
    )
    return result


def _build_insert_preview(
    table: str,
    columns: list[str],
    rows: list[dict],
    max_rows: int | None = 50,
) -> str:
    col_list = ", ".join(columns)
    subset = rows if max_rows is None else rows[:max_rows]
    lines = []
    for row in subset:
        vals = ", ".join(
            f"'{v}'" if v is not None else "NULL"
            for v in (row.get(c) for c in columns)
        )
        lines.append(f"  ({vals})")
    sql = f"INSERT INTO {table} ({col_list}) VALUES\n" + ",\n".join(lines) + ";"
    if max_rows is not None and len(rows) > max_rows:
        sql += f"\n-- ... dan {len(rows) - max_rows} baris lainnya (total {len(rows)} baris)"
    return sql


# ──────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────
def main() -> None:
    st.set_page_config(page_title="Excel → DB Import", layout="wide", page_icon="📥")

    # ── Sidebar: DB connection ──────────────────────────────
    with st.sidebar:
        st.header("🔌 Database Connection")
        db_url = st.text_input(
            "SQLAlchemy URL",
            value="sqlite:///import_tool.db",
            help=(
                "SQLite  : sqlite:///file.db\n"
                "MySQL   : mysql+pymysql://user:pass@host:3306/dbname\n"
                "Postgres: postgresql+psycopg2://user:pass@host:5432/dbname"
            ),
        )
        engine = None
        try:
            engine = get_cached_engine(db_url)
            with engine.connect() as conn:
                conn.execute(text("SELECT 1"))
            st.success("Connected ✅")
        except Exception as exc:
            st.error(f"Koneksi gagal: {exc}")
            st.stop()

        st.divider()
        with st.expander("🐟 Pakai MySQL lokal?"):
            st.markdown(
                "**Gunakan `localhost` (bukan `127.0.0.1`) agar koneksi pakai Unix socket:**\n\n"
                "```\n"
                "# Tanpa password:\n"
                "mysql+pymysql://root@localhost/nama_db\n\n"
                "# Dengan password:\n"
                "mysql+pymysql://root:password@localhost/nama_db\n\n"
                "# Eksplisit via socket:\n"
                "mysql+pymysql://root@localhost/nama_db?unix_socket=/tmp/mysql.sock\n"
                "```\n\n"
                "⚠️ `127.0.0.1` memaksa TCP — gunakan `localhost` untuk socket lokal.\n\n"
                "Pastikan DB sudah dibuat dan user punya hak `CREATE`, `INSERT`, `SELECT`."
            )
        st.caption("Stack: Streamlit · SQLAlchemy · Pandas · PyMySQL")
        st.divider()
        st.caption("🧹 Gunakan menu **Cleansing Data** di sidebar untuk membersihkan Excel sebelum import.")

    st.title("📥 Excel → Database Import Tool")

    # ── Step 1: SQL INSERT ──────────────────────────────────
    st.subheader("① SQL INSERT — Struktur & Data Dummy/Referensi")
    st.caption(
        "Paste SQL INSERT untuk mendefinisikan **nama table** dan **urutan kolom** "
        "target DB. Data VALUES (opsional) bisa langsung di-seed ke DB sebagai referensi."
    )

    sql_input = st.text_area(
        "SQL INSERT statement",
        value=SAMPLE_INSERT,
        height=200,
        placeholder="INSERT INTO table (col1, col2) VALUES (v1, v2);",
    )

    parsed = None
    if sql_input.strip():
        try:
            parsed = parse_insert_sql(sql_input)
        except SQLParseError as exc:
            st.error(str(exc))

    if parsed:
        info_col, ref_col = st.columns([1, 2])
        with info_col:
            st.markdown(f"**Table target:** `{parsed.table_name}`")
            st.markdown("**Kolom terdeteksi:**")
            for i, c in enumerate(parsed.columns, 1):
                st.markdown(f"`{i}.` `{c}`")
        with ref_col:
            if parsed.sample_rows:
                st.markdown("**Preview data referensi dari SQL:**")
                st.dataframe(_safe_df(pd.DataFrame(parsed.sample_rows)), width="stretch")
            else:
                st.info("Tidak ada baris VALUES — hanya definisi kolom yang digunakan.")

        if parsed.sample_rows:
            if st.button("💾 Seed dummy data ke DB", help="Insert baris dari SQL ke database (opsional)."):
                try:
                    n, skipped = insert_rows(engine, parsed.table_name, parsed.sample_rows, conflict_strategy="skip")
                    log_action(
                        "seed_dummy_data",
                        status="success",
                        payload={"table": parsed.table_name, "inserted": n, "skipped": skipped},
                    )
                    if skipped:
                        st.warning(f"{skipped} baris dilewati (sudah ada).")
                    st.success(f"{n} baris dummy berhasil dimasukkan ke `{parsed.table_name}`.)")
                except Exception as exc:
                    log_action(
                        "seed_dummy_data",
                        status="error",
                        payload={"table": parsed.table_name, "error": str(exc)},
                    )
                    st.error(f"Gagal seed: {exc}")

    st.divider()

    # ── Step 2: Upload Excel ────────────────────────────────
    st.subheader("② Upload File Excel")

    uploaded = st.file_uploader("Pilih file Excel (.xlsx / .xls)", type=["xlsx", "xls"])
    excel_df: pd.DataFrame | None = None
    excel_cols: list[str] = []

    # image_rows cached in session_state to avoid re-processing on every rerun
    image_rows: dict[int, ExtractedImage] = {}

    if uploaded:
        xlsx = pd.ExcelFile(uploaded)
        selected_sheet = st.selectbox("Pilih sheet", options=xlsx.sheet_names)

        # Cache key: file identity + sheet — only re-process when file or sheet changes
        cache_key = (uploaded.name, uploaded.size, selected_sheet)
        if st.session_state.get("_excel_cache_key") != cache_key:
            uploaded.seek(0)
            excel_df_raw = pd.read_excel(uploaded, sheet_name=selected_sheet)
            # Extract embedded images (expensive — only run once per file+sheet)
            try:
                uploaded.seek(0)
                sheet_idx = xlsx.sheet_names.index(selected_sheet)
                extracted = extract_images_from_sheet(uploaded, sheet_name=sheet_idx)
                img_rows_raw = map_images_to_rows(extracted, header_row=1)
            except Exception:
                img_rows_raw = {}
            st.session_state["_excel_cache_key"] = cache_key
            st.session_state["_excel_df"] = excel_df_raw
            st.session_state["_image_rows"] = img_rows_raw
            log_action(
                "upload_excel",
                status="success",
                payload={
                    "file": uploaded.name,
                    "sheet": selected_sheet,
                    "rows": len(excel_df_raw),
                    "columns": len(excel_df_raw.columns),
                    "images": len(img_rows_raw),
                },
            )

        excel_df = st.session_state.get("_excel_df")
        image_rows = st.session_state.get("_image_rows", {})
        excel_cols = list(excel_df.columns.astype(str)) if excel_df is not None else []

        stat_col1, stat_col2, stat_col3, stat_col4 = st.columns(4)
        stat_col1.metric("Baris", len(excel_df) if excel_df is not None else 0)
        stat_col2.metric("Kolom", len(excel_cols))
        stat_col3.metric("Sheet", selected_sheet)
        stat_col4.metric("Gambar terdeteksi", len(image_rows))
        if excel_df is not None:
            st.dataframe(_safe_df(excel_df.head(5)), width="stretch")

    st.divider()

    # ── Step 3: Mapping UI ──────────────────────────────────
    st.subheader("③ Mapping Kolom Excel → Kolom DB")

    if not parsed:
        st.info("⬆ Isi SQL INSERT di step ① terlebih dahulu.")
        st.stop()

    mapping: dict[str, str | None] = {}

    if not excel_cols:
        st.info("⬆ Upload file Excel di step ② terlebih dahulu.")
        unique_cols: list[str] = []
        autogen_patterns: dict[str, str] = {}
        relation_configs: dict[str, dict] = {}
        image_config: dict = {}
    else:
        st.caption(
            "Setiap baris = 1 kolom DB. Pilih kolom Excel yang sesuai dari dropdown. "
            f"Pilih **{NONE_LABEL}** untuk melewati kolom tersebut."
        )

        options = [NONE_LABEL] + excel_cols
        lower_excel = [c.lower() for c in excel_cols]

        # Render grid: 3 kolom per baris
        db_cols = parsed.columns
        rows_of_cols = [db_cols[i:i+3] for i in range(0, len(db_cols), 3)]

        for row_group in rows_of_cols:
            grid = st.columns(3)
            for col_idx, db_col in enumerate(row_group):
                default_idx = 0
                if db_col.lower() in lower_excel:
                    default_idx = lower_excel.index(db_col.lower()) + 1
                with grid[col_idx]:
                    chosen = st.selectbox(
                        f"🗂 `{db_col}`",
                        options=options,
                        index=default_idx,
                        key=f"map_{db_col}",
                    )
                    mapping[db_col] = None if chosen == NONE_LABEL else chosen

        active = {k: v for k, v in mapping.items() if v is not None}
        st.caption(f"✅ {len(active)} kolom terpetakan · ⏭ {len(mapping) - len(active)} dilewati")

        st.markdown("**🚫 Kolom yang tidak boleh duplikat (unique check)**")
        st.caption(
            "Pilih kolom DB yang nilainya harus unik dalam file Excel. "
            "Baris duplikat akan dideteksi dan dipisahkan sebelum diimport."
        )
        unique_cols: list[str] = st.multiselect(
            "Pilih kolom unique",
            options=list(active.keys()),
            default=[],
            placeholder="Kosongkan jika tidak ada pengecekan duplikat...",
            key="unique_cols",
        )

        # ── Auto-generate section ───────────────────────────
        skipped_cols = [c for c in parsed.columns if mapping.get(c) is None]
        st.markdown("**⚙️ Auto-generate nilai kolom**")
        st.caption(
            "Isi pola nilai otomatis untuk kolom yang tidak diambil dari Excel. "
            "Didukung: `{YmdHis}` `{Y}` `{m}` `{d}` `{H}` `{i}` `{s}` `{uuid}` `{n}` `{n:03}`"
        )

        autogen_patterns: dict[str, str] = {}
        all_db_cols = parsed.columns
        ag_rows = [all_db_cols[i:i+2] for i in range(0, len(all_db_cols), 2)]
        for ag_row in ag_rows:
            ag_grid = st.columns(2)
            for ag_idx, db_col in enumerate(ag_row):
                with ag_grid[ag_idx]:
                    is_skipped = mapping.get(db_col) is None
                    pattern = st.text_input(
                        f"{'🔁' if is_skipped else '✏️'} `{db_col}`"
                        + (" _(tidak dari Excel)_" if is_skipped else " _(override Excel)_"),
                        value="",
                        placeholder=f"e.g. SMP-{{YmdHis}}-{{n:03}}" if is_skipped else "",
                        key=f"autogen_{db_col}",
                    )
                    if pattern.strip():
                        autogen_patterns[db_col] = pattern.strip()

        if autogen_patterns:
            st.caption(
                "Preview nilai (baris ke-1): "
                + " · ".join(
                    f"`{c}` → `{resolve_autogen(p, 1, datetime.now())}`"
                    for c, p in autogen_patterns.items()
                )
            )

        # ── Relation / Foreign Key section ──────────────────
        all_tables = list_tables(engine)
        st.markdown("**🔗 Relasi Table (Foreign Key Lookup)**")
        st.caption(
            "Untuk kolom yang nilainya harus di-lookup ke table master, "
            "atur relasi di sini. Nilai dari Excel akan diganti dengan ID dari table master."
        )

        relation_configs: dict[str, dict] = {}
        active_mapped_cols = [c for c in parsed.columns if mapping.get(c) is not None]

        if not active_mapped_cols:
            st.info("Mapping kolom terlebih dahulu untuk mengatur relasi.")
        elif not all_tables:
            st.info("Belum ada table di database untuk dijadikan master.")
        else:
            for db_col in active_mapped_cols:
                with st.expander(f"🔗 Relasi untuk kolom `{db_col}`", expanded=False):
                    enable_rel = st.checkbox(
                        "Aktifkan lookup ke table master",
                        key=f"rel_enable_{db_col}",
                    )
                    if enable_rel:
                        r_col1, r_col2, r_col3 = st.columns(3)
                        with r_col1:
                            master_tbl = st.selectbox(
                                "Table master",
                                options=all_tables,
                                key=f"rel_tbl_{db_col}",
                            )
                        master_cols = []
                        if master_tbl:
                            try:
                                from app.db import get_table_columns as _gtc
                            except ModuleNotFoundError:
                                from db import get_table_columns as _gtc  # type: ignore
                            master_cols = _gtc(engine, master_tbl)
                        with r_col2:
                            match_col = st.selectbox(
                                "Kolom pencarian (match)",
                                options=master_cols,
                                help="Kolom di table master yang nilainya sama dengan data Excel",
                                key=f"rel_match_{db_col}",
                            )
                        with r_col3:
                            return_col = st.selectbox(
                                "Kolom return (ambil nilai ini)",
                                options=master_cols,
                                help="Kolom yang nilainya akan dimasukkan ke DB (biasanya id)",
                                key=f"rel_return_{db_col}",
                            )

                        if master_tbl and match_col and return_col:
                            relation_configs[db_col] = {
                                "master_table": master_tbl,
                                "match_col": match_col,
                                "return_col": return_col,
                            }
                            # Preview sample lookup
                            try:
                                lmap = load_lookup_map(engine, master_tbl, match_col, return_col)
                                sample = dict(list(lmap.items())[:5])
                                st.caption(
                                    f"Preview lookup `{master_tbl}` ({len(lmap)} entri): "
                                    + " · ".join(f"`{k}` → `{v}`" for k, v in sample.items())
                                    + (" …" if len(lmap) > 5 else "")
                                )
                            except Exception as exc:
                                st.warning(f"Tidak bisa load preview: {exc}")

        # ── Image Import section ────────────────────────────
        st.markdown("**🖼️ Import Gambar dari Excel**")

        image_config: dict = {}

        if not image_rows:
            st.caption("Tidak ada gambar terdeteksi di file Excel ini.")
        else:
            st.caption(
                f"{len(image_rows)} gambar terdeteksi di file Excel. "
                "Tentukan kolom DB tujuan dan format penyimpanan."
            )
            img_col1, img_col2, img_col3 = st.columns(3)
            with img_col1:
                img_enable = st.checkbox("Aktifkan import gambar", key="img_enable")
            if img_enable:
                with img_col2:
                    img_db_col = st.selectbox(
                        "Kolom DB tujuan gambar",
                        options=parsed.columns,
                        key="img_db_col",
                        help="Kolom yang akan menyimpan path/base64/binary gambar",
                    )
                with img_col3:
                    img_format = st.selectbox(
                        "Format simpan",
                        options=["base64", "path", "binary"],
                        format_func=lambda x: {
                            "base64": "Base64 string (TEXT/LONGTEXT)",
                            "path": "File path (simpan ke disk)",
                            "binary": "Binary / BLOB",
                        }[x],
                        key="img_format",
                    )

                name_col1, name_col2, name_col3 = st.columns(3)
                with name_col1:
                    img_naming_col = st.selectbox(
                        "📛 Nama file dari kolom",
                        options=[NONE_LABEL] + excel_cols,
                        index=0,
                        key="img_naming_col",
                        help="Nilai kolom ini dipakai sebagai nama file gambar. "
                             "Kosongkan untuk pakai nomor baris.",
                    )
                with name_col2:
                    img_prefix = st.text_input(
                        "Prefix nama file",
                        value="",
                        placeholder="e.g. product_",
                        key="img_prefix",
                    )
                with name_col3:
                    img_suffix = st.text_input(
                        "Suffix nama file",
                        value="",
                        placeholder="e.g. _foto",
                        key="img_suffix",
                    )

                if img_format == "path":
                    img_save_dir = st.text_input(
                        "Folder simpan gambar",
                        value="uploads/images",
                        key="img_save_dir",
                        help="Relatif dari direktori project. Akan dibuat otomatis.",
                    )
                else:
                    img_save_dir = "uploads/images"

                image_config = {
                    "enabled": True,
                    "db_col": img_db_col,
                    "format": img_format,
                    "save_dir": img_save_dir,
                    "naming_col": None if img_naming_col == NONE_LABEL else img_naming_col,
                    "prefix": img_prefix.strip(),
                    "suffix": img_suffix.strip(),
                }

                # Show image preview grid with naming preview
                sample_indices = sorted(image_rows.keys())[:6]
                if sample_indices:
                    st.caption("Preview gambar (maks 6 baris pertama yang punya gambar):")
                    preview_cols = st.columns(min(6, len(sample_indices)))
                    for pi, df_idx in enumerate(sample_indices):
                        img_obj = image_rows[df_idx]
                        with preview_cols[pi]:
                            # Build preview filename
                            if image_config["naming_col"] and excel_df is not None:
                                raw_name = str(excel_df.iloc[df_idx].get(image_config["naming_col"], df_idx + 2))
                            else:
                                raw_name = str(df_idx + 2)
                            preview_name = f"{image_config['prefix']}{raw_name}{image_config['suffix']}.{img_obj.ext}"
                            st.image(img_obj.data, use_container_width=True)
                            st.caption(f"`{preview_name}`")

    st.divider()

    # ── Step 4: Import & Result ─────────────────────────────
    st.subheader("④ Import & Hasil Query")

    has_autogen = bool(autogen_patterns) if excel_cols else False
    if excel_df is None or (not any(v is not None for v in mapping.values()) and not has_autogen):
        st.info("Selesaikan step ② dan ③ untuk melanjutkan.")
        st.stop()

    active_mapping = {k: v for k, v in mapping.items() if v is not None}
    # Merge autogen + image cols into active_cols for preview
    img_db_col_active = image_config.get("db_col") if image_config.get("enabled") else None
    extra_cols = list(autogen_patterns.keys() if excel_cols else [])
    if img_db_col_active:
        extra_cols.append(img_db_col_active)
    active_cols = list(dict.fromkeys(list(active_mapping.keys()) + extra_cols))

    # Pre-load all relation lookup maps (bulk, once before row loop)
    lookup_maps: dict[str, dict] = {}
    rel_errors: list[str] = []
    for db_col, rel in (relation_configs.items() if excel_cols else {}.items()):
        try:
            lookup_maps[db_col] = load_lookup_map(
                engine, rel["master_table"], rel["match_col"], rel["return_col"]
            )
        except Exception as exc:
            rel_errors.append(f"Lookup `{db_col}` gagal: {exc}")
    for err in rel_errors:
        st.warning(err)

    # Build target rows
    import_ts = datetime.now()  # single timestamp for entire import batch
    all_rows: list[dict[str, Any]] = []
    unresolved_lookup: list[tuple[int, str, Any]] = []  # (row_idx, db_col, raw_val)
    insert_seq = 0  # sequential counter for autogen {n}, only increments for non-skipped rows

    for row_idx, (_, source_row) in enumerate(excel_df.iterrows(), start=1):
        # Skip rows where ALL mapped Excel columns are empty / NaN
        if active_mapping and all(
            pd.isna(source_row[excel_col])
            for excel_col in active_mapping.values()
        ):
            continue

        insert_seq += 1
        row: dict[str, Any] = {}
        # 1. Values from Excel mapping
        for db_col, excel_col in active_mapping.items():
            val = source_row[excel_col]
            raw = None if pd.isna(val) else val
            # Apply relation lookup if configured
            if db_col in lookup_maps:
                looked_up = lookup_maps[db_col].get(str(raw))
                if looked_up is None:
                    unresolved_lookup.append((row_idx, db_col, raw))
                row[db_col] = looked_up
            else:
                row[db_col] = raw
        # 2. Auto-generated values
        for db_col, pattern in (autogen_patterns.items() if excel_cols else {}.items()):
            row[db_col] = resolve_autogen(pattern, insert_seq, import_ts)
        # 3. Embedded image from Excel
        if image_config.get("enabled") and image_rows:
            df_index = row_idx - 1  # row_idx is 1-based, df_index is 0-based
            img_obj = image_rows.get(df_index)
            img_col_name = image_config["db_col"]
            img_fmt = image_config["format"]
            prefix = image_config.get("prefix", "")
            suffix = image_config.get("suffix", "")
            naming_col = image_config.get("naming_col")

            if img_obj:
                # Build file stem from naming column or fallback to row number
                if naming_col and naming_col in source_row.index:
                    raw_name = str(source_row[naming_col]).strip()
                    # Sanitize: replace characters not safe for filenames
                    import re as _re
                    raw_name = _re.sub(r'[^\w\-.]', '_', raw_name)
                else:
                    raw_name = str(row_idx)
                stem = f"{prefix}{raw_name}{suffix}"

                if img_fmt == "base64":
                    row[img_col_name] = img_obj.to_base64()
                elif img_fmt == "binary":
                    row[img_col_name] = img_obj.data
                elif img_fmt == "path":
                    from pathlib import Path as _Path
                    saved = img_obj.save(image_config["save_dir"], stem)
                    row[img_col_name] = str(saved)
            else:
                row[img_col_name] = None
        all_rows.append(row)

    if unresolved_lookup:
        with st.expander(f"⚠️ {len(unresolved_lookup)} nilai tidak ditemukan di table master", expanded=True):
            st.caption("Baris berikut tidak memiliki pasangan di table master — nilai akan `NULL`.")
            udf = pd.DataFrame(unresolved_lookup, columns=["Baris (Excel)", "Kolom DB", "Nilai tidak ditemukan"])
            st.dataframe(_safe_df(udf), width="stretch")

    # ── Duplicate check ─────────────────────────────────────
    result_rows = all_rows
    dup_rows: list[dict[str, Any]] = []

    if unique_cols:
        seen: set[tuple] = set()
        clean: list[dict[str, Any]] = []
        dups: list[dict[str, Any]] = []
        for r in all_rows:
            key = tuple(str(r.get(c)) for c in unique_cols)
            if key in seen:
                dups.append(r)
            else:
                seen.add(key)
                clean.append(r)
        result_rows = clean
        dup_rows = dups

        if dup_rows:
            st.warning(
                f"⚠️ **{len(dup_rows)} baris duplikat ditemukan** berdasarkan kolom: "
                + ", ".join(f"`{c}`" for c in unique_cols)
                + f" — akan **dilewati**. {len(result_rows)} baris akan diimport."
            )
            with st.expander(f"🔎 Lihat {len(dup_rows)} baris duplikat"):
                st.dataframe(_safe_df(pd.DataFrame(dup_rows)), width="stretch")
        else:
            st.success(f"✅ Tidak ada duplikat pada kolom: " + ", ".join(f"`{c}`" for c in unique_cols))

    prev_col, data_col = st.columns([3, 2])
    with prev_col:
        with st.expander("🔍 Preview INSERT query (maks 50 baris)", expanded=True):
            st.code(
                _build_insert_preview(parsed.table_name, active_cols, result_rows),
                language="sql",
            )
    with data_col:
        with st.expander("📋 Preview data (10 baris pertama)", expanded=True):
            st.dataframe(_safe_df(pd.DataFrame(result_rows).head(10)), width="stretch")

    conflict_strategy = st.radio(
        "Jika baris sudah ada di DB (duplicate/conflict):",
        options=["skip", "replace", "error"],
        format_func=lambda x: {
            "skip": "⏭ Skip — lewati baris yang konflik",
            "replace": "♻️ Replace — timpa baris yang konflik",
            "error": "🚫 Error — hentikan import jika ada konflik",
        }[x],
        horizontal=True,
        index=0,
    )

    import_btn = st.button(
        f"🚀 Import {len(result_rows)} baris ke `{parsed.table_name}`",
        type="primary",
        use_container_width=True,
    )

    if import_btn:
        try:
            inserted, skipped = insert_rows(
                engine, parsed.table_name, result_rows, conflict_strategy=conflict_strategy
            )
            log_action(
                "import_rows",
                status="success",
                payload={
                    "table": parsed.table_name,
                    "sent": len(result_rows),
                    "inserted": inserted,
                    "skipped": skipped,
                    "strategy": conflict_strategy,
                },
            )
            if skipped:
                st.warning(f"⚠️ {skipped} baris dilewati (duplicate/conflict).")
            st.success(f"✅ Berhasil import **{inserted} baris** ke tabel `{parsed.table_name}`.")
            st.balloons()

            # Rows yang benar-benar masuk ke DB
            confirmed_rows = result_rows[:inserted] if conflict_strategy == "skip" else result_rows

            res_tab1, res_tab2 = st.tabs(["📊 Data Terimport", "📝 Full INSERT Query"])
            with res_tab1:
                st.caption(f"Menampilkan {len(confirmed_rows)} dari {len(result_rows)} baris yang dikirim.")
                st.dataframe(_safe_df(pd.DataFrame(confirmed_rows)), width="stretch")
            with res_tab2:
                st.caption(f"Full INSERT query — {len(confirmed_rows)} baris.")
                st.code(
                    _build_insert_preview(parsed.table_name, active_cols, confirmed_rows, max_rows=None),
                    language="sql",
                )
        except Exception as exc:
            log_action(
                "import_rows",
                status="error",
                payload={
                    "table": parsed.table_name,
                    "sent": len(result_rows),
                    "strategy": conflict_strategy,
                    "error": str(exc),
                },
            )
            st.error(f"❌ Import gagal: {exc}")

    st.divider()
    st.subheader("🧾 Log Action")
    st.caption("Menampilkan 50 action terbaru dari file log aplikasi.")
    logs = read_action_logs(limit=50)
    if logs:
        log_df = pd.DataFrame(logs)
        if "payload" in log_df.columns:
            log_df["payload"] = log_df["payload"].apply(
                lambda p: json.dumps(p, ensure_ascii=False) if isinstance(p, dict) else str(p)
            )
        st.dataframe(_safe_df(log_df.iloc[::-1]), width="stretch")
    else:
        st.info("Belum ada log action.")

    # ── Reset / Clear Cache ─────────────────────────────────
    st.divider()
    with st.expander("🗑️ Reset & Bersihkan Memori"):
        st.caption(
            "Hapus data Excel dan gambar dari memori session. "
            "Lakukan setelah import selesai untuk membebaskan RAM."
        )
        mem_cols = ["_excel_df", "_image_rows", "_excel_cache_key"]
        cached_keys = [k for k in mem_cols if k in st.session_state]
        if cached_keys:
            img_count = len(st.session_state.get("_image_rows", {}))
            df_rows = len(st.session_state.get("_excel_df", pd.DataFrame()))
            st.info(f"Cache aktif: {df_rows} baris Excel + {img_count} gambar di memori.")
        else:
            st.info("Tidak ada data di cache session.")
        if st.button("🗑️ Hapus cache Excel & gambar", key="btn_clear_cache"):
            for k in mem_cols:
                st.session_state.pop(k, None)
            log_action("clear_cache", status="success", payload={"keys": mem_cols})
            st.success("Cache dibersihkan. Upload file baru untuk memulai lagi.")
            st.rerun()


if __name__ == "__main__":
    main()
