from __future__ import annotations

import textwrap
from typing import Any

import pandas as pd
import streamlit as st
from sqlalchemy import text

try:
    from app.db import get_engine, insert_rows
    from app.sql_parser import SQLParseError, parse_insert_sql
except ModuleNotFoundError:
    from db import get_engine, insert_rows  # type: ignore
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


def _build_insert_preview(table: str, columns: list[str], rows: list[dict]) -> str:
    col_list = ", ".join(columns)
    lines = []
    for row in rows[:50]:
        vals = ", ".join(
            f"'{v}'" if v is not None else "NULL"
            for v in (row.get(c) for c in columns)
        )
        lines.append(f"  ({vals})")
    return f"INSERT INTO {table} ({col_list}) VALUES\n" + ",\n".join(lines) + ";"


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
                    if skipped:
                        st.warning(f"{skipped} baris dilewati (sudah ada).")
                    st.success(f"{n} baris dummy berhasil dimasukkan ke `{parsed.table_name}`.)")
                except Exception as exc:
                    st.error(f"Gagal seed: {exc}")

    st.divider()

    # ── Step 2: Upload Excel ────────────────────────────────
    st.subheader("② Upload File Excel")

    uploaded = st.file_uploader("Pilih file Excel (.xlsx / .xls)", type=["xlsx", "xls"])
    excel_df: pd.DataFrame | None = None
    excel_cols: list[str] = []

    if uploaded:
        xlsx = pd.ExcelFile(uploaded)
        selected_sheet = st.selectbox("Pilih sheet", options=xlsx.sheet_names)
        uploaded.seek(0)
        excel_df = pd.read_excel(uploaded, sheet_name=selected_sheet)
        excel_cols = list(excel_df.columns.astype(str))

        stat_col1, stat_col2, stat_col3 = st.columns(3)
        stat_col1.metric("Baris", len(excel_df))
        stat_col2.metric("Kolom", len(excel_cols))
        stat_col3.metric("Sheet", selected_sheet)
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

    st.divider()

    # ── Step 4: Import & Result ─────────────────────────────
    st.subheader("④ Import & Hasil Query")

    if excel_df is None or not any(v is not None for v in mapping.values()):
        st.info("Selesaikan step ② dan ③ untuk melanjutkan.")
        st.stop()

    active_mapping = {k: v for k, v in mapping.items() if v is not None}
    active_cols = list(active_mapping.keys())

    # Build target rows
    all_rows: list[dict[str, Any]] = []
    for _, source_row in excel_df.iterrows():
        row: dict[str, Any] = {}
        for db_col, excel_col in active_mapping.items():
            val = source_row[excel_col]
            row[db_col] = None if pd.isna(val) else val
        all_rows.append(row)

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
            if len(result_rows) > 50:
                st.caption(f"… dan {len(result_rows) - 50} baris lainnya tidak ditampilkan.")
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
            if skipped:
                st.warning(f"⚠️ {skipped} baris dilewati (duplicate/conflict).")
            st.success(f"✅ Berhasil import **{inserted} baris** ke tabel `{parsed.table_name}`.")
            st.balloons()

            res_tab1, res_tab2 = st.tabs(["📊 Data Terimport", "📝 Full INSERT Query"])
            with res_tab1:
                st.dataframe(_safe_df(pd.DataFrame(result_rows)), width="stretch")
            with res_tab2:
                st.code(
                    _build_insert_preview(parsed.table_name, active_cols, result_rows),
                    language="sql",
                )
        except Exception as exc:
            st.error(f"❌ Import gagal: {exc}")


if __name__ == "__main__":
    main()
