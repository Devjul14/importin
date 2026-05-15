from __future__ import annotations

import socket
import sys

if sys.platform == "win32" and not hasattr(socket, "AF_UNIX"):
    socket.AF_UNIX = None  # type: ignore[attr-defined]

import pandas as pd
import streamlit as st
from sqlalchemy import text

try:
    from app.db import get_engine, get_table_columns, list_tables, bulk_update_by_join
    from app.action_log import log_action
except ModuleNotFoundError:
    from db import get_engine, get_table_columns, list_tables, bulk_update_by_join  # type: ignore
    from action_log import log_action  # type: ignore

NONE_LABEL = "— skip —"


# ──────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────
@st.cache_resource
def _get_engine(db_url: str):
    return get_engine(db_url)


def _safe(df: pd.DataFrame) -> pd.DataFrame:
    return df.astype(str)


# ──────────────────────────────────────────────────────────────
# Page
# ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="BOM Update by Join", layout="wide", page_icon="🔄")

# ── Sidebar ────────────────────────────────────────────────────
with st.sidebar:
    st.header("🔌 Database Connection")
    db_url = st.text_input(
        "SQLAlchemy URL",
        value="sqlite:///import_tool.db",
        help=(
            "SQLite  : sqlite:///file.db\n"
            "MySQL   : mysql+pymysql://user:pass@host:3306/dbname\n"
        ),
        key="bom_db_url",
    )
    engine = None
    try:
        engine = _get_engine(db_url)
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        st.success("Connected ✅")
    except Exception as exc:
        st.error(f"Koneksi gagal: {exc}")
        st.stop()
    st.caption("Update baris DB menggunakan kolom join dari Excel.")

st.title("🔄 Update BOM by Nama Produk")
st.caption(
    "Upload Excel hasil **unpivot** (nama_produk | kode_bahan_baku | qty_per_pcs), "
    "lalu update kolom di database berdasarkan kecocokan **nama_produk**."
)

# ── Step 1: Target table & join config ────────────────────────
st.subheader("① Konfigurasi Table & Join")

all_tables = list_tables(engine)
if not all_tables:
    st.error("Tidak ada table di database. Pastikan koneksi sudah benar.")
    st.stop()

col1, col2 = st.columns(2)
with col1:
    target_table = st.selectbox("Table target (yang akan di-UPDATE)", options=all_tables, key="upd_table")
with col2:
    table_cols = get_table_columns(engine, target_table) if target_table else []
    join_col_db = st.selectbox(
        "Kolom JOIN di DB (key matching)",
        options=table_cols,
        index=table_cols.index("nama_produk") if "nama_produk" in table_cols else 0,
        help="Kolom di database yang nilainya cocok dengan kolom Excel (biasanya nama_produk).",
        key="upd_join_col",
    )

update_cols_db: list[str] = st.multiselect(
    "Kolom yang akan di-UPDATE",
    options=[c for c in table_cols if c != join_col_db],
    default=[c for c in ["kode_bahan_baku", "qty_per_pcs"] if c in table_cols],
    help="Pilih kolom mana saja yang nilainya akan diupdate dari Excel.",
    key="upd_cols",
)

if not update_cols_db:
    st.info("Pilih minimal satu kolom untuk di-update.")
    st.stop()

st.divider()

# ── Step 2: Upload Excel ───────────────────────────────────────
st.subheader("② Upload File Excel")

uploaded = st.file_uploader("Pilih file Excel (.xlsx / .xls)", type=["xlsx", "xls"], key="upd_file")
excel_df: pd.DataFrame | None = None

if uploaded:
    xlsx = pd.ExcelFile(uploaded)
    selected_sheet = st.selectbox("Pilih sheet", options=xlsx.sheet_names, key="upd_sheet")

    cache_key = (uploaded.name, uploaded.size, selected_sheet)
    if st.session_state.get("_bom_cache_key") != cache_key:
        uploaded.seek(0)
        st.session_state["_bom_df"] = pd.read_excel(uploaded, sheet_name=selected_sheet)
        st.session_state["_bom_cache_key"] = cache_key

    excel_df = st.session_state.get("_bom_df")

    if excel_df is not None:
        c1, c2 = st.columns(2)
        c1.metric("Baris", len(excel_df))
        c2.metric("Kolom", len(excel_df.columns))
        st.dataframe(_safe(excel_df.head(5)), use_container_width=True)

st.divider()

# ── Step 3: Column Mapping ─────────────────────────────────────
st.subheader("③ Mapping Kolom Excel → Kolom DB")

if excel_df is None:
    st.info("Upload file Excel di step ② terlebih dahulu.")
    st.stop()

excel_cols = list(excel_df.columns.astype(str))
options = [NONE_LABEL] + excel_cols
lower_excel = [c.lower() for c in excel_cols]

def _default_idx(db_col: str) -> int:
    if db_col.lower() in lower_excel:
        return lower_excel.index(db_col.lower()) + 1
    return 0

all_map_cols = [join_col_db] + update_cols_db
mapping: dict[str, str | None] = {}

grid_rows = [all_map_cols[i:i+3] for i in range(0, len(all_map_cols), 3)]
for grp in grid_rows:
    cols = st.columns(3)
    for idx, db_col in enumerate(grp):
        with cols[idx]:
            label = f"🔑 `{db_col}` _(join key)_" if db_col == join_col_db else f"✏️ `{db_col}`"
            chosen = st.selectbox(label, options=options, index=_default_idx(db_col), key=f"bom_map_{db_col}")
            mapping[db_col] = None if chosen == NONE_LABEL else chosen

# Validate join key is mapped
if mapping.get(join_col_db) is None:
    st.error(f"Kolom join `{join_col_db}` harus dipetakan ke kolom Excel.")
    st.stop()

mapped_update_cols = [c for c in update_cols_db if mapping.get(c) is not None]
if not mapped_update_cols:
    st.warning("Pilih setidaknya satu kolom update yang dipetakan ke Excel.")
    st.stop()

st.caption(
    f"Join key: `{join_col_db}` ← `{mapping[join_col_db]}`  |  "
    f"Update: " + ", ".join(f"`{c}` ← `{mapping[c]}`" for c in mapped_update_cols)
)

st.divider()

# ── Step 4: Preview & Execute ──────────────────────────────────
st.subheader("④ Preview & Eksekusi Update")

# Build rows dict for update
join_excel_col = mapping[join_col_db]
rows_to_update: list[dict] = []

for _, src in excel_df.iterrows():
    join_val = src[join_excel_col]
    if pd.isna(join_val) or str(join_val).strip() == "":
        continue
    row: dict = {join_col_db: str(join_val).strip()}
    for db_col in mapped_update_cols:
        excel_col = mapping[db_col]
        val = src[excel_col]
        row[db_col] = None if pd.isna(val) else val
    rows_to_update.append(row)

st.info(f"**{len(rows_to_update)}** baris dari Excel siap di-update ke `{target_table}` via `{join_col_db}`.")

# Preview table
preview_df = pd.DataFrame(rows_to_update)
with st.expander("🔍 Preview data yang akan diupdate (maks 50 baris)", expanded=True):
    st.dataframe(_safe(preview_df.head(50)), use_container_width=True)

# Show sample UPDATE SQL
if rows_to_update:
    sample = rows_to_update[0]
    set_part = ", ".join(f"{c} = '{sample.get(c)}'" for c in mapped_update_cols)
    sample_sql = (
        f"UPDATE {target_table}\n"
        f"   SET {set_part}\n"
        f" WHERE {join_col_db} = '{sample[join_col_db]}';\n"
        f"-- ... ({len(rows_to_update)} total statements)"
    )
    with st.expander("📝 Contoh SQL UPDATE"):
        st.code(sample_sql, language="sql")

# Execute
update_btn = st.button(
    f"🚀 Eksekusi UPDATE {len(rows_to_update)} baris ke `{target_table}`",
    type="primary",
    use_container_width=True,
    key="upd_execute",
)

if update_btn:
    if not rows_to_update:
        st.error("Tidak ada data untuk diupdate.")
    else:
        try:
            updated, not_found = bulk_update_by_join(
                engine,
                table_name=target_table,
                rows=rows_to_update,
                join_col=join_col_db,
                update_cols=mapped_update_cols,
            )
            log_action(
                "bom_update_by_join",
                status="success",
                payload={
                    "table": target_table,
                    "join_col": join_col_db,
                    "update_cols": mapped_update_cols,
                    "sent": len(rows_to_update),
                    "updated": updated,
                    "not_found": not_found,
                },
            )
            st.success(f"✅ **{updated} baris** berhasil diupdate di `{target_table}`.")
            if not_found:
                st.warning(
                    f"⚠️ **{not_found} baris** tidak ditemukan di DB "
                    f"(tidak ada {join_col_db} yang cocok)."
                )
            st.balloons()

            # Show summary of not-found rows
            if not_found:
                # Find which join values did not match
                updated_vals = set()
                # Re-query to find what is in DB
                try:
                    from sqlalchemy import MetaData, Table, select as sa_select
                    meta = MetaData()
                    tbl = Table(target_table, meta, autoload_with=engine)
                    with engine.connect() as conn:
                        existing = {
                            str(r[0])
                            for r in conn.execute(sa_select(tbl.c[join_col_db])).fetchall()
                        }
                    missing_rows = [r for r in rows_to_update if str(r[join_col_db]) not in existing]
                    if missing_rows:
                        with st.expander(f"🔎 {len(missing_rows)} nilai tidak ditemukan di DB"):
                            st.dataframe(_safe(pd.DataFrame(missing_rows)), use_container_width=True)
                except Exception:
                    pass

        except Exception as exc:
            log_action(
                "bom_update_by_join",
                status="error",
                payload={"table": target_table, "error": str(exc)},
            )
            st.error(f"❌ Update gagal: {exc}")
