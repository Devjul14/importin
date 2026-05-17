from __future__ import annotations

import socket
import sys

if sys.platform == "win32" and not hasattr(socket, "AF_UNIX"):
    socket.AF_UNIX = None  # type: ignore[attr-defined]

import pandas as pd
import streamlit as st
from sqlalchemy import text

try:
    from app.db import get_engine, get_table_columns, list_tables, bulk_update_by_join, bom_replace_by_join
    from app.action_log import log_action
except ModuleNotFoundError:
    from db import get_engine, get_table_columns, list_tables, bulk_update_by_join, bom_replace_by_join  # type: ignore
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


def _get_col_types(eng, table: str) -> dict[str, str]:
    """Return {col_name: mysql_type_str} using INFORMATION_SCHEMA."""
    dialect = eng.dialect.name
    if dialect == "mysql":
        sql = text(
            "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS "
            "WHERE TABLE_NAME = :tbl AND TABLE_SCHEMA = DATABASE()"
        )
        with eng.connect() as conn:
            rows = conn.execute(sql, {"tbl": table}).fetchall()
        return {r[0]: r[1].lower() for r in rows}
    # SQLite fallback via PRAGMA
    with eng.connect() as conn:
        rows = conn.execute(text(f"PRAGMA table_info({table})")).fetchall()
    return {r[1]: r[2].lower() for r in rows}


_INTEGER_TYPES = {"int", "integer", "bigint", "smallint", "tinyint", "mediumint"}


def _series_is_non_numeric(series: pd.Series) -> bool:
    """Return True if the column contains at least one non-numeric string value."""
    sample = series.dropna().astype(str)
    if sample.empty:
        return False
    try:
        pd.to_numeric(sample, errors="raise")
        return False
    except (ValueError, TypeError):
        return True


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

# ── Schema type check ─────────────────────────────────────────
with st.expander("🔧 Cek & Perbaiki Tipe Kolom DB", expanded=False):
    st.caption(
        "Periksa apakah tipe kolom DB kompatibel dengan data Excel. "
        "Jika kolom DB bertipe INT namun data adalah string (misal `BB00012`), "
        "klik **ALTER** untuk mengubah tipe ke `VARCHAR(100)`."
    )
    try:
        col_types = _get_col_types(engine, target_table)
        type_rows = []
        for c in [join_col_db] + update_cols_db:
            db_type = col_types.get(c, "unknown")
            type_rows.append({"Kolom DB": c, "Tipe Saat Ini": db_type})
        st.dataframe(_safe(pd.DataFrame(type_rows)), use_container_width=True, hide_index=True)

        int_update_cols = [
            c for c in update_cols_db
            if any(t in col_types.get(c, "") for t in _INTEGER_TYPES)
        ]
        if int_update_cols:
            st.warning(
                f"⚠️ Kolom berikut bertipe INTEGER di DB: "
                + ", ".join(f"`{c}`" for c in int_update_cols)
                + ". Jika data Excel berupa teks (misal `BB00001`), UPDATE akan gagal."
            )
            alter_col = st.selectbox(
                "Pilih kolom yang ingin di-ALTER ke VARCHAR(100)",
                options=int_update_cols,
                key="alter_col_select",
            )
            alter_len = st.number_input("Panjang VARCHAR", min_value=10, max_value=500, value=100, step=10, key="alter_len")
            if st.button(f"⚡ ALTER `{alter_col}` → VARCHAR({int(alter_len)})", key="btn_alter"):
                try:
                    with engine.begin() as conn:
                        conn.execute(text(
                            f"ALTER TABLE {target_table} "
                            f"MODIFY COLUMN {alter_col} VARCHAR({int(alter_len)})"
                        ))
                    st.success(f"✅ Kolom `{alter_col}` berhasil diubah ke VARCHAR({int(alter_len)}). Refresh halaman untuk melihat tipe terbaru.")
                    st.cache_resource.clear()
                except Exception as exc:
                    st.error(f"ALTER gagal: {exc}")
        else:
            st.success("✅ Semua tipe kolom terlihat kompatibel.")
    except Exception as exc:
        st.warning(f"Tidak bisa cek tipe kolom: {exc}")

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

# ── Match check against DB ─────────────────────────────────────
if rows_to_update:
    excel_unique = list({r[join_col_db] for r in rows_to_update})
    try:
        from sqlalchemy import MetaData as _Meta, Table as _Tbl, select as _sel
        _meta = _Meta()
        _tbl = _Tbl(target_table, _meta, autoload_with=engine)
        with engine.connect() as _conn:
            db_vals = {
                str(r[0])
                for r in _conn.execute(_sel(_tbl.c[join_col_db])).fetchall()
            }
        matched = [v for v in excel_unique if v in db_vals]
        unmatched = [v for v in excel_unique if v not in db_vals]
        mc1, mc2 = st.columns(2)
        mc1.metric("✅ Produk cocok (akan diproses)", len(matched))
        mc2.metric("⚠️ Produk tidak ditemukan di DB", len(unmatched))
        if unmatched:
            with st.expander(f"🔎 {len(unmatched)} nama produk tidak ada di DB (akan dilewati)"):
                st.dataframe(pd.DataFrame({join_col_db: unmatched}), use_container_width=True, hide_index=True)
        if not matched:
            st.error(
                f"❌ Tidak ada satu pun `{join_col_db}` dari Excel yang cocok dengan data di DB. "
                "Pastikan kolom join sudah benar dan DB URL sudah terhubung ke database yang tepat."
            )
            st.stop()
    except Exception as _exc:
        st.warning(f"Tidak bisa cek kecocokan: {_exc}")

st.info(f"**{len(rows_to_update)}** baris dari Excel siap diproses ke `{target_table}` via `{join_col_db}`.")

# ── Strategy selector ─────────────────────────────────────────
strategy = st.radio(
    "Strategi import:",
    options=["replace", "update"],
    format_func=lambda x: {
        "replace": "♻️ Delete lama → Insert baru  (aman untuk multi-baris per produk & unique key)",
        "update":  "✏️ UPDATE per baris  (cocok jika 1 baris per produk, tanpa unique constraint)",
    }[x],
    index=0,
    key="upd_strategy",
)

carry_cols: list[str] = []
delete_col: str | None = None
if strategy == "replace":
    st.caption(
        "Mode **Delete → Insert**: baris lama untuk produk yang ada di Excel dihapus, "
        "lalu baris baru di-insert dengan `id_produk` diambil otomatis dari DB. "
        "Produk yang tidak ditemukan di DB akan **dilewati** (tidak dihapus, tidak diinsert)."
    )
    rc1, rc2 = st.columns(2)
    with rc1:
        carry_candidates = [c for c in table_cols if c not in mapped_update_cols and c != join_col_db]
        carry_cols = st.multiselect(
            "Kolom carry-over dari DB",
            options=carry_candidates,
            default=[c for c in ["id_produk", "created_by", "created_at"] if c in carry_candidates],
            help="Kolom yang diambil dari baris lama sebelum delete, lalu disertakan di INSERT baru.",
            key="upd_carry_cols",
        )
    with rc2:
        delete_col_options = ["(sama dengan join col)"] + [c for c in carry_cols if c in table_cols]
        delete_col_sel = st.selectbox(
            "DELETE by kolom",
            options=delete_col_options,
            index=1 if "id_produk" in carry_cols else 0,
            help="Gunakan id_produk untuk DELETE agar lebih aman dari konflik unique key.",
            key="upd_delete_col",
        )
        delete_col = None if delete_col_sel == "(sama dengan join col)" else delete_col_sel

# Preview table
preview_df = pd.DataFrame(rows_to_update)
with st.expander("🔍 Preview data yang akan diproses (maks 50 baris)", expanded=True):
    st.dataframe(_safe(preview_df.head(50)), use_container_width=True)

# Show sample SQL
if rows_to_update:
    sample = rows_to_update[0]
    if strategy == "replace":
        unique_names = list({r[join_col_db] for r in rows_to_update})[:3]
        del_col_display = delete_col if delete_col else join_col_db
        in_list = ", ".join(f"'{v}'" for v in unique_names)
        col_list = ", ".join([join_col_db] + mapped_update_cols + [c for c in carry_cols if c != join_col_db])
        val_ex = ", ".join([f"'{sample.get(c, '?')}'" for c in [join_col_db] + mapped_update_cols])
        sample_sql = (
            f"-- Step 1: Lookup id_produk by nama_produk (internal)\n\n"
            f"-- Step 2: Hapus baris lama by {del_col_display}\n"
            f"DELETE FROM {target_table} WHERE {del_col_display} IN (id1, id2, ...);\n\n"
            f"-- Step 3: Insert baris baru ({join_col_db} sebagai referensi + id_produk dari carry)\n"
            f"INSERT INTO {target_table} ({col_list}) VALUES ({val_ex}, ...carry...);\n"
            f"-- Total: {len(rows_to_update)} baris dari {len(set(r[join_col_db] for r in rows_to_update))} produk"
        )
    else:
        set_part = ", ".join(f"{c} = '{sample.get(c)}'" for c in mapped_update_cols)
        sample_sql = (
            f"UPDATE {target_table}\n"
            f"   SET {set_part}\n"
            f" WHERE {join_col_db} = '{sample[join_col_db]}';\n"
            f"-- ... ({len(rows_to_update)} total statements)"
        )
    with st.expander("📝 Contoh SQL"):
        st.code(sample_sql, language="sql")

# Execute
update_btn = st.button(
    f"🚀 Eksekusi {'Delete+Insert' if strategy == 'replace' else 'UPDATE'} "
    f"{len(rows_to_update)} baris ke `{target_table}`",
    type="primary",
    use_container_width=True,
    key="upd_execute",
)

if update_btn:
    if not rows_to_update:
        st.error("Tidak ada data untuk diproses.")
    elif strategy == "replace":
        try:
            inserted, skipped, not_found_vals = bom_replace_by_join(
                engine,
                table_name=target_table,
                rows=rows_to_update,
                join_col=join_col_db,
                carry_cols=carry_cols,
                delete_col=delete_col,
            )
            log_action(
                "bom_replace_by_join",
                status="success",
                payload={
                    "table": target_table,
                    "join_col": join_col_db,
                    "delete_col": delete_col,
                    "carry_cols": carry_cols,
                    "sent": len(rows_to_update),
                    "inserted": inserted,
                    "skipped": skipped,
                    "not_found_count": len(not_found_vals),
                },
            )
            st.success(f"✅ **{inserted} baris** berhasil di-insert ke `{target_table}`.")
            if skipped:
                st.warning(
                    f"⚠️ **{skipped} baris** dilewati karena `{join_col_db}` tidak ditemukan di DB:"
                )
                with st.expander(f"🔎 {len(not_found_vals)} nama produk tidak ditemukan di DB"):
                    st.dataframe(
                        pd.DataFrame({join_col_db: not_found_vals}),
                        use_container_width=True, hide_index=True,
                    )
            st.balloons()
        except Exception as exc:
            log_action(
                "bom_replace_by_join",
                status="error",
                payload={"table": target_table, "error": str(exc)},
            )
            st.error(f"❌ Gagal: {exc}")
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
                st.warning(f"⚠️ **{not_found} baris** tidak ditemukan di DB.")
            st.balloons()
        except Exception as exc:
            log_action(
                "bom_update_by_join",
                status="error",
                payload={"table": target_table, "error": str(exc)},
            )
            st.error(f"❌ Update gagal: {exc}")
