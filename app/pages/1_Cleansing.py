"""
Cleansing Page — interactive data cleansing before import.
"""

from __future__ import annotations

import sys
from pathlib import Path

# Ensure app/ is importable when run as a Streamlit page
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
import streamlit as st

try:
    from app.cleansing import (
        coerce_datetime, coerce_numeric, drop_duplicates,
        drop_empty_columns, drop_empty_rows, fill_missing, find_replace,
        load_excel_unmerged, normalize_whitespace, null_summary,
        rename_columns, select_columns, split_column, strip_unit, to_csv_bytes,
        to_excel_bytes, to_lowercase, to_titlecase, to_uppercase, trim_whitespace,
        unpivot,
    )
except ModuleNotFoundError:
    from cleansing import (  # type: ignore
        coerce_datetime, coerce_numeric, drop_duplicates,
        drop_empty_columns, drop_empty_rows, fill_missing, find_replace,
        load_excel_unmerged, normalize_whitespace, null_summary,
        rename_columns, select_columns, split_column, strip_unit, to_csv_bytes,
        to_excel_bytes, to_lowercase, to_titlecase, to_uppercase, trim_whitespace,
        unpivot,
    )

st.set_page_config(page_title="Cleansing Data", page_icon="🧹", layout="wide")
st.title("🧹 Cleansing Data")
st.caption("Bersihkan file Excel sebelum digunakan di modul import.")

# ─── Session state ───────────────────────────────────────────────────────────
if "clean_df" not in st.session_state:
    st.session_state.clean_df = None
if "clean_history" not in st.session_state:
    st.session_state.clean_history: list[tuple[str, pd.DataFrame]] = []


def _push(label: str, df: pd.DataFrame) -> None:
    """Save snapshot to history and update working df."""
    st.session_state.clean_history.append((label, st.session_state.clean_df.copy()))
    st.session_state.clean_df = df.copy()


def _safe(df: pd.DataFrame) -> pd.DataFrame:
    # Rename duplicate column names before converting (PyArrow requires unique cols)
    cols = pd.Series(df.columns.astype(str))
    for dup in cols[cols.duplicated()].unique():
        mask = cols == dup
        cols[mask] = [dup if i == 0 else f"{dup}.{i}" for i, _ in enumerate(mask[mask].index)]
    df = df.copy()
    df.columns = cols.tolist()
    return df.astype(str).replace("nan", "").replace("<NA>", "")


# ─── Upload ───────────────────────────────────────────────────────────────────
st.subheader("① Upload File Excel")

uploaded = st.file_uploader("Pilih file Excel (.xlsx / .xls)", type=["xlsx", "xls"], key="clean_upload")

if uploaded:
    xlsx = pd.ExcelFile(uploaded)
    col_sheet, col_hdr = st.columns(2)
    with col_sheet:
        sheet = st.selectbox("Sheet", xlsx.sheet_names, key="clean_sheet")
    with col_hdr:
        header_row = st.number_input(
            "Baris header (1-based)", min_value=1, max_value=20, value=1, key="clean_hdr",
            help="Nomor baris yang dijadikan nama kolom. Baris di atasnya diabaikan."
        )

    unmerge = st.checkbox(
        "🔓 Un-merge cells otomatis",
        value=True,
        key="clean_unmerge",
        help="Isi sel kosong hasil merge dengan nilai sel pertama di range merge.",
    )

    if st.button("📂 Muat File", type="primary"):
        with st.spinner("Memuat..."):
            try:
                uploaded.seek(0)
                if unmerge:
                    df = load_excel_unmerged(uploaded, sheet_name=sheet, header_row=int(header_row) - 1)
                else:
                    uploaded.seek(0)
                    df = pd.read_excel(uploaded, sheet_name=sheet, header=int(header_row) - 1)

                st.session_state.clean_df = df.copy()
                st.session_state.clean_history = []
                st.success(f"Dimuat: {len(df)} baris × {len(df.columns)} kolom")
            except Exception as exc:
                st.error(f"Gagal memuat: {exc}")

# ─── Working area ─────────────────────────────────────────────────────────────
df: pd.DataFrame | None = st.session_state.clean_df

if df is None:
    st.info("⬆ Upload dan muat file Excel terlebih dahulu.")
    st.stop()

st.divider()

# ─── Stats bar ────────────────────────────────────────────────────────────────
m1, m2, m3, m4 = st.columns(4)
m1.metric("Baris", len(df))
m2.metric("Kolom", len(df.columns))
m3.metric("Sel kosong", int(df.isna().sum().sum()))
m4.metric("Duplikat", int(df.duplicated().sum()))

with st.expander("📊 Preview data sekarang", expanded=True):
    st.dataframe(_safe(df.head(20)), use_container_width=True)

with st.expander("🔍 Null summary per kolom"):
    st.dataframe(null_summary(df), use_container_width=True)

st.divider()

# ─── Toolbar operations ───────────────────────────────────────────────────────
st.subheader("② Operasi Cleansing")

tabs = st.tabs([
    "📌 Pilih Kolom",
    "🗑️ Hapus Kosong",
    "✂️ Trim & Case",
    "🔁 Rename",
    "🚫 Deduplikasi",
    "🔢 Tipe Data",
    "🔍 Find & Replace",
    "❓ Isi Missing",
    "✂️ Split Kolom",
    "🔄 Unpivot",
])

# ────── TAB 1: Select columns ──────────────────────────────────────────────────
with tabs[0]:
    st.markdown("**Pilih kolom yang dibutuhkan** (kolom lain akan dihapus).")
    keep = st.multiselect(
        "Kolom yang disimpan", df.columns.tolist(),
        default=df.columns.tolist(), key="t_keep_cols"
    )
    reorder = st.checkbox("Urutkan sesuai pilihan di atas", value=True, key="t_reorder")
    if st.button("✅ Terapkan pilihan kolom", key="btn_select"):
        new_df = select_columns(df, keep)
        _push(f"Pilih {len(keep)} kolom", new_df)
        st.success(f"{len(new_df.columns)} kolom tersimpan.")
        st.rerun()

# ────── TAB 2: Drop empty ──────────────────────────────────────────────────────
with tabs[1]:
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("**Hapus baris kosong**")
        row_how = st.radio("Hapus baris jika", ["Semua kolom kosong", "Salah satu kolom kosong"], key="t_row_how")
        row_subset = st.multiselect("Cek kolom (kosong = semua)", df.columns.tolist(), key="t_row_subset")
        if st.button("🗑️ Hapus baris kosong", key="btn_drop_rows"):
            how = "all" if "Semua" in row_how else "any"
            new_df = drop_empty_rows(df, how=how, subset=row_subset or None)
            removed = len(df) - len(new_df)
            _push(f"Hapus {removed} baris kosong", new_df)
            st.success(f"{removed} baris dihapus.")
            st.rerun()

    with c2:
        st.markdown("**Hapus kolom kosong**")
        threshold = st.slider(
            "Hapus kolom jika % kosong ≥", 0, 100, 100, format="%d%%", key="t_col_thresh"
        )
        if st.button("🗑️ Hapus kolom kosong", key="btn_drop_cols"):
            new_df = drop_empty_columns(df, threshold=threshold / 100)
            removed = len(df.columns) - len(new_df.columns)
            _push(f"Hapus {removed} kolom kosong", new_df)
            st.success(f"{removed} kolom dihapus.")
            st.rerun()

# ────── TAB 3: Trim & Case ─────────────────────────────────────────────────────
with tabs[2]:
    trim_cols = st.multiselect("Kolom (kosong = semua kolom teks)", df.columns.tolist(), key="t_trim_cols")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✂️ Trim whitespace", key="btn_trim"):
            new_df = trim_whitespace(df, trim_cols or None)
            _push("Trim whitespace", new_df)
            st.success("Whitespace di-trim.")
            st.rerun()
        if st.button("↔️ Normalize spasi berganda", key="btn_norm"):
            new_df = normalize_whitespace(df, trim_cols or None)
            _push("Normalize whitespace", new_df)
            st.success("Spasi berganda dinormalisasi.")
            st.rerun()
    with c2:
        case_cols = st.multiselect("Kolom untuk ubah case", df.columns.tolist(), key="t_case_cols")
        case_type = st.radio("Ubah ke", ["UPPERCASE", "lowercase", "Title Case"], key="t_case_type")
        if st.button("🔡 Terapkan case", key="btn_case"):
            if not case_cols:
                st.warning("Pilih minimal satu kolom.")
            else:
                if case_type == "UPPERCASE":
                    new_df = to_uppercase(df, case_cols)
                elif case_type == "lowercase":
                    new_df = to_lowercase(df, case_cols)
                else:
                    new_df = to_titlecase(df, case_cols)
                _push(f"Case → {case_type}", new_df)
                st.success("Case diubah.")
                st.rerun()

# ────── TAB 4: Rename ──────────────────────────────────────────────────────────
with tabs[3]:
    st.markdown("Edit nama kolom di bawah ini:")
    rename_data = pd.DataFrame({"Nama lama": df.columns.tolist(), "Nama baru": df.columns.tolist()})
    edited = st.data_editor(rename_data, use_container_width=True, hide_index=True, key="t_rename_editor")
    if st.button("✅ Terapkan rename", key="btn_rename"):
        rmap = {row["Nama lama"]: row["Nama baru"] for _, row in edited.iterrows() if row["Nama lama"] != row["Nama baru"]}
        if rmap:
            new_df = rename_columns(df, rmap)
            _push(f"Rename {len(rmap)} kolom", new_df)
            st.success(f"{len(rmap)} kolom diubah namanya.")
            st.rerun()
        else:
            st.info("Tidak ada perubahan nama.")

# ────── TAB 5: Deduplicate ─────────────────────────────────────────────────────
with tabs[4]:
    dup_subset = st.multiselect("Cek duplikat berdasarkan kolom (kosong = semua)", df.columns.tolist(), key="t_dup_subset")
    dup_keep = st.radio("Simpan", ["Pertama (first)", "Terakhir (last)"], key="t_dup_keep")
    dups_preview = df[df.duplicated(subset=dup_subset or None, keep=False)]
    st.caption(f"{len(dups_preview)} baris terduplikasi terdeteksi.")
    if not dups_preview.empty:
        with st.expander("Lihat baris duplikat"):
            st.dataframe(_safe(dups_preview), use_container_width=True)
    if st.button("🚫 Hapus duplikat", key="btn_dedup"):
        keep_val = "first" if "first" in dup_keep else "last"
        new_df, removed_df = drop_duplicates(df, subset=dup_subset or None, keep=keep_val)
        _push(f"Dedup {len(removed_df)} baris", new_df)
        st.success(f"{len(removed_df)} duplikat dihapus.")
        st.rerun()

# ────── TAB 6: Type coercion ───────────────────────────────────────────────────
with tabs[5]:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**🔢 Konversi ke Angka**")
        num_cols = st.multiselect("Kolom angka", df.columns.tolist(), key="t_num_cols")
        if st.button("🔢 Konversi numeric", key="btn_numeric"):
            if not num_cols:
                st.warning("Pilih kolom.")
            else:
                new_df, errs = coerce_numeric(df, num_cols)
                _push(f"Numeric: {', '.join(num_cols)}", new_df)
                for col, cnt in errs.items():
                    if cnt:
                        st.warning(f"`{col}`: {cnt} nilai tidak bisa dikonversi → NaN")
                st.success("Konversi selesai.")
                st.rerun()

    with c2:
        st.markdown("**📅 Konversi ke Tanggal**")
        date_cols = st.multiselect("Kolom tanggal", df.columns.tolist(), key="t_date_cols")
        date_fmt_in = st.text_input("Format input (opsional)", placeholder="%d/%m/%Y", key="t_date_fmt_in")
        date_fmt_out = st.text_input("Format output (opsional)", placeholder="%Y-%m-%d", key="t_date_fmt_out")
        if st.button("📅 Konversi datetime", key="btn_datetime"):
            if not date_cols:
                st.warning("Pilih kolom.")
            else:
                new_df, errs = coerce_datetime(
                    df, date_cols,
                    fmt=date_fmt_in or None,
                    output_fmt=date_fmt_out or None,
                )
                _push(f"Datetime: {', '.join(date_cols)}", new_df)
                for col, cnt in errs.items():
                    if cnt:
                        st.warning(f"`{col}`: {cnt} nilai tidak bisa diparse → NaT")
                st.success("Konversi selesai.")
                st.rerun()

# ────── TAB 7: Find & Replace ─────────────────────────────────────────────────
with tabs[6]:
    fr_col = st.selectbox("Kolom", df.columns.tolist(), key="t_fr_col")
    c1, c2 = st.columns(2)
    with c1:
        fr_find = st.text_input("Cari", key="t_fr_find")
    with c2:
        fr_replace = st.text_input("Ganti dengan", key="t_fr_replace")
    c3, c4 = st.columns(2)
    with c3:
        fr_regex = st.checkbox("Regex", key="t_fr_regex")
    with c4:
        fr_case = st.checkbox("Case-sensitive", value=True, key="t_fr_case")
    if st.button("🔍 Terapkan Find & Replace", key="btn_fr"):
        if not fr_find:
            st.warning("Isi teks yang ingin dicari.")
        else:
            new_df, count = find_replace(df, fr_col, fr_find, fr_replace, is_regex=fr_regex, case_sensitive=fr_case)
            _push(f"Replace `{fr_find}` di `{fr_col}`", new_df)
            st.success(f"{count} nilai diganti.")
            st.rerun()

# ────── TAB 8: Fill missing ────────────────────────────────────────────────────
with tabs[7]:
    fm_col = st.selectbox("Kolom", df.columns.tolist(), key="t_fm_col")
    fm_method = st.selectbox(
        "Metode isi",
        ["value", "ffill", "bfill", "mean", "median", "mode"],
        format_func=lambda x: {
            "value": "Nilai tetap",
            "ffill": "Forward fill (isi dari baris atas)",
            "bfill": "Backward fill (isi dari baris bawah)",
            "mean": "Rata-rata",
            "median": "Median",
            "mode": "Modus (nilai terbanyak)",
        }[x],
        key="t_fm_method",
    )
    fm_value = ""
    if fm_method == "value":
        fm_value = st.text_input("Nilai pengisi", key="t_fm_value")

    na_count = int(df[fm_col].isna().sum()) if fm_col in df.columns else 0
    st.caption(f"{na_count} nilai kosong di kolom `{fm_col}`.")

    if st.button("❓ Isi missing values", key="btn_fill"):
        new_df = fill_missing(df, fm_col, method=fm_method, value=fm_value or None)
        _push(f"Fill `{fm_col}` [{fm_method}]", new_df)
        st.success("Missing values diisi.")
        st.rerun()

# ────── TAB 9: Split column ────────────────────────────────────────────────────
with tabs[8]:
    sp_col = st.selectbox("Kolom sumber", df.columns.tolist(), key="t_sp_col")
    sp_delim = st.text_input("Delimiter", value=",", key="t_sp_delim")
    sp_names = st.text_input(
        "Nama kolom baru (pisahkan koma)",
        placeholder="kota, provinsi",
        key="t_sp_names",
    )
    if sp_col in df.columns:
        sample_vals = df[sp_col].dropna().head(3).tolist()
        st.caption(f"Contoh nilai: {sample_vals}")
    if st.button("✂️ Split kolom", key="btn_split"):
        names = [n.strip() for n in sp_names.split(",") if n.strip()]
        if not names:
            st.warning("Isi nama kolom baru.")
        else:
            new_df = split_column(df, sp_col, sp_delim, names)
            _push(f"Split `{sp_col}` → {names}", new_df)
            st.success(f"Kolom `{sp_col}` dipecah menjadi {len(names)} kolom.")
            st.rerun()

# ────── TAB 10: Unpivot ───────────────────────────────────────────────────────
with tabs[9]:
    st.markdown(
        "**Ubah format wide → long** (unpivot / melt). "
        "Cocok untuk tabel BOM di mana setiap kolom adalah kode bahan baku."
    )
    st.caption(
        "Contoh: kolom `Nama Barang | BB00022 | BB00003 | ...` "
        "→ baris `nama_barang | kode_bahan_baku | qty`"
    )

    uv_id_cols = st.multiselect(
        "📌 Kolom ID (tetap — tidak di-unpivot)",
        df.columns.tolist(),
        default=[df.columns[0]] if len(df.columns) > 0 else [],
        key="t_uv_id",
        help="Pilih kolom yang menjadi identitas baris, misal: Nama Barang, SKU, dll.",
    )

    remaining_cols = [c for c in df.columns if c not in uv_id_cols]
    uv_val_cols = st.multiselect(
        "📊 Kolom yang di-unpivot (nilai)",
        remaining_cols,
        default=remaining_cols,
        key="t_uv_vals",
        help="Pilih kolom bahan baku / material. Header kolom akan jadi value di kolom baru.",
    )

    c1, c2 = st.columns(2)
    with c1:
        uv_var_name = st.text_input(
            "Nama kolom baru — header asli",
            value="kode_bahan_baku",
            key="t_uv_varname",
        )
    with c2:
        uv_val_name = st.text_input(
            "Nama kolom baru — nilai",
            value="qty",
            key="t_uv_valname",
        )

    c3, c4, c5, c6 = st.columns(4)
    with c3:
        uv_drop_null = st.checkbox("Hapus baris dengan nilai NULL", value=True, key="t_uv_null")
    with c4:
        uv_drop_dash = st.checkbox("Hapus baris dengan nilai ' -' / '-'", value=True, key="t_uv_dash")
    with c5:
        uv_strip_unit = st.checkbox(
            "Hapus satuan (cm / Yrd / Pcs dll)",
            value=True,
            key="t_uv_strip",
            help="Contoh: '150 cm' → 150, '2743 Yrd' → 2743, '30 Pcs' → 30",
        )
    with c6:
        uv_sort = st.checkbox(
            "Urutkan per produk",
            value=True,
            key="t_uv_sort",
            help="Kelompokkan baris per Nama Barang sehingga semua bahan baku tiap produk berurutan.",
        )

    if uv_id_cols and uv_val_cols:
        preview_melt = unpivot(
            df,
            id_cols=uv_id_cols,
            value_cols=uv_val_cols,
            var_name=uv_var_name or "kode_bahan_baku",
            value_name=uv_val_name or "qty",
            drop_null_values=uv_drop_null,
            drop_dash_values=uv_drop_dash,
            strip_units=uv_strip_unit,
            sort_by_id=uv_sort,
        )
        st.caption(
            f"Preview hasil unpivot: **{len(preview_melt)} baris** × **{len(preview_melt.columns)} kolom** "
            f"(dari {len(df)} baris × {len(df.columns)} kolom)"
        )
        st.dataframe(_safe(preview_melt.head(20)), use_container_width=True)

        if st.button("🔄 Terapkan Unpivot", key="btn_unpivot", type="primary"):
            _push(
                f"Unpivot {len(uv_val_cols)} kolom → [{uv_var_name}, {uv_val_name}]",
                preview_melt,
            )
            st.success(
                f"Unpivot selesai: {len(preview_melt)} baris · "
                f"kolom: {', '.join(preview_melt.columns.tolist())}"
            )
            st.rerun()
    else:
        st.info("Pilih minimal 1 kolom ID dan 1 kolom nilai untuk melanjutkan.")

# ─── History & Undo ───────────────────────────────────────────────────────────
st.divider()
st.subheader("③ History & Undo")

if st.session_state.clean_history:
    hist_labels = [f"[{i+1}] {label}" for i, (label, _) in enumerate(st.session_state.clean_history)]
    hist_labels.insert(0, "— State sekarang —")
    selected_hist = st.selectbox("Restore ke state", hist_labels, index=0, key="t_hist")

    c_undo, c_restore = st.columns(2)
    with c_undo:
        if st.button("↩️ Undo (1 langkah)", key="btn_undo"):
            label, prev_df = st.session_state.clean_history.pop()
            st.session_state.clean_df = prev_df
            st.success(f"Undo: `{label}`")
            st.rerun()
    with c_restore:
        if st.button("⏮️ Restore ke state terpilih", key="btn_restore"):
            idx = hist_labels.index(selected_hist) - 1  # offset because of "— sekarang —"
            if idx >= 0:
                restore_label, restore_df = st.session_state.clean_history[idx]
                st.session_state.clean_df = restore_df.copy()
                st.session_state.clean_history = st.session_state.clean_history[:idx]
                st.success(f"Restored ke: `{restore_label}`")
                st.rerun()
else:
    st.caption("Belum ada operasi yang dilakukan.")

# ─── Export ───────────────────────────────────────────────────────────────────
st.divider()
st.subheader("④ Download Hasil")

out_df = st.session_state.clean_df
if out_df is not None:
    st.dataframe(_safe(out_df.head(10)), use_container_width=True)
    st.caption(f"Total: {len(out_df)} baris × {len(out_df.columns)} kolom")

    c1, c2 = st.columns(2)
    with c1:
        xlsx_bytes = to_excel_bytes(out_df)
        st.download_button(
            "⬇️ Download Excel (.xlsx)",
            data=xlsx_bytes,
            file_name="cleaned_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c2:
        csv_bytes = to_csv_bytes(out_df)
        st.download_button(
            "⬇️ Download CSV",
            data=csv_bytes,
            file_name="cleaned_data.csv",
            mime="text/csv",
            use_container_width=True,
        )
