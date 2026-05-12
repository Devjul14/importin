# 📥 Excel to Database Import Tool

Tool berbasis web untuk mengimpor data dari file Excel ke database secara visual — tanpa menulis kode.

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-1.44%2B-red)
![SQLAlchemy](https://img.shields.io/badge/SQLAlchemy-2.0%2B-green)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

---

## ✨ Fitur

| Fitur | Keterangan |
|---|---|
| **SQL INSERT sebagai schema** | Paste SQL INSERT untuk definisikan nama tabel dan kolom target |
| **Upload Excel** | Support `.xlsx` dan `.xls`, multi-sheet |
| **Visual column mapping** | Dropdown per kolom DB — auto-match by nama |
| **Auto-generate nilai** | Generate nilai per baris: timestamp, UUID, nomor urut, prefix/suffix |
| **Foreign key lookup** | Lookup ID dari tabel master by nilai kolom Excel |
| **Unique / duplicate check** | Deteksi baris duplikat sebelum import berdasarkan kolom yang dipilih |
| **Conflict strategy** | Skip / Replace / Error saat data sudah ada di DB |
| **Import gambar embedded** | Ekstrak gambar dari sel Excel, simpan sebagai base64 / file / binary |
| **Penamaan gambar dinamis** | Nama file gambar dari nilai kolom Excel + prefix/suffix |
| **Multi-database** | SQLite (default), MySQL, PostgreSQL via SQLAlchemy URL |

---

## 🚀 Quick Start

### 1. Clone & Setup

```bash
git clone https://github.com/<username>/importin.git
cd importin

python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Jalankan

```bash
streamlit run app/main.py
```

Buka `http://localhost:8501` di browser.

---

## 🗂️ Struktur Project

```
importin/
├── app/
│   ├── main.py             # UI utama Streamlit
│   ├── db.py               # SQLAlchemy engine, insert, lookup
│   ├── sql_parser.py       # Parser SQL INSERT → schema
│   ├── image_extractor.py  # Ekstrak gambar embedded dari Excel
│   ├── validators.py       # Validasi JSON input
│   ├── importer.py         # Build rows dari DataFrame + mapping
│   └── sample_data.py      # Seed dummy data ke DB
├── requirements.txt
└── README.md
```

---

## 🔄 Alur Import (4 Step)

### ① SQL INSERT — Struktur & Data Referensi

Paste SQL INSERT untuk mendefinisikan **nama tabel** dan **urutan kolom** target.

```sql
INSERT INTO products (id, kode, nama, harga, foto)
VALUES (1, 'P001', 'Produk A', 50000, NULL);
```

Baris `VALUES` opsional — bisa langsung di-seed ke DB sebagai data referensi.

---

### ② Upload File Excel

- Upload file `.xlsx` / `.xls`
- Pilih sheet yang diinginkan
- Gambar embedded di sel akan otomatis terdeteksi

---

### ③ Mapping Kolom

**Column Mapping** — dropdown per kolom DB:
- Auto-match jika nama kolom sama (case-insensitive)
- Pilih `— skip —` untuk melewati kolom

**Unique Check** — pilih kolom yang nilainya tidak boleh duplikat dalam file Excel

**Auto-generate** — isi pola nilai otomatis per baris:

| Token | Output |
|---|---|
| `{YmdHis}` | `20260423143005` |
| `{Y}` `{m}` `{d}` | `2026` `04` `23` |
| `{H}` `{i}` `{s}` | `14` `30` `05` |
| `{n}` | `1`, `2`, `3` (nomor baris) |
| `{n:03}` | `001`, `002`, `003` |
| `{uuid}` | UUID4 hex unik per baris |

Contoh: `SMP-{YmdHis}-{n:03}` → `SMP-20260423143005-001`

**Foreign Key Lookup** — per kolom, konfigurasi:
- Table master yang jadi referensi
- Kolom pencarian (nilai Excel akan di-match ke kolom ini)
- Kolom return (nilai yang diambil, biasanya `id`)

**Import Gambar** — untuk kolom gambar embedded di Excel:
- Pilih kolom DB tujuan
- Format: `base64` / `path` / `binary`
- **Nama file dari kolom**: pilih kolom Excel untuk penamaan file (misal `kode_barang` → `B001.png`)
- Prefix / suffix nama file opsional

---

### ④ Import & Hasil

- Preview INSERT query (maks 50 baris)
- Preview data sebelum dikirim
- Pilih conflict strategy: **Skip** / **Replace** / **Error**
- Klik **🚀 Import**
- Hasil: tab data terimport + full INSERT query

---

## 🔌 Koneksi Database

### SQLite (default)
```
sqlite:///import_tool.db
```

### MySQL lokal
```
mysql+pymysql://root@localhost/nama_db
mysql+pymysql://root:password@localhost/nama_db
mysql+pymysql://root@localhost/nama_db?unix_socket=/tmp/mysql.sock
```

> ⚠️ Gunakan `localhost` bukan `127.0.0.1` agar koneksi pakai Unix socket.

Install driver MySQL:
```bash
pip install pymysql
```

### PostgreSQL
```
postgresql+psycopg2://user:pass@localhost:5432/nama_db
```

Install driver:
```bash
pip install psycopg2-binary
```

---

## 🖼️ Import Gambar dari Excel

Gambar embedded di sel Excel (bukan hyperlink) otomatis diekstrak.

**Penamaan file:**
- Pilih kolom Excel (misal `kode_barang`) → file saved sebagai `B001.png`
- Tambah prefix: `product_` → `product_B001.png`
- Tambah suffix: `_foto` → `B001_foto.png`
- Tanpa kolom → nama default nomor baris: `1.png`, `2.png`

**Format penyimpanan:**
| Format | Cocok untuk | Tipe kolom DB |
|---|---|---|
| `base64` | Simpan langsung di DB | `TEXT` / `LONGTEXT` |
| `path` | Simpan ke disk, path di DB | `VARCHAR` |
| `binary` | BLOB di DB | `BLOB` / `LONGBLOB` |

---

## 🧾 Log Action

Semua aksi penting (upload Excel, seed data, import, clear cache) dicatat ke file log JSONL.

### Lihat dari UI
- Jalankan aplikasi: `streamlit run app/main.py`
- Buka `http://localhost:8501`
- Scroll ke section **🧾 Log Action** untuk melihat 50 log terbaru

### Lihat dari terminal
```bash
# 50 baris terakhir
tail -n 50 uploads/action_log.jsonl

# pantau realtime
tail -f uploads/action_log.jsonl
```

Lokasi file log: `uploads/action_log.jsonl`

---

## 📦 Dependencies

```
streamlit>=1.44.0
pandas>=2.2.0
openpyxl>=3.1.0
SQLAlchemy>=2.0.0
pymysql>=1.1.0
pillow>=10.0.0
```

---

## 📄 License

MIT
