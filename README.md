# Excel to Database Import Tool

Tool minimal untuk import data dari Excel ke database dengan alur terstruktur:

1. Schema setup (user define struktur tabel)
2. Dummy data reference (user input data referensi)
3. Mapping JSON (mapping kolom spreadsheet ke kolom database)
4. Import execution

## Stack

- Python
- Streamlit (UI)
- SQLAlchemy (DB)
- Pandas + OpenPyXL (Excel reader)

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run app/main.py
```

## Input JSON Examples

### Schema JSON

```json
[
  {"name": "id", "type": "integer", "primary_key": true, "nullable": false},
  {"name": "name", "type": "string", "length": 255},
  {"name": "email", "type": "string", "unique": true},
  {"name": "age", "type": "integer"}
]
```

### Dummy data JSON

```json
[
  {"name": "Demo User 1", "email": "demo1@local", "age": 28},
  {"name": "Demo User 2", "email": "demo2@local", "age": 34}
]
```

### Mapping JSON

```json
{
  "name": "Nama",
  "email": "Email",
  "age": {"source": "Umur"},
  "status": {"value": "active"}
}
```

## Notes

- Default DB URL: `sqlite:///import_tool.db`
- Bisa ganti ke DB lain via SQLAlchemy URL di sidebar.
