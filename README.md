# 📊 SQL Query Exporter — Automated SQL-to-Excel/CSV Pipeline

An automated export pipeline that reads `.sql` files from a defined input folder, executes the queries against a Microsoft SQL Server, and exports the results as formatted Excel or CSV files — locally runnable, modular, and production-ready.

---

## Architecture Overview

```
Input Folder (*.sql Files)
        │
        ▼
┌─────────────────────────┐
│   SQL Query Exporter    │  ← Local execution (manual or scheduled)
│   (Python 3.x)          │
│                         │
│  1. Load .sql files      │
│  2. Connect to SQL DB    │
│  3. Execute queries      │
│  4. Export to Excel/CSV  │
└─────────────────────────┘
        │
        ▼
┌─────────────────────────┐
│   Output Folder         │  ← .xlsx / .csv per query file
│   (Excel / CSV)         │
└─────────────────────────┘
        │
        ▼
┌─────────────────────────┐
│   Logging               │  ← Daily rotating .log file
└─────────────────────────┘
```

---

## Tech Stack

| Layer | Technology |
| --- | --- |
| Runtime | Python 3.x |
| Data Processing | Pandas |
| Database Driver | SQLAlchemy + pyodbc (ODBC Driver 17) |
| Excel Export | openpyxl |
| Configuration | python-dotenv (.env file) |
| Logging | Python logging module |

---

## Features

- **Multi-file processing** — All `.sql` files in the input folder are picked up and processed automatically
- **Flexible authentication** — Supports both Windows Authentication (Trusted Connection) and SQL login via username & password
- **Formatted Excel output** — Auto-filter, frozen header row, and auto-fitted column widths out of the box
- **CSV export option** — Alternative export as semicolon-separated CSV with UTF-8 BOM for Excel compatibility
- **Export timestamp** — Every output file contains a dedicated `Export_Zeitstempel` column
- **Structured logging** — Daily log files written to a configurable path, mirrored to the console

---

## Pipeline Flow

### 1. `verbinde_mit_sql_datenbank()` — DB Connection

Reads connection parameters from the `.env` file and creates a SQLAlchemy engine. Supports two authentication modes:
- **Trusted Connection** — Windows Authentication (no credentials required)
- **SQL Login** — Username & password passed via ODBC connection string

### 2. `lade_sql_dateien()` — Load SQL Files

Scans the input directory for all `*.sql` files and reads their content into a dictionary (`filename → query string`). Raises a `FileNotFoundError` if no `.sql` files are found.

### 3. `sql_dataframe_erstellen()` — Execute Queries

Iterates over all loaded queries and executes them against the database via `pd.read_sql()`. Returns a dictionary of (`filename → DataFrame`).

### 4. `export_to_excel()` / `export_to_csv()` — Export

- **Excel**: Writes each DataFrame to a separate `.xlsx` file with formatted worksheet (freeze panes, auto-filter, adjusted column widths)
- **CSV**: Writes each DataFrame to a separate `.csv` file using `;` as separator and `utf-8-sig` encoding for Excel compatibility

---

## Configuration

All settings are managed via a `.env` file (never hardcoded). Copy `.env.example` and fill in your values:

| Variable | Description |
| --- | --- |
| `INPUT_PATH` | Path to the folder containing `.sql` files |
| `OUTPUT_PATH` | Path to the folder where exports will be saved |
| `LOGGER_PATH` | Path to the folder where log files will be written |
| `DB_SERVER` | SQL Server hostname or IP address |
| `DATABASE` | Target database name |
| `TRUSTED_CONNECTION` | `yes` for Windows Auth, `no` for SQL login |
| `DB_USER` | SQL login username (only if `TRUSTED_CONNECTION=no`) |
| `DB_PASSWORD` | SQL login password (only if `TRUSTED_CONNECTION=no`) |
| `ENCODING` | File encoding for reading `.sql` files (e.g. `utf-8`) |

---

## Local Setup

```bash
# 1. Clone the repository
git clone https://github.com/your-username/sql-query-exporter.git
cd sql-query-exporter

# 2. Install dependencies
pip install -r requirements.txt

# 3. Configure environment
cp .env.example .env
# → Open .env and fill in your values

# 4. Add your SQL files to the input folder
# → e.g. input/my_query.sql

# 5. Run the exporter
python sql_query_exporter.py
```

---

## Project Structure

```
├── sql_query_exporter.py   # Main ETL logic
├── .env.example            # Environment variable template
├── .gitignore              # Excludes .env, logs, __pycache__
├── requirements.txt        # Python dependencies
└── README.md               # Project documentation
```

---

## Key Design Decisions

**Dictionary-based processing** — SQL files and DataFrames are managed as `filename → content` dictionaries throughout the pipeline, making it easy to trace which query produced which output.

**Engine vs. Connection separation** — SQLAlchemy engine creation and connection handling are kept separate, allowing clean resource management and explicit connection closing.

**Configurable export format** — Excel is the default export format. The CSV export function is available as a drop-in alternative with a single line change, without touching any other logic.

**No hardcoded paths or credentials** — All environment-specific values live in the `.env` file, making the script portable across machines and environments without code changes.
