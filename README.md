# update_precious_metals

Small utility to fetch current gold and silver spot prices and update the
`Current Prices` sheet in the Excel workbook `Precious Metal Investments.xlsx`.

What it does
- Fetches Gold (`gc.f`) and Silver (`si.f`) from Stooq (with a fallback CSV URL).
- Updates `Current Prices` sheet cells:
  - `B2` ← Gold (USD/oz)
  - `B3` ← Silver (USD/oz)
  - `C2`/`C3` ← today's date (formatted `m/d/yy`)
- Creates a timestamped backup of the workbook before writing.

Files
- `update_precious_metals.py` — main script
- `requirements.txt` — runtime dependencies

Install

1. (Optional) create and activate a virtualenv:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

2. Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Usage

- Dry-run (fetch prices but don't write):

```bash
python3 update_precious_metals.py --dry-run
```

- Write to the default workbook (creates a backup first):

```bash
python3 update_precious_metals.py
```

- Use a different workbook path:

```bash
python3 update_precious_metals.py --workbook "/path/to/Precious Metal Investments.xlsx"
```

Notes
- Default workbook path is `~/Dropbox/Excel Shared/Precious Metal Investments.xlsx`.
- Backups are created in the same folder with the pattern
  `Precious Metal Investments.backup.YYYYMMDDHHMMSS.xlsx`.
- The script formats prices as currency (`$#,##0.00`) and dates as `m/d/yy`.

Scheduling
- macOS (launchd) or cron can call the script periodically. Example cron (daily at 17:00):

```cron
0 17 * * * /usr/bin/python3 /path/to/update_precious_metals.py
```

Troubleshooting
- If `requests` warns about character detection, install `charset_normalizer` (included in `requirements.txt`).
- If the sheet name or workbook path differs, pass `--workbook` or edit `WORKBOOK_PATH` in the script.
