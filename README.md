# data-backup-reports

## weekly-csv-backup-to-dropbox
Back up recent CSVs to Dropbox as a dated zip, with optional retention.

### Setup
1. Install Python 3.9+.
2. `python -m venv .venv && source .venv/bin/activate` (Windows: `. .venv/Scripts/activate`)
3. `pip install -r requirements.txt`
4. Copy `.env.example` → `.env`, fill values.
5. Test: `python backup.py`

### Config notes
- `SOURCE_DIRS`: comma‑separated folders. Supports network paths (e.g., `\\\\SERVER\\Share`).
- `FILE_GLOB`: defaults to `**/*.csv` (recursive). Use `*.csv` for flat directories.
- `LOOKBACK_DAYS`: files with modified time ≥ now‑N days.
- `DROPBOX_DEST_FOLDER`: must begin with `/`. If your app has **App folder** access, `/` refers to the app‑scoped root.
- `RETAIN_DAYS`: set `0` to disable deletion.
- `DRY_RUN`: `true` to simulate.

### Security
Prefer storing `DROPBOX_ACCESS_TOKEN` in your OS secret store or CI secret rather than plain `.env`.

### Troubleshooting
- **Auth failed**: confirm token scopes and that the token hasn’t expired.
- **Nothing to back up**: widen `LOOKBACK_DAYS` or verify paths/glob.
- **Permission denied**: ensure the Dropbox app has rights to `DROPBOX_DEST_FOLDER`.

## excel-report-generator

A small but mighty CLI that turns your CSV/TSV/JSON/Parquet files into a well-formatted Excel workbook with optional summary pivots and charts.

### Quick start

```bash
python -m venv .venv && . .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt

# Run with a folder of data files
python report_generator.py -i example_data -o report.xlsx

# With a config for pivots
python report_generator.py -i example_data -o report.xlsx -c config.example.yaml
```

Open `report.xlsx` — you’ll see one sheet per file and any summary sheets you configured.

### Inputs
- **Files**: `.csv`, `.tsv`, `.json`, `.parquet`
- **Directories**: pass a folder and we’ll ingest all supported files inside.

### Options
- `--input, -i` one or more files/folders (required)
- `--out, -o` output .xlsx path (required)
- `--config, -c` YAML defining date columns and pivots (optional)
- `--sheet-prefix` prefix sheet names with a label (optional)
- `--no-table` disable Excel table styling (optional)
- `--no-autofilter` disable header filters (optional)
- `--no-freeze` don’t freeze header row (optional)

### Pivots & Charts
Define pivots in YAML. Each pivot builds a new "Summary - <name>" sheet. Minimal example:

```yaml
pivots:
  - name: Revenue by Month
    source: sales_q1
    index: [order_month]
    columns: [region]
    values: revenue
    aggfunc: sum
    fillna: 0
    margins: true
```

> Tip: Precompute helper columns (like `order_month`) in your CSV or add them to your raw data source.

### Formatting smarts
- **Column widths** sized to content
- **Number formats** auto-detected (currency for columns like `amount`, `revenue`, etc.)
- **Dates** shown as `yyyy-mm-dd`
- **Frozen header + filters** by default

### Extending
- Add validation, custom conditional formatting, or more chart types in `report_generator.py`.
- Make it a package + console entrypoint if you want to `pipx install` it later.

### Troubleshooting
- If you see `PyYAML is not installed`, run `pip install pyyaml` or remove the `--config` flag.
- If Excel complains about sheet names, we truncate to 31 chars.
