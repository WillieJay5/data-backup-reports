# excel-report-generator

A small but mighty CLI that turns your CSV/TSV/JSON/Parquet files into a well-formatted Excel workbook with optional summary pivots and charts.

## Quick start

```bash
python -m venv .venv && . .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt

# Run with a folder of data files
python report_generator.py -i example_data -o report.xlsx

# With a config for pivots
python report_generator.py -i example_data -o report.xlsx -c config.example.yaml
```

Open `report.xlsx` — you’ll see one sheet per file and any summary sheets you configured.

## Inputs
- **Files**: `.csv`, `.tsv`, `.json`, `.parquet`
- **Directories**: pass a folder and we’ll ingest all supported files inside.

## Options
- `--input, -i` one or more files/folders (required)
- `--out, -o` output .xlsx path (required)
- `--config, -c` YAML defining date columns and pivots (optional)
- `--sheet-prefix` prefix sheet names with a label (optional)
- `--no-table` disable Excel table styling (optional)
- `--no-autofilter` disable header filters (optional)
- `--no-freeze` don’t freeze header row (optional)

## Pivots & Charts
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

## Formatting smarts
- **Column widths** sized to content
- **Number formats** auto-detected (currency for columns like `amount`, `revenue`, etc.)
- **Dates** shown as `yyyy-mm-dd`
- **Frozen header + filters** by default

## Extending
- Add validation, custom conditional formatting, or more chart types in `report_generator.py`.
- Make it a package + console entrypoint if you want to `pipx install` it later.

## Troubleshooting
- If you see `PyYAML is not installed`, run `pip install pyyaml` or remove the `--config` flag.
- If Excel complains about sheet names, we truncate to 31 chars.
