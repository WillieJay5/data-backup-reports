from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

import pandas as pd
import numpy as np

try:
    import yaml  # type: ignore
except Exception:  # pragma: no cover
    yaml = None


SUPPORTED_EXTS = {".csv", ".tsv", ".json", ".parquet"}


def read_frame(path: Path, date_cols: Optional[List[str]] = None) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path, parse_dates=date_cols or True)
    if ext == ".tsv":
        return pd.read_csv(path, sep="\t", parse_dates=date_cols or True)
    if ext == ".json":
        return pd.read_json(path)
    if ext == ".parquet":
        return pd.read_parquet(path)
    raise ValueError(f"Unsupported file extension: {ext}")


def auto_number_format(series: pd.Series) -> Optional[str]:
    if pd.api.types.is_datetime64_any_dtype(series):
        return "yyyy-mm-dd"
    if pd.api.types.is_bool_dtype(series):
        return None
    if pd.api.types.is_integer_dtype(series) or pd.api.types.is_float_dtype(series):
        # detect currency-ish by column name
        name = str(series.name).lower()
        if any(k in name for k in ["amount", "revenue", "price", "cost", "total", "sales"]):
            return "$#,##0.00"
        # default numeric format with thousand separator
        return "#,##0.00"
    return None


def best_width(series: pd.Series, max_width: int = 60) -> int:
    # Estimate width based on max string length in column (including header)
    s = series.astype(str)
    maxlen = max([len(str(series.name))] + s.map(len).tolist())
    # Excel uses ~1.2 scale; add a small buffer
    return min(maxlen + 2, max_width)


def write_df_to_sheet(df: pd.DataFrame, writer: pd.ExcelWriter, sheet_name: str, *,
                      make_table: bool = True,
                      table_style: str = "Table Style Light 9",
                      autofilter: bool = True,
                      freeze_header: bool = True) -> None:
    df = df.copy()
    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1 if make_table else 0)
    wb = writer.book
    ws = writer.sheets[sheet_name]

    # header format
    header_fmt = wb.add_format({"bold": True, "text_wrap": False, "valign": "bottom"})

    # Write our header manually if making a table (we left a row for it)
    if make_table:
        for col_idx, col_name in enumerate(df.columns):
            ws.write(0, col_idx, col_name, header_fmt)
        # Add table
        last_row = len(df)
        last_col = len(df.columns) - 1
        ws.add_table(0, 0, last_row, last_col, {
            "name": f"tbl_{sheet_name[:20].replace(' ', '_')}",
            "style": table_style,
            "columns": [{"header": str(c)} for c in df.columns],
            "autofilter": True,
        })
    else:
        # Apply header format to first row that pandas wrote
        for col_idx in range(len(df.columns)):
            ws.write(0, col_idx, df.columns[col_idx], header_fmt)

    # Column widths + number formats
    for col_idx, col in enumerate(df.columns):
        fmt_str = auto_number_format(df[col])
        fmt = wb.add_format({"num_format": fmt_str}) if fmt_str else None
        width = best_width(df[col])
        ws.set_column(col_idx, col_idx, width, fmt)

    if autofilter and not make_table:
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)

    if freeze_header:
        ws.freeze_panes(1 if make_table else 1, 0)


def build_pivots(summary_cfg: List[Dict[str, Any]],
                 frames: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    pivots = {}
    for p in summary_cfg:
        src = p.get("source")
        if src not in frames:
            print(f"[warn] pivot source not found: {src}", file=sys.stderr)
            continue
        df = frames[src]
        index = p.get("index")
        columns = p.get("columns")
        values = p.get("values")
        aggfunc = p.get("aggfunc", "sum")
        fillna = p.get("fillna", 0)
        # Support multi-value aggregations
        pivot = pd.pivot_table(df, index=index, columns=columns, values=values,
                               aggfunc=aggfunc, fill_value=fillna, dropna=False, margins=p.get("margins", False))
        # Flatten columns if multiindex
        if isinstance(pivot.columns, pd.MultiIndex):
            pivot.columns = [" ".join([str(c) for c in tup if str(c) != ""]) for tup in pivot.columns]
        pivots[p.get("name", f"pivot_{src}")] = pivot.reset_index()
    return pivots


def add_chart_from_table(writer: pd.ExcelWriter, sheet_name: str, table_range: str,
                         chart_type: str = "column", title: Optional[str] = None) -> None:
    wb = writer.book
    ws = writer.sheets[sheet_name]
    chart = wb.add_chart({"type": chart_type})
    # Assume first column is categories, rest are series
    first_col_letter = table_range.split(":")[0].rstrip("0123456789")
    start_row = int("".join([c for c in table_range.split(":")[0] if c.isdigit()]))
    # We will infer categories and series based on the Excel range; however building robustly
    # would require parsing A1 notation—keep it simple with XlsxWriter built-ins:
    chart.add_series({
        "categories": f"={sheet_name}!{table_range.split(':')[0]}:{first_col_letter}{table_range.split(':')[1][len(first_col_letter):]}",
        "values": f"={sheet_name}!{chr(ord(first_col_letter)+1)}{start_row}:{table_range.split(':')[1]}",
    })
    if title:
        chart.set_title({"name": title})
    ws.insert_chart("B2", chart)


def load_config(path: Optional[Path]) -> Dict[str, Any]:
    if not path:
        return {}
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")
    if yaml is None:
        raise RuntimeError("PyYAML is not installed; cannot load YAML config.")
    with path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def gather_inputs(inputs: List[str]) -> List[Path]:
    paths: List[Path] = []
    for inp in inputs:
        p = Path(inp)
        if p.is_dir():
            for ext in SUPPORTED_EXTS:
                paths.extend(sorted(p.glob(f"*{ext}")))
        elif p.is_file():
            paths.append(p)
        else:
            print(f"[warn] input not found: {inp}", file=sys.stderr)
    return paths


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Generate a polished Excel workbook from data files")
    ap.add_argument("--input", "-i", nargs="+", required=True, help="Input files and/or directories")
    ap.add_argument("--out", "-o", required=True, help="Output .xlsx path")
    ap.add_argument("--config", "-c", help="Optional YAML config for pivots & options")
    ap.add_argument("--sheet-prefix", default="", help="Prefix to apply to each sheet name")
    ap.add_argument("--no-table", action="store_true", help="Disable Excel table styling")
    ap.add_argument("--no-autofilter", action="store_true")
    ap.add_argument("--no-freeze", action="store_true")
    args = ap.parse_args(argv)

    cfg = load_config(Path(args.config)) if args.config else {}

    input_paths = gather_inputs(args.input)
    if not input_paths:
        print("No input files found.")
        return 2

    writer = pd.ExcelWriter(args.out, engine="xlsxwriter")
    frames: Dict[str, pd.DataFrame] = {}

    # Write each source to its own sheet
    for p in input_paths:
        name = p.stem[:31]  # Excel sheet name limit
        sheet_name = f"{args.sheet_prefix}{name}"
        date_cols = None
        # allow optional date parsing hints via config: sources: [{name: sales_q1, date_cols:["order_date"]}, ...]
        for s in (cfg.get("sources") or []):
            if s.get("name") == name and s.get("date_cols"):
                date_cols = s.get("date_cols")
                break
        df = read_frame(p, date_cols=date_cols)
        frames[name] = df
        write_df_to_sheet(
            df,
            writer,
            sheet_name,
            make_table=not args.no_table,
            table_style=cfg.get("table_style", "Table Style Light 9"),
            autofilter=not args.no_autofilter,
            freeze_header=not args.no_freeze,
        )

    # Optional summary pivots
    summary_cfg = cfg.get("pivots")
    if summary_cfg:
        pivots = build_pivots(summary_cfg, frames)
        for piv_name, piv_df in pivots.items():
            sname = f"Summary - {piv_name}"[:31]
            write_df_to_sheet(piv_df, writer, sname, make_table=True)
            # Optional chart per pivot
            # We don't know final written range until after write; reconstruct basic A1 range
            rows, cols = piv_df.shape
            # A1 helpers
            def a1(col_idx: int, row_idx: int) -> str:
                col_letters = ""
                col_idx += 1
                while col_idx:
                    col_idx, remainder = divmod(col_idx - 1, 26)
                    col_letters = chr(65 + remainder) + col_letters
                return f"{col_letters}{row_idx+1}"
            start = a1(0, 0)
            end = a1(cols - 1, rows)
            add_chart_from_table(writer, sname, f"{start}:{end}", chart_type="column", title=piv_name)

    writer.close()
    print(f"Wrote {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
