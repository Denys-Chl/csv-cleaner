#!/usr/bin/env python3
"""Merge messy CSV files into one clean table.

What it does:
  * reads many CSV files (auto-detects encoding and delimiter: , ; or tab)
  * normalises column names ("E-mail", "Full Name" -> email, name)
  * trims spaces, lowercases e-mails, tidies phone numbers
  * removes empty rows and duplicates (by key column(s) or by the whole row)
  * saves the result as CSV or Excel and prints a short report

Usage:
  python clean_merge.py sample_data -o result.csv --key email
  python clean_merge.py a.csv b.csv -o result.xlsx --key email phone
"""
import argparse
import csv
import re
import sys
from pathlib import Path

ALIASES = {
    "e_mail": "email", "mail": "email", "email_address": "email",
    "tel": "phone", "telephone": "phone", "mobile": "phone", "phone_number": "phone",
    "full_name": "name", "fullname": "name", "customer": "name", "client": "name",
}
ENCODINGS = ("utf-8-sig", "cp1251", "latin-1")


def normalize_header(h: str) -> str:
    h = re.sub(r"[^0-9a-zA-Zа-яА-ЯіїєґІЇЄҐ]+", "_", (h or "").strip().lower()).strip("_")
    return ALIASES.get(h, h or "column")


def clean_value(column: str, value: str) -> str:
    value = re.sub(r"\s+", " ", (value or "").strip())
    if column == "email":
        value = value.lower()
    elif column == "phone":
        digits = re.sub(r"\D", "", value)
        if digits:
            value = ("+" if value.lstrip().startswith("+") else "") + digits
    return value


def read_csv(path: Path):
    """Returns (headers, rows) where rows is a list of dicts. Tries several encodings."""
    raw = path.read_bytes()
    text = None
    for enc in ENCODINGS:
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    reader = csv.reader(text.splitlines(), dialect)
    rows = [r for r in reader]
    if not rows:
        return [], []
    headers = [normalize_header(h) for h in rows[0]]
    # make duplicate column names unique: phone, phone_2
    seen = {}
    for i, h in enumerate(headers):
        seen[h] = seen.get(h, 0) + 1
        if seen[h] > 1:
            headers[i] = f"{h}_{seen[h]}"
    data = []
    for r in rows[1:]:
        r = (r + [""] * len(headers))[: len(headers)]
        data.append(dict(zip(headers, r)))
    return headers, data


def merge(files, key=None):
    """Merge files. Returns (headers, rows, stats)."""
    stats = {"files": 0, "rows_in": 0, "empty_removed": 0, "duplicates_removed": 0, "rows_out": 0}
    headers = []
    out = []
    seen_keys = set()
    for path in files:
        file_headers, rows = read_csv(Path(path))
        stats["files"] += 1
        for h in file_headers:
            if h not in headers:
                headers.append(h)
        for row in rows:
            stats["rows_in"] += 1
            row = {c: clean_value(c, v) for c, v in row.items()}
            if not any(row.values()):
                stats["empty_removed"] += 1
                continue
            if key:
                k = tuple(row.get(c, "") for c in key)
                if not any(k):          # no key values: can't compare, keep the row
                    k = None
            else:
                k = tuple(sorted(row.items()))
            if k is not None:
                if k in seen_keys:
                    stats["duplicates_removed"] += 1
                    continue
                seen_keys.add(k)
            row["source_file"] = Path(path).name
            out.append(row)
    headers.append("source_file")
    stats["rows_out"] = len(out)
    return headers, out, stats


def save(path: Path, headers, rows):
    if path.suffix.lower() == ".xlsx":
        from openpyxl import Workbook  # optional dependency
        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for r in rows:
            ws.append([r.get(h, "") for h in headers])
        for col in ws.columns:  # auto column width
            ws.column_dimensions[col[0].column_letter].width = min(40, max(len(str(c.value or "")) for c in col) + 2)
        ws.freeze_panes = "A2"
        wb.save(path)
    else:
        with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig: Excel opens it correctly
            w = csv.DictWriter(f, fieldnames=headers)
            w.writeheader()
            w.writerows(rows)


def collect_files(inputs):
    files = []
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            files += sorted(p.glob("*.csv"))
        elif p.is_file():
            files.append(p)
        else:
            sys.exit(f"Not found: {item}")
    if not files:
        sys.exit("No CSV files found.")
    return files


def main(argv=None):
    ap = argparse.ArgumentParser(description="Merge and clean messy CSV files.")
    ap.add_argument("inputs", nargs="+", help="CSV files or folders with CSV files")
    ap.add_argument("-o", "--output", default="merged.csv", help="result file (.csv or .xlsx)")
    ap.add_argument("--key", nargs="+", help="column(s) that identify a duplicate, e.g. --key email")
    args = ap.parse_args(argv)

    files = collect_files(args.inputs)
    headers, rows, stats = merge(files, key=args.key)
    save(Path(args.output), headers, rows)
    print(f"Files processed:     {stats['files']}")
    print(f"Rows read:           {stats['rows_in']}")
    print(f"Empty rows removed:  {stats['empty_removed']}")
    print(f"Duplicates removed:  {stats['duplicates_removed']}")
    print(f"Rows saved:          {stats['rows_out']}  ->  {args.output}")


if __name__ == "__main__":
    main()
