# CSV Cleaner: merge messy CSV files into one clean table

Command-line tool in pure Python (standard library; `openpyxl` only if you want `.xlsx` output).

## What it does
- Reads many CSV files at once, auto-detects **encoding** (UTF-8 / Windows-1251) and **delimiter** (`,` `;` tab)
- Normalises column names: `E-mail`, `mail`, `Email Address` → `email`; `Full Name`, `Client` → `name`; `Tel`, `Mobile` → `phone`
- Trims extra spaces, lowercases e-mails, tidies phone numbers
- Removes empty rows and **duplicates** (by `--key` columns, or by the whole row)
- Adds a `source_file` column so you can see where each row came from
- Saves to **CSV** (opens correctly in Excel) or **XLSX**, and prints a short report

## Usage
```bash
python clean_merge.py sample_data -o result.csv --key email
python clean_merge.py a.csv b.csv -o result.xlsx --key email phone
```
Example report:
```
Files processed:     3
Rows read:           10
Empty rows removed:  2
Duplicates removed:  3
Rows saved:          5  ->  result.csv
```

## Tests
```bash
python -m unittest discover -s tests -v
```
