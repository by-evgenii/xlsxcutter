# XLSX Cutter

XLSX Cutter is a desktop utility for slicing huge Excel workbooks into smaller,
manageable files. The refreshed interface also adds conversion tools so you can
turn worksheets into CSV, JSON or ODS documents and reconstruct large CSV/JSON
tables back into XLSX workbooks automatically.

## Features

- Choose any sheet from an Excel workbook directly in the UI.
- Split a sheet into multiple files while keeping the header row.
- Export each chunk as XLSX, CSV, JSON and/or ODS (OpenDocument Spreadsheet).
- Rebuild large CSV or JSON tables into a multi-sheet XLSX workbook when the
  data exceeds Excel's row limit.
- Modernised Tk/ttk interface with status feedback.
- Ready for packaging with PyInstaller on Windows (.exe), Linux (ELF) and
  macOS (Mach-O/.app).

## Requirements

- Python 3.10+
- The packages listed in [`requirements.txt`](requirements.txt)

Install the dependencies with:

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\Scripts\activate
pip install -r requirements.txt
```

## Running the app

```bash
python xlsxcutter.py
```

## Splitting a worksheet

1. Launch the app and open the **Split Excel** tab.
2. Click **Browse** next to *Workbook* and select an `.xlsx` file. The sheet
   drop-down automatically populates.
3. Choose the sheet to process and the number of rows per output file.
4. Select the output folder and tick one or more export formats.
5. Press **Split** to generate the files. Each export reuses the sheet header
   and appends an incrementing suffix (`_part001`, `_part002`, ...).

## Building XLSX files from CSV/JSON

1. Switch to the **Build XLSX from CSV/JSON** tab.
2. Choose a `.csv` or `.json` table. The output filename is pre-filled for you.
3. Set the maximum rows per Excel sheet. The value is capped at Excel's limit of
   1,048,576 rows.
4. Click **Build XLSX**. The tool creates as many sheets as necessary to fit the
   entire dataset.

## Packaging into executables

XLSX Cutter is designed to work with PyInstaller, making it easy to distribute
native executables for the major platforms.

```bash
pyinstaller xlsxcutter.py --name xlsxcutter --onefile --windowed
```

The command above produces:

- Windows: `dist/xlsxcutter.exe`
- Linux: `dist/xlsxcutter` (ELF executable)
- macOS: `dist/xlsxcutter` (Mach-O binary). Use `--onedir` instead of `--onefile`
  if you prefer a `.app` bundle: `pyinstaller xlsxcutter.py --name xlsxcutter --onedir --windowed`.

## Automated releases

A GitHub Actions workflow (`.github/workflows/release.yml`) builds PyInstaller
artifacts for Windows, macOS and Linux whenever you push a tag that starts with
`v` (for example `v1.0.0`). The job uploads the executables to the matching
GitHub release automatically.

To cut a release:

```bash
git tag v1.0.0
git push origin v1.0.0
```

Once the workflow completes, download the platform-specific binaries from the
release page.

## License

This project is released under the terms of the MIT license. See
[`LICENSE`](LICENSE) for details.
