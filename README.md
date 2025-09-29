# XLSX Cutter

XLSX Cutter is a desktop utility for slicing huge Excel workbooks into smaller,
manageable files. The refreshed interface also adds conversion tools so you can
turn worksheets into CSV, JSON or ODS documents and reconstruct large CSV/JSON
tables back into XLSX workbooks automatically.

## Features

- Choose any sheet from an Excel workbook directly in the UI.
- Split a sheet into multiple files while keeping the header row.
- Define the table range by selecting the start cell and optional end cell before
  splitting.
- Export each chunk as XLSX, CSV, JSON and/or ODS (OpenDocument Spreadsheet).
- Rebuild large CSV or JSON tables into a multi-sheet XLSX workbook when the
  data exceeds Excel's row limit.
- Modernised Tk/ttk interface with status feedback.
- Ready for packaging with PyInstaller on Windows (.exe), Linux (ELF) and
  macOS (Mach-O/.app).

## Requirements

- Python 3.10+
- The Python packages listed in [`requirements.txt`](requirements.txt)

The app now installs its core Python dependencies automatically the first time
you run it. If you prefer to manage packages manually (for example when working
offline), create a virtual environment and install the requirements yourself:

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
3. Choose the sheet to process, the number of rows per output file, and the
   top-left cell of the table. Provide an optional end cell to narrow the range.
4. Select the output folder and tick one or more export formats.
5. Press **Split** to generate the files. Each export reuses the sheet header
   and appends an incrementing suffix (`_part001`, `_part002`, ...). Selecting
   the ODS format automatically installs `odfpy` if it is not already available.

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

## Continuous integration builds

The GitHub Actions workflow at [`.github/workflows/build.yml`](.github/workflows/build.yml)
creates runnable artifacts for each push and pull request:

- Windows 11 (x64) standalone executable built with PyInstaller.
- macOS (ARM64) application bundle built with PyInstaller.
- Ubuntu (x64) `.deb` package containing the PyInstaller binary.

Each job uploads its artifact so you can download and test the latest build
directly from the workflow run.

## License

This project is released under the terms of the MIT license. See
[`LICENSE`](LICENSE) for details.
