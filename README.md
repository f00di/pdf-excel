# PDF Excel

PDF Excel converts readable PDF content into Excel workbooks. It includes a browser app for GitHub Pages and a Python command line tool.

The online app runs fully in the browser and downloads an `.xlsx` file. The CLI uses `pdfplumber` to extract PDF tables and `openpyxl` to write `.xlsx` files. If a page does not contain a detected table, both paths can fall back to writing the page text as a one-column sheet.

The Python package also includes a bank-statement append workflow. It can extract normalized transactions from statement PDFs, analyze an existing workbook, suggest field-to-column mappings, validate the append plan, and write an updated workbook while preserving styles and formulas.

## Project Status

This project includes a GitHub Pages app, a runnable Python package, a `pdf-excel` CLI command, unit tests, and GitHub Actions CI.

## Use Online

Open the GitHub Pages site:

```text
https://f00di.github.io/pdf-excel/
```

Drop a text-based PDF into the page, convert it, and download the generated workbook.

The browser app is static HTML, CSS, and JavaScript in `index.html` and `web/`. It uses Mozilla PDF.js for browser PDF text extraction and SheetJS for workbook export.

## Features

- Run online from GitHub Pages without a server.
- Extract tables from text-based PDFs.
- Export each detected table to its own Excel sheet.
- Add text-only fallback sheets for pages without detected tables.
- Extract bank-statement transactions, including a Bank of America statement parser.
- Append extracted transactions into an existing `.xlsx` or `.xlsm` workbook.
- Detect target sheet headers, formulas, protected sheets, and the next append row.
- Suggest mappings from PDF fields to columns such as date, description, debit, credit, balance, and notes.
- Skip duplicate-looking rows or overwrite from a chosen append row.
- Preserve row styling, extend formulas where possible, highlight new cells, and add an append log sheet.
- Refuse to overwrite existing outputs unless `--overwrite` is passed.
- Run as either `pdf-excel` or `python -m pdf_excel`.

Scanned image-only PDFs usually need OCR before this tool can extract useful data.

## Publish on GitHub Pages

For a fork or a new repository, enable GitHub Pages with:

```text
Settings -> Pages -> Build and deployment -> Deploy from a branch
Branch: main
Folder: / (root)
```

After GitHub publishes the branch, the root `index.html` becomes the online converter at:

```text
https://<github-user>.github.io/pdf-excel/
```

## Getting Started

Clone the repository:

```sh
git clone <repository-url>
cd pdf-excel
```

Create a virtual environment and install the package:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e .
```

Optional capabilities can be installed with extras:

```sh
python -m pip install -e ".[ocr,encrypted,xls,ai]"
```

Convert a PDF:

```sh
pdf-excel path/to/input.pdf -o output.xlsx
```

Or run the module directly:

```sh
python -m pdf_excel path/to/input.pdf --overwrite
```

Append bank-statement transactions into an existing workbook:

```sh
pdf-excel statement.pdf --append-to budget.xlsx --sheet BoA -o budget_updated.xlsx
```

Skip duplicate-looking rows while appending:

```sh
pdf-excel statement.pdf --append-to budget.xlsx --write-mode skip_duplicates --overwrite
```

Append multiple PDFs to the default target sheet:

```sh
pdf-excel april.pdf may.pdf --append-to budget.xlsx
```

## Command Options

```text
usage: pdf-excel [-h] [-o OUTPUT] [--append-to WORKBOOK] [--sheet SHEET]
                 [--write-mode {append,skip_duplicates,overwrite}]
                 [--overwrite] [--no-text-fallback] [--ocr]
                 [--pdf-password PDF_PASSWORD]
                 [--excel-password EXCEL_PASSWORD] [--no-highlight]
                 [--use-ai-mapping]
                 pdf [pdf ...]
```

- `pdf`: source PDF file, or multiple PDFs when using `--append-to`.
- `-o, --output`: output workbook path. Defaults to the input PDF name for conversion, or `<workbook>_updated.xlsx` / `<workbook>_updated.xlsm` for append mode.
- `--append-to`: existing Excel workbook to update from statement PDF transactions.
- `--sheet`: target sheet for append mode. Repeat for multiple sheets. Defaults to `BoA` when present, otherwise the first sheet.
- `--write-mode`: append rows, skip duplicate-looking rows, or overwrite from the detected starting row.
- `--overwrite`: replace an existing output file.
- `--no-text-fallback`: only export detected tables.
- `--ocr`: try OCR fallback for scanned PDFs when optional OCR dependencies and the Tesseract system binary are installed.
- `--pdf-password`: password for encrypted statement PDFs.
- `--excel-password`: password for encrypted Excel workbooks.
- `--no-highlight`: do not highlight newly written cells in append mode.
- `--use-ai-mapping`: optionally use OpenAI mapping when `OPENAI_API_KEY` and `OPENAI_MAPPING_MODEL` are configured; otherwise deterministic mapping is used.

Append mode writes a new workbook and leaves the original workbook unchanged.

## Development

Run the tests without installing the package:

```sh
PYTHONPATH=src python3 -m unittest discover -s tests
```

Run tests after installing the package:

```sh
python -m unittest discover -s tests
```

## Repository Structure

```text
.
├── .github/              # GitHub templates and CI workflow
├── web/                  # Browser app assets
├── src/pdf_excel/        # CLI and conversion package
├── tests/                # Unit tests
├── index.html            # GitHub Pages browser app
├── CODE_OF_CONDUCT.md    # Community behavior expectations
├── CONTRIBUTING.md       # How to contribute
├── LICENSE               # MIT license
├── pyproject.toml        # Python package metadata
├── README.md             # Project overview
└── SECURITY.md           # Security reporting guidance
```

## Contributing

Contributions are welcome. Before opening a pull request, read [CONTRIBUTING.md](CONTRIBUTING.md) for the expected workflow.

Useful contributions include:

- Improvements to PDF extraction and spreadsheet export.
- Tests and sample fixtures.
- Documentation improvements.
- Bug reports with reproducible examples.
- Feature requests that describe a real workflow.

## License

This project is licensed under the [MIT License](LICENSE).
