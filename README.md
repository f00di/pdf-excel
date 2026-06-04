# PDF Excel

PDF Excel converts readable PDF content into Excel workbooks. It includes a browser app for GitHub Pages and a Python command line tool.

The online app runs fully in the browser and downloads an `.xlsx` file. The CLI uses `pdfplumber` to extract PDF tables and `openpyxl` to write `.xlsx` files. If a page does not contain a detected table, both paths can fall back to writing the page text as a one-column sheet.

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

Convert a PDF:

```sh
pdf-excel path/to/input.pdf -o output.xlsx
```

Or run the module directly:

```sh
python -m pdf_excel path/to/input.pdf --overwrite
```

## Command Options

```text
usage: pdf-excel [-h] [-o OUTPUT] [--overwrite] [--no-text-fallback] pdf
```

- `pdf`: source PDF file.
- `-o, --output`: output `.xlsx` path. Defaults to the input filename with `.xlsx`.
- `--overwrite`: replace an existing output file.
- `--no-text-fallback`: only export detected tables.

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
