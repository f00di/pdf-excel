# PDF Excel

Converts PDF tables to Excel. Two parts: a browser app (`index.html` + `web/`, PDF.js + SheetJS) on https://f00di.github.io/pdf-excel/, and a Python package/CLI in `src/pdf_excel` (pdfplumber + openpyxl), including bank-statement append into existing workbooks.

- Usage and features: `README.md`. Contributing and conventions: `CONTRIBUTING.md`.
- CLI: `pdf-excel` or `python -m pdf_excel`. Tests: `python -m pytest`.
- Don't overwrite outputs unless `--overwrite` is passed. Preserve workbook styles and formulas when appending.
- Never commit real bank statements or personal financial data as fixtures.
