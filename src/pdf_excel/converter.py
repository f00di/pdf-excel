"""Core PDF to Excel conversion logic."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


INVALID_SHEET_TITLE_CHARS = set("[]:*?/\\")
MAX_SHEET_TITLE_LENGTH = 31


class PdfExcelError(RuntimeError):
    """Base exception for expected PDF Excel failures."""


class MissingDependencyError(PdfExcelError):
    """Raised when a runtime conversion dependency is not installed."""


@dataclass(frozen=True)
class ExtractedTable:
    """A table or fallback text block extracted from a PDF page."""

    page_number: int
    index: int
    rows: list[list[str]]
    source: str = "table"


@dataclass(frozen=True)
class ConversionResult:
    """Summary of a completed conversion."""

    output_path: Path
    page_count: int
    sheet_count: int
    row_count: int


def convert_pdf_to_excel(
    pdf_path: str | Path,
    *,
    output_path: str | Path | None = None,
    overwrite: bool = False,
    text_fallback: bool = True,
) -> ConversionResult:
    """Extract PDF tables and write them to an Excel workbook."""

    source_path = _validate_pdf_path(pdf_path)
    destination_path = _resolve_output_path(source_path, output_path)

    if destination_path.exists() and not overwrite:
        raise PdfExcelError(
            f"{destination_path} already exists. Use --overwrite to replace it."
        )

    destination_path.parent.mkdir(parents=True, exist_ok=True)

    tables, page_count = extract_tables(source_path, text_fallback=text_fallback)
    if not tables:
        raise PdfExcelError(
            "No tables were found in the PDF. "
            "Try a PDF with embedded text or omit --no-text-fallback."
        )

    row_count = write_workbook(destination_path, tables)
    return ConversionResult(
        output_path=destination_path,
        page_count=page_count,
        sheet_count=len(tables),
        row_count=row_count,
    )


def extract_tables(
    pdf_path: Path,
    *,
    text_fallback: bool = True,
) -> tuple[list[ExtractedTable], int]:
    """Extract table rows from every page in a PDF."""

    try:
        import pdfplumber
    except ImportError as exc:
        raise MissingDependencyError(
            "Missing PDF extraction dependency. Install the project with "
            "`python -m pip install -e .` and run the command again."
        ) from exc

    extracted: list[ExtractedTable] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            page_tables = page.extract_tables() or []
            table_index = 1

            for raw_table in page_tables:
                rows = _clean_rows(raw_table)
                if rows:
                    extracted.append(
                        ExtractedTable(
                            page_number=page_number,
                            index=table_index,
                            rows=rows,
                        )
                    )
                    table_index += 1

            if table_index == 1 and text_fallback:
                rows = _extract_text_rows(page.extract_text())
                if rows:
                    extracted.append(
                        ExtractedTable(
                            page_number=page_number,
                            index=1,
                            rows=rows,
                            source="text",
                        )
                    )

        return extracted, len(pdf.pages)


def write_workbook(path: Path, tables: Iterable[ExtractedTable]) -> int:
    """Write extracted tables to an Excel workbook and return the row count."""

    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise MissingDependencyError(
            "Missing Excel writing dependency. Install the project with "
            "`python -m pip install -e .` and run the command again."
        ) from exc

    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    existing_titles: set[str] = set()
    total_rows = 0

    for table in tables:
        title = _unique_sheet_title(
            existing_titles,
            f"Page {table.page_number} {table.source.title()} {table.index}",
        )
        worksheet = workbook.create_sheet(title=title)

        for row in table.rows:
            worksheet.append(row)
            total_rows += 1

    workbook.save(path)
    return total_rows


def _validate_pdf_path(pdf_path: str | Path) -> Path:
    path = Path(pdf_path).expanduser().resolve()

    if not path.exists():
        raise PdfExcelError(f"{path} does not exist.")
    if not path.is_file():
        raise PdfExcelError(f"{path} is not a file.")
    if path.suffix.lower() != ".pdf":
        raise PdfExcelError(f"{path} is not a PDF file.")

    return path


def _resolve_output_path(
    source_path: Path,
    output_path: str | Path | None,
) -> Path:
    if output_path is None:
        return source_path.with_suffix(".xlsx")

    path = Path(output_path).expanduser()
    if path.suffix.lower() != ".xlsx":
        raise PdfExcelError("Output path must end with .xlsx.")

    return path.resolve()


def _clean_rows(rows: Iterable[Iterable[object | None]]) -> list[list[str]]:
    cleaned: list[list[str]] = []

    for row in rows:
        cleaned_row = [_clean_cell(cell) for cell in row]
        if any(cell for cell in cleaned_row):
            cleaned.append(cleaned_row)

    return cleaned


def _clean_cell(cell: object | None) -> str:
    if cell is None:
        return ""
    return " ".join(str(cell).split())


def _extract_text_rows(text: str | None) -> list[list[str]]:
    if not text:
        return []
    return [[line.strip()] for line in text.splitlines() if line.strip()]


def _sanitize_sheet_title(title: str) -> str:
    sanitized = "".join(
        "_" if char in INVALID_SHEET_TITLE_CHARS else char for char in title
    ).strip()
    if not sanitized:
        sanitized = "Sheet"
    return sanitized[:MAX_SHEET_TITLE_LENGTH]


def _unique_sheet_title(existing_titles: set[str], title: str) -> str:
    base = _sanitize_sheet_title(title)
    candidate = base
    counter = 2

    while candidate in existing_titles:
        suffix = f" {counter}"
        candidate = f"{base[: MAX_SHEET_TITLE_LENGTH - len(suffix)]}{suffix}"
        counter += 1

    existing_titles.add(candidate)
    return candidate
