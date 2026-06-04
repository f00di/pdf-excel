"""High-level workflow for appending statement PDFs into Excel workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from pdf_excel.converter import PdfExcelError
from pdf_excel.mapping import suggest_mapping
from pdf_excel.models import (
    AppendOperation,
    MappingSuggestion,
    PDFExtractionResult,
    ValidationIssue,
    WorkbookAnalysis,
    WriteSummary,
)
from pdf_excel.statement_extractor import extract_pdf_file
from pdf_excel.validation import validate_append_plan
from pdf_excel.workbook_analyzer import analyze_excel_file
from pdf_excel.workbook_writer import append_transactions_to_workbook


WRITE_MODES = {"append", "skip_duplicates", "overwrite"}


@dataclass(frozen=True)
class AppendResult:
    """Summary for a completed PDF transaction append workflow."""

    output_path: Path
    source_excel_path: Path
    pdf_results: list[PDFExtractionResult]
    analysis: WorkbookAnalysis
    mappings: list[MappingSuggestion]
    validation_issues: list[ValidationIssue]
    summary: WriteSummary

    @property
    def transaction_count(self) -> int:
        return sum(len(result.transactions) for result in self.pdf_results)


def append_pdf_transactions_to_excel(
    pdf_paths: Iterable[str | Path],
    excel_path: str | Path,
    *,
    output_path: str | Path | None = None,
    sheets: Iterable[str] | None = None,
    write_mode: str = "append",
    overwrite: bool = False,
    use_ocr: bool = False,
    pdf_password: str | None = None,
    excel_password: str | None = None,
    highlight_new_cells: bool = True,
    use_ai_mapping: bool = False,
) -> AppendResult:
    """Extract statement transactions from PDFs and append them to a workbook."""

    pdf_file_paths = [_validate_pdf_path(path) for path in pdf_paths]
    if not pdf_file_paths:
        raise PdfExcelError("At least one PDF path is required.")

    workbook_path = _validate_workbook_path(excel_path)
    if write_mode not in WRITE_MODES:
        raise PdfExcelError(
            f"Unsupported write mode {write_mode!r}. Choose one of: "
            f"{', '.join(sorted(WRITE_MODES))}."
        )

    destination_path = _resolve_output_path(workbook_path, output_path)
    if destination_path.exists() and not overwrite:
        raise PdfExcelError(
            f"{destination_path} already exists. Use --overwrite to replace it."
        )

    excel_bytes = workbook_path.read_bytes()
    analysis = analyze_excel_file(
        file_bytes=excel_bytes,
        file_name=workbook_path.name,
        password=excel_password,
    )
    if not analysis.writable:
        raise PdfExcelError(
            "This workbook can be previewed but not written. Save it as .xlsx first."
        )

    pdf_results = [
        extract_pdf_file(
            file_name=pdf_path.name,
            file_bytes=pdf_path.read_bytes(),
            use_ocr=use_ocr,
            password=pdf_password,
        )
        for pdf_path in pdf_file_paths
    ]
    transactions = [
        transaction
        for result in pdf_results
        for transaction in result.transactions
    ]
    if not transactions:
        warnings = [
            warning
            for result in pdf_results
            for warning in result.warnings
        ]
        detail = f" {' '.join(warnings[:3])}" if warnings else ""
        raise PdfExcelError(f"No transaction rows were extracted from the PDF(s).{detail}")

    target_sheets = _resolve_target_sheets(analysis, sheets)
    mappings: list[MappingSuggestion] = []
    validation_issues: list[ValidationIssue] = []
    operations: list[AppendOperation] = []

    for sheet_name in target_sheets:
        sheet = analysis.sheets[sheet_name]
        mapping = suggest_mapping(
            sheet,
            transactions,
            use_ai=use_ai_mapping,
        )
        mappings.append(mapping)
        issues = validate_append_plan(
            transactions=transactions,
            sheet=sheet,
            mapping=mapping,
            existing_rows=sheet.preview_rows,
        )
        validation_issues.extend(issues)
        operations.append(
            AppendOperation(
                target_sheet=sheet_name,
                header_row=mapping.header_row,
                start_row=mapping.start_row,
                field_mapping=mapping.field_mapping,
                write_mode=write_mode,
                highlight_new_cells=highlight_new_cells,
            )
        )

    errors = [issue for issue in validation_issues if issue.severity == "error"]
    if errors:
        details = " ".join(issue.message for issue in errors[:4])
        raise PdfExcelError(f"Append validation failed. {details}")

    output_bytes, _, summary = append_transactions_to_workbook(
        excel_bytes=excel_bytes,
        source_excel_name=workbook_path.name,
        transactions=transactions,
        operations=operations,
        password=excel_password,
    )
    summary.output_file_name = destination_path.name

    destination_path.parent.mkdir(parents=True, exist_ok=True)
    destination_path.write_bytes(output_bytes)

    return AppendResult(
        output_path=destination_path,
        source_excel_path=workbook_path,
        pdf_results=pdf_results,
        analysis=analysis,
        mappings=mappings,
        validation_issues=validation_issues,
        summary=summary,
    )


def _validate_pdf_path(pdf_path: str | Path) -> Path:
    path = Path(pdf_path).expanduser().resolve()
    if not path.exists():
        raise PdfExcelError(f"{path} does not exist.")
    if not path.is_file():
        raise PdfExcelError(f"{path} is not a file.")
    if path.suffix.lower() != ".pdf":
        raise PdfExcelError(f"{path} is not a PDF file.")
    return path


def _validate_workbook_path(workbook_path: str | Path) -> Path:
    path = Path(workbook_path).expanduser().resolve()
    if not path.exists():
        raise PdfExcelError(f"{path} does not exist.")
    if not path.is_file():
        raise PdfExcelError(f"{path} is not a file.")
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise PdfExcelError(f"{path} is not a supported Excel workbook.")
    return path


def _resolve_output_path(
    source_path: Path,
    output_path: str | Path | None,
) -> Path:
    if output_path is None:
        return source_path.with_name(f"{source_path.stem}_updated{source_path.suffix}")

    path = Path(output_path).expanduser()
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise PdfExcelError("Output path must end with .xlsx or .xlsm.")
    return path.resolve()


def _resolve_target_sheets(
    analysis: WorkbookAnalysis,
    sheets: Iterable[str] | None,
) -> list[str]:
    if sheets:
        target_sheets = [sheet for sheet in sheets if sheet]
    elif "BoA" in analysis.sheets:
        target_sheets = ["BoA"]
    else:
        target_sheets = list(analysis.sheets)[:1]

    if not target_sheets:
        raise PdfExcelError("At least one target sheet is required.")

    missing = [sheet for sheet in target_sheets if sheet not in analysis.sheets]
    if missing:
        available = ", ".join(analysis.sheets)
        raise PdfExcelError(
            f"Target sheet not found: {', '.join(missing)}. Available sheets: {available}."
        )
    return target_sheets

