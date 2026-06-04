"""Shared data models for statement extraction and workbook appends."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


PDF_FIELDS = [
    "transaction_date",
    "description",
    "debit",
    "credit",
    "balance",
    "reference_number",
    "account_number",
    "statement_period",
    "opening_balance",
    "closing_balance",
    "category",
    "notes",
]


@dataclass
class ExtractedTransaction:
    """A normalized transaction row extracted from a statement PDF."""

    transaction_date: str = ""
    description: str = ""
    debit: float | None = None
    credit: float | None = None
    balance: float | None = None
    reference_number: str = ""
    account_number: str = ""
    statement_period: str = ""
    opening_balance: float | None = None
    closing_balance: float | None = None
    category: str = ""
    notes: str = ""
    source_pdf: str = ""
    source_page: int | None = None
    raw_text: str = ""
    confidence: float = 0.0
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class PDFExtractionResult:
    """Statement extraction details for one PDF."""

    file_name: str
    raw_text: str = ""
    tables: list[dict[str, Any]] = field(default_factory=list)
    transactions: list[ExtractedTransaction] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    ocr_used: bool = False

    def transaction_dicts(self) -> list[dict[str, Any]]:
        return [transaction.to_dict() for transaction in self.transactions]


@dataclass
class SheetStructure:
    """Detected structure for one workbook sheet."""

    sheet_name: str
    max_row: int
    max_column: int
    header_row: int
    headers: dict[str, str]
    preview_rows: list[dict[str, Any]]
    empty_areas: list[str]
    tables: list[dict[str, Any]]
    merged_cells: list[str]
    data_validations: list[str]
    formulas: list[dict[str, str]]
    column_formats: dict[str, dict[str, Any]]
    column_widths: dict[str, float]
    next_row: int
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class WorkbookAnalysis:
    """Summary of workbook sheets and write suitability."""

    file_name: str
    file_type: str
    sheets: dict[str, SheetStructure]
    writable: bool = True
    warnings: list[str] = field(default_factory=list)


@dataclass
class MappingSuggestion:
    """Mapping between normalized PDF fields and Excel columns."""

    target_sheet: str
    header_row: int
    start_row: int
    field_mapping: dict[str, str]
    confidence_score: float
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class ValidationIssue:
    """Problem or warning found before appending transactions."""

    severity: str
    message: str
    row_index: int | None = None
    field: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class AppendOperation:
    """One append/write operation against a target worksheet."""

    target_sheet: str
    header_row: int
    start_row: int
    field_mapping: dict[str, str]
    write_mode: str = "append"
    highlight_new_cells: bool = True


@dataclass
class WriteSummary:
    """Details about rows and cells written to a workbook."""

    source_excel_name: str
    target_sheets: list[str]
    rows_added: int
    rows_skipped: int
    output_file_name: str
    warnings: list[str] = field(default_factory=list)
    changed_cells: list[dict[str, Any]] = field(default_factory=list)
    backup_created: bool = True

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

