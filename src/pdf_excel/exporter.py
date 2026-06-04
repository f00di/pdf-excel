"""Export extracted transactions and append summaries."""

from __future__ import annotations

import csv
import json
from datetime import datetime
from io import StringIO
from typing import Any

from pdf_excel.models import ExtractedTransaction, PDFExtractionResult, WriteSummary


def transactions_to_rows(
    transactions: list[ExtractedTransaction],
) -> list[dict[str, Any]]:
    """Convert transactions to serializable dictionaries."""

    return [transaction.to_dict() for transaction in transactions]


def transactions_to_csv(transactions: list[ExtractedTransaction]) -> bytes:
    """Serialize extracted transactions to CSV bytes."""

    rows = transactions_to_rows(transactions)
    buffer = StringIO()
    if not rows:
        return b""
    writer = csv.DictWriter(buffer, fieldnames=list(rows[0]))
    writer.writeheader()
    writer.writerows(rows)
    return buffer.getvalue().encode("utf-8")


def transactions_to_json(transactions: list[ExtractedTransaction]) -> bytes:
    """Serialize extracted transactions to JSON bytes."""

    return json.dumps(transactions_to_rows(transactions), indent=2, default=str).encode(
        "utf-8"
    )


def build_summary_report(
    pdf_results: list[PDFExtractionResult],
    summary: WriteSummary | None,
    validation_warnings: list[str],
) -> bytes:
    """Build a plain-text append summary report."""

    lines = [
        "PDF to Excel Append Summary",
        f"Created: {datetime.now().isoformat(timespec='seconds')}",
        "",
        "Source PDFs:",
    ]
    for result in pdf_results:
        lines.extend(
            [
                f"- {result.file_name}",
                f"  OCR used: {'yes' if result.ocr_used else 'no'}",
                f"  Transactions extracted: {len(result.transactions)}",
            ]
        )
        for warning in result.warnings:
            lines.append(f"  Warning: {warning}")

    if summary:
        lines.extend(
            [
                "",
                "Excel Output:",
                f"- Source Excel: {summary.source_excel_name}",
                f"- Output file: {summary.output_file_name}",
                f"- Target sheets: {', '.join(summary.target_sheets)}",
                f"- Rows added: {summary.rows_added}",
                f"- Rows skipped: {summary.rows_skipped}",
                f"- Backup created: {'yes' if summary.backup_created else 'no'}",
            ]
        )
        for warning in summary.warnings:
            lines.append(f"- Write warning: {warning}")

    if validation_warnings:
        lines.extend(["", "Validation Warnings:"])
        for warning in validation_warnings:
            lines.append(f"- {warning}")

    return "\n".join(lines).encode("utf-8")

