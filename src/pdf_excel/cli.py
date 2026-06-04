"""Command line interface for PDF Excel."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from pdf_excel.append import append_pdf_transactions_to_excel
from pdf_excel.converter import PdfExcelError, convert_pdf_to_excel
from pdf_excel.workbook_analyzer import ExcelAnalysisError
from pdf_excel.workbook_writer import ExcelWriteError


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf-excel",
        description=(
            "Extract PDF tables to Excel, or append bank-statement PDF "
            "transactions into an existing workbook."
        ),
    )
    parser.add_argument(
        "pdf",
        nargs="+",
        type=Path,
        help="Path to one or more source PDF files.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help=(
            "Path for the generated workbook. Defaults to the PDF name for "
            "conversion, or <workbook>_updated.xlsx/.xlsm for append mode."
        ),
    )
    parser.add_argument(
        "--append-to",
        type=Path,
        metavar="WORKBOOK",
        help="Append extracted bank-statement transactions into this workbook.",
    )
    parser.add_argument(
        "--sheet",
        action="append",
        help=(
            "Target sheet for append mode. Repeat to append to multiple sheets. "
            "Defaults to BoA when present, otherwise the first sheet."
        ),
    )
    parser.add_argument(
        "--write-mode",
        choices=["append", "skip_duplicates", "overwrite"],
        default="append",
        help="How append mode should write rows. Defaults to append.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace the output file if it already exists.",
    )
    parser.add_argument(
        "--no-text-fallback",
        action="store_true",
        help="Do not create text-only sheets when no tables are found on a page.",
    )
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="Use OCR fallback for scanned PDFs in append mode when dependencies exist.",
    )
    parser.add_argument(
        "--pdf-password",
        help="Password for encrypted PDF statements in append mode.",
    )
    parser.add_argument(
        "--excel-password",
        help="Password for encrypted Excel workbooks in append mode.",
    )
    parser.add_argument(
        "--no-highlight",
        action="store_true",
        help="Do not highlight newly written cells in append mode.",
    )
    parser.add_argument(
        "--use-ai-mapping",
        action="store_true",
        help=(
            "Use OpenAI field mapping when OPENAI_API_KEY and "
            "OPENAI_MAPPING_MODEL are configured; otherwise use heuristics."
        ),
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.append_to:
            result = append_pdf_transactions_to_excel(
                args.pdf,
                args.append_to,
                output_path=args.output,
                sheets=args.sheet,
                write_mode=args.write_mode,
                overwrite=args.overwrite,
                use_ocr=args.ocr,
                pdf_password=args.pdf_password,
                excel_password=args.excel_password,
                highlight_new_cells=not args.no_highlight,
                use_ai_mapping=args.use_ai_mapping,
            )
            _print_append_result(result)
            return 0

        if len(args.pdf) != 1:
            parser.error("table conversion accepts exactly one PDF unless --append-to is used")

        result = convert_pdf_to_excel(
            args.pdf[0],
            output_path=args.output,
            overwrite=args.overwrite,
            text_fallback=not args.no_text_fallback,
        )
    except (PdfExcelError, ExcelAnalysisError, ExcelWriteError, OSError, ValueError) as exc:
        print(f"pdf-excel: error: {exc}", file=sys.stderr)
        return 1

    print(
        "Wrote "
        f"{result.output_path} "
        f"({result.sheet_count} sheet(s), "
        f"{result.row_count} row(s), "
        f"{result.page_count} page(s))."
    )
    return 0


def _print_append_result(result) -> None:
    warning_count = sum(
        1 for issue in result.validation_issues if issue.severity == "warning"
    )
    sheet_names = ", ".join(result.summary.target_sheets)
    print(
        "Wrote "
        f"{result.output_path} "
        f"({result.transaction_count} transaction(s), "
        f"{result.summary.rows_added} row(s) added, "
        f"{result.summary.rows_skipped} row(s) skipped, "
        f"sheet(s): {sheet_names})."
    )
    if warning_count:
        print(
            f"pdf-excel: warning: {warning_count} validation warning(s) were found.",
            file=sys.stderr,
        )
    for warning in result.summary.warnings:
        print(f"pdf-excel: warning: {warning}", file=sys.stderr)
