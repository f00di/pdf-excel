"""Command line interface for PDF Excel."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from pdf_excel.converter import PdfExcelError, convert_pdf_to_excel


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf-excel",
        description="Extract tables from a PDF and write them to an Excel workbook.",
    )
    parser.add_argument("pdf", type=Path, help="Path to the source PDF file.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Path for the generated .xlsx file. Defaults to the PDF name.",
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
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        result = convert_pdf_to_excel(
            args.pdf,
            output_path=args.output,
            overwrite=args.overwrite,
            text_fallback=not args.no_text_fallback,
        )
    except (PdfExcelError, OSError, ValueError) as exc:
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
