"""Tests for conversion helpers that do not require PDF dependencies."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pdf_excel.converter import (
    PdfExcelError,
    _clean_rows,
    _extract_text_rows,
    _resolve_output_path,
    _sanitize_sheet_title,
    _unique_sheet_title,
    _validate_pdf_path,
)


class ConverterHelperTests(unittest.TestCase):
    def test_clean_rows_strips_cells_and_skips_empty_rows(self) -> None:
        rows = [
            [" Name ", None, "Amount\nDue"],
            ["", " ", None],
            [123, " paid ", ""],
        ]

        self.assertEqual(
            _clean_rows(rows),
            [["Name", "", "Amount Due"], ["123", "paid", ""]],
        )

    def test_extract_text_rows_uses_non_empty_lines(self) -> None:
        self.assertEqual(
            _extract_text_rows(" Header \n\nValue 1\n Value 2 "),
            [["Header"], ["Value 1"], ["Value 2"]],
        )

    def test_sanitize_sheet_title_replaces_invalid_characters(self) -> None:
        title = _sanitize_sheet_title("[]:*?/\\")

        self.assertEqual(title, "_______")

    def test_sanitize_sheet_title_limits_excel_length(self) -> None:
        title = _sanitize_sheet_title("A" * 40)

        self.assertEqual(len(title), 31)

    def test_unique_sheet_title_adds_suffix(self) -> None:
        existing = {"Page 1 Table 1"}

        self.assertEqual(
            _unique_sheet_title(existing, "Page 1 Table 1"),
            "Page 1 Table 1 2",
        )

    def test_default_output_path_uses_xlsx_suffix(self) -> None:
        self.assertEqual(
            _resolve_output_path(Path("/tmp/example.pdf"), None),
            Path("/tmp/example.xlsx"),
        )

    def test_rejects_non_xlsx_output_path(self) -> None:
        with self.assertRaises(PdfExcelError):
            _resolve_output_path(Path("/tmp/example.pdf"), "/tmp/example.csv")

    def test_validate_pdf_path_rejects_missing_file(self) -> None:
        with self.assertRaises(PdfExcelError):
            _validate_pdf_path("/tmp/does-not-exist.pdf")

    def test_validate_pdf_path_accepts_pdf_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "sample.pdf"
            path.write_bytes(b"%PDF-1.4\n")

            self.assertEqual(_validate_pdf_path(path), path.resolve())


if __name__ == "__main__":
    unittest.main()
