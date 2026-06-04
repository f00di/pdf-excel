"""Tests for workbook analysis, mapping, validation, and append workflow."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from pdf_excel.append import append_pdf_transactions_to_excel
from pdf_excel.mapping import suggest_mapping
from pdf_excel.models import ExtractedTransaction, SheetStructure
from pdf_excel.validation import validate_append_plan
from pdf_excel.workbook_analyzer import find_next_append_row


def _escape_pdf_text(value: str) -> str:
    return value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _simple_text_pdf(lines: list[str]) -> bytes:
    text_lines = ["BT", "/F1 12 Tf", "72 720 Td"]
    for index, line in enumerate(lines):
        if index:
            text_lines.append("0 -18 Td")
        text_lines.append(f"({_escape_pdf_text(line)}) Tj")
    text_lines.append("ET")

    stream = "\n".join(text_lines).encode("ascii")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        (
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            b"/Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>"
        ),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
    ]

    pdf = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for number, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{number} 0 obj\n".encode("ascii"))
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")

    xref_offset = len(pdf)
    pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    pdf.extend(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref_offset}\n%%EOF\n"
        ).encode("ascii")
    )
    return bytes(pdf)


class AppendFeatureTests(unittest.TestCase):
    def test_mapping_detects_common_bank_columns(self) -> None:
        sheet = SheetStructure(
            sheet_name="Transactions",
            max_row=10,
            max_column=5,
            header_row=1,
            headers={
                "A": "Date",
                "B": "Description",
                "C": "Withdrawals",
                "D": "Deposits",
                "E": "Running Balance",
            },
            preview_rows=[],
            empty_areas=[],
            tables=[],
            merged_cells=[],
            data_validations=[],
            formulas=[],
            column_formats={},
            column_widths={},
            next_row=11,
        )
        transactions = [
            ExtractedTransaction(
                transaction_date="2026-01-01",
                description="Coffee",
                debit=4.25,
            )
        ]

        suggestion = suggest_mapping(sheet, transactions)

        self.assertEqual(suggestion.field_mapping["transaction_date"], "A")
        self.assertEqual(suggestion.field_mapping["description"], "B")
        self.assertEqual(suggestion.field_mapping["debit"], "C")
        self.assertEqual(suggestion.field_mapping["credit"], "D")
        self.assertEqual(suggestion.field_mapping["balance"], "E")

    def test_validation_requires_date_and_description_mapping(self) -> None:
        sheet = SheetStructure(
            sheet_name="Transactions",
            max_row=10,
            max_column=2,
            header_row=1,
            headers={"A": "Date", "B": "Description"},
            preview_rows=[],
            empty_areas=[],
            tables=[],
            merged_cells=[],
            data_validations=[],
            formulas=[],
            column_formats={},
            column_widths={},
            next_row=11,
        )
        suggestion = suggest_mapping(sheet, [])
        suggestion.field_mapping.pop("description", None)

        issues = validate_append_plan(
            transactions=[
                ExtractedTransaction(
                    transaction_date="2026-01-01",
                    description="Coffee",
                    debit=4.25,
                )
            ],
            sheet=sheet,
            mapping=suggestion,
        )

        self.assertTrue(
            any(
                issue.severity == "error" and "description" in issue.message
                for issue in issues
            )
        )

    def test_next_append_row_uses_date_column_before_formula_templates(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Date", "Description", "Credit", "Debit", "Balance"])
        sheet.append(["2026-03-01", "Existing transaction", None, 10, 90])
        sheet.append(["2026-03-02", None, None, None, "=E2+C3-D3"])
        sheet.append([None, None, None, None, "=E3+C4-D4"])
        sheet.append(
            [
                None,
                "Template text that should not move the insertion point",
                None,
                None,
                "=E4+C5-D5",
            ]
        )

        self.assertEqual(find_next_append_row(sheet, header_row=1), 4)

    def test_appends_boa_pdf_transactions_to_existing_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            pdf_path = tmp_path / "boa.pdf"
            workbook_path = tmp_path / "budget.xlsx"
            output_path = tmp_path / "budget_updated.xlsx"

            pdf_path.write_bytes(
                _simple_text_pdf(
                    [
                        "Your Adv SafeBalance Banking",
                        (
                            "for April 23, 2026 to May 20, 2026 Account number: "
                            "3252 0963 1573"
                        ),
                        "Account summary",
                        "Beginning balance on April 23, 2026 $19.38",
                        "Ending balance on May 20, 2026 $1,294.14",
                        "Withdrawals and other subtractions",
                        "Deposits and other additions",
                        "Date Description Amount",
                        (
                            "04/30/26 Zelle payment from BATOOL AFSHAR "
                            "Conf# otuapdnyg 57.50"
                        ),
                        "Total deposits and other additions $57.50",
                        "ATM and debit card subtractions",
                        "Date Description Amount",
                        "05/04/26 PURCHASE 0502 UBER EATS San FranciscoCA -16.53",
                        "Total ATM and debit card subtractions -$16.53",
                        "Bank of America",
                    ]
                )
            )

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "BoA"
            worksheet.append(["Date", "Description", "Withdrawals", "Deposits", "Notes"])
            worksheet.append(["2026-04-29", "Existing", 1.0, None, ""])
            workbook.save(workbook_path)

            result = append_pdf_transactions_to_excel(
                [pdf_path],
                workbook_path,
                output_path=output_path,
            )

            self.assertEqual(result.transaction_count, 2)
            self.assertEqual(result.summary.rows_added, 2)
            self.assertTrue(output_path.exists())

            updated = load_workbook(output_path)
            sheet = updated["BoA"]
            self.assertEqual(sheet["D3"].value, 57.5)
            self.assertEqual(sheet["C4"].value, 16.53)
            self.assertEqual(sheet["B3"].value[:8], "04/30/26")
            self.assertIn("_PDF_Append_Log", updated.sheetnames)


if __name__ == "__main__":
    unittest.main()

