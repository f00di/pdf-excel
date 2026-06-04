"""Tests for bank-statement extraction helpers."""

from __future__ import annotations

import unittest

from pdf_excel.statement_extractor import extract_statement_metadata, extract_transactions


BOA_TEXT = """
--- Page 1 ---
Your Adv SafeBalance Banking
for April 23, 2026 to May 20, 2026 Account number: 3252 0963 1573
Account summary
Beginning balance on April 23, 2026 $19.38
Ending balance on May 20, 2026 $1,294.14
--- Page 3 ---
Deposits and other additions
Date Description Amount
04/30/26 Zelle payment from BATOOL AFSHAR Conf# otuapdnyg 57.50
Total deposits and other additions $57.50
Withdrawals and other subtractions
ATM and debit card subtractions
Date Description Amount
05/04/26 PURCHASE 0502 UBER * EATS San FranciscoCA -16.53
Total ATM and debit card subtractions -$16.53
Other subtractions
Date Description Amount
05/14/26 Online Banking transfer to SAV 9909 Confirmation# 7880621814 -76.85
Total other subtractions -$76.85
Service fees
Date Transaction description Amount
05/18/26 PURCHASE 0517 Amazon Prime Subscript Dubai -0.13
Total service fees -$0.13
Bank of America
"""


class StatementExtractorTests(unittest.TestCase):
    def test_boa_metadata_is_precise(self) -> None:
        metadata = extract_statement_metadata(BOA_TEXT)

        self.assertEqual(metadata["account_number"], "3252 0963 1573")
        self.assertEqual(
            metadata["statement_period"],
            "April 23, 2026 to May 20, 2026",
        )
        self.assertEqual(metadata["statement_start"], "2026-04-23")
        self.assertEqual(metadata["statement_end"], "2026-05-20")
        self.assertEqual(metadata["opening_balance"], 19.38)
        self.assertEqual(metadata["closing_balance"], 1294.14)

    def test_boa_parser_ignores_summary_and_verifies_section_totals(self) -> None:
        metadata = extract_statement_metadata(BOA_TEXT)
        transactions = extract_transactions("boa.pdf", BOA_TEXT, [], metadata)

        self.assertEqual(len(transactions), 4)
        self.assertEqual(transactions[0].credit, 57.50)
        self.assertIsNone(transactions[0].debit)
        self.assertEqual(transactions[1].debit, 16.53)
        self.assertEqual(transactions[2].reference_number, "7880621814")
        self.assertEqual(transactions[3].debit, 0.13)
        self.assertTrue(all(transaction.confidence == 0.96 for transaction in transactions))


if __name__ == "__main__":
    unittest.main()

