"""Tests for the command line interface."""

from __future__ import annotations

import unittest
from contextlib import redirect_stdout
from io import StringIO

from pdf_excel.cli import main


class CliTests(unittest.TestCase):
    def test_help_command_exits_successfully(self) -> None:
        with redirect_stdout(StringIO()), self.assertRaises(SystemExit) as caught:
            main(["--help"])

        self.assertEqual(caught.exception.code, 0)


if __name__ == "__main__":
    unittest.main()
