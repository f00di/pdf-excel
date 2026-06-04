"""Validation helpers for PDF-to-Excel append plans."""

from __future__ import annotations

from collections import Counter
from typing import Any

from pdf_excel.models import (
    ExtractedTransaction,
    MappingSuggestion,
    SheetStructure,
    ValidationIssue,
)
from pdf_excel.statement_extractor import normalize_date


def validate_append_plan(
    transactions: list[ExtractedTransaction],
    sheet: SheetStructure,
    mapping: MappingSuggestion,
    existing_rows: list[dict[str, Any]] | None = None,
) -> list[ValidationIssue]:
    """Validate the transactions, target sheet, and suggested mapping."""

    issues: list[ValidationIssue] = []
    if not transactions:
        issues.append(
            ValidationIssue("error", "No extracted transactions are available to append.")
        )
        return issues

    issues.extend(_validate_mapping(sheet, mapping))
    issues.extend(_validate_transactions(transactions))
    issues.extend(_detect_duplicate_transactions(transactions, existing_rows or [], mapping))
    issues.extend(_validate_balances(transactions))
    return issues


def _validate_mapping(
    sheet: SheetStructure,
    mapping: MappingSuggestion,
) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    if mapping.start_row <= mapping.header_row:
        issues.append(
            ValidationIssue(
                "error",
                "Starting row must be below the detected header row.",
                row_index=mapping.start_row,
            )
        )

    for required_field in ["transaction_date", "description"]:
        if required_field not in mapping.field_mapping:
            issues.append(
                ValidationIssue("error", f"Required field is not mapped: {required_field}.")
            )

    for field, column in mapping.field_mapping.items():
        if column not in sheet.headers:
            issues.append(
                ValidationIssue(
                    "error",
                    f"Mapped column {column} for {field} does not exist in the "
                    "selected sheet.",
                )
            )
    return issues


def _validate_transactions(
    transactions: list[ExtractedTransaction],
) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    for index, transaction in enumerate(transactions, start=1):
        if not transaction.transaction_date:
            issues.append(
                ValidationIssue(
                    "error",
                    "Missing transaction date.",
                    index,
                    "transaction_date",
                )
            )
        elif not normalize_date(transaction.transaction_date):
            issues.append(
                ValidationIssue(
                    "warning",
                    "Transaction date could not be validated.",
                    index,
                    "transaction_date",
                )
            )

        if not transaction.description:
            issues.append(
                ValidationIssue("error", "Missing description.", index, "description")
            )

        if transaction.debit is not None and transaction.credit is not None:
            issues.append(
                ValidationIssue("warning", "Both debit and credit are populated.", index)
            )
        if transaction.debit is None and transaction.credit is None:
            issues.append(
                ValidationIssue("warning", "Neither debit nor credit is populated.", index)
            )

        for field in ["debit", "credit", "balance"]:
            value = getattr(transaction, field)
            if value is None:
                continue
            try:
                float(value)
            except (TypeError, ValueError):
                issues.append(
                    ValidationIssue("error", f"{field} is not numeric.", index, field)
                )

        for warning in transaction.warnings:
            issues.append(ValidationIssue("warning", warning, index))
    return issues


def _detect_duplicate_transactions(
    transactions: list[ExtractedTransaction],
    existing_rows: list[dict[str, Any]],
    mapping: MappingSuggestion,
) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    extracted_keys = [
        _transaction_key(
            transaction.transaction_date,
            transaction.description,
            transaction.debit,
            transaction.credit,
        )
        for transaction in transactions
    ]
    duplicates = [key for key, count in Counter(extracted_keys).items() if count > 1 and key]
    for key in duplicates:
        issues.append(
            ValidationIssue("warning", f"Duplicate-looking extracted transaction: {key}.")
        )

    if not existing_rows:
        return issues

    date_col = mapping.field_mapping.get("transaction_date")
    desc_col = mapping.field_mapping.get("description")
    debit_col = mapping.field_mapping.get("debit")
    credit_col = mapping.field_mapping.get("credit")
    existing_keys = set()
    for row in existing_rows:
        existing_keys.add(
            _transaction_key(
                row.get(date_col, ""),
                row.get(desc_col, ""),
                row.get(debit_col, None),
                row.get(credit_col, None),
            )
        )

    for index, transaction in enumerate(transactions, start=1):
        key = _transaction_key(
            transaction.transaction_date,
            transaction.description,
            transaction.debit,
            transaction.credit,
        )
        if key and key in existing_keys:
            issues.append(
                ValidationIssue(
                    "warning",
                    "Possible duplicate already exists in the sheet.",
                    index,
                )
            )
    return issues


def _validate_balances(transactions: list[ExtractedTransaction]) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    previous_balance: float | None = None
    for index, transaction in enumerate(transactions, start=1):
        if previous_balance is not None and transaction.balance is not None:
            debit = float(transaction.debit or 0)
            credit = float(transaction.credit or 0)
            expected = round(previous_balance - debit + credit, 2)
            actual = round(float(transaction.balance), 2)
            if abs(expected - actual) > 0.05:
                issues.append(
                    ValidationIssue(
                        "warning",
                        f"Balance continuity check differs. Expected {expected}, "
                        f"saw {actual}.",
                        index,
                        "balance",
                    )
                )
        if transaction.balance is not None:
            previous_balance = float(transaction.balance)
    return issues


def _transaction_key(date: Any, description: Any, debit: Any, credit: Any) -> str:
    date_text = normalize_date(date) or str(date or "").strip()
    description_text = " ".join(str(description or "").lower().split())
    debit_text = _amount_key(debit)
    credit_text = _amount_key(credit)
    if not any([date_text, description_text, debit_text, credit_text]):
        return ""
    return "|".join([date_text, description_text, debit_text, credit_text])


def _amount_key(value: Any) -> str:
    if value in (None, ""):
        return ""
    try:
        return f"{float(str(value).replace(',', '').replace('$', '')):.2f}"
    except ValueError:
        return str(value).strip()

