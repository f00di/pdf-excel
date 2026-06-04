"""Suggest mappings from extracted statement fields to Excel columns."""

from __future__ import annotations

import json
import os
import re
from typing import Any

from pdf_excel.models import ExtractedTransaction, MappingSuggestion, SheetStructure


FIELD_ALIASES = {
    "transaction_date": [
        "date",
        "transaction date",
        "trans date",
        "posting date",
        "value date",
    ],
    "description": [
        "description",
        "details",
        "narration",
        "particulars",
        "memo",
        "payee",
        "transaction",
    ],
    "debit": [
        "debit",
        "withdrawal",
        "withdrawals",
        "paid out",
        "charge",
        "spent",
        "outflow",
    ],
    "credit": [
        "credit",
        "deposit",
        "deposits",
        "paid in",
        "received",
        "inflow",
    ],
    "balance": ["balance", "running balance", "available balance"],
    "reference_number": [
        "reference",
        "reference number",
        "ref",
        "check",
        "cheque",
        "trace",
        "id",
    ],
    "account_number": ["account", "account number", "acct", "acct number"],
    "statement_period": ["statement period", "period"],
    "opening_balance": ["opening balance", "beginning balance"],
    "closing_balance": ["closing balance", "ending balance"],
    "category": ["category", "type", "class"],
    "notes": ["notes", "note", "comments", "memo", "remarks", "details"],
}


def suggest_mapping(
    sheet: SheetStructure,
    transactions: list[ExtractedTransaction],
    *,
    use_ai: bool = False,
) -> MappingSuggestion:
    """Suggest a field-to-column mapping for a selected sheet."""

    heuristic = _heuristic_mapping(sheet, transactions)
    if not use_ai:
        return heuristic

    ai_suggestion = _try_openai_mapping(sheet, transactions)
    if ai_suggestion:
        return ai_suggestion

    heuristic.warnings.append(
        "AI mapping was requested, but no AI provider was configured. Used "
        "deterministic mapping instead."
    )
    return heuristic


def _heuristic_mapping(
    sheet: SheetStructure,
    transactions: list[ExtractedTransaction],
) -> MappingSuggestion:
    field_mapping: dict[str, str] = {}
    field_scores: dict[str, float] = {}
    used_columns: set[str] = set()
    headers = sheet.headers or {}

    for field, aliases in FIELD_ALIASES.items():
        best_column = ""
        best_score = 0.0
        for column, header in headers.items():
            if column in used_columns:
                continue
            if _should_preserve_formula_column(sheet, field, column):
                continue
            score = _score_header(field, header, aliases)
            if score > best_score:
                best_column = column
                best_score = score
        minimum_score = (
            0.75
            if field
            in {"statement_period", "account_number", "opening_balance", "closing_balance"}
            else 0.45
        )
        if best_column and best_score >= minimum_score:
            field_mapping[field] = best_column
            field_scores[field] = min(best_score, 1.0)
            used_columns.add(best_column)

    warnings = []
    for required_field in ["transaction_date", "description"]:
        if required_field not in field_mapping:
            warnings.append(f"No confident mapping found for {required_field}.")
    if not any(field in field_mapping for field in ["debit", "credit"]):
        warnings.append("No confident debit or credit mapping found.")
    for column, info in sheet.column_formats.items():
        if info.get("formula_count", 0) and _normalize(info.get("header")) == "balance":
            warnings.append(
                f"Detected formulas in balance column {column}; balance will be "
                "extended from the sheet formula instead of imported from the PDF."
            )

    populated_fields = _populated_transaction_fields(transactions)
    relevant_scores = [
        score
        for field, score in field_scores.items()
        if field in populated_fields or field in {"transaction_date", "description"}
    ]
    confidence = (
        round(sum(relevant_scores) / len(relevant_scores), 2)
        if relevant_scores
        else 0.0
    )

    return MappingSuggestion(
        target_sheet=sheet.sheet_name,
        header_row=sheet.header_row,
        start_row=sheet.next_row,
        field_mapping=field_mapping,
        confidence_score=confidence,
        warnings=warnings,
    )


def _try_openai_mapping(
    sheet: SheetStructure,
    transactions: list[ExtractedTransaction],
) -> MappingSuggestion | None:
    model = os.getenv("OPENAI_MAPPING_MODEL")
    if not os.getenv("OPENAI_API_KEY") or not model:
        return None

    try:
        from openai import OpenAI
    except ImportError:
        return None

    sample_rows = [transaction.to_dict() for transaction in transactions[:5]]
    prompt_payload = {
        "sheet": {
            "name": sheet.sheet_name,
            "header_row": sheet.header_row,
            "start_row": sheet.next_row,
            "headers": sheet.headers,
            "column_formats": sheet.column_formats,
        },
        "extracted_pdf_rows": sample_rows,
        "required_output_shape": {
            "target_sheet": "Sheet1",
            "header_row": 1,
            "start_row": 25,
            "field_mapping": {
                "transaction_date": "A",
                "description": "B",
                "debit": "C",
                "credit": "D",
                "balance": "E",
            },
            "confidence_score": 0.92,
            "warnings": [],
        },
    }

    try:
        client = OpenAI()
        response = client.responses.create(
            model=model,
            input=(
                "Recommend a PDF bank-statement field to Excel-column mapping. "
                "Return strict JSON only. Never suggest writing without user "
                "confirmation.\n\n"
                + json.dumps(prompt_payload, default=str)
            ),
        )
        payload = json.loads(response.output_text)
    except Exception:
        return None

    field_mapping = {
        field: column
        for field, column in payload.get("field_mapping", {}).items()
        if isinstance(field, str)
        and isinstance(column, str)
        and column in sheet.headers
        and not _should_preserve_formula_column(sheet, field, column)
    }
    if not field_mapping:
        return None

    return MappingSuggestion(
        target_sheet=sheet.sheet_name,
        header_row=int(payload.get("header_row", sheet.header_row)),
        start_row=int(payload.get("start_row", sheet.next_row)),
        field_mapping=field_mapping,
        confidence_score=float(payload.get("confidence_score", 0.0)),
        warnings=list(payload.get("warnings", [])),
    )


def _score_header(field: str, header: str, aliases: list[str]) -> float:
    normalized_header = _normalize(header)
    if not normalized_header:
        return 0.0
    if field == "statement_period" and normalized_header == "bank statement":
        return 0.2

    normalized_aliases = [_normalize(alias) for alias in aliases]
    if normalized_header in normalized_aliases:
        return 1.0
    if any(alias in normalized_header for alias in normalized_aliases):
        return 0.88

    header_tokens = set(normalized_header.split())
    best_overlap = 0.0
    for alias in normalized_aliases:
        alias_tokens = set(alias.split())
        overlap = len(header_tokens & alias_tokens) / max(len(alias_tokens), 1)
        best_overlap = max(best_overlap, overlap)

    if field == "debit" and any(
        token in normalized_header for token in ["withdraw", "payment", "out"]
    ):
        best_overlap = max(best_overlap, 0.78)
    if field == "credit" and any(
        token in normalized_header for token in ["deposit", "received", "in"]
    ):
        best_overlap = max(best_overlap, 0.78)
    if field == "transaction_date" and "date" in normalized_header:
        best_overlap = max(best_overlap, 0.82)
    if field == "description" and any(
        token in normalized_header for token in ["detail", "memo", "name"]
    ):
        best_overlap = max(best_overlap, 0.74)

    return best_overlap


def _should_preserve_formula_column(
    sheet: SheetStructure,
    field: str,
    column: str,
) -> bool:
    if field not in {"balance", "opening_balance", "closing_balance"}:
        return False
    column_info = sheet.column_formats.get(column, {})
    return int(column_info.get("formula_count", 0) or 0) > 0


def _populated_transaction_fields(transactions: list[ExtractedTransaction]) -> set[str]:
    populated = set()
    for transaction in transactions[:50]:
        for field, value in transaction.to_dict().items():
            if field.startswith("source_") or field in {
                "raw_text",
                "warnings",
                "confidence",
            }:
                continue
            if value not in (None, "", []):
                populated.add(field)
    return populated


def _normalize(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()

