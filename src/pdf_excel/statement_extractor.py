"""Extract normalized bank-statement transactions from PDF files."""

from __future__ import annotations

import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any

from pdf_excel.models import ExtractedTransaction, PDFExtractionResult


DATE_PATTERN = re.compile(
    r"(?P<date>\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|"
    r"\d{4}-\d{1,2}-\d{1,2}|"
    r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)"
    r"[a-z]*\.?\s+\d{1,2},?\s+\d{2,4})\b)",
    re.IGNORECASE,
)
AMOUNT_PATTERN = re.compile(
    r"(?<![\w])(?P<amount>-?\(?\$?\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})\)?|"
    r"-?\(?\$?\s*\d+\.\d{2}\)?)(?![\w])"
)
BOA_TRANSACTION_LINE_PATTERN = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{2,4})\s+"
    r"(?P<description>.+?)\s+"
    r"(?P<amount>-?\$?\d[\d,]*\.\d{2})$"
)


TABLE_FIELD_ALIASES = {
    "transaction_date": [
        "date",
        "trans date",
        "transaction date",
        "posting date",
        "value date",
    ],
    "description": [
        "description",
        "details",
        "narration",
        "particulars",
        "memo",
        "transaction",
    ],
    "debit": ["debit", "withdrawal", "withdrawals", "paid out", "payment", "charge"],
    "credit": ["credit", "deposit", "deposits", "paid in", "receipt"],
    "signed_amount": ["amount", "transaction amount", "value"],
    "balance": ["balance", "running balance", "available balance"],
    "reference_number": ["ref", "reference", "reference number", "check", "cheque", "trace"],
}

BOA_SECTION_TYPES = {
    "deposits and other additions": "credit",
    "atm and debit card subtractions": "debit",
    "other subtractions": "debit",
    "service fees": "debit",
}


def extract_pdf_path(
    pdf_path: str | Path,
    *,
    use_ocr: bool = False,
    password: str | None = None,
) -> PDFExtractionResult:
    """Extract statement transactions from a PDF path."""

    path = Path(pdf_path).expanduser().resolve()
    return extract_pdf_file(
        file_name=path.name,
        file_bytes=path.read_bytes(),
        use_ocr=use_ocr,
        password=password,
    )


def extract_pdf_files(
    uploaded_files: list[tuple[str, bytes]],
    *,
    use_ocr: bool = False,
    password: str | None = None,
) -> list[PDFExtractionResult]:
    """Extract statement transactions from in-memory PDF files."""

    return [
        extract_pdf_file(
            file_name=file_name,
            file_bytes=file_bytes,
            use_ocr=use_ocr,
            password=password,
        )
        for file_name, file_bytes in uploaded_files
    ]


def extract_pdf_file(
    file_name: str,
    file_bytes: bytes,
    *,
    use_ocr: bool = False,
    password: str | None = None,
) -> PDFExtractionResult:
    """Extract raw text, tables, metadata, and transaction rows from a PDF."""

    warnings: list[str] = []
    raw_text, tables, extraction_warnings = _extract_with_pdfplumber(
        file_bytes,
        password=password,
    )
    warnings.extend(extraction_warnings)
    ocr_used = False

    if use_ocr and len(raw_text.strip()) < 80:
        ocr_text, ocr_warnings = _extract_with_ocr(file_bytes, password=password)
        warnings.extend(ocr_warnings)
        if ocr_text.strip():
            raw_text = ocr_text
            ocr_used = True

    metadata = extract_statement_metadata(raw_text)
    transactions = extract_transactions(
        file_name=file_name,
        raw_text=raw_text,
        tables=tables,
        metadata=metadata,
    )

    if not transactions:
        warnings.append(
            "No transaction rows were detected automatically. Review raw text "
            "or export the PDF text and map rows manually."
        )

    return PDFExtractionResult(
        file_name=file_name,
        raw_text=raw_text,
        tables=tables,
        transactions=transactions,
        metadata=metadata,
        warnings=warnings,
        ocr_used=ocr_used,
    )


def _extract_with_pdfplumber(
    file_bytes: bytes,
    *,
    password: str | None = None,
) -> tuple[str, list[dict[str, Any]], list[str]]:
    warnings: list[str] = []
    try:
        import pdfplumber
    except ImportError:
        return "", [], ["pdfplumber is not installed; PDF extraction is unavailable."]

    page_texts: list[str] = []
    tables: list[dict[str, Any]] = []
    try:
        with pdfplumber.open(BytesIO(file_bytes), password=password or "") as pdf:
            for index, page in enumerate(pdf.pages, start=1):
                text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
                page_texts.append(f"--- Page {index} ---\n{text}".strip())
                for table in page.extract_tables() or []:
                    cleaned_rows = [
                        [_clean_cell(cell) for cell in row]
                        for row in table
                        if row and any(_clean_cell(cell) for cell in row)
                    ]
                    if cleaned_rows:
                        tables.append({"page": index, "rows": cleaned_rows})
    except Exception as exc:
        message = str(exc)
        if "password" in message.lower() or "encrypted" in message.lower():
            warnings.append("This PDF appears to be password-protected or encrypted.")
        else:
            warnings.append(f"PDF parsing failed: {message}")

    return "\n\n".join(page_texts).strip(), tables, warnings


def _extract_with_ocr(
    file_bytes: bytes,
    *,
    password: str | None = None,
) -> tuple[str, list[str]]:
    warnings: list[str] = []
    try:
        import fitz
        import pytesseract
        from PIL import Image
    except ImportError:
        return (
            "",
            [
                "OCR fallback requires PyMuPDF, Pillow, pytesseract, and the "
                "Tesseract system binary."
            ],
        )

    page_texts: list[str] = []
    try:
        document = fitz.open(stream=file_bytes, filetype="pdf")
        if document.needs_pass and not document.authenticate(password or ""):
            return "", ["This PDF is password-protected. Enter the password and try again."]
        for index, page in enumerate(document, start=1):
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image = Image.open(BytesIO(pixmap.tobytes("png")))
            text = pytesseract.image_to_string(image)
            page_texts.append(f"--- Page {index} OCR ---\n{text}".strip())
    except Exception as exc:
        warnings.append(
            f"OCR fallback failed: {exc}. Tesseract must be installed separately "
            "from Python packages."
        )

    return "\n\n".join(page_texts).strip(), warnings


def extract_statement_metadata(raw_text: str) -> dict[str, Any]:
    """Extract account and statement summary metadata from raw statement text."""

    metadata: dict[str, Any] = {}
    condensed = " ".join(raw_text.split())

    account_match = _first_match(
        [
            r"Account\s+number:\s*([0-9 ]{6,})",
            r"Account\s*#\s*([0-9 ]{6,})",
            r"(?:account|acct)\s*(?:number|no\.?|#)?\s*[:\-]?"
            r"\s*([*Xx0-9\- ]{4,})",
        ],
        condensed,
    )
    if account_match:
        metadata["account_number"] = " ".join(account_match.group(1).split())

    boa_period_match = re.search(
        r"\bfor\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})\s+to\s+"
        r"([A-Za-z]+\s+\d{1,2},\s+\d{4})\b",
        condensed,
        re.IGNORECASE,
    )
    if boa_period_match:
        start = boa_period_match.group(1)
        end = boa_period_match.group(2)
        metadata["statement_period"] = f"{start} to {end}"
        metadata["statement_start"] = normalize_date(start)
        metadata["statement_end"] = normalize_date(end)

    period_match = re.search(
        r"(?:statement\s+period|period)\s*[:\-]?\s*("
        r".{0,40}?\d{1,2}[/-]\d{1,2}[/-]\d{2,4}"
        r".{0,20}?\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
        condensed,
        re.IGNORECASE,
    )
    if period_match and "statement_period" not in metadata:
        metadata["statement_period"] = period_match.group(1).strip()

    opening_match = re.search(
        r"(?:opening|beginning)\s+balance"
        r"(?:\s+on\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})?"
        r"\s*[:\-]?\s*(\$?\(?-?[\d,]+\.\d{2}\)?)",
        condensed,
        re.IGNORECASE,
    )
    if opening_match:
        metadata["opening_balance"] = parse_amount(opening_match.group(1))

    closing_match = re.search(
        r"(?:closing|ending)\s+balance"
        r"(?:\s+on\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})?"
        r"\s*[:\-]?\s*(\$?\(?-?[\d,]+\.\d{2}\)?)",
        condensed,
        re.IGNORECASE,
    )
    if closing_match:
        metadata["closing_balance"] = parse_amount(closing_match.group(1))

    return metadata


def extract_transactions(
    file_name: str,
    raw_text: str,
    tables: list[dict[str, Any]],
    metadata: dict[str, Any],
) -> list[ExtractedTransaction]:
    """Extract normalized transactions from specialized, table, or text sources."""

    boa_transactions = _transactions_from_boa_statement(file_name, raw_text, metadata)
    if boa_transactions:
        return boa_transactions

    table_transactions = _transactions_from_tables(file_name, tables, metadata)
    if table_transactions:
        return table_transactions
    return _transactions_from_lines(file_name, raw_text, metadata)


def _transactions_from_boa_statement(
    file_name: str,
    raw_text: str,
    metadata: dict[str, Any],
) -> list[ExtractedTransaction]:
    if not _looks_like_boa_statement(raw_text):
        return []

    rows: list[dict[str, Any]] = []
    totals: dict[str, float] = {}
    current_page: int | None = None
    current_section = ""
    active_table = False
    pending: dict[str, Any] | None = None

    def finish_pending() -> None:
        nonlocal pending
        if pending:
            rows.append(pending)
            pending = None

    for raw_line in raw_text.splitlines():
        line = " ".join(raw_line.split()).strip()
        if not line:
            continue

        page_match = re.match(r"--- Page (\d+)", line)
        if page_match:
            current_page = int(page_match.group(1))
            continue

        total_match = re.match(
            r"^Total\s+(.+?)\s+(-?\$?[\d,]+\.\d{2})$",
            line,
            re.IGNORECASE,
        )
        if total_match:
            finish_pending()
            total_section = _normalize_boa_section(total_match.group(1))
            if total_section in BOA_SECTION_TYPES:
                total_amount = parse_amount(total_match.group(2))
                if total_amount is not None:
                    totals[total_section] = abs(total_amount)
            active_table = False
            continue

        section = _normalize_boa_section(line)
        if section in BOA_SECTION_TYPES:
            finish_pending()
            current_section = section
            active_table = False
            continue

        if current_section and line.lower() in {
            "date description amount",
            "date transaction description amount",
        }:
            active_table = True
            continue

        if not active_table:
            continue

        if _is_boa_noise_line(line):
            finish_pending()
            active_table = "continued on the next page" in line.lower()
            continue

        transaction_match = BOA_TRANSACTION_LINE_PATTERN.match(line)
        if transaction_match:
            finish_pending()
            amount_text = transaction_match.group("amount")
            pending = {
                "date_text": transaction_match.group("date"),
                "description_parts": [transaction_match.group("description")],
                "amount_text": amount_text,
                "amount": parse_amount(amount_text),
                "section": current_section,
                "page": current_page,
            }
            continue

        if pending:
            pending["description_parts"].append(line)

    finish_pending()

    if not rows:
        return []

    section_sums: dict[str, float] = {}
    for row in rows:
        amount = row.get("amount")
        if amount is None:
            continue
        section = row["section"]
        section_sums[section] = section_sums.get(section, 0.0) + abs(amount)

    section_totals_match = {
        section: abs(round(section_sums.get(section, 0.0) - total, 2)) <= 0.02
        for section, total in totals.items()
    }

    transactions = []
    for row in rows:
        transaction = _boa_row_to_transaction(
            row=row,
            file_name=file_name,
            metadata=metadata,
            total_verified=section_totals_match.get(row["section"], False),
        )
        if transaction:
            transactions.append(transaction)

    return _deduplicate_transactions(transactions)


def _boa_row_to_transaction(
    row: dict[str, Any],
    file_name: str,
    metadata: dict[str, Any],
    total_verified: bool,
) -> ExtractedTransaction | None:
    amount = row.get("amount")
    if amount is None:
        return None

    normalized_date = normalize_date(row.get("date_text", ""))
    description_body = " ".join(row.get("description_parts", [])).strip()
    amount_text = str(row.get("amount_text", "")).strip()
    signed_description = (
        f"{row.get('date_text', '')} {description_body} {amount_text}".strip()
    )
    section_type = BOA_SECTION_TYPES.get(row.get("section", ""), "")

    debit: float | None = None
    credit: float | None = None
    warnings = []
    if amount < 0 or section_type == "debit":
        debit = abs(amount)
    elif section_type == "credit":
        credit = abs(amount)
    else:
        warnings.append("Could not determine Bank of America section for amount direction.")

    if not normalized_date:
        warnings.append("Could not normalize transaction date.")
    if not description_body:
        warnings.append("Missing transaction description.")

    if total_verified and not warnings:
        confidence = 0.96
    elif not warnings:
        confidence = 0.91
    else:
        confidence = 0.74

    reference = _extract_reference_number(description_body)
    notes = f"BoA section: {row.get('section', '')}"
    if total_verified:
        notes += "; section total verified"

    return ExtractedTransaction(
        transaction_date=normalized_date,
        description=signed_description,
        debit=debit,
        credit=credit,
        reference_number=reference,
        account_number=str(metadata.get("account_number", "")),
        statement_period=str(metadata.get("statement_period", "")),
        opening_balance=metadata.get("opening_balance"),
        closing_balance=metadata.get("closing_balance"),
        source_pdf=file_name,
        source_page=row.get("page"),
        raw_text=signed_description,
        confidence=confidence,
        notes=notes,
        warnings=warnings,
    )


def _looks_like_boa_statement(raw_text: str) -> bool:
    normalized = raw_text.lower()
    return (
        "bank of america" in normalized
        and "account summary" in normalized
        and "deposits and other additions" in normalized
        and "withdrawals and other subtractions" in normalized
    )


def _normalize_boa_section(line: str) -> str:
    normalized = re.sub(r"\s+-\s+continued$", "", line.strip(), flags=re.IGNORECASE)
    return normalized.lower()


def _is_boa_noise_line(line: str) -> bool:
    lower = line.lower()
    return (
        lower.startswith("page ")
        or lower.startswith("available in ")
        or lower == "continued on the next page"
        or lower.startswith("make bank transfers")
        or lower.startswith("use our app")
        or lower.startswith("scan the code")
        or lower.startswith("when you use")
        or lower.startswith("mobile banking requires")
        or lower.startswith("message and data rates")
        or lower.startswith("fees or other costs")
        or lower.startswith("note your ending balance")
        or lower.startswith("braille and large print")
        or lower.startswith("bankofamerica.com")
    )


def _extract_reference_number(description: str) -> str:
    match = re.search(r"\bConf#\s*([A-Z0-9]+)", description, re.IGNORECASE)
    if match:
        return match.group(1)
    match = re.search(
        r"\b(?:ID|Confirmation#)\s*:?\s*([A-Z0-9]+)",
        description,
        re.IGNORECASE,
    )
    if match:
        return match.group(1)
    long_number = re.search(r"\b(\d{12,})\b", description)
    return long_number.group(1) if long_number else ""


def _transactions_from_tables(
    file_name: str,
    tables: list[dict[str, Any]],
    metadata: dict[str, Any],
) -> list[ExtractedTransaction]:
    transactions: list[ExtractedTransaction] = []
    for table in tables:
        rows = table.get("rows", [])
        header_index, mapping = _detect_table_header(rows)
        if header_index is None:
            continue

        for row in rows[header_index + 1 :]:
            transaction = _transaction_from_table_row(
                row=row,
                mapping=mapping,
                file_name=file_name,
                page=table.get("page"),
                metadata=metadata,
            )
            if transaction:
                transactions.append(transaction)
    return _deduplicate_transactions(transactions)


def _detect_table_header(rows: list[list[str]]) -> tuple[int | None, dict[str, int]]:
    best_index: int | None = None
    best_mapping: dict[str, int] = {}
    best_score = 0

    for index, row in enumerate(rows[:8]):
        mapping: dict[str, int] = {}
        for col_index, value in enumerate(row):
            normalized = _normalize_label(value)
            for field, aliases in TABLE_FIELD_ALIASES.items():
                if field not in mapping and any(alias in normalized for alias in aliases):
                    mapping[field] = col_index

        score = len(mapping)
        if "transaction_date" in mapping:
            score += 2
        if "description" in mapping:
            score += 2
        if score > best_score:
            best_score = score
            best_index = index
            best_mapping = mapping

    if best_score < 3:
        return None, {}
    return best_index, best_mapping


def _transaction_from_table_row(
    row: list[str],
    mapping: dict[str, int],
    file_name: str,
    page: int | None,
    metadata: dict[str, Any],
) -> ExtractedTransaction | None:
    def value(field: str) -> str:
        index = mapping.get(field)
        if index is None or index >= len(row):
            return ""
        return _clean_cell(row[index])

    date_value = normalize_date(value("transaction_date"))
    description = value("description")
    debit = parse_amount(value("debit"))
    credit = parse_amount(value("credit"))
    signed_amount = parse_amount(value("signed_amount"))
    balance = parse_amount(value("balance"))
    reference = value("reference_number")

    if debit is None and credit is None and signed_amount is not None:
        if signed_amount < 0:
            debit = abs(signed_amount)
        else:
            credit = abs(signed_amount)

    if not any(
        [date_value, description, debit is not None, credit is not None, balance is not None]
    ):
        return None

    warnings = []
    if value("transaction_date") and not date_value:
        warnings.append("Could not normalize transaction date.")
    if debit is not None and debit < 0:
        debit = abs(debit)
    if credit is not None and credit < 0:
        debit = abs(credit)
        credit = None
    if debit is not None and credit is not None:
        warnings.append("Both debit and credit were populated in the source row.")

    return ExtractedTransaction(
        transaction_date=date_value,
        description=description,
        debit=debit,
        credit=credit,
        balance=balance,
        reference_number=reference,
        account_number=str(metadata.get("account_number", "")),
        statement_period=str(metadata.get("statement_period", "")),
        opening_balance=metadata.get("opening_balance"),
        closing_balance=metadata.get("closing_balance"),
        source_pdf=file_name,
        source_page=page,
        raw_text=" | ".join(row),
        confidence=0.86 if not warnings else 0.68,
        notes="; ".join(warnings),
        warnings=warnings,
    )


def _transactions_from_lines(
    file_name: str,
    raw_text: str,
    metadata: dict[str, Any],
) -> list[ExtractedTransaction]:
    transactions: list[ExtractedTransaction] = []
    current_page: int | None = None
    for line in raw_text.splitlines():
        page_match = re.match(r"--- Page (\d+)", line)
        if page_match:
            current_page = int(page_match.group(1))
            continue

        line = " ".join(line.split())
        if not line:
            continue
        date_match = DATE_PATTERN.search(line)
        if not date_match:
            continue

        amounts = list(AMOUNT_PATTERN.finditer(line))
        if not amounts:
            continue

        transaction = _transaction_from_line(
            line=line,
            date_match=date_match,
            amounts=amounts,
            file_name=file_name,
            page=current_page,
            metadata=metadata,
        )
        if transaction:
            transactions.append(transaction)

    return _deduplicate_transactions(transactions)


def _transaction_from_line(
    line: str,
    date_match: re.Match[str],
    amounts: list[re.Match[str]],
    file_name: str,
    page: int | None,
    metadata: dict[str, Any],
) -> ExtractedTransaction | None:
    date_text = date_match.group("date")
    normalized_date = normalize_date(date_text)
    amount_values = [parse_amount(match.group("amount")) for match in amounts]
    amount_values = [value for value in amount_values if value is not None]
    if not amount_values:
        return None

    balance = amount_values[-1] if len(amount_values) >= 2 else None
    transaction_amount = amount_values[-2] if len(amount_values) >= 2 else amount_values[-1]
    amount_start = amounts[-2].start() if len(amounts) >= 2 else amounts[-1].start()
    description = line[date_match.end() : amount_start].strip(" -:\t")
    if not description:
        description = "Unlabeled transaction"

    warnings = []
    debit: float | None = None
    credit: float | None = None
    upper_line = line.upper()
    amount_text = amounts[-2].group("amount") if len(amounts) >= 2 else amounts[-1].group("amount")

    if transaction_amount < 0 or _amount_text_is_negative(amount_text):
        debit = abs(transaction_amount)
    elif any(
        token in upper_line
        for token in [
            " CR",
            "CREDIT",
            "DEPOSIT",
            "REFUND",
            "TRANSFER FROM",
            "WIRE IN",
            "ACH CREDIT",
            "ZELLE PAYMENT FROM",
            "INTEREST PAID",
        ]
    ):
        credit = abs(transaction_amount)
    elif any(
        token in upper_line
        for token in [
            " DR",
            "DEBIT",
            "WITHDRAWAL",
            "PURCHASE",
            "PAYMENT TO",
            "FEE",
            "ATM",
            "POS",
            "CHECKCARD",
            "MOBILE PURCHASE",
            "ZELLE PAYMENT TO",
        ]
    ):
        debit = abs(transaction_amount)
    else:
        credit = abs(transaction_amount)
        warnings.append(
            "Amount direction was ambiguous; assumed credit because the amount was positive."
        )

    if date_text and not normalized_date:
        warnings.append("Could not normalize transaction date.")

    return ExtractedTransaction(
        transaction_date=normalized_date,
        description=description,
        debit=debit,
        credit=credit,
        balance=balance,
        account_number=str(metadata.get("account_number", "")),
        statement_period=str(metadata.get("statement_period", "")),
        opening_balance=metadata.get("opening_balance"),
        closing_balance=metadata.get("closing_balance"),
        source_pdf=file_name,
        source_page=page,
        raw_text=line,
        confidence=0.66 if warnings else 0.76,
        notes="; ".join(warnings),
        warnings=warnings,
    )


def parse_amount(value: Any) -> float | None:
    """Normalize currency text into a signed float."""

    if value is None:
        return None
    text = str(value).strip()
    if not text or text in {"-", "--"}:
        return None

    negative = _amount_text_is_negative(text)
    text = re.sub(r"\b(?:CR|DR)\b", "", text, flags=re.IGNORECASE)
    text = (
        text.replace("$", "")
        .replace(",", "")
        .replace("(", "")
        .replace(")", "")
        .strip()
    )
    try:
        amount = float(text)
    except ValueError:
        return None
    return -abs(amount) if negative else amount


def normalize_date(value: Any) -> str:
    """Normalize common bank-statement date values to ISO date strings."""

    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    try:
        from dateutil import parser

        parsed = parser.parse(text, fuzzy=False, dayfirst=False)
        return parsed.date().isoformat()
    except Exception:
        pass

    formats = (
        "%m/%d/%Y",
        "%m/%d/%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y-%m-%d",
        "%B %d, %Y",
        "%b %d, %Y",
        "%B %d %Y",
        "%b %d %Y",
        "%B %d, %y",
        "%b %d, %y",
    )
    for date_format in formats:
        try:
            return datetime.strptime(text, date_format).date().isoformat()
        except ValueError:
            continue
    return ""


def _deduplicate_transactions(
    transactions: list[ExtractedTransaction],
) -> list[ExtractedTransaction]:
    seen = set()
    unique_transactions = []
    for transaction in transactions:
        key = (
            transaction.transaction_date,
            transaction.description.lower().strip(),
            transaction.debit,
            transaction.credit,
            transaction.balance,
            transaction.source_pdf,
            transaction.source_page,
        )
        if key in seen:
            continue
        seen.add(key)
        unique_transactions.append(transaction)
    return unique_transactions


def _first_match(patterns: list[str], text: str) -> re.Match[str] | None:
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match
    return None


def _amount_text_is_negative(text: str) -> bool:
    stripped = text.strip()
    return stripped.startswith("-") or (
        stripped.startswith("(") and stripped.endswith(")")
    )


def _normalize_label(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def _clean_cell(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).split()).strip()

