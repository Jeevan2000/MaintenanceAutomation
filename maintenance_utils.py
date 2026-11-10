"""Workbook helpers for the Ferro Towers append-only maintenance ledger."""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import calendar
import re

from openpyxl import load_workbook

MONTHLY_MAINTENANCE_AMOUNT = 1500
OWNER_MASTER = {
    "101": "SANIDHYA SURANA",
    "102": "SUNIL CHHAJED",
    "201": "SANIDHYA SURANA",
    "202": "SAMPATLAL DESLADA",
    "301": "SANDIP LAGAD",
    "302": "SWANAND THORAT",
    "401": "DIPAK BADGUJAR",
    "402": "VIJAY KHADE",
    "501": "MACCHINDRA RASKAR",
    "502": "SUDHIR SHIRKE",
    "601": "PRACHEER JOSHI",
}
REQUIRED_HEADERS = {
    "flat": ("flat no", "flat no."),
    "owner": ("owners", "owner"),
    "pending": ("pending balance",),
    "date": ("paid date",),
    "amount": ("amount",),
    "balance": ("balance",),
    "paid_till": ("paid till",),
}


def normalize_text(value) -> str:
    return "" if value is None else str(value).strip()


def parse_date(value) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _header_key(value) -> str:
    return re.sub(r"\s+", " ", normalize_text(value).lower()).rstrip(".")


def _find_ledger_sheet(workbook):
    for ws in workbook.worksheets:
        for row_idx in range(1, min(ws.max_row, 20) + 1):
            headers = {_header_key(ws.cell(row_idx, col).value): col for col in range(1, ws.max_column + 1)}
            found = {}
            for key, aliases in REQUIRED_HEADERS.items():
                found[key] = next((headers[_header_key(alias)] for alias in aliases if _header_key(alias) in headers), None)
            if all(found.values()):
                return ws, row_idx, found
    raise ValueError("Could not find a worksheet with Flat No., Owners, Pending Balance, Paid Date, Amount, Balance, and Paid Till headers")


def _flat_key(value) -> str:
    text = normalize_text(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _number(value, field_name: str) -> Decimal:
    if value is None or normalize_text(value) == "":
        return Decimal("0")
    try:
        parsed = Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise ValueError(f"Existing {field_name} is not numeric: {value!r}") from exc
    if not parsed.is_finite():
        raise ValueError(f"Existing {field_name} is not a finite number")
    return parsed


def _numeric_output(value: Decimal):
    return int(value) if value == value.to_integral_value() else float(value)


def _month_start(value: date) -> date:
    return date(value.year, value.month, 1)


def _add_months(value: date, months: int) -> date:
    month_index = value.year * 12 + value.month - 1 + months
    year, month_index = divmod(month_index, 12)
    month = month_index + 1
    day = min(value.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def _paid_till_date(value) -> date | None:
    if value is None or normalize_text(value) == "":
        return None
    parsed = parse_date(value)
    if parsed is None:
        raise ValueError(f"Existing Paid Till is not a valid date: {value!r}")
    return parsed


def append_payment(path: Path, flat: str, amount: float, paid_date: date) -> dict:
    """Append a payment and calculate carried-forward balance and Paid Till date.

    The latest payment is located by scanning from the bottom. On a flat's first
    payment, its existing Balance and Paid Till in the master row are opening values.
    """
    flat = _flat_key(flat)
    if flat not in OWNER_MASTER:
        raise ValueError(f"Unknown flat number: {flat}")
    try:
        amount = Decimal(str(amount))
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise ValueError("Amount must be a positive number") from exc
    if not amount.is_finite() or amount <= 0:
        raise ValueError("Amount must be a positive number")
    if not isinstance(paid_date, date):
        paid_date = parse_date(paid_date)
    if paid_date is None:
        raise ValueError("A valid payment date is required")

    wb = load_workbook(path)
    ws, header_row, cols = _find_ledger_sheet(wb)
    roster_rows = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        if _flat_key(ws.cell(row_idx, cols["flat"]).value) == flat:
            roster_rows.append(row_idx)
    if not roster_rows:
        raise ValueError(f"Flat {flat} is missing from the workbook master list")

    latest_payment_row = next(
        (row_idx for row_idx in reversed(roster_rows)
         if ws.cell(row_idx, cols["date"]).value not in (None, "")
         or ws.cell(row_idx, cols["amount"]).value not in (None, "")),
        None,
    )
    if latest_payment_row is None:
        opening_row = roster_rows[0]
        opening_balance = ws.cell(opening_row, cols["balance"]).value
        if opening_balance in (None, ""):
            opening_balance = ws.cell(opening_row, cols["pending"]).value
        pending_balance = _number(opening_balance, "Balance")
        previous_paid_till = _paid_till_date(ws.cell(opening_row, cols["paid_till"]).value)
    else:
        pending_balance = _number(ws.cell(latest_payment_row, cols["balance"]).value, "Balance")
        previous_paid_till = _paid_till_date(ws.cell(latest_payment_row, cols["paid_till"]).value)

    balance = pending_balance - amount
    months_paid = int(amount // Decimal(MONTHLY_MAINTENANCE_AMOUNT))
    if previous_paid_till is not None:
        paid_till = _add_months(previous_paid_till, months_paid)
    elif months_paid > 0:
        paid_till = _add_months(_month_start(paid_date), months_paid - 1)
    else:
        paid_till = None
    last_data_row = max(
        row_idx for row_idx in range(header_row + 1, ws.max_row + 1)
        if normalize_text(ws.cell(row_idx, cols["flat"]).value)
    )
    new_row = last_data_row + 1
    # Copy formatting from the prior ledger row if available, without copying values.
    template_row = latest_payment_row or roster_rows[-1]
    for col in range(1, ws.max_column + 1):
        source = ws.cell(template_row, col)
        target = ws.cell(new_row, col)
        if source.has_style:
            target._style = source._style

    values = {
        "flat": int(flat) if flat.isdigit() else flat,
        "owner": OWNER_MASTER[flat],
        "pending": _numeric_output(pending_balance),
        "date": datetime.combine(paid_date, datetime.min.time()),
        "amount": _numeric_output(amount),
        "balance": _numeric_output(balance),
        "paid_till": datetime.combine(paid_till, datetime.min.time()) if paid_till else None,
    }
    for key, value in values.items():
        ws.cell(new_row, cols[key]).value = value
    currency_format = r'"₹"\ #,##0.00'
    long_date_format = r"[$-F800]dddd\,\ mmmm\ dd\,\ yyyy"
    for key in ("pending", "amount", "balance"):
        ws.cell(new_row, cols[key]).number_format = currency_format
    for key in ("date", "paid_till"):
        ws.cell(new_row, cols[key]).number_format = long_date_format
    wb.save(path)
    values["row"] = new_row
    return values


def get_maintenance_report(path: Path, as_of: date | None = None) -> dict:
    """Return each flat's balance and overdue-month count as of the report month."""
    as_of = as_of or date.today()
    wb = load_workbook(path, data_only=True, read_only=True)
    ws, header_row, cols = _find_ledger_sheet(wb)
    last_rows = {}
    for row_idx in range(header_row + 1, ws.max_row + 1):
        flat = _flat_key(ws.cell(row_idx, cols["flat"]).value)
        if flat in OWNER_MASTER:
            last_rows[flat] = row_idx

    rows = []
    for flat, owner in OWNER_MASTER.items():
        row_idx = last_rows.get(flat)
        pending = None
        paid_till = None
        if row_idx is not None:
            balance_value = ws.cell(row_idx, cols["balance"]).value
            pending = _number(balance_value, f"Balance for flat {flat}")
            paid_till = _paid_till_date(ws.cell(row_idx, cols["paid_till"]).value)

        if pending is None or paid_till is None:
            overdue_months = 0
            reported_balance = pending or Decimal("0")
        else:
            elapsed_months = max(
                0,
                (as_of.year - paid_till.year) * 12 + as_of.month - paid_till.month,
            )
            reported_balance = pending + elapsed_months * Decimal(MONTHLY_MAINTENANCE_AMOUNT)
            credit_months = int((-pending) // Decimal(MONTHLY_MAINTENANCE_AMOUNT)) if pending < 0 else 0
            overdue_months = max(0, elapsed_months - credit_months)
        rows.append({
            "flat": flat,
            "owner": owner,
            "pending_balance": _numeric_output(reported_balance),
            "paid_till": paid_till,
            "pending_months": overdue_months,
        })

    wb.close()
    return {"as_of": as_of, "rows": rows}
