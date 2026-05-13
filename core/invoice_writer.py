import shutil
from datetime import date
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import column_index_from_string
from config import (
    INVOICE_TEMPLATE_PATH,
    INV_ROW_INV_NO,
    INV_ROW_INV_DATE,
    INV_ROW_PARTY,
    INV_ROW_ADDR1,
    INV_ROW_ADDR2,
    INV_ROW_SERVICE,
    INV_ROW_TOTAL,
    INV_ROW_WORDS,
    INV_ROW_PERIOD,
    INV_COL_VALUE,
    INV_COL_AMT,
)
from core.amount_words import to_indian_words


def _safe_set(ws, col_letter: str, row: int, value):
    """Write to a cell only if it is not a non-anchor MergedCell."""
    col_idx = column_index_from_string(col_letter)
    cell = ws.cell(row=row, column=col_idx)
    if isinstance(cell, MergedCell):
        return
    cell.value = value


def write_invoice(
    party_full_name: str,
    party_address: str,
    invoice_no: str,
    invoice_date: date,
    billing_month: str,
    total_amount: float,
    out_path: str,
) -> None:
    """Copy Invoice template and fill all required fields."""
    shutil.copy2(INVOICE_TEMPLATE_PATH, out_path)

    wb = openpyxl.load_workbook(out_path)
    ws = wb.active

    _safe_set(ws, INV_COL_VALUE, INV_ROW_INV_NO,   invoice_no)
    _safe_set(ws, INV_COL_VALUE, INV_ROW_INV_DATE,  invoice_date.strftime("%d.%m.%Y"))
    _safe_set(ws, "B",           INV_ROW_PARTY,      party_full_name)

    addr1, addr2 = _split_address(str(party_address))
    _safe_set(ws, "B", INV_ROW_ADDR1, addr1)
    _safe_set(ws, "B", INV_ROW_ADDR2, addr2)

    _safe_set(ws, "B",        INV_ROW_SERVICE, _service_description(billing_month))
    _safe_set(ws, INV_COL_AMT, INV_ROW_SERVICE, total_amount)
    _safe_set(ws, INV_COL_AMT, INV_ROW_TOTAL,   total_amount)
    _safe_set(ws, "B",        INV_ROW_WORDS,   to_indian_words(total_amount))
    _safe_set(ws, "B",        INV_ROW_PERIOD,  _invoice_period(billing_month))

    try:
        wb.save(out_path)
    except Exception as e:
        raise ValueError(f"Could not save Invoice file: {e}")


def _split_address(address: str) -> tuple[str, str]:
    address = address.strip()
    last_comma = address.rfind(",")
    if last_comma == -1:
        return address, ""
    return address[:last_comma].strip(), address[last_comma + 1:].strip()


def _service_description(billing_month: str) -> str:
    return f'"Lab Test Service" – For the period {billing_month}'


def _invoice_period(billing_month: str) -> str:
    from datetime import datetime
    import calendar

    parsed = None
    for fmt in ("%b'%y", "%B %Y", "%b %Y", "%B'%y"):
        try:
            parsed = datetime.strptime(billing_month, fmt)
            break
        except ValueError:
            continue

    if parsed is None:
        return f"1st {billing_month} to 31 {billing_month}"

    last_day = calendar.monthrange(parsed.year, parsed.month)[1]
    month_abbr = parsed.strftime("%b")
    year_short = parsed.strftime("'%y")
    return f"1st {month_abbr}{year_short} to {last_day} {month_abbr}{year_short}"
