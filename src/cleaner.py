"""Read, validate and summarize single-currency sales exports."""

import csv
import logging
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from zipfile import BadZipFile

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

COLUMNS = ("sale_id", "sale_date", "product", "region", "quantity", "unit_price")
DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y")
LOGGER = logging.getLogger(__name__)
FORMULA = object()


class ReportError(ValueError):
    """An input or output problem the operator can correct."""


def read_sales(path: Path) -> pd.DataFrame:
    """Read UTF-8 CSV or the first XLSX worksheet, retaining source row numbers."""
    extension = path.suffix.lower()
    if extension not in {".csv", ".xlsx"}:
        raise ReportError("Use a .csv or .xlsx input file.")
    try:
        if extension == ".csv":
            with path.open(encoding="utf-8-sig", newline="") as source:
                records = list(csv.reader(source, strict=True))
        else:
            workbook = load_workbook(path, read_only=True, data_only=False)
            try:
                records = [
                    tuple(FORMULA if cell.data_type == "f" else cell.value for cell in row)
                    for row in workbook.worksheets[0].iter_rows()
                ]
            finally:
                workbook.close()
    except (OSError, UnicodeError, csv.Error, BadZipFile, InvalidFileException,
            KeyError, ValueError, SyntaxError) as exc:
        raise ReportError(f"Cannot read {path.name}. Check the file format and access. ({exc})") from exc

    if not records or not any(value is not None and str(value).strip() for value in records[0]):
        raise ReportError("The input is empty or has no header in the first row.")

    headers = [str(value).strip().lower() if value is not None else "" for value in records[0]]
    named_headers = [name for name in headers if name]
    if len(named_headers) != len(set(named_headers)):
        raise ReportError("The header contains duplicate column names.")
    missing = [name for name in COLUMNS if name not in headers]
    if missing:
        raise ReportError(f"Missing required columns: {', '.join(missing)}.")

    positions = [headers.index(name) for name in COLUMNS]
    rows = {}
    for number, values in enumerate(records[1:], start=2):
        if not any(value is not None and str(value).strip() for value in values):
            continue
        if len(values) != len(headers):
            raise ReportError(f"Source row {number} has {len(values)} fields. Expected {len(headers)}.")
        selected = [values[index] for index in positions]
        formula_fields = [field for field, value in zip(COLUMNS, selected) if value is FORMULA]
        if formula_fields:
            raise ReportError(
                f"Source row {number} contains formulas in {', '.join(formula_fields)}. "
                "Export these columns as values and rerun."
            )
        rows[number] = selected
    if not rows:
        raise ReportError("The input has headers but no sales rows.")
    return pd.DataFrame.from_dict(rows, orient="index", columns=COLUMNS, dtype=object)


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for pattern in DATE_FORMATS:
        try:
            return datetime.strptime(value, pattern).date()
        except ValueError:
            continue
    return None


def clean_sales(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return accepted sales and one audit entry for each affected source row."""
    sales = []
    events = []
    rows = raw.map(lambda value: "" if pd.isna(value) else str(value).strip())
    duplicates = rows.duplicated()
    conflicts = rows.loc[~duplicates, "sale_id"].duplicated(keep=False)
    conflicting_ids = set(conflicts[conflicts].index)
    for number, sale_id, raw_date, product, region, quantity, price in rows.itertuples():
        if duplicates.loc[number]:
            events.append((number, sale_id, "Duplicate removed", "Repeated sales fields after trimming."))
            continue
        problems = [f"Missing {field}" for field in ("sale_id", "sale_date", "product",
                    "region", "quantity", "unit_price") if not rows.at[number, field]]
        original_date = raw.at[number, "sale_date"]
        parsed_date = parse_date(original_date if isinstance(original_date, date) else raw_date)
        if raw_date and (parsed_date is None or parsed_date.year < 1900):
            problems.append("Invalid sale_date")
        if quantity and (not quantity.isascii() or not quantity.isdecimal()
                         or len(quantity) > 9 or int(quantity) == 0):
            problems.append("Quantity must be a positive whole number (up to 9 digits)")
        # Plain INR values only. Reject excess precision instead of rounding the input.
        whole, point, fraction = price.partition(".")
        digits = whole + fraction
        if price and (not digits.isascii() or not digits.isdecimal() or not whole
                      or len(whole) > 9 or (point and not 1 <= len(fraction) <= 2)):
            problems.append("Unit price must be a nonnegative INR amount with up to 2 decimals")
        if sale_id and number in conflicting_ids:
            problems.append("Sale ID appears on different rows")
        if problems:
            events.append((number, sale_id, "Excluded", " / ".join(problems)))
            continue

        units = int(quantity)
        unit_price = Decimal(price)
        revenue = units * unit_price
        if revenue > Decimal("9999999999999.99"):
            events.append((number, sale_id, "Excluded", "Line revenue exceeds Excel's cent precision."))
            continue
        sales.append((number, sale_id, parsed_date, product, region, units, unit_price, revenue))
        changes = []
        if raw_date != parsed_date.isoformat():
            changes.append("Date normalized to " + parsed_date.isoformat())
        if any(isinstance(original, str) and original != original.strip()
               for original in raw.loc[number]):
            changes.append("Trimmed surrounding whitespace")
        if changes:
            events.append((number, sale_id, "Normalized", " / ".join(changes)))

    cleaned = pd.DataFrame(sales, columns=("source_row", *COLUMNS, "revenue"))
    audit = pd.DataFrame(events, columns=("source_row", "sale_id", "action", "details"))
    counts = audit["action"].value_counts()
    LOGGER.info("Read %d rows | kept %d | duplicates removed %d | excluded %d | normalized %d",
                len(raw), len(cleaned), counts.get("Duplicate removed", 0),
                counts.get("Excluded", 0), counts.get("Normalized", 0))
    if cleaned.empty:
        examples = " / ".join(f"row {row.source_row}: {row.details}"
                              for row in audit.itertuples() if row.action == "Excluded")
        raise ReportError("No valid sales remain. " + examples[:450])
    return cleaned.sort_values(["sale_date", "sale_id"]), audit


def summarize(sales: pd.DataFrame, key: str) -> pd.DataFrame:
    """Aggregate accepted sales with exact decimal arithmetic."""
    result = sales.groupby(key, as_index=False).agg(
        sales=("sale_id", "size"), units=("quantity", "sum"), revenue=("revenue", "sum")
    )
    if key != "month":
        result = result.sort_values(["revenue", key], ascending=[False, True])
    return result
