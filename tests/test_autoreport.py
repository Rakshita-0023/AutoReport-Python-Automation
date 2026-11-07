import csv
from datetime import date
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook

from src.cleaner import COLUMNS, ReportError, clean_sales, read_sales
from generate_sample import generate
from main import main

VALID = ("0001", "04/05/2025", "Stand", "North", "2", "12.15")


def write_csv(path, rows):
    with path.open("w", encoding="utf-8", newline="") as target:
        csv.writer(target).writerows(rows)
    return path


@pytest.mark.parametrize(("rows", "message"), [
    ([], "empty"),
    ([COLUMNS], "no sales rows"),
    ([("sale_id",)], "Missing required columns"),
    ([(*COLUMNS, " SALE_ID ")], "duplicate column"),
    ([COLUMNS, ("A", "2025-01-01")], "row 2 has 2 fields"),
])
def test_bad_input(tmp_path, rows, message):
    with pytest.raises(ReportError, match=message):
        read_sales(write_csv(tmp_path / "input.csv", rows))


def test_cleaning(tmp_path):
    path = write_csv(tmp_path / "input.csv", [
        COLUMNS, VALID, VALID,
        ("B", "2025-05-05", "Stand", "North", "1", ""),
        ("C", "2025-05-05", "Stand", "North", "1", "0"),
        ("D", "31/02/2025", "Stand", "North", "two", "-2"),
        ("E", "2025-05-05", "Stand", "North", "1", "2"),
        ("E", "2025-05-05", "Stand", "North", "1", "3"),
    ])
    sales, audit = clean_sales(read_sales(path))
    assert list(sales.sale_id) == ["0001", "C"]
    assert list(sales.sale_date) == [date(2025, 5, 4), date(2025, 5, 5)]
    assert list(sales.revenue) == [Decimal("24.30"), Decimal("0")]
    assert audit.action.value_counts().to_dict() == {
        "Excluded": 4, "Duplicate removed": 1, "Normalized": 1,
    }
    assert "Missing unit_price" in audit.set_index("source_row").loc[4, "details"]
    assert "Invalid sale_date" in audit.set_index("source_row").loc[6, "details"]


def test_cli_refuses_invalid_data_without_replacing_output(tmp_path, caplog):
    path = write_csv(tmp_path / "input.csv", [COLUMNS, (*VALID[:-1], "")])
    output = tmp_path / "report.xlsx"
    output.write_bytes(b"previous report")
    assert main(["--input", str(path), "--output", str(output)]) == 1
    assert "No valid sales remain" in caplog.text
    assert "Missing unit_price" in caplog.text
    assert output.read_bytes() == b"previous report"


def test_xlsx_round_trip_and_literal_labels(tmp_path):
    source = tmp_path / "input.xlsx"
    workbook = Workbook()
    workbook.active.append(COLUMNS)
    workbook.active.append(["0001", date(2025, 5, 4), "=2+2", "North", 2, 12.15])
    workbook.active["C2"].data_type = "s"
    workbook.save(source)
    assert main(["--input", str(source), "--output", str(source)]) == 1

    output = tmp_path / "report.xlsx"
    assert main(["--input", str(source), "--output", str(output)]) == 0
    report = load_workbook(output)
    detail = report["Cleaned Sales Data"]
    assert detail["B6"].value == "0001"
    assert detail["C6"].value.date() == date(2025, 5, 4)
    assert detail["D6"].value == "=2+2" and detail["D6"].data_type == "s"
    assert detail["H6"].value == 24.30


def test_sample_report_reconciles(tmp_path):
    source = tmp_path / "sample.csv"
    generate(source)
    with source.open(newline="", encoding="utf-8") as stream:
        assert len(list(csv.DictReader(stream))) == 1000
    output = tmp_path / "nested" / "report.xlsx"
    assert main(["--input", str(source), "--output", str(output)]) == 0

    report = load_workbook(output)
    assert report.sheetnames == ["Sales Summary", "Product Performance", "Region Performance",
                                 "Cleaned Sales Data", "Data Quality Report"]
    summary, detail, quality = (report[name] for name in
                                ("Sales Summary", "Cleaned Sales Data", "Data Quality Report"))
    assert detail.freeze_panes == "C6"
    assert detail.auto_filter.ref == "A5:H947"
    assert "₹" in detail["H6"].number_format
    assert [quality[f"B{row}"].value for row in range(6, 10)] == [1000, 942, 30, 28]
    total = Decimal(str(summary["A6"].value))
    assert sum(Decimal(str(cell.value)) for cell in detail["H"][5:]) == total
    for sheet, start in [(summary, 10), (report["Product Performance"], 6),
                         (report["Region Performance"], 6)]:
        assert len(sheet._charts) == 1
        assert sum(Decimal(str(cell.value)) for cell in sheet["D"][start - 1:]) == total


def test_xlsx_requires_values_in_sales_columns(tmp_path):
    path = tmp_path / "input.xlsx"
    workbook = Workbook()
    workbook.active.append(COLUMNS)
    workbook.active.append([*VALID[:-1], "=2+2"])
    workbook.save(path)
    with pytest.raises(ReportError, match="Source row 2 contains formulas in unit_price"):
        read_sales(path)
