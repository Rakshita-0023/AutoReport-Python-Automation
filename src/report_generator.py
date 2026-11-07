"""Build the Excel snapshot and replace the target only after a successful save."""

from decimal import Decimal
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrData, StrRef, StrVal
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .cleaner import ReportError, summarize

NAVY = "203B55"
TEAL = "00796B"
MUTED = "526579"
CURRENCY = '"₹"#,##0.00'
COUNT = "#,##0"


def heading(sheet, title, context):
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.tabColor = NAVY
    sheet["A2"] = title
    sheet["A2"].font = Font(name="Arial", size=16, bold=True, color=NAVY)
    sheet.row_dimensions[2].height = 24
    sheet["A3"] = context
    sheet["A3"].font = Font(name="Arial", size=10, italic=True, color=MUTED)


def table(sheet, start, headers, rows, formats, *, filter_rows=False):
    for column, label in enumerate(headers, 1):
        cell = sheet.cell(start, column, label)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[start].height = 23
    last = start
    for last, values in enumerate(rows, start + 1):
        for column, value in enumerate(values, 1):
            cell = sheet.cell(last, column, value)
            # Source labels are always literal, including values that start with "=".
            if isinstance(value, str):
                cell.data_type = "s"
            cell.font = Font(name="Arial", size=10, color=NAVY)
            cell.alignment = Alignment(vertical="center", wrap_text=headers[column - 1] == "Details")
            if last % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="EFF5F8")
            if column in formats:
                cell.number_format = formats[column]
        if headers[-1] == "Details" and len(values[-1]) > 85:
            sheet.row_dimensions[last].height = 48
    if filter_rows and last > start:
        sheet.auto_filter.ref = f"A{start}:{get_column_letter(len(headers))}{last}"
    return last


def widths(sheet, *values):
    for column, width in enumerate(values, 1):
        sheet.column_dimensions[get_column_letter(column)].width = width


def revenue_chart(sheet, start, end, title, anchor, *, trend=False):
    chart = LineChart() if trend else BarChart()
    if not trend:
        chart.type = "bar"
        chart.x_axis.scaling.orientation = "maxMin"
        chart.y_axis.crosses = "max"
    chart.title = title
    chart.legend = None
    chart.width = 17
    chart.height = 8
    chart.add_data(Reference(sheet, min_col=4, min_row=start + 1, max_row=end))
    categories = Reference(sheet, min_col=1, min_row=start + 1, max_row=end)
    series = chart.series[0]
    labels = [StrVal(idx=index, v=sheet.cell(row, 1).value)
              for index, row in enumerate(range(start + 1, end + 1))]
    cache = StrData(ptCount=len(labels), pt=labels)
    series.cat = AxDataSource(strRef=StrRef(f=str(categories), strCache=cache))
    if trend:
        series.graphicalProperties.line.solidFill = TEAL
        series.marker.symbol = "circle"
        series.smooth = False
    else:
        series.graphicalProperties.solidFill = TEAL
        series.graphicalProperties.line.noFill = True
    chart.y_axis.numFmt = "#,##0"
    sheet.add_chart(chart, anchor)


def write_report(sales, audit, source_name: str, output: Path) -> None:
    total = sum(sales["revenue"])
    if total > Decimal("9999999999999.99"):
        raise ReportError("Total revenue exceeds Excel's cent precision. Report a shorter period.")

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Sales Summary"
    products = workbook.create_sheet("Product Performance")
    regions = workbook.create_sheet("Region Performance")
    detail = workbook.create_sheet("Cleaned Sales Data")
    quality = workbook.create_sheet("Data Quality Report")

    monthly = summarize(sales.assign(month=sales["sale_date"].map(lambda day: day.strftime("%Y-%m"))),
                        "month")
    first, last = sales["sale_date"].min(), sales["sale_date"].max()
    counts = audit["action"].value_counts()
    duplicates = int(counts.get("Duplicate removed", 0))
    excluded = int(counts.get("Excluded", 0))

    heading(summary, "Sales Summary", f"{first:%d %b %Y} to {last:%d %b %Y} | Accepted sales only")
    widths(summary, 24, 16, 16, 22, 4)
    summary.sheet_view.zoomScale = 85
    table(summary, 5, ("Revenue (INR)", "Sales", "Units", "Rows excluded"),
          [(total, len(sales), int(sales["quantity"].sum()), excluded)],
          {1: CURRENCY, 2: COUNT, 3: COUNT, 4: COUNT})
    summary.row_dimensions[6].height = 23
    for cell in summary[6][:4]:
        cell.font = Font(name="Arial", size=12, bold=True, color=NAVY)

    metrics = ("Sales", "Units", "Revenue (INR)")
    formats = {2: COUNT, 3: COUNT, 4: CURRENCY}
    end = table(summary, 9, ("Month", *metrics), monthly.itertuples(index=False, name=None), formats)
    revenue_chart(summary, 9, end, "Monthly revenue (INR)", "F5", trend=True)

    for sheet, key in ((products, "product"), (regions, "region")):
        results = summarize(sales, key)
        heading(sheet, sheet.title, "Accepted sales, ranked by revenue")
        widths(sheet, 26, 14, 14, 22, 4)
        sheet.sheet_view.zoomScale = 85
        end = table(sheet, 5, (key.title(), *metrics), results.itertuples(index=False, name=None),
                    formats, filter_rows=True)
        sheet.freeze_panes = "A6"
        revenue_chart(sheet, 5, min(end, 13), f"Leading {key}s by revenue (INR)", "F5")

    heading(detail, "Cleaned Sales Data", "Source row refers to the input record, including the header.")
    widths(detail, 15, 17, 17, 27, 17, 14, 22, 22)
    table(detail, 5, ("Source row", "Sale ID", "Sale date", "Product", "Region", "Quantity",
                      "Unit price (INR)", "Revenue (INR)"),
          sales.itertuples(index=False, name=None),
          {3: "yyyy-mm-dd", 6: COUNT, 7: CURRENCY, 8: CURRENCY}, filter_rows=True)
    detail.freeze_panes = "C6"

    heading(quality, "Data Quality Report", f"Source file · {source_name}")
    widths(quality, 26, 17, 24, 88)
    quality.sheet_properties.tabColor = "A46021"
    table(quality, 5, ("Check", "Rows"), [
        ("Input sales rows", len(sales) + duplicates + excluded),
        ("Accepted rows", len(sales)),
        ("Duplicates removed", duplicates),
        ("Rows excluded", excluded),
        ("Accepted rows normalized", int(counts.get("Normalized", 0))),
    ], {2: COUNT})
    quality["A11"] = "Excluded rows need source corrections before they can count toward revenue."
    quality["A11"].font = Font(name="Arial", size=10, italic=True, color=MUTED)
    table(quality, 13, ("Source row", "Sale ID", "Action", "Details"),
          audit.itertuples(index=False, name=None), {}, filter_rows=True)
    quality.freeze_panes = "A14"

    output.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as temporary:
        temp_path = Path(temporary.name)
    try:
        workbook.save(temp_path)
        temp_path.replace(output)
    finally:
        temp_path.unlink(missing_ok=True)
