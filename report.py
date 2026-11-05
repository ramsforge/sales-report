"""Turn a messy sales CSV into a clean Excel report.

Usage: python report.py sample_sales.csv report.xlsx
"""
import csv
import sys
from collections import defaultdict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill

DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y")


def parse_date(text):
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            pass
    return None


def parse_number(text):
    try:
        return float(text.replace("$", "").replace(",", "").strip())
    except ValueError:
        return None


def clean(rows):
    """Return (clean_rows, problems). Normalizes text, drops bad rows and duplicates."""
    good, problems, seen = [], [], set()
    for line, row in enumerate(rows, start=2):  # line 1 is the header
        date = parse_date(row["date"])
        qty = parse_number(row["quantity"])
        price = parse_number(row["unit_price"])
        if date is None or qty is None or price is None:
            problems.append((line, "invalid date, quantity or price", dict(row)))
            continue
        record = (date, row["product"].strip().title(), row["region"].strip().title(), int(qty), price)
        if record in seen:
            problems.append((line, "duplicate row", dict(row)))
            continue
        seen.add(record)
        good.append(record)
    return good, problems


def summarize(records, key_index):
    totals = defaultdict(float)
    for r in records:
        totals[r[key_index]] += r[3] * r[4]
    return sorted(totals.items(), key=lambda kv: kv[1], reverse=True)


def write_sheet(ws, header, rows):
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")
    for row in rows:
        ws.append(list(row))
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = max(len(str(c.value)) for c in col) + 4


def build_report(records, problems, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Clean Data"
    write_sheet(ws, ["Date", "Product", "Region", "Quantity", "Unit Price", "Revenue"],
                [(*r, r[3] * r[4]) for r in records])
    for cell in ws["A"][1:]:
        cell.number_format = "yyyy-mm-dd"

    for title, idx in (("By Product", 1), ("By Region", 2)):
        s = wb.create_sheet(title)
        data = summarize(records, idx)
        write_sheet(s, [title.split()[-1], "Revenue"], data)
        chart = BarChart()
        chart.title = f"Revenue {title.lower()}"
        chart.add_data(Reference(s, min_col=2, min_row=1, max_row=len(data) + 1), titles_from_data=True)
        chart.set_categories(Reference(s, min_col=1, min_row=2, max_row=len(data) + 1))
        s.add_chart(chart, "D2")

    write_sheet(wb.create_sheet("Problems"), ["CSV Line", "Reason", "Original Row"],
                [(line, reason, str(row)) for line, reason, row in problems])
    wb.save(out_path)


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    with open(sys.argv[1], newline="", encoding="utf-8") as f:
        records, problems = clean(csv.DictReader(f))
    build_report(records, problems, sys.argv[2])
    print(f"{len(records)} clean rows, {len(problems)} problems -> {sys.argv[2]}")


if __name__ == "__main__":
    main()
