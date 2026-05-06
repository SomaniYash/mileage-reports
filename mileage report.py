import os, re, glob, csv, io
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def parse_csv_file(filepath):
    """Parse a mileage CSV and return a dict with staff, member, KMs, parking."""
    with open(filepath, encoding="utf-8") as f:
        content = f.read()
    lines = [l.rstrip() for l in content.strip().split("\n")]

    header_match = re.search(
        r"Name:\s*(.+?),\s*Member:\s*(.+?),\s*Month/Year:\s*(.+)", lines[0]
    )
    if not header_match:
        return None

    staff_name  = header_match.group(1).strip()
    member_name = header_match.group(2).strip()
    month_year  = header_match.group(3).strip()

    total_km = 0.0
    total_parking = 0.0

    for line in lines:
        km_match = re.match(r"Total Kilometers,\s*([\d.]+)", line)
        if km_match:
            total_km = float(km_match.group(1))

    data_started = False
    for line in lines:
        if line.startswith("Date,"):
            data_started = True
            continue
        if data_started:
            if line == "" or line.startswith(("Total", "Approved", "Month Billed")):
                continue
            try:
                row = next(csv.reader(io.StringIO(line)))
                if len(row) >= 6:
                    p = row[5].strip()
                    total_parking += float(p) if p not in ("", "null") else 0.0
            except Exception:
                pass

    return {"staff": staff_name, "member": member_name,
            "month_year": month_year, "total_km": total_km,
            "total_parking": total_parking}


def build_excel(folder_path, output_path="mileage_summary.xlsx"):
    """Process all CSVs in a folder and write a formatted summary Excel file."""
    csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
    if not csv_files:
        print("No CSV files found in:", folder_path)
        return

    records = [r for fp in csv_files if (r := parse_csv_file(fp))]

    members = defaultdict(list)
    for r in records:
        members[r["member"]].append(r)

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    header_font = Font(bold=True, size=11, color="FFFFFF")
    member_fill = PatternFill("solid", fgColor="2E5D9E")
    label_fill  = PatternFill("solid", fgColor="D9E1F2")
    data_fill   = PatternFill("solid", fgColor="EEF2FB")
    thin   = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left",   vertical="center")

    row_cursor = 1

    for idx, (member_name, staff_list) in enumerate(sorted(members.items()), start=1):
        num_staff = len(staff_list)

        # Calculate total KMs across all staff for this member
        member_total_km = round(sum(s["total_km"] for s in staff_list), 2)

        # Member header (merged across all staff columns)
        ws.merge_cells(start_row=row_cursor, start_column=1,
                       end_row=row_cursor, end_column=1 + num_staff)
        cell = ws.cell(row=row_cursor, column=1,
                       value=f"Member {idx}: {member_name}  —  {member_total_km} km")
        cell.font = header_font; cell.fill = member_fill
        cell.alignment = center; cell.border = border
        row_cursor += 1

        # Staff names sub-header row
        ws.cell(row=row_cursor, column=1, value="").border = border
        for col, staff in enumerate(staff_list, start=2):
            c = ws.cell(row=row_cursor, column=col, value=staff["staff"])
            c.font = Font(bold=True, size=10)
            c.fill = label_fill; c.alignment = center; c.border = border
        row_cursor += 1

        # Total KMs row
        lbl = ws.cell(row=row_cursor, column=1, value="Total KMs")
        lbl.font = Font(bold=True); lbl.fill = label_fill
        lbl.alignment = left; lbl.border = border
        for col, staff in enumerate(staff_list, start=2):
            c = ws.cell(row=row_cursor, column=col, value=staff["total_km"])
            c.fill = data_fill; c.alignment = center; c.border = border
        row_cursor += 1

        # Parking Expense row
        lbl = ws.cell(row=row_cursor, column=1, value="Parking Expense ($)")
        lbl.font = Font(bold=True); lbl.fill = label_fill
        lbl.alignment = left; lbl.border = border
        for col, staff in enumerate(staff_list, start=2):
            c = ws.cell(row=row_cursor, column=col, value=staff["total_parking"])
            c.fill = data_fill; c.alignment = center; c.border = border
        row_cursor += 1

        row_cursor += 1  # blank spacer between members

    # Auto-size columns
    for col in ws.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 18)

    wb.save(output_path)
    print(f"Saved: {output_path} | Members: {len(members)} | Records: {len(records)}")


# -------------------------------------------------------
# UPDATE THESE TWO PATHS BEFORE RUNNING
# -------------------------------------------------------
FOLDER_PATH = r"/Users/yshmani/Desktop/mileage_entries"   # folder containing your CSVs
OUTPUT_FILE = r"/Users/yshmani/Desktop/mileage_entries/mileage_summary.xlsx"
build_excel(FOLDER_PATH, OUTPUT_FILE)