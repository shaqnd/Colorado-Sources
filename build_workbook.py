import csv
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter

# ── colour palette ──────────────────────────────────────────────────────────
HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")   # dark blue
HEADER_FONT   = Font(bold=True, color="FFFFFF", size=11)

# municipality row colours (alternating pairs for readability)
MUNI_COLOURS = {
    "Adams County (Unincorporated)": "D6E4F0",
    "Brighton":                       "D5E8D4",
    "Commerce City":                  "FFF2CC",
    "Federal Heights":                "FCE4D6",
    "Northglenn":                     "E1D5E7",
    "Thornton":                       "DAE8FC",
    "Westminster":                    "F8CECC",
    "Aurora":                         "D5E8D4",
    "Bennett":                        "FFF2CC",
    "Lochbuie":                       "FCE4D6",
    "Arvada":                         "E1D5E7",
    "REGIONAL / ALL MUNICIPALITIES":  "EEEEEE",
}

LINK_FONT = Font(color="1155CC", underline="single", size=10)
DEFAULT_FONT = Font(size=10)
BOLD_FONT    = Font(bold=True, size=10)

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

wb = Workbook()
ws = wb.active
ws.title = "All Resources"

# ── read CSV ────────────────────────────────────────────────────────────────
with open("/home/user/Colorado-Sources/Adams_County_CO_Comprehensive_RE_Resources.csv",
          encoding="utf-8") as f:
    rows = list(csv.reader(f))

headers = rows[0]   # Municipality, Category, Resource Name, URL, Notes
data    = rows[1:]

# ── write header row ─────────────────────────────────────────────────────────
for col_idx, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx, value=h)
    cell.fill      = PatternFill("solid", fgColor="1F4E79")
    cell.font      = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = thin_border()

ws.row_dimensions[1].height = 28

# ── write data rows ──────────────────────────────────────────────────────────
for row_idx, row in enumerate(data, 2):
    muni = row[0] if len(row) > 0 else ""
    bg   = MUNI_COLOURS.get(muni, "FFFFFF")
    fill = PatternFill("solid", fgColor=bg)

    for col_idx, value in enumerate(row, 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.border    = thin_border()
        cell.fill      = fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)

        # col 4 = URL → hyperlink
        if col_idx == 4 and value.startswith("http"):
            cell.value     = value
            cell.hyperlink = value
            cell.font      = LINK_FONT
        elif col_idx == 1:
            cell.value = value
            cell.font  = BOLD_FONT
        else:
            cell.value = value
            cell.font  = DEFAULT_FONT

    ws.row_dimensions[row_idx].height = 18

# ── column widths ─────────────────────────────────────────────────────────────
col_widths = {1: 34, 2: 26, 3: 36, 4: 62, 5: 55}
for col_idx, width in col_widths.items():
    ws.column_dimensions[get_column_letter(col_idx)].width = width

# ── auto-filter on all columns ────────────────────────────────────────────────
last_col = get_column_letter(len(headers))
ws.auto_filter.ref = f"A1:{last_col}{len(data)+1}"

# ── freeze header row ─────────────────────────────────────────────────────────
ws.freeze_panes = "A2"

# ── second sheet: Category Index ─────────────────────────────────────────────
ws2 = wb.create_sheet("Category Index")
ws2.freeze_panes = "A2"

cat_headers = ["Category", "Municipality Count", "Resource Count"]
for col_idx, h in enumerate(cat_headers, 1):
    cell = ws2.cell(row=1, column=col_idx, value=h)
    cell.fill      = PatternFill("solid", fgColor="1F4E79")
    cell.font      = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = thin_border()
ws2.row_dimensions[1].height = 28

from collections import defaultdict
cat_map = defaultdict(lambda: {"munis": set(), "count": 0})
for row in data:
    cat  = row[1] if len(row) > 1 else ""
    muni = row[0] if len(row) > 0 else ""
    cat_map[cat]["munis"].add(muni)
    cat_map[cat]["count"] += 1

alt_fills = [PatternFill("solid", fgColor="EBF3FB"),
             PatternFill("solid", fgColor="FFFFFF")]
for i, (cat, info) in enumerate(sorted(cat_map.items()), 2):
    fill = alt_fills[i % 2]
    for col_idx, val in enumerate([cat, len(info["munis"]), info["count"]], 1):
        cell = ws2.cell(row=i, column=col_idx, value=val)
        cell.fill      = fill
        cell.font      = DEFAULT_FONT
        cell.alignment = Alignment(vertical="center")
        cell.border    = thin_border()
    ws2.row_dimensions[i].height = 18

ws2.column_dimensions["A"].width = 38
ws2.column_dimensions["B"].width = 22
ws2.column_dimensions["C"].width = 18
ws2.auto_filter.ref = f"A1:C{len(cat_map)+1}"

# ── third sheet: Municipality Index ──────────────────────────────────────────
ws3 = wb.create_sheet("Municipality Index")
ws3.freeze_panes = "A2"

muni_headers = ["Municipality", "Type / County", "Resource Count", "Categories Covered"]
for col_idx, h in enumerate(muni_headers, 1):
    cell = ws3.cell(row=1, column=col_idx, value=h)
    cell.fill      = PatternFill("solid", fgColor="1F4E79")
    cell.font      = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = thin_border()
ws3.row_dimensions[1].height = 28

muni_types = {
    "Adams County (Unincorporated)": "County / Adams County",
    "Brighton":                       "City — County Seat / Adams County",
    "Commerce City":                  "City / Adams County",
    "Federal Heights":                "City / Adams County",
    "Northglenn":                     "City / Adams County",
    "Thornton":                       "City / Adams County",
    "Westminster":                    "City / Adams + Jefferson Counties",
    "Aurora":                         "City / Adams + Arapahoe + Douglas Counties",
    "Bennett":                        "Town / Adams County",
    "Lochbuie":                       "Town / Adams + Weld Counties",
    "Arvada":                         "City / Jefferson + small Adams County",
    "REGIONAL / ALL MUNICIPALITIES":  "Regional / State Resources",
}

muni_map = defaultdict(lambda: {"cats": set(), "count": 0})
for row in data:
    muni = row[0] if len(row) > 0 else ""
    cat  = row[1] if len(row) > 1 else ""
    muni_map[muni]["cats"].add(cat)
    muni_map[muni]["count"] += 1

order = list(muni_types.keys())
for i, muni in enumerate(order, 2):
    info  = muni_map.get(muni, {"cats": set(), "count": 0})
    mtype = muni_types.get(muni, "")
    fill  = PatternFill("solid", fgColor=MUNI_COLOURS.get(muni, "FFFFFF"))
    cats  = ", ".join(sorted(info["cats"]))
    for col_idx, val in enumerate([muni, mtype, info["count"], cats], 1):
        cell = ws3.cell(row=i, column=col_idx, value=val)
        cell.fill      = fill
        cell.font      = BOLD_FONT if col_idx == 1 else DEFAULT_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border    = thin_border()
    ws3.row_dimensions[i].height = 30

ws3.column_dimensions["A"].width = 34
ws3.column_dimensions["B"].width = 32
ws3.column_dimensions["C"].width = 18
ws3.column_dimensions["D"].width = 70
ws3.auto_filter.ref = f"A1:D{len(order)+1}"

# ── save ──────────────────────────────────────────────────────────────────────
out = "/home/user/Colorado-Sources/Adams_County_CO_RE_Resources.xlsx"
wb.save(out)
print(f"Saved: {out}")
print(f"Rows in All Resources sheet: {len(data)}")
