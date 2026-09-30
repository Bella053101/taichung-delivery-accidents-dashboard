import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
import pandas as pd
import os

# Dataset information
dataset_title = "臺中市外送員交通事故統計（114年第四季）"
source_url = "https://opendata.taichung.gov.tw/search/20238ddd-628e-48eb-818f-d02bf69bca39"
dept = "臺中市政府警察局"
phone = "(04)23274275"
city_code = "66000"

# Data table:
# 編號, 縣市別代碼, 市話, 類別, 發生件數, 死亡人數, 受傷人數
# 1, 66000, (04)23274275, 總計, 421, -, 432
# 2, 66000, (04)23274275, A1類, -, -, -
# 3, 66000, (04)23274275, A2類, 310, -, 432
# 4, 66000, (04)23274275, A3類, 111, -, -

# Create Workbook
wb = openpyxl.Workbook()

# Sheet 1: 外送員交通事故統計表
ws = wb.active
ws.title = "外送員交通事故統計"
ws.views.sheetView[0].showGridLines = True

# Colors & Styles
NAVY = "1B365D"
LIGHT_BLUE = "E8EEF5"
HEADER_FILL = "2E5B88"
ACCENT_ROW = "F7F9FC"
BORDER_COLOR = "D0D7DE"

font_title = Font(name="微軟正黑體", size=16, bold=True, color="1B365D")
font_sub = Font(name="微軟正黑體", size=10, italic=True, color="555555")
font_header = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
font_data = Font(name="微軟正黑體", size=11, bold=False, color="000000")
font_total = Font(name="微軟正黑體", size=11, bold=True, color="1B365D")
font_note_head = Font(name="微軟正黑體", size=10, bold=True, color="2E5B88")
font_note = Font(name="微軟正黑體", size=9.5, color="333333")

fill_header = PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")
fill_total = PatternFill(start_color="D9E2EC", end_color="D9E2EC", fill_type="solid")
fill_alt = PatternFill(start_color=ACCENT_ROW, end_color=ACCENT_ROW, fill_type="solid")
fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color=BORDER_COLOR),
    right=Side(style="thin", color=BORDER_COLOR),
    top=Side(style="thin", color=BORDER_COLOR),
    bottom=Side(style="thin", color=BORDER_COLOR)
)

double_bottom_border = Border(
    left=Side(style="thin", color=BORDER_COLOR),
    right=Side(style="thin", color=BORDER_COLOR),
    top=Side(style="thin", color=BORDER_COLOR),
    bottom=Side(style="double", color="1B365D")
)

align_center = Alignment(horizontal="center", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")

# Title Block
ws.merge_cells("A1:G1")
ws["A1"] = dataset_title
ws["A1"].font = font_title
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 35

ws.merge_cells("A2:G2")
ws["A2"] = f"資料來源：臺中市政府資料開放平臺 ({source_url}) ｜ 權管機關：{dept}"
ws["A2"].font = font_sub
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[2].height = 20

# Empty Row
ws.row_dimensions[3].height = 10

# Table Headers
headers = ["編號", "縣市別代碼", "聯絡市話", "事故類別", "發生件數 (件)", "死亡人數 (人)", "受傷人數 (人)"]
ws.row_dimensions[4].height = 26

for col_idx, text in enumerate(headers, 1):
    cell = ws.cell(row=4, column=col_idx, value=text)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = thin_border

# Data Rows
data = [
    [1, "66000", "(04)23274275", "總計", 421, 0, 432],
    [2, "66000", "(04)23274275", "A1類 (人員當場或24小時內死亡)", 0, 0, 0],
    [3, "66000", "(04)23274275", "A2類 (人員受傷或超過24小時死亡)", 310, 0, 432],
    [4, "66000", "(04)23274275", "A3類 (僅車損/財物損壞，無人傷亡)", 111, 0, 0]
]

for row_idx, row_data in enumerate(data, 5):
    ws.row_dimensions[row_idx].height = 24
    is_total = (row_idx == 5)
    row_fill = fill_total if is_total else (fill_alt if row_idx % 2 == 1 else fill_white)
    row_font = font_total if is_total else font_data
    
    for col_idx, val in enumerate(row_data, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.font = row_font
        cell.fill = row_fill
        cell.border = double_bottom_border if is_total else thin_border
        
        # Alignment & Number formatting
        if col_idx in [1, 2, 3]:
            cell.alignment = align_center
        elif col_idx == 4:
            cell.alignment = align_left
        else:
            cell.alignment = align_right
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0"

# Note & Definition Section
start_note_row = 10
ws.row_dimensions[start_note_row].height = 20
ws.cell(row=start_note_row, column=1, value="【備註與名詞定義說明】").font = font_note_head

notes = [
    "1. 本資料統計期間為民國114年第四季（114/10/01 ~ 114/12/31）臺中市外送員涉入之交通事故。",
    "2. A1類交通事故：造成人員當場或二十四小時內死亡之交通事故。",
    "3. A2類交通事故：造成人員受傷或超過二十四小時死亡之交通事故。",
    "4. A3類交通事故：僅有車輛或財物損壞，無人員傷亡之交通事故。",
    "5. 原開放資料集標示「-」代表統計值為 0 或無涉該項數據，本表數值已標準化為數值格式俾利統計計算。"
]

for idx, note in enumerate(notes, start_note_row + 1):
    ws.row_dimensions[idx].height = 18
    c = ws.cell(row=idx, column=1, value=note)
    c.font = font_note

# Auto-adjust column widths
for col in ws.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        if cell.row > 8:  # ignore note rows for width calculation
            continue
        if cell.value:
            val_str = str(cell.value)
            # Estimate Chinese character length
            length = sum(2 if ord(char) > 127 else 1 for char in val_str)
            if length > max_len:
                max_len = length
    ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

# Specific adjustments
ws.column_dimensions["A"].width = 10
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 38
ws.column_dimensions["E"].width = 16
ws.column_dimensions["F"].width = 16
ws.column_dimensions["G"].width = 16

# Add Bar Chart
chart = BarChart()
chart.type = "col"
chart.style = 10
chart.title = "114年第四季臺中市外送員各類事故件數與受傷人數"
chart.y_axis.title = "數量 (件 / 人)"
chart.x_axis.title = "事故類別"

# Data for A1, A2, A3 (rows 6 to 8, cols 5 to 7)
data_ref = Reference(ws, min_col=5, min_row=5, max_col=7, max_row=8) # include headers on row 5 or 4?
# We want categories from col 4 (A1, A2, A3) -> rows 6 to 8
cats = Reference(ws, min_col=4, min_row=6, max_row=8)
data_ref = Reference(ws, min_col=5, min_row=4, max_col=7, max_row=8)

chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats)
chart.height = 12
chart.width = 18
ws.add_chart(chart, "A18")

# Save files
excel_path = "114年第四季臺中市外送員交通事故統計.xlsx"
csv_path = "114年第四季臺中市外送員交通事故統計.csv"

wb.save(excel_path)
print(f"Excel file saved successfully to: {excel_path}")

# Also output standard UTF-8-BOM CSV
df_out = pd.DataFrame([
    {"編號": 1, "縣市別代碼": "66000", "市話": "(04)23274275", "類別": "總計", "發生件數": 421, "死亡人數": "-", "受傷人數": 432},
    {"編號": 2, "縣市別代碼": "66000", "市話": "(04)23274275", "類別": "A1類", "發生件數": "-", "死亡人數": "-", "受傷人數": "-"},
    {"編號": 3, "縣市別代碼": "66000", "市話": "(04)23274275", "類別": "A2類", "發生件數": 310, "死亡人數": "-", "受傷人數": 432},
    {"編號": 4, "縣市別代碼": "66000", "市話": "(04)23274275", "類別": "A3類", "發生件數": 111, "死亡人數": "-", "受傷人數": "-"}
])
df_out.to_csv(csv_path, index=False, encoding="utf-8-sig")
print(f"CSV file saved to: {csv_path}")
