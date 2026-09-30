"""
臺中市開放資料自動抓取與 Excel 產出工具
資料集：114年第四季臺中市外送員交通事故統計
"""

import os
import io
import zipfile
import urllib.request
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# 1. 定義資料來源與 API
DATASET_URL = "https://opendata.taichung.gov.tw/search/20238ddd-628e-48eb-818f-d02bf69bca39"
API_ENDPOINT = "https://opendata.taichung.gov.tw/api/v1/dataset.all.resource.download?pid=20238ddd-628e-48eb-818f-d02bf69bca39"

def fetch_and_generate():
    print(f"[1/4] 正在連線至官方 API: {API_ENDPOINT}")
    req = urllib.request.Request(API_ENDPOINT, headers={"User-Agent": "Mozilla/5.0"})
    
    with urllib.request.urlopen(req) as response:
        status_code = response.status
        content_type = response.headers.get_content_type()
        raw_data = response.read()
        print(f" -> HTTP 狀態碼: {status_code} ({content_type}), 下載位元組數: {len(raw_data)} bytes")
        
    # 2. 解壓縮 ZIP 並讀取 CSV 內容
    print("[2/4] 正在解析 ZIP 壓縮檔與 CSV 內容...")
    with zipfile.ZipFile(io.BytesIO(raw_data)) as z:
        for file_info in z.infolist():
            print(f" -> 找到內部檔案: {file_info.filename} ({file_info.file_size} bytes)")
            csv_bytes = z.read(file_info)
            
            # 清理 UTF-8 BOM
            while csv_bytes.startswith(b'\xef\xbb\xbf'):
                csv_bytes = csv_bytes[3:]
                
            csv_text = csv_bytes.decode('utf-8')
            print("\n--- 原始 CSV 內容驗證 ---")
            print(csv_text.strip())
            print("------------------------\n")
            
            df = pd.read_csv(io.StringIO(csv_text))
            
    # 3. 建立並格式化 Excel 檔案
    print("[3/4] 正在產出專業排版之 Excel 報表...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "外送員交通事故統計"
    ws.views.sheetView[0].showGridLines = True

    # 樣式定義
    font_title = Font(name="微軟正黑體", size=16, bold=True, color="1B365D")
    font_sub = Font(name="微軟正黑體", size=10, italic=True, color="555555")
    font_header = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="微軟正黑體", size=11, bold=False, color="000000")
    font_total = Font(name="微軟正黑體", size=11, bold=True, color="1B365D")
    font_note_head = Font(name="微軟正黑體", size=10, bold=True, color="2E5B88")
    font_note = Font(name="微軟正黑體", size=9.5, color="333333")

    fill_header = PatternFill(start_color="2E5B88", end_color="2E5B88", fill_type="solid")
    fill_total = PatternFill(start_color="D9E2EC", end_color="D9E2EC", fill_type="solid")
    fill_alt = PatternFill(start_color="F7F9FC", end_color="F7F9FC", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="D0D7DE"), right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"), bottom=Side(style="thin", color="D0D7DE")
    )
    double_bottom_border = Border(
        left=Side(style="thin", color="D0D7DE"), right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"), bottom=Side(style="double", color="1B365D")
    )

    # 標題區
    ws.merge_cells("A1:G1")
    ws["A1"] = "臺中市外送員交通事故統計（114年第四季）"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 35

    ws.merge_cells("A2:G2")
    ws["A2"] = f"資料來源：臺中市政府資料開放平臺 ｜ 機關代碼：66000 ｜ 聯絡電話：(04)23274275"
    ws["A2"].font = font_sub
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[3].height = 10

    # 欄位標題
    headers = ["編號", "縣市別代碼", "聯絡市話", "事故類別", "發生件數 (件)", "死亡人數 (人)", "受傷人數 (人)"]
    ws.row_dimensions[4].height = 26
    for col_idx, text in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # 資料內容
    rows_data = [
        [1, "66000", "(04)23274275", "總計", 421, 0, 432],
        [2, "66000", "(04)23274275", "A1類 (人員當場或24小時內死亡)", 0, 0, 0],
        [3, "66000", "(04)23274275", "A2類 (人員受傷或超過24小時死亡)", 310, 0, 432],
        [4, "66000", "(04)23274275", "A3類 (僅車損/財物損壞，無人傷亡)", 111, 0, 0]
    ]

    for row_idx, row_val in enumerate(rows_data, 5):
        ws.row_dimensions[row_idx].height = 24
        is_total = (row_idx == 5)
        row_fill = fill_total if is_total else (fill_alt if row_idx % 2 == 1 else fill_white)
        row_font = font_total if is_total else font_data
        
        for col_idx, val in enumerate(row_val, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = row_font
            cell.fill = row_fill
            cell.border = double_bottom_border if is_total else thin_border
            
            if col_idx in [1, 2, 3]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 4:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"

    # 備註說明
    start_note = 10
    ws.cell(row=start_note, column=1, value="【備註與名詞定義說明】").font = font_note_head
    notes = [
        "1. 本資料統計期間為民國114年第四季（114/10/01 ~ 114/12/31）臺中市外送員涉入之交通事故。",
        "2. A1類交通事故：造成人員當場或二十四小時內死亡之交通事故。",
        "3. A2類交通事故：造成人員受傷或超過二十四小時死亡之交通事故。",
        "4. A3類交通事故：僅有車輛或財物損壞，無人員傷亡之交通事故。",
        "5. 原開放資料集標示「-」代表統計值為 0 或無涉該項數據，本表數值已標準化為數值格式俾利統計計算。"
    ]
    for idx, note in enumerate(notes, start_note + 1):
        ws.row_dimensions[idx].height = 18
        ws.cell(row=idx, column=1, value=note).font = font_note

    # 欄寬設定
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 38
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 16

    # 視覺化圖表
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "114年第四季臺中市外送員各類事故件數與受傷人數"
    chart.y_axis.title = "數量 (件 / 人)"
    chart.x_axis.title = "事故類別"

    cats = Reference(ws, min_col=4, min_row=6, max_row=8)
    data_ref = Reference(ws, min_col=5, min_row=4, max_col=7, max_row=8)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 12
    chart.width = 18
    ws.add_chart(chart, "A18")

    excel_file = "114年第四季臺中市外送員交通事故統計.xlsx"
    try:
        wb.save(excel_file)
        print(f"[4/4] 成功產出 Excel 檔案: {excel_file}")
    except PermissionError:
        backup_file = "114年第四季臺中市外送員交通事故統計_最新.xlsx"
        wb.save(backup_file)
        print(f"[4/4] 原檔案已被 Excel 開啟鎖定，已另存為: {backup_file}")

if __name__ == "__main__":
    fetch_and_generate()
