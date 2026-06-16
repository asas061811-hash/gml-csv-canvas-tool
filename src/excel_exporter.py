# -*- coding: utf-8 -*-
"""
excel_exporter.py
-------------------
使用 openpyxl 將處理結果輸出為 Excel 檔案：
    - output/解析後點位表.xlsx
    - output/異常清單.xlsx
    - output/重複量測比對表.xlsx
    - output/人工判讀紀錄範本.xlsx  （第二版）
    - output/修正紀錄範本.xlsx      （第二版）
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from anomaly_detector import ANOMALY_COLUMNS, DUPLICATE_COLUMNS


PARSED_COLUMNS = [
    "原始識別碼",
    "管線識別碼",
    "點號",
    "類別碼",
    "X",
    "Y",
    "Z",
    "X_原始值",
    "Y_原始值",
    "Z_原始值",
    "測量日期",
    "來源CSV",
    "原始列號",
]

HEADER_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
HEADER_FONT = Font(bold=True)


def _write_sheet(ws, columns, rows):
    ws.append(columns)
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    # 簡單調整欄寬
    for idx, col in enumerate(columns, start=1):
        max_len = len(str(col))
        for row in rows:
            val = row.get(col, "")
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 10), 50)

    ws.freeze_panes = "A2"


def export_parsed_records(records, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "解析後點位表"

    rows = []
    for rec in records:
        rows.append({
            "原始識別碼": rec.get("原始識別碼", ""),
            "管線識別碼": rec.get("管線識別碼", ""),
            "點號": rec.get("點號"),
            "類別碼": rec.get("類別碼", ""),
            "X": rec.get("X"),
            "Y": rec.get("Y"),
            "Z": rec.get("Z"),
            "X_原始值": rec.get("X_raw", ""),
            "Y_原始值": rec.get("Y_raw", ""),
            "Z_原始值": rec.get("Z_raw", ""),
            "測量日期": rec.get("測量日期", ""),
            "來源CSV": rec.get("來源CSV", ""),
            "原始列號": rec.get("原始列號", ""),
        })

    _write_sheet(ws, PARSED_COLUMNS, rows)
    wb.save(output_path)


def export_anomalies(anomalies, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "異常清單"
    _write_sheet(ws, ANOMALY_COLUMNS, anomalies)
    wb.save(output_path)


def export_duplicates(duplicates, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "重複量測比對表"
    _write_sheet(ws, DUPLICATE_COLUMNS, duplicates)
    wb.save(output_path)


# ── 第二版：下拉選單輔助 ────────────────────────────────────────────────────

def _add_dropdown(ws, col_letter, start_row, end_row, choices):
    """在指定欄範圍加入下拉選單（DataValidation list）。"""
    formula = '"' + ",".join(choices) + '"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showDropDown=False)
    dv.sqref = f"{col_letter}{start_row}:{col_letter}{end_row}"
    ws.add_data_validation(dv)


INPUT_FILL = PatternFill(start_color="FFFCE5", end_color="FFFCE5", fill_type="solid")   # 淺黃：人工填寫欄
LOCK_FILL  = PatternFill(start_color="EEF2F7", end_color="EEF2F7", fill_type="solid")   # 淺灰：資料來源欄


def _style_header_row(ws, data_cols, input_cols):
    """對標題列套用顏色：資料欄灰底、人工填寫欄黃底。"""
    for idx, col in enumerate(data_cols + input_cols, start=1):
        cell = ws.cell(row=1, column=idx)
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
        if col in input_cols:
            cell.fill = INPUT_FILL
        else:
            cell.fill = LOCK_FILL


# ── 人工判讀紀錄範本 ─────────────────────────────────────────────────────────

REVIEW_DATA_COLS = [
    "來源CSV", "原始列號", "原始識別碼", "管線識別碼", "點號",
    "X", "Y", "Z", "測量日期",
]
REVIEW_INPUT_COLS = [
    "判讀狀態", "問題類型", "判讀說明", "建議回覆施工廠商內容",
    "確認人員", "確認日期",
]
REVIEW_COLUMNS = REVIEW_DATA_COLS + REVIEW_INPUT_COLS

REVIEW_STATUS_CHOICES = ["正常", "待確認", "需廠商補測", "排除", "採用"]
REVIEW_ISSUE_CHOICES  = ["點位偏移", "點位缺漏", "點位重複", "座標異常",
                          "日期不一致", "點序錯誤", "其他"]

MAX_DROPDOWN_ROWS = 5000   # 下拉選單預先覆蓋到第幾列（不含標題）


def export_review_template(records, output_path):
    """
    輸出人工判讀紀錄範本。
    - 前 9 欄由 records 帶入資料（唯讀底色）
    - 後 6 欄留空供人工填寫（黃色底色），含判讀狀態與問題類型下拉選單
    - 不修改原始 CSV
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "人工判讀紀錄"

    # 標題列
    ws.append(REVIEW_COLUMNS)
    _style_header_row(ws, REVIEW_DATA_COLS, REVIEW_INPUT_COLS)

    # 資料列
    for rec in records:
        ws.append([
            rec.get("來源CSV", ""),
            rec.get("原始列號", ""),
            rec.get("原始識別碼", ""),
            rec.get("管線識別碼", ""),
            rec.get("點號", ""),
            rec.get("X", ""),
            rec.get("Y", ""),
            rec.get("Z", ""),
            rec.get("測量日期", ""),
            "",  # 判讀狀態
            "",  # 問題類型
            "",  # 判讀說明
            "",  # 建議回覆施工廠商內容
            "",  # 確認人員
            "",  # 確認日期
        ])

    # 人工填寫欄背景色
    input_start_col = len(REVIEW_DATA_COLS) + 1
    data_rows = len(records)
    for r in range(2, data_rows + 2):
        for c in range(input_start_col, len(REVIEW_COLUMNS) + 1):
            ws.cell(row=r, column=c).fill = INPUT_FILL

    # 下拉選單（涵蓋資料列 + 預留空白列）
    dropdown_end = max(data_rows + 1, MAX_DROPDOWN_ROWS)
    status_col = get_column_letter(REVIEW_COLUMNS.index("判讀狀態") + 1)
    issue_col  = get_column_letter(REVIEW_COLUMNS.index("問題類型") + 1)
    _add_dropdown(ws, status_col, 2, dropdown_end, REVIEW_STATUS_CHOICES)
    _add_dropdown(ws, issue_col,  2, dropdown_end, REVIEW_ISSUE_CHOICES)

    # 欄寬
    col_widths = {
        "來源CSV": 22, "原始列號": 10, "原始識別碼": 28, "管線識別碼": 28,
        "點號": 8, "X": 16, "Y": 16, "Z": 12, "測量日期": 14,
        "判讀狀態": 14, "問題類型": 14, "判讀說明": 30,
        "建議回覆施工廠商內容": 40, "確認人員": 12, "確認日期": 14,
    }
    for idx, col in enumerate(REVIEW_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = col_widths.get(col, 16)

    ws.freeze_panes = "J2"
    wb.save(output_path)


# ── 修正紀錄範本 ─────────────────────────────────────────────────────────────

CORRECTION_COLUMNS = [
    "動作類型", "原始識別碼", "管線識別碼", "點號",
    "原X", "原Y", "原Z",
    "修正後X", "修正後Y", "修正後Z",
    "修正原因", "是否納入後續GML",
    "確認人員", "確認日期",
]

ACTION_CHOICES = ["新增", "刪除", "採用", "排除", "修正"]


def export_correction_template(output_path):
    """
    輸出修正紀錄範本（空白範本，全部欄位由人工填寫）。
    - 動作類型欄含下拉選單
    - 不修改原始 CSV
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "修正紀錄"

    # 標題列（全欄黃底，表示人工填寫）
    ws.append(CORRECTION_COLUMNS)
    for idx in range(1, len(CORRECTION_COLUMNS) + 1):
        cell = ws.cell(row=1, column=idx)
        cell.font = HEADER_FONT
        cell.fill = INPUT_FILL
        cell.alignment = Alignment(horizontal="center")

    # 動作類型下拉選單
    action_col = get_column_letter(CORRECTION_COLUMNS.index("動作類型") + 1)
    _add_dropdown(ws, action_col, 2, MAX_DROPDOWN_ROWS, ACTION_CHOICES)

    # 欄寬
    col_widths = {
        "動作類型": 12, "原始識別碼": 28, "管線識別碼": 28, "點號": 8,
        "原X": 16, "原Y": 16, "原Z": 12,
        "修正後X": 16, "修正後Y": 16, "修正後Z": 12,
        "修正原因": 30, "是否納入後續GML": 18,
        "確認人員": 12, "確認日期": 14,
    }
    for idx, col in enumerate(CORRECTION_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = col_widths.get(col, 16)

    ws.freeze_panes = "A2"
    wb.save(output_path)
