"""
main.py - GML CSV 點位處理工具主程式

用法：
  python src/main.py
  （請在 gml-csv-canvas-tool/ 目錄下執行，或使用 run.bat）
"""

import os
import sys

# 確保 src/ 可 import 同層模組
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from csv_loader import load_all_csvs
from id_parser import parse_all_ids
from anomaly_detector import detect_anomalies
from canvas_generator import generate_canvas
from excel_exporter import export_all

# ---- 路徑設定 ----
BASE_DIR   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_DIR  = os.path.join(BASE_DIR, 'input_csv')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
CONFIG_PATH = os.path.join(BASE_DIR, 'config', 'rules.yaml')


def main():
    print('=' * 55)
    print('  GML CSV 點位處理工具  v2.0')
    print('=' * 55)
    print()

    # 建立輸出目錄
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── 步驟 1：讀取 CSV ──────────────────────────────────────
    print('【步驟 1】讀取 CSV 檔案')
    print(f'  輸入目錄: {INPUT_DIR}')

    df, file_meta, load_errors = load_all_csvs(INPUT_DIR, CONFIG_PATH)

    if load_errors:
        print(f'\n  [警告] {len(load_errors)} 個檔案讀取失敗：')
        for err in load_errors:
            print(f'    - {err["file"]}：{err["error"]}')

    if df.empty:
        print('\n[錯誤] 未讀取到任何資料。')
        print('  請確認 input_csv/ 資料夾內有 CSV 檔案。')
        sys.exit(1)

    print(f'\n  共載入 {len(df)} 筆資料，來自 {len(file_meta)} 個檔案。\n')

    # ── 步驟 2：解析識別碼 ───────────────────────────────────
    print('【步驟 2】解析識別碼')
    df = parse_all_ids(df)

    pipeline_count = df['pipeline_id'].nunique()
    print(f'  共識別出 {pipeline_count} 條管線識別碼。\n')

    # ── 步驟 3：異常偵測 ────────────────────────────────────
    print('【步驟 3】異常偵測')
    df, anomaly_list, duplicate_list = detect_anomalies(df, CONFIG_PATH, file_meta)

    # 建立異常點集合（用於畫布標示）
    anomaly_rows = set()
    for item in anomaly_list:
        src = item.get('來源CSV')
        row_no = item.get('原始列號')
        if src and row_no and row_no != '（整份檔案）':
            try:
                anomaly_rows.add((src, int(row_no)))
            except (ValueError, TypeError):
                pass

    print(f'  偵測到 {len(anomaly_list)} 筆異常。')
    print(f'  偵測到 {len(duplicate_list)} 筆重複量測記錄。\n')

    # ── 步驟 4：產生點位畫布 ─────────────────────────────────
    print('【步驟 4】產生互動式 HTML 點位畫布')
    canvas_path = os.path.join(OUTPUT_DIR, '點位畫布.html')
    generate_canvas(df, canvas_path, anomaly_rows=anomaly_rows)

    # ── 步驟 5：輸出 Excel ───────────────────────────────────
    print('\n【步驟 5】輸出 Excel 成果檔案')
    export_all(df, anomaly_list, duplicate_list, OUTPUT_DIR)

    # ── 完成摘要 ────────────────────────────────────────────
    print()
    print('=' * 55)
    print('  完成！成果檔案位於：')
    print(f'  {OUTPUT_DIR}')
    print()
    print('  ├─ 點位畫布.html         ← 用瀏覽器開啟，互動瀏覽點位')
    print('  ├─ 解析後點位表.xlsx     ← 所有點位的解析結果')
    print('  ├─ 異常清單.xlsx         ← 需人工確認的異常')
    print('  ├─ 重複量測比對表.xlsx   ← 重複量測點的座標比對')
    print('  ├─ 人工判讀紀錄範本.xlsx ← [v2] 逐點填寫判讀狀態（含下拉選單）')
    print('  └─ 修正紀錄範本.xlsx     ← [v2] 填寫修正決策（含下拉選單）')
    print()
    print('  注意：異常標示僅供參考，所有修改決策請由人工確認。')
    print('=' * 55)


if __name__ == '__main__':
    main()
