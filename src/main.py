# -*- coding: utf-8 -*-
"""
main.py
-------
GML CSV 點位畫布工具 - 主程式

執行方式（於專案根目錄）：
    python src/main.py

會依序：
    1. 讀取 config/rules.yaml
    2. 讀取 input_csv 內所有 CSV（不修改原始檔案）
    3. 解析識別碼，拆出管線識別碼與點號
    4. 進行各項異常檢查與重複點位比對
    5. 輸出：
        output/點位畫布.html
        output/解析後點位表.xlsx
        output/異常清單.xlsx
        output/重複量測比對表.xlsx
"""

import os
import sys

import yaml

# 確保可以匯入同目錄下的模組
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import csv_loader
import id_parser
import anomaly_detector
import canvas_generator
import excel_exporter


def get_project_root():
    # src/ 的上一層即為專案根目錄
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    root = get_project_root()
    input_dir = os.path.join(root, "input_csv")
    output_dir = os.path.join(root, "output")
    rules_path = os.path.join(root, "config", "rules.yaml")

    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    print("=" * 60)
    print("GML CSV 點位畫布工具")
    print("=" * 60)

    # 1. 讀取設定檔
    if not os.path.exists(rules_path):
        print(f"[錯誤] 找不到設定檔: {rules_path}")
        sys.exit(1)

    with open(rules_path, "r", encoding="utf-8") as f:
        rules = yaml.safe_load(f)

    # 2. 讀取所有 CSV
    print(f"\n[1/5] 讀取 CSV 檔案：{input_dir}")
    records, warnings = csv_loader.load_all_csv(input_dir, rules)

    if warnings:
        print("  -- 讀取警告 --")
        for w in warnings:
            print(f"  * {w}")

    if not records:
        print("\n[結束] 沒有讀取到任何資料列，請確認 input_csv 內是否放置了 CSV 檔案。")
        return

    print(f"  共讀取 {len(records)} 筆資料列。")

    # 3. 解析識別碼
    print("\n[2/5] 解析識別碼（管線識別碼 / 點號）")
    records = id_parser.parse_records(records, rules)

    # 4. 異常檢查與重複點位比對
    print("\n[3/5] 進行異常檢查與重複點位比對")
    records, anomalies, duplicates = anomaly_detector.process_records(records, rules)
    print(f"  發現異常項目：{len(anomalies)} 筆")
    print(f"  重複量測比對項目：{len(duplicates)} 筆")

    # 5. 輸出 HTML 畫布
    print("\n[4/5] 產生互動式 HTML 畫布")
    canvas_path = os.path.join(output_dir, "點位畫布.html")
    canvas_generator.generate_canvas(records, canvas_path)
    print(f"  已輸出: {canvas_path}")

    # 6. 輸出 Excel（第一版）
    print("\n[5/5] 輸出 Excel 報表（第一版）")
    parsed_path = os.path.join(output_dir, "解析後點位表.xlsx")
    anomalies_path = os.path.join(output_dir, "異常清單.xlsx")
    duplicates_path = os.path.join(output_dir, "重複量測比對表.xlsx")

    excel_exporter.export_parsed_records(records, parsed_path)
    print(f"  已輸出: {parsed_path}")

    excel_exporter.export_anomalies(anomalies, anomalies_path)
    print(f"  已輸出: {anomalies_path}")

    excel_exporter.export_duplicates(duplicates, duplicates_path)
    print(f"  已輸出: {duplicates_path}")

    # 7. 輸出 Excel（第二版：人工判讀與修正範本）
    print("\n[+] 輸出 Excel 判讀範本（第二版）")
    review_path = os.path.join(output_dir, "人工判讀紀錄範本.xlsx")
    correction_path = os.path.join(output_dir, "修正紀錄範本.xlsx")

    excel_exporter.export_review_template(records, review_path)
    print(f"  已輸出: {review_path}")

    excel_exporter.export_correction_template(correction_path)
    print(f"  已輸出: {correction_path}")

    print("\n完成！請至 output/ 資料夾查看結果。")
    print("提醒：所有重複點位、座標差異與異常項目皆需人工確認，本工具不會自動修改或刪除任何資料。")


if __name__ == "__main__":
    main()
