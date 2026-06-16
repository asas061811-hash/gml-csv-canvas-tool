# -*- coding: utf-8 -*-
"""
anomaly_detector.py
--------------------
針對解析後的點位資料進行各項檢查，產出：
    1. 補強後的點位資料（含轉換後的 X/Y/Z 數值與格式是否正確的標記）
    2. 異常清單（list[dict]）
    3. 重複量測比對表（list[dict]）

本模組不會刪除、合併或修改任何資料，僅標記與彙整供人工判讀。
"""

import math
from collections import defaultdict


ANOMALY_COLUMNS = [
    "異常類型",
    "異常說明",
    "原始識別碼",
    "管線識別碼",
    "點號",
    "X",
    "Y",
    "Z",
    "測量日期",
    "來源CSV",
    "原始列號",
    "建議處理方式",
]

DUPLICATE_COLUMNS = [
    "來源CSV",
    "原始列號",
    "原始識別碼",
    "管線識別碼",
    "點號",
    "測量日期",
    "X",
    "Y",
    "Z",
    "與比對點XY差距",
    "與比對點Z差距",
    "判定結果",
    "建議處理方式",
]


def _to_float(raw):
    """
    將字串轉換為浮點數。

    回傳 (value, status)
        status = "empty"   -> 原始字串為空
        status = "ok"      -> 轉換成功
        status = "invalid" -> 非空但無法轉換為數字
    """
    if raw is None:
        return None, "empty"
    s = str(raw).strip()
    if s == "":
        return None, "empty"
    # 允許千分位逗號
    s2 = s.replace(",", "")
    try:
        return float(s2), "ok"
    except ValueError:
        return None, "invalid"


def _new_anomaly(a_type, desc, rec, x_val, y_val, z_val, suggestion):
    return {
        "異常類型": a_type,
        "異常說明": desc,
        "原始識別碼": rec.get("原始識別碼", ""),
        "管線識別碼": rec.get("管線識別碼", ""),
        "點號": rec.get("點號"),
        "X": x_val,
        "Y": y_val,
        "Z": z_val,
        "測量日期": rec.get("測量日期", ""),
        "來源CSV": rec.get("來源CSV", ""),
        "原始列號": rec.get("原始列號", ""),
        "建議處理方式": suggestion,
    }


def process_records(records, rules):
    """
    處理所有資料列。

    回傳:
        records   : 補強後的點位資料（加入 X, Y, Z 數值、X_格式正確 等欄位）
        anomalies : 異常清單 list[dict]
        duplicates: 重複量測比對表 list[dict]
    """
    anomalies = []

    thresholds = rules.get("duplicate_thresholds", {})
    xy_tol = float(thresholds.get("xy_tolerance_m", 0.30))
    z_tol = float(thresholds.get("z_tolerance_m", 0.10))

    # ---------- 1. 欄位缺漏 / 座標格式 / 座標空值 / 日期空值 ----------
    for rec in records:
        missing_fields = rec.get("缺漏欄位", [])
        field_name_map = {
            "id": "識別碼",
            "category": "類別碼",
            "x": "X",
            "y": "Y",
            "z": "Z",
            "date": "測量日期",
        }
        for mf in missing_fields:
            anomalies.append(_new_anomaly(
                "欄位缺漏",
                f"來源 CSV 找不到對應「{field_name_map.get(mf, mf)}」欄位，請確認原始檔案或於 config/rules.yaml 新增欄位對應。",
                rec, None, None, None,
                "請人工確認來源 CSV 是否確實缺少此欄位，或於 config/rules.yaml 新增欄位別名後重新執行。",
            ))

        # X / Y 轉換與格式檢查
        x_val, x_status = _to_float(rec.get("X_raw", ""))
        y_val, y_status = _to_float(rec.get("Y_raw", ""))
        z_val, z_status = _to_float(rec.get("Z_raw", ""))

        rec["X"] = x_val
        rec["Y"] = y_val
        rec["Z"] = z_val

        if x_status == "invalid":
            anomalies.append(_new_anomaly(
                "X 座標格式異常",
                f"X 欄位原始值「{rec.get('X_raw')}」無法轉換為數字。",
                rec, rec.get("X_raw"), rec.get("Y_raw"), rec.get("Z_raw"),
                "請人工核對原始 CSV 中的 X 欄位內容是否正確。",
            ))
        if y_status == "invalid":
            anomalies.append(_new_anomaly(
                "Y 座標格式異常",
                f"Y 欄位原始值「{rec.get('Y_raw')}」無法轉換為數字。",
                rec, rec.get("X_raw"), rec.get("Y_raw"), rec.get("Z_raw"),
                "請人工核對原始 CSV 中的 Y 欄位內容是否正確。",
            ))

        if x_status == "empty" or y_status == "empty":
            anomalies.append(_new_anomaly(
                "座標空值",
                "X 或 Y 座標為空值，無法繪製於畫布上。",
                rec, x_val, y_val, z_val,
                "請人工確認該點位是否漏量測，並補齊座標資料。",
            ))

        if rec.get("測量日期", "") == "":
            anomalies.append(_new_anomaly(
                "日期空值",
                "測量日期為空值。",
                rec, x_val, y_val, z_val,
                "請人工確認該點位的測量日期。",
            ))

        # 點號無法解析（識別碼結尾無數字）
        if rec.get("點號") is None and rec.get("原始識別碼", "") != "":
            anomalies.append(_new_anomaly(
                "欄位缺漏",
                f"識別碼「{rec.get('原始識別碼')}」結尾無數字，無法解析出點號。",
                rec, x_val, y_val, z_val,
                "請人工確認此識別碼格式是否正確，或於 config/rules.yaml 調整解析規則。",
            ))

    # ---------- 2. 同一管線日期不一致 ----------
    pipeline_groups = defaultdict(list)
    for rec in records:
        pid = rec.get("管線識別碼", "")
        if pid:
            pipeline_groups[pid].append(rec)

    for pid, group in pipeline_groups.items():
        dates = set(r.get("測量日期", "") for r in group if r.get("測量日期", "") != "")
        if len(dates) > 1:
            for rec in group:
                anomalies.append(_new_anomaly(
                    "同一管線日期不一致",
                    f"管線識別碼「{pid}」內各點位測量日期不一致，出現的日期有：{sorted(dates)}。",
                    rec, rec.get("X"), rec.get("Y"), rec.get("Z"),
                    "請人工確認是否為分批測量，並確認日期記錄是否正確。",
                ))

    # ---------- 3. 同一管線識別碼 + 點號重複 / 重複點位比對 ----------
    key_groups = defaultdict(list)
    for rec in records:
        pid = rec.get("管線識別碼", "")
        pno = rec.get("點號")
        if pid and pno is not None:
            key_groups[(pid, pno)].append(rec)

    duplicates = []

    for (pid, pno), group in key_groups.items():
        if len(group) < 2:
            continue

        # 3a. 標記「識別碼+點號重複」
        for rec in group:
            anomalies.append(_new_anomaly(
                "識別碼+點號重複",
                f"管線識別碼「{pid}」的點號「{pno}」共出現 {len(group)} 筆資料。",
                rec, rec.get("X"), rec.get("Y"), rec.get("Z"),
                "請人工確認是否為重複測量、補測或誤植，並決定保留哪一筆資料。",
            ))

        # 3b. 兩兩比對 XY / Z 差距
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                rec_a = group[i]
                rec_b = group[j]

                xa, ya, za = rec_a.get("X"), rec_a.get("Y"), rec_a.get("Z")
                xb, yb, zb = rec_b.get("X"), rec_b.get("Y"), rec_b.get("Z")

                if xa is None or ya is None or xb is None or yb is None:
                    xy_diff = None
                else:
                    xy_diff = math.sqrt((xa - xb) ** 2 + (ya - yb) ** 2)

                if za is None or zb is None:
                    z_diff = None
                else:
                    z_diff = abs(za - zb)

                if xy_diff is None or z_diff is None:
                    verdict = "座標不完整，無法比對差距"
                    suggestion = "請人工確認兩筆資料的座標是否完整後再行比對。"
                elif xy_diff <= xy_tol and z_diff <= z_tol:
                    verdict = "近似重複，待確認"
                    suggestion = "兩筆資料座標相近，請人工確認是否為同一點位的重複測量，並決定保留哪一筆。"
                else:
                    verdict = "重複點位座標差異過大，待確認"
                    suggestion = "兩筆資料座標差異較大，請人工確認是否為不同點位誤用相同識別碼，或其中一筆有誤。"

                for rec_self, rec_other in ((rec_a, rec_b), (rec_b, rec_a)):
                    duplicates.append({
                        "來源CSV": rec_self.get("來源CSV", ""),
                        "原始列號": rec_self.get("原始列號", ""),
                        "原始識別碼": rec_self.get("原始識別碼", ""),
                        "管線識別碼": pid,
                        "點號": pno,
                        "測量日期": rec_self.get("測量日期", ""),
                        "X": rec_self.get("X"),
                        "Y": rec_self.get("Y"),
                        "Z": rec_self.get("Z"),
                        "與比對點XY差距": xy_diff,
                        "與比對點Z差距": z_diff,
                        "判定結果": verdict,
                        "建議處理方式": suggestion,
                    })

                if verdict != "座標不完整，無法比對差距":
                    a_type = (
                        "重複點位近似（XY/Z差距在容許範圍內）"
                        if verdict == "近似重複，待確認"
                        else "重複點位座標差異過大"
                    )
                    desc = (
                        f"與來源 CSV「{rec_b.get('來源CSV')}」第 {rec_b.get('原始列號')} 列"
                        f"（識別碼「{rec_b.get('原始識別碼')}」）比對，"
                        f"XY 差距約 {xy_diff:.3f} m，Z 差距約 {z_diff:.3f} m。"
                    )
                    anomalies.append(_new_anomaly(
                        a_type, desc, rec_a, rec_a.get("X"), rec_a.get("Y"), rec_a.get("Z"), suggestion,
                    ))
                    desc_b = (
                        f"與來源 CSV「{rec_a.get('來源CSV')}」第 {rec_a.get('原始列號')} 列"
                        f"（識別碼「{rec_a.get('原始識別碼')}」）比對，"
                        f"XY 差距約 {xy_diff:.3f} m，Z 差距約 {z_diff:.3f} m。"
                    )
                    anomalies.append(_new_anomaly(
                        a_type, desc_b, rec_b, rec_b.get("X"), rec_b.get("Y"), rec_b.get("Z"), suggestion,
                    ))

    return records, anomalies, duplicates
