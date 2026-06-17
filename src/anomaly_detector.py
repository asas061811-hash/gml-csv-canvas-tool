"""
anomaly_detector.py - 偵測各類異常點位，輸出異常清單與重複量測比對表。

只標示異常，不刪除、不修改、不取平均、不判斷哪筆正確。
所有決策保留給人工確認。
"""

import math
import yaml
import os
import pandas as pd


DEFAULT_XY_TOL = 0.30  # 公尺
DEFAULT_Z_TOL = 0.10   # 公尺


def load_thresholds(config_path):
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                cfg = yaml.safe_load(f) or {}
            thr = cfg.get('thresholds', {})
            xy = float(thr.get('xy_tolerance', DEFAULT_XY_TOL))
            z = float(thr.get('z_tolerance', DEFAULT_Z_TOL))
            return xy, z
        except Exception:
            pass
    return DEFAULT_XY_TOL, DEFAULT_Z_TOL


def safe_float(val):
    """嘗試轉 float，失敗回傳 None。"""
    if val is None:
        return None
    try:
        result = float(str(val).strip())
        return None if math.isnan(result) else result
    except (ValueError, TypeError):
        return None


def _is_valid(v):
    """檢查數值是否有效（非 None、非 NaN）。"""
    if v is None:
        return False
    if isinstance(v, float) and math.isnan(v):
        return False
    return True


def _anomaly_row(atype, desc, row, suggestion):
    """建立一筆異常記錄。"""
    return {
        '異常類型': atype,
        '異常說明': desc,
        '原始識別碼': row.get('raw_id'),
        '管線識別碼': row.get('pipeline_id'),
        '點號': row.get('point_no'),
        'X': row.get('raw_x'),
        'Y': row.get('raw_y'),
        'Z': row.get('raw_z'),
        '測量日期': row.get('raw_date'),
        '來源CSV': row.get('_source_file'),
        '原始列號': row.get('_original_row'),
        '建議處理方式': suggestion,
    }


def detect_anomalies(df, config_path, file_meta,
                     xy_tolerance=None, z_tolerance=None):
    """
    執行所有異常偵測，回傳：
      df_processed  : 加入數值欄位 (x_val, y_val, z_val) 的 DataFrame
      anomaly_list  : list of dict（異常清單）
      duplicate_list: list of dict（重複量測比對表）

    xy_tolerance / z_tolerance: 若提供則覆蓋 config 設定值。
    """
    _xy, _z = load_thresholds(config_path)
    xy_tol = float(xy_tolerance) if xy_tolerance is not None else _xy
    z_tol  = float(z_tolerance)  if z_tolerance  is not None else _z
    df = df.copy()

    # --- 數值轉換 ---
    df['x_val'] = df['raw_x'].apply(safe_float)
    df['y_val'] = df['raw_y'].apply(safe_float)
    df['z_val'] = df['raw_z'].apply(safe_float)

    anomalies = []
    duplicates = []

    # ----------------------------------------------------------------
    # A. 檔案層級：欄位缺漏
    # ----------------------------------------------------------------
    for filename, meta in file_meta.items():
        for missing_col in meta.get('missing_cols', []):
            col_label = {
                'id': '識別碼', 'category': '類別碼',
                'x': 'X 座標', 'y': 'Y 座標',
                'z': 'Z 高程', 'date': '測量日期',
            }.get(missing_col, missing_col)
            anomalies.append({
                '異常類型': '欄位缺漏',
                '異常說明': f'檔案「{filename}」缺少欄位：{col_label}',
                '原始識別碼': None,
                '管線識別碼': None,
                '點號': None,
                'X': None, 'Y': None, 'Z': None,
                '測量日期': None,
                '來源CSV': filename,
                '原始列號': '（整份檔案）',
                '建議處理方式': f'請確認欄位名稱，或在 config/rules.yaml 的 column_aliases.{missing_col} 中新增對應名稱。',
            })

    # ----------------------------------------------------------------
    # B. 逐列異常檢查
    # ----------------------------------------------------------------
    for _, row in df.iterrows():
        r = row.to_dict()

        # B1. 識別碼空值
        if r.get('raw_id') is None:
            anomalies.append(_anomaly_row(
                '識別碼空值',
                '識別碼欄位為空，無法解析管線與點號',
                r, '請向廠商確認該筆資料的識別碼。',
            ))

        # B2. X 座標空值
        if r.get('raw_x') is None:
            anomalies.append(_anomaly_row(
                '座標空值', 'X 座標欄位為空',
                r, '請向廠商確認該點位的 X 座標。',
            ))
        # B3. X 座標格式異常（非空但無法轉數字）
        elif r['x_val'] is None:
            anomalies.append(_anomaly_row(
                'X 座標格式異常',
                f'X 座標值「{r["raw_x"]}」無法解析為數字',
                r, '請確認 X 座標欄位是否含有非數字字元（單位、符號等）。',
            ))

        # B4. Y 座標空值
        if r.get('raw_y') is None:
            anomalies.append(_anomaly_row(
                '座標空值', 'Y 座標欄位為空',
                r, '請向廠商確認該點位的 Y 座標。',
            ))
        # B5. Y 座標格式異常
        elif r['y_val'] is None:
            anomalies.append(_anomaly_row(
                'Y 座標格式異常',
                f'Y 座標值「{r["raw_y"]}」無法解析為數字',
                r, '請確認 Y 座標欄位是否含有非數字字元（單位、符號等）。',
            ))

        # B6. Z 高程空值
        if r.get('raw_z') is None:
            anomalies.append(_anomaly_row(
                '座標空值', 'Z 高程欄位為空',
                r, '請向廠商確認該點位的 Z 高程。',
            ))

        # B7. 日期空值
        if r.get('raw_date') is None:
            anomalies.append(_anomaly_row(
                '日期空值', '測量日期欄位為空',
                r, '請向廠商確認該點位的測量日期。',
            ))

    # ----------------------------------------------------------------
    # C. 管線層級異常：依 pipeline_id 分組
    # ----------------------------------------------------------------
    valid_pid = df[df['pipeline_id'].notna()]

    for pid, grp in valid_pid.groupby('pipeline_id', sort=False):

        # C1. 同一管線日期不一致
        dates = grp['raw_date'].dropna().unique()
        if len(dates) > 1:
            dates_str = '、'.join(str(d) for d in dates)
            for _, row in grp.iterrows():
                anomalies.append(_anomaly_row(
                    '同一管線日期不一致',
                    f'管線「{pid}」存在多個測量日期：{dates_str}',
                    row.to_dict(),
                    '請向廠商確認該管線是否分多次測量，並說明各點位對應的測量日期。',
                ))

        # C2. 同一管線識別碼 + 點號重複
        pno_groups = grp.groupby('point_no', sort=False)
        for pno, pno_grp in pno_groups:
            if pno is None:
                continue
            if len(pno_grp) <= 1:
                continue

            # 加入異常清單
            for _, row in pno_grp.iterrows():
                anomalies.append(_anomaly_row(
                    '同一管線識別碼 + 點號重複',
                    f'管線「{pid}」點號「{pno}」出現 {len(pno_grp)} 筆資料',
                    row.to_dict(),
                    '請確認是否為重複量測，詳見「重複量測比對表」。',
                ))

            # 計算各筆與第一筆的距離，加入重複量測比對表
            rows_list = [r.to_dict() for _, r in pno_grp.iterrows()]
            ref = rows_list[0]
            ref_x = ref.get('x_val'); ref_x = ref_x if _is_valid(ref_x) else None
            ref_y = ref.get('y_val'); ref_y = ref_y if _is_valid(ref_y) else None
            ref_z = ref.get('z_val'); ref_z = ref_z if _is_valid(ref_z) else None

            for i, r in enumerate(rows_list):
                rx = r.get('x_val')
                ry = r.get('y_val')
                rz = r.get('z_val')

                if i == 0:
                    xy_dist = 0.0
                    z_dist = 0.0
                    verdict = '（首筆，作為比對基準）'
                    suggestion = '請與其他重複筆資料比對後，確認哪一筆座標為正確量測值。'
                else:
                    rx = r.get('x_val')
                    ry = r.get('y_val')
                    rz = r.get('z_val')
                    # 將 pandas NaN 統一視為 None
                    rx = rx if _is_valid(rx) else None
                    ry = ry if _is_valid(ry) else None
                    rz = rz if _is_valid(rz) else None

                    if _is_valid(ref_x) and _is_valid(ref_y) and rx is not None and ry is not None:
                        xy_dist = math.sqrt((rx - ref_x) ** 2 + (ry - ref_y) ** 2)
                    else:
                        xy_dist = None

                    if _is_valid(ref_z) and rz is not None:
                        z_dist = abs(rz - ref_z)
                    else:
                        z_dist = None

                    # 判定
                    if xy_dist is None or z_dist is None:
                        verdict = '無法比對（座標缺漏）'
                        suggestion = '請先補齊座標後再行比對。'
                    elif xy_dist <= xy_tol and z_dist <= z_tol:
                        verdict = '近似重複，待確認'
                        suggestion = (
                            f'XY 差距 {xy_dist:.4f}m（≤{xy_tol}m）、Z 差距 {z_dist:.4f}m（≤{z_tol}m），'
                            '座標接近，請確認哪一筆為正確量測值，或說明重複測量原因。'
                        )
                    else:
                        verdict = '重複點位座標差異過大，待確認'
                        suggestion = (
                            f'XY 差距 {xy_dist:.4f}m、Z 差距 {z_dist:.4f}m，'
                            '差距超過容差，請向廠商說明原因，確認哪一筆為正確量測值。'
                        )

                    # 加入「差距過大」異常清單
                    if xy_dist is not None and z_dist is not None:
                        if xy_dist > xy_tol or z_dist > z_tol:
                            anomalies.append(_anomaly_row(
                                '重複點位座標差異過大',
                                (f'管線「{pid}」點號「{pno}」重複點與首筆 XY 差距 {xy_dist:.4f}m、'
                                 f'Z 差距 {z_dist:.4f}m，超過容差（XY≤{xy_tol}m, Z≤{z_tol}m）'),
                                r,
                                '請向廠商說明座標差距原因，並確認哪一筆為正確量測值。',
                            ))

                xy_dist_fmt = f'{xy_dist:.4f}' if isinstance(xy_dist, float) else xy_dist
                z_dist_fmt = f'{z_dist:.4f}' if isinstance(z_dist, float) else z_dist

                duplicates.append({
                    '來源CSV': r['_source_file'],
                    '原始列號': r['_original_row'],
                    '原始識別碼': r.get('raw_id'),
                    '管線識別碼': pid,
                    '點號': pno,
                    '測量日期': r.get('raw_date'),
                    'X': r.get('raw_x'),
                    'Y': r.get('raw_y'),
                    'Z': r.get('raw_z'),
                    '比對基準來源CSV': ref['_source_file'],
                    '比對基準原始列號': ref['_original_row'],
                    '與比對點XY差距(m)': xy_dist_fmt,
                    '與比對點Z差距(m)': z_dist_fmt,
                    '判定結果': verdict,
                    '建議處理方式': suggestion,
                })

    return df, anomalies, duplicates
