"""
attribute_merger.py - 分類表屬性合併與代碼轉換

流程：
  1. 讀取人工填寫的「分類表_定義版」Excel
  2. 與目前點位分類表合併（只補屬性，不覆蓋座標）
  3. 將定義文字轉換為標準代碼 → 產生代碼版
  4. 產生代碼定義對照版
"""

import io
import os
import math

import openpyxl
import pandas as pd
import yaml


# ── 代碼對照表載入 ──────────────────────────────────

def load_code_mappings(config_path):
    """從 rules.yaml 載入 code_mappings，回傳 dict[欄位名, dict[代碼str, 定義str]]。"""
    if not os.path.exists(config_path):
        return {}
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            cfg = yaml.safe_load(f) or {}
        return cfg.get('code_mappings', {})
    except Exception:
        return {}


def build_def_to_code_map(code_mappings):
    """
    從 code_mappings 建立反向對照：{欄位名: {定義文字: 代碼str}}。
    支援不區分前後空白。
    """
    result = {}
    for field, mapping in code_mappings.items():
        rev = {}
        for code, definition in mapping.items():
            rev[str(definition).strip()] = str(code)
        result[field] = rev
    return result


def build_code_to_def_map(code_mappings):
    """從 code_mappings 建立正向對照：{欄位名: {代碼str: 定義文字}}。"""
    result = {}
    for field, mapping in code_mappings.items():
        result[field] = {str(k): str(v) for k, v in mapping.items()}
    return result


# ── Excel 讀取 ──────────────────────────────────────

_SHEET_NAME_MAP = {
    '管線': '管線', '分類表_管線': '管線', '分類表_管線_定義版': '管線',
    '人手孔': '人手孔', '分類表_人手孔': '人手孔', '分類表_人手孔_定義版': '人手孔',
    '開關閥': '開關閥', '分類表_開關閥': '開關閥', '分類表_開關閥_定義版': '開關閥',
    '消防栓': '消防栓', '分類表_消防栓': '消防栓', '分類表_消防栓_定義版': '消防栓',
    '電桿': '電桿', '分類表_電桿': '電桿', '分類表_電桿_定義版': '電桿',
    '號誌': '號誌', '分類表_號誌': '號誌', '分類表_號誌_定義版': '號誌',
    '其他設施': '其他設施', '分類表_其他設施': '其他設施', '分類表_其他設施_定義版': '其他設施',
    '維護口': '維護口', '分類表_維護口': '維護口', '分類表_維護口_定義版': '維護口',
    '場站': '場站', '分類表_場站': '場站', '分類表_場站_定義版': '場站',
}


def detect_facility_sheet_name(sheet_name):
    """自動辨識工作表名屬於哪個設施分類。回傳設施類型名或 None。"""
    s = str(sheet_name).strip()
    if s in _SHEET_NAME_MAP:
        return _SHEET_NAME_MAP[s]
    for key, val in _SHEET_NAME_MAP.items():
        if key in s:
            return val
    return None


def load_definition_excel(uploaded_file):
    """
    讀取人工上傳的分類表_定義版 Excel。
    回傳 dict[設施類型, DataFrame]。
    """
    result = {}
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception as e:
        return result, f'Excel 讀取失敗：{e}'

    for sheet_name in xls.sheet_names:
        ft = detect_facility_sheet_name(sheet_name)
        if ft is None:
            continue
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
            df = df.fillna('')
            df.columns = [str(c).strip() for c in df.columns]
            result[ft] = df
        except Exception:
            continue
    return result, None


# ── 合併邏輯 ────────────────────────────────────────

_COORD_PREFIXES = ('E', 'N', 'Z', 'd', '第')

def _is_coord_col(col):
    """判斷是否為座標欄位（不可被覆蓋）。"""
    s = str(col).strip()
    if s in ('E', 'N', 'Z', 'd', 'X', 'Y'):
        return True
    if s.startswith('第') and any(s.endswith(f'點{x}') for x in ('X','Y','Z','d')):
        return True
    if s.startswith('E(') or s.startswith('N(') or s.startswith('Z(') or s.startswith('d('):
        return True
    return False


def _find_match_key(def_row, point_row, facility_type):
    """嘗試多種匹配方式，回傳是否匹配。"""
    def _s(val):
        if val is None:
            return ''
        return str(val).strip()

    id_d = _s(def_row.get('識別碼'))
    id_p = _s(point_row.get('識別碼'))
    if id_d and id_p and id_d == id_p:
        return True

    if facility_type == '管線':
        pn_d = _s(def_row.get('管線編號'))
        pn_p = _s(point_row.get('管線編號'))
        if pn_d and pn_p and pn_d == pn_p:
            return True

    cat_d = _s(def_row.get('類別碼'))
    cat_p = _s(point_row.get('類別碼'))
    if cat_d and id_d and cat_d == cat_p and id_d == id_p:
        return True

    return False


def merge_definition_attributes(point_cls_df, definition_df, facility_type):
    """
    將定義版屬性合併到點位分類表。
    只補入屬性欄位，不覆蓋座標與 point_key。
    回傳 (merged_df, unmatched_def_rows, unfilled_point_rows)。
    """
    if point_cls_df is None or point_cls_df.empty:
        return point_cls_df, definition_df, pd.DataFrame()
    if definition_df is None or definition_df.empty:
        return point_cls_df, pd.DataFrame(), point_cls_df

    merged = point_cls_df.copy()
    def_used = set()
    unfilled_idxs = []

    def_records = definition_df.to_dict('records')

    for pidx in range(len(merged)):
        point_row = merged.iloc[pidx].to_dict()
        matched = False

        for didx, def_row in enumerate(def_records):
            if didx in def_used:
                continue
            if _find_match_key(def_row, point_row, facility_type):
                for col, val in def_row.items():
                    if _is_coord_col(col):
                        continue
                    if col in ('point_key', '_source_file', '_original_row'):
                        continue
                    v = str(val).strip()
                    if v and col in merged.columns:
                        cur = str(merged.iloc[pidx][col]).strip()
                        if not cur:
                            merged.iat[pidx, merged.columns.get_loc(col)] = v
                def_used.add(didx)
                matched = True
                break

        if not matched:
            unfilled_idxs.append(pidx)

    unmatched_def = definition_df.iloc[[i for i in range(len(def_records)) if i not in def_used]].copy()
    unfilled_points = merged.iloc[unfilled_idxs].copy() if unfilled_idxs else pd.DataFrame()

    return merged, unmatched_def, unfilled_points


# ── 代碼轉換 ────────────────────────────────────────

def convert_definition_to_code(df, code_mappings):
    """
    將 DataFrame 中的定義文字轉成代碼。
    回傳 (converted_df, errors)。
    errors: list of dict {row, field, value, error}。
    """
    if df is None or df.empty:
        return df, []

    def_to_code = build_def_to_code_map(code_mappings)
    code_to_def = build_code_to_def_map(code_mappings)
    result = df.copy()
    errors = []

    for col in result.columns:
        if col not in def_to_code and col not in code_to_def:
            continue
        d2c = def_to_code.get(col, {})
        c2d = code_to_def.get(col, {})

        for idx in range(len(result)):
            val = str(result.iat[idx, result.columns.get_loc(col)]).strip()
            if not val:
                continue
            if val in c2d:
                pass
            elif val in d2c:
                result.iat[idx, result.columns.get_loc(col)] = d2c[val]
            else:
                errors.append({
                    'row': idx + 1,
                    'field': col,
                    'value': val,
                    'error': f'無法將「{val}」轉換為代碼',
                })

    return result, errors


def convert_code_to_definition(df, code_mappings):
    """將 DataFrame 中的代碼轉回定義文字。"""
    if df is None or df.empty:
        return df

    code_to_def = build_code_to_def_map(code_mappings)
    result = df.copy()

    for col in result.columns:
        if col not in code_to_def:
            continue
        c2d = code_to_def[col]
        for idx in range(len(result)):
            val = str(result.iat[idx, result.columns.get_loc(col)]).strip()
            if val in c2d:
                result.iat[idx, result.columns.get_loc(col)] = c2d[val]

    return result


def build_code_definition_table(df, code_mappings):
    """
    產生代碼定義對照版：原始代碼欄 + 新增 {欄位}_定義 欄。
    """
    if df is None or df.empty:
        return df

    code_to_def = build_code_to_def_map(code_mappings)
    result = df.copy()

    insert_after = []
    for col in list(result.columns):
        if col not in code_to_def:
            continue
        c2d = code_to_def[col]
        def_col = f'{col}_定義'
        vals = []
        for idx in range(len(result)):
            v = str(result.iat[idx, result.columns.get_loc(col)]).strip()
            vals.append(c2d.get(v, v))
        insert_after.append((col, def_col, vals))

    for col, def_col, vals in reversed(insert_after):
        loc = result.columns.get_loc(col) + 1
        result.insert(loc, def_col, vals)

    return result


# ── 檢核訊息 ────────────────────────────────────────

def validate_merge_result(merged_df, unmatched_def, unfilled_points, convert_errors,
                          excluded_count=0):
    """產生合併與轉換檢核摘要。"""
    n_merged = len(merged_df) if merged_df is not None else 0
    n_unmatched = len(unmatched_def) if unmatched_def is not None else 0
    n_unfilled = len(unfilled_points) if unfilled_points is not None else 0
    n_convert_err = len(convert_errors) if convert_errors else 0

    return {
        'success_count': n_merged,
        'unmatched_def_count': n_unmatched,
        'unfilled_point_count': n_unfilled,
        'convert_error_count': n_convert_err,
        'excluded_count': excluded_count,
        'convert_errors': convert_errors or [],
    }
