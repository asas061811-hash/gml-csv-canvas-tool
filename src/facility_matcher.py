"""
facility_matcher.py - 設施表匹配功能

參考 SurveyDataProcessor.js 的 refMap 概念，
提供設施屬性表（設計資料）與實測點位的比對與屬性補入。

第一版預留函式介面，後續整合至主畫面。
"""

import re

import pandas as pd


def normalize_facility_id(id_value):
    """
    正規化設施識別碼：去除前後空白。
    回傳 str（空值回傳空字串）。
    """
    if id_value is None:
        return ''
    return str(id_value).strip()


def match_master_sub_id(id_value):
    """
    主次編號智慧配對。

    例如 P-ABb094-1 → 主編號 P-ABb094。
    若識別碼結尾為 -數字 格式，回傳去掉後綴的主編號；
    否則回傳 None（表示不需要回退配對）。

    參考 SurveyDataProcessor.js:
        const suffixMatch = group.baseId.match(/^(.*?)-\\d+$/);
    """
    s = normalize_facility_id(id_value)
    if not s:
        return None
    m = re.match(r'^(.*?)-\d+$', s)
    if m:
        return m.group(1)
    return None


def build_reference_map(ref_df):
    """
    從設施屬性表 DataFrame 建立快速查詢字典。

    嘗試以下欄位作為 key：
        識別碼 > 管線編號 > 設施編號 > 人手孔編號 > 人孔編號 > ID

    回傳 dict: {正規化識別碼: row_dict}

    參考 SurveyDataProcessor.js constructor 的 refMap 建立邏輯。
    """
    if ref_df is None or ref_df.empty:
        return {}

    id_cols = ['識別碼', '管線編號', '設施編號', '人手孔編號', '人孔編號', 'ID']
    ref_map = {}

    for _, row in ref_df.iterrows():
        row_dict = row.to_dict()
        raw_id = None
        for col in id_cols:
            v = row_dict.get(col)
            if v is not None and str(v).strip():
                raw_id = str(v).strip()
                break
        if raw_id:
            ref_map[raw_id] = row_dict

    return ref_map


def match_facility_attributes(points_df, ref_df):
    """
    將設施屬性表的屬性比對到點位 DataFrame。

    對每個點位，依識別碼查詢 ref_map：
    1. 精確比對
    2. 若找不到，嘗試主次編號回退配對

    回傳:
        matched_df: 原 DataFrame 加上匹配到的屬性欄位
        unmatched_ids: 未匹配到的識別碼清單

    參考 SurveyDataProcessor.js process() 的 mode='full' 邏輯。

    第一版為預留介面，後續整合至主畫面時再完善欄位映射。
    """
    ref_map = build_reference_map(ref_df)
    if not ref_map:
        return points_df.copy(), list(points_df.get('edited_raw_id', pd.Series()).dropna().unique())

    matched_df = points_df.copy()
    unmatched_ids = []
    matched_attrs = []

    id_col = 'edited_raw_id' if 'edited_raw_id' in points_df.columns else 'raw_id'

    for _, row in points_df.iterrows():
        raw_id = normalize_facility_id(row.get(id_col))
        ref_row = ref_map.get(raw_id)

        if ref_row is None:
            master_id = match_master_sub_id(raw_id)
            if master_id:
                ref_row = ref_map.get(master_id)

        if ref_row:
            matched_attrs.append(ref_row)
        else:
            matched_attrs.append({})
            if raw_id:
                unmatched_ids.append(raw_id)

    return matched_df, list(set(unmatched_ids))
