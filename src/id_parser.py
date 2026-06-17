"""
id_parser.py - 解析識別碼，拆解出管線識別碼與點號。

規則：識別碼末尾的連續數字為「點號」，其餘為「管線識別碼」。
例：
  1-9-30mTE04~TB06-C1  → 管線識別碼: 1-9-30mTE04~TB06-C  點號: 1
  1-9-30mTE04~TB06-C6  → 管線識別碼: 1-9-30mTE04~TB06-C  點號: 6
  1-9-30mTE04~TB06-C12 → 管線識別碼: 1-9-30mTE04~TB06-C  點號: 12
"""

import re
import pandas as pd


def parse_id(raw_id):
    """
    回傳 (pipeline_id, point_no_str)。
    無法解析時回傳 (raw_id, None)。
    """
    if raw_id is None:
        return None, None

    s = str(raw_id).strip()
    if not s:
        return None, None

    # 末尾為數字序列：分割為前綴（管線識別碼）+ 數字（點號）
    m = re.match(r'^(.*\D)(\d+)$', s)
    if m:
        return m.group(1), m.group(2)

    # 整個識別碼全為數字（特殊情形：無法區分管線與點號）
    if re.match(r'^\d+$', s):
        return s, s

    # 末尾無數字（無法拆出點號）
    return s, None


def try_int(s):
    """嘗試將字串轉為整數，失敗時回傳 None。"""
    if s is None:
        return None
    try:
        return int(s)
    except (ValueError, TypeError):
        return None


def parse_all_ids(df):
    """
    在 df 中新增欄位：
      pipeline_id  - 管線識別碼
      point_no     - 點號（字串）
      point_no_int - 點號（整數，用於排序；無法轉換時為 None）
    回傳新的 DataFrame（不修改原始 df）。
    """
    df = df.copy()

    pipeline_ids = []
    point_nos = []
    point_no_ints = []

    for raw_id in df['raw_id']:
        pid, pno = parse_id(raw_id)
        pipeline_ids.append(pid)
        point_nos.append(pno)
        point_no_ints.append(try_int(pno))

    df['pipeline_id'] = pipeline_ids
    df['point_no'] = point_nos
    df['point_no_int'] = pd.array(point_no_ints, dtype=pd.Int64Dtype())

    return df
