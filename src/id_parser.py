# -*- coding: utf-8 -*-
"""
id_parser.py
------------
解析「原始識別碼」，拆出「管線識別碼」與「點號」。

規則：抓取識別碼結尾的連續數字作為點號，其餘部分為管線識別碼。

範例：
    1-9-30mTE04~TB06-C1  ->  管線識別碼: 1-9-30mTE04~TB06-C   點號: 1
    1-9-30mTE04~TB06-C6  ->  管線識別碼: 1-9-30mTE04~TB06-C   點號: 6

若識別碼結尾沒有數字，則：
    管線識別碼 = 原始識別碼（整串）
    點號 = None
"""

import re


def parse_id(raw_id, pattern=r"^(.*?)(\d+)$"):
    """
    解析單一識別碼字串。

    回傳:
        (管線識別碼, 點號)
        點號為 int，若無法解析則為 None。
    """
    if raw_id is None:
        return "", None

    raw_id = str(raw_id).strip()
    if raw_id == "":
        return "", None

    m = re.match(pattern, raw_id)
    if m:
        pipeline_id = m.group(1)
        point_no = int(m.group(2))
        return pipeline_id, point_no
    else:
        # 結尾沒有數字，整串視為管線識別碼，點號留空
        return raw_id, None


def parse_records(records, rules):
    """
    為每筆 record 加入「管線識別碼」與「點號」欄位。
    直接修改傳入的 records（list[dict]），並回傳同一份 list。
    """
    pattern = rules.get("id_parse", {}).get("trailing_digit_regex", r"^(.*?)(\d+)$")

    for rec in records:
        pipeline_id, point_no = parse_id(rec.get("原始識別碼", ""), pattern)
        rec["管線識別碼"] = pipeline_id
        rec["點號"] = point_no

    return records
