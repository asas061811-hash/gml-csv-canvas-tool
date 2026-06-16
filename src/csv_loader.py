# -*- coding: utf-8 -*-
"""
csv_loader.py
--------------
讀取 input_csv 資料夾內所有 CSV 檔案，並依據 config/rules.yaml 中的
field_aliases 自動偵測欄位名稱，輸出統一格式的資料列（list of dict）。

不會修改原始 CSV 檔案，所有讀取皆為唯讀。
"""

import io
import os
import glob

import pandas as pd

# 標準欄位名稱（程式內部使用）
STANDARD_FIELDS = ["id", "category", "x", "y", "z", "date"]


def _normalize_col_name(name):
    """正規化欄位名稱：去除空白、統一全形括號為半形括號。"""
    if name is None:
        return ""
    name = str(name).strip()
    name = name.replace("（", "(").replace("）", ")")
    return name


def _build_alias_lookup(field_aliases):
    """
    建立 {正規化後的別名: 標準欄位名稱} 的對照表，
    方便用實際欄位名稱反查標準欄位。
    """
    lookup = {}
    for std_field, aliases in field_aliases.items():
        for alias in aliases:
            norm = _normalize_col_name(alias).lower()
            lookup[norm] = std_field
    return lookup


def _detect_column_mapping(header, alias_lookup):
    """
    比對 CSV 表頭與別名清單，回傳 {標準欄位名稱: 實際欄位名稱}。
    若某標準欄位找不到對應的實際欄位，則該標準欄位不會出現在回傳的字典中。
    """
    mapping = {}
    for col in header:
        norm = _normalize_col_name(col).lower()
        if norm in alias_lookup:
            std_field = alias_lookup[norm]
            # 若已經有對應到的欄位，保留第一個找到的，不覆蓋
            if std_field not in mapping:
                mapping[std_field] = col
    return mapping


def _read_csv_with_fallback_encoding(path, encodings):
    """
    使用 pandas 嘗試以多種編碼讀取 CSV，回傳 (header, rows, used_encoding)。
    所有欄位皆以字串（dtype=str）讀入，避免 pandas 自動轉型造成資料失真，
    rows 為 list[dict]，key 為原始欄位名稱，value 為字串。
    """
    last_error = None
    for enc in encodings:
        try:
            df = pd.read_csv(path, dtype=str, encoding=enc, keep_default_na=False)
            header = list(df.columns)
            rows = df.to_dict(orient="records")
            return header, rows, enc
        except (UnicodeDecodeError, UnicodeError) as e:
            last_error = e
            continue
    # 全部編碼都失敗，最後再用 errors="replace" 強制讀取，避免整個程式中斷
    with open(path, "r", encoding="utf-8", errors="replace", newline="") as f:
        df = pd.read_csv(f, dtype=str, keep_default_na=False)
    header = list(df.columns)
    rows = df.to_dict(orient="records")
    return header, rows, "utf-8(replace, 編碼可能有誤)"


def load_all_csv(input_dir, rules):
    """
    讀取 input_dir 內所有 .csv 檔案。

    回傳:
        records: list[dict]，每個 dict 為一個資料列，包含以下 key：
            - 原始識別碼 (str)
            - 類別碼 (str)
            - X_raw, Y_raw, Z_raw (str，尚未轉換為數字的原始值)
            - 測量日期 (str)
            - 來源CSV (str)
            - 原始列號 (int，對應 CSV 中的資料行號，標頭為第1行，第一筆資料為第2行)
            - 缺漏欄位 (list[str]，此列因整個 CSV 缺少對應欄位而標記的標準欄位名稱)
        file_warnings: list[str]，讀取過程中的警告訊息（例如缺少欄位）
    """
    field_aliases = rules.get("field_aliases", {})
    encodings = rules.get("csv_encodings", ["utf-8-sig", "utf-8", "big5", "cp950", "gb18030"])
    alias_lookup = _build_alias_lookup(field_aliases)

    records = []
    file_warnings = []

    csv_paths = sorted(glob.glob(os.path.join(input_dir, "*.csv")))
    if not csv_paths:
        file_warnings.append(f"在 {input_dir} 中找不到任何 CSV 檔案。")
        return records, file_warnings

    for path in csv_paths:
        filename = os.path.basename(path)
        header, rows, used_enc = _read_csv_with_fallback_encoding(path, encodings)

        if "編碼可能有誤" in used_enc:
            file_warnings.append(f"{filename}: 所有預設編碼皆讀取失敗，已以 utf-8 並替換無法解碼字元的方式讀取，內容可能有誤。")

        file_records = _process_file_rows(filename, header, rows, alias_lookup, file_warnings)
        records.extend(file_records)

    return records, file_warnings


# ── Streamlit 網頁版入口 ──────────────────────────────────────────────────────

def _read_bytes_with_fallback_encoding(raw_bytes, encodings):
    """
    以多種編碼嘗試解碼 bytes 並讀入 CSV。
    回傳 (header, rows, used_encoding)，格式同 _read_csv_with_fallback_encoding。
    """
    for enc in encodings:
        try:
            df = pd.read_csv(io.BytesIO(raw_bytes), dtype=str, encoding=enc, keep_default_na=False)
            return list(df.columns), df.to_dict(orient="records"), enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    df = pd.read_csv(
        io.BytesIO(raw_bytes), dtype=str, encoding="utf-8", errors="replace", keep_default_na=False
    )
    return list(df.columns), df.to_dict(orient="records"), "utf-8(replace, 編碼可能有誤)"


def _process_file_rows(filename, header, rows, alias_lookup, file_warnings):
    """
    load_all_csv 與 load_from_uploads 共用的內部處理邏輯。
    回傳該檔案的 records list。
    """
    if not header:
        file_warnings.append(f"{filename}: 無法讀取表頭，已略過此檔案。")
        return []

    mapping = _detect_column_mapping(header, alias_lookup)

    missing_fields = [f for f in STANDARD_FIELDS if f not in mapping]
    if missing_fields:
        file_warnings.append(
            f"{filename}: 找不到對應欄位 {missing_fields}，"
            f"請確認該檔案是否有此欄位，或於 config/rules.yaml 中新增別名對應。"
        )

    records = []
    for idx, row in enumerate(rows):
        line_no = idx + 2

        def get_raw(std_field, _mapping=mapping, _row=row):
            col = _mapping.get(std_field)
            if col is None:
                return ""
            val = _row.get(col, "")
            return str(val).strip() if val is not None else ""

        records.append({
            "原始識別碼": get_raw("id"),
            "類別碼": get_raw("category"),
            "X_raw": get_raw("x"),
            "Y_raw": get_raw("y"),
            "Z_raw": get_raw("z"),
            "測量日期": get_raw("date"),
            "來源CSV": filename,
            "原始列號": line_no,
            "缺漏欄位": list(missing_fields),
        })

    return records


def load_from_uploads(uploaded_files, rules):
    """
    讀取 Streamlit file_uploader 回傳的 UploadedFile 物件清單。

    uploaded_files: list of streamlit UploadedFile
                    每個物件需有 .name (str) 與 .read() → bytes。

    回傳與 load_all_csv 相同的 (records, file_warnings)。
    不會修改任何原始 CSV，所有讀取皆為 in-memory。
    """
    field_aliases = rules.get("field_aliases", {})
    encodings = rules.get("csv_encodings", ["utf-8-sig", "utf-8", "big5", "cp950", "gb18030"])
    alias_lookup = _build_alias_lookup(field_aliases)

    records = []
    file_warnings = []

    if not uploaded_files:
        file_warnings.append("沒有上傳任何 CSV 檔案。")
        return records, file_warnings

    for uf in uploaded_files:
        filename = uf.name
        raw_bytes = uf.read()

        header, rows, used_enc = _read_bytes_with_fallback_encoding(raw_bytes, encodings)

        if "編碼可能有誤" in used_enc:
            file_warnings.append(
                f"{filename}: 所有預設編碼皆讀取失敗，已以 utf-8 並替換無法解碼字元的方式讀取，內容可能有誤。"
            )

        file_records = _process_file_rows(filename, header, rows, alias_lookup, file_warnings)
        records.extend(file_records)

    return records, file_warnings
