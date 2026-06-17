"""
csv_loader.py - 讀取 input_csv 資料夾內所有 CSV，自動偵測編碼與欄位，
輸出標準化 DataFrame。不修改原始 CSV。

另提供 load_uploaded_csvs() 供 Streamlit 使用。
"""

import io
import os
import glob
import pandas as pd
import yaml

ENCODINGS_TO_TRY = ['utf-8-sig', 'utf-8', 'big5', 'cp950', 'gbk', 'latin-1']

DEFAULT_COLUMN_ALIASES = {
    'id':       ['識別碼', 'ID', 'id', '點位識別碼', 'point_id', '編號', '點號'],
    'category': ['類別碼', '類別', '管種', 'category', 'type', 'CAT', 'cat'],
    'x':        ['X', 'E', 'X(E)', 'X座標', 'Easting', 'x', 'e'],
    'y':        ['Y', 'N', 'Y(N)', 'Y座標', 'Northing', 'y', 'n'],
    'z':        ['Z', 'H', 'Z(H)', '高程', 'Height', 'Elevation', 'elevation', 'z', 'h'],
    'date':     ['日期', '測量日期', '設置日期', '量測日期', 'survey_date', 'date', 'DATE'],
}

STD_FIELDS = list(DEFAULT_COLUMN_ALIASES.keys())

# 判斷完全空白列的欄位清單（全空 → 排除，不進入任何成果表）
_BLANK_CHECK_FIELDS = ['raw_id', 'raw_category', 'raw_x', 'raw_y', 'raw_z']


def _is_empty_val(v):
    """判斷一個值是否為空（None、NaN、空字串）。"""
    if v is None:
        return True
    try:
        import math
        if isinstance(v, float) and math.isnan(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip() == ''


def filter_blank_rows(df):
    """
    排除無任何有效點位資訊的空白列。
    判斷條件：raw_id、raw_category、raw_x、raw_y、raw_z 全部為空（None/NaN/空字串）。
    回傳 (filtered_df, n_removed)。
    """
    if df.empty:
        return df, 0
    def _is_blank(row):
        return all(_is_empty_val(row.get(f)) for f in _BLANK_CHECK_FIELDS)
    mask = df.apply(_is_blank, axis=1)
    n_removed = int(mask.sum())
    return df[~mask].reset_index(drop=True), n_removed


def load_config(config_path):
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            print(f'  [警告] 設定檔讀取失敗：{e}')
    return {}


def detect_encoding(file_path):
    """嘗試各種編碼，回傳第一個成功讀取的編碼名稱。"""
    for enc in ENCODINGS_TO_TRY:
        try:
            with open(file_path, encoding=enc, errors='strict') as f:
                f.read(8192)
            return enc
        except (UnicodeDecodeError, LookupError):
            continue
    return 'utf-8'


def read_csv_with_encoding(file_path):
    """讀取 CSV，自動偵測編碼。回傳 (DataFrame, encoding_str)。"""
    enc = detect_encoding(file_path)
    candidates = [enc] + [e for e in ENCODINGS_TO_TRY if e != enc]
    last_err = None
    for attempt_enc in candidates:
        try:
            df = pd.read_csv(
                file_path,
                encoding=attempt_enc,
                dtype=str,
                keep_default_na=False,  # 避免 'N'、'NA' 等被轉成 NaN
                na_values=[''],
            )
            # 去除欄位名稱前後空白
            df.columns = [c.strip() for c in df.columns]
            return df, attempt_enc
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f'無法讀取 CSV（已嘗試所有編碼）：{last_err}')


def merge_aliases(config_aliases):
    """將設定檔的別名合併到預設別名（設定檔優先）。"""
    merged = {k: list(v) for k, v in DEFAULT_COLUMN_ALIASES.items()}
    for key, vals in config_aliases.items():
        if key in merged:
            # 設定檔中的別名放最前面（優先比對）
            existing = merged[key]
            new_unique = [v for v in vals if v not in existing]
            merged[key] = new_unique + existing
        else:
            merged[key] = list(vals)
    return merged


def map_columns(df_columns, aliases):
    """
    依別名清單找出 CSV 欄位對應的標準欄位名。
    回傳 dict: {std_name: csv_column_name}
    """
    col_lower = {c.strip().lower(): c for c in df_columns}
    result = {}
    for std_name, alias_list in aliases.items():
        for alias in alias_list:
            alias_stripped = alias.strip()
            # 精確比對
            if alias_stripped in df_columns:
                result[std_name] = alias_stripped
                break
            # 大小寫不分比對
            if alias_stripped.lower() in col_lower:
                result[std_name] = col_lower[alias_stripped.lower()]
                break
    return result


def load_all_csvs(input_dir, config_path):
    """
    讀取 input_dir 內所有 CSV，回傳：
      df_std   : 標準化 DataFrame
      file_meta: {filename: {'col_map': dict, 'missing_cols': list, 'encoding': str}}
      load_errors: [{'file': str, 'error': str}]
    """
    config = load_config(config_path)
    config_aliases = config.get('column_aliases', {})
    aliases = merge_aliases(config_aliases)

    # 尋找 CSV（不遞迴，只找第一層）
    pattern_lower = os.path.join(input_dir, '*.csv')
    pattern_upper = os.path.join(input_dir, '*.CSV')
    csv_files = sorted(set(
        glob.glob(pattern_lower) + glob.glob(pattern_upper)
    ))

    if not csv_files:
        print(f'[警告] {input_dir} 內找不到任何 CSV 檔案。')
        return pd.DataFrame(), {}, []

    all_records = []
    file_meta = {}
    load_errors = []

    for csv_file in csv_files:
        filename = os.path.basename(csv_file)
        print(f'\n讀取: {filename}')

        try:
            df_raw, encoding = read_csv_with_encoding(csv_file)
        except RuntimeError as e:
            msg = str(e)
            print(f'  [錯誤] {msg}')
            load_errors.append({'file': filename, 'error': msg})
            continue

        print(f'  編碼: {encoding}，共 {len(df_raw)} 列，欄位: {list(df_raw.columns)}')

        col_map = map_columns(list(df_raw.columns), aliases)
        missing_cols = [k for k in STD_FIELDS if k not in col_map]

        print(f'  欄位對應: {col_map}')
        if missing_cols:
            print(f'  [警告] 未找到對應欄位: {missing_cols}')

        file_meta[filename] = {
            'col_map': col_map,
            'missing_cols': missing_cols,
            'encoding': encoding,
        }

        # 建立標準化列
        records = []
        for i, (_, row) in enumerate(df_raw.iterrows()):
            rec = {
                '_source_file': filename,
                '_original_row': i + 2,  # CSV 第 1 列為標題，資料從第 2 列起
            }
            for std_name in STD_FIELDS:
                if std_name in col_map:
                    raw_val = row.get(col_map[std_name])
                    # 空字串統一轉 None
                    rec[f'raw_{std_name}'] = raw_val if (raw_val is not None and str(raw_val).strip() != '') else None
                else:
                    rec[f'raw_{std_name}'] = None
            records.append(rec)

        df_std, n_blank = filter_blank_rows(pd.DataFrame(records))
        if n_blank:
            print(f'  [清理] 排除 {n_blank} 筆空白列')
        all_records.append(df_std)

    if not all_records:
        return pd.DataFrame(), file_meta, load_errors

    combined = pd.concat(all_records, ignore_index=True)
    print(f'\n共載入 {len(combined)} 筆資料（來自 {len(all_records)} 個檔案）')
    return combined, file_meta, load_errors


# ============================================================
# Streamlit 版：接受 UploadedFile 物件列表
# ============================================================

def _parse_uploaded_file(uploaded_file, aliases):
    """讀取單一 Streamlit UploadedFile，回傳 (df_std, meta) 或 (None, error_str)。"""
    filename = uploaded_file.name
    content = uploaded_file.read()

    df_raw = None
    encoding = None
    for enc in ENCODINGS_TO_TRY:
        try:
            df_raw = pd.read_csv(
                io.BytesIO(content),
                encoding=enc,
                dtype=str,
                keep_default_na=False,
                na_values=[''],
            )
            df_raw.columns = [c.strip() for c in df_raw.columns]
            encoding = enc
            break
        except Exception:
            continue

    if df_raw is None:
        return None, None, f'無法解析（已嘗試所有編碼）'

    col_map = map_columns(list(df_raw.columns), aliases)
    missing_cols = [k for k in STD_FIELDS if k not in col_map]

    meta = {'col_map': col_map, 'missing_cols': missing_cols, 'encoding': encoding}

    records = []
    for i, (_, row) in enumerate(df_raw.iterrows()):
        rec = {'_source_file': filename, '_original_row': i + 2}
        for std_name in STD_FIELDS:
            if std_name in col_map:
                raw_val = row.get(col_map[std_name])
                rec[f'raw_{std_name}'] = raw_val if (raw_val is not None and str(raw_val).strip() != '') else None
            else:
                rec[f'raw_{std_name}'] = None
        records.append(rec)

    raw_df = pd.DataFrame(records)
    cleaned_df, n_blank = filter_blank_rows(raw_df)
    if n_blank:
        meta['blank_rows_removed'] = n_blank
    return cleaned_df, meta, None


def load_uploaded_csvs(uploaded_files, config_path):
    """
    Streamlit 版 load_all_csvs：接受 st.file_uploader 回傳的 UploadedFile 列表。
    回傳 (df_std, file_meta, load_errors)，格式與 load_all_csvs 相同。
    """
    config = load_config(config_path)
    config_aliases = config.get('column_aliases', {})
    aliases = merge_aliases(config_aliases)

    all_records = []
    file_meta = {}
    load_errors = []

    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        df_std, meta, err = _parse_uploaded_file(uploaded_file, aliases)
        if err:
            load_errors.append({'file': filename, 'error': err})
            continue
        file_meta[filename] = meta
        all_records.append(df_std)

    if not all_records:
        return pd.DataFrame(), file_meta, load_errors

    combined = pd.concat(all_records, ignore_index=True)
    return combined, file_meta, load_errors
