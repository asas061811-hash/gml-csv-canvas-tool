"""
gml_generator.py - GML 產製引擎

將修正後點位資料或 GML 專用寬表 CSV 轉換為 GML XML 格式。
支援目標規範：桃園市、國土署、南科/世曦。
第一版完整實作：桃園市 + 管線。
"""

import io
import math
import re
from xml.dom.minidom import getDOMImplementation
from xml.etree.ElementTree import Element, SubElement, tostring

import pandas as pd

# ── 目標規範常數 ──────────────────────────────────
TARGET_TAOYUAN = 'Taoyuan'
TARGET_NLMA = 'NLMA'
TARGET_STEC = 'STEC'

TARGET_OPTIONS = {
    TARGET_TAOYUAN: '桃園市',
    TARGET_NLMA: '國土署',
    TARGET_STEC: '南科 / 世曦',
}

ELEVATION_ABSOLUTE = 'absolute'
ELEVATION_GROUND = 'ground'

ELEVATION_OPTIONS = {
    ELEVATION_ABSOLUTE: 'Z 已是絕對高程，直接輸出',
    ELEVATION_GROUND: 'Z 為地面高程，輸出 Z − d',
}

# 目前已實作的組合
_IMPLEMENTED = {
    (TARGET_TAOYUAN, '管線'),
    (TARGET_TAOYUAN, '人手孔'),
    (TARGET_TAOYUAN, '開關閥'),
    (TARGET_TAOYUAN, '消防栓'),
    (TARGET_TAOYUAN, '電桿'),
    (TARGET_TAOYUAN, '號誌'),
    (TARGET_TAOYUAN, '其他設施'),
    (TARGET_TAOYUAN, '維護口'),
    (TARGET_TAOYUAN, '場站'),
}


def is_implemented(target, facility_type):
    return (target, facility_type) in _IMPLEMENTED


# ============================================================
# 1. get_feature_config - 依類別碼判斷 GML tag 與幾何類型
# ============================================================

def get_feature_config(category_code, target=TARGET_TAOYUAN):
    """
    依類別碼判斷 GML featureMember 的 tagName 與幾何類型。
    回傳 (tag_name, geom_type)。
    geom_type: 'LineString' 或 'Point'。
    """
    code = str(category_code or '').strip()
    if len(code) != 7 or not code.isdigit():
        return ('UTL_其他設施', 'Point')

    end = code[5:7]   # fine code
    mid1 = code[1:3]  # mid code
    mid2 = code[3:5]  # minor code

    if end == '01':
        tag = 'UTL_管線'
        if mid1 == '03':
            tag = 'UTL_管線_自來水'
        elif mid1 == '05':
            tag = 'UTL_管線_供氣'
        elif mid1 == '07':
            tag = 'UTL_管線_輸油'

        if target == TARGET_TAOYUAN:
            if mid1 == '02':
                tag = 'UTL_管道'
            elif mid1 == '04' and mid2 == '01':
                tag = 'UTL_管線_汙水'

        return (tag, 'LineString')

    tag = 'UTL_其他設施'
    if end == '02':
        tag = 'UTL_人手孔'
    elif end == '03':
        if mid1 in ('01', '02'):
            tag = 'UTL_電桿'
        elif mid1 == '03':
            tag = 'UTL_消防栓'
        elif mid1 == '04':
            tag = 'UTL_人手孔'
        elif mid1 in ('05', '06', '07'):
            tag = 'UTL_開關閥'
        else:
            tag = 'UTL_維護口'
    elif end == '04':
        if mid1 == '01' and mid2 == '05':
            tag = 'UTL_號誌'
        elif mid1 in ('02', '03'):
            tag = 'UTL_開關閥'
    elif end == '96':
        tag = '其他設施'
    elif end == '97':
        tag = '場站'

    return (tag, 'Point')


# ============================================================
# 2. get_schema_for_tag - 依 tag 回傳 GML 屬性欄位清單
# ============================================================

def get_schema_for_tag(tag_name, category_code, target=TARGET_TAOYUAN):
    """回傳該 GML tag 對應的屬性欄位清單（不含座標）。"""
    code = str(category_code or '').strip()
    end = code[5:7] if len(code) == 7 else ''
    mid1 = code[1:3] if len(code) >= 3 else ''

    if tag_name.startswith('UTL_管線') or tag_name == 'UTL_管道':
        schema = ['類別碼', '識別碼', '起點編號', '終點編號', '管理單位', '作業區分', '設置日期']
        if tag_name == 'UTL_管道':
            schema += ['IndexNo', '管道編號']
        else:
            schema.append('管線編號')
        schema += [
            '尺寸單位', '管徑寬度', '管徑高度', '涵管條數', '管線材料',
            '起點埋設深度', '終點埋設深度', '管線長度', '管線型態',
            '使用狀態', '資料狀態', '備註',
        ]
        if mid1 in ('04', '03'):
            schema.append('輸送物質')
        elif mid1 in ('05', '07', '91'):
            schema += ['壓力區分', '輸送物質']
        schema += ['參考模型代碼', '旋轉角']
        return schema

    if tag_name == 'UTL_人手孔':
        schema = ['類別碼', '識別碼', '管理單位', '作業區分', '設置日期',
                  '人手孔編號', '孔蓋種類', '尺寸單位', '蓋部寬度', '蓋部長度']
        if mid1 == '06':
            schema.append('閘門名稱')
        schema += ['地盤高', '孔深', '孔蓋型態', '使用狀態', '資料狀態']
        if mid1 in ('03', '05', '07', '91'):
            schema.append('內容物')
        schema += ['備註', '旋轉角']
        return schema

    if tag_name == 'UTL_電桿':
        return ['類別碼', '識別碼', '管理單位', '作業區分', '設置日期',
                '電桿編號', '長度', '材質', '使用狀態', '資料狀態', '備註',
                '參考模型代碼', '旋轉角']

    if tag_name == 'UTL_消防栓':
        return ['類別碼', '識別碼', '管理單位', '作業區分', '設置日期',
                '消防栓編號', '管身口徑', '出水口口徑', '埋設深度', '消防栓型態',
                '使用狀態', '資料狀態', '備註', '參考模型代碼', '旋轉角']

    if tag_name == 'UTL_開關閥':
        schema = ['類別碼', '識別碼', '管理單位', '作業區分', '設置日期', '開關閥編號']
        if end == '03':
            schema += ['閥類編號', '口徑', '名稱']
        elif end == '04':
            schema += ['口徑', '閥類編號']
            if mid1 == '02':
                schema.append('名稱')
        schema += ['地盤高', '埋設深度', '開關閥型態', '使用狀態', '資料狀態', '備註',
                   '參考模型代碼', '旋轉角']
        return schema

    if tag_name == 'UTL_維護口':
        return ['類別碼', '識別碼', '管理單位', '作業區分', '設置日期',
                '維護口編號', '名稱', '使用狀態', '資料狀態', '備註',
                '參考模型代碼', '旋轉角']

    if tag_name == 'UTL_號誌':
        return ['類別碼', '識別碼', '管理單位', '作業區分', '設置日期',
                '號誌編號', '號誌種類', '號誌架設方式', '長度',
                '使用狀態', '資料狀態', '備註', '參考模型代碼', '旋轉角']

    if tag_name in ('其他設施', 'UTL_其他設施'):
        return ['類別碼', '識別碼', '管理單位', '作業區分', '設置日期',
                '設施編號', '設施名稱', '設施長度', '設施寬度', '設施高度', '設施型態',
                '使用狀態', '資料狀態', '備註', '參考模型代碼', '旋轉角']

    if tag_name in ('場站', 'UTL_場站'):
        return ['類別碼', '識別碼', '管理單位', '作業區分', '設置日期',
                '場站名稱', '使用狀態', '資料狀態', '備註', '參考模型代碼', '旋轉角']

    return []


# ============================================================
# 3. format_coord_string - 座標格式化
# ============================================================

def format_coord_string(x, y, z, d, geom_type, target=TARGET_TAOYUAN,
                        elevation_mode=ELEVATION_ABSOLUTE):
    """
    產生單點座標字串。
    桃園市：X Y Z d（四維）。
    國土署/南科：依 elevation_mode 決定 Z 值（三維）。
    """
    try:
        xf = float(x)
        yf = float(y)
        zf = float(z)
        df = float(d) if d else 0.0
    except (ValueError, TypeError):
        return None

    if target == TARGET_TAOYUAN:
        return f'{xf} {yf} {zf} {df}'
    else:
        if elevation_mode == ELEVATION_GROUND:
            return f'{xf} {yf} {(zf - df):.3f}'
        return f'{xf} {yf} {zf:.3f}'


# ============================================================
# 4. build_coordinates - 從分類表資料列建立座標字串
# ============================================================

def _safe_float(val):
    if val is None:
        return None
    s = str(val).strip()
    if s == '':
        return None
    try:
        v = float(s)
        return None if math.isnan(v) else v
    except (ValueError, TypeError):
        return None


def build_coordinates_pipeline(row_dict, target=TARGET_TAOYUAN,
                               elevation_mode=ELEVATION_ABSOLUTE):
    """
    管線：從分類表列的 第N點X/Y/Z/d 欄位組合座標字串。
    回傳座標字串或 None。
    """
    coords = []
    n = 1
    while True:
        x = _safe_float(row_dict.get(f'第{n}點X'))
        y = _safe_float(row_dict.get(f'第{n}點Y'))
        z = _safe_float(row_dict.get(f'第{n}點Z'))
        d_val = _safe_float(row_dict.get(f'第{n}點d'))
        if x is None or y is None:
            break
        if z is None:
            z = 0.0
        if d_val is None:
            d_val = 0.0
        cs = format_coord_string(x, y, z, d_val, 'LineString', target, elevation_mode)
        if cs:
            coords.append(cs)
        n += 1
    return ' '.join(coords) if coords else None


def build_coordinates_point(row_dict, target=TARGET_TAOYUAN,
                            elevation_mode=ELEVATION_ABSOLUTE):
    """
    點狀設施：從 X/Y/Z/d 欄位組合座標字串。
    回傳座標字串或 None。
    """
    x = _safe_float(row_dict.get('X'))
    y = _safe_float(row_dict.get('Y'))
    z = _safe_float(row_dict.get('Z'))
    d_val = _safe_float(row_dict.get('d') or row_dict.get('埋設深度') or row_dict.get('孔深'))
    if x is None or y is None:
        return None
    if z is None:
        z = 0.0
    if d_val is None:
        d_val = 0.0
    cs = format_coord_string(x, y, z, d_val, 'Point', target, elevation_mode)
    return cs


# ============================================================
# 5. _smart_fill_missing - GML 必填欄位智能補值
# ============================================================

def _smart_fill_missing(field, value, mid1):
    """依 JS 版邏輯，對缺漏的必填欄位進行智能補值。"""
    s = str(value or '').strip()
    is_missing = (s == '' or s.lower() == 'empty')

    if is_missing:
        if field in ('輸送物質', '內容物'):
            if mid1 == '03':
                return '自來水'
            elif mid1 == '04':
                return '汙水'
            elif mid1 == '05':
                return '瓦斯'
            elif mid1 == '07':
                return '石油'
        if field == '壓力區分':
            return '0'
    return value


# ============================================================
# 6. generate_gml - 主函式：產生 GML XML 文字
# ============================================================

def generate_gml(cls_df, target=TARGET_TAOYUAN, facility_type='管線',
                 elevation_mode=ELEVATION_ABSOLUTE):
    """
    從分類表 DataFrame 產生 GML XML。

    參數:
        cls_df: build_classification_df 產生的分類表（管線為寬表，點位型為一列一點）
        target: 目標規範（TARGET_TAOYUAN, TARGET_NLMA, TARGET_STEC）
        facility_type: 設施類型
        elevation_mode: 高程模式

    回傳:
        (gml_text, success_count, error_list)
    """
    if cls_df is None or cls_df.empty:
        return ('', 0, [{'error': '無資料可產製 GML'}])

    is_pipeline = (facility_type == '管線')
    success_count = 0
    error_list = []

    # 建構 XML
    schema_loc = ('http://standards.moi.gov.tw/schema/utilityex utilityex.xsd'
                  if target == TARGET_TAOYUAN else
                  'https://standards.moi.gov.tw/schema/utilityex utilityex.xsd')

    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8"?>')
    lines.append(f'<UTL xmlns:gml="http://www.opengis.net/gml"'
                 f' xmlns:xlink="http://www.w3.org/1999/xlink"'
                 f' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
                 f' xmlns:gco="http://www.isotc211.org/2005/gco"'
                 f' xmlns:gmd="http://www.isotc211.org/2005/gmd"'
                 f' xsi:schemaLocation="{schema_loc}"'
                 f' xmlns="http://standards.moi.gov.tw/schema/utilityex">')

    for idx, row in cls_df.iterrows():
        row_dict = row.to_dict()
        cat_code = str(row_dict.get('類別碼', '') or '').strip()

        if not cat_code or len(cat_code) != 7:
            error_list.append({
                'row': idx,
                'id': row_dict.get('識別碼', ''),
                'error': f'類別碼無效: {cat_code}',
            })
            continue

        tag_name, geom_type = get_feature_config(cat_code, target)
        mid1 = cat_code[1:3]

        # 建構座標
        if is_pipeline:
            coord_str = build_coordinates_pipeline(row_dict, target, elevation_mode)
        else:
            coord_str = build_coordinates_point(row_dict, target, elevation_mode)

        if not coord_str:
            error_list.append({
                'row': idx,
                'id': row_dict.get('識別碼', ''),
                'error': '缺少有效座標',
            })
            continue

        # 取得 GML schema
        feature_schema = get_schema_for_tag(tag_name, cat_code, target)

        # 座標維度
        srs_dim = '3'

        # geometry 節點
        if geom_type == 'LineString':
            pos_tag = 'gml:posList'
        else:
            pos_tag = 'gml:coordinates'

        # 組裝 XML
        lines.append('    <gml:featureMember>')
        lines.append(f'        <{tag_name}>')

        lines.append(f'            <geometry>')
        lines.append(f'                <gml:{geom_type} srsName="EPSG:3826" srsDimension="{srs_dim}">')
        lines.append(f'                    <{pos_tag}>{coord_str}</{pos_tag}>')
        lines.append(f'                </gml:{geom_type}>')
        lines.append(f'            </geometry>')

        for field in feature_schema:
            if target != TARGET_TAOYUAN and field in ('IndexNo', '管道編號'):
                continue

            val = row_dict.get(field)
            if val is None and row_dict.get(f'{field}*') is not None:
                val = row_dict.get(f'{field}*')

            val = _smart_fill_missing(field, val, mid1)

            val_str = str(val).strip() if val is not None else ''
            is_empty = (val_str == '' or val_str.lower() == 'empty')

            if field == '設置日期':
                lines.append(f'            <設置日期>')
                lines.append(f'                <gml:TimeInstant>')
                if not is_empty:
                    lines.append(f'                    <gml:timePosition>{_xml_escape(val_str)}</gml:timePosition>')
                else:
                    lines.append(f'                    <gml:timePosition/>')
                lines.append(f'                </gml:TimeInstant>')
                lines.append(f'            </設置日期>')
            else:
                if not is_empty:
                    lines.append(f'            <{field}>{_xml_escape(val_str)}</{field}>')
                else:
                    lines.append(f'            <{field}/>')

        lines.append(f'        </{tag_name}>')
        lines.append('    </gml:featureMember>')
        success_count += 1

    lines.append('</UTL>')
    gml_text = '\n'.join(lines)

    return (gml_text, success_count, error_list)


def _xml_escape(s):
    """XML 特殊字元跳脫。"""
    return (s.replace('&', '&amp;')
             .replace('<', '&lt;')
             .replace('>', '&gt;')
             .replace('"', '&quot;')
             .replace("'", '&apos;'))


# ============================================================
# 7. export_gml_wide_csv - 產生 GML 專用寬表 CSV
# ============================================================

def export_gml_wide_csv_bytes(cls_df, facility_type='管線'):
    """
    將分類表 DataFrame 輸出為 GML 專用寬表 CSV (BOM UTF-8)。
    管線已是寬表格式（第N點X/Y/Z/d），直接輸出。
    點狀設施直接輸出其欄位。
    """
    if cls_df is None or cls_df.empty:
        return b''
    buf = io.BytesIO()
    buf.write(b'\xef\xbb\xbf')  # BOM
    cls_df.to_csv(buf, index=False, encoding='utf-8')
    buf.seek(0)
    return buf.read()
