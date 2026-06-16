# -*- coding: utf-8 -*-
"""
app.py — GML CSV 點位畫布工具 Streamlit 網頁版
執行方式（於專案根目錄）：
    streamlit run app.py
"""

import io
import os
import sys
import zipfile

import streamlit as st
import pandas as pd
import yaml

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

import csv_loader
import id_parser
import anomaly_detector
import canvas_generator
import excel_exporter


# ── 頁面設定 ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GML CSV 點位畫布工具",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ── 讀取 rules.yaml ───────────────────────────────────────────────────────────
RULES_PATH = os.path.join(os.path.dirname(__file__), "config", "rules.yaml")


@st.cache_resource
def _load_base_rules():
    with open(RULES_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


BASE_RULES = _load_base_rules()


# ── Excel → bytes 輔助 ────────────────────────────────────────────────────────
def _to_excel_bytes(export_fn, *args):
    """呼叫任一 export_* 函式，將結果寫入 BytesIO 並回傳 bytes。"""
    bio = io.BytesIO()
    export_fn(*args, bio)
    return bio.getvalue()


# ── 側邊欄 ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("GML CSV 點位畫布工具")
    st.caption("第一版 + 第二版")
    st.divider()

    uploaded_files = st.file_uploader(
        "上傳 CSV 檔案（可多選）",
        type=["csv"],
        accept_multiple_files=True,
        help="支援 UTF-8、Big5、CP950、GB18030 等常見編碼",
    )

    st.divider()
    st.subheader("重複點位判定門檻")
    xy_tol = st.slider("XY 平面差距 (m)", 0.01, 1.0, 0.30, 0.01)
    z_tol  = st.slider("Z 高程差距 (m)",  0.01, 0.5, 0.10, 0.01)

    st.divider()
    run_btn = st.button(
        "▶ 開始處理",
        type="primary",
        use_container_width=True,
        disabled=not uploaded_files,
    )

    if not uploaded_files:
        st.info("請先上傳 CSV 檔案")


# ── 主標題 ────────────────────────────────────────────────────────────────────
st.title("GML CSV 點位畫布工具")

if not uploaded_files:
    st.markdown("""
### 使用方式

1. 在左側上傳一或多個廠商提供的 CSV 檔案
2. 視需要調整**重複點位判定門檻**
3. 點選「**▶ 開始處理**」

程式將自動：
- 解析識別碼（管線識別碼 / 點號）
- 偵測各類異常（座標空值、格式異常、重複點位等）
- 產生互動式點位畫布
- 輸出各類 Excel 報表（可直接下載，含人工判讀與修正紀錄範本）

> ⚠️ 本工具不會修改任何原始 CSV，所有判讀與修正均由人工確認。
    """)
    st.stop()


# ── Session state：依上傳檔案組合 + 門檻決定是否需要重新處理 ─────────────────
_file_key = (
    tuple(sorted(f.name for f in uploaded_files)),
    round(xy_tol, 2),
    round(z_tol, 2),
)

if st.session_state.get("_file_key") != _file_key:
    # 上傳內容或門檻改變 → 清除舊結果，提示重新執行
    st.session_state.pop("results", None)
    st.session_state["_file_key"] = _file_key

# ── 執行處理管線 ──────────────────────────────────────────────────────────────
if run_btn:
    for f in uploaded_files:
        f.seek(0)

    rules = dict(BASE_RULES)
    rules["duplicate_thresholds"] = {"xy_tolerance_m": xy_tol, "z_tolerance_m": z_tol}

    with st.spinner("處理中，請稍候…"):
        records, warnings = csv_loader.load_from_uploads(uploaded_files, rules)

        if records:
            records = id_parser.parse_records(records, rules)
            records, anomalies, duplicates = anomaly_detector.process_records(records, rules)
        else:
            anomalies, duplicates = [], []

    st.session_state["results"] = (records, warnings, anomalies, duplicates)

# ── 尚未執行 ─────────────────────────────────────────────────────────────────
if "results" not in st.session_state:
    st.info("請點選左側「**▶ 開始處理**」按鈕開始分析")
    st.stop()

records, warnings, anomalies, duplicates = st.session_state["results"]

# ── 讀取警告 ──────────────────────────────────────────────────────────────────
if warnings:
    with st.expander(f"⚠️ 讀取警告（共 {len(warnings)} 則）", expanded=bool(warnings)):
        for w in warnings:
            st.warning(w)

if not records:
    st.error("沒有讀取到任何資料列，請確認 CSV 檔案內容。")
    st.stop()

# ── 摘要指標 ──────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
c1.metric("資料列數",     len(records))
c2.metric("管線數量",     len({r.get("管線識別碼") for r in records if r.get("管線識別碼")}))
c3.metric("異常項目",     len(anomalies))
c4.metric("重複量測比對", len(duplicates))

st.divider()

# ── 預先建立 Plotly Figure（各 tab 共用）────────────────────────────────────
fig = canvas_generator.build_figure(records)

# ── 分頁 ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🗺️ 點位畫布",
    "📋 解析後點位表",
    "⚠️ 異常清單",
    "🔁 重複量測比對表",
    "📝 人工判讀紀錄範本",
    "✏️ 修正紀錄範本",
    "📦 下載全部",
])


# ── Tab 1：點位畫布 ───────────────────────────────────────────────────────────
with tab1:
    st.plotly_chart(fig, use_container_width=True)

    html_buf = io.StringIO()
    fig.write_html(html_buf, include_plotlyjs="cdn", full_html=True)
    st.download_button(
        "⬇️ 下載 點位畫布.html",
        data=html_buf.getvalue().encode("utf-8"),
        file_name="點位畫布.html",
        mime="text/html",
    )


# ── Tab 2：解析後點位表 ───────────────────────────────────────────────────────
with tab2:
    df_parsed = pd.DataFrame([
        {col: r.get(col, "") for col in excel_exporter.PARSED_COLUMNS}
        for r in records
    ])
    st.dataframe(df_parsed, use_container_width=True, height=500)

    st.download_button(
        "⬇️ 下載 解析後點位表.xlsx",
        data=_to_excel_bytes(excel_exporter.export_parsed_records, records),
        file_name="解析後點位表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Tab 3：異常清單 ───────────────────────────────────────────────────────────
with tab3:
    if not anomalies:
        st.success("✅ 未發現任何異常")
    else:
        df_anomalies = pd.DataFrame(anomalies)[anomaly_detector.ANOMALY_COLUMNS]

        def _color_anomaly(row):
            t = str(row.get("異常類型", ""))
            if "重複" in t:
                return ["background-color:#FFF3CD"] * len(row)
            if "格式" in t or "空值" in t:
                return ["background-color:#F8D7DA"] * len(row)
            return ["background-color:#D1ECF1"] * len(row)

        st.dataframe(
            df_anomalies.style.apply(_color_anomaly, axis=1),
            use_container_width=True,
            height=500,
        )
        st.caption("🟡 重複點位　🔴 座標/格式問題　🔵 其他")

        st.download_button(
            "⬇️ 下載 異常清單.xlsx",
            data=_to_excel_bytes(excel_exporter.export_anomalies, anomalies),
            file_name="異常清單.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ── Tab 4：重複量測比對表 ─────────────────────────────────────────────────────
with tab4:
    if not duplicates:
        st.success("✅ 未發現重複量測資料")
    else:
        df_dup = pd.DataFrame(duplicates)[anomaly_detector.DUPLICATE_COLUMNS]

        def _color_dup(row):
            v = str(row.get("判定結果", ""))
            if "差異過大" in v:
                return ["background-color:#F8D7DA"] * len(row)
            if "近似" in v:
                return ["background-color:#FFF3CD"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df_dup.style.apply(_color_dup, axis=1),
            use_container_width=True,
            height=500,
        )
        st.caption("🟡 近似重複　🔴 座標差異過大")

        st.download_button(
            "⬇️ 下載 重複量測比對表.xlsx",
            data=_to_excel_bytes(excel_exporter.export_duplicates, duplicates),
            file_name="重複量測比對表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ── Tab 5：人工判讀紀錄範本 ───────────────────────────────────────────────────
with tab5:
    st.markdown("""
#### 人工判讀紀錄範本

此 Excel 已預填所有點位資料（**灰底**欄位），請在 Excel 中填寫**黃底**欄位：

| 欄位 | 下拉選單選項 |
|---|---|
| **判讀狀態** | 正常 / 待確認 / 需廠商補測 / 排除 / 採用 |
| **問題類型** | 點位偏移 / 點位缺漏 / 點位重複 / 座標異常 / 日期不一致 / 點序錯誤 / 其他 |
| **判讀說明** | 自由填寫 |
| **建議回覆施工廠商內容** | 自由填寫 |
| **確認人員 / 確認日期** | 自由填寫 |

> ⚠️ 請勿修改灰底欄位（來源CSV、原始列號、原始識別碼等資料來源欄）
    """)

    st.download_button(
        "⬇️ 下載 人工判讀紀錄範本.xlsx",
        data=_to_excel_bytes(excel_exporter.export_review_template, records),
        file_name="人工判讀紀錄範本.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Tab 6：修正紀錄範本 ───────────────────────────────────────────────────────
with tab6:
    st.markdown("""
#### 修正紀錄範本（空白範本）

請依判讀結果逐筆填入修正動作：

| 欄位 | 下拉選單選項 / 說明 |
|---|---|
| **動作類型** | 新增 / 刪除 / 採用 / 排除 / 修正 |
| 原始識別碼 / 管線識別碼 / 點號 | 對應原始資料（可從解析後點位表複製） |
| 原X / 原Y / 原Z | 原始座標 |
| 修正後X / 修正後Y / 修正後Z | 修正後座標（刪除 / 排除時留空） |
| 修正原因 | 說明修正依據 |
| **是否納入後續GML** | 是 / 否 |
| 確認人員 / 確認日期 | 修正負責人與完成日期 |

> ⚠️ 本範本為人工登記表，程式不會讀取此檔案自動修改任何資料。
    """)

    st.download_button(
        "⬇️ 下載 修正紀錄範本.xlsx",
        data=_to_excel_bytes(excel_exporter.export_correction_template),
        file_name="修正紀錄範本.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Tab 7：下載全部（ZIP）────────────────────────────────────────────────────
with tab7:
    st.markdown("""
#### 一次下載全部輸出

打包所有輸出為一個 ZIP 檔案，包含：

| 檔案 | 說明 |
|---|---|
| `點位畫布.html` | 互動式 Plotly 畫布 |
| `解析後點位表.xlsx` | 所有點位解析結果 |
| `異常清單.xlsx` | 自動偵測異常 |
| `重複量測比對表.xlsx` | 重複點位兩兩比對 |
| `人工判讀紀錄範本.xlsx` | 預填點位資料的判讀表單 |
| `修正紀錄範本.xlsx` | 空白修正動作登記表 |
    """)

    if st.button("建立 ZIP 包", type="secondary"):
        with st.spinner("打包中…"):
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # HTML
                html_sio = io.StringIO()
                fig.write_html(html_sio, include_plotlyjs="cdn", full_html=True)
                zf.writestr("點位畫布.html", html_sio.getvalue().encode("utf-8"))

                # Excel 檔
                for fname, fn, args in [
                    ("解析後點位表.xlsx",     excel_exporter.export_parsed_records, [records]),
                    ("異常清單.xlsx",         excel_exporter.export_anomalies,      [anomalies]),
                    ("重複量測比對表.xlsx",   excel_exporter.export_duplicates,     [duplicates]),
                    ("人工判讀紀錄範本.xlsx", excel_exporter.export_review_template,[records]),
                    ("修正紀錄範本.xlsx",     excel_exporter.export_correction_template, []),
                ]:
                    bio = io.BytesIO()
                    fn(*args, bio)
                    zf.writestr(fname, bio.getvalue())

        st.download_button(
            "⬇️ 下載 GML輸出.zip",
            data=zip_buf.getvalue(),
            file_name="GML輸出.zip",
            mime="application/zip",
        )
