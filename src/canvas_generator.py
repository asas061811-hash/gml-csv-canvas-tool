# -*- coding: utf-8 -*-
"""
canvas_generator.py
--------------------
使用 Plotly 將點位資料畫成互動式 HTML 畫布。

- 每條管線（管線識別碼）以一條 trace 呈現，依點號排序後以「線+點」連接。
- 滑鼠移到點位上，或點選圖例，可顯示該點的詳細資訊。
- 座標為空值或格式異常（無法轉換為數字）的點位不會畫在圖上，
  但仍會出現在「異常清單」中，供人工另行確認。
"""

from collections import defaultdict

import plotly.graph_objects as go


def _build_hover_text(rec):
    z_val = rec.get("Z")
    z_str = "" if z_val is None else f"{z_val}"
    return (
        f"原始識別碼: {rec.get('原始識別碼', '')}<br>"
        f"管線識別碼: {rec.get('管線識別碼', '')}<br>"
        f"點號: {rec.get('點號', '')}<br>"
        f"類別碼: {rec.get('類別碼', '')}<br>"
        f"X/E: {rec.get('X', '')}<br>"
        f"Y/N: {rec.get('Y', '')}<br>"
        f"Z/H: {z_str}<br>"
        f"測量日期: {rec.get('測量日期', '')}<br>"
        f"來源 CSV: {rec.get('來源CSV', '')}<br>"
        f"原始列號: {rec.get('原始列號', '')}"
    )


def build_figure(records):
    """
    建立 Plotly Figure 並回傳（供 Streamlit 網頁版與 CLI 共用）。
    """
    plottable = [
        r for r in records
        if r.get("X") is not None and r.get("Y") is not None and r.get("管線識別碼")
    ]

    groups = defaultdict(list)
    for rec in plottable:
        groups[rec["管線識別碼"]].append(rec)

    fig = go.Figure()

    if not groups:
        fig.add_annotation(
            text="沒有可繪製的點位資料（座標可能為空值或格式異常，請參考異常清單）",
            xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False,
            font=dict(size=16),
        )
    else:
        for pid in sorted(groups.keys()):
            group = groups[pid]
            group_sorted = sorted(
                group,
                key=lambda r: (r.get("點號") is None, r.get("點號") if r.get("點號") is not None else 0),
            )

            xs = [r["X"] for r in group_sorted]
            ys = [r["Y"] for r in group_sorted]
            hover_texts = [_build_hover_text(r) for r in group_sorted]
            labels = [str(r.get("點號", "")) for r in group_sorted]

            fig.add_trace(go.Scatter(
                x=xs,
                y=ys,
                mode="lines+markers+text",
                name=pid,
                text=labels,
                textposition="top center",
                hovertext=hover_texts,
                hoverinfo="text",
                marker=dict(size=7),
                line=dict(width=2),
            ))

    fig.update_layout(
        title="GML 點位互動式畫布（依管線識別碼分組、依點號排序連線）",
        xaxis_title="X / E",
        yaxis_title="Y / N",
        legend_title="管線識別碼",
        hovermode="closest",
        height=900,
    )
    fig.update_yaxes(scaleanchor="x", scaleratio=1)

    return fig


def generate_canvas(records, output_path):
    """
    產生互動式 HTML 畫布並輸出至 output_path（CLI 用）。
    """
    fig = build_figure(records)
    fig.write_html(output_path, include_plotlyjs="cdn", full_html=True)
