"""
canvas_generator.py - 產生互動式 HTML 點位畫布（Plotly）。

- 依管線識別碼分組，依點號排序後畫線連點
- 可縮放、平移、hover 顯示詳細資訊
- 異常點位以特殊符號標示
"""

import os
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

# Plotly 顏色序列（超過數量後循環）
COLOR_PALETTE = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
    '#aec7e8', '#ffbb78', '#98df8a', '#ff9896', '#c5b0d5',
    '#c49c94', '#f7b6d2', '#c7c7c7', '#dbdb8d', '#9edae5',
]


def _hover_text(row):
    """產生 hover 時顯示的 HTML 文字。"""
    lines = [
        f'<b>原始識別碼：</b>{row.get("raw_id", "")}',
        f'<b>管線識別碼：</b>{row.get("pipeline_id", "")}',
        f'<b>點號：</b>{row.get("point_no", "")}',
        f'<b>類別碼：</b>{row.get("raw_category", "")}',
        f'<b>X(E)：</b>{row.get("raw_x", "")}',
        f'<b>Y(N)：</b>{row.get("raw_y", "")}',
        f'<b>Z(H)：</b>{row.get("raw_z", "")}',
        f'<b>測量日期：</b>{row.get("raw_date", "")}',
        f'<b>來源 CSV：</b>{row.get("_source_file", "")}',
        f'<b>原始列號：</b>{row.get("_original_row", "")}',
    ]
    return '<br>'.join(lines)


def build_figure(df, anomaly_rows=None):
    """
    建立並回傳 Plotly Figure 物件（不存檔）。
    供 Streamlit 直接使用 st.plotly_chart()，也供 generate_canvas() 呼叫。
    """
    if anomaly_rows is None:
        anomaly_rows = set()

    df_plot = df[df['x_val'].notna() & df['y_val'].notna()].copy()

    if df_plot.empty:
        fig = go.Figure()
        fig.update_layout(title='GML 點位畫布（無有效座標）')
        return fig

    traces = []
    pipeline_ids = sorted(df_plot['pipeline_id'].dropna().unique())

    for i, pid in enumerate(pipeline_ids):
        color = COLOR_PALETTE[i % len(COLOR_PALETTE)]
        grp = df_plot[df_plot['pipeline_id'] == pid].copy()

        grp['_sort_key'] = grp['point_no_int'].apply(
            lambda v: (0, int(v)) if pd.notna(v) else (1, 0)
        )
        grp = grp.sort_values('_sort_key')

        xs = grp['x_val'].tolist()
        ys = grp['y_val'].tolist()
        hover_texts = [_hover_text(r) for _, r in grp.iterrows()]

        normal_mask = [
            (r['_source_file'], r['_original_row']) not in anomaly_rows
            for _, r in grp.iterrows()
        ]
        anomaly_mask = [not m for m in normal_mask]

        traces.append(go.Scatter(
            x=xs, y=ys,
            mode='lines',
            name=str(pid),
            line=dict(color=color, width=1.5),
            showlegend=False,
            hoverinfo='skip',
        ))

        norm_xs = [x for x, m in zip(xs, normal_mask) if m]
        norm_ys = [y for y, m in zip(ys, normal_mask) if m]
        norm_texts = [t for t, m in zip(hover_texts, normal_mask) if m]
        if norm_xs:
            traces.append(go.Scatter(
                x=norm_xs, y=norm_ys,
                mode='markers',
                name=str(pid),
                marker=dict(color=color, size=7, symbol='circle'),
                hovertemplate='%{text}<extra></extra>',
                text=norm_texts,
                legendgroup=str(pid),
                showlegend=True,
            ))

        anom_xs = [x for x, m in zip(xs, anomaly_mask) if m]
        anom_ys = [y for y, m in zip(ys, anomaly_mask) if m]
        anom_texts = [t for t, m in zip(hover_texts, anomaly_mask) if m]
        if anom_xs:
            traces.append(go.Scatter(
                x=anom_xs, y=anom_ys,
                mode='markers',
                name=f'{pid}（異常）',
                marker=dict(color='red', size=10, symbol='x', line=dict(width=2)),
                hovertemplate='⚠ 異常點<br>%{text}<extra></extra>',
                text=anom_texts,
                legendgroup=str(pid),
                showlegend=True,
            ))

    df_no_pid = df_plot[df_plot['pipeline_id'].isna()]
    if not df_no_pid.empty:
        hover_texts = [_hover_text(r) for _, r in df_no_pid.iterrows()]
        traces.append(go.Scatter(
            x=df_no_pid['x_val'].tolist(),
            y=df_no_pid['y_val'].tolist(),
            mode='markers',
            name='（無管線識別碼）',
            marker=dict(color='gray', size=7, symbol='diamond'),
            hovertemplate='%{text}<extra></extra>',
            text=hover_texts,
        ))

    fig = go.Figure(data=traces)
    fig.update_layout(
        title=dict(text='GML 點位畫布', font=dict(size=20)),
        xaxis=dict(
            title='X (E) 座標',
            scaleanchor='y', scaleratio=1,
            showgrid=True, gridcolor='#e0e0e0',
        ),
        yaxis=dict(title='Y (N) 座標', showgrid=True, gridcolor='#e0e0e0'),
        hovermode='closest',
        legend=dict(
            title='管線識別碼<br><i>(點選可顯示/隱藏)</i>',
            bgcolor='rgba(255,255,255,0.8)',
            bordercolor='#cccccc', borderwidth=1,
        ),
        plot_bgcolor='#fafafa',
        paper_bgcolor='white',
        font=dict(family='Microsoft JhengHei, Arial, sans-serif'),
        margin=dict(l=60, r=40, t=60, b=60),
    )

    total_pts = len(df_plot)
    anom_pts = sum(
        1 for _, r in df_plot.iterrows()
        if (r['_source_file'], r['_original_row']) in anomaly_rows
    )
    fig.add_annotation(
        text=(
            f'共 {total_pts} 個有效座標點，{anom_pts} 個異常點（紅 ✕）'
            '｜紅色 ✕ 代表異常，詳見異常清單'
        ),
        xref='paper', yref='paper',
        x=0, y=-0.08,
        showarrow=False,
        font=dict(size=11, color='gray'),
        align='left',
    )
    return fig


def generate_canvas(df, output_path, anomaly_rows=None):
    """建立畫布並儲存為 HTML 檔案（CLI 用）。"""
    if anomaly_rows is None:
        anomaly_rows = set()
    df_plot = df[df['x_val'].notna() & df['y_val'].notna()]
    if df_plot.empty:
        print('  [警告] 無有效座標可繪製，產生空白畫布。')
    fig = build_figure(df, anomaly_rows)
    _save_html(fig, output_path)


def _save_html(fig, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    pio.write_html(
        fig,
        file=output_path,
        include_plotlyjs=True,     # 完全離線可用
        full_html=True,
        auto_open=False,
        config={
            'scrollZoom': True,
            'displayModeBar': True,
            'modeBarButtonsToAdd': ['drawline', 'drawopenpath', 'eraseshape'],
            'toImageButtonOptions': {
                'format': 'png',
                'filename': 'GML_點位畫布',
                'height': 900,
                'width': 1600,
            },
        },
    )
    print(f'  已輸出: {output_path}')
