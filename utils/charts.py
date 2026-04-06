"""
圖表模組 - 使用 Plotly 產生互動式圖表
"""

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd


COLORS = {
    'Alan': '#636EFA',
    'Lydia': '#EF553B',
    '四大超商': '#00CC96',
    '一般': '#AB63FA',
    '總金額': '#FFA15A',
}

MONTH_LABELS = {i: f'{i}月' for i in range(1, 13)}


def monthly_bar_chart(summary_df: pd.DataFrame, year: int) -> go.Figure:
    """每月消費長條圖（by Owner）"""
    if summary_df.empty:
        return _empty_chart("無資料")

    owners = [c for c in summary_df.columns if c not in ['結算月份', '總金額']]

    fig = go.Figure()
    for owner in owners:
        if owner in summary_df.columns:
            fig.add_trace(go.Bar(
                name=owner,
                x=summary_df['結算月份'].map(MONTH_LABELS),
                y=summary_df[owner],
                marker_color=COLORS.get(owner, None),
                text=summary_df[owner].apply(lambda x: f'${x:,.0f}'),
                textposition='auto',
            ))

    fig.update_layout(
        title=f'{year} 年每月信用卡消費（by Owner）',
        xaxis_title='月份',
        yaxis_title='金額 (TWD)',
        barmode='group',
        template='plotly_white',
        height=400,
    )
    return fig


def category_bar_chart(cat_df: pd.DataFrame, year: int) -> go.Figure:
    """每月消費長條圖（by 消費類別）"""
    if cat_df.empty:
        return _empty_chart("無資料")

    categories = [c for c in cat_df.columns if c != '結算月份']

    fig = go.Figure()
    for cat in categories:
        if cat in cat_df.columns:
            fig.add_trace(go.Bar(
                name=cat,
                x=cat_df['結算月份'].map(MONTH_LABELS),
                y=cat_df[cat],
                marker_color=COLORS.get(cat, None),
                text=cat_df[cat].apply(lambda x: f'${x:,.0f}'),
                textposition='auto',
            ))

    fig.update_layout(
        title=f'{year} 年每月消費類別分佈',
        xaxis_title='月份',
        yaxis_title='金額 (TWD)',
        barmode='stack',
        template='plotly_white',
        height=400,
    )
    return fig


def reward_chart(rewards_by_month: dict, year: int) -> go.Figure:
    """回饋金額圖表"""
    if not rewards_by_month:
        return _empty_chart("無資料")

    months = sorted(rewards_by_month.keys())
    conv_rewards = [rewards_by_month[m]['四大超商_回饋'] for m in months]
    gen_rewards = [rewards_by_month[m]['一般_回饋'] for m in months]
    totals = [rewards_by_month[m]['回饋總額'] for m in months]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='四大超商回饋',
        x=[MONTH_LABELS.get(m, str(m)) for m in months],
        y=conv_rewards,
        marker_color=COLORS['四大超商'],
        text=[f'${v:,.0f}' for v in conv_rewards],
        textposition='auto',
    ))
    fig.add_trace(go.Bar(
        name='一般回饋',
        x=[MONTH_LABELS.get(m, str(m)) for m in months],
        y=gen_rewards,
        marker_color=COLORS['一般'],
        text=[f'${v:,.1f}' for v in gen_rewards],
        textposition='auto',
    ))
    fig.add_trace(go.Scatter(
        name='回饋總額',
        x=[MONTH_LABELS.get(m, str(m)) for m in months],
        y=totals,
        mode='lines+markers+text',
        text=[f'${v:,.1f}' for v in totals],
        textposition='top center',
        line=dict(color=COLORS['總金額'], width=2),
    ))

    fig.update_layout(
        title=f'{year} 年信用卡回饋金額',
        xaxis_title='月份',
        yaxis_title='回饋金額 (TWD)',
        barmode='stack',
        template='plotly_white',
        height=400,
    )
    return fig


def owner_pie_chart(df: pd.DataFrame, year: int, month: int = None) -> go.Figure:
    """Owner 佔比圓餅圖"""
    if df.empty:
        return _empty_chart("無資料")

    filtered = df[df['結算年份'] == year]
    if month:
        filtered = filtered[filtered['結算月份'] == month]

    owner_totals = filtered.groupby('Owner')['清算消費金額'].sum().reset_index()

    period = f'{year}年{month}月' if month else f'{year}年'

    fig = px.pie(
        owner_totals,
        values='清算消費金額',
        names='Owner',
        title=f'{period} 消費佔比（by Owner）',
        color='Owner',
        color_discrete_map=COLORS,
        hole=0.4,
    )
    fig.update_traces(
        textinfo='label+percent+value',
        texttemplate='%{label}<br>$%{value:,.0f}<br>(%{percent})',
    )
    fig.update_layout(template='plotly_white', height=350)
    return fig


def category_pie_chart(df: pd.DataFrame, year: int, month: int = None) -> go.Figure:
    """類別佔比圓餅圖"""
    if df.empty:
        return _empty_chart("無資料")

    filtered = df[df['結算年份'] == year]
    if month:
        filtered = filtered[filtered['結算月份'] == month]

    cat_totals = filtered.groupby('消費類別')['清算消費金額'].sum().reset_index()

    period = f'{year}年{month}月' if month else f'{year}年'

    fig = px.pie(
        cat_totals,
        values='清算消費金額',
        names='消費類別',
        title=f'{period} 消費佔比（by 類別）',
        color='消費類別',
        color_discrete_map=COLORS,
        hole=0.4,
    )
    fig.update_traces(
        textinfo='label+percent+value',
        texttemplate='%{label}<br>$%{value:,.0f}<br>(%{percent})',
    )
    fig.update_layout(template='plotly_white', height=350)
    return fig


def trend_line_chart(df: pd.DataFrame, year: int) -> go.Figure:
    """年度消費趨勢折線圖"""
    if df.empty:
        return _empty_chart("無資料")

    filtered = df[df['結算年份'] == year]
    monthly = filtered.groupby('結算月份')['清算消費金額'].sum().reset_index()
    monthly['月份'] = monthly['結算月份'].map(MONTH_LABELS)

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=monthly['月份'],
        y=monthly['清算消費金額'],
        mode='lines+markers+text',
        text=monthly['清算消費金額'].apply(lambda x: f'${x:,.0f}'),
        textposition='top center',
        line=dict(color=COLORS['總金額'], width=3),
        marker=dict(size=10),
    ))

    fig.update_layout(
        title=f'{year} 年月度消費趨勢',
        xaxis_title='月份',
        yaxis_title='金額 (TWD)',
        template='plotly_white',
        height=400,
    )
    return fig


def _empty_chart(msg: str) -> go.Figure:
    """產生空白圖表（帶提示訊息）"""
    fig = go.Figure()
    fig.add_annotation(
        text=msg,
        xref="paper", yref="paper",
        x=0.5, y=0.5,
        showarrow=False,
        font=dict(size=20, color="gray"),
    )
    fig.update_layout(
        template='plotly_white',
        height=300,
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
    )
    return fig
