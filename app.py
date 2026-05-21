# -*- coding: utf-8 -*-
from __future__ import annotations
import datetime
import os
import pathlib
import streamlit as st
try:
    import pandas as pd
except Exception:
    pd = None
import plotly.express as px
import plotly.graph_objects as go
from dashboard_core import parse_sales_workbook, month_shift, rows_to_csv

# 上次数据缓存路径
_CACHE_DIR = pathlib.Path(__file__).parent / '.data_cache'
_CACHE_DIR.mkdir(exist_ok=True)
_CACHE_FILE = _CACHE_DIR / 'last_upload.xlsx'

def _slicer(label, options, key):
    """Multi-select dropdown, default all selected"""
    sk = f'slicer_{key}'
    # 空选项防护
    if not options:
        st.caption(f'{label}: 无可用选项')
        return []
    # 初始化session_state（multiselect会自动管理）
    if sk not in st.session_state:
        st.session_state[sk] = list(options)
    sel = st.multiselect(
        label,
        options=list(options),
        default=st.session_state[sk],
        key=sk
    )
    return list(sel) if sel else list(options)


st.set_page_config(page_title='小豚当家BI看板', layout='wide', initial_sidebar_state='expanded')

CSS = '''
<style>
:root {--navy:#07111f;--blue:#1d4ed8;--cyan:#06b6d4;--green:#22c55e;--orange:#f59e0b;--red:#ef4444;--muted:#64748b;}
.block-container {padding-top: 1.1rem; padding-bottom: 2rem; max-width: 1800px;}
[data-testid="stSidebar"] {background: #081324;}
[data-testid="stSidebar"] * {color: #e5f0ff;}
[data-testid="stFileUploader"] {border: 1px dashed rgba(255,255,255,.35); border-radius: 16px; padding: 8px;}
.hero {border-radius: 28px; padding: 24px 28px; margin-bottom: 16px; color: white; background: radial-gradient(circle at 12% 18%, rgba(35,116,255,.75), transparent 30%), linear-gradient(135deg,#06101f 0%, #0b2242 52%, #123f7f 100%); box-shadow: 0 18px 40px rgba(15,23,42,.18);}
.hero-title {font-size: 31px; font-weight: 900; margin: 0; letter-spacing: .5px;}
.hero-sub {color: #cfe3ff; margin-top: 8px; font-size: 14px;}
.badge {display:inline-block; padding: 5px 10px; border-radius: 999px; background: rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.22); margin-right:8px; font-size:12px;}
.section-title {font-size: 18px; font-weight: 800; margin: 16px 0 8px; color: #0f172a;}
[data-testid="stMetric"] {background: linear-gradient(180deg,#ffffff,#f8fbff); border: 1px solid #e8eef8; padding: 15px 16px; border-radius: 20px; box-shadow: 0 12px 28px rgba(15,23,42,.07);}
[data-testid="stMetricLabel"] {color:#64748b;}
[data-testid="stMetricValue"] {font-size: 25px; font-weight: 900;}
.card-note {font-size: 13px; color:#64748b; margin-top:-4px;}
.stTabs [data-baseweb="tab-list"] {gap: 8px;}
.stTabs [data-baseweb="tab"] {background:#f1f5f9; border-radius:999px; padding: 8px 16px;}
.stTabs [aria-selected="true"] {background:#dbeafe; color:#1d4ed8;}
.comp-card {background:#f0f6ff; border:1px solid #c7d9f5; border-radius:16px; padding:16px; margin-bottom:8px;}
.comp-period {font-size:13px; color:#64748b; margin-bottom:4px;}
.comp-value {font-size:22px; font-weight:900; color:#1d4ed8;}
.delta-up {color:#22c55e; font-weight:700;}
.delta-down {color:#ef4444; font-weight:700;}
.diag-card {border-radius:16px; padding:16px; margin-bottom:12px;}
.diag-warn {background:#fff7ed; border:1px solid #fdba74;}
.diag-ok {background:#f0fdf4; border:1px solid #86efac;}
.diag-danger {background:#fef2f2; border:1px solid #fca5a5;}
.diag-title {font-weight:800; font-size:15px; margin-bottom:4px;}
.diag-body {font-size:13px; color:#374151; line-height:1.7;}
.drill-table {width:100%; border-collapse:collapse; font-size:12.5px; margin-top:8px;}
.drill-table th {background:#f0f6ff; font-weight:700; text-align:left; padding:6px 10px; border-bottom:2px solid #1d4ed8;}
.drill-table td {padding:5px 10px; border-bottom:1px solid #e5e7eb; white-space:nowrap; overflow:hidden; max-width:200px;text-overflow:ellipsis;}
.drill-table tr:hover {background:#f8fbff;}
.action-tag {display:inline-block; padding:3px 10px; border-radius:999px; font-size:11px; font-weight:600; margin:2px 3px 2px 0;}
.tag-p0 {background:#fee2e2;color:#dc2626;border:1px solid #fecaca;}
.tag-p1 {background:#fff7ed;color:#ea580c;border:1px solid #fed7aa;}
.tag-p2 {background:#fefce8;color:#ca8a04;border:1px solid #fde68a;}
.tag-p3 {background:#ecfdf5;color:#059669;border:1px solid #a7f3d0;}
/* 自定义HTML表格样式 */
.styled-table-wrap {overflow-x:auto; border-radius:12px; border:1px solid #e2e8f0; margin-top:6px;}
.styled-table {width:100%; border-collapse:collapse; font-size:12.5px;}
.styled-table thead th {background:#1e293b; color:#fff; font-weight:700; text-align:left; padding:9px 10px; white-space:nowrap; position:sticky; top:0; z-index:1;}
.styled-table tbody td {padding:7px 10px; border-bottom:1px solid #e5e7eb; vertical-align:middle;}
.styled-table tbody tr:hover {background:#eff6ff;}
.styled-table td span {white-space:normal;}
</style>
'''
st.markdown(CSS, unsafe_allow_html=True)
st.markdown('''<div class="hero"><div><span class="badge">影锋BI风格</span><span class="badge">全域电商经营驾驶舱</span><span class="badge">上传即更新</span></div><h1 class="hero-title">小豚当家销售经营BI看板</h1><div class="hero-sub">经营总览 · 时间段对比 · 趋势分析 · 渠道矩阵 · 商品诊断 · 智能诊断, 一页完成日常复盘。</div></div>''', unsafe_allow_html=True)

with st.sidebar:
    st.header('数据源更新')
    uploaded = st.file_uploader('上传最新 Excel 数据源', type=['xlsx'])
    if uploaded is not None:
        _CACHE_FILE.write_bytes(uploaded.getvalue())
        st.caption('✅ 已保存，下次打开自动加载此文件。')
    elif _CACHE_FILE.exists():
        mtime = datetime.datetime.fromtimestamp(_CACHE_FILE.stat().st_mtime)
        st.caption(f'📂 自动加载上次数据（{mtime.strftime("%Y-%m-%d %H:%M")} 更新）')
    else:
        st.caption('上传后自动刷新全页，并记住数据方便下次使用。')
    st.divider()
    st.markdown('**建议数据源工作表**')
    st.caption('天猫数据源 / 京东抖音数据源')
    st.markdown('**核心口径**')
    st.caption('转化率=支付买家数/商品访客数；客单价=支付金额/支付买家数。')

@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes):
    return parse_sales_workbook(file_bytes)

if uploaded is not None:
    _file_bytes = uploaded.getvalue()
elif _CACHE_FILE.exists():
    _file_bytes = _CACHE_FILE.read_bytes()
else:
    _file_bytes = None

if not _file_bytes:
    st.info('请在左侧上传最新 Excel 数据源。上传后，这个页面会直接变成可筛选、可导出的 BI 看板。')
    st.stop()

try:
    with st.spinner('正在解析 Excel 并生成经营看板...'):
        data = load_data(_file_bytes)
except Exception as e:
    st.error(f'解析失败：{e}')
    st.stop()

meta = data['meta']
st.success(f"数据已更新：{meta['dateRange'][0]} 至 {meta['dateRange'][1]}，共 {meta['rows']:,} 行；解析工作表：{'、'.join(meta.get('usedSheets', []))}")

# 全局筛选
fc = st.container(border=True)
with fc:
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        start = st.date_input('开始日期', value=datetime.date.fromisoformat(meta['dateRange'][0]))
    with c2:
        end = st.date_input('结束日期', value=datetime.date.fromisoformat(meta['dateRange'][1]))
    with c3:
        channel = _slicer('渠道', data['filters']['channels'], 'ch')
    with c4:
        store = _slicer('店铺', data['filters']['stores'], 'st', max_show=5)
    with c5:
        category = _slicer('品类', data['filters']['categories'], 'cat', max_show=6)
    with c6:
        model = _slicer('型号', data['filters']['models'], 'mdl', max_show=5)

s = str(start)
e = str(end)

METRICS = ['商品访客数', '商品浏览量', '商品加购人数', '商品加购件数', '支付买家数', '支付件数', '支付金额', '成功退款金额']

def get_period_rows(all_rows, s0: str, e0: str, date_key='日期'):
    """Extract rows from all_rows where date is between s0 and e0 (ignoring global channel/store/category/model filters)"""
    out = []
    for r in all_rows:
        d = r.get(date_key, '')
        if len(d) == 7:
            d = d + '-01'
        if d and s0 <= d <= e0:
            out.append(r)
    return out

def filter_rows(rows, date_key='日期'):
    out = []
    for r in rows:
        d = r.get(date_key, '')
        if len(d) == 7:
            d = d + '-01'
        if d and (d < s or d > e):
            continue
        if channel and r.get('渠道') not in channel:
            continue
        if store and r.get('店铺') not in store:
            continue
        if category and r.get('品类') not in category:
            continue
        if model and r.get('型号') not in model:
            continue
        out.append(r)
    return out

def summarize(rows):
    t = {m: 0.0 for m in METRICS}
    for r in rows:
        for m in METRICS:
            t[m] += float(r.get(m, 0) or 0)
    t['客单价'] = t['支付金额'] / t['支付买家数'] if t['支付买家数'] else 0
    t['支付转化率'] = t['支付买家数'] / t['商品访客数'] if t['商品访客数'] else 0
    t['加购率'] = t['商品加购人数'] / t['商品访客数'] if t['商品访客数'] else 0
    t['退款率'] = t['成功退款金额'] / t['支付金额'] if t['支付金额'] else 0
    return t

def group(rows, key):
    d = {}
    for r in rows:
        k = r.get(key) or '未标注'
        d.setdefault(k, {m: 0.0 for m in METRICS})
        for m in METRICS:
            d[k][m] += float(r.get(m, 0) or 0)
    out = []
    for k, v in d.items():
        v[key] = k
        v['客单价'] = v['支付金额'] / v['支付买家数'] if v['支付买家数'] else 0
        v['支付转化率'] = v['支付买家数'] / v['商品访客数'] if v['商品访客数'] else 0
        v['加购率'] = v['商品加购人数'] / v['商品访客数'] if v['商品访客数'] else 0
        v['退款率'] = v['成功退款金额'] / v['支付金额'] if v['支付金额'] else 0
        out.append(v)
    return sorted(out, key=lambda x: x['支付金额'], reverse=True)

def df(rows):
    if pd is None or not rows:
        return rows
    return pd.DataFrame(rows)

def delta_badge(d):
    if d is None:
        return '--'
    sign = '+' if d >= 0 else ''
    cls = 'delta-up' if d >= 0 else 'delta-down'
    return f'<span class="{cls}">{sign}{d*100:.1f}%</span>'

# ──────────────────────────────────────────────
# 基于选定时间段计算同比/环比
# ──────────────────────────────────────────────
def _period_sum(metric_key, s0, e0, apply_filter=True):
    """Calculate metric sum within [s0,e0]; if apply_filter=True, apply channel/store/category/model filters"""
    rows = []
    for r in data['daily']:
        d = r.get('日期', '')
        if len(d) == 7:
            d = d + '-01'
        if not d or not (s0 <= d <= e0):
            continue
        if apply_filter:
            if channel and r.get('渠道') not in channel:
                continue
            if store and r.get('店铺') not in store:
                continue
            if category and r.get('品类') not in category:
                continue
            if model and r.get('型号') not in model:
                continue
        rows.append(r)
    return summarize(rows)



def period_delta_text(metric_key):
    """Calculate MoM and YoY delta for given metric, return display text"""
    cur_days = (end - start).days + 1
    # MoM: same-length prior period
    mom_end = start - datetime.timedelta(days=1)
    mom_start = mom_end - datetime.timedelta(days=cur_days - 1)
    # 同比: 去年同期
    try:
        yoy_start = start.replace(year=start.year - 1)
    except ValueError:
        yoy_start = start.replace(year=start.year - 1, day=28)
    try:
        yoy_end = end.replace(year=end.year - 1)
    except ValueError:
        yoy_end = end.replace(year=end.year - 1, day=28)

    cur_v = _period_sum(metric_key, s, e)[metric_key]
    mom_v = _period_sum(metric_key, str(mom_start), str(mom_end))[metric_key]
    yoy_v = _period_sum(metric_key, str(yoy_start), str(yoy_end))[metric_key]

    mo = (cur_v - mom_v) / mom_v if mom_v else None
    yy = (cur_v - yoy_v) / yoy_v if yoy_v else None
    a = '--' if mo is None else f'环比 {mo*100:+.1f}%'
    b = '--' if yy is None else f'同比 {yy*100:+.1f}%'
    return f'{a} / {b}'


def _html_table(rows, col_widths=None, height=None):
    """Render dict list as styled HTML table, supports HTML tags inside cells"""
    if not rows:
        return '<div style="color:#94a3b8;padding:10px;">暂无数据</div>'
    cols = list(rows[0].keys())
    w = col_widths or {}
    h = f' style="max-height:{height}px;overflow-y:auto;"' if height else ''
    html = f'<div class="styled-table-wrap"{h}><table class="styled-table"><thead><tr>'
    for c in cols:
        cw = w.get(c, '')
        st = f' style="min-width:{cw}"' if cw else ''
        html += f'<th{st}>{c}</th>'
    html += '</tr></thead><tbody>'
    for i, r in enumerate(rows):
        bg = '#fafafa' if i % 2 == 0 else 'white'
        html += f'<tr style="background:{bg}">'
        for c in cols:
            val = r.get(c, '')
            html += f'<td>{val}</td>'
        html += '</tr>'
    html += '</tbody></table></div>'
    return html


# 当前筛选数据
daily = filter_rows(data['daily'], '日期')
totals = summarize(daily)

# Monthly data (with dimensions, for channel trend)
monthly_raw = data.get('monthly', [])
# all_months: 仅月份维度汇总
all_months = data.get('all_months', [])
mm = {r['月份']: r for r in all_months}

ch_rows = group(daily, '渠道')
cat_rows = group(daily, '品类')
store_rows = group(daily, '店铺')

# ─────────────────────────────────────────────────────────────
# Tab 结构
# ─────────────────────────────────────────────────────────────
tabs = st.tabs(['经营总览', '时间段对比', '趋势分析', '渠道矩阵', '商品诊断', '🔍 智能诊断'])

# ═══════════════════════════════════════════════════════════════
# TAB 1: 经营总览
# ═══════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown('<div class="section-title">经营总览</div>', unsafe_allow_html=True)
    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    k1.metric('支付金额', f"¥{totals['支付金额']:,.0f}", period_delta_text('支付金额'))
    k2.metric('支付件数', f"{totals['支付件数']:,.0f}", period_delta_text('支付件数'))
    k3.metric('支付买家', f"{totals['支付买家数']:,.0f}", period_delta_text('支付买家数'))
    k4.metric('访客数', f"{totals['商品访客数']:,.0f}", period_delta_text('商品访客数'))
    k5.metric('支付转化率', f"{totals['支付转化率']*100:.2f}%", period_delta_text('支付转化率'))
    k6.metric('客单价', f"¥{totals['客单价']:,.0f}", period_delta_text('客单价'))
    k7.metric('退款率', f"{totals['退款率']*100:.2f}%", period_delta_text('退款率'))

    st.markdown('<div class="section-title">全域趋势与结构</div>', unsafe_allow_html=True)
    trend = [{'月份': r['月份'], '支付金额': r['支付金额'], '访客数': r['商品访客数'],
               '支付件数': r['支付件数'], '转化率': r['支付转化率']} for r in all_months]
    a_col, b_col = st.columns([2, 1])
    with a_col:
        fig = go.Figure()
        if trend:
            fig.add_trace(go.Bar(x=[r['月份'] for r in trend], y=[r['支付金额'] for r in trend],
                                  name='支付金额', marker_color='#1d4ed8'))
            fig.add_trace(go.Scatter(x=[r['月份'] for r in trend], y=[r['访客数'] for r in trend],
                                      name='访客数', yaxis='y2', line=dict(color='#06b6d4', width=3)))
            fig.add_trace(go.Scatter(x=[r['月份'] for r in trend], y=[r['支付件数'] for r in trend],
                                      name='支付件数', yaxis='y2', line=dict(color='#22c55e', width=3)))
        fig.update_layout(height=390, template='plotly_white', margin=dict(l=20, r=20, t=35, b=20),
                          legend=dict(orientation='h'), yaxis_title='支付金额',
                          yaxis2=dict(title='流量/销量', overlaying='y', side='right'))
        st.plotly_chart(fig, use_container_width=True)
    with b_col:
        fig = px.pie(df(ch_rows[:8]), names='渠道', values='支付金额', hole=.55,
                     color_discrete_sequence=px.colors.qualitative.Set2)
        fig.update_layout(height=390, margin=dict(l=10, r=10, t=35, b=10), title='渠道销售占比')
        st.plotly_chart(fig, use_container_width=True)

    c_col, d_col, e_col = st.columns(3)
    with c_col:
        fig = px.bar(df(store_rows[:12]), x='支付金额', y='店铺', orientation='h',
                     title='店铺销售排行', color='支付转化率', color_continuous_scale='Blues')
        fig.update_layout(height=430, template='plotly_white', yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig, use_container_width=True)
    with d_col:
        fig = px.bar(df(cat_rows[:12]), x='支付金额', y='品类', orientation='h',
                     title='品类销售排行', color='加购率', color_continuous_scale='Teal')
        fig.update_layout(height=430, template='plotly_white', yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig, use_container_width=True)
    with e_col:
        bubble = [{'品类': r['品类'], '支付金额': r['支付金额'], '访客数': r['商品访客数'],
                   '转化率': r['支付转化率'], '客单价': r['客单价'],
                   'size_val': max(abs(r['支付金额']), 0) or 1} for r in cat_rows[:20]]
        fig = px.scatter(df(bubble), x='访客数', y='转化率', size='size_val',
                         color='品类', hover_data=['客单价', '支付金额'], title='品类流量-转化矩阵')
        fig.update_layout(height=430, template='plotly_white')
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title">导出与留档</div>', unsafe_allow_html=True)
    d1, d2 = st.columns(2)
    with d1:
        comp = []
        for r in reversed(all_months):
            prev = mm.get(month_shift(r['月份'], -1))
            ly = mm.get(month_shift(r['月份'], -12))
            comp.append({
                '月份': r['月份'], '支付金额': round(r['支付金额'], 2), '支付件数': round(r['支付件数'], 0),
                '访客数': round(r['商品访客数'], 0), '转化率': round(r['支付转化率'], 4),
                '金额环比': None if not prev or not prev['支付金额'] else round((r['支付金额'] - prev['支付金额']) / prev['支付金额'], 4),
                '金额同比': None if not ly or not ly['支付金额'] else round((r['支付金额'] - ly['支付金额']) / ly['支付金额'], 4),
            })
        st.download_button('下载月度同比环比 CSV', rows_to_csv(comp, ['月份', '支付金额', '支付件数', '访客数', '转化率', '金额环比', '金额同比']), file_name='monthly_yoy_mom.csv', mime='text/csv')
    with d2:
        st.download_button('下载当前筛选日汇总 CSV', rows_to_csv(daily, ['日期', '渠道', '店铺', '品类', '型号', '支付金额', '支付件数', '商品访客数', '支付转化率', '客单价', '退款率']), file_name='filtered_daily_summary.csv', mime='text/csv')

# ═══════════════════════════════════════════════════════════════
# TAB 2: 时间段对比
# ═══════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<div class="section-title">时间段对比分析</div>', unsafe_allow_html=True)

    comp_mode = st.radio('对比模式', ['本期 vs 上期(环比)', '本期 vs 去年同期(同比)', '自定义时间段对比'], horizontal=True)

    today_s = str(start)
    today_e = str(end)

    if comp_mode == '本期 vs 上期(环比)':
        cur_days = (end - start).days + 1
        b_end = start - datetime.timedelta(days=1)
        b_start = b_end - datetime.timedelta(days=cur_days - 1)
        prev_s = str(b_start)
        prev_e = str(b_end)
        label_a = f'本期 {today_s} ~ {today_e}'
        label_b = f'上期 {prev_s} ~ {prev_e}'
    elif comp_mode == '本期 vs 去年同期(同比)':
        try:
            y_start = start.replace(year=start.year - 1)
        except ValueError:
            y_start = start.replace(year=start.year - 1, day=28)
        try:
            y_end = end.replace(year=end.year - 1)
        except ValueError:
            y_end = end.replace(year=end.year - 1, day=28)
        prev_s = str(y_start)
        prev_e = str(y_end)
        label_a = f'本期 {today_s} ~ {today_e}'
        label_b = f'去年同期 {prev_s} ~ {prev_e}'
    else:
        st.markdown('**选择对比时间段**')
        cp1, cp2, cp3, cp4 = st.columns(4)
        with cp1:
            cmp_start = st.date_input('对比期 开始', value=start - datetime.timedelta(days=30), key='cmp_start')
        with cp2:
            cmp_end = st.date_input('对比期 结束', value=end - datetime.timedelta(days=30), key='cmp_end')
        with cp3:
            st.empty()
        with cp4:
            st.empty()
        prev_s = str(cmp_start)
        prev_e = str(cmp_end)
        label_a = f'A期 {today_s} ~ {today_e}'
        label_b = f'B期 {prev_s} ~ {prev_e}'

    def calc_period_summary(s0, e0):
        rows = []
        for r in data['daily']:
            d = r.get('日期', '')
            if len(d) == 7:
                d = d + '-01'
            if not d or not (s0 <= d <= e0):
                continue
            if channel and r.get('渠道') not in channel:
                continue
            if store and r.get('店铺') not in store:
                continue
            if category and r.get('品类') not in category:
                continue
            if model and r.get('型号') not in model:
                continue
            rows.append(r)
        return summarize(rows)

    cur_sum = calc_period_summary(today_s, today_e)
    prev_sum = calc_period_summary(prev_s, prev_e)

    comp_kpis = [
        ('支付金额', '支付金额', '¥', False),
        ('支付件数', '支付件数', '', False),
        ('支付买家', '支付买家数', '', False),
        ('访客数', '商品访客数', '', False),
        ('转化率', '支付转化率', '', True),
        ('客单价', '客单价', '¥', False),
        ('退款率', '退款率', '', True),
    ]

    st.markdown('---')
    kpi_cols = st.columns(7)
    for idx, (k_name, k_key, prefix, is_pct) in enumerate(comp_kpis):
        cur_v = cur_sum.get(k_key, 0)
        prev_v = prev_sum.get(k_key, 0)
        delta_v = (cur_v - prev_v) / prev_v if prev_v else None
        if is_pct:
            cur_str = f"{cur_v*100:.2f}%"
            prev_str = f"{prev_v*100:.2f}%"
            diff_pp = (cur_v - prev_v) * 100
            sign = '+' if diff_pp >= 0 else ''
            cls = 'delta-up' if diff_pp >= 0 else 'delta-down'
            delta_label = f'<span class="{cls}">{sign}{diff_pp:.2f}pp</span>'
        else:
            cur_str = f"{prefix}{cur_v:,.0f}"
            prev_str = f"{prefix}{prev_v:,.0f}"
            delta_label = delta_badge(delta_v)
        with kpi_cols[idx]:
            st.markdown(
                f'<p style="font-weight:800;color:#1d4ed8;font-size:13px;margin:0 0 6px 0;text-align:center;">{k_name}</p>'
                f'<div class="comp-card" style="padding:10px;"><div class="comp-period">{label_a[:16]}</div><div class="comp-value" style="font-size:18px;">{cur_str}</div></div>'
                f'<div class="comp-card" style="padding:10px;"><div class="comp-period">{label_b[:16]}</div><div class="comp-value" style="font-size:18px;color:#64748b;">{prev_str}</div></div>'
                f'<div class="comp-card" style="padding:10px;background:#f0f9ff;"><div class="comp-period">变化率</div><div style="font-size:16px;font-weight:700;">{delta_label}</div></div>',
                unsafe_allow_html=True
            )

    st.markdown('---')
    st.markdown('<div class="section-title">指标变化详情</div>', unsafe_allow_html=True)
    compare_rows = []
    for k_name, k_key, prefix, is_pct in comp_kpis:
        cur_v = cur_sum.get(k_key, 0)
        prev_v = prev_sum.get(k_key, 0)
        chg = (cur_v - prev_v) / prev_v if prev_v else None
        diff = cur_v - prev_v
        if is_pct:
            cur_str = f"{cur_v*100:.2f}%"
            prev_str = f"{prev_v*100:.2f}%"
            diff_str = f"{diff*100:+.2f}pp"
        else:
            cur_str = f"{prefix}{cur_v:,.0f}"
            prev_str = f"{prefix}{prev_v:,.0f}"
            diff_str = f"{prefix}{diff:+,.0f}"
        compare_rows.append({
            '指标': k_name, '本期数值': cur_str, '对比期数值': prev_str,
            '变化量': diff_str, '变化率(%)': f'{chg*100:+.1f}%' if chg is not None else '--'
        })
    st.dataframe(df(compare_rows), use_container_width=True, hide_index=True)

    st.markdown('---')
    p1, p2 = st.columns(2)
    key_map = {'支付金额': '支付金额', '访客数': '商品访客数', '支付件数': '支付件数', '支付买家': '支付买家数'}
    chart_data = [{'指标': k, '本期': cur_sum.get(v, 0), '对比期': prev_sum.get(v, 0)} for k, v in key_map.items()]
    with p1:
        fig = px.bar(chart_data, x='指标', y=['本期', '对比期'], barmode='group',
                     color_discrete_sequence=['#1d4ed8', '#f59e0b'])
        fig.update_layout(height=350, template='plotly_white', title='核心指标对比', legend_title='时间段')
        st.plotly_chart(fig, use_container_width=True)
    with p2:
        ch_data = []
        for k, v in key_map.items():
            cur_v = cur_sum.get(v, 0)
            prev_v = prev_sum.get(v, 0)
            chg = (cur_v - prev_v) / prev_v if prev_v else 0
            ch_data.append({'指标': k, '变化率': chg})
        colors = ['#22c55e' if x['变化率'] >= 0 else '#ef4444' for x in ch_data]
        fig = go.Figure(go.Bar(x=[x['指标'] for x in ch_data], y=[x['变化率'] for x in ch_data],
                                marker_color=colors))
        fig.update_layout(height=350, template='plotly_white', title='各指标变化率', yaxis_tickformat='.1%')
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('---')
    st.markdown('<div class="section-title">渠道维度对比</div>', unsafe_allow_html=True)
    cur_ch = group(get_period_rows(data['daily'], today_s, today_e), '渠道')
    prev_ch = group(get_period_rows(data['daily'], prev_s, prev_e), '渠道')
    prev_ch_map = {r['渠道']: r for r in prev_ch}
    ch_compare = []
    for r in cur_ch:
        name = r['渠道']
        prev_r = prev_ch_map.get(name, {})
        prev_amt = prev_r.get('支付金额', 0)
        cur_amt = r['支付金额']
        chg = (cur_amt - prev_amt) / prev_amt if prev_amt else None
        ch_compare.append({'渠道': name, '本期金额': f"¥{cur_amt:,.0f}",
                            '对比期金额': f"¥{prev_amt:,.0f}",
                            '变化率(%)': f'{chg*100:+.1f}%' if chg is not None else '--'})
    if ch_compare:
        st.dataframe(df(ch_compare), use_container_width=True, hide_index=True)
    st.download_button('下载时间段对比 CSV', rows_to_csv(compare_rows, ['指标', '本期数值', '对比期数值', '变化量', '变化率(%)']), file_name='period_comparison.csv', mime='text/csv')

# ═══════════════════════════════════════════════════════════════
# TAB 3: 趋势分析
# ═══════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<div class="section-title">趋势分析</div>', unsafe_allow_html=True)
    gran = st.radio('趋势粒度', ['月度', '周度', '日度'], horizontal=True, key='granularity')

    if gran == '月度':
        tr_data = [{'周期': r['月份'], '支付金额': r['支付金额'], '访客数': r['商品访客数'],
                    '支付件数': r['支付件数'], '转化率': r['支付转化率'], '加购率': r['加购率']}
                   for r in all_months]
    elif gran == '周度':
        week_dict = {}   # key: (yr, wk_int)  -> {metrics..., week_start, week_end}
        for r in daily:
            try:
                ds = r.get('日期', '')
                if len(ds) == 7:
                    ds = ds + '-01'
                dt = datetime.date.fromisoformat(ds[:10])
                iso = dt.isocalendar()
                yr, wk = iso[0], iso[1]
                wkey = (yr, wk)
            except Exception:
                continue
            if wkey not in week_dict:
                week_dict[wkey] = {m: 0.0 for m in METRICS}
                week_dict[wkey]['_dates'] = []
            for m in METRICS:
                week_dict[wkey][m] += float(r.get(m, 0) or 0)
            week_dict[wkey]['_dates'].append(dt)
        week_rows = []
        for k in sorted(week_dict.keys()):
            v = week_dict[k]
            dates = v['_dates']
            w_start = min(dates).strftime('%m/%d')
            w_end   = max(dates).strftime('%m/%d')
            label = f'{w_start}-{w_end}'
            byr = v['支付买家数']
            vis = v['商品访客数']
            week_rows.append({
                '周期': label, '支付金额': v['支付金额'], '访客数': vis,
                '支付件数': v['支付件数'],
                '转化率': byr / vis if vis else 0,
                '加购率': v['商品加购人数'] / vis if vis else 0
            })
        tr_data = week_rows
    else:
        tr_data = []
        for r in daily:
            byr = float(r.get('支付买家数', 0) or 0)
            vis = float(r.get('商品访客数', 0) or 0)
            tr_data.append({
                '周期': r.get('日期', ''), '支付金额': float(r.get('支付金额', 0) or 0),
                '访客数': vis, '支付件数': float(r.get('支付件数', 0) or 0),
                '转化率': byr / vis if vis else 0,
                '加购率': float(r.get('商品加购人数', 0) or 0) / vis if vis else 0
            })

    t1, t2 = st.columns(2)
    with t1:
        fig = go.Figure()
        if tr_data:
            fig.add_trace(go.Bar(x=[r['周期'] for r in tr_data], y=[r['支付金额'] for r in tr_data],
                                  name='支付金额', marker_color='#1d4ed8', opacity=0.85))
            fig.add_trace(go.Scatter(x=[r['周期'] for r in tr_data], y=[r['访客数'] for r in tr_data],
                                      name='访客数', yaxis='y2', line=dict(color='#06b6d4', width=2)))
        fig.update_layout(height=350, template='plotly_white', legend=dict(orientation='h'),
                          yaxis_title='支付金额', yaxis2=dict(title='访客数', overlaying='y', side='right'))
        st.plotly_chart(fig, use_container_width=True)
    with t2:
        fig = go.Figure()
        if tr_data:
            fig.add_trace(go.Scatter(x=[r['周期'] for r in tr_data], y=[r['转化率'] * 100 for r in tr_data],
                                      name='支付转化率(%)', line=dict(color='#22c55e', width=2)))
            fig.add_trace(go.Scatter(x=[r['周期'] for r in tr_data], y=[r['加购率'] * 100 for r in tr_data],
                                      name='加购率(%)', line=dict(color='#f59e0b', width=2)))
        fig.update_layout(height=350, template='plotly_white', legend=dict(orientation='h'), yaxis_title='比率(%)')
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('---')
    st.markdown('<div class="section-title">同比趋势叠加（月度）</div>', unsafe_allow_html=True)
    if all_months:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=[r['月份'] for r in all_months], y=[r['支付金额'] for r in all_months],
                              name='本期月度金额', marker_color='#1d4ed8'))
        ly_data = [mm.get(month_shift(r['月份'], -12), {}).get('支付金额', 0) for r in all_months]
        fig.add_trace(go.Scatter(x=[r['月份'] for r in all_months], y=ly_data,
                                  name='去年同期金额', line=dict(color='#f59e0b', width=2, dash='dash')))
        fig.update_layout(height=380, template='plotly_white', legend=dict(orientation='h'))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('---')
    st.markdown('<div class="section-title">周内趋势（每日均值）</div>', unsafe_allow_html=True)
    dow_map = {0: '周一', 1: '周二', 2: '周三', 3: '周四', 4: '周五', 5: '周六', 6: '周日'}
    dow_dict = {v: [] for v in dow_map.values()}
    for r in daily:
        try:
            ds = r.get('日期', '')
            if len(ds) == 7:
                ds = ds + '-01'
            dt = datetime.date.fromisoformat(ds[:10])
            dow_dict[dow_map[dt.weekday()]].append(r)
        except Exception:
            continue
    dow_avg = []
    for dow_name in ['周一', '周二', '周三', '周四', '周五', '周六', '周日']:
        rows = dow_dict[dow_name]
        if rows:
            sv = summarize(rows)
            dow_avg.append({'星期': dow_name, '支付金额': sv['支付金额'],
                             '访客数': sv['商品访客数'], '转化率': sv['支付转化率']})
    if dow_avg:
        fig = px.bar(df(dow_avg), x='星期', y='支付金额', color='转化率',
                     color_continuous_scale='RdYlGn', title='各星期日均支付金额（颜色=转化率）')
        fig.update_layout(height=340, template='plotly_white')
        st.plotly_chart(fig, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# TAB 4: 渠道矩阵
# ═══════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<div class="section-title">渠道矩阵分析</div>', unsafe_allow_html=True)

    # 渠道表现: 基于已过滤 daily 数据
    ch_all = group(daily, '渠道')
    if ch_all:
        ch_display = []
        for r in ch_all:
            ch_name = r['渠道']
            # 计算环比: 取选定时间段等长的上一期
            cur_days = (end - start).days + 1
            mom_end = start - datetime.timedelta(days=1)
            mom_start = mom_end - datetime.timedelta(days=cur_days - 1)
            prev_ch_rows = [x for x in data['daily']
                            if x.get('渠道') == ch_name
                            and str(mom_start) <= x.get('日期', '') <= str(mom_end)]
            prev_ch_amt = sum(float(x.get('支付金额', 0) or 0) for x in prev_ch_rows)
            mo_chg = (r['支付金额'] - prev_ch_amt) / prev_ch_amt if prev_ch_amt else None
            ch_display.append({
                '渠道': ch_name, '支付金额': f"¥{r['支付金额']:,.0f}",
                '支付件数': f"{r['支付件数']:,.0f}", '访客数': f"{r['商品访客数']:,.0f}",
                '转化率': f"{r['支付转化率']*100:.2f}%", '客单价': f"¥{r['客单价']:,.0f}",
                '退款率': f"{r['退款率']*100:.2f}%",
                '环比': f'{mo_chg*100:+.1f}%' if mo_chg is not None else '--',
            })
        st.dataframe(df(ch_display), use_container_width=True, hide_index=True)

    # 渠道可视化
    if ch_all:
        vis1, vis2 = st.columns(2)
        with vis1:
            fig = px.bar(df(ch_all), x='渠道', y='支付金额', color='渠道',
                         title='渠道支付金额', color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_layout(height=340, template='plotly_white', showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        with vis2:
            fig = px.scatter(df(ch_all), x='商品访客数', y='支付转化率', size='支付金额',
                              color='渠道', hover_data=['客单价'],
                              title='渠道流量-转化散点',
                              color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_layout(height=340, template='plotly_white')
            st.plotly_chart(fig, use_container_width=True)

    # 渠道内店铺明细
    st.markdown('---')
    st.markdown('<div class="section-title">渠道内店铺明细</div>', unsafe_allow_html=True)
    store_rows2 = group(daily, '店铺')
    store_display = []
    for r in store_rows2:
        # 找到该店铺对应渠道
        ch_of_store = '未标注'
        for dr in daily:
            if dr.get('店铺') == r['店铺']:
                ch_of_store = dr.get('渠道', '未标注')
                break
        store_display.append({
            '店铺': r['店铺'], '渠道': ch_of_store,
            '支付金额': f"¥{r['支付金额']:,.0f}", '支付件数': f"{r['支付件数']:,.0f}",
            '访客数': f"{r['商品访客数']:,.0f}", '转化率': f"{r['支付转化率']*100:.2f}%",
            '客单价': f"¥{r['客单价']:,.0f}"
        })
    if store_display:
        st.dataframe(df(store_display), use_container_width=True, hide_index=True)
        st.download_button('下载店铺明细 CSV', rows_to_csv(store_display, ['店铺', '渠道', '支付金额', '支付件数', '访客数', '转化率', '客单价']), file_name='channel_store.csv', mime='text/csv')

    # 渠道 × 品类矩阵
    st.markdown('---')
    st.markdown('<div class="section-title">渠道 × 品类 金额矩阵</div>', unsafe_allow_html=True)
    cross = {}
    for r in daily:
        ch = r.get('渠道') or '未标注'
        cat = r.get('品类') or '未标注'
        cross.setdefault(ch, {}).setdefault(cat, 0.0)
        cross[ch][cat] += float(r.get('支付金额', 0) or 0)
    cross_rows = []
    all_cats = sorted({cat for d in cross.values() for cat in d.keys()})
    for ch_key, cats in sorted(cross.items()):
        row = {'渠道': ch_key}
        for cat in all_cats:
            row[cat] = f"¥{cats.get(cat, 0):,.0f}"
        cross_rows.append(row)
    if cross_rows:
        st.dataframe(df(cross_rows), use_container_width=True, hide_index=True)

    # 渠道月度趋势（从 daily 重新按月+渠道汇总）
    st.markdown('---')
    st.markdown('<div class="section-title">渠道月度销售趋势</div>', unsafe_allow_html=True)
    ch_monthly_dict = {}
    for r in data['daily']:
        d = r.get('日期', '')
        if len(d) < 7:
            continue
        mk = d[:7]
        ck = r.get('渠道') or '未标注'
        ch_monthly_dict.setdefault(mk, {})[ck] = ch_monthly_dict.get(mk, {}).get(ck, 0) + float(r.get('支付金额', 0) or 0)
    if ch_monthly_dict:
        all_chs = sorted({ck for mdata in ch_monthly_dict.values() for ck in mdata})
        all_months_sorted = sorted(ch_monthly_dict.keys())
        fig = go.Figure()
        colors = px.colors.qualitative.Set2
        for i, ch_name in enumerate(all_chs):
            fig.add_trace(go.Scatter(
                x=all_months_sorted,
                y=[ch_monthly_dict[m].get(ch_name, 0) for m in all_months_sorted],
                name=ch_name, mode='lines+markers',
                line=dict(color=colors[i % len(colors)], width=2)
            ))
        fig.update_layout(height=360, template='plotly_white',
                          legend=dict(orientation='h', y=-0.18),
                          xaxis_title='月份', yaxis_title='支付金额')
        st.plotly_chart(fig, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# TAB 5: 商品诊断
# ═══════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<div class="section-title">商品诊断</div>', unsafe_allow_html=True)

    products = data.get('products', [])
    styles = data.get('styles', [])
    models_data = data.get('models', [])

    # TOP 商品
    st.markdown('<div class="section-title">TOP 销售额商品</div>', unsafe_allow_html=True)
    prod = [r for r in products if
            (channel == '全部' or r.get('渠道') == channel) and
            (category == '全部' or r.get('品类') == category) and
            (model == '全部' or r.get('型号') == model)]
    if not prod:
        # 降级: 从 daily 按商品聚合
        prod_dict = {}
        for r in daily:
            pname = r.get('商品名称') or r.get('款式') or r.get('型号') or '未知商品'
            pid = r.get('商品ID') or ''
            pkey = (pname[:60], pid, r.get('渠道', ''), r.get('品类', ''), r.get('型号', ''))
            prod_dict.setdefault(pkey, {m: 0.0 for m in METRICS})
            for m in METRICS:
                prod_dict[pkey][m] += float(r.get(m, 0) or 0)
        for (pname, pid, pch, pcat, pmod), vals in prod_dict.items():
            vals['商品名称'] = pname
            vals['商品ID'] = pid
            vals['渠道'] = pch
            vals['品类'] = pcat
            vals['型号'] = pmod
            vals['支付转化率'] = vals['支付买家数'] / vals['商品访客数'] if vals['商品访客数'] else 0
            vals['客单价'] = vals['支付金额'] / vals['支付买家数'] if vals['支付买家数'] else 0
            prod.append(vals)
        prod = sorted(prod, key=lambda x: x['支付金额'], reverse=True)

    prod_display = [{'商品名称': str(r.get('商品名称', ''))[:40],
                     '商品ID': r.get('商品ID', ''),
                     '渠道': r.get('渠道', ''), '品类': r.get('品类', ''), '型号': r.get('型号', ''),
                     '支付金额': f"¥{r.get('支付金额', 0):,.0f}",
                     '支付件数': f"{r.get('支付件数', 0):,.0f}",
                     '访客数': f"{r.get('商品访客数', 0):,.0f}",
                     '转化率': f"{r.get('支付转化率', 0)*100:.2f}%",
                     '客单价': f"¥{r.get('客单价', 0):,.0f}"} for r in prod[:200]]
    if prod_display:
        st.dataframe(df(prod_display), use_container_width=True, hide_index=True, height=420)
        st.download_button('下载 TOP 商品 CSV', rows_to_csv(prod[:200],
            ['商品名称', '商品ID', '渠道', '品类', '型号', '支付金额', '支付件数', '商品访客数', '支付转化率', '客单价']),
            file_name='top_products.csv', mime='text/csv')
    else:
        st.info('暂无商品数据，请检查数据源是否包含商品名称字段。')

    # 款式 & 型号 Tab
    st.markdown('---')
    tab_style, tab_model_tab = st.tabs(['款式分布', '型号分布'])
    with tab_style:
        sty = [r for r in styles if
               (channel == '全部' or r.get('渠道') == channel) and
               (category == '全部' or r.get('品类') == category) and
               (model == '全部' or r.get('型号') == model)]
        if not sty:
            sty = group(daily, '款式') if any(r.get('款式') for r in daily) else []
        sty_display = [{'款式': r.get('款式', ''), '渠道': r.get('渠道', ''),
                         '品类': r.get('品类', ''), '型号': r.get('型号', ''),
                         '支付金额': f"¥{r.get('支付金额', 0):,.0f}",
                         '支付件数': f"{r.get('支付件数', 0):,.0f}",
                         '转化率': f"{r.get('支付转化率', 0)*100:.2f}%",
                         '客单价': f"¥{r.get('客单价', 0):,.0f}"} for r in sty[:300]]
        if sty_display:
            st.dataframe(df(sty_display), use_container_width=True, hide_index=True, height=380)
            st.download_button('下载款式 CSV', rows_to_csv(sty[:300], ['款式', '渠道', '品类', '型号', '支付金额', '支付件数', '支付转化率', '客单价']), file_name='styles.csv', mime='text/csv')
        else:
            st.info('暂无款式数据。')

    with tab_model_tab:
        mdl = [r for r in models_data if
               (channel == '全部' or r.get('渠道') == channel) and
               (category == '全部' or r.get('品类') == category) and
               (model == '全部' or r.get('型号') == model)]
        if not mdl:
            mdl = group(daily, '型号')
        mdl_display = [{'型号': r.get('型号', ''), '渠道': r.get('渠道', ''),
                         '品类': r.get('品类', ''), '店铺': r.get('店铺', ''),
                         '支付金额': f"¥{r.get('支付金额', 0):,.0f}",
                         '支付件数': f"{r.get('支付件数', 0):,.0f}",
                         '转化率': f"{r.get('支付转化率', 0)*100:.2f}%",
                         '客单价': f"¥{r.get('客单价', 0):,.0f}"} for r in mdl[:300]]
        if mdl_display:
            st.dataframe(df(mdl_display), use_container_width=True, hide_index=True, height=380)
            st.download_button('下载型号 CSV', rows_to_csv(mdl[:300], ['型号', '渠道', '品类', '店铺', '支付金额', '支付件数', '支付转化率', '客单价']), file_name='models.csv', mime='text/csv')
        else:
            st.info('暂无型号数据。')

    # 品类 TOP 图 + 下钻
    st.markdown('---')
    st.markdown('<div class="section-title">品类销售额 TOP10</div>', unsafe_allow_html=True)
    if cat_rows:
        fig = px.bar(df(cat_rows[:10]), x='支付金额', y='品类', orientation='h',
                     color='支付转化率', color_continuous_scale='Blues', title='品类排行（颜色=转化率）')
        fig.update_layout(height=380, template='plotly_white', yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('---')
    st.markdown('<div class="section-title">商品层级快速筛选</div>', unsafe_allow_html=True)
    x1, x2, x3 = st.columns(3)
    with x1:
        top_ch = st.selectbox('渠道', ['全部'] + data['filters']['channels'], key='prod_ch')
    with x2:
        top_cat = st.selectbox('品类', ['全部'] + data['filters']['categories'], key='prod_cat')
    with x3:
        top_mdl = st.selectbox('型号', ['全部'] + data['filters']['models'], key='prod_mdl')
    filtered_prod = [r for r in products if
                     (top_ch == '全部' or r.get('渠道') == top_ch) and
                     (top_cat == '全部' or r.get('品类') == top_cat) and
                     (top_mdl == '全部' or r.get('型号') == top_mdl)]
    fp_display = [{'商品名称': str(r.get('商品名称', ''))[:50], '渠道': r.get('渠道', ''),
                   '品类': r.get('品类', ''), '型号': r.get('型号', ''),
                   '支付金额': f"¥{r.get('支付金额', 0):,.0f}",
                   '转化率': f"{r.get('支付转化率', 0)*100:.2f}%",
                   '客单价': f"¥{r.get('客单价', 0):,.0f}"} for r in filtered_prod[:100]]
    if fp_display:
        st.dataframe(df(fp_display), use_container_width=True, hide_index=True, height=350)

# ═══════════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════════
# TAB 6: 智能诊断 V2（多因子归因 + 健康评分 + 可执行措施 + 正向亮点）
# ═══════════════════════════════════════════════════════════════
# 构建筛选标签文字
_filter_parts = []
if channel: _filter_parts.append(f'渠道={"+".join(channel)}')
if store: _filter_parts.append(f'店铺={"+".join(store)}')
if category: _filter_parts.append(f'品类={"+".join(category)}')
if model: _filter_parts.append(f'型号={"+".join(model)}')
_filter_label = ' | '.join(_filter_parts) if _filter_parts else '全域'

with tabs[5]:
    st.markdown('<div class="section-title">🔍 智能问题定位诊断 & 优化措施</div>', unsafe_allow_html=True)
    st.caption(f'诊断区间：{s} ~ {e} | 筛选范围：{_filter_label} | 对比区间：相同天数的上一期')

    # ══════════════════════════════════════
    # A. 核心数据准备（基于当前筛选条件）
    # ══════════════════════════════════════
    cur_days = (end - start).days + 1
    mom_end_dt = start - datetime.timedelta(days=1)
    mom_start_dt = mom_end_dt - datetime.timedelta(days=cur_days - 1)
    prev_s_d = str(mom_start_dt)
    prev_e_d = str(mom_end_dt)

    cur_sum = summarize(daily)

    def _agg_by_dims(rows, dims):
        out = {}
        for r in rows:
            key = tuple(r.get(d) or '未标注' for d in dims)
            out.setdefault(key, {m: 0.0 for m in METRICS})
            for m in METRICS:
                out[key][m] += float(r.get(m, 0) or 0)
        for k, v in out.items():
            v['支付转化率'] = v['支付买家数'] / v['商品访客数'] if v['商品访客数'] else 0
            v['客单价'] = v['支付金额'] / v['支付买家数'] if v['支付买家数'] else 0
            v['退款率'] = v['成功退款金额'] / v['支付金额'] if v['支付金额'] else 0
        return out

    # 关键修改: 使用已筛选的 daily 数据, 而非全量 data['daily']
    cur_rows_all = list(daily)
    prev_rows_all_raw = []
    for r in data['daily']:
        d = r.get('日期', '')
        if len(d) == 7: d = d + '-01'
        if prev_s_d <= d <= prev_e_d: prev_rows_all_raw.append(r)
    # 对比期数据也应用同样的筛选条件
    prev_rows_all = []
    for r in prev_rows_all_raw:
        if channel and r.get('渠道') not in channel:
            continue
        if store and r.get('店铺') not in store:
            continue
        if category and r.get('品类') not in category:
            continue
        if model and r.get('型号') not in model:
            continue
        prev_rows_all.append(r)

    cur_by_model   = _agg_by_dims(cur_rows_all, ['渠道','品类','型号'])
    cur_by_cat     = _agg_by_dims(cur_rows_all, ['渠道','品类'])
    cur_by_channel = _agg_by_dims(cur_rows_all, ['渠道'])
    prev_by_model  = _agg_by_dims(prev_rows_all, ['渠道','品类','型号'])
    prev_by_cat    = _agg_by_dims(prev_rows_all, ['渠道','品类'])
    prev_by_channel= _agg_by_dims(prev_rows_all, ['渠道'])

    prev_sum_all = summarize(prev_rows_all)
    def _global_chg(k):
        c = cur_sum.get(k, 0); p = prev_sum_all.get(k, 0)
        return (c - p) / p if p else None

    gmv_g = _global_chg('支付金额')
    vis_g = _global_chg('商品访客数')
    cvr_g = _global_chg('支付转化率')
    aov_g = _global_chg('客单价')
    ref_g = _global_chg('退款率')

    WARN_T = -0.05; DANGER_T = -0.15

    def _pct(v): return f'{v*100:+.1f}%' if v is not None else '--'

    # ──────────────────────────────────────
    # A0. 总体健康评分 & 一句话结论（新增）
    # ──────────────────────────────────────
    scores = {}
    for name, val in [('GMV',gmv_g),('流量',vis_g),('转化率',cvr_g),('客单价',aov_g),('退款率',ref_g)]:
        if val is None: scores[name] = 100
        elif val > WARN_T: scores[name] = 100
        elif val > DANGER_T: scores[name] = 60 + int((val - WARN_T) / (0 - WARN_T) * 40)
        else: scores[name] = max(0, int(val / DANGER_T * 60))
    if ref_g is not None and ref_g > 0:
        scores['退款率'] = max(0, 100 - abs(ref_g) * 300)
    health_score = sum(scores.values()) / len(scores)

    if health_score >= 90: hv = ('🟢 整体健康', '#22c55e', '各项核心指标表现良好，继续保持现有经营策略。')
    elif health_score >= 70: hv = ('🟡 需要关注', '#f59e0b', f'部分指标出现波动（综合得分{health_score:.0f}/100），建议重点关注下方标红项。')
    elif health_score >= 50: hv = ('🠤 存在风险', '#ef4444', f'多项指标明显下滑（综合得分{health_score:.0f}/100），建议立即执行P0优先级行动。')
    else: hv = ('⚠️ 紧急告警', '#dc2626', f'整体经营状况堪忧（综合得分{health_score:.0f}/100），请优先处理所有P0任务！')

    sc1, sc2, sc3 = st.columns([1, 4, 2])
    with sc1:
        st.metric('健康评分', f'{health_score:.0f}', delta=None,
                  help=f"GMV:{scores['GMV']} | 流量:{scores['流量']} | 转化:{scores['转化率']} | 客单价:{scores['客单价']} | 退款:{scores['退款率']}")
    with sc2:
        st.markdown(f"<div style='background:{hv[1]}15;border-left:4px solid {hv[1]};border-radius:8px;padding:12px 16px;margin-top:16px;'>"
                     f"<strong>{hv[0]}</strong>&nbsp;&nbsp;{hv[2]}</div>", unsafe_allow_html=True)
    with sc3:
        st.metric('异常型号数', f'{len([m for m in cur_by_model.values()])}' if False else '-', delta=None)

    # ══════════════════════════════════════
    # B. 第一层: 全局健康度总览（5卡片 + 归因提示）
    # ══════════════════════════════════════
    s1,s2,s3,s4,s5 = st.columns(5)
    metrics_info = [
        ('💰 支付金额',gmv_g,cur_sum.get('支付金额',0),prev_sum_all.get('支付金额',0),'¥',False),
        ('👁 商品访客数',vis_g,cur_sum.get('商品访客数',0),prev_sum_all.get('商品访客数',0),'',False),
        ('🔄 支付转化率',cvr_g,cur_sum.get('支付转化率',0)*100,prev_sum_all.get('支付转化率',0)*100,'',True),
        ('🎫 客单价',aov_g,cur_sum.get('客单价',0),prev_sum_all.get('客单价',0),'¥',False),
        ('↩️ 退款率',ref_g,cur_sum.get('退款率',0)*100,prev_sum_all.get('退款率',0)*100,'',True),
    ]
    for col,(mname,mch,cv,pv,pre,ispct) in zip([s1,s2,s3,s4,s5],metrics_info):
        with col:
            lvl='ok' if mch is None or mch>WARN_T else ('warn' if mch>DANGER_T else 'danger')
            icon={'danger':'🔴','warn':'🟡','ok':'🟢'}[lvl]
            cv_s=f'{cv:,.0f}' if not ispct else f'{cv:.2f}%'
            pv_s=f'{pv:,.0f}' if not ispct else f'{pv:.2f}%'
            # 确保数值有效, 避免显示 {}
            try:
                _ = float(cv); _ = float(pv)
            except:
                cv_s='-'; pv_s='-'
            ch_s=_pct(mch)
            bg={'danger':'#fef2f2','warn':'#fff7ed','ok':'#f0fdf4'}[lvl]
            border={'danger':'#fca5a5','warn':'#fdba74','ok':'#86efac'}[lvl]
            cause_hint = ''
            if '支付金额' in mname and mch is not None and mch < 0:
                parts=[]
                if vis_g and vis_g<-0.03: parts.append(f'流量{_pct(vis_g)}')
                if cvr_g and cvr_g<-0.02: parts.append(f'转化{_pct(cvr_g)}')
                if aov_g and aov_g<-0.02: parts.append(f'客单价{_pct(aov_g)}')
                cause_hint=f'<div style="font-size:10px;color:#ea580c;margin-top:2px;">主因：{"+".join(parts) if parts else "多因子"}</div>'
            elif '访客' in mname and mch is not None and mch < -0.05:
                cause_hint='<div style="font-size:10px;color:#ea580c;margin-top:2px;">→ 可能原因：推广降权/搜索排名下降/季节性波动</div>'
            elif '转化' in mname and mch is not None and mch < -0.03:
                cause_hint='<div style="font-size:10px;color:#ea580c;margin-top:2px;">→ 可能原因：价格竞争力下降/差评累积/详情页体验差</div>'
            st.markdown(
                f'<div style="background:{bg};border:1px solid {border};border-radius:14px;padding:12px;text-align:center;">'
                f'<div style="font-size:11.5px;color:#64748b;font-weight:700;">{mname}</div>'
                f'<div style="font-size:21px;font-weight:900;color:#0f172a;margin:4px 0;">{pre}{cv_s}</div>'
                f'<div style="font-size:11px;color:#94a3b8;">vs 上期 {pre}{pv_s} ({ch_s})</div>{cause_hint}</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════
    # C. 第二层: 根因下钻 — 型号级定位（增强归因）
    # ══════════════════════════════════════
    st.markdown('<hr style="margin:18px 0;border:none;border-top:1px dashed #cbd5e1;">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🎯 根因定位：下钻到型号级</div>', unsafe_allow_html=True)
    st.caption('展示各维度中「环比下滑最严重」的具体型号。归因列使用多因子加权分析，自动识别主要衰退驱动因素。')

    # --- C1. 各渠道内 GMV 下滑最严重的 TOP 型号 ---
    st.markdown('#### 📍 各渠道内 GMV 下滑最严重的 TOP 型号（含多因子归因）')
    ch_model_issues=[]
    total_cur_gmv = cur_sum.get('支付金额', 1)
    for ch_key,cur_v in cur_by_channel.items():
        ch_name=ch_key[0]
        prev_v=prev_by_channel.get(ch_key,{})
        ch_gmv_c=cur_v.get('支付金额',0); ch_gmv_p=prev_v.get('支付金额',0)
        ch_chg=(ch_gmv_c-ch_gmv_p)/ch_gmv_p if ch_gmv_p else None
        if ch_chg and ch_chg<0:
            worst_models=[]
            for mk_key,mv in cur_by_model.items():
                if mk_key[0]!=ch_name: continue
                pv_m=prev_by_model.get(mk_key,{})
                mc=mv.get('支付金额',0); mp=pv_m.get('支付金额',0)
                m_chg=(mc-mp)/mp if mp else None
                if m_chg is not None and m_chg<0:
                    worst_models.append({
                        '渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                        '本期GMV':mc,'上期GMV':mp,'环比':m_chg,
                        '本期转化率':mv.get('支付转化率',0),'上期转化率':pv_m.get('支付转化率',0),
                        '本期访客':mv.get('商品访客数',0),'上期访客':pv_m.get('商品访客数',0),
                        '本期件数':mv.get('支付件数',0),'上期件数':pv_m.get('支付件数',0),
                        '本期客单价':mv.get('客单价',0),'上期客单价':pv_m.get('客单价',0),
                    })
            worst_models.sort(key=lambda x:x['环比'])
            for wm in worst_models[:3]: ch_model_issues.append(wm)

    if ch_model_issues:
        tbl_rows=[]
        for idx,w in enumerate(ch_model_issues[:25]):
            vis_c=w['本期访客'];vis_p=w['上期访客']
            vis_chg_m=(vis_c-vis_p)/vis_p if vis_p else None
            cvr_c=w['本期转化率'];cvr_p=w['上期转化率']
            cvr_diff=(cvr_c-cvr_p)*100 if cvr_p else None
            aov_c=w['本期客单价'];aov_p=w['上期客单价']
            aov_chg_m=(aov_c-aov_p)/aov_p if aov_p else None
            factors=[]
            if vis_chg_m is not None:
                if vis_chg_m<=-0.20: factors.append(('流量崩塌',f'\u2193{_pct(vis_chg_m)}',3))
                elif vis_chg_m<=-0.10: factors.append(('流量大幅下降',f'\u2193{_pct(vis_chg_m)}',2))
                elif vis_chg_m<=-0.05: factors.append(('流量小幅下滑',f'\u2193{_pct(vis_chg_m)}',1))
            if cvr_diff is not None:
                if cvr_diff<=-5: factors.append(('转化率崩溃',f'\u2193{cvr_diff:.1f}pp',3))
                elif cvr_diff<=-2: factors.append(('转化率明显下降',f'\u2193{cvr_diff:.1f}pp',2))
                elif cvr_diff<=-0.5: factors.append(('转化率微降',f'\u2193{cvr_diff:.1f}pp',1))
            if aov_chg_m is not None:
                if aov_chg_m<=-0.20: factors.append(('客单价暴跌',f'\u2193{_pct(aov_chg_m)}',3))
                elif aov_chg_m<=-0.10: factors.append(('客单价明显下跌',f'\u2193{_pct(aov_chg_m)}',2))
                elif aov_chg_m<=-0.03: factors.append(('客单价微跌',f'\u2193{_pct(aov_chg_m)}',1))
            factors.sort(key=lambda x:x[2], reverse=True)
            if factors:
                reason_parts=[]
                for fname,fdetail,fw in factors[:3]:
                    c='#dc2626' if fw==3 else ('#ea580c' if fw==2 else '#f59e0b')
                    reason_parts.append(f"<span style='color:{c}'>\u25CF {fname}:{fdetail}</span>")
                reason_str=' '.join(reason_parts)
            else:
                reason_str='<span style=\'color:#64748b\'>多因素平稳下滑</span>'
            impact_amt=max(0, w['上期GMV']-w['本期GMV'])
            impact_pct=(impact_amt/total_cur_gmv*100) if total_cur_gmv else 0
            tbl_rows.append({
                '序号':idx+1,
                '严重度':'🔴' if w['环比']<DANGER_T else ('🟠' if w['环比']<WARN_T else '🟡'),
                '渠道':w['渠道'],'品类':w['品类'],'型号':w['型号'],
                '本期GMV':f"\u00A5{w['本期GMV']:,.0f}",'上期GMV':f"\u00A5{w['上期GMV']:,.0f}",
                '环比变化':f"<span style='color:#dc2626;font-weight:700'>{_pct(w['环比'])}</span>",
                '拖累总额':f"\u00A5{-impact_amt:,.0f}({impact_pct:.1f}%)" if impact_amt>0 else '-',
                '归因分析(主因)':reason_str,
                '转化率变化':f"{cvr_diff:+.1f}pp" if cvr_diff is not None else '--',
                '流量变化':_pct(vis_chg_m)})
        st.markdown(_html_table(tbl_rows, height=max(280,min(520,len(tbl_rows)*38+40))), unsafe_allow_html=True)
        top_drag=sorted(ch_model_issues,key=lambda x:x['环比'])[:5]
        drag_summary=' | '.join([f"[{t['型号']}]{_pct(t['环比'])}" for t in top_drag])
        st.caption(f'📌 最大拖累TOP5: {drag_summary}')
    else:
        st.info('\u2705 所有渠道的各型号表现稳定，未发现显著异常下滑。')

    # --- C2. 转化率骤降型号 ---
    st.markdown('#### 📉 转化率骤降型号（降幅>20%且访客>50）')
    cvr_drop_models=[]
    for mk_key,mv in cur_by_model.items():
        pv_m=prev_by_model.get(mk_key,{})
        cvr_c=mv.get('支付转化率',0);cvr_p=pv_m.get('支付转化率',0)
        if cvr_p>=0.005 and cvr_c<cvr_p:
            cvr_drop=(cvr_c-cvr_p)/cvr_p
            if cvr_drop<-0.20 and mv.get('商品访客数',0)>50:
                cvr_drop_models.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '本期转化率':cvr_c*100,'上期转化率':cvr_p*100,'降幅':cvr_drop,
                    '本期访客':mv.get('商品访客数',0),'本期GMV':mv.get('支付金额',0),
                    '本期加购人数':mv.get('商品加购人数',0),'上期加购人数':pv_m.get('商品加购人数',0)})
    cvr_drop_models.sort(key=lambda x:x['降幅'])
    if cvr_drop_models:
        cdr=[]
        for c in cvr_drop_models[:15]:
            cart_c=c.get('本期加购人数',0); cart_p=c.get('上期加购人数',0)
            vis_c=c.get('本期访客',0)
            fn=''
            if cart_c and cart_p and vis_c:
                crc=cart_c/vis_c*100; crp=cart_p/vis_c*100
                if crc<crp-2: fn=f'(加购率也\u2193{crc-crp:.1f}pp\u2192详情页吸引力下降)'
                else: fn=f'(加购率正常\u2192可能为价格/评价/库存因素)'
            cdr.append({'渠道':c['渠道'],'品类':c['品类'],'型号':c['型号'],
                '上期转化':f"{c['上期转化率']:.2f}%",'本期转化':f"{c['本期转化率']:.2f}%",
                '降幅':f"<span style='color:#dc2626;font-weight:700'>{_pct(c['降幅'])}</span>",
                '本期访客':f"{c['本期访客']:,.0f}",'本期GMV':f"\u00A5{c['本期GMV']:,.0f}",
                '漏斗判断':fn})
        st.markdown(_html_table(cdr, height=min(450,len(cdr)*36+40)), unsafe_allow_html=True)
    else:
        st.info('\u2705 未发现转化率骤降型号（阈值：降幅>20%，访客>50）。')

    # --- C3. 客单价明显下跌型号 ---
    st.markdown('#### 💰 客单价明显下跌型号（降幅超10%）')
    aov_drop=[]
    for mk_key,mv in cur_by_model.items():
        pv_m=prev_by_model.get(mk_key,{})
        ac=mv.get('客单价',0);ap=pv_m.get('客单价',0)
        if ap>10 and ac<ap:
            ad=(ac-ap)/ap
            if ad<-0.10 and mv.get('支付件数',0)>10:
                aov_drop.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '上期客单':ap,'本期客单':ac,'降幅':ad,
                    '本期件数':mv.get('支付件数',0),'上期件数':pv_m.get('支付件数',0),
                    '本期GMV':mv.get('支付金额',0)})
    aov_drop.sort(key=lambda x:x['降幅'])
    if aov_drop:
        adr=[]
        for a in aov_drop[:15]:
            pc_c=a.get('本期件数',0);pc_p=a.get('上期件数',0)
            ar=''
            if pc_c>pc_p*1.3: ar='(件数\u2191但均价\u2193\u2192低价SKU占比提升/折扣加大)'
            elif pc_c<pc_p*0.7: ar='(件数\u2193且均价\u2193\u2192高客单SKU销量萎缩)'
            else: ar='(件数持平\u2192直接降价/促销力度加大)'
            adr.append({'渠道':a['渠道'],'品类':a['品类'],'型号':a['型号'],
                '上期客单价':f"\u00A5{a['上期客单']:,.0f}",'本期客单价':f"\u00A5{a['本期客单']:,.0f}",
                '降幅':f"<span style='color:#ea580c;font-weight:700'>{_pct(a['降幅'])}</span>",
                '本期件数':f"{a['本期件数']:,.0f}",'上期件数':f"{a['上期件数']:,.0f}",
                '初步判断':ar})
        st.markdown(_html_table(adr, height=min(420,len(adr)*36+40)), unsafe_allow_html=True)
    else:
        st.info('\u2705 未发现客单价明显下跌型号（阈值：降幅>10%，件数>10）。')

    # --- C4. 爆款断崖掉量 ---
    st.markdown('#### \u26A1 爆款断崖式掉量型号（上期TOP20\u2192本期缩水>30%）')
    prev_top20=sorted(prev_by_model.items(),key=lambda x:x[1].get('支付金额',0),reverse=True)[:20]
    drop_stars=[]
    for mk_key,pv in prev_top20:
        mv=cur_by_model.get(mk_key,{})
        pc=mv.get('支付金额',0);pp=pv.get('支付金额',0)
        if pp>0:
            drop=(pc-pp)/pp
            if drop<-0.30:
                drop_stars.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '上期GMV':pp,'本期GMV':pc,'缩水幅度':drop})
    if drop_stars:
        dsr=[]
        for d in drop_stars[:15]:
            loss=d['上期GMV']-d['本期GMV']
            dsr.append({'渠道':d['渠道'],'品类':d['品类'],'型号':d['型号'],
                '上期GMV':f"\u00A5{d['上期GMV']:,.0f}",'本期GMV':f"\u00A5{d['本期GMV']:,.0f}",
                '缩水':f"<span style='color:#dc2626;font-weight:700'>{_pct(d['缩水幅度'])}</span>",
                '损失金额':f"\u00A5{loss:,.0f}",
                '占上期份额':f"{d['上期GMV']/max(prev_sum_all.get('支付金额',1),1)*100:.1f}%"})
        st.markdown(_html_table(dsr, height=min(400,len(dsr)*36+40)), unsafe_allow_html=True)
    else:
        st.info('\u2705 上期TOP20爆款型号均保持稳定。')

    # --- C5. 新晋增长亮点（新增正向反馈）---
    st.markdown('#### 🌟 新晋增长亮点（增速>50%或新上榜）')
    cur_top20=sorted(cur_by_model.items(),key=lambda x:x[1].get('支付金额',0),reverse=True)[:20]
    rising_stars=[]
    for mk_key,mv in cur_top20:
        pv_m=prev_by_model.get(mk_key,{})
        mc=mv.get('支付金额',0);mp=pv_m.get('支付金额',0)
        if mp>0:
            growth=(mc-mp)/mp
            if growth>0.50 and mc>1000:
                rising_stars.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '上期GMV':mp,'本期GMV':mc,'增速':growth})
        else:
            if mc>2000:
                rising_stars.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '上期GMV':0,'本期GMV':mc,'增速':float('inf')})
    rising_stars.sort(key=lambda x:x['本期GMV'],reverse=True)
    if rising_stars:
        rsr=[]
        for r in rising_stars[:10]:
            sp="<span style='color:#22c55e;font-weight:700'>🚀 新品/爆发</span>" if r['增速']==float('inf') else f"<span style='color:#22c55e;font-weight:700'>+{r['增速']*100:.0f}%</span>"
            rsr.append({'渠道':r['渠道'],'品类':r['品类'],'型号':r['型号'],
                '上期GMV':f"\u00A5{r['上期GMV']:,.0f}" if r['上期GMV']>0 else '新上榜',
                '本期GMV':f"\u00A5{r['本期GMV']:,.0f}",'增速':sp})
        st.markdown(_html_table(rsr, height=min(340,len(rsr)*34+40)), unsafe_allow_html=True)
    else:
        st.info('\u2139\ufe0f 本周期未发现显著增长新星（阈值：增速>50% 或 新上榜且GMV>\u00A52000）。')

    # ══════════════════════════════════════
    # D. 第三层: 具体可执行的优化措施（绑定真实数据值）
    # ══════════════════════════════════════
    st.markdown('<hr style="margin:18px 0;border:none;border-top:1px dashed #cbd5e1;">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🛠️ 具体执行措施清单</div>',unsafe_allow_html=True)
    st.caption(f"以下根据上述诊断结果自动生成（共识别 {len(ch_model_issues)} 个异常型号）。每条措施绑定实际数据值。")

    actions=[]

    def add_action(priority,title,detail,owner,timeline,metric_target):
        actions.append({'p':priority,'t':title,'d':detail,'o':owner,'tl':timeline,'mt':metric_target})

    # ======== GMV 问题 ========
    if gmv_g is not None and gmv_g<0:
        gmv_loss=max(0, prev_sum_all.get('支付金额',0)-cur_sum.get('支付金额',0))
        if vis_g is not None and vis_g<-0.08:
            vis_loss_share = abs(vis_g)/(abs(vis_g)+abs(cvr_g or 0)+abs(aov_g or 0)+0.001)
            add_action('P0','【流量】紧急排查核心渠道流量断崖',
                f'<b>现状：</b>访客从 {prev_sum_all.get("商品访客数",0):,.0f}\u2192{cur_sum.get("商品访客数",0):,.0f}（{_pct(vis_g)}），'
                f'潜在影响GMV约 \u00A5{gmv_loss*vis_loss_share:,.0f}<br><br>'
                f'<b>排查步骤：</b><br>'
                f'\u2460 直通车后台\u2192推广计划列表\u2192筛选近7天展现量\u219330%\u2192检查预算耗尽/质量分下降<br>'
                f'\u2461 生意参谋\u2192市场\u2192搜索分析\u2192核心类目Top10词\u2192对比搜索人气和点击率<br>'
                f'\u2462 直播间：查看回放\u2192流量来源分析\u2192确认付费/免费比例<br>'
                f'\u2463 <b>应急：</b>表现最差的3个计划日预算+50%，观察3天',
                '运营负责人','24小时内',f'访客\u2265{prev_sum_all.get("商品访客数",0)*0.95:,.0f}')
        if cvr_g is not None and cvr_g<-0.08:
            ccvr=cur_sum.get('支付转化率',0)*100; pcvr=prev_sum_all.get('支付转化率',0)*100
            lost_orders=cur_sum.get('商品访客数',0)*(prev_sum_all.get('支付转化率',0)-cur_sum.get('支付转化率',0))
            lost_gmv_val=lost_orders*cur_sum.get('客单价',0)
            add_action('P0','【转化】全店转化率紧急提升行动',
                f'<b>现状：</b>{pcvr:.2f}%\u2192{ccvr:.2f}%（\u2193{pcvr-ccvr:.2f}pp），'
                f'少成交约{lost_orders:,.0f}单，影响\u00A5{lost_gmv_val:,.0f}<br><br>'
                f'<b>执行步骤：</b><br>'
                f'\u2460 从「转化率骤降表」提取全部异常SKU\u2192逐一做首屏3秒测试<br>'
                f'\u2461 导出近90天评价\u2192词频统计\u2192负面Top3\u2192修改FAQ和卖点<br>'
                f'\u2462 平台搜索这些型号\u2192对比竞品前3名\u2192高于竞品8%则设限时9折<br>'
                f'\u2463 检查大促是否刚结束\u2192价格回调导致骤降\u2192延长优惠3天',
                '运营+美工','3天内',f'转化率\u2265{pcvr*0.97:.2f}%')
        if aov_g is not None and aov_g<-0.06:
            add_action('P1','【客单价】高客单价SKU曝光恢复',
                f'<b>现状：</b>\u00A5{prev_sum_all.get("客单价",0):.0f}\u2192\u00A5{cur_sum.get("客单价",0):.0f}（{_pct(aov_g)}）<br><br>'
                f'\u2460 提取客单价前20SKU\u2192核对本周vs上周访客\u2192圈出降幅最大的5个<br>'
                f'\u2461 设置关联推荐：「搭配购买减X」「买二送一」，放在加购区下方<br>'
                f'\u2462 满减门槛：均值\u00A5{cur_sum.get("客单价",0):.0f}\u2192满减线设\u00A5{cur_sum.get("客单价",0)*1.3:,.0f}<br>'
                f'\u2463 直通车\u2192「高消费力」人群溢价+20%',
                '运营','1周内',f'客单价\u2265\u00A5{prev_sum_all.get("客单价",0)*0.97:,.0f}')

    # ======== 渠道级别措施 ========
    if ch_model_issues:
        ch_gmv_changes={}
        for ck,cv_ch in cur_by_channel.items():
            pv_ch=prev_by_channel.get(ck,{})
            cc=cv_ch.get('支付金额',0);pp=pv_ch.get('支付金额',0)
            ch_gmv_changes[ck[0]]=(cc-pp)/pp if pp else None
        sorted_ch=sorted(ch_gmv_changes.items(),key=lambda x:x[1] if x[1] else 0)
        worst_ch=sorted_ch[0] if sorted_ch else (None,None)
        if worst_ch[1] and worst_ch[1]<-0.05:
            ch_nm=worst_ch[0];ch_pct=_pct(worst_ch[1])
            bad_in=[m for m in ch_model_issues if m['渠道']==ch_nm][:3]
            ml=', '.join([f"[{m['型号']}]({m['品类']})" for m in bad_in]) or '\u591A\u4e2a\u578b\u53f7'
            add_action('P0',f'【渠道】{ch_nm}专项整改（GMV{ch_pct}）',
                f'该渠道GMV{ch_pct}，集中在：{ml}<br>'
                f'\u2460 检查上述型号推广状态（\u505c/\u964d\u6743/\u8FDD\u89C4）<br>'
                f'\u2461 DSR\u5206\u6570\uff1aDSR&lt;4.7\u5F71\u54CD\u641C\u7D22\u6743\u91CD<br>'
                f'\u2462 \u6838\u5BF9\u6D3B\u52A8\u62A5\u540D\u60C5\u51B5\uff0C\u91CD\u8981\u4F1A\u573A\u8865\u62A5<br>'
                f'\u2463 \u6296\u97F3\u6E20\u9053\uFF1A\u67E5\u8FD17\u5929\u76F4\u64AD\u65F6\u957F\u548CGMV/\u5C0F\u65F6',
                f'{ch_nm}\u6E20\u9053\u8D1F\u8D23\u4EBA','48\u5C0F\u65F6\u5185',f'{ch_nm} GMV\u73AF\u6BD4\u8F6C\u6B63')

    # ======== 型号级别具体措施（Top3异常型号逐一给方案）=====
    if ch_model_issues:
        top3_bad=ch_model_issues[:3]
        for bm in top3_bad:
            mod=bm['型号'];cat=bm['品类'];ch=bm['渠道']
            chg_pct=_pct(bm['环比'])
            vc=bm['本期访客'];vp=bm['上期访客']
            vm=(vc-vp)/vp if vp else None
            ccr=bm['本期转化率']*100;cpr=bm['上期转化率']*100
            cdr=ccr-cpr
            root='\u6D41\u91CF\u65AD\u5D16' if (vm and vm<-0.15) else ('\u8F6C\u5316\u5931\u6548' if cdr<-2 else '\u590D\u5408\u8870\u9000')
            pri='P0' if bm['环比']<DANGER_T else 'P1'
            loss_amnt=bm['上期GMV']-bm['本期GMV']
            detail=(
                f'<b>数据：</b>\u00A5{bm["\u672C\u671FGMV"]:,.0f} vs \u00A5{bm["\u4E0A\u671FGMV"]:,.0f}'
                f'(\u635F\u5931\u00A5{loss_amnt:,.0f})\uff0c\u8F6C\u5316{ccr:.2f}% vs {cpr:.2f}%<br><br>'
                f'<b>[{root}] \u5B9A\u5411\u65BD\uFF1A</b><br>')
            if root=='\u6D41\u91CF\u65AD\u5D16':
                detail+=(
                    f'\u2460 <b>\u6D41\u91CF\u7AEF</b>\uFF1A\u68C0\u67E5\u8BE5\u578B\u53F7\u5728{ch}\u6E20\u9053\u7684\u641C\u7D22\u6392\u540D\u3001\u4E3B\u56FECTR\uff0C'
                    f'CTR={vc/(vp+0.001)*100:.1f}%({"<3%\u9700\u6362\u4E3B\u56FE" if vp and vc/vp<0.03 else "\u6B63\u5E38"})<br>'
                    f'\u2461 \u68C0\u67E5\u662F\u5426\u6709\u63A8\u5E7F\u8BA1\u5212\u88AB\u7CFB\u7EDF\u9650\u6D41/\u9884\u7B97\u8017\u5C3D<br>'
                    f'\u2462 <b>\u6025\u6551</b>\uFF1A\u4E34\u65F6\u589E\u52A0\u76F4\u901A\u8F66\u65E5\u9884\u7B97+50%\uFF0C\u6301\u7EED3\u5929\u89C2\u5BDF<br>')
            else:
                detail+=(
                    f'\u2460 <b>\u8F6C\u5316\u7AEF</b>\uFF1A\u6253\u5F00\u8BE5\u578B\u53F7\u8BE6\u60C5\u9875\u6A21\u62DF\u4E70\u5BB6\u6D4F\u89C8\u2014\u2014\u9996\u5C4F3\u79D2\u80FD\u5426\u770B\u6E05\u6838\u5FC3\u5356\u70B9\uFF1F<br>'
                    f'\u2461 <b>\u8BC4\u4EF7\u5BA1\u8BA1</b>\uFF1A\u5BFC\u51FA\u8FD160\u5929\u8BC4\u4EF7\u2192\u8BCD\u9891\u7EDF\u8BA1\u2192\u8D1F\u9762Top3\u2192\u4F18\u5316\u8BDD\u672F<br>'
                    f'\u2462 <b>\u4EF7\u683C\u5BF9\u6807</b>\uFF1A\u641C\u7D22\u540C\u6B3E\u7ADE\u54C13\u5BB6\uFF0C'
                    f'{"\u9AD8\u4E8E\u7ADE\u54C18\u5219\u8BBE\u9650\u65F6\u6298" if bm["\u672C\u671F\u5BA2\u5355\u4EF7"] > 0 and bm["\u4E0A\u671F\u5BA2\u5355\u4EF7"] > 0 and bm["\u672C\u671F\u5BA2\u5355\u4EF7"] >= bm["\u4E0A\u671F\u5BA2\u5355\u4EF7"]*1.08 else "\u4EF7\u683C\u57FA\u672C\u5408\u7406"}<br>'
                    f'\u2463 <b>\u5E93\u5B58\u68C0\u67E5</b>\uFF1A\u786E\u8BA4\u8BE5\u578B\u53F7\u65E0\u7F3A\u8D27/\u9884\u552E\u72B6\u6001<br>'
                    f'\u2464 <b>\u5DEE\u8BC4\u5904\u7406</b>\uFF1A\u7B5B\u9009\u51FA\u73B0\u22652\u6B21\u4EE5\u4E0A\u7684\u8D1F\u9762\u6807\u7B7E\u96C6\u4E2D\u5904\u7406')
            add_action(pri,f'[{mod}]({cat}/{ch}){root}-GMV{chg_pct}',
                detail,f'\u8FD0\u8425-{cat}\u7EC4','3-5\u5929\u89C1\u6548',f'{mod} GMV\u73AF\u6BD4\u2191>-5%')

    # ======== 退款率措施 ========
    if ref_g is not None and ref_g>0.05:
        crp=cur_sum.get('退款率',0)*100
        if crp>8 or ref_g>0.10:
            add_action('P1','【\u552E\u540E】\u9000\u6B3E\u7387\u5F02\u5E38\u5347\u9AD8',
                f'\u5F53\u524D{crp:.1f}%\uff0C\u53D8\u5316{_pct(ref_g)}<br>'
                f'\u2460 \u5BFC\u51FA\u8FD130\u5929\u9000\u6B3E\u8BA2\u5355\u6309\u300C\u9000\u6B3E\u539F\u56E0\u300D\u5E01\u5E15\u62C9\u5206析\uFF0CTop3\u539F\u56E0<br>'
                f'\u2461 「\u8D28\u91CF\u95EE\u9898/\u63CF\u8FF0\u4E0D\u7B26」&gt;40%\uFF1A\u8D28\u68C0\u56E2\u961F\u62BD\u68C0<br>'
                f'\u2462 「\u7269\u6D41\u6162/\u7834\u635F」&gt;30%\uFF1A\u6539\u8FDB\u5305\u88C5+\u5207\u6362\u5FEB\u9012<br>'
                f'\u2463 「\u4E0D\u60F3\u8981\u4E86\"\uFF1C\u8BF4\u660E\u8BE6\u60C5\u9875\u8BEF\u5BFC\u4FE1\u606F\uFF0C\u9700\u4FEE\u6B63',
                '\u5BA2\u670D+\u4ED3\u50A8+\u8D28\u68C0','2\u5468\u5185','\u9000\u6B3E\u7387&lt;5%')

    # ======== 常规措施 ========
    add_action('P3','【\u5E38\u89C4】\u6BCF\u5468\u4E00\u4E0A\u5348\u5065\u5EB7\u68C0\u67E5',
        '\u6BCF\u5468\u4E00 10:00\u5B8C\u6210：<br>'
        '\u2460 \u6253\u5F00\u672C\u770B\u677F\u300C\u667A\u80FD\u8BCA\u65AD\u300DTab\u622A\u56FE\u5B58\u6863<br>'
        '\u2461 \u5BF9\u6BD4\u4E0A\u5468\u540C\u671F\u6807\u8BB0\u53D8\u5316\u00B15%<br>'
        '\u2462 \u8FDE\u7EED2\u5468\u540C\u4E00\u6307\u6807\u4E0B\u6ED1\u2192\u4E13\u9879\u4F1A\u8BAE<br>'
        '\u2463 \u68C0\u67E5\u672C\u5468\u5230\u671F\u6D3B\u52A8/\u4F18\u60E0\u5238\u7EED\u671F',
        '\u8FD0\u8425\u8D1F\u8D23\u4EBA','\u6BCF\u5468\u4E00\u56FA\u5B9A','\u5468\u62A5\u5B58\u6863')
    add_action('P3','【\u5E38\u89C4】\u6708\u5EA6\u6E20\u9053ROI\u590D\u76D8',
        '\u6BCF\u67085\u65E5\u524D\u5B8C\u6210\u4E0A\u6708\u5404\u6E20\u9053ROI\uFF1A<br>'
        f'ROI=(\u6E20\u9053\u9500\u552E\u989D-\u9000\u8D27\u989D)/\u6E20\u9053\u63A8\u5E7F\u8D39\u7528<br>'
        '\u2460 &gt;5 \u52A0\u5927\u6295\u5165 / 2-5 \u7EF4\u6301 / &lt;2 \u7F29\u51CF\u6216\u4F18\u5316<br>'
        '\u2461 ROI&lt;2\u8F93\u51FA\u300AXX\u6E20\u9053\u4F18\u5316\u65B9\u6848\u300B<br>',
        '\u8FD0\u8425+\u8D22\u52A1','\u6BCF\u67085\u53F7\u524D','\u5168\u6E20\u9053\u5747ROI&gt;3')

    # ---- 展示 ----
    actions_sorted=sorted(actions,key=lambda x:['P0','P1','P2','P3'].index(x['p']))
    for act in actions_sorted:
        cls={'P0':'tag-p0','P1':'tag-p1','P2':'tag-p2','P3':'tag-p3'}[act['p']]
        tag_html=f"<span class='action-tag {cls}'>{act['p']}</span>"
        exp_title=f"{tag_html} **{act['t']}** <small style='color:#94a3b8;'>| {act['o']} | \u76EE\u6807: {act['mt']} | \u89C1\u6548: {act['tl']}</small>"
        with st.expander(exp_title, expanded=(act['p']=='P0')):
            st.markdown(act['d'], unsafe_allow_html=True)
    if not actions_sorted:
        st.success('\u2705 \u5F53\u524D\u6240\u6709\u6838\u5FC3\u6307\u6807\u5065\u5EB7\uFF0C\u65E0\u989D\u5916\u5E72\u9884\u3002')

    # ══════════════════════════════════════
    # E. 一键下载诊断报告（修复列错乱问题）
    # ══════════════════════════════════════
    st.markdown("<hr style='margin:18px 0;border:none;border-top:1px dashed #cbd5e1;'>", unsafe_allow_html=True)
    rep_header=['\u8BCA\u65AD\u65F6\u95F4','\u8BCA\u65AD\u533A\u95F4','\u5BF9\u6BD4\u533A\u95F4','GMV\u53D8\u5316','\u8BBF\u5BA2\u53D8\u5316',
               '\u8F6C\u5316\u7387\u53D8\u5316','\u5BA2\u5355\u4EF7\u53D8\u5316','\u9000\u6B3E\u7387','\u5065\u5EB7\u8BC4\u5206',
               '\u5F02\u5E38\u578B\u53F7\u6570','\u8F6C\u5316\u9AA4\u964D\u578B\u53F7\u6570','\u7206\u6B3E\u6389\u91CF\u6570','\u65B0\u661F\u589E\u957F\u6570',
               'P0\u4EFB\u52A1\u6570','P1\u4EFB\u52A1\u6570']
    dl_data=[{
        '诊断时间': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), '诊断区间': f'{s}~{e}', '对比区间': f'{prev_s_d}~{prev_e_d}',
        'GMV变化': _pct(gmv_g), '访客变化': _pct(vis_g),
        '转化率变化': f"{(cur_sum.get('支付转化率',0)-prev_sum_all.get('支付转化率',0))*100:+.2f}pp" if prev_sum_all.get('支付转化率',0) else '--',
        '客单价变化': _pct(aov_g), '退款率变化': _pct(ref_g),
        '健康评分': f'{health_score:.0f}/100', '健康结论': hv[0],
        '异常型号数': len(ch_model_issues), '转化骤降数': len(cvr_drop_models),
        '爆款掉量数': len(drop_stars), '新星增长数': len(rising_stars) if 'rising_stars' in dir() else 0,
        'P0任务数': sum(1 for a in actions if a['p']=='P0'), 'P1任务数': sum(1 for a in actions if a['p']=='P1'),
    }]
    for act in actions_sorted:
        dl_data.append({
            '诊断时间': '', '诊断区间': '', '对比区间': '',
            'GMV变化': '', '访客变化': '', '转化率变化': '', '客单价变化': '', '退款率变化': '',
            '健康评分': '', '健康结论': '',
            '异常型号数': '', '转化骤降数': '', '爆款掉量数': '', '新星增长数': '',
            'P0任务数': '', 'P1任务数': '',
            '优先级': act['p'], '措施标题': act['t'], '负责人': act['o'], '见效周期': act['tl'], '量化目标': act['mt'],
        })
    st.download_button('📥 下载完整诊断报告 CSV',
        rows_to_csv(dl_data, list(dl_data[0].keys()) if dl_data else rep_header),
        file_name=f'xiaotunbi_diagnosis_{s.replace("-","")}_{e.replace("-","")}.csv',
        mime='text/csv')

