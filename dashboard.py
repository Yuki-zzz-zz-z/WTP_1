"""WWTP operational review. Run: streamlit run dashboard.py

Input: original WWTP_Analytics_Database.xlsx / Daily_KPI, one row per day.
Put the workbook beside this file or in data/, or upload it in the sidebar.
No original workbook is modified. No model forecast or automatic dosing command.
"""
from __future__ import annotations

import hashlib
import html
from io import BytesIO
from pathlib import Path
import math
import os

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

GREEN = '#167455'
COLORS = [GREEN, '#388bb0', '#a3b9a8', '#e5ad51', '#7775a6', '#d07861']
LIMITS = {'EO_COD': 30., 'EO_TN': 10., 'EO_NH3N': 5., 'EO_TP': .5, 'EO_SS': 10.}
SHORT = {'EO_COD': 'COD', 'EO_TN': 'TN', 'EO_NH3N': 'NH₃-N', 'EO_TP': 'TP', 'EO_SS': 'SS'}
CORE = ['EO_Flow', 'AEQ_COD', 'EOA_SS', 'EOA_MLVSS', *LIMITS]
RAW = list(dict.fromkeys(CORE + ['AEQ_TN','AEQ_TP','AEQ_NH3N','AEQ_SS','AEQ_pH',
    'EO_pH','EO_NO3N','EO_BOD5','EO_Conductivity','EOA_SVI','EOA_30min_Settling',
    'SRT','SC_Sludge_Blanket','ES_Flow','ES_SS','ES_SS_Online','RS_Flow','RS_SS',
    'EOA_Recycle_to_DN','DN_Outlet_COD','DN_Outlet_NH3N','DN_Outlet_NO3N',
    'Nutrient_N_kg_d','Nutrient_P_kg_d','Aeration_Inlet_Temperature',
    'DryCake_Weight_t','DryCake_Consistency_pct','Floc_Load_Excel'] +
    [f'PM{i}_{field}' for i in [1,2,3] for field in ['COD','Flow','TN','TP','SS']]))
MIN_HISTORY = 60


def divide(a, b):
    return a / b.where(b > 0)


@st.cache_data(show_spinner=False, max_entries=6)
def load_book(content: bytes):
    """Cache by content, never filename; ignore trailing formula-only rows."""
    book = pd.ExcelFile(BytesIO(content), engine='openpyxl')
    if 'Daily_KPI' not in book.sheet_names:
        raise ValueError('工作簿缺少 Daily_KPI 工作表。请上传原结构的数据表。')
    header = book.parse('Daily_KPI', header=None, nrows=1).iloc[0].dropna().astype(str)
    if header.duplicated().any():
        raise ValueError('Daily_KPI 存在重复列名，请先修正。')
    raw = book.parse('Daily_KPI')
    raw.columns = raw.columns.astype(str).str.strip()
    if raw.columns.duplicated().any() or 'Date' not in raw:
        raise ValueError('需要唯一的 Date 日期列和不重复的字段名。')
    dates = pd.to_datetime(raw.Date, errors='coerce', format='mixed')
    # Explicitly accept Excel serial dates in subsequently uploaded workbooks.
    serial = pd.to_numeric(raw.Date, errors='coerce')
    numeric = raw.Date.map(lambda x: isinstance(x, (int, float)) and not pd.isna(x))
    dates.loc[numeric] = pd.to_datetime(serial.loc[numeric], unit='D', origin='1899-12-30', errors='coerce')
    bad_date = dates.isna() & raw.drop(columns='Date').notna().any(axis=1)
    if bad_date.any():
        raise ValueError(f'有 {int(bad_date.sum())} 行日期无法识别，请先核对。')
    raw = raw.loc[dates.notna()].copy()
    raw.index = pd.DatetimeIndex(dates.loc[dates.notna()]).normalize()
    raw = raw.drop(columns='Date').sort_index()
    if raw.index.duplicated().any():
        raise ValueError('同一天出现多行数据，请确认每日汇总口径后再上传。')
    if not any(c in raw for c in CORE):
        raise ValueError('未识别到有效业务字段，请保留原数据库列名。')
    issues = []
    d = pd.DataFrame(index=raw.index)
    missing = [c for c in RAW if c not in raw]
    for c in RAW:
        values = raw[c] if c in raw else pd.Series(np.nan, index=raw.index)
        nums = pd.to_numeric(values, errors='coerce').astype(float)
        invalid = (values.notna() & nums.isna()) | (nums.notna() & ~np.isfinite(nums)) | nums.lt(0)
        if c in ['EOA_SS', 'EOA_MLVSS', 'SRT']:
            invalid |= nums.eq(0)
        if c.endswith('_pH'):
            invalid |= nums.gt(14)
        if c == 'EOA_30min_Settling':
            invalid |= nums.gt(1000)
        if c == 'DryCake_Consistency_pct':
            invalid |= nums.gt(100)
        # These are data-quality guards, not operating control limits.
        if c == 'SC_Sludge_Blanket':
            invalid |= nums.gt(20)
        if c in ['ES_SS','ES_SS_Online','RS_SS']:
            invalid |= nums.gt(100000)
        for date in d.index[invalid]:
            issues.append({'日期':date,'字段':c,'原值':str(values.loc[date]),'处理':'异常或疑似单位错误，按缺测处理；请核对原表'})
        d[c] = nums.mask(invalid)
    active_fields = [c for c in RAW if c not in ['SRT','EOA_SVI','Floc_Load_Excel']]
    active = d[active_fields].notna().any(axis=1)
    if not active.any():
        raise ValueError('没有有效的流量、负荷、污泥或出水数据。')
    latest, first = d.index[active].max(), d.index[active].min()
    trimmed = int((d.index > latest).sum())
    d = d.loc[first:latest].reindex(pd.date_range(first, latest, name='Date'))
    d['low_flow'] = d.EO_Flow.lt(5000)
    if 'Operating_Status' in raw:
        d['low_flow'] |= raw.Operating_Status.reindex(d.index).eq('Low_Flow_Shutdown_Like')
    invalid_fraction = d.EOA_MLVSS.gt(d.EOA_SS)
    for date in d.index[invalid_fraction]:
        issues.append({'日期':date,'字段':'EOA_MLVSS','原值':str(d.loc[date,'EOA_MLVSS']),
                       '处理':'MLVSS 大于 MLSS，MLVSS及相关计算按缺测处理'})
    d.loc[invalid_fraction, 'EOA_MLVSS'] = np.nan
    params = book.parse('Parameters') if 'Parameters' in book.sheet_names else pd.DataFrame()
    volume = 27200.
    if {'Parameter','Value'}.issubset(params.columns):
        vals = params.loc[params.Parameter.eq('TOTAL_BIO_VOLUME_M3'),'Value']
        if len(vals):
            volume = float(vals.iloc[0])
    if not math.isfinite(volume) or volume <= 0:
        raise ValueError('Parameters 中 TOTAL_BIO_VOLUME_M3 必须为正数。')
    units = book.parse('KPI_Units') if 'KPI_Units' in book.sheet_names else pd.DataFrame()
    return d, pd.DataFrame(issues, columns=['日期','字段','原值','处理']), {
        'latest':latest, 'trimmed':trimmed, 'missing':missing, 'volume':volume,
        'hash':hashlib.sha256(content).hexdigest()[:12], 'units':units}


def calculate(d, volume):
    x = d.copy()
    x['COD_load'] = x.AEQ_COD * x.EO_Flow / 1000
    x['MLSS'] = x.EOA_SS / 1000
    x['MLVSS'] = x.EOA_MLVSS / 1000
    x['fraction'] = divide(x.EOA_MLVSS, x.EOA_SS)
    x['FM'] = divide(x.COD_load, x.MLVSS * volume)
    x['ES_solids'] = x.ES_Flow * x.ES_SS / 1000
    x['EO_solids'] = x.EO_Flow * x.EO_SS / 1000
    x['ES_online_solids'] = x.ES_Flow * x.ES_SS_Online / 1000
    x['total_solids'] = x.ES_solids + x.EO_solids
    x['SRT_balance'] = divide(x.MLSS * volume, x.total_solids)
    # Use paired inventory / output observations in this secondary diagnostic.
    pair = x[['MLSS','total_solids']].dropna().reindex(x.index)
    x['SRT_balance7'] = divide((pair.MLSS*volume).rolling(7,min_periods=4).mean(),
                               pair.total_solids.rolling(7,min_periods=4).mean())
    x['SV30_pct'] = x.EOA_30min_Settling / 10
    x['SVI_calc'] = divide(x.EOA_30min_Settling * 1000, x.EOA_SS)
    x['RAS_pct'] = divide(x.RS_Flow * 100, x.EO_Flow)
    x['Recycle_pct'] = divide(x.EOA_Recycle_to_DN * 100, x.EO_Flow)
    x['Cake_dry'] = x.DryCake_Weight_t*x.DryCake_Consistency_pct/100
    for i in [1,2,3]:
        x[f'PM{i}_load'] = x[f'PM{i}_COD']*x[f'PM{i}_Flow']/1000
    for elem in ['N','P']:
        x[f'{elem}_in'] = x[f'AEQ_T{elem}']*x.EO_Flow/1000
        x[f'{elem}_dose'] = x[f'Nutrient_{elem}_kg_d']
        x[f'{elem}_supply'] = x[f'{elem}_in']+x[f'{elem}_dose']
        x[f'{elem}_ratio'] = divide(x[f'{elem}_supply']*1000,x.COD_load)
    for label, inlet, outlet in [('COD','AEQ_COD','EO_COD'),('TN','AEQ_TN','EO_TN'),('NH3','AEQ_NH3N','EO_NH3N')]:
        x[f'{label}_removal'] = 100*divide(x[inlet]-x[outlet],x[inlet])
    x['complete_eo'] = x[list(LIMITS)].notna().all(axis=1)
    x['pass_eo'] = x.complete_eo & x[list(LIMITS)].le(pd.Series(LIMITS)).all(axis=1)
    x['historical_ok'] = x.pass_eo & ~x.low_flow & x.EO_Flow.notna() & (x.index >= '2024-01-01')
    return x.replace([np.inf,-np.inf],np.nan)


def reference(x, date):
    h = x.loc[(x.index >= date-pd.Timedelta(days=365)) & (x.index < date) & x.historical_ok]
    out = {'hist_n':len(h)}
    for field in ['FM','fraction','N_ratio','P_ratio','SRT','EOA_SVI','SC_Sludge_Blanket']:
        values = h[field].dropna()
        if field in ['FM','fraction']: values = values[values > 0]
        out[field+'_n'] = len(values)
        q = values.quantile([.25,.5,.75]) if len(values) >= MIN_HISTORY else [np.nan]*3
        for key,value in zip(['lo','mid','hi'],q): out[field+'_'+key] = value
    for elem in ['N','P']:
        pairs = h[[f'{elem}_supply','COD_load']].dropna()
        pairs = pairs[pairs.COD_load.gt(0)]
        out[elem+'_weighted'] = pairs[f'{elem}_supply'].sum()*1000/pairs.COD_load.sum() if len(pairs)>=MIN_HISTORY else np.nan
    return out


def add_references(x, dates, volume):
    r = pd.DataFrame([reference(x,t) for t in dates], index=dates)
    r['MLSS_lo'] = x.loc[dates,'COD_load']/(volume*r.fraction_mid*r.FM_hi)
    r['MLSS_hi'] = x.loc[dates,'COD_load']/(volume*r.fraction_mid*r.FM_lo)
    # Current low-flow/shutdown-like conditions are outside the calibration pool.
    invalid = x.loc[dates,'low_flow'] | x.loc[dates,'EO_Flow'].isna()
    r.loc[invalid,['MLSS_lo','MLSS_hi']] = np.nan
    return x.loc[dates].join(r)


def fmt(value, digits=1):
    return '—' if pd.isna(value) or not np.isfinite(value) else f'{value:,.{digits}f}'


def average(s, minimum=.7):
    return s.mean() if len(s) and s.notna().mean() >= minimum else np.nan


def change(now, prev, field):
    a,b = average(now[field]),average(prev[field])
    return (a/b-1)*100 if pd.notna(a) and pd.notna(b) and b>0 else np.nan


def supply_mean(frame, elem):
    p = frame[[f'{elem}_supply','COD_load']].dropna()
    p = p[p.COD_load.gt(0)]
    if len(p) < math.ceil(.7*len(frame)) or p.COD_load.sum()<=0:
        return np.nan,len(p)
    return p[f'{elem}_supply'].sum()*1000/p.COD_load.sum(),len(p)


def position(v,lo,hi):
    if any(pd.isna(q) for q in [v,lo,hi]): return '暂无法比较'
    return '低于历史集中区间' if v<lo else '高于历史集中区间' if v>hi else '处于历史集中区间'


def summarize(now, prev):
    """Deterministic Chinese findings with explicit observation coverage."""
    n=len(now); last=now.iloc[-1]; findings=[]; actions=[]
    complete=int(now.complete_eo.sum()); passed=int(now.pass_eo.sum())
    exceed={SHORT[c]:int(now[c].gt(v).sum()) for c,v in LIMITS.items() if now[c].gt(v).any()}
    if exceed:
        title='出水指标需关注'
        eo='出水超过研究筛选线：'+ '、'.join(f'{k} {v}天' for k,v in exceed.items())+'。'
        actions.append('优先核对出水超线日的化验结果及同期进水、回流和排泥记录。')
    elif complete==n:
        title='出水满足筛选条件，继续关注污泥与营养供应'
        eo=f'出水五项 {passed}/{n} 天满足筛选条件。'
    else:
        title='出水数据不完整，健康状态需补充核实'
        eo=f'出水五项齐全 {complete}/{n} 天，其中 {passed} 天满足筛选条件；其余日期不能作完整判断。'
        actions.append('补齐缺失出水检测后再确认完整运行状态。')
    findings.append(('出水表现',eo+'筛选线用于运行回顾，不等同于排放合规结论。'))
    delta=change(now,prev,'COD_load')
    load=f'COD负荷均值 {fmt(average(now.COD_load),0)} kg/d'
    if pd.notna(delta): load+=f'，较前一等长窗口{("增加" if delta>=0 else "减少")} {abs(delta):.1f}%。'
    else: load+='；有效数据不足，暂不比较变化。'
    findings.append(('负荷变化',load))
    mlss=average(now.MLSS); fm=average(now.FM); srt=average(now.SRT)
    joint=now[['MLSS','MLSS_lo','MLSS_hi']].dropna()
    high=int(joint.MLSS.gt(joint.MLSS_hi).sum()); low=int(joint.MLSS.lt(joint.MLSS_lo).sum())
    sludge=f'MLSS均值 {fmt(mlss,2)} g/L；{len(joint)}/{n} 天可比较参考范围，其中偏高 {high} 天、偏低 {low} 天。'
    fm_state=position(fm,last.FM_lo,last.FM_hi)
    sludge+=f' F/M均值 {fmt(fm,3)}，{fm_state}；SRT均值 {fmt(srt)} d。'
    if pd.notna(fm) and pd.notna(last.FM_lo) and fm<last.FM_lo:
        sludge+='单位MLVSS承担的COD负荷较低。'
    elif pd.notna(fm) and pd.notna(last.FM_hi) and fm>last.FM_hi:
        sludge+='单位MLVSS承担的COD负荷较高。'
    es_delta=change(now,prev,'ES_solids'); mlss_delta=change(now,prev,'MLSS')
    if pd.notna(es_delta): sludge+=f' ES固体排出较前窗{("增加" if es_delta>=0 else "减少")} {abs(es_delta):.1f}%。'
    if pd.notna(mlss_delta) and pd.notna(es_delta) and mlss_delta>10 and es_delta < -10:
        sludge+='污泥浓度上升与排出减少同时出现，排泥变化可能是库存增加的因素，尚不能确定因果。'
    if high or low: actions.append('结合ES固体排出、SRT、SV30及泥层趋势复核污泥状态，再评估排泥调整。')
    findings.append(('污泥状态',sludge))
    nutrients=[]
    for elem in ['N','P']:
        value,count=supply_mean(now,elem)
        state=position(value,last[elem+'_ratio_lo'],last[elem+'_ratio_hi'])
        nutrients.append(f'{elem}总供应 {fmt(value,2)} g/kgCOD，{state}（有效{count}/{n}天）')
    missing_dose=[e for e in ['N','P'] if pd.isna(last[e+'_dose'])]
    if missing_dose:
        nutrients.append('截止日'+'/'.join(missing_dose)+'投加缺测，不能判断该日供应')
        actions.append('补录截止日缺失的营养盐投加量；零投加请明确填写0。')
    findings.append(('营养供应','；'.join(nutrients)+'。窗口加权供应与截止日历史范围比较；偏离不直接表示不足或过量。'))
    settling=[]
    for field,label in [('EOA_SVI','SVI'),('SV30_pct','SV30'),('SC_Sludge_Blanket','泥层高度')]:
        v=change(now,prev,field)
        if pd.notna(v): settling.append(f'{label}较前窗{("上升" if v>=0 else "下降")} {abs(v):.1f}%')
    findings.append(('沉降观察','；'.join(settling)+'。' if settling else '沉降观测不足，暂不判断趋势。'))
    if any(pd.notna(change(now,prev,c)) and change(now,prev,c)>10 for c in ['SV30_pct','SC_Sludge_Blanket']):
        actions.append('SV30或泥层高度上升，建议核对测量并结合出水SS持续观察固液分离。')
    if not actions: actions.append('继续跟踪出水、排泥固体量及营养供应变化，保留每日检测和投加记录。')
    return title,findings,actions


def plot(frame, mapping, unit, title, mode='bar', band=None, line=None, stacked=False):
    fig=go.Figure()
    if band:
        lo,hi=band
        fig.add_trace(go.Scatter(x=frame.index,y=frame[lo],mode='lines',line=dict(width=0),showlegend=False,hoverinfo='skip',connectgaps=False))
        fig.add_trace(go.Scatter(x=frame.index,y=frame[hi],mode='lines',line=dict(width=0),fill='tonexty',
                      fillcolor='rgba(22,116,85,.12)',name='历史参考范围',hovertemplate='%{y:.3f}<extra>参考上界</extra>',connectgaps=False))
    for i,(col,label) in enumerate(mapping.items()):
        trace=go.Bar if mode=='bar' else go.Scatter
        kwargs={'marker_color':COLORS[i%len(COLORS)]} if mode=='bar' else {'line':dict(color=COLORS[i%len(COLORS)],width=2.4),'mode':'lines+markers','connectgaps':False}
        fig.add_trace(trace(x=frame.index,y=frame[col],name=label,hovertemplate='%{x|%Y-%m-%d}<br>%{y:,.3f} '+unit+'<extra>'+label+'</extra>',**kwargs))
    if line is not None:
        fig.add_hline(y=line,line_dash='dot',line_color='#cc8b41',annotation_text='研究筛选线',annotation_position='top left')
    fig.update_layout(title=dict(text=title,font=dict(size=16,color='#243c33')),height=315,
        paper_bgcolor='white',plot_bgcolor='white',font=dict(family='Segoe UI, Microsoft YaHei, sans-serif',size=12,color='#51655c'),
        margin=dict(l=12,r=12,t=50,b=25),legend=dict(orientation='h',y=1.03,x=0,font=dict(size=11)),
        hovermode='x unified',barmode='stack' if stacked else 'group',bargap=.3)
    fig.update_xaxes(tickformat='%m/%d',showgrid=False,zeroline=False)
    fig.update_yaxes(title=unit,gridcolor='#edf1ed',zeroline=False,rangemode='tozero')
    st.plotly_chart(fig,width='stretch',config={'displaylogo':False,'modeBarButtonsToRemove':['select2d','lasso2d']})


def chart_pair(now, specs):
    for start in range(0,len(specs),2):
        for col,spec in zip(st.columns(2),specs[start:start+2]):
            with col:
                with st.container(border=True): plot(now,**spec)


def style():
    st.markdown('''<style>
    .stApp {background:#f4f7f4;color:#233b30;font-family:Inter,'Microsoft YaHei',sans-serif}
    .block-container {padding-top:2rem;padding-bottom:3rem;max-width:1530px}
    [data-testid="stSidebar"] {background:#eaf0eb;border-right:1px solid #dce5dd}
    [data-testid="stHeader"] {background:transparent}
    h1 {font-size:2.2rem!important;letter-spacing:-.04em;font-weight:650!important}
    h2,h3 {letter-spacing:-.02em}
    .eyebrow {font-size:11px;letter-spacing:.18em;color:#648273;font-weight:700;margin-bottom:10px}
    .sub {color:#728078;font-size:13px;margin:0 0 20px}
    .summary {background:#fff;border:1px solid #dde7de;border-left:5px solid #167455;border-radius:12px;padding:22px 26px;margin:5px 0 20px}
    .summary h3 {font-size:19px!important;margin:0 0 12px;padding:0;color:#214335}
    .finding {font-size:13px;line-height:1.8;margin:7px 0;color:#4d6055}
    .finding b {display:inline-block;color:#244d38;min-width:76px}
    .actions {background:#edf4ef;color:#305c45;padding:11px 14px;margin-top:14px;border-radius:6px;font-size:13px;line-height:1.8}
    [data-testid="stMetric"] {background:#fff;border:1px solid #dfe8e0;border-radius:10px;padding:16px}
    [data-testid="stMetricLabel"] {color:#738478;font-size:12px}
    [data-testid="stMetricValue"] {font-size:27px;font-weight:600;color:#234b36}
    [data-testid="stVerticalBlockBorderWrapper"]>div {border-color:#dfe8e0!important;border-radius:10px}
    .stTabs [data-baseweb="tab-list"] {gap:26px;border-bottom:1px solid #dce5dd;margin:16px 0 20px}
    .stTabs [data-baseweb="tab"] {font-size:14px;height:44px}
    .stTabs [aria-selected="true"] {color:#167455}
    .stTabs [data-baseweb="tab-highlight"] {background:#167455}
    div[data-testid="stAlert"] {font-size:13px}
    @media(max-width:800px) {.summary{padding:16px}.block-container{padding:1rem}.finding b{display:block}}
    </style>''',unsafe_allow_html=True)


def main():
    st.set_page_config(page_title='WWTP · 运行健康看板',page_icon='🌿',layout='wide')
    style()
    st.sidebar.markdown('## WWTP ANALYTICS')
    st.sidebar.caption('运行健康 · 数据驱动的日常回顾')
    upload=st.sidebar.file_uploader('更新每日数据',type=['xlsx'],help='上传完整历史数据库，保留 Daily_KPI 工作表和原列名。内容变化即重新计算。')
    default=Path(os.environ.get('WWTP_DATA_PATH',Path(__file__).parent/'WWTP_Analytics_Database.xlsx'))
    if not default.exists(): default=Path(__file__).parent/'data'/'WWTP_Analytics_Database.xlsx'
    if upload:
        content=upload.getvalue();name=upload.name
    elif default.exists():
        content=default.read_bytes();name=default.name
    else:
        st.markdown('<div class="eyebrow">WATER & ENVIRONMENT</div>',unsafe_allow_html=True)
        st.title('WWTP 运行健康看板')
        st.info('请在左侧上传 WWTP_Analytics_Database.xlsx。也可以将文件放在 dashboard.py 同目录或 data 文件夹，刷新后读取。')
        st.stop()
    try:
        with st.spinner('读取最新数据并核对字段…'): d,issues,meta=load_book(content)
    except Exception as exc:
        st.error(f'数据读取未完成：{exc}');st.stop()
    st.sidebar.caption(f'已读取：{name}\n\n数据版本：{meta["hash"]}')
    st.sidebar.button('重新读取本地文件',help='替换本地工作簿后点击。上传文件则重新选择新版文件。')
    st.sidebar.divider()
    window=st.sidebar.radio('分析窗口',[7,10,15,30],format_func=lambda n:f'{n} 天',horizontal=True,key='window')
    follow=st.sidebar.toggle('跟随最新有效日期',value=True)
    if follow: end=meta['latest']
    else:
        end=pd.Timestamp(st.sidebar.date_input('截止日期',value=meta['latest'].date(),min_value=d.index.min().date(),max_value=meta['latest'].date()))
    start=end-pd.Timedelta(days=window-1)
    st.sidebar.caption(f'本窗口  {start:%Y/%m/%d} — {end:%Y/%m/%d}\n\n比较窗口  {start-pd.Timedelta(days=window):%m/%d} — {start-pd.Timedelta(days=1):%m/%d}')
    volume=meta['volume']
    st.sidebar.caption(f'总生化有效容积：{volume:,.0f} m³\n\n参考池：2024年起、评价日前365天内满足五项出水筛选且非低流量的历史日。')
    x=calculate(d,volume)
    dates=pd.date_range(start,end,name='Date')
    all_dates=x.index.union(pd.date_range(start-pd.Timedelta(days=window),end,name='Date'))
    x=x.reindex(all_dates)
    for c in ['complete_eo','pass_eo','historical_ok','low_flow']: x[c]=x[c].fillna(False).astype(bool)
    now=add_references(x,dates,volume)
    prev=x.reindex(pd.date_range(start-pd.Timedelta(days=window),start-pd.Timedelta(days=1),name='Date'))
    title,findings,actions=summarize(now,prev)
    st.markdown('<div class="eyebrow">OPERATIONS INTELLIGENCE / WATER & ENVIRONMENT</div>',unsafe_allow_html=True)
    left,right=st.columns([3,1])
    with left: st.title('WWTP 运行健康看板')
    with right: st.caption(f'分析周期  {window} 天\n\n{start:%Y.%m.%d} — {end:%Y.%m.%d}')
    st.markdown('<p class="sub">污泥状态、处理表现与营养供应的综合回顾</p>',unsafe_allow_html=True)
    last=now.iloc[-1]
    n_value,n_count=supply_mean(now,'N');p_value,p_count=supply_mean(now,'P')
    compact=[('出水与负荷', findings[0][1].split('筛选线用于')[0]+' '+findings[1][1]),
             ('污泥与排泥',f'F/M {fmt(average(now.FM),3)}，{position(average(now.FM),last.FM_lo,last.FM_hi)}；MLSS {fmt(average(now.MLSS),2)} g/L，SRT {fmt(average(now.SRT))} d；ES固体排出较前窗 {fmt(change(now,prev,"ES_solids"))}%。'),
             ('营养供应', f'N / P总供应 {fmt(n_value,2)} / {fmt(p_value,2)} g/kgCOD（N有效{n_count}/{window}天，P有效{p_count}/{window}天）。'+('截止日投加缺测，需补录。' if pd.isna(last.N_dose) or pd.isna(last.P_dose) else '详见供应参考与趋势。'))]
    lines=''.join(f'<div class="finding"><b>{html.escape(k)}</b>{html.escape(v)}</div>' for k,v in compact)
    advice=' '.join(actions[:3])
    st.markdown(f'<div class="summary"><div class="eyebrow">运行健康摘要</div><h3>{html.escape(title)}</h3>{lines}<div class="actions"><b>跟进重点　</b>{html.escape(advice)}</div></div>',unsafe_allow_html=True)
    with st.expander('展开分析依据与完整结论'):
        for label,detail in findings: st.markdown(f'**{label}**　{detail}')
    if end.date()<pd.Timestamp.now().date()-pd.Timedelta(days=3):
        st.caption(f'当前为历史数据回顾：最新有效记录 {meta["latest"]:%Y-%m-%d}。更新数据表后，摘要和图表随数据重算。')
    for col,(field,label,unit,digits) in zip(st.columns(6),[
        ('COD_load','COD负荷','kg/d',0),('MLSS','MLSS','g/L',2),('FM','F/M','kgCOD/kgMLVSS·d',3),
        ('SRT','SRT','d',1),('ES_solids','ES固体排出','kg/d',0),('EO_COD','出水COD','mg/L',1)]):
        v=average(now[field]);delta=change(now,prev,field)
        with col:
            st.metric(label,fmt(v,digits),delta=None if pd.isna(delta) else f'{delta:+.1f}% 较前窗',delta_color='off')
            st.caption(f'{unit} · 窗口均值 · {now[field].count()}/{window}天')
    tabs=st.tabs(['处理表现','污泥与排泥','营养供应 N / P','沉降与回流','数据与方法'])
    with tabs[0]:
        st.caption('每日柱形图与同日进出水表观去除率。各采样点的COD口径以数据库为准。')
        chart_pair(now,[
            dict(mapping={'COD_load':'生化进水负荷'},unit='kg COD/d',title='生化COD负荷'),
            dict(mapping={f'PM{i}_load':f'PM{i}' for i in [1,2,3]},unit='kg COD/d',title='生产线COD来源',stacked=True),
            dict(mapping={f'PM{i}_COD':f'PM{i}' for i in [1,2,3]},unit='mg/L',title='生产线COD浓度'),
            dict(mapping={'AEQ_COD':'AEQ','DN_Outlet_COD':'DN出口','EO_COD':'最终出水'},unit='mg/L',title='沿程COD浓度'),
            dict(mapping={'AEQ_TN':'AEQ TN','EO_TN':'出水TN'},unit='mg/L',title='进出水总氮'),
            dict(mapping={'AEQ_NH3N':'AEQ NH₃-N','EO_NH3N':'出水NH₃-N','EO_NO3N':'出水NO₃-N'},unit='mg/L',title='氮形态变化'),
            dict(mapping={'COD_removal':'COD','TN_removal':'TN','NH3_removal':'NH₃-N'},unit='%',title='同日浓度表观去除率'),
            dict(mapping={'EO_Flow':'出水流量'},unit='m³/d',title='过程流量'),
        ])
        st.caption('PM来源合计与AEQ负荷属于不同采样边界，不直接等同。柱图不填补缺测，组成缺失时不能视为完整总量。浓度去除率未校正水力停留时间；TN未计入过程投加，不作总氮质量平衡。')
        st.subheader('出水观察')
        chart_pair(now,[dict(mapping={c:SHORT[c]},unit='mg/L',title=f'出水 {SHORT[c]}',line=limit) for c,limit in LIMITS.items()])
    with tabs[1]:
        st.info('MLSS范围由当日COD负荷、此前365天F/M P25–P75和历史MLVSS/MLSS中位数换算。范围表示历史运行对应关系，偏离不直接触发排泥。')
        plot(now,{'MLSS':'实测MLSS'},'g/L','MLSS与负荷对应参考范围',mode='line',band=('MLSS_lo','MLSS_hi'))
        last=now.iloc[-1]
        st.caption(f'截止日参考：{fmt(last.MLSS_lo,2)}–{fmt(last.MLSS_hi,2)} g/L；F/M有效历史 {int(last.FM_n)} 天；f中位数 {fmt(last.fraction_mid,3)}。每一天均使用该日之前的历史，最少需{MIN_HISTORY}个有效历史样本。')
        chart_pair(now,[
            dict(mapping={'FM':'重算F/M'},unit='kgCOD/(kgMLVSS·d)',title='F/M与历史集中区间',band=('FM_lo','FM_hi')),
            dict(mapping={'MLSS':'MLSS','MLVSS':'MLVSS'},unit='g/L',title='污泥浓度组成'),
            dict(mapping={'ES_solids':'ES排出','EO_solids':'出水携出'},unit='kg TSS/d',title='系统固体排出',stacked=True),
            dict(mapping={'ES_Flow':'ES流量'},unit='m³/d',title='剩余污泥排出流量'),
            dict(mapping={'ES_SS':'实验室SS','ES_SS_Online':'在线SS'},unit='mg/L',title='ES固体浓度'),
            dict(mapping={'SRT':'数据库SRT','SRT_balance7':'7日TSS平衡诊断'},unit='d',title='SRT与辅助诊断',mode='line'),
        ])
        st.caption('ES固体量 = ES流量 × 实验室SS ÷ 1000。在线SS仅作对照，不自动填补化验缺测。数据库SRT与TSS平衡诊断口径不同。F/M与MLSS共享库存信息，不作为相互独立的证据。')
    with tabs[2]:
        st.subheader('Historical Normal Supply Reference')
        st.caption('历史正常供应参考：总供应 = 进水元素负荷 + 实际元素投加。按同日完整数据计算，缺测不等于零投加。')
        for col,elem in zip(st.columns(2),['N','P']):
            with col:
                last=now.iloc[-1];value,count=supply_mean(now,elem)
                st.metric(f'{elem} 总供应 / COD',fmt(value,2)+' g/kgCOD')
                st.caption(f'截止日前365天 P25–P75：{fmt(last[elem+"_ratio_lo"],2)}–{fmt(last[elem+"_ratio_hi"],2)}；负荷加权中心 {fmt(last[elem+"_weighted"],2)}；历史n={int(last[elem+"_ratio_n"])}')
                plot(now,{elem+'_ratio':'实际总供应'},f'g{elem}/kgCOD',f'{elem}供应与逐日历史参考',band=(elem+'_ratio_lo',elem+'_ratio_hi'))
                paired=now.copy()
                valid=paired[[elem+'_in',elem+'_dose']].notna().all(axis=1)
                paired.loc[~valid,[elem+'_in',elem+'_dose']]=np.nan
                plot(paired,{elem+'_in':'进水自带',elem+'_dose':'实际投加'},f'kg{elem}/d',f'{elem}总供应构成（完整日）',stacked=True)
                plot(now,{elem+'_dose':'实际投加'},f'kg{elem}/d',f'{elem}营养盐投加')
                st.caption(f'窗口加权比值有效 {count}/{window} 天。投加字段已是元素质量，不再换算药剂含量。')
        st.info('参考区间是历史中间50%的供应水平，不能证明生物最低需求。供应偏低时先核对进水自带N/P、投加记录和出水变化；供应偏高也不足以单独认定药剂浪费。')
        with st.expander('既有研究中的低供应证据'):
            st.dataframe(pd.DataFrame({'连续窗口':['7天','14天','30天'],'N平均供应':[16.34,20.77,23.16],'P平均供应':[1.484,2.445,2.924]}),hide_index=True)
            st.caption('单位g元素/kgCOD。以上为既有研究固定结果，非本次窗口重算；N和P的最低窗口不同。窗口内仍有高供应日，不能作为Minimum Need或每日投加设定值。机理参数尚未独立标定，不参与自动诊断。')
    with tabs[3]:
        chart_pair(now,[
            dict(mapping={'EOA_SVI':'数据库SVI'},unit='mL/g',title='污泥容积指数 SVI'),
            dict(mapping={'SV30_pct':'SV30'},unit='%',title='30分钟沉降体积'),
            dict(mapping={'SC_Sludge_Blanket':'泥层高度'},unit='m',title='二沉池泥层'),
            dict(mapping={'EO_SS':'出水SS'},unit='mg/L',title='固体随出水流失',line=10),
            dict(mapping={'RAS_pct':'回流污泥比','Recycle_pct':'内回流比'},unit='%',title='回流运行'),
            dict(mapping={'RS_SS':'回流污泥SS'},unit='mg/L',title='回流污泥浓度'),
            dict(mapping={'Cake_dry':'估算干固体'},unit='t/d',title='泥饼干固体量'),
            dict(mapping={'Aeration_Inlet_Temperature':'温度'},unit='°C',title='生化进水温度'),
        ])
        st.caption('SV30由mL/L换算为%；SVI受污泥浓度影响。泥饼可能含不同来源固体，不作为ES生物污泥排出的等价量。仅报告表中可见趋势，不推断菌相、泡沫或设备运行原因。')
    with tabs[4]:
        st.subheader('数据覆盖与可追溯性')
        coverage=pd.DataFrame({'指标':RAW,'有效天数':[int(now[c].count()) for c in RAW],
                               '窗口天数':window,'最新有效日期':[x.loc[:end,c].last_valid_index() for c in RAW]})
        st.dataframe(coverage,hide_index=True,width='stretch')
        recent_issues=issues.loc[issues['日期'].between(start,end)] if len(issues) else issues
        st.write(f'当前窗口数据核查项：{len(recent_issues)} 条')
        st.dataframe(recent_issues,hide_index=True,width='stretch')
        with st.expander(f'全部历史核查记录（{len(issues)}条）'):
            st.dataframe(issues,hide_index=True,width='stretch')
        st.caption(f'已忽略最新原始观测日期后的 {meta["trimmed"]} 行预留日期或公式残值。版本指纹 {meta["hash"]}。')
        if meta['missing']: st.warning('未找到这些字段，相应模块保留缺测：'+', '.join(meta['missing']))
        with st.expander('计算口径与更新方式',expanded=True):
            st.markdown(f'''- 每天上传完整新版Excel，或替换同目录工作簿后点击“重新读取本地文件”。分析日期默认跟随最新有效原始观测；历史数据不足时显示缺测。
- 历史筛选：COD≤30、TN≤10、NH₃-N≤5、TP≤0.5、SS≤10 mg/L，五项齐全；排除流量<5000 m³/d或表内低流量标记。筛选条件沿用研究定义，并非法律合规限值。
- 历史范围：2024年起、每个评价日前365天，不含当天。每个指标独立计算有效样本数，少于{MIN_HISTORY}天不生成区间。
- 生化COD负荷 = AEQ_COD × EO_Flow / 1000。EO_Flow作为共同过程流量代理。
- F/M = COD负荷 /（{volume:,.0f} m³ × MLVSS g/L）；EOA浓度作为总库存代表，属于诊断假设。
- MLSS下界/上界 = 当日COD负荷 /（池容 × 历史f中位数 × 历史F/M P75/P25）。f=MLVSS/MLSS。低流量日不生成MLSS范围。
- 原表Floc_Load_Excel与重算总生化F/M定义不同；原表目标未混用到F/M图。
- N/P总供应分别为AEQ TN/TP × EO_Flow / 1000 + 元素投加。窗口比值为完整配对日Σ总供应/ΣCOD×1000；历史P25/P75按每日等权，中心按COD负荷加权。
- 窗口与前一相邻等长日历窗口比较。指标覆盖至少70%才显示均值与变化；无填零、无前填、无插值。摘要中>10%的变化仅触发关注文字，不是工艺控制阈值。
- 数据保护：负数、非数字、无穷值、零MLSS/MLVSS/SRT、MLVSS>MLSS按缺测；泥层>20m、SS>100000mg/L、SV30>1000mL/L为疑似录入/单位错误，不自动修正。核查记录保留原值。
- 范围表示历史对应关系，非安全边界。SRT、F/M、SVI可能共享输入，不能按独立证据投票。结论不包含无法由日数据支持的微生物、设备故障或生产状态判断。''')
        st.subheader('当前窗口明细')
        st.dataframe(now.reset_index(),hide_index=True,width='stretch')
        st.download_button('下载本窗口计算数据',now.to_csv(index=True).encode('utf-8-sig'),f'WWTP_{end:%Y%m%d}_{window}d.csv','text/csv')
        report=f'WWTP运行回顾\n{start:%Y-%m-%d} 至 {end:%Y-%m-%d}\n\n{title}\n\n'+'\n\n'.join(k+'：'+v for k,v in findings)+'\n\n跟进重点：'+advice
        st.download_button('下载中文分析摘要',report.encode('utf-8-sig'),f'WWTP_summary_{end:%Y%m%d}.txt','text/plain')
    st.caption('WWTP ANALYTICS　/　预览迭代版　·　仅基于已记录数据解释运行状态')


if __name__=='__main__':
    main()
