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
import re

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
def read_shutdown(content):
    raw=pd.read_excel(BytesIO(content),sheet_name='Shutdown_History')
    records=[]; invalid=0
    for _,r in raw.iterrows():
        if pd.isna(r.get('machine')): continue
        machine=str(r['machine']).strip().upper()
        start=pd.to_datetime(str(r.get('stop_start','')).strip(),errors='coerce')
        end=pd.to_datetime(str(r.get('stop_end','')).strip(),errors='coerce')
        if machine not in ['PM1','PM2','PM3'] or pd.isna(start) or pd.isna(end) or end<=start:
            invalid+=1;continue
        records.append((machine,start,end))
    return records,invalid


def describe_shutdown(records,now,history):
    start=now.index.min();end=now.index.max()+pd.Timedelta(days=1)
    paragraphs=[]
    for machine in ['PM1','PM2','PM3']:
        intervals=sorted((max(a,start),min(b,end)) for m,a,b in records if m==machine and a<end and b>start)
        if not intervals: continue
        merged=[]
        for a,b in intervals:
            if merged and a<=merged[-1][1]: merged[-1]=(merged[-1][0],max(b,merged[-1][1]))
            else: merged.append((a,b))
        hours=sum((b-a).total_seconds()/3600 for a,b in merged)
        text=f'{machine}本期已记录停机{hours:.1f}小时，占本期时长{hours/(24*len(now)):.0%}。'
        text+=' '.join(f'{a:%m/%d %H:%M}—{b:%m/%d %H:%M}' for a,b in merged)+'。'
        onset=min(a for m,a,b in records if m==machine and a<end and b>start)
        # Only complete stopped calendar days; partial days do not represent pure shutdown conditions.
        days=[d for d in now.index if any(a<=d and b>=d+pd.Timedelta(days=1) for a,b in merged)]
        baseline=history.loc[(history.index>=onset.normalize()-pd.Timedelta(days=7))&(history.index<onset.normalize())]
        stopped=now.loc[days]
        changes={}
        for c in [machine+'_load','COD_load']:
            if len(stopped)>=2 and len(baseline)>=3 and stopped[c].notna().mean()>=.7 and baseline[c].notna().mean()>=.7 and baseline[c].mean()>0:
                changes[c]=(stopped[c].mean()/baseline[c].mean()-1)*100
        if changes:
            labels={machine+'_load':machine+'排水COD负荷','COD_load':'总进水COD负荷','MLVSS':'MLVSS','ES_solids':'排泥固体量'}
            dates='、'.join(f'{d:%m/%d}' for d in days)
            text+=f'完整停机日（{dates}）与停机前（{baseline.index.min():%m/%d}—{baseline.index.max():%m/%d}）的日均负荷比较：'+'、'.join(f'{labels[c]}{v:+.1f}%' for c,v in changes.items())+'。'
            if changes.get(machine+'_load',0)<0 and changes.get('COD_load',0)<0:
                text+='该纸机排水负荷与总负荷同步下降，支持停机与本期低负荷变化有关。'
            elif changes.get('COD_load',0)>0:
                text+='停机期间总负荷未同步降低，应结合其他生产线及进水变化分析。'
        paragraphs.append(text)
    return ' '.join(paragraphs)


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
FM_CONTROL = (0.10, 0.25)  # Plant-confirmed experience range; independent of Floc_Load_Excel.


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
    for field in ['FM','fraction','N_ratio','P_ratio','N_dose','P_dose','SRT','EOA_SVI','SC_Sludge_Blanket']:
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


def directional_advice(now, prev):
    """Preview decision rules: latest three consecutive days; no dose setpoints."""
    healthy=bool(len(now)>=3 and now.tail(3).complete_eo.all() and now.tail(3).pass_eo.all())
    recent=now.tail(3)
    continuous=len(recent)==3 and (recent.index[-1]-recent.index[0]).days==2
    low=continuous and recent.FM.notna().all() and recent.FM.lt(.10).all()
    high=continuous and recent.FM.notna().all() and recent.FM.gt(.25).all()
    changes={c:change(now,prev,c) for c in ['MLSS','MLVSS','ES_solids','COD_load','SV30_pct','SC_Sludge_Blanket','EO_SS']}
    up=lambda c: pd.notna(changes[c]) and changes[c]>0
    down=lambda c: pd.notna(changes[c]) and changes[c]<0
    settle_complete=all(pd.notna(changes[c]) for c in ['SV30_pct','SC_Sludge_Blanket','EO_SS'])
    # An increase in any of these indicators blocks a positive "increase wasting" trial.
    no_settle_counter=settle_complete and not any(up(c) for c in ['SV30_pct','SC_Sludge_Blanket','EO_SS'])
    all_in=now.FM.notna().all() and now.FM.between(.10,.25).all()
    inventory_adjustment=(all_in and now.complete_eo.all() and now.pass_eo.all()
        and now.MLSS.notna().all() and now.MLVSS.notna().all()
        and now.iloc[-1].MLSS>now.iloc[0].MLSS
        and 3<now.iloc[-1].MLSS<=4 and 3<=average(now.MLSS)<=4
        and down('COD_load') and down('ES_solids') and change(now,prev,'FM')<0
        and change(now,prev,'SRT')>0
        and pd.notna(changes['SC_Sludge_Blanket']) and changes['SC_Sludge_Blanket']<=0)
    split=len(now)//2
    early,late=now.iloc[:split],now.iloc[split:]
    trends={c:change(late,early,c) for c in ['MLSS','MLVSS','ES_solids','COD_load','FM']}
    rising=lambda c: pd.notna(trends[c]) and trends[c]>0
    falling=lambda c: pd.notna(trends[c]) and trends[c]<0
    latest=now.iloc[-1]
    separation_ok=(now.complete_eo.all() and now.pass_eo.all()
        and pd.notna(changes['SC_Sludge_Blanket']) and changes['SC_Sludge_Blanket']<=0)
    if now.low_flow.any():
        heading='低流量工况：暂不套用常规排泥方向'
        sludge='窗口含低流量日，先按当前实际运行构型确认负荷和有效池容，再制定排泥方向。'
    elif not recent.complete_eo.all():
        heading='出水数据未齐，暂缓主动改变排泥强度'
        sludge='补齐出水检测前，暂不启动基于本看板的增减排泥试调；当前结果不足以确认调整后的处理余量。'
    elif not healthy:
        if bool(now.iloc[-1].pass_eo):
            heading='出水已回到参考线内，先观察恢复与排泥响应'
            sludge='截止日出水已回到参考线内，但最近3天出现过波动，恢复观察尚短。暂维持现有排泥安排，跟踪后续出水与污泥逐日变化，不因期间污泥均值偏高追加排泥。'
        else:
            heading='截止日出水仍高于参考线，优先核查异常'
            sludge='先核对异常日负荷、检测与固体流失情况，暂不为追求历史MLSS范围主动增减排泥。'
    elif separation_ok and falling('MLSS') and falling('MLVSS') and rising('FM') and (rising('ES_solids') or up('ES_solids')) and pd.notna(latest.MLSS) and latest.FM>=.10:
        heading='可考虑减少排泥，保留污泥量'
        sludge='本期后段MLSS与MLVSS较前段下降、F/M升高，同时排泥固体量增加，支持评估适度减少排泥。结合负荷是否继续增加及出水响应确定调整；MLSS低于3000 mg/L时应优先关注生物量保留，回到控制范围且趋势稳定后重新评估。'
    elif separation_ok and rising('COD_load') and rising('FM') and not falling('MLVSS') and latest.FM>.25:
        heading='优先应对负荷增加，暂不追加排泥'
        sludge='本期后段负荷与F/M升高，MLVSS未同步下降，截止日F/M高于经验上限。优先核对负荷来源并评估均衡进水；避免为短时负荷峰值追调污泥库存，继续观察负荷与出水是否恢复。'
    elif separation_ok and rising('MLSS') and rising('MLVSS') and falling('FM') and not rising('COD_load') and latest.MLSS>3 and latest.FM<=.25 and not rising('ES_solids'):
        heading='可考虑适度增加排泥，缓解污泥积累'
        sludge='本期后段MLSS与MLVSS增加、F/M下降，负荷未增加，排泥也未增强，变化支持评估适度增加排泥。控制MLSS向3000–4000 mg/L范围内稳定运行；若接近下限、F/M升高或出水变差，应重新评估。'
    elif separation_ok and falling('MLSS') and falling('MLVSS') and falling('ES_solids') and 3<=latest.MLSS<=4 and .1<=latest.FM<=.25:
        heading='排泥已减少，先观察污泥量是否趋稳'
        sludge='本期后段排泥固体量已经减少，但MLSS与MLVSS仍下降；截止日MLSS和F/M均在控制范围内。先观察现有调整后的变化，避免连续追加调整；若污泥量继续降低并接近3000 mg/L，或F/M持续上升，再评估进一步减少排泥。'
    elif separation_ok and falling('MLSS') and rising('ES_solids') and 3<=latest.MLSS<=4 and .1<=latest.FM<=.25:
        heading='排泥增强后污泥量回落，先观察现有调整'
        sludge='本期后段排泥固体量增加、MLSS下降，截止日MLSS与F/M仍在控制范围内。当前先观察现有排泥安排的效果；MLSS继续接近3000 mg/L或F/M升高时评估回调，避免重复加大排泥。'
    elif low and up('MLSS') and up('MLVSS') and down('ES_solids') and no_settle_counter:
        heading='倾向逐步增加排泥，降低污泥库存'
        sludge='最近连续3天F/M低于0.10，MLSS与MLVSS较前窗增加、ES固体排出减少；出水满足筛选条件，沉降与出水SS没有同步上升信号。建议在现场确认后逐步增加排泥，观察F/M、库存和出水响应；不直接给出未经验证的调整幅度。'
    elif high and down('MLSS') and down('MLVSS') and up('ES_solids') and no_settle_counter:
        heading='倾向减少排泥，保留生物量'
        sludge='最近连续3天F/M高于0.25，MLSS与MLVSS较前窗下降、ES固体排出增加，支持先减少排泥以保留生物量；同步观察沉降和出水，避免仅追求F/M回落。'
    elif high and up('COD_load') and pd.notna(changes['MLVSS']) and changes['MLVSS']>=0:
        heading='优先缓冲进水负荷，暂不增加排泥'
        sludge='最近连续3天F/M高于0.25，COD负荷增加而MLVSS未下降，变化更支持负荷端压力。优先评估均衡进水与利用现有调节能力，暂不增加排泥；不将高F/M直接归为生物量流失。'
    elif inventory_adjustment:
        heading='可考虑适度增加排泥，匹配当前负荷'
        sludge='本期负荷与排泥固体量下降、SRT升高，期末MLSS高于期初，F/M较前期降低。MLSS总体仍在3000–4000 mg/L控制范围内，可结合后续MLSS、MLVSS及泥层变化评估小幅增加排泥。若MLSS接近3000 mg/L、F/M升高或出水出现不利变化，应重新评估排泥强度；不直接给出未经验证的调整幅度。'
    elif all_in and settle_complete and not (up('SC_Sludge_Blanket') or up('EO_SS')):
        heading='建议维持当前排泥强度'
        sludge='F/M各日均在0.10–0.25内、最近3天出水满足筛选条件，泥层与出水SS窗口均值未上升。建议维持当前排泥强度，不为贴合历史MLSS范围追加排泥；继续跟踪库存、SRT和固体排出。'
    elif up('SC_Sludge_Blanket') or up('EO_SS'):
        heading='先跟踪泥层与出水SS，暂不追加排泥'
        sludge='泥层或出水SS较前窗上升，应单独关注，不能被SV30下降抵消。先核对逐日变化及二沉池回流输送情况；暂不为贴合历史MLSS范围追加排泥。若上升持续或出水SS越过参考线，优先现场检查固液分离。'
    elif low or high:
        heading='F/M持续偏离，但增减排泥证据尚不一致'
        reasons=[]
        if not settle_complete: reasons.append('沉降或出水SS比较数据不足')
        elif not no_settle_counter: reasons.append('沉降或出水SS存在上升信号')
        if not all(pd.notna(changes[c]) for c in ['MLSS','MLVSS','ES_solids']): reasons.append('污泥库存或ES比较数据不足')
        if not reasons: reasons.append('库存与ES变化未形成支持同一排泥方向的组合')
        sludge='暂不启动增减排泥试调：'+'；'.join(reasons)+'。先解决这些具体矛盾，再选定调整方向。'
    else:
        heading='暂维持排泥强度，跟踪F/M偏离是否持续'
        sludge='最新3天未形成同方向的连续F/M超界，或关键观测不完整。暂维持排泥强度，逐日跟踪负荷与MLVSS；窗口内较早或孤立的超界不自动触发当前增减排泥。'
    nutrients=[]
    for elem in ['N','P']:
        value,count=supply_mean(now,elem);last=now.iloc[-1]
        lo,hi=last[elem+'_ratio_lo'],last[elem+'_ratio_hi']
        needed=[elem+'_ratio',elem+'_ratio_hi',elem+'_dose',elem+'_dose_hi',elem+'_in']
        complete=continuous and recent[needed].notna().all().all()
        high_supply=complete and recent[elem+'_ratio'].gt(recent[elem+'_ratio_hi']).all()
        high_dose=complete and recent[elem+'_dose'].gt(recent[elem+'_dose_hi']).all()
        inlet_change=change(now,prev,elem+'_in')
        if pd.isna(value) or pd.isna(lo) or pd.isna(hi):
            advice='暂不启动加减药试验，先补齐供应数据或积累足够参考样本。'
        elif healthy and high_supply and high_dose and pd.notna(inlet_change) and inlet_change>0:
            advice='优先评估下调外加营养盐的小步试验：最近3天总供应和实际投加均高于各自历史P75，进水自带量较前窗增加且最近3天出水满足筛选条件。先制定现场试验、监测和停止条件，再确定减量幅度；不能直接认定药剂过量。'
        elif healthy and value<lo:
            advice='不因低于历史区间而自动追加，暂沿用已记录期间的投加方式并观察出水响应。'
        elif not healthy and value<lo:
            advice='优先验证营养供应是否为限制因素，同时检查负荷和固体流失；暂不直接追加，避免把出水变化单独归因为营养不足。'
        elif healthy:
            advice='暂维持已记录期间的投加方式，不为贴合历史中心主动加减药。'
        else:
            advice='优先查明出水变化，暂不启动以历史供应位置为依据的减药试验。'
        if pd.isna(last[elem+'_dose']): advice+='截止日投加缺测，此方向仅适用于已有记录的期间，补录后再评价当天。'
        nutrients.append(f'{elem}：'+advice)
    return heading,sludge,nutrients


def effluent_review(now):
    """Separate dated events from the condition at the selected cutoff."""
    details=[]
    for field,limit in LIMITS.items():
        events=now.loc[now[field].gt(limit),field]
        if events.empty: continue
        shown=events if len(events)<=3 else events.loc[sorted(set([events.index[-1],events.idxmax()]))]
        dates='、'.join(f'{d:%Y-%m-%d}：{fmt(v,2)}' for d,v in shown.items())
        after=now.loc[now.index>events.index[-1],field]
        if after.empty:
            recovery='截止日仍高于参考线'
        elif after.isna().any():
            recovery='之后存在缺测，不能确认连续恢复'
        else:
            recovery=f'之后连续{len(after)}天回到参考线内'
        details.append(f'{SHORT[field]}共{len(events)}天高于参考线（{fmt(limit,2)} mg/L）：{dates} mg/L；{recovery}。')
    latest=now.iloc[-1]
    current=[SHORT[c] for c,v in LIMITS.items() if pd.notna(latest[c]) and latest[c]>v]
    if current: state='截止日仍有'+'、'.join(current)+'高于参考线'
    elif not latest.complete_eo: state='截止日出水检测未齐，当前状态待确认'
    elif details: state='本期出现过出水波动，截止日五项均回到参考线内'
    else: state='截止日五项出水均在参考线内'
    if not details:
        details=[f'本期五项出水齐全{int(now.complete_eo.sum())}/{len(now)}天，满足筛选条件{int(now.pass_eo.sum())}/{len(now)}天；缺测日不作正常判断。']
    return state,' '.join(details)


def fm_review(now, prev):
    """Date-specific alerts are independent of persistence-based adjustment rules."""
    abnormal=now.loc[now.FM.lt(.10)|now.FM.gt(.25),'FM']
    if abnormal.empty:
        return '本期未记录到F/M超出经验范围；缺测日不作正常判断。',None,None
    peak=abnormal.index[((abnormal-.175).abs()).argmax()]
    latest=abnormal.index[-1]
    dates=sorted(set([peak,latest]))
    parts=[]
    for date in dates:
        row=now.loc[date];v=row.FM;bound=.25 if v>.25 else .10
        text=f'{date:%Y-%m-%d} F/M {v:.3f}，'+('高于上限' if v>.25 else '低于下限')+f'{abs(v/bound-1)*100:.1f}%；MLSS {fmt(row.MLSS,2)} g/L。'
        prior_date=date-pd.Timedelta(days=1)
        history=pd.concat([prev,now])
        if prior_date in history.index:
            prior=history.loc[prior_date]
            for col,label in [('COD_load','COD负荷'),('MLVSS','MLVSS')]:
                if pd.notna(row[col]) and pd.notna(prior[col]) and prior[col]>0:
                    text+=f'{label}较前日{(row[col]/prior[col]-1)*100:+.1f}%。'
        parts.append(text)
    after=now.loc[now.index>latest,'FM']
    if after.empty: recovery='截止日仍越界，需当日关注。'
    elif after.isna().any(): recovery='后续存在缺测，无法确认连续恢复。'
    else: recovery=f'但截至{now.index[-1].month}月{now.index[-1].day}日，F/M已回到经验范围内。'
    current=now.iloc[-1].FM
    heading=advice=None
    if pd.notna(current) and (current>.25 or current<.1):
        heading='截止日F/M高于经验上限，优先核查负荷与生物量' if current>.25 else '截止日F/M低于经验下限，关注负荷与污泥积累'
        advice=('先核对当日COD负荷、流量和MLVSS。确认负荷增加后，优先评估均衡进水，暂不追加排泥；若高负荷持续且MLVSS继续下降，再结合沉降与出水评估减少排泥。' if current>.25 else '先区分负荷下降与污泥增加；不因单日偏低立即加大排泥。若低F/M持续且污泥继续积累、出水及沉降允许，再评估逐步增加排泥。')
    return ' '.join(parts)+recovery,heading,advice


def summarize(now, prev):
    n=len(now);last=now.iloc[-1]
    state,events=effluent_review(now)
    # Current decisions use the latest three calendar days; past events remain visible.
    recent=now.tail(3)
    title,sludge_direction,nutrients=directional_advice(now,prev)
    fm=now.FM.dropna()
    inside=int(fm.between(.10,.25).sum())
    changes={c:change(now,prev,c) for c in ['COD_load','MLSS','MLVSS','ES_solids','SRT','SV30_pct','SC_Sludge_Blanket','EO_SS']}
    def changed(c,label):
        v=changes[c]
        return f'{label}较前期{("增加" if v>0 else "减少" if v<0 else "持平")}'+(f'{abs(v):.1f}%' if v else '') if pd.notna(v) else f'{label}比较数据不足'
    def positive(c): return pd.notna(changes[c]) and changes[c]>0
    stock='；'.join(changed(c,label) for c,label in [('COD_load','进入系统的有机负荷'),('MLSS','池内污泥浓度'),('ES_solids','排出的污泥固体量')])+'。'
    if positive('MLSS') and positive('ES_solids'):
        stock+='排泥增强与本期污泥浓度上升同时存在，不能仅凭期间均值认定排泥无效或库存已开始下降。'
    elif positive('MLSS') and pd.notna(changes['COD_load']) and changes['COD_load']<0:
        stock+='负荷减少而污泥浓度增加，应关注两者的匹配，但这不等于生物活性下降。'
    stock+=f' 污泥浓度均值{fmt(average(now.MLSS),2)} g/L；'+changed('SRT','泥龄')+'（同期变化不单独证明排泥效果）。'
    # Compare complete adjacent three-day blocks, never silently bridge missing days.
    six=now.tail(6)
    trajectory=[]
    for c,label in [('MLSS','污泥浓度'),('MLVSS','挥发性污泥浓度'),('ES_solids','排泥固体量')]:
        if len(six)==6 and (six.index[-1]-six.index[0]).days==5 and six[c].notna().all():
            a,b=six[c].iloc[:3].mean(),six[c].iloc[3:].mean()
            if a>0: trajectory.append(f'{label}{("上升" if b>a else "下降" if b<a else "持平")}'+(f'{abs(b/a-1)*100:.1f}%' if b!=a else ''))
    trend=('最近3天均值与此前3天比较：'+'、'.join(trajectory)+'。这描述近期方向，不能确认是哪天调整排泥后的响应。') if trajectory else '最近两个连续3天的记录不足，暂不能确认污泥是否开始回落。'
    control=f'单位污泥承担的负荷（F/M）均值{fmt(average(now.FM),3)}；{inside}/{n}天在经验控制范围0.10–0.25内'
    control+=f'，低于下限{int(fm.lt(.1).sum())}天、高于上限{int(fm.gt(.25).sum())}天，缺测{n-len(fm)}天。'
    if inside==n: control+='目前没有F/M越界依据，不把污泥浓度偏离历史范围直接解释为污泥过多或偏老。'
    settle='；'.join(changed(c,label) for c,label in [('SC_Sludge_Blanket','泥层高度'),('EO_SS','出水SS'),('SV30_pct','SV30')])+'。'
    settle+=f'截止日：泥层{fmt(last.SC_Sludge_Blanket,2)} m，出水SS {fmt(last.EO_SS,2)} mg/L，SV30 {fmt(last.SV30_pct,1)}%。'
    if positive('SC_Sludge_Blanket') or positive('EO_SS'):
        settle+='泥层或出水SS的上升值得单独跟踪，SV30下降不能抵消这些信号；泥层尚无已确认控制线，不据百分比认定严重程度。'
    else: settle+='目前未见泥层或出水SS的窗口均值上升证据；缺失项目仍不能判断。'
    nutrient=[]
    for elem in ['N','P']:
        value,count=supply_mean(now,elem)
        text=f'{elem}总供应{fmt(value,2)} g/kgCOD，{position(value,last[elem+"_ratio_lo"],last[elem+"_ratio_hi"])}（{count}/{n}天）'
        for suffix,label in [('_in','进水自带量'),('_dose','外加量')]:
            v=change(now,prev,elem+suffix)
            text+=f'；{label}较前期'+(f'{v:+.1f}%' if pd.notna(v) else '数据不足')
        nutrient.append(text+'。')
    fm_events,alert,alert_advice=fm_review(now,prev)
    if alert and bool(last.complete_eo) and bool(last.pass_eo) and not bool(last.low_flow):
        title,sludge_direction=alert,alert_advice
    period=f'{now.index[0]:%Y-%m-%d}至{now.index[-1]:%Y-%m-%d}，共{n}天。五项出水齐全{int(now.complete_eo.sum())}天，其中{int(now.pass_eo.sum())}天满足参考筛选条件；F/M {inside}天在经验范围，{len(fm)-inside}天越界，{n-len(fm)}天缺测。'
    if len(fm)>inside:
        period+='本期存在负荷匹配偏离，不能以窗口平均值处于范围内概括为全期平稳。'
    elif inside==n and bool(now.pass_eo.all()):
        period+='本期出水与F/M均满足上述评价条件，仍需结合污泥及沉降变化判断调整方向。'
    split=n//2
    early,late=now.iloc[:split],now.iloc[split:]
    intra=[]
    for col,label in [('COD_load','COD负荷'),('MLSS','MLSS'),('ES_solids','排泥固体量')]:
        delta=change(late,early,col)
        if pd.notna(delta): intra.append(f'{label}{delta:+.1f}%')
    within=f'本期后{len(late)}天与前{len(early)}天的日均值比较：'+('、'.join(intra)+'。' if intra else '有效记录不足。')
    period=f'{now.index[0]:%Y-%m-%d}至{now.index[-1]:%Y-%m-%d}，出水五项满足参考条件{int(now.pass_eo.sum())}/{n}天，检测不完整{n-int(now.complete_eo.sum())}天；F/M有{len(fm)-inside}天超出经验范围。'+state+'。'
    within=f'{late.index[0]:%m/%d}—{late.index[-1]:%m/%d}与{early.index[0]:%m/%d}—{early.index[-1]:%m/%d}日均值比较：'+('、'.join(intra)+'。' if intra else '有效记录不足。')
    if len(late.COD_load.dropna())>=3 and late.COD_load.notna().all() and early.COD_load.notna().all():
        baseline=early.COD_load.mean()
        if baseline>0:
            typical=late.COD_load.drop(late.COD_load.idxmax()).mean()
            if late.COD_load.mean()>baseline and typical<=baseline:
                within+='后段负荷均值上升主要受最高负荷日拉动，其余日期平均负荷未高于前段。'
    stock=stock.replace('（同期变化不单独证明排泥效果）','')
    if changes['COD_load']<0 and changes['ES_solids']>0:
        stock+='负荷减少而排泥增强，需关注当前排泥强度与负荷的匹配。'
    elif changes['COD_load']<0 and changes['ES_solids']<0 and changes['MLVSS']<0:
        stock+='负荷与排泥量同步下降，但MLVSS仍降低，污泥量尚未表现为稳定。'
    event_text=[]
    if len(fm)>inside: event_text.append(fm_events)
    if any(now[c].gt(v).any() for c,v in LIMITS.items()): event_text.append(events)
    findings=[('综合结论',period)]
    if event_text: findings.append(('重点事件',' '.join(event_text)))
    findings.extend([('负荷与污泥',stock),('区间内变化',within),('负荷匹配',control),('沉降与出水',settle),('营养盐供应',' '.join(nutrient))])
    mlss_valid=now.MLSS.dropna()
    findings.insert(next(i for i,(k,v) in enumerate(findings) if k=='负荷匹配'),
        ('污泥控制',f'MLSS控制范围3000–4000 mg/L。本期{int(mlss_valid.between(3,4).sum())}/{n}天在范围内，低于下限{int(mlss_valid.lt(3).sum())}天，高于上限{int(mlss_valid.gt(4).sum())}天；截止日{fmt(last.MLSS*1000,0)} mg/L。'))
    vss_delta=change(now,prev,'MLVSS')
    fm_delta=change(now,prev,'FM')
    fraction_now=average(now.fraction)*100
    fraction_prev=average(prev.fraction)*100
    composition=f'MLVSS均值{fmt(average(now.MLVSS)*1000,0)} mg/L，截止日{fmt(last.MLVSS*1000,0)} mg/L'
    if pd.notna(vss_delta): composition+=f'，均值较前期{vss_delta:+.1f}%'
    composition+=f'；MLVSS/MLSS平均占比{fmt(fraction_now,1)}%'
    if pd.notna(fraction_now) and pd.notna(fraction_prev): composition+=f'，较前期{fraction_now-fraction_prev:+.1f}个百分点'
    composition+='。该占比反映挥发性固体组成，不等同于活性微生物比例。'
    explanation=f'F/M截止日{fmt(last.FM,3)} kgCOD/(kgMLVSS·d)'
    if pd.notna(fm_delta): explanation+=f'，本期均值较前期{fm_delta:+.1f}%'
    explanation+='。'
    load_delta=changes['COD_load']
    if pd.notna(load_delta) and pd.notna(vss_delta):
        if load_delta<0 and vss_delta>0:
            explanation+='负荷减少、MLVSS增加，两者均推动单位污泥负荷降低。'
        elif load_delta>0 and vss_delta<0:
            explanation+='负荷增加、MLVSS减少，两者均推动单位污泥负荷升高。'
        elif load_delta<0 and vss_delta<=0:
            explanation+='负荷与MLVSS均减少；负荷下降使F/M降低，MLVSS减少则起相反作用，应结合每日配对变化判断。'
        elif load_delta>0 and vss_delta>=0:
            explanation+='负荷与MLVSS均增加；负荷增加使F/M升高，MLVSS增加则起相反作用，应结合每日配对变化判断。'
        else:
            explanation+='F/M由同日COD负荷和MLVSS共同决定，不能仅根据MLSS高低判断。'
    findings=[('MLSS控制' if k=='污泥控制' else k,v+(' '+explanation if k=='负荷匹配' else '')) for k,v in findings]
    findings.insert(next(i for i,(k,v) in enumerate(findings) if k=='负荷匹配'),('MLVSS与组成',composition))
    sludge_summary=f'MLSS均值{fmt(average(now.MLSS)*1000,0)} mg/L，截止日{fmt(last.MLSS*1000,0)} mg/L；本期{int(mlss_valid.between(3,4).sum())}/{n}天处于3000–4000 mg/L控制范围。'
    sludge_summary+=f'MLVSS均值{fmt(average(now.MLVSS)*1000,0)} mg/L（较前期{fmt(vss_delta,1)}%），MLVSS/MLSS占比{fmt(fraction_now,1)}%。'
    sludge_summary+=f'F/M均值{fmt(average(now.FM),3)}，{inside}/{n}天在0.10–0.25范围内。'+explanation
    sludge_summary+=' '+changed('ES_solids','排泥固体量')+'，'+changed('SRT','泥龄')+'。'
    period_parts=[]
    for c,label in [('MLSS','MLSS'),('MLVSS','MLVSS'),('ES_solids','排泥固体量')]:
        delta=change(late,early,c)
        if pd.notna(delta): period_parts.append(f'{label}{delta:+.1f}%')
    if period_parts: sludge_summary+=f'本期后{len(late)}天较前{len(early)}天：'+ '、'.join(period_parts)+'。'
    load_summary=changed('COD_load','进水有机负荷')+f'，均值{fmt(average(now.COD_load),0)} kg COD/d。'
    load_intra=change(late,early,'COD_load')
    if pd.notna(load_intra): load_summary+=f'{late.index[0]:%m/%d}—{late.index[-1]:%m/%d}较{early.index[0]:%m/%d}—{early.index[-1]:%m/%d}日均负荷{load_intra:+.1f}%。'
    if '主要受最高负荷日拉动' in within: load_summary+='后段均值上升主要受最高负荷日拉动，其余日期平均负荷未高于前段。'
    findings=[(k,v) for k,v in findings if k in ['综合结论','重点事件']]+[
        ('生产与负荷',load_summary),('污泥分析',sludge_summary),('沉降与出水',settle),('营养盐供应',' '.join(nutrient))]
    actions=['负荷与排泥：'+sludge_direction,*nutrients]
    if len(nutrients)==2 and nutrients[0][2:]==nutrients[1][2:]:
        actions=['负荷与排泥：'+sludge_direction,'营养盐：'+nutrients[0][2:]]
    if now.low_flow.any(): findings.append(('适用限制','本期存在低流量日，常规排泥方向需先确认实际运行池容与构型。'))
    current_dates=f'{now.index[0]:%m/%d}—{now.index[-1]:%m/%d}'
    previous_dates=f'{prev.index[0]:%m/%d}—{prev.index[-1]:%m/%d}'
    first_dates=f'{early.index[0]:%m/%d}—{early.index[-1]:%m/%d}'
    second_dates=f'{late.index[0]:%m/%d}—{late.index[-1]:%m/%d}'
    def movement(value):
        return '比较数据不足' if pd.isna(value) else f'{"增加" if value>=0 else "减少"}{abs(value):.1f}%'
    es_valid=now.ES_solids.dropna()
    sludge_summary=('本期每天均有排泥记录，' if len(es_valid)==n and es_valid.gt(0).all() else f'本期有{int(es_valid.gt(0).sum())}天记录排泥，缺测{n-len(es_valid)}天，')
    sludge_summary+=f'日均排泥固体量（ES solids）{fmt(average(now.ES_solids),0)} kg TSS/d；{current_dates}较{previous_dates}{movement(changes["ES_solids"])}。'
    sludge_summary+=f'本期内，{second_dates}较{first_dates}日均排泥固体量{movement(change(late,early,"ES_solids"))}，同期MLSS{movement(change(late,early,"MLSS"))}、MLVSS{movement(change(late,early,"MLVSS"))}。'
    sludge_summary+=f'全期MLSS均值{fmt(average(now.MLSS)*1000,0)} mg/L，{int(mlss_valid.between(3,4).sum())}/{n}天在3000–4000 mg/L内；MLVSS均值{fmt(average(now.MLVSS)*1000,0)} mg/L，平均MLVSS/MLSS占比{fmt(fraction_now,1)}%。'
    sludge_summary+=f'F/M均值{fmt(average(now.FM),3)} kgCOD/(kgMLVSS·d)，{inside}/{n}天在0.10–0.25内。'
    sludge_summary+=f'平均泥龄（SRT）{fmt(average(now.SRT),1)} d，较{previous_dates}{movement(change(now,prev,"SRT"))}。'
    settle=f'{current_dates}与{previous_dates}比较：泥层平均高度{movement(changes["SC_Sludge_Blanket"])}，出水SS均值{movement(changes["EO_SS"])}，SV30均值{movement(changes["SV30_pct"])}。'
    settle+=f'本期均值分别为{fmt(average(now.SC_Sludge_Blanket),2)} m、{fmt(average(now.EO_SS),2)} mg/L和{fmt(average(now.SV30_pct),1)}%；出水SS有效范围{fmt(now.EO_SS.min(),1)}–{fmt(now.EO_SS.max(),1)} mg/L。'
    if now.EO_SS.gt(LIMITS['EO_SS']).any(): settle+='出水SS存在高于参考线的记录，见重点事件。'
    elif now.EO_SS.notna().all(): settle+='本期出水SS均在10 mg/L参考线内。'
    findings=[(k,sludge_summary if k=='污泥分析' else settle if k=='沉降与出水' else v.replace('较前期',f'较{previous_dates}')) for k,v in findings]
    actions=[v.replace('本期后段',f'{second_dates}较{first_dates}').replace('较前期',f'较{previous_dates}') for v in actions]
    if title=='排泥增强后污泥量回落，先观察现有调整':
        title='本期后段排泥增加、污泥浓度回落，暂不进一步加大排泥'
        actions[0]=f'负荷与排泥：{second_dates}较{first_dates}排泥增强，MLSS已有回落。建议先观察现有排泥安排，暂不进一步加大；若MLSS继续接近3000 mg/L，同时MLVSS下降、F/M升高，再评估适度减少排泥。'
    return title,findings,actions


def chart_formulas(frame, mapping, band=None, ratio=False):
    fields=set(mapping); notes=[]
    volume=frame.attrs.get('volume')
    v=f'{volume:,.0f} m³' if volume is not None else '生化有效池容（m³）'
    if 'COD_load' in fields:
        notes.append('COD负荷（kg COD/d）＝AEQ COD（mg/L）× EO出水流量（m³/d）÷1000。此处以EO流量作为过程流量代理。')
    if any(f'PM{i}_load' in fields for i in [1,2,3]):
        notes.append('各纸机COD负荷（kg COD/d）＝该纸机COD（mg/L）×该纸机排水流量（m³/d）÷1000。')
    if any(c.endswith('_removal') for c in fields):
        notes.append('表观去除率（%）＝（同日AEQ浓度－EO出水浓度）÷AEQ浓度×100。未校正停留时间，不代表质量去除率。')
    if 'FM' in fields:
        notes.append(f'F/M＝COD负荷（kg COD/d）÷［{v}×MLVSS（g/L）］，单位kgCOD/(kgMLVSS·d)；g/L与kg/m³数值相同。')
    if fields & {'MLSS','MLVSS'}:
        notes.append('MLSS、MLVSS（g/L）＝原表相应浓度（mg/L）÷1000。')
    if ratio:
        notes.append('右轴MLVSS/MLSS（%）＝同日MLVSS÷MLSS×100，两者使用相同浓度单位。')
    if fields & {'ES_solids','EO_solids'}:
        notes.append('ES排出（kg TSS/d）＝ES流量（m³/d）×ES实验室SS（mg/L）÷1000；出水携出＝EO流量×EO SS÷1000。堆叠总量为两者之和，任一缺测时不能视作完整总量。')
    if 'SRT_balance7' in fields:
        notes.append(f'7日TSS平衡诊断（d）＝近7日平均［MLSS（g/L）×{v}］÷近7日平均［ES排出＋出水携出（kg TSS/d）］。只取库存与排出均有效的配对日，至少4日；数据库SRT直接读取原表，未按此式重算。')
    if 'SV30_pct' in fields:
        notes.append('SV30（%）＝原表30分钟沉降体积（mL/L）÷10。')
    if fields & {'RAS_pct','Recycle_pct'}:
        notes.append('回流污泥比（%）＝RS回流污泥流量÷EO出水流量×100；内回流比（%）＝EOA回DN流量÷EO出水流量×100。所有流量单位均为m³/d；本图分母为出水流量，并非单独测得的进水流量。')
    if 'Cake_dry' in fields:
        notes.append('泥饼干固体量（t/d）＝湿泥饼量（t/d）×泥饼含固率（%）÷100。')
    for elem in ['N','P']:
        if fields & {elem+'_ratio',elem+'_in'}:
            notes.append(f'{elem}进水自带量（kg/d）＝AEQ T{elem}（mg/L）×EO流量（m³/d）÷1000；总供应＝进水自带量＋实际元素投加量（kg/d）。')
        if elem+'_ratio' in fields:
            notes.append(f'每日供应比（g{elem}/kgCOD）＝当日{elem}总供应÷当日COD负荷×1000。上方窗口指标＝完整配对日总供应之和÷同批日期COD负荷之和×1000。')
        if elem+'_dose' in fields:
            notes.append(f'{elem}投加直接读取原表元素质量（kg{elem}/d），不再换算药剂含量。')
    if band:
        if band[0]=='MLSS_lo':
            notes.append(f'MLSS参考下限/上限＝当日COD负荷÷［{v}×历史MLVSS/MLSS中位数×历史F/M的P75/P25］，结果为g/L。')
        notes.append('历史P25–P75：每个日期取此前365天、2024年起的有效记录，筛选五项出水齐全且COD≤30、TN≤10、氨氮≤5、TP≤0.5、SS≤10 mg/L，排除低流量日；每项至少60个有效样本，不含当天。')
    if notes: notes.append('计算所需数据缺失或分母为零时不生成该值，不按零填补。')
    return notes


def plot(frame, mapping, unit, title, mode='bar', band=None, line=None, stacked=False, control=None, ratio=False, caption=None, operating_band=None):
    st.markdown(f'<div class="chart-heading">{html.escape(title)}</div>',unsafe_allow_html=True)
    fig=go.Figure()
    if operating_band:
        fig.add_hrect(y0=operating_band[0],y1=operating_band[1],fillcolor='rgba(204,139,65,.12)',line_width=0,layer='below')
        for level in operating_band:
            fig.add_hline(y=level,line_dash='dash',line_color='#cc8b41')
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
    if control:
        for level,label in zip(control,['经验下限 0.10','经验上限 0.25']):
            fig.add_hline(y=level,line_dash='dash',line_color='#cc8b41',annotation_text=label,annotation_position='top right')
    legend_count=len(mapping)+int(bool(band))+int(ratio)
    legend_space=28*legend_count+45
    fig.update_layout(height=300+legend_space,
        paper_bgcolor='white',plot_bgcolor='white',font=dict(family='Segoe UI, Microsoft YaHei, sans-serif',size=12,color='#51655c'),
        margin=dict(l=15,r=20,t=20,b=legend_space),legend=dict(orientation='v',y=-.22,yanchor='top',x=0,font=dict(size=11)),
        hovermode='x unified',barmode='stack' if stacked else 'group',bargap=.3)
    fig.update_xaxes(tickformat='%m/%d',showgrid=False,zeroline=False)
    values=frame[list(mapping)].to_numpy(dtype=float).ravel().tolist()
    if stacked: values+=frame[list(mapping)].sum(axis=1,min_count=1).tolist()
    if band: values+=frame[list(band)].to_numpy(dtype=float).ravel().tolist()
    if line is not None: values.append(line)
    if control: values+=list(control)
    if operating_band: values+=list(operating_band)
    values=[v for v in values if pd.notna(v) and np.isfinite(v)]
    lower=(min(values) if mode!='bar' else min(0,min(values))) if values else 0
    upper=max(values) if values else 1
    span=max(upper-lower,abs(upper)*.1,1e-6)
    axis_min=(max(0,lower-.15*span) if lower>=0 else lower-.15*span) if mode!='bar' else (lower-.06*span if lower<0 else 0)
    axis_max=upper+.18*span
    axis_step=None
    if mode!='bar' and values:
        # Broad, rounded bounds: avoid both zero anchoring and tight magnification.
        broad_span=max(upper-lower,abs(upper)*.5,1e-6)
        raw_step=broad_span/3
        power=10**math.floor(math.log10(raw_step))
        axis_step=next(k*power for k in [1,2,2.5,5,10] if k*power>=raw_step)
        axis_min=math.floor((lower-.2*broad_span)/axis_step)*axis_step
        if lower>=0: axis_min=max(0,axis_min)
        axis_max=math.ceil((upper+.2*broad_span)/axis_step)*axis_step
    fig.update_yaxes(title=unit,gridcolor='#edf1ed',zeroline=False,
        range=[axis_min,axis_max],dtick=axis_step,automargin=True)
    fig.update_xaxes(automargin=True,nticks=5)
    fig.update_annotations(font=dict(size=10),bgcolor='rgba(255,255,255,.85)')
    if ratio:
        fig.add_trace(go.Scatter(x=frame.index,y=frame['fraction'],name='MLVSS/MLSS',
            mode='lines+markers',yaxis='y2',connectgaps=False,
            line=dict(color='#cc8b41',width=2.4),marker=dict(size=5),
            hovertemplate='%{x|%Y-%m-%d}<br>%{y:.1%}<extra>MLVSS/MLSS</extra>'))
        fig.update_layout(yaxis2=dict(title='MLVSS/MLSS',overlaying='y',side='right',
            range=[0,1],tickformat='.0%',showgrid=False,zeroline=False),
            legend=dict(font=dict(size=10)),margin=dict(r=48))
    st.plotly_chart(fig,width='stretch',config={'displaylogo':False,'modeBarButtonsToRemove':['select2d','lasso2d']})
    if caption:
        st.caption(caption)
    formulas=chart_formulas(frame,mapping,band,ratio)
    if formulas:
        st.markdown('<div class="chart-formula"><b>计算口径</b><br>'+ '<br>'.join(html.escape(s) for s in formulas)+'</div>',unsafe_allow_html=True)


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
    .chart-heading {font-size:16px;font-weight:650;color:#243c33;line-height:1.5;margin:2px 0 10px;overflow-wrap:anywhere}
    .chart-formula {font-size:11px!important;line-height:1.5;color:#78867e;background:transparent;padding:5px 2px;border-top:1px solid #e6ece7;margin:4px 0 8px;overflow-wrap:anywhere}
    .chart-formula b {font-weight:500;font-size:11px!important}
    .summary {background:#fff;border:1px solid #dde7de;border-left:5px solid #167455;border-radius:12px;padding:22px 26px;margin:5px 0 20px}
    .summary h3 {font-size:19px!important;margin:0 0 12px;padding:0;color:#214335}
    .finding {font-size:13px;line-height:1.9;margin:16px 0;color:#4d6055}
    .finding > b {display:block;color:#244d38;margin-bottom:5px;font-size:13px}
    .summary p {margin:0 0 8px;line-height:1.9;overflow-wrap:break-word}
    .summary .reading-token {white-space:nowrap}
    .summary-section {font-size:12px;font-weight:700;color:#167455;letter-spacing:.08em;border-top:1px solid #e3ebe5;padding-top:12px;margin-top:15px;margin-bottom:8px}
    .event {background:#fff4e3;border-left:4px solid #cc8b41;padding:12px 15px;margin:14px 0;border-radius:6px;font-size:13px;line-height:1.8}
    .event.resolved {background:#f0f4f2;border-left-color:#86a697}
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


def metric_labels(text):
    """Expand narrative terms without changing chart series or field identifiers."""
    replacements={
        '挥发性污泥浓度':'挥发性污泥浓度（MLVSS，g/L）',
        '池内污泥浓度':'池内污泥浓度（MLSS，g/L）',
        '污泥浓度均值':'污泥浓度（MLSS）均值',
        '排泥固体量':'排泥固体量（ES solids，kg TSS/d）',
        '排出的污泥固体量':'排出的污泥固体量（ES solids，kg TSS/d）',
        '泥层高度':'二沉池泥层高度（SC blanket，m）',
        '泥龄':'泥龄（SRT，d）',
        '进入系统的有机负荷':'进水有机负荷（COD load，kg COD/d）',
        '总进水COD负荷':'总进水COD负荷（kg COD/d）',
        '进水自带量':'进水自带量（kg元素/d）',
        '外加量':'外加量（kg元素/d）',
    }
    for old,new in replacements.items(): text=text.replace(old,new)
    return text


def report_html(text):
    """Readable prose only; keep the numeric and chart data untouched."""
    text=re.sub(r'(?<=\d),(?=\d{3}(?:\D|$))','',text)
    text=text.replace('（ES solids，kg TSS/d） (ES solids)','（ES solids，kg TSS/d）')
    text=text.replace('（ES solids，kg TSS/d）（ES solids）','（ES solids，kg TSS/d）')
    text=text.replace('（SRT，d）（SRT）','（SRT，d）')
    text=text.replace('（SRT，d） (SRT)','（SRT，d）')
    # Replace signed changes with words, so a date does not run into a minus sign.
    text=re.sub(r'([+-])(\d+(?:\.\d+)?)%',lambda m:('增加 ' if m[1]=='+' else '减少 ')+m[2]+'%',text)
    text=re.sub(r'([\u4e00-\u9fff])([A-Za-z0-9])',r'\1 \2',text)
    text=re.sub(r'([A-Za-z0-9%])([\u4e00-\u9fff])',r'\1 \2',text)
    sentences=[p.strip()+'。' for p in text.split('。') if p.strip()]
    paragraphs=[];current=''
    for sentence in sentences:
        if current and len(current)+len(sentence)>125:
            paragraphs.append(current);current=''
        current+=sentence
    if current: paragraphs.append(current)
    def escape_tokens(p):
        escaped=html.escape(p)
        return re.sub(r'\d{2}/\d{2}(?:—\d{2}/\d{2})?|\d+(?:\.\d+)?\s*(?:kg TSS/d|kg COD/d|mg/L|g/L|%)',
            lambda m:'<span class="reading-token">'+m[0]+'</span>',escaped)
    return ''.join('<p>'+escape_tokens(p)+'</p>' for p in paragraphs)


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
    shutdown_upload=st.sidebar.file_uploader('更新停机记录（可选）',type=['xlsx'],key='shutdown_upload')
    shutdown_records=[];shutdown_invalid=0
    shutdown_path=Path(__file__).parent/'data'/'PM_shutdown_history.xlsx'
    if shutdown_upload or shutdown_path.exists():
        try:
            shutdown_records,shutdown_invalid=read_shutdown(shutdown_upload.getvalue() if shutdown_upload else shutdown_path.read_bytes())
        except Exception as exc:
            st.sidebar.warning(f'停机记录暂不可用：{exc}')
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
    now.attrs['volume']=volume
    prev=x.reindex(pd.date_range(start-pd.Timedelta(days=window),start-pd.Timedelta(days=1),name='Date'))
    title,findings,actions=summarize(now,prev)
    production=describe_shutdown(shutdown_records,now,x)
    if production:
        findings=[(k,production+' '+v if k=='生产与负荷' else v) for k,v in findings]
    findings=[(k,metric_labels(v)) for k,v in findings]
    actions=[metric_labels(v) for v in actions]
    st.markdown('<div class="eyebrow">OPERATIONS INTELLIGENCE / WATER & ENVIRONMENT</div>',unsafe_allow_html=True)
    left,right=st.columns([3,1])
    with left: st.title('WWTP 运行健康看板')
    with right: st.caption(f'分析周期  {window} 天\n\n{start:%Y.%m.%d} — {end:%Y.%m.%d}')
    st.markdown('<p class="sub">污泥状态、处理表现与营养供应的综合回顾</p>',unsafe_allow_html=True)
    last=now.iloc[-1]
    n_value,n_count=supply_mean(now,'N');p_value,p_count=supply_mean(now,'P')
    lines=''
    for k,v in findings:
        if k=='重点事件':
            active=any(pd.notna(last[c]) and last[c]>v for c,v in LIMITS.items()) or (pd.notna(last.FM) and not .1<=last.FM<=.25)
            lines+=f'<div class="event{"" if active else " resolved"}"><b>{"重点关注" if active else "期间事件"}</b>{report_html(v)}</div>'
            continue
        if k=='生产与负荷':
            lines+='<div class="summary-section">区间运行分析</div>'
        lines+=f'<div class="finding"><b>{html.escape(k)}</b>{report_html(v)}</div>'
    advice=''.join(report_html(a) for a in actions)
    st.markdown(f'<div class="summary"><div class="eyebrow">运行判断与调整意见</div><h3>{html.escape(title)}</h3>{lines}<div class="actions"><b>调整与跟进　</b>{advice}</div></div>',unsafe_allow_html=True)
    st.caption('评价口径：出水使用既定研究筛选条件，非正式排放合规判定；F/M使用已确认的经验控制范围；历史P25–P75仅作运行参照。')
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
        plot(now,{'MLSS':'实测MLSS'},'g/L','MLSS与运行控制范围',mode='line',operating_band=(3,4),caption='浅橙色控制带：3000–4000 mg/L（3.0–4.0 g/L）。范围内仍可结合负荷、F/M、排泥及出水表现评估库存调整。')
        st.info('MLSS范围由当日COD负荷、此前365天F/M P25–P75和历史MLVSS/MLSS中位数换算。范围表示历史运行对应关系，偏离不直接触发排泥。')
        plot(now,{'MLSS':'实测MLSS'},'g/L','MLSS与负荷对应参考范围',mode='line',band=('MLSS_lo','MLSS_hi'))
        last=now.iloc[-1]
        st.caption(f'截止日参考：{fmt(last.MLSS_lo,2)}–{fmt(last.MLSS_hi,2)} g/L；F/M有效历史 {int(last.FM_n)} 天；f中位数 {fmt(last.fraction_mid,3)}。每一天均使用该日之前的历史，最少需{MIN_HISTORY}个有效历史样本。')
        chart_pair(now,[
            dict(mapping={'FM':'重算F/M'},unit='kgCOD/(kgMLVSS·d)',title='F/M：经验控制与历史参考',band=('FM_lo','FM_hi'),control=FM_CONTROL,
                 caption='经验控制范围：0.10–0.25 kgCOD/(kgMLVSS·d)。橙色虚线表示经验上下限；绿色阴影表示每日历史P25–P75，仅作运行参照。'),
            dict(mapping={'MLSS':'MLSS','MLVSS':'MLVSS'},unit='g/L',title='污泥浓度组成',ratio=True),
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
        st.caption('停机分析只使用有效起止时间，按实际重叠时长计算，重叠记录合并。无记录不视为正常开机。完整停机日与停机前7天比较，描述同期关联，不推断操作目的或未记录的加药。')
        if shutdown_invalid: st.caption(f'停机表有{shutdown_invalid}条记录起止时间或机台无效，未参与分析。')
        st.caption('排泥方向中的持续偏离暂采用最新连续3个日历日同向超界，缺测中断；这是待现场验证的预览规则，不影响单日异常提示。')
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
- F/M经验控制范围为0.10–0.25 kgCOD/(kgMLVSS·d)，由用户确认。按每个有效日检查是否超界，等于上下限视为范围内。原表Floc_Load_Excel的70–150范围不参与此评价。历史F/M P25–P75用于描述既往集中水平，不能替代经验控制范围；MLSS参考范围继续从历史F/M反算，不把经验上下限当成新的MLSS目标。
- N/P总供应分别为AEQ TN/TP × EO_Flow / 1000 + 元素投加。窗口比值为完整配对日Σ总供应/ΣCOD×1000；历史P25/P75按每日等权，中心按COD负荷加权。
- 窗口与前一相邻等长日历窗口比较。指标覆盖至少70%才显示均值与变化；无填零、无前填、无插值。总结先检查出水、再检查逐日F/M经验范围，沉降只描述变化方向；未设定未经确认的“明显恶化”百分比阈值，不由同期变化直接推断因果。
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
