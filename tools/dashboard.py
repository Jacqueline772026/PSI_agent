"""PSI 周度总览看板：从 PSI 导出生成一个本地 HTML 文件，用浏览器打开就能看。

数据只在本地：HTML 里内嵌了这一周的汇总数，不联网、不上传。发给别人走飞书或公司渠道。

看什么：
  - 本周 SO（全部 / 公开市场 / 运营商）：环比、对近 4 周平均、对 SO 预测
  - 渠道库存和可卖天数（DOS = 库存 ÷ 近 4 周平均 SO × 7）
  - DB 库存（AB / ALSO / IM）
  - 近 16 周 SO、渠道库存走势（公开市场 vs 运营商）
  - 分渠道、分系列明细；本周涨跌最多的机型；自动提示

用法：
  .venv/bin/python -m tools.dashboard --week 45
  .venv/bin/python -m tools.dashboard --week 45 --psi "数据/PSI历史版本/xxx.xlsx" --out "W45"
"""
import argparse
import datetime as dt
import html
import json
from pathlib import Path

import openpyxl
import pandas as pd

from checks.structure_check import COL_C, COL_E, COL_LABEL, OTHERS, is_number, load
from psi_tool.paths import OUTPUT, PSI_FILE, ROOT
from psi_tool.sources import week_range

OM = ["EURO", "MEX", "MSH", "X-KOM", "Michain", "OM Others"]
OP = ["Play", "TMO", "Plus", "Orange"]
LABELS = {"SO": "SO", "Channel Inventory Actual": "库存", "SO FCST": "预测"}
SERIES = {"RM Note": "Redmi Note", "Redmi": "Redmi", "xiaomi": "Xiaomi", "POCO": "POCO", "pad": "Pad"}
TREND_WEEKS = 16
DOS_HIGH, DOS_LOW = 120, 21  # 可卖天数超过 120 天提示压货，低于 21 天提示断货风险（阈值待和导师确认）
MOVE_PCT = 0.2  # 环比变化超过 20% 才提示


def extract(psi_path):
    """PSI 各 KA 页 → 长表：sheet、渠道、代码、产品、系列、指标、周、数值（只取手填的数值，公式格跳过）。"""
    psi = load(psi_path)
    rows = []
    for name in OM + OP:
        sh = psi[name]
        ch = "公开市场" if name in OM else "运营商"
        for r in range(1, sh.last_label_row + 1):
            lab = sh.get(r, COL_LABEL)
            if lab not in LABELS:
                continue
            others = sh.get(r, COL_C) in OTHERS
            if not others and r < sh.first_product:
                continue  # 小计行
            code = "Others" if others else sh.code(r)
            for w in range(1, 53):
                v = sh.get(r, sh.week_col(w))
                if is_number(v):
                    rows.append((name, ch, code, "Others" if others else sh.get(r, COL_E), sh.get(r, COL_C),
                                 LABELS[lab], w, float(v)))
    return pd.DataFrame(rows, columns=["sheet", "ch", "code", "name", "series", "lab", "w", "v"])


def db_stock(psi_path):
    wb = openpyxl.load_workbook(psi_path, read_only=True)
    ws = wb["2026 DB STOCK"]
    ws.reset_dimensions()
    rows = [(r[0], r[1], r[8] or 0) for r in ws.iter_rows(min_row=3, values_only=True)
            if r and isinstance(r[0], int) and r[1] in ("AB", "ALSO", "IM") and is_number(r[8] or 0)]
    return pd.DataFrame(rows, columns=["w", "db", "v"])


def filled(d, sheet, week):
    """这张 sheet 这一周 SO 有没有填（至少 5 格是数值）。"""
    return (d[(d.sheet == sheet) & (d.lab == "SO") & (d.w == week)].shape[0]) >= 5


def pct(a, b):
    return None if not b else a / b - 1


def build(d, db, week):
    so, inv, fc = (d[d.lab == x] for x in ("SO", "库存", "预测"))
    sheets_ok = {s: filled(d, s, week) for s in OM + OP}
    weeks = list(range(max(1, week - TREND_WEEKS + 1), week + 1))

    def total(frame, w, sheets):
        return frame[(frame.w == w) & frame.sheet.isin(sheets)].v.sum()

    def kpi(sheets, label):
        sheets = [s for s in sheets if sheets_ok[s]]
        cur, prev = total(so, week, sheets), total(so, week - 1, sheets)
        avg4 = sum(total(so, w, sheets) for w in range(week - 3, week + 1)) / 4
        fcst = total(fc, week, sheets)
        stock = total(inv, week, sheets)
        return dict(label=label, so=cur, wow=pct(cur, prev), vs4=pct(cur, avg4), fcst=pct(cur, fcst) if fcst else None,
                    stock=stock, dos=stock / avg4 * 7 if avg4 else None, n=len(sheets))

    kpis = [kpi(OM + OP, "全部"), kpi(OM, "公开市场"), kpi(OP, "运营商")]
    dbw = db[db.w == week].v.sum()
    dbp = db[db.w == week - 1].v.sum()
    db_kpi = dict(stock=dbw, wow=pct(dbw, dbp), by=db[db.w == week].groupby("db").v.sum().to_dict())

    trend = {}
    for ch, sheets in (("公开市场", OM), ("运营商", OP)):
        trend[ch] = dict(so=[total(so, w, sheets) for w in weeks], stock=[total(inv, w, sheets) for w in weeks])

    # 分渠道
    channels = []
    for s in OM + OP:
        cur, prev = total(so, week, [s]), total(so, week - 1, [s])
        avg4 = sum(total(so, w, [s]) for w in range(week - 3, week + 1)) / 4
        fcst, stock = total(fc, week, [s]), total(inv, week, [s])
        channels.append(dict(sheet=s, ch="公开市场" if s in OM else "运营商", filled=sheets_ok[s], so=cur,
                             wow=pct(cur, prev), avg4=avg4, fcst=pct(cur, fcst) if fcst else None,
                             stock=stock, dos=stock / avg4 * 7 if avg4 else None))

    # 分系列（只算本周已填的 sheet）
    ok = [s for s in OM + OP if sheets_ok[s]]
    ser = []
    cur_all = total(so, week, ok)
    for key, label in SERIES.items():
        f = so[so.sheet.isin(ok) & (so.series == key)]
        c, p = f[f.w == week].v.sum(), f[f.w == week - 1].v.sum()
        ser.append(dict(series=label, so=c, wow=pct(c, p), share=c / cur_all if cur_all else None))

    # 机型涨跌（全部已填 sheet 合计）
    f = so[so.sheet.isin(ok) & (so.code != "Others")]
    by = f[f.w.isin([week, week - 1])].pivot_table(index=["code", "name"], columns="w", values="v", aggfunc="sum").fillna(0)
    by = by.reindex(columns=[week - 1, week], fill_value=0)
    by["diff"] = by[week] - by[week - 1]
    movers = [dict(code=c, name=n, prev=r[week - 1], cur=r[week], diff=r["diff"]) for (c, n), r in by.iterrows()]
    ups = sorted([m for m in movers if m["diff"] > 0], key=lambda m: -m["diff"])[:5]
    downs = sorted([m for m in movers if m["diff"] < 0], key=lambda m: m["diff"])[:5]

    # 提示
    notes = []
    missing = [s for s in OM + OP if not sheets_ok[s]]
    if missing:
        notes.append(("warning", f"本周还没填 SO：{'、'.join(missing)}。这几家不计入本周合计和环比"))
    for c in channels:
        if not c["filled"]:
            continue
        if c["wow"] is not None and abs(c["wow"]) >= MOVE_PCT and abs(c["so"] - c["so"] / (1 + c["wow"])) >= 100:
            notes.append(("info", f"{c['sheet']} SO 环比 {c['wow']:+.0%}（{fmt(c['so'] / (1 + c['wow']))} → {fmt(c['so'])}）"))
        if c["dos"] and c["dos"] > DOS_HIGH:
            notes.append(("serious", f"{c['sheet']} 库存够卖 {c['dos']:.0f} 天（{fmt(c['stock'])} 台），超过 {DOS_HIGH} 天"))
        if c["dos"] is not None and c["dos"] < DOS_LOW and c["avg4"] >= 50:
            notes.append(("critical", f"{c['sheet']} 库存只够卖 {c['dos']:.0f} 天，有断货风险"))
        if c["fcst"] is not None and abs(c["fcst"]) >= 0.3:
            notes.append(("info", f"{c['sheet']} SO 比预测{'高' if c['fcst'] > 0 else '低'} {abs(c['fcst']):.0%}，预测可能没填全或偏差大"))
    return dict(week=week, weeks=weeks, kpis=kpis, db=db_kpi, trend=trend, channels=channels, series=ser,
                ups=ups, downs=downs, notes=notes)


def fmt(x):
    return "—" if x is None or pd.isna(x) else f"{x:,.0f}"


def fpct(x):
    return "—" if x is None or pd.isna(x) else f"{x:+.0%}"


STATUS = {"good": ("✔", "正常"), "warning": ("⚠", "注意"), "serious": ("▲", "偏高"), "critical": ("●", "风险"),
          "info": ("ℹ", "变化")}


def dos_status(dos):
    if dos is None:
        return None
    return "serious" if dos > DOS_HIGH else "critical" if dos < DOS_LOW else "good"


def render(m, psi_name):
    start, end = week_range(m["week"])
    esc = html.escape

    def tile(k):
        sub = [f"环比 {fpct(k['wow'])}", f"对 4 周均 {fpct(k['vs4'])}"]
        if k["fcst"] is not None:
            sub.append(f"对预测 {fpct(k['fcst'])}")
        return (f'<div class="tile"><div class="tile-label">{esc(k["label"])} SO</div>'
                f'<div class="hero">{fmt(k["so"])}</div><div class="tile-sub">{" · ".join(sub)}</div>'
                f'<div class="tile-sub">库存 {fmt(k["stock"])} · 够卖 {fmt(k["dos"])} 天</div></div>')

    tiles = "".join(tile(k) for k in m["kpis"])
    dbk = m["db"]
    tiles += (f'<div class="tile"><div class="tile-label">DB 库存</div><div class="hero">{fmt(dbk["stock"])}</div>'
              f'<div class="tile-sub">环比 {fpct(dbk["wow"])}</div>'
              f'<div class="tile-sub">{" · ".join(f"{k} {fmt(v)}" for k, v in dbk["by"].items())}</div></div>')

    def status_cell(st):
        if not st:
            return "<td></td>"
        icon, label = STATUS[st]
        return f'<td><span class="status {st}"><span aria-hidden="true">{icon}</span> {label}</span></td>'

    ch_rows = []
    for c in m["channels"]:
        if not c["filled"]:
            ch_rows.append(f'<tr class="muted"><td>{esc(c["sheet"])}</td><td>{esc(c["ch"])}</td>'
                           f'<td colspan="7">本周还没填</td></tr>')
            continue
        ch_rows.append(f'<tr><td>{esc(c["sheet"])}</td><td>{esc(c["ch"])}</td><td class="num">{fmt(c["so"])}</td>'
                       f'<td class="num">{fpct(c["wow"])}</td><td class="num">{fmt(c["avg4"])}</td>'
                       f'<td class="num">{fpct(c["fcst"])}</td><td class="num">{fmt(c["stock"])}</td>'
                       f'<td class="num">{fmt(c["dos"])}</td>{status_cell(dos_status(c["dos"]))}</tr>')
    ser_rows = "".join(f'<tr><td>{esc(s["series"])}</td><td class="num">{fmt(s["so"])}</td>'
                       f'<td class="num">{fpct(s["wow"])}</td><td class="num">'
                       f'{"—" if s["share"] is None else f"{s["share"]:.0%}"}</td></tr>' for s in m["series"])

    def mover_rows(lst):
        return "".join(f'<tr><td>{esc(str(x["name"]))}</td><td class="code">{esc(str(x["code"]))}</td>'
                       f'<td class="num">{fmt(x["prev"])}</td><td class="num">{fmt(x["cur"])}</td>'
                       f'<td class="num">{x["diff"]:+,.0f}</td></tr>' for x in lst) or '<tr><td colspan="5">—</td></tr>'

    notes = "".join(f'<li><span class="status {k}"><span aria-hidden="true">{STATUS[k][0]}</span> {STATUS[k][1]}</span> '
                    f'{esc(t)}</li>' for k, t in m["notes"]) or "<li>本周没有需要特别注意的变化。</li>"
    data = json.dumps(dict(weeks=m["weeks"], trend=m["trend"]), ensure_ascii=False)
    return TEMPLATE.replace("{{TITLE}}", f"PSI 周度总览 W{m['week']}").replace(
        "{{WEEK}}", f"2026 W{m['week']}（{start:%m-%d} ~ {end:%m-%d}）").replace(
        "{{SOURCE}}", esc(f"{psi_name}，生成于 {dt.datetime.now():%Y-%m-%d %H:%M}")).replace(
        "{{TILES}}", tiles).replace("{{NOTES}}", notes).replace("{{CHANNELS}}", "".join(ch_rows)).replace(
        "{{SERIES}}", ser_rows).replace("{{UPS}}", mover_rows(m["ups"])).replace(
        "{{DOWNS}}", mover_rows(m["downs"])).replace("{{DATA}}", data).replace(
        "{{DOS_HIGH}}", str(DOS_HIGH)).replace("{{DOS_LOW}}", str(DOS_LOW))


TEMPLATE = r"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{TITLE}}</title>
<style>
.viz-root{color-scheme:light;--surface-0:#f3f2ef;--surface-1:#fcfcfb;--text-primary:#0b0b0b;--text-secondary:#52514e;
--text-muted:#7a7974;--grid:#e4e3df;--border:#dcdbd6;--series-1:#2a78d6;--series-2:#eb6834;
--good:#0ca30c;--warning:#fab219;--serious:#ec835a;--critical:#d03b3b}
@media (prefers-color-scheme:dark){:root:where(:not([data-theme="light"])) .viz-root{color-scheme:dark;--surface-0:#111110;
--surface-1:#1a1a19;--text-primary:#f0efec;--text-secondary:#c3c2b7;--text-muted:#94938b;--grid:#2e2e2b;--border:#3a3a36;
--series-1:#3987e5;--series-2:#d95926}}
:root[data-theme="dark"] .viz-root{color-scheme:dark;--surface-0:#111110;--surface-1:#1a1a19;--text-primary:#f0efec;
--text-secondary:#c3c2b7;--text-muted:#94938b;--grid:#2e2e2b;--border:#3a3a36;--series-1:#3987e5;--series-2:#d95926}
*{box-sizing:border-box}body{margin:0}
.viz-root{background:var(--surface-0);color:var(--text-primary);font:14px/1.5 -apple-system,BlinkMacSystemFont,"PingFang SC",
"Microsoft YaHei",sans-serif;min-height:100vh;padding:24px 16px 48px}
.wrap{max-width:1180px;margin:0 auto}
header h1{font-size:22px;margin:0 0 4px}header p{margin:0;color:var(--text-secondary)}
.card{background:var(--surface-1);border:1px solid var(--border);border-radius:12px;padding:16px 18px;margin-top:16px}
.card h2{font-size:15px;margin:0 0 12px}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-top:16px}
.tile{background:var(--surface-1);border:1px solid var(--border);border-radius:12px;padding:14px 16px}
.tile-label{color:var(--text-secondary);font-size:13px}.hero{font-size:30px;font-weight:650;margin:2px 0 4px}
.tile-sub{color:var(--text-secondary);font-size:12.5px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}@media(max-width:820px){.grid2{grid-template-columns:1fr}}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:7px 8px;border-bottom:1px solid var(--grid);text-align:left}
th{color:var(--text-secondary);font-weight:500}td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}
td.code{color:var(--text-muted);font-size:12px}tr.muted td{color:var(--text-muted)}
.status{display:inline-flex;gap:4px;align-items:center;font-size:12px;padding:1px 7px;border-radius:99px;
border:1px solid currentColor;white-space:nowrap}
.status.good{color:var(--good)}.status.warning{color:#b37a00}.status.serious{color:#c5552b}.status.critical{color:var(--critical)}
.status.info{color:var(--text-secondary)}
ul.notes{margin:0;padding-left:0;list-style:none}ul.notes li{padding:5px 0;border-bottom:1px solid var(--grid)}
.legend{display:flex;gap:16px;color:var(--text-secondary);font-size:12.5px;margin-bottom:6px}
.legend i{display:inline-block;width:16px;height:2px;vertical-align:middle;margin-right:6px}
.chart{position:relative}.chart svg{width:100%;height:auto;display:block;overflow:visible}
.tip{position:absolute;pointer-events:none;background:var(--surface-1);border:1px solid var(--border);border-radius:8px;
padding:6px 10px;font-size:12.5px;box-shadow:0 4px 14px rgba(0,0,0,.12);display:none;white-space:nowrap}
.tip b{font-variant-numeric:tabular-nums}.tip i{display:inline-block;width:12px;height:2px;vertical-align:middle;margin-right:6px}
.foot{color:var(--text-muted);font-size:12px;margin-top:16px}
details summary{cursor:pointer;color:var(--text-secondary);font-size:12.5px;margin-top:8px}
</style></head>
<body><div class="viz-root"><div class="wrap">
<header><h1>PSI 周度总览</h1><p>{{WEEK}} · 数据来源：{{SOURCE}}</p></header>
<div class="tiles">{{TILES}}</div>
<div class="card"><h2>本周提示</h2><ul class="notes">{{NOTES}}</ul></div>
<div class="grid2">
 <div class="card"><h2>每周 SO</h2><div class="legend"><span><i style="background:var(--series-1)"></i>公开市场</span>
 <span><i style="background:var(--series-2)"></i>运营商</span></div><div class="chart" id="c-so"></div>
 <details><summary>表格</summary><div id="t-so"></div></details></div>
 <div class="card"><h2>渠道库存</h2><div class="legend"><span><i style="background:var(--series-1)"></i>公开市场</span>
 <span><i style="background:var(--series-2)"></i>运营商</span></div><div class="chart" id="c-stock"></div>
 <details><summary>表格</summary><div id="t-stock"></div></details></div>
</div>
<div class="card"><h2>分渠道</h2><table><thead><tr><th>渠道</th><th>类型</th><th class="num">本周 SO</th><th class="num">环比</th>
<th class="num">近 4 周均</th><th class="num">对预测</th><th class="num">库存</th><th class="num">够卖天数</th><th>状态</th></tr></thead>
<tbody>{{CHANNELS}}</tbody></table></div>
<div class="grid2">
 <div class="card"><h2>分系列</h2><table><thead><tr><th>系列</th><th class="num">本周 SO</th><th class="num">环比</th>
 <th class="num">占比</th></tr></thead><tbody>{{SERIES}}</tbody></table></div>
 <div class="card"><h2>本周涨得最多</h2><table><thead><tr><th>机型</th><th>代码</th><th class="num">上周</th><th class="num">本周</th>
 <th class="num">变化</th></tr></thead><tbody>{{UPS}}</tbody></table>
 <h2 style="margin-top:16px">本周跌得最多</h2><table><thead><tr><th>机型</th><th>代码</th><th class="num">上周</th>
 <th class="num">本周</th><th class="num">变化</th></tr></thead><tbody>{{DOWNS}}</tbody></table></div>
</div>
<p class="foot">够卖天数 = 本周渠道库存 ÷ 近 4 周平均 SO × 7。超过 {{DOS_HIGH}} 天标"偏高"，低于 {{DOS_LOW}} 天标"风险"（阈值待确认）。
本周还没填的渠道不计入合计和环比。数值来自 PSI 里手填的格子，公式格（未来周的预测）不算。内含 L3 数据，只在公司渠道内传阅。</p>
</div></div>
<script>
const D={{DATA}};
function css(v){return getComputedStyle(document.querySelector('.viz-root')).getPropertyValue(v).trim()}
function fmt(x){return Math.round(x).toLocaleString('en-US')}
function line(id, key){
  const box=document.getElementById(id), W=560, H=230, L=52, R=12, T=12, B=28;
  const series=[['公开市场','--series-1'],['运营商','--series-2']];
  const vals=series.flatMap(([n])=>D.trend[n][key]); const max=Math.max(...vals,1)*1.1;
  const x=i=>L+(W-L-R)*i/(D.weeks.length-1), y=v=>T+(H-T-B)*(1-v/max);
  const ns='http://www.w3.org/2000/svg', svg=document.createElementNS(ns,'svg'); svg.setAttribute('viewBox',`0 0 ${W} ${H}`);
  svg.setAttribute('role','img'); svg.setAttribute('aria-label',(key==='so'?'每周 SO':'渠道库存')+'，公开市场和运营商');
  const el=(t,a)=>{const e=document.createElementNS(ns,t);for(const k in a)e.setAttribute(k,a[k]);svg.appendChild(e);return e};
  for(let k=0;k<=4;k++){const v=max/1.1*k/4, yy=y(v); el('line',{x1:L,x2:W-R,y1:yy,y2:yy,stroke:css('--grid'),'stroke-width':1});
    const t=el('text',{x:L-6,y:yy+4,'text-anchor':'end','font-size':11,fill:css('--text-muted')}); t.textContent=fmt(v)}
  D.weeks.forEach((w,i)=>{if(i%2===0||i===D.weeks.length-1){const t=el('text',{x:x(i),y:H-8,'text-anchor':'middle','font-size':11,fill:css('--text-muted')});t.textContent='W'+w}});
  series.forEach(([n,c])=>{const d=D.trend[n][key].map((v,i)=>(i?'L':'M')+x(i)+','+y(v)).join('');
    el('path',{d,fill:'none',stroke:css(c),'stroke-width':2,'stroke-linejoin':'round','stroke-linecap':'round'});
    const last=D.trend[n][key].length-1, v=D.trend[n][key][last];
    el('circle',{cx:x(last),cy:y(v),r:4,fill:css(c),stroke:css('--surface-1'),'stroke-width':2})});
  const hair=el('line',{y1:T,y2:H-B,stroke:css('--text-muted'),'stroke-width':1,visibility:'hidden'});
  const hit=el('rect',{x:L,y:T,width:W-L-R,height:H-T-B,fill:'transparent',tabindex:0});
  box.appendChild(svg); const tip=document.createElement('div'); tip.className='tip'; box.appendChild(tip);
  function show(i){hair.setAttribute('x1',x(i));hair.setAttribute('x2',x(i));hair.setAttribute('visibility','visible');
    tip.replaceChildren(); const h=document.createElement('div'); h.textContent='W'+D.weeks[i]; h.style.color=css('--text-secondary'); tip.appendChild(h);
    series.forEach(([n,c])=>{const r=document.createElement('div'),k=document.createElement('i'),b=document.createElement('b');
      k.style.background=css(c); b.textContent=fmt(D.trend[n][key][i]); r.append(k,b,document.createTextNode('  '+n)); tip.appendChild(r)});
    const bb=svg.getBoundingClientRect(), px=x(i)/W*bb.width; tip.style.display='block';
    tip.style.left=Math.min(px+12, bb.width-tip.offsetWidth)+'px'; tip.style.top='8px'}
  hit.addEventListener('pointermove',e=>{const bb=svg.getBoundingClientRect(),px=(e.clientX-bb.left)/bb.width*W;
    show(Math.max(0,Math.min(D.weeks.length-1,Math.round((px-L)/(W-L-R)*(D.weeks.length-1)))))});
  hit.addEventListener('pointerleave',()=>{tip.style.display='none';hair.setAttribute('visibility','hidden')});
  let fi=D.weeks.length-1; hit.addEventListener('focus',()=>show(fi));
  hit.addEventListener('keydown',e=>{if(e.key==='ArrowLeft')fi=Math.max(0,fi-1);if(e.key==='ArrowRight')fi=Math.min(D.weeks.length-1,fi+1);show(fi)});
  hit.addEventListener('blur',()=>{tip.style.display='none';hair.setAttribute('visibility','hidden')});
  // 表格视图
  const tb=document.createElement('table'), hd=tb.insertRow(); ['周','公开市场','运营商'].forEach(t=>{const c=document.createElement('th');c.textContent=t;hd.appendChild(c)});
  D.weeks.forEach((w,i)=>{const r=tb.insertRow(); r.insertCell().textContent='W'+w;
    series.forEach(([n])=>{const c=r.insertCell(); c.className='num'; c.textContent=fmt(D.trend[n][key][i])})});
  document.getElementById('t-'+id.slice(2)).appendChild(tb);
}
function draw(){['c-so','c-stock'].forEach(id=>{document.getElementById(id).replaceChildren();document.getElementById('t-'+id.slice(2)).replaceChildren()});
  line('c-so','so'); line('c-stock','stock')}
draw(); matchMedia('(prefers-color-scheme: dark)').addEventListener('change',draw);
</script></body></html>
"""


def latest_week(d):
    counts = d[d.lab == "SO"].groupby(["w", "sheet"]).size().unstack(fill_value=0)
    ok = (counts >= 5).sum(axis=1)
    return int(ok[ok >= len(OM + OP) / 2].index.max())


def main():
    ap = argparse.ArgumentParser(description="生成 PSI 周度总览看板（本地 HTML）")
    ap.add_argument("--week", type=int, help="默认：PSI 里一半以上渠道已经填了 SO 的最近一周")
    ap.add_argument("--psi", default=str(PSI_FILE))
    ap.add_argument("--out", help="输出文件夹；默认 W{周}/，没有这个文件夹就放 output/")
    a = ap.parse_args()
    d = extract(a.psi)
    week = a.week or latest_week(d)
    m = build(d, db_stock(a.psi), week)
    out = Path(a.out) if a.out else (ROOT / f"W{week}" if (ROOT / f"W{week}").is_dir() else OUTPUT)
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"W{week}_周度总览.html"
    path.write_text(render(m, Path(a.psi).name), encoding="utf-8")
    print(path)


if __name__ == "__main__":
    main()
