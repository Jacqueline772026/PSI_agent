"""PSI 看板的其余几页：库存健康、预警清单、新品与退市（关联 Roadmap）、预测准确度、数据质量。

每个函数接收 tools.dashboard.extract() 的长表 d（sheet / ch / code / name / series / lab / w / v），返回一段 HTML。
"""
import datetime as dt
import html
import re
from pathlib import Path

import openpyxl
import pandas as pd

from checks.structure_check import COL_E, COL_H, load
from psi_tool.mapping import canonical
from psi_tool.paths import DATA
from psi_tool.sources import week_range
from tools.learn_mappings import name_key

OM = ["EURO", "MEX", "MSH", "X-KOM", "Michain", "OM Others"]
OP = ["Play", "TMO", "Plus", "Orange"]
SERIES = {"RM Note": "Redmi Note", "Redmi": "Redmi", "xiaomi": "Xiaomi", "POCO": "POCO", "pad": "Pad"}
DOS_LOW, DOS_OK, DOS_HIGH = 21, 60, 120  # <21 断货风险；21–60 健康；60–120 偏高；>120 压货（阈值待和导师确认）
DOS_TARGET = 60  # 算"超出目标的库存"时用的目标天数
ALERT_MIN_STOCK = 1000  # 压货清单只列渠道库存 ≥ 1000 台的机型
STOCKOUT_MIN_SO = 30  # 断货清单只列近 4 周周均 SO ≥ 30 台的
FCST_FLAG = 0.3  # 预测偏差超过 ±30% 标出来
ROADMAP_DIR = DATA / "参考" / "Roadmap"
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}

esc = html.escape
BADGE = {"critical": ("●", "断货风险"), "good": ("✔", "健康"), "warning": ("⚠", "偏高"), "serious": ("▲", "压货"),
         "info": ("ℹ", "提示"), "missing": ("—", "未填")}


def fmt(x):
    return "—" if x is None or pd.isna(x) else f"{x:,.0f}"


def fpct(x):
    return "—" if x is None or pd.isna(x) else f"{x:+.0%}"


def badge(kind, label=None):
    icon, text = BADGE[kind]
    return f'<span class="status {kind}"><span aria-hidden="true">{icon}</span> {esc(label or text)}</span>'


def dos_kind(dos):
    if dos is None or pd.isna(dos):
        return None
    return "critical" if dos < DOS_LOW else "good" if dos <= DOS_OK else "warning" if dos <= DOS_HIGH else "serious"


def filled_sheets(d, week):
    c = d[(d.lab == "SO") & (d.w == week)].groupby("sheet").size()
    return [s for s in OM + OP if c.get(s, 0) >= 5]


def per_code(d, week, sheets, by=("code", "name")):
    """近 4 周周均 SO、本周库存、DOS。"""
    f = d[d.sheet.isin(sheets)]
    so4 = f[(f.lab == "SO") & f.w.between(week - 3, week)].groupby(list(by)).v.sum() / 4
    stock = f[(f.lab == "库存") & (f.w == week)].groupby(list(by)).v.sum()
    t = pd.DataFrame({"avg4": so4, "stock": stock}).fillna(0)
    t["dos"] = t.stock / t.avg4.where(t.avg4 > 0) * 7
    return t


def spark(values, w=110, h=26):
    """近 12 周 SO 的小折线（单一系列，颜色用 series-1）。"""
    vals = [0 if pd.isna(v) else v for v in values]
    mx = max(vals) or 1
    pts = " ".join(f"{2 + i * (w - 4) / max(len(vals) - 1, 1):.1f},{h - 3 - v / mx * (h - 6):.1f}" for i, v in enumerate(vals))
    return (f'<svg class="spark" viewBox="0 0 {w} {h}" width="{w}" height="{h}" role="img" aria-label="近 {len(vals)} 周 SO">'
            f'<polyline points="{pts}" fill="none" stroke="var(--series-1)" stroke-width="1.6" stroke-linejoin="round"/></svg>')


def weekly_so(d, sheets, week, n=12):
    f = d[(d.lab == "SO") & d.sheet.isin(sheets) & d.w.between(week - n + 1, week)]
    return f.pivot_table(index="code", columns="w", values="v", aggfunc="sum").reindex(
        columns=range(week - n + 1, week + 1)).fillna(0)


# ---------- B. 库存健康 ----------

def page_inventory(d, week):
    sheets = filled_sheets(d, week)
    t = per_code(d, week, sheets, by=("sheet", "series"))
    tot = per_code(d, week, sheets, by=("sheet",))
    cols = list(SERIES)
    head = "".join(f"<th class='num'>{esc(SERIES[c])}</th>" for c in cols)
    rows = []
    for s in OM + OP:
        if s not in sheets:
            rows.append(f"<tr class='muted'><td>{esc(s)}</td><td colspan='{len(cols) + 2}'>本周还没填</td></tr>")
            continue
        cells = []
        for c in cols:
            if (s, c) in t.index and t.loc[(s, c), "stock"] > 0:
                r = t.loc[(s, c)]
                k = dos_kind(r.dos)
                title = f"{s} · {SERIES[c]}：库存 {fmt(r.stock)}，周均 SO {fmt(r.avg4)}，够卖 {fmt(r.dos)} 天"
                cells.append(f"<td class='num heat {k or ''}' title='{esc(title)}' tabindex='0'>{fmt(r.dos)}"
                             f"<small>{fmt(r.stock)}</small></td>")
            else:
                cells.append("<td class='num muted'>—</td>")
        r = tot.loc[s]
        rows.append(f"<tr><td>{esc(s)}</td>{''.join(cells)}<td class='num'><b>{fmt(r.dos)}</b><small>{fmt(r.stock)}</small></td>"
                    f"<td>{badge(dos_kind(r.dos)) if dos_kind(r.dos) else ''}</td></tr>")
    legend = " ".join(badge(k, f"{lab}") for k, lab in (("critical", f"< {DOS_LOW} 天 断货风险"), ("good", f"{DOS_LOW}–{DOS_OK} 天 健康"),
                                                          ("warning", f"{DOS_OK}–{DOS_HIGH} 天 偏高"), ("serious", f"> {DOS_HIGH} 天 压货")))
    return f"""
<div class="card"><h2>渠道 × 系列：库存够卖天数</h2>
<p class="hint">格子里大字是够卖天数（库存 ÷ 近 4 周周均 SO × 7），小字是库存台数。鼠标停在格子上看明细。</p>
<div class="legend-row">{legend}</div>
<div class="scroll"><table class="matrix"><thead><tr><th>渠道</th>{head}<th class="num">合计</th><th>状态</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table></div></div>"""


# ---------- C. 预警清单 ----------

def page_alerts(d, week):
    sheets = filled_sheets(d, week)
    # 同一代码在不同 sheet 里名称可能不一样（Redmi 17 / Redmi 17 5G），按代码汇总，名称取最常见的
    names = d[d.code != "Others"].groupby("code").name.agg(lambda x: x.mode().iat[0])
    code = per_code(d, week, sheets, by=("code",))
    code = code[code.index != "Others"]
    code.index = pd.MultiIndex.from_arrays([code.index, code.index.map(names)], names=["code", "name"])
    ka = per_code(d, week, sheets, by=("sheet", "code", "name"))
    spk = weekly_so(d, sheets, week)

    over = code[(code.stock >= ALERT_MIN_STOCK) & (code.dos > DOS_HIGH)].copy()
    over["excess"] = over.stock - over.avg4 * DOS_TARGET / 7
    over = over.sort_values("excess", ascending=False)
    rows = []
    for (c, n), r in over.iterrows():
        where = ka.xs(c, level="code")
        where = where[where.stock > 0].sort_values("stock", ascending=False).head(3)
        where = "、".join(f"{s} {fmt(v)}" for (s, _), v in where.stock.items())
        rows.append(f"<tr><td>{esc(str(n))}<div class='code'>{esc(c)}</div></td>"
                    f"<td>{spark(spk.loc[c]) if c in spk.index else ''}</td><td class='num'>{fmt(r.avg4)}</td>"
                    f"<td class='num'>{fmt(r.stock)}</td><td class='num'><b>{fmt(r.dos)}</b></td>"
                    f"<td class='num'>{fmt(r.excess)}</td><td class='small'>{esc(where)}</td></tr>")
    over_html = "".join(rows) or "<tr><td colspan='7'>没有符合条件的机型。</td></tr>"

    out = ka[(ka.avg4 >= STOCKOUT_MIN_SO) & (ka.dos < DOS_LOW)].sort_values("dos")
    rows = [f"<tr><td>{esc(s)}</td><td>{esc(str(n))}<div class='code'>{esc(c)}</div></td><td class='num'>{fmt(r.avg4)}</td>"
            f"<td class='num'>{fmt(r.stock)}</td><td class='num'><b>{fmt(r.dos)}</b></td></tr>"
            for (s, c, n), r in out.iterrows() if c != "Others"]
    out_html = "".join(rows) or "<tr><td colspan='5'>没有符合条件的组合。</td></tr>"

    zero = ka[(ka.stock == 0) & (ka.avg4 >= STOCKOUT_MIN_SO)]
    rows = [f"<tr><td>{esc(s)}</td><td>{esc(str(n))}<div class='code'>{esc(c)}</div></td><td class='num'>{fmt(r.avg4)}</td></tr>"
            for (s, c, n), r in zero.iterrows() if c != "Others"]
    zero_html = "".join(rows)
    return f"""
<div class="card"><h2>压货：库存高、卖得慢</h2>
<p class="hint">全部已填渠道合计。条件：渠道库存 ≥ {ALERT_MIN_STOCK:,} 台，且够卖超过 {DOS_HIGH} 天。
"超出目标"= 库存 − 按 {DOS_TARGET} 天目标应有的库存，按这一列从大到小排。可以拿来讨论促销、调拨或减少 SI。</p>
<div class="scroll"><table><thead><tr><th>机型</th><th>近 12 周 SO</th><th class="num">周均 SO</th><th class="num">库存</th>
<th class="num">够卖天数</th><th class="num">超出目标</th><th>库存最多的渠道</th></tr></thead><tbody>{over_html}</tbody></table></div></div>
<div class="card"><h2>断货风险：卖得快、库存少</h2>
<p class="hint">按渠道 × 机型。条件：近 4 周周均 SO ≥ {STOCKOUT_MIN_SO} 台，且库存够卖不到 {DOS_LOW} 天。</p>
<div class="scroll"><table><thead><tr><th>渠道</th><th>机型</th><th class="num">周均 SO</th><th class="num">库存</th>
<th class="num">够卖天数</th></tr></thead><tbody>{out_html}</tbody></table></div>
{f'<h2 style="margin-top:16px">在卖但库存已经是 0</h2><table><thead><tr><th>渠道</th><th>机型</th><th class="num">周均 SO</th></tr></thead><tbody>{zero_html}</tbody></table>' if zero_html else ''}
</div>"""


# ---------- D. 新品与退市（Roadmap） ----------

def latest_roadmap():
    files = sorted(ROADMAP_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime) if ROADMAP_DIR.exists() else []
    return files[-1] if files else None


def parse_roadmap(path, psi_path):
    """Roadmap 的 Smartphone / Tab 页 → 每个机型一行：名称、零售价、型号、PSI 代码、每月状态。"""
    psi = load(psi_path)
    codes, names = set(), {}
    for s in OM + OP:
        sh = psi[s]
        for r in sh.product_rows("SO"):
            c = sh.code(r)
            codes.add(c)
            e = sh.get(r, COL_E)
            if isinstance(e, str) and not e.startswith("="):
                names.setdefault((name_key(e)[0], str(sh.get(r, COL_H)).strip()), c)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for sheet in ("Smartphone", "Tab"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        ws.reset_dimensions()
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        h = next((i for i, r in enumerate(rows[:6]) if any(str(x).strip().upper() == "MODEL" for x in r if x)), None)
        if h is None:
            continue
        head = [str(x).strip() if x else "" for x in rows[h]]
        c_model, c_name = head.index("MODEL"), head.index("PRODUCT NAME")
        c_price = c_name - 1
        # 月份列：表头那一行的月份名，从第一个年份（表头上一行）开始，过了 12 月年份 +1
        year_row = rows[h - 1] if h else []
        year = next((int(re.match(r"(20\d\d)", str(x)).group(1)) for x in year_row if x and re.match(r"20\d\d", str(x))), 2026)
        months, last = {}, None
        for j, x in enumerate(head):
            m = MONTHS.get(x.strip().lower()[:3]) if x else None
            if m:
                if last and m < last:
                    year += 1
                months[j], last = dt.date(year, m, 1), m
        series = None
        for i in range(h + 1, len(rows)):
            r = rows[i] + [None] * 30
            model = str(r[c_model] or "").strip()
            if r[0]:
                series = str(r[0]).strip()
            if not model or not re.search(r"\d+\s*\+\s*\d+", model):
                continue
            m = re.match(r"(.+?)\s*(\d+\s*\+\s*\d+)\s*$", model)
            code = canonical(m.group(1).replace(" ", "") + m.group(2).replace(" ", "")) if m else model
            name = str(r[c_name] or "").strip()
            if code not in codes and name and name.upper() != "TBD":
                code = names.get((name_key(name)[0], m.group(2).replace(" ", "") if m else ""), code)
            events = [(dt_, str(r[j]).strip()) for j, dt_ in months.items() if r[j] not in (None, "")]
            price_row = rows[i + 1] + [None] * 30 if i + 1 < len(rows) else []
            price_notes = [(dt_, str(price_row[j]).strip()) for j, dt_ in months.items()
                           if price_row and price_row[j] not in (None, "")]
            out.append(dict(sheet=sheet, series=series, rrp=r[c_price], name=name if name.upper() != "TBD" else "",
                            model=model, code=code, in_psi=code in codes, events=events, price=price_notes))
    return out


def classify_event(text):
    t = text.lower()
    if "eol" in t:
        return "EOL"
    if "last buy" in t:
        return "last buy"
    if "clear" in t:
        return "清库存"
    if t.startswith("so"):
        return "SO"
    if t.startswith("si"):
        return "SI"
    return text


def page_roadmap(d, week, roadmap, roadmap_name):
    if not roadmap:
        return "<div class='card'><h2>新品与退市</h2><p>没找到 Roadmap 文件（数据/参考/Roadmap/*.xlsx）。</p></div>"
    sheets = filled_sheets(d, week)
    stats = per_code(d, week, sheets, by=("code",))
    spk = weekly_so(d, sheets, week)
    _, today = week_range(week)
    exiting, launching = [], []
    for m in roadmap:
        ev = sorted((t, classify_event(x)) for t, x in m["events"])
        end = next(((t, k) for t, k in ev if k in ("last buy", "EOL")), None)
        start = next(((t, k) for t, k in ev if k in ("SI", "SO")), None)
        s = stats.loc[m["code"]] if m["code"] in stats.index else None
        row = dict(m=m, ev=ev, s=s)
        if end:
            row["end"] = end
            weeks_left = max((end[0] - today).days / 7, 0)
            row["weeks_left"] = weeks_left
            if s is not None:
                row["left_at_end"] = s.stock - s.avg4 * weeks_left
            exiting.append(row)
        elif start:
            row["start"] = start
            launching.append(row)

    def timeline(ev):
        return " → ".join(f"{t:%y-%m} {esc(k)}" for t, k in ev) or "—"

    def price(m):
        return "；".join(f"{t:%y-%m} {esc(x)}" for t, x in m["price"]) or ""

    rows = []
    for r in sorted(exiting, key=lambda r: -(r.get("left_at_end") or -1e9)):
        m, s = r["m"], r["s"]
        left = r.get("left_at_end")
        flag = badge("serious", "到期会剩货") if left and left > 300 else (badge("good", "能卖完") if left is not None else "")
        rows.append(f"<tr><td>{esc(m['name'] or m['model'])}<div class='code'>{esc(m['model'])} → {esc(m['code'])}"
                    f"{'' if m['in_psi'] else '（PSI 无此行）'}</div></td><td class='num'>{esc(str(m['rrp'] or ''))}</td>"
                    f"<td>{r['end'][0]:%Y-%m} {esc(r['end'][1])}</td><td class='num'>{r['weeks_left']:.0f}</td>"
                    f"<td>{spark(spk.loc[m['code']]) if m['code'] in spk.index else ''}</td>"
                    f"<td class='num'>{fmt(s.avg4) if s is not None else '—'}</td><td class='num'>{fmt(s.stock) if s is not None else '—'}</td>"
                    f"<td class='num'><b>{fmt(left)}</b></td><td>{flag}</td><td class='small'>{price(m)}</td></tr>")
    exit_html = "".join(rows) or "<tr><td colspan='10'>Roadmap 里没有 last buy / EOL 的机型。</td></tr>"

    rows = []
    for r in sorted(launching, key=lambda r: r["start"][0]):
        m = r["m"]
        rows.append(f"<tr><td>{esc(m['name'] or 'TBD')}<div class='code'>{esc(m['model'])} → {esc(m['code'])}</div></td>"
                    f"<td>{esc(m['series'] or '')}</td><td class='num'>{esc(str(m['rrp'] or ''))}</td>"
                    f"<td class='small'>{timeline(r['ev'])}</td><td class='small'>{price(m)}</td></tr>")
    launch_html = "".join(rows) or "<tr><td colspan='5'>—</td></tr>"
    return f"""
<div class="card"><h2>即将退市：到 last buy / EOL 时还会剩多少渠道库存</h2>
<p class="hint">按现在的卖速推算：到期剩余 = 本周渠道库存 − 近 4 周周均 SO × 距到期的周数。剩得多的（> 300 台）要提前安排促销或调拨。
到期时间和价格调整来自 Roadmap（{esc(roadmap_name)}）；Roadmap 的旧代码（Q16 等）已换成 PSI 代码。</p>
<div class="scroll"><table><thead><tr><th>机型</th><th class="num">零售价 PLN</th><th>到期</th><th class="num">还剩几周</th>
<th>近 12 周 SO</th><th class="num">周均 SO</th><th class="num">渠道库存</th><th class="num">到期剩余</th><th>判断</th><th>价格调整</th>
</tr></thead><tbody>{exit_html}</tbody></table></div></div>
<div class="card"><h2>即将上市</h2>
<p class="hint">Roadmap 里有 SI / SO 计划、还没有退市计划的机型。上市后会出现在 PSI 里，可以在"周度总览"跟踪起量。</p>
<div class="scroll"><table><thead><tr><th>机型</th><th>系列</th><th class="num">零售价 PLN</th><th>时间线</th><th>价格调整</th></tr></thead>
<tbody>{launch_html}</tbody></table></div></div>"""


# ---------- E. 预测准确度 ----------

def page_forecast(d, week, n=8):
    weeks = list(range(week - n + 1, week + 1))
    so, fc = d[d.lab == "SO"], d[d.lab == "预测"]
    rows = []
    for s in OM + OP:
        cells, biases = [], []
        for w in weeks:
            a = so[(so.sheet == s) & (so.w == w)]
            f = fc[(fc.sheet == s) & (fc.w == w)]
            if len(a) < 5 or f.v.sum() == 0:
                cells.append("<td class='num muted'>—</td>")
                continue
            covered = a[a.code.isin(f[f.v > 0].code)].v.sum()
            bias = covered / f.v.sum() - 1
            biases.append(bias)
            k = "serious" if abs(bias) > FCST_FLAG else ""
            cells.append(f"<td class='num {k}' title='实际（有预测的机型）{fmt(covered)} / 预测 {fmt(f.v.sum())}'>{fpct(bias)}</td>")
        a = so[(so.sheet == s) & (so.w == week)]
        f = fc[(fc.sheet == s) & (fc.w == week)]
        cov = a[a.code.isin(f[f.v > 0].code)].v.sum() / a.v.sum() if a.v.sum() else None
        avg = sum(biases) / len(biases) if biases else None
        rows.append(f"<tr><td>{esc(s)}</td>{''.join(cells)}<td class='num'><b>{fpct(avg)}</b></td>"
                    f"<td class='num'>{'—' if cov is None else f'{cov:.0%}'}</td></tr>")
    head = "".join(f"<th class='num'>W{w}</th>" for w in weeks)
    return f"""
<div class="card"><h2>SO 预测准确度（实际 ÷ 预测 − 1）</h2>
<p class="hint">只比有预测的机型：正数 = 卖得比预测多，负数 = 卖得比预测少。超过 ±{FCST_FLAG:.0%} 的标色。
"本周预测覆盖"= 有预测的机型的 SO 占本周 SO 的比例；覆盖低说明 KAM 很多机型没填预测，偏差参考价值有限。</p>
<div class="scroll"><table><thead><tr><th>渠道</th>{head}<th class="num">平均偏差</th><th class="num">本周预测覆盖</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table></div></div>"""


# ---------- F. 数据质量 ----------

def page_quality(d, week, n=8):
    weeks = list(range(week - n + 1, week + 1))
    so, inv = d[d.lab == "SO"], d[d.lab == "库存"]
    rows = []
    for s in OM + OP:
        cells = []
        for w in weeks:
            a, b = so[(so.sheet == s) & (so.w == w)], inv[(inv.sheet == s) & (inv.w == w)]
            ok = len(a) >= 5 and len(b) >= 5
            cells.append(f"<td class='num'>{badge('good', '已填') if ok else badge('missing')}</td>")
        cur = so[(so.sheet == s) & (so.w == week)]
        oth = cur[cur.code == "Others"].v.sum() / cur.v.sum() if cur.v.sum() else None
        rows.append(f"<tr><td>{esc(s)}</td>{''.join(cells)}<td class='num'>{'—' if oth is None else f'{oth:.0%}'}</td></tr>")
    head = "".join(f"<th class='num'>W{w}</th>" for w in weeks)
    return f"""
<div class="card"><h2>填报情况</h2>
<p class="hint">某周 SO 和库存都有数值（至少 5 格）算"已填"。Others 占比 = Others 行 SO 占这家本周 SO 的比例；
占比高说明很多型号没对上 PSI 的行，可能是新品没加行或映射缺失。库存恒等式等更细的检查见每周的核对报告（W{{周}}_report.md）。</p>
<div class="scroll"><table><thead><tr><th>渠道</th>{head}<th class="num">本周 Others 占比</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table></div></div>"""
