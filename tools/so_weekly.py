"""公开市场（OM）SO 周报 / 月报：给"波兰公开市场 SO预实分析"每周那一页准备数。

从 PSI 的 OM 各 KA 页（EURO / MEX / X-KOM / MSH / Michain / OM Others）汇总，等于 PSI 'OM' 页的 SO，
不用等飞书的 IMPORTRANGE 算出来。当周可以用工具结果覆盖（PSI 还没填或填错时，比如 W44 的 KTR 填反）。

输出 W{周}_SO周报.xlsx：
  - 本周（和预实页同样的结构）：每个代码一行，W{起}…W{周} 的 SO 趋势、WoW、近 4 周均值、
    分渠道（MEX / EURO / MSH / X-KOM / Michain / OM Others）本周和上周、OM 库存、DOS、MEX 促销价，
    以及"事实提示"（库存变化、DOS 过低、渠道 SO 大变、促销价变化），写原因备注时参考
  - 系列汇总：ALL 和 Redmi / RM Note / Xiaomi / POCO / Pad 每周 SO
  - 月度：按月汇总 SO（周按周四所在的月份归月），环比
用法：
  .venv/bin/python -m tools.so_weekly --week 44 [--psi PSI.xlsx] [--inbox W44测试/W44原始文件] [--out W44测试/工具结果]
"""
import argparse
import datetime as dt
from pathlib import Path

import pandas as pd

from checks.structure_check import COL_C, COL_E, COL_LABEL, is_number, load
from psi_tool.paths import CONFIG, MAPPING, PSI_FILE, TTL_DIR

KA = ["MEX", "EURO", "MSH", "X-KOM", "Michain", "OM Others"]
SERIES = {"redmi": "Redmi", "rm note": "RM Note", "xiaomi": "Xiaomi", "poco": "POCO", "pad": "Pad", "poco_pad": "Pad"}


def num(x):
    return x if is_number(x) else None  # 公式 = 还没填（PSI 里会显示 FCST 或 0），按没数处理


def week_month(week, year=2026):
    thursday = dt.date.fromisocalendar(year, week, 4)
    return thursday.month


def collect(psi, weeks):
    """{(sheet, code): {"so": {w: n}, "stock": {w: n}, "price": {w: n}}}，以及 OM 页的系列 / 名称。"""
    data = {}
    for name in KA:
        s = psi[name]
        blocks = {}
        for r in s.product_rows():
            code = s.code(r)
            lab = s.get(r, COL_LABEL)
            if code:
                blocks.setdefault(code, {})[lab] = r
        for r in s.others_rows:
            blocks.setdefault("Others", {})[s.get(r, COL_LABEL)] = r
        for code, b in blocks.items():
            d = data.setdefault((name, code), {"so": {}, "stock": {}, "price": {}})
            for w in weeks:
                c = s.week_col(w)
                if "SO" in b:
                    d["so"][w] = num(s.get(b["SO"], c))
                    if name == "MEX" and is_number(s.get(b["SO"] - 1, c)) and s.get(b["SO"] - 1, COL_LABEL) == "Promo Price":
                        d["price"][w] = s.get(b["SO"] - 1, c)
                if "Channel Inventory Actual" in b:
                    d["stock"][w] = num(s.get(b["Channel Inventory Actual"], c))
    om = psi["OM"]
    info = {}
    for r in om.product_rows("SO"):
        code = om.code(r)
        if code:
            info.setdefault(code, (SERIES.get(str(om.get(r, COL_C)).strip().lower(), str(om.get(r, COL_C))), om.get(r, COL_E)))
    return data, info


def override_current(data, week, inbox, psi_path):
    """当周用工具结果（只覆盖周报收齐的那几页）。"""
    from psi_tool.run import build, summarize
    results, _ = build(week, inbox, psi_path, CONFIG, MAPPING, TTL_DIR)
    used = []
    for name in KA:
        res = results.get(name)
        if not res or any(not p.files for p in res["parts"]):
            continue
        t = summarize(res)
        for x in t.itertuples():
            d = data.setdefault((name, x.code), {"so": {}, "stock": {}, "price": {}})
            d["so"][week], d["stock"][week] = x.so, x.stock
        used.append(name)
    return used


def build_tables(data, info, week, first):
    weeks = list(range(first, week + 1))
    codes = [c for c in info] + ["Others"]
    rows = []
    for code in codes:
        series, pname = info.get(code, ("Others", "Others"))
        so = {w: sum((data.get((k, code), {}).get("so", {}).get(w) or 0) for k in KA) for w in weeks}
        if not any(so[w] for w in weeks[-8:]) and not any((data.get((k, code), {}).get("stock", {}).get(week) or 0) for k in KA):
            continue  # 近 8 周没销量、本周也没库存的老品不列
        cur, prev = so[week], so.get(week - 1, 0)
        avg4 = sum(so.get(w, 0) for w in range(week - 4, week)) / 4
        stock = sum((data.get((k, code), {}).get("stock", {}).get(week) or 0) for k in KA)
        stock_prev = sum((data.get((k, code), {}).get("stock", {}).get(week - 1) or 0) for k in KA)
        dos = round(stock / cur * 7) if cur > 0 else None
        price = data.get(("MEX", code), {}).get("price", {})
        p_now, p_prev = price.get(week), price.get(week - 1)
        row = {"系列": series, "代码": code, "产品": pname}
        row.update({f"W{w}": so[w] for w in weeks})
        row.update({"WoW": (cur / prev - 1) if prev else None, "近4周均值": round(avg4),
                    "本周 vs 4周均": (cur / avg4 - 1) if avg4 else None})
        facts = []
        for k in KA:
            a = data.get((k, code), {}).get("so", {})
            n, p = a.get(week) or 0, a.get(week - 1) or 0
            row[f"{k} W{week - 1}"], row[f"{k} W{week}"] = p, n
            if abs(n - p) >= 30 and (p == 0 or abs(n / p - 1) >= 0.3):
                facts.append(f"{k} {p}→{n}")
            ks, kp = (data.get((k, code), {}).get("stock", {}).get(w) or 0 for w in (week, week - 1))
            if kp >= 50 and ks <= 0.2 * kp:
                facts.append(f"{k} 库存 {kp}→{ks}（快断货）")
            elif ks >= 100 and ks >= 2 * max(kp, 1):
                facts.append(f"{k} 库存 {kp}→{ks}（补货）")
        row.update({"OM 库存": stock, "上周 OM 库存": stock_prev, "DOS（天）": dos,
                    "MEX 促销价": p_now, "MEX 上周促销价": p_prev})
        if p_now and p_prev and p_now != p_prev:
            facts.insert(0, f"MEX 促销价 {p_prev}→{p_now}")
        if dos is not None and dos < 14 and cur >= 20:
            facts.append(f"DOS 只有 {dos} 天")
        if cur == 0 and stock > 0 and prev > 0:
            facts.append("本周没销量但有库存")
        row["事实提示（写原因时参考）"] = "；".join(facts)
        rows.append(row)
    detail = pd.DataFrame(rows)
    wk = [f"W{w}" for w in weeks]
    order = ["Redmi", "RM Note", "Xiaomi", "POCO", "Pad", "Others"]
    detail["_o"] = detail["系列"].map(lambda s: order.index(s) if s in order else 9)
    detail = detail.sort_values(["_o"], kind="stable").drop(columns="_o")
    series = detail.groupby("系列", sort=False)[wk].sum()
    series = series.reindex([s for s in order if s in series.index])
    series.loc["ALL"] = series.sum()
    series = series.reindex(["ALL"] + [s for s in series.index if s != "ALL"])
    series["WoW"] = series[f"W{week}"] / series[f"W{week - 1}"] - 1
    return detail, series


def monthly(data, info, week, year=2026):
    weeks = list(range(1, week + 1))
    months = sorted({week_month(w, year) for w in weeks})
    rows = []
    for code in list(info) + ["Others"]:
        series, pname = info.get(code, ("Others", "Others"))
        m = {mo: 0 for mo in months}
        for w in weeks:
            m[week_month(w, year)] += sum((data.get((k, code), {}).get("so", {}).get(w) or 0) for k in KA)
        if not any(m.values()):
            continue
        rows.append({"系列": series, "代码": code, "产品": pname, **{f"{mo}月": m[mo] for mo in months}})
    df = pd.DataFrame(rows)
    cols = [f"{mo}月" for mo in months]
    ser = df.groupby("系列")[cols].sum()
    ser.loc["ALL"] = ser.sum()
    mom = ser[cols].pct_change(axis=1)
    return df, ser, mom, {mo: [w for w in weeks if week_month(w, year) == mo] for mo in months}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--week", type=int, required=True)
    ap.add_argument("--psi", default=str(PSI_FILE))
    ap.add_argument("--inbox", help="当周原始文件夹：给了就用工具结果覆盖当周（周报收齐的页）")
    ap.add_argument("--first", type=int, default=4, help="趋势从第几周开始（预实页从 W4 开始）")
    ap.add_argument("--out", default=".")
    a = ap.parse_args()
    psi = load(a.psi)
    data, info = collect(psi, range(1, a.week + 1))
    used = override_current(data, a.week, a.inbox, a.psi) if a.inbox else []
    detail, series = build_tables(data, info, a.week, a.first)
    mdf, mser, mmom, mweeks = monthly(data, info, a.week)
    out = Path(a.out) / f"W{a.week}_SO周报.xlsx"
    with pd.ExcelWriter(out) as xw:
        series.to_excel(xw, sheet_name="系列汇总")
        detail.to_excel(xw, sheet_name="本周明细", index=False)
        mser.to_excel(xw, sheet_name="月度-系列")
        mmom.to_excel(xw, sheet_name="月度-系列环比")
        mdf.to_excel(xw, sheet_name="月度-代码", index=False)
        pd.DataFrame({"说明": [
            f"数据来源：{Path(a.psi).name} 的 OM 各 KA 页（EURO / MEX / X-KOM / MSH / Michain / OM Others）合计，等于 PSI 'OM' 页的 SO。",
            f"W{a.week} 用工具结果覆盖的页：{'、'.join(used) or '无（全部用 PSI 里填的数）'}。",
            "PSI 里还是公式的格子（没填、显示的是 FCST）按 0 算。",
            "DOS = OM 库存 ÷ 本周 SO × 7。MEX 促销价取 MEX 页 Promo Price 行。",
            "月份：每周按周四所在的月份归月，" + "；".join(f"{m}月 = W{ws[0]}–W{ws[-1]}" for m, ws in mweeks.items()),
            "'事实提示'只是根据数字自动列出的变化，原因（竞品、活动、渠道动作）还要人补。",
        ]}).to_excel(xw, sheet_name="说明", index=False)
    print(out)
    print(series[[f"W{a.week - 1}", f"W{a.week}", "WoW"]].to_string())


if __name__ == "__main__":
    main()
