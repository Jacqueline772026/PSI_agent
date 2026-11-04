"""Auchan 库存口径改成 门店 + 仓库 之后，算出 PSI 历史周要改哪些格子。

PSI「OM Others」的 Channel Inventory Actual 里，历史上只含 Auchan 的门店库存。
这里按周、按代码算出 Auchan 的仓库库存（Zapas Magazynowy Ilość），列出 修正后 = PSI 现值 + 仓库库存。
同时用 KTR + MAX + Auchan 门店 的合计核验 PSI 那一周是否确实只含门店，核验不过的周不要直接套用。

用法：.venv/bin/python -m tools.auchan_backfill --from-week 1 --to-week 43 --inbox 数据/历史周报
结果：前期资料/分析报告/Auchan仓库库存补算表.xlsx（含公司数据，不进仓库）
"""
import argparse
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

from checks.structure_check import COL_E, COL_LABEL, is_number, load
from psi_tool.mapping import Mapper, latest_ttl
from psi_tool.paths import CONFIG, MAPPING, PSI_FILE, ROOT, TTL_DIR
from psi_tool.run import ACTUAL, psi_layout
from psi_tool.sources import load_config, load_source

OUT = ROOT / "前期资料" / "分析报告" / "Auchan仓库库存补算表.xlsx"


def stock_total(inbox, cfg, name, week, column=None):
    c = dict(cfg[name])
    if column:
        c["stock"] = column
    rows = load_source(inbox, name, c, week).rows
    return None if rows is None else rows


def main():
    ap = argparse.ArgumentParser(description="Auchan 仓库库存历史补算")
    ap.add_argument("--from-week", type=int, default=1)
    ap.add_argument("--to-week", type=int, default=43)
    ap.add_argument("--inbox", default=str(ROOT / "数据" / "历史周报"))
    a = ap.parse_args()

    cfg = load_config(CONFIG)
    sheet = load(PSI_FILE)["OM Others"]
    mapper = Mapper(latest_ttl(TTL_DIR), MAPPING)
    layout = psi_layout(sheet)
    rows_by_code = {c: (n, act) for c, n, _, act in layout}
    summary, cells = [], []
    for week in range(a.from_week, a.to_week + 1):
        col = sheet.week_col(week)
        store = stock_total(a.inbox, cfg, "AUCHAN", week, "Zapas Bieżący Ilość")
        wh = stock_total(a.inbox, cfg, "AUCHAN", week, "Zapas Magazynowy Ilość")
        psi_total = sum(v for _, _, _, r in layout if r for v in [sheet.get(r, col)] if is_number(v))
        if wh is None:
            summary.append({"周": f"W{week}", "核验": "❌ 没有 Auchan 文件，无法补算", "PSI 现合计": int(psi_total)})
            continue
        others = [stock_total(a.inbox, cfg, s, week) for s in ("KTR", "MAX")]
        expected = None if any(x is None for x in others) else int(store["stock"].sum() + sum(x["stock"].sum() for x in others))
        check = ("⚠️ 缺 KTR 或 MAX 文件，未核验" if expected is None else
                 "✅ PSI 只含门店" if expected == int(psi_total) else
                 "❗ PSI 可能已含仓库，不要再加" if int(psi_total) - expected == int(wh["stock"].sum()) else
                 f"⚠️ PSI 和门店口径重算差 {int(psi_total) - expected}（不等于仓库量，是别的差异，如 KTR 用错周、未填的格子）")
        wh = wh[wh["stock"] != 0].copy()
        wh["code"] = [mapper.lookup("EAN", k)[0] for k in wh["key"]]
        wh["psi"] = [c if c in rows_by_code else "Others" for c in wh["code"]]
        for code, qty in wh.groupby("psi")["stock"].sum().items():
            name, r = rows_by_code[code]
            now = sheet.get(r, col)
            cells.append({"周": f"W{week}", "PSI 单元格": f"{get_column_letter(col + 1)}{r}", "代码": code, "产品": name,
                          "PSI 现值": now if is_number(now) else "（公式，未填）", "加 Auchan 仓库": int(qty),
                          "修正后": int(now + qty) if is_number(now) else None})
        summary.append({"周": f"W{week}", "Auchan 门店": int(store["stock"].sum()), "Auchan 仓库": int(wh["stock"].sum()),
                        "PSI 现合计": int(psi_total), "PSI 修正后合计": int(psi_total + wh["stock"].sum()), "核验": check})
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        pd.DataFrame(summary).to_excel(xw, sheet_name="按周总览", index=False)
        pd.DataFrame(cells).to_excel(xw, sheet_name="逐格修改", index=False)
    print(pd.DataFrame(summary).to_string(index=False))
    print(f"\n{len(cells)} 格：{OUT}")


if __name__ == "__main__":
    main()
