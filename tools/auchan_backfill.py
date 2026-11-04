"""Auchan 库存口径改成 门店 + 仓库 之后，算出 PSI 历史周要改哪些格子。

PSI「OM Others」的 Channel Inventory Actual 里，W1–W9 已经含 Auchan 仓库（逐代码核对过，不用补），
W10 起只含门店。所以默认从 W10 开始。
这里按周、按代码算出 Auchan 的仓库库存（Zapas Magazynowy Ilość），列出 修正后 = PSI 现值 + 仓库库存。
同时用 KTR + MAX + Auchan 门店 的合计核验 PSI 那一周是否确实只含门店，核验不过的周不要直接套用。

用法：
  .venv/bin/python -m tools.auchan_backfill --force     生成补算表（必须用贴之前最新导出的 PSI）
  .venv/bin/python -m tools.auchan_backfill --verify    贴完重新导出 PSI 后，逐格核对
结果：前期资料/分析报告/Auchan仓库库存补算表.xlsx（含公司数据，不进仓库）
"""
import argparse
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

from checks.structure_check import is_formula, is_number, load
from psi_tool.mapping import Mapper, latest_ttl
from psi_tool.paths import CONFIG, MAPPING, PSI_FILE, ROOT, TTL_DIR
from psi_tool.run import ACTUAL, psi_layout
from psi_tool.sources import load_config, load_source

OUT = ROOT / "前期资料" / "分析报告" / "Auchan仓库库存补算表.xlsx"
# 这两周 PSI 用了上一周的 KTR 文件（当周: 实际用的周），补算时一起改成当周的，SO 和库存都改
KTR_WRONG = {37: 36, 41: 40}


def stock_total(inbox, cfg, name, week, column=None):
    c = dict(cfg[name])
    if column:
        c["stock"] = column
    rows = load_source(inbox, name, c, week).rows
    return None if rows is None else rows


def product_rows(layout):
    return [(c, n, act) for c, n, _, act in layout[1:] if act]  # 不含 Others：Auchan 仓库都落在产品行上


def ktr_delta(inbox, cfg, mapper, rows_by_code, week):
    """当周 KTR − 实际用的那周 KTR，按 PSI 行汇总。返回 {代码: (SO 差, 库存差)}。"""
    out = {}
    for sign, w in ((1, week), (-1, KTR_WRONG[week])):
        rows = stock_total(inbox, cfg, "KTR", w).copy()
        rows["code"] = [mapper.lookup("EAN", k)[0] for k in rows["key"]]
        rows["psi"] = [c if c in rows_by_code else "Others" for c in rows["code"]]
        for code, g in rows.groupby("psi"):
            so, st = out.get(code, (0, 0))
            out[code] = (so + sign * int(g["so"].sum()), st + sign * int(g["stock"].sum()))
    return {c: v for c, v in out.items() if v != (0, 0)}


def build_blocks(sheet, layout, corrected, weeks):
    """粘贴块：行顺序 = PSI 筛选 Channel Inventory Actual 后 Others 下面的顺序。
    某周的产品行里有公式时（历史上 P854+64 那行），整块只到公式行上面，防止把公式盖掉。"""
    rows = product_rows(layout)
    whole, partial = [], []
    cut = len(rows)
    for w in weeks:
        hits = [i for i, (_, _, r) in enumerate(rows) if is_formula(sheet.get(r, sheet.week_col(w)))]
        if hits:
            partial.append(w)
            cut = min(cut, hits[0])
        else:
            whole.append(w)

    def block(ws, n):
        out = pd.DataFrame({"代码": [c for c, _, _ in rows[:n]], "产品": [x for _, x, _ in rows[:n]],
                            "PSI 库存行": [r for _, _, r in rows[:n]]})
        for w in ws:
            col = sheet.week_col(w)
            vals = []
            for c, _, r in rows[:n]:
                now = sheet.get(r, col)
                add = corrected.get((w, c), 0)
                vals.append(now + add if is_number(now) else (add or None))
            out[f"W{w}"] = pd.array([None if v is None else int(round(v)) for v in vals], dtype="Int64")
        return out

    return block(whole, len(rows)), block(partial, cut), whole, partial, cut, rows


def verify(sheet):
    """贴完重新导出 PSI 后运行：逐格对照粘贴块。"""
    blocks = pd.read_excel(OUT, sheet_name=None)
    so_rows = "SO 修正"
    done, todo, other = 0, [], []
    cells = blocks["逐格修改"]
    original = {c: n for c, n in zip(cells["PSI 单元格"], cells["PSI 现值"])}
    for name, df in blocks.items():
        if not name.startswith(("粘贴块", so_rows)):
            continue
        row_col = "PSI SO 行" if name.startswith(so_rows) else "PSI 库存行"
        for col in [c for c in df.columns if str(c).startswith("W")]:
            pc = sheet.week_col(int(col[1:]))
            for _, r in df.iterrows():
                want, row = r[col], int(r[row_col])
                got = sheet.get(row, pc)
                if pd.isna(want) or not is_number(got):
                    continue
                addr = f"{get_column_letter(pc + 1)}{row}"
                if int(round(got)) == int(want):
                    done += 1
                elif addr in original and int(round(got)) == int(original[addr]):
                    todo.append(addr)
                else:
                    other.append(f"{addr}（应为 {int(want)}，现为 {int(round(got))}）")
    for _, c in cells[cells["代码"] == "Others"].iterrows():  # Others 行不在粘贴块里，手工改的
        row, col = coordinate_to_tuple(c["PSI 单元格"])
        got = sheet.get(row, col - 1)
        if is_number(got) and int(round(got)) == int(c["修正后"]):
            done += 1
        elif is_number(got) and int(round(got)) == int(c["PSI 现值"]):
            todo.append(c["PSI 单元格"])
        else:
            other.append(f"{c['PSI 单元格']}（应为 {int(c['修正后'])}，现为 {got}）")
    print(f"已正确 {done} 格；还没改 {len(todo)} 格；数对不上 {len(other)} 格")
    for name, xs in (("还没改", todo), ("数对不上", other)):
        if xs:
            print(f"{name}：" + "、".join(xs[:40]) + ("…" if len(xs) > 40 else ""))
    return done, todo, other


def main():
    ap = argparse.ArgumentParser(description="Auchan 仓库库存历史补算")
    ap.add_argument("--from-week", type=int, default=10)
    ap.add_argument("--to-week", type=int, default=43)
    ap.add_argument("--inbox", default=str(ROOT / "数据" / "历史周报"))
    ap.add_argument("--force", action="store_true", help="覆盖已有的补算表（PSI 还没贴时才能用）")
    ap.add_argument("--verify", action="store_true", help="贴完并重新导出 PSI 后，逐格核对")
    a = ap.parse_args()

    cfg = load_config(CONFIG)
    sheet = load(PSI_FILE)["OM Others"]
    if a.verify:
        verify(sheet)
        return
    if OUT.exists() and not a.force:
        raise SystemExit(f"{OUT.name} 已存在。如果 PSI 已经贴过，再算会把仓库加两遍；确认还没贴才加 --force。贴完核对请用 --verify")
    mapper = Mapper(latest_ttl(TTL_DIR), MAPPING)
    layout = psi_layout(sheet)
    rows_by_code = {c: (n, act) for c, n, _, act in layout}
    so_row = {c: r for c, _, r, _ in layout}
    summary, cells, corrected, weeks, so_fix = [], [], {}, [], {}
    for week in range(a.from_week, a.to_week + 1):
        col = sheet.week_col(week)
        store = stock_total(a.inbox, cfg, "AUCHAN", week, "Zapas Bieżący Ilość")
        wh = stock_total(a.inbox, cfg, "AUCHAN", week, "Zapas Magazynowy Ilość")
        psi_total = sum(v for _, _, _, r in layout if r for v in [sheet.get(r, col)] if is_number(v))
        if wh is None:
            summary.append({"周": f"W{week}", "核验": "❌ 没有 Auchan 文件，无法补算", "PSI 现合计": int(psi_total)})
            continue
        weeks.append(week)
        others = [stock_total(a.inbox, cfg, "KTR", KTR_WRONG.get(week, week)), stock_total(a.inbox, cfg, "MAX", week)]
        expected = None if any(x is None for x in others) else int(store["stock"].sum() + sum(x["stock"].sum() for x in others))
        wh = wh[wh["stock"] != 0].copy()
        check = ("⚠️ 缺 KTR 或 MAX 文件，未核验" if expected is None else
                 "✅ PSI 只含门店" if expected == int(psi_total) else
                 "❗ PSI 可能已含仓库，不要再加" if int(psi_total) - expected == int(wh["stock"].sum()) else
                 f"⚠️ PSI 和门店口径重算差 {int(psi_total) - expected}（不等于仓库量，是别的差异，如 KTR 用错周、未填的格子）")
        wh["code"] = [mapper.lookup("EAN", k)[0] for k in wh["key"]]
        wh["psi"] = [c if c in rows_by_code else "Others" for c in wh["code"]]
        for code, qty in wh.groupby("psi")["stock"].sum().items():
            name, r = rows_by_code[code]
            now = sheet.get(r, col)
            corrected[(week, code)] = int(qty)
            cells.append({"周": f"W{week}", "PSI 单元格": f"{get_column_letter(col + 1)}{r}", "代码": code, "产品": name,
                          "PSI 现值": now if is_number(now) else "（公式或空）", "加 Auchan 仓库": int(qty),
                          "修正后": int(now + qty) if is_number(now) else int(qty)})
        ktr_stock = 0
        if week in KTR_WRONG:
            for code, (d_so, d_st) in ktr_delta(a.inbox, cfg, mapper, rows_by_code, week).items():
                ktr_stock += d_st
                name, r = rows_by_code[code]
                if d_st:
                    now = sheet.get(r, col)
                    base = now if is_number(now) else 0
                    done = corrected.get((week, code), 0)
                    corrected[(week, code)] = done + d_st
                    cells.append({"周": f"W{week}", "PSI 单元格": f"{get_column_letter(col + 1)}{r}", "代码": code, "产品": name,
                                  "PSI 现值": base, "加 Auchan 仓库": done, "KTR 换当周": d_st, "修正后": int(base + done + d_st)})
                if d_so:
                    r2 = so_row[code]
                    now = sheet.get(r2, col)
                    so_fix[(week, code)] = d_so
                    cells.append({"周": f"W{week}", "PSI 单元格": f"{get_column_letter(col + 1)}{r2}", "代码": code, "产品": name,
                                  "PSI 现值": now if is_number(now) else 0, "SO：KTR 换当周": d_so,
                                  "修正后": int((now if is_number(now) else 0) + d_so)})
        summary.append({"周": f"W{week}", "Auchan 门店": int(store["stock"].sum()), "Auchan 仓库": int(wh["stock"].sum()),
                        "KTR 库存修正": ktr_stock, "PSI 现合计": int(psi_total),
                        "PSI 修正后合计": int(psi_total + wh["stock"].sum() + ktr_stock), "核验": check})
    whole, partial, wk_whole, wk_partial, cut, rows = build_blocks(sheet, layout, corrected, weeks)
    # 公式行及以下的格子如果也要改，块覆盖不到，要手工
    manual = [c for c in cells if int(c["周"][1:]) in wk_partial and
              [x for x, _, _ in rows].index(c["代码"]) >= cut] if cut < len(rows) else []
    # 同一格既加仓库又改 KTR 时，前面会有两条：只留最后一条（修正后是累计值）
    seen = {}
    for c in cells:
        seen[c["PSI 单元格"]] = c
    cells = list(seen.values())
    others_cells = [c for c in cells if c["代码"] == "Others"]
    if others_cells:
        print("⚠️ Others 行不在粘贴块里，要手工改：" + "、".join(f"{c['PSI 单元格']} → {c['修正后']}" for c in others_cells))
    so_rows = [(c, n, so_row[c]) for c, n, _ in rows]
    so_block = pd.DataFrame({"代码": [c for c, _, _ in so_rows], "产品": [n for _, n, _ in so_rows],
                             "PSI SO 行": [r for _, _, r in so_rows]})
    for w in sorted(KTR_WRONG):
        col = sheet.week_col(w)
        vals = []
        for c, _, r in so_rows:
            now = sheet.get(r, col)
            add = so_fix.get((w, c), 0)
            vals.append(now + add if is_number(now) else (add or None))
        so_block[f"W{w}"] = pd.array([None if v is None else int(round(v)) for v in vals], dtype="Int64")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        pd.DataFrame(summary).to_excel(xw, sheet_name="按周总览", index=False)
        whole.to_excel(xw, sheet_name=f"粘贴块1 W{wk_whole[0]}-W{wk_whole[-1]}"[:31], index=False)
        if wk_partial:
            partial.to_excel(xw, sheet_name=f"粘贴块2 W{wk_partial[0]}-W{wk_partial[-1]}"[:31], index=False)
        so_block.to_excel(xw, sheet_name="SO 修正 W37 W41", index=False)
        pd.DataFrame(cells).to_excel(xw, sheet_name="逐格修改", index=False)
    print(pd.DataFrame(summary).to_string(index=False))
    print(f"\n整列块周：{wk_whole}\n到公式行上面的周：{wk_partial}（共 {cut}/{len(rows)} 行）；块外要手工的格子 {len(manual)}")
    print(f"{len(cells)} 格：{OUT}")


if __name__ == "__main__":
    main()
