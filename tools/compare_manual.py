"""你的 Excel 手工结果 vs 工具结果，逐格对数。

第 1 步：生成模板（工具跑完之后）
  .venv/bin/python -m tools.compare_manual --week 44 --template
  → W44测试/Excel手工/W44_手工结果.xlsx
    每张 PSI sheet 一页，代码已经按 PSI 筛选后的顺序排好（和你往 PSI 里粘贴的顺序一样），
    你把自己 SUMIF / 透视表的结果粘贴进"你的 SO""你的库存"两列即可；运营商只填有数的代码也行。
    "DB 合计"页填 AB / ALSO / IM 的 ST 合计和库存合计。

第 2 步：对数
  .venv/bin/python -m tools.compare_manual --week 44
  → W44测试/Excel手工/W44_对数结果.xlsx（不一样的行标红）和 W44_对数结果.md（总结）

空着没填的格子不比；填了 0 的会比。
"""
import argparse
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from psi_tool.paths import ROOT

DB_SHEETS = ("2026 DB ST", "2026 DB STOCK")
MINE_SO, MINE_STOCK = "你的 SO", "你的库存"


def paths(week, tool, manual):
    base = ROOT / f"W{week}测试" if (ROOT / f"W{week}测试").exists() else ROOT / f"W{week}"
    tool = Path(tool) if tool else base / "工具结果" / f"W{week}_paste.xlsx"
    manual = Path(manual) if manual else base / "Excel手工" / f"W{week}_手工结果.xlsx"
    return tool, manual


def read_tool(tool_path, week):
    sheets = pd.read_excel(tool_path, sheet_name=None)
    om_op = {}
    for name, df in sheets.items():
        if name in DB_SHEETS or name == "说明":
            continue
        df = df[df["代码"] != "合计"]
        om_op[name] = df[["代码", "产品", f"W{week} SO", f"W{week} 库存", "PSI SO 行", "PSI 库存行"]].rename(
            columns={f"W{week} SO": "工具 SO", f"W{week} 库存": "工具库存"})
    db = {}
    if "2026 DB ST" in sheets:
        db["ST"] = sheets["2026 DB ST"].groupby("B DB")["J ST qty"].sum()
    if "2026 DB STOCK" in sheets:
        db["库存"] = sheets["2026 DB STOCK"].groupby("B DB")["I STOCK"].sum()
    return om_op, db


def make_template(tool_path, manual_path, week):
    om_op, _ = read_tool(tool_path, week)
    manual_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(manual_path, engine="openpyxl") as xw:
        pd.DataFrame({"DB": ["AB", "ALSO", "IM"], "ST 合计": None, "库存合计": None}).to_excel(
            xw, sheet_name="DB 合计", index=False)
        for name, df in om_op.items():
            out = df[["代码", "产品", "PSI SO 行", "PSI 库存行"]].copy()
            out[MINE_SO], out[MINE_STOCK] = None, None
            out.to_excel(xw, sheet_name=name, index=False)
    wb = load_workbook(manual_path)
    fill = PatternFill("solid", fgColor="E2EFDA")
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.column_dimensions["B"].width = 28
        ws.freeze_panes = "A2"
        for col in range(1, ws.max_column + 1):  # 要你填的列标绿
            if ws.cell(1, col).value in (MINE_SO, MINE_STOCK, "ST 合计", "库存合计"):
                for row in range(1, ws.max_row + 1):
                    ws.cell(row, col).fill = fill
    wb.save(manual_path)
    print(f"模板：{manual_path}")


def num(x):
    try:
        return None if pd.isna(x) or str(x).strip() == "" else int(round(float(x)))
    except (TypeError, ValueError):
        return None


def compare(tool_path, manual_path, week):
    om_op, db = read_tool(tool_path, week)
    mine = pd.read_excel(manual_path, sheet_name=None)
    summary, detail = [], {}

    if "DB 合计" in mine:
        rows = []
        for _, r in mine["DB 合计"].iterrows():
            for field, col in (("ST", "ST 合计"), ("库存", "库存合计")):
                m = num(r.get(col))
                t = int(db.get(field, pd.Series(dtype=int)).get(r["DB"], 0))
                if m is not None:
                    rows.append({"DB": r["DB"], "指标": field, "工具": t, "你": m, "差": m - t})
        d = pd.DataFrame(rows)
        if len(d):
            detail["DB 合计"] = d
            summary.append(("DB 合计", len(d), int((d["差"] != 0).sum())))

    for name, t in om_op.items():
        if name not in mine:
            continue
        m = mine[name]
        if MINE_SO not in m.columns:
            continue
        # 按行号对齐（代码可能重复，比如 TMO 的 P15AE4+128 有两块）
        j = t.merge(m[["PSI SO 行", MINE_SO, MINE_STOCK]], on="PSI SO 行", how="left")
        rows, checked = [], 0
        for _, r in j.iterrows():
            for field, tcol, mcol in (("SO", "工具 SO", MINE_SO), ("库存", "工具库存", MINE_STOCK)):
                mv = num(r[mcol])
                if mv is None:
                    continue
                checked += 1
                tv = int(r[tcol])
                if mv != tv:
                    rows.append({"代码": r["代码"], "产品": r["产品"], "指标": field, "工具": tv, "你": mv,
                                 "差（你 − 工具）": mv - tv, "PSI 行": int(r["PSI SO 行"] if field == "SO" else r["PSI 库存行"])})
        if checked:
            summary.append((name, checked, len(rows)))
            detail[name] = pd.DataFrame(rows, columns=["代码", "产品", "指标", "工具", "你", "差（你 − 工具）", "PSI 行"])
    return summary, detail


def write_result(summary, detail, out_xlsx, out_md, week):
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xw:
        pd.DataFrame(summary, columns=["sheet", "比对格数", "不一样"]).to_excel(xw, sheet_name="总览", index=False)
        for name, d in detail.items():
            d.to_excel(xw, sheet_name=name[:31], index=False)
    wb = load_workbook(out_xlsx)
    red = PatternFill("solid", fgColor="F8CBAD")
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = Font(bold=True)
        if ws.title == "总览":
            for row in ws.iter_rows(min_row=2):
                if row[2].value:
                    for c in row:
                        c.fill = red
        elif ws.title == "DB 合计":
            for row in ws.iter_rows(min_row=2):
                if row[4].value:
                    for c in row:
                        c.fill = red
    wb.save(out_xlsx)

    L = [f"# W{week} 对数结果：你的 Excel vs 工具", "",
         "| sheet | 比对格数 | 不一样 |", "|---|---|---|"]
    L += [f"| {n} | {c} | {'✅ 0' if d == 0 else f'⚠️ {d}'} |" for n, c, d in summary]
    for name, d in detail.items():
        bad = d[d["差"] != 0] if name == "DB 合计" else d  # 其他 sheet 的明细本来就只有不一样的行
        if len(bad):
            L += ["", f"## {name}", "", "| " + " | ".join(bad.columns) + " |", "|" + "---|" * len(bad.columns)]
            L += ["| " + " | ".join(str(x) for x in r) + " |" for r in bad.values]
    L += ["", "差异怎么查：先看工具报告（W{周}_report.md）里这一家有没有提示（未匹配、冲正、排除的样机等）；"
          "再看你的 Excel 这个代码是哪几行加起来的。查不出来把这一节发给我。"]
    Path(out_md).write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L[:len(summary) + 4]))
    print(f"\n明细：{out_xlsx}")


def main():
    ap = argparse.ArgumentParser(description="Excel 手工结果和工具结果对数")
    ap.add_argument("--week", type=int, required=True)
    ap.add_argument("--tool", help="工具的 W{周}_paste.xlsx")
    ap.add_argument("--manual", help="你的 W{周}_手工结果.xlsx")
    ap.add_argument("--template", action="store_true", help="生成手工结果模板")
    a = ap.parse_args()
    tool, manual = paths(a.week, a.tool, a.manual)
    if a.template:
        make_template(tool, manual, a.week)
        return
    summary, detail = compare(tool, manual, a.week)
    write_result(summary, detail, manual.with_name(f"W{a.week}_对数结果.xlsx"),
                 manual.with_name(f"W{a.week}_对数结果.md"), a.week)


if __name__ == "__main__":
    main()
