"""手工版 PSI 和工具结果逐格对比，生成"修正清单"（带单元格位置和可能原因）。

  .venv/bin/python -m tools.fix_list --week 45 --inbox "W45/W45原始文件" --psi 手工版PSI.xlsx [--out W45/工具结果]

列出：
  - OM / 运营商各页：PSI 已填但和工具不一样的格子；PSI 空着、工具有数的格子
  - DB 两页：按 经销商 + 识别码 + 客户 对不上的组
可能原因（自动判断，只是提示）：
  - "和上周一样"：这一格可能没更新，沿用了上周的数
  - "等于 XX 这一行的工具值"：可能填串了行（如 P7E8+256 填成了 O7E8+256 的数）
  - "SO 和库存对调"：两格的数正好互换
周报没收齐的页（比如缺 X-KOM 文件）不比。
"""
import argparse
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from checks.structure_check import is_number, load
from psi_tool.db import build_db, compare_db
from psi_tool.paths import CONFIG, MAPPING, TTL_DIR
from psi_tool.run import build, summarize


def ka_fixes(week, inbox, psi_path):
    results, _ = build(week, inbox, psi_path, CONFIG, MAPPING, TTL_DIR)
    out, skipped = [], []
    for sheet, res in results.items():
        t = summarize(res)
        if t is None or any(not p.files for p in res["parts"]):
            skipped.append(sheet)
            continue
        s = res["sheet"]
        c, p = s.week_col(week), s.week_col(week - 1)
        col = get_column_letter(c + 1)  # get() 的列号比 Excel 少 1
        by_value = {}
        for x in t.itertuples():
            by_value.setdefault((x.so, x.stock), []).append(x.code)
        for x in t.itertuples():
            cells = {}
            for f, rw, lab in (("so", x.so_row, "SO"), ("stock", x.actual_row, "库存")):
                if rw is None or f not in res["fields"]:
                    continue
                v, prev = s.get(int(rw), c), s.get(int(rw), p)
                cur = int(round(v)) if is_number(v) else None
                cells[f] = (int(rw), lab, cur, getattr(x, f), prev)
            for f, (rw, lab, cur, new, prev) in cells.items():
                if cur == new or (cur is None and new == 0):
                    continue
                why = []
                if cur is None:
                    why.append("PSI 空着，周报有数")
                elif is_number(prev) and cur == int(round(prev)) and cur != new:
                    why.append("和上周一样：可能没更新，沿用了上周的数")
                other = cells.get("stock" if f == "so" else "so")
                if other and cur is not None and other[2] is not None and cur == other[3] and other[2] == new:
                    why.append("SO 和库存对调了")
                if "so" in cells and "stock" in cells:
                    pair = (cells["so"][2], cells["stock"][2])
                    hit = [k for k in by_value.get(pair, []) if k != x.code]
                    if hit:
                        why.append(f"PSI 这两格等于 {hit[0]} 这一行的工具值：可能填串了行")
                out.append(dict(sheet=sheet, cell=f"{col}{rw}", code=x.code, metric=lab,
                                cur="空" if cur is None else cur, new=new, why="；".join(why)))
    return out, skipped


def db_fixes(week, inbox, psi_path):
    st, stock, _ = build_db(week, inbox, psi_path)
    cmp = compare_db(psi_path, week, st, stock)
    out = []
    for name, (n, diff, _) in cmp.items():
        for _, d in diff.iterrows():
            key = " / ".join(str(d[k]) for k in d.index if k not in ("PSI", "工具"))
            out.append(dict(sheet=name, cell="（按组）", code=key, metric="ST" if name.endswith("ST") else "库存",
                            cur=int(d["PSI"]), new=int(d["工具"]), why="PSI 里没有这一组" if d["PSI"] == 0 else ""))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--week", type=int, required=True)
    ap.add_argument("--inbox", required=True)
    ap.add_argument("--psi", required=True, help="手工做好的这一周 PSI 导出")
    ap.add_argument("--out", default=".")
    a = ap.parse_args()
    ka, skipped = ka_fixes(a.week, a.inbox, a.psi)
    rows = ka + db_fixes(a.week, a.inbox, a.psi)
    out = Path(a.out) / f"W{a.week}_PSI核对修正清单.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "修正清单"
    ws.append(["PSI sheet", "单元格", "代码 / 组", "指标", "PSI 现在", "工具（原始周报）", "可能原因（自动判断）", "处理结论（人工填）"])
    for r in rows:
        ws.append([r["sheet"], r["cell"], r["code"], r["metric"], r["cur"], r["new"], r["why"], ""])
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFGH", (14, 10, 40, 8, 10, 14, 52, 30)):
        ws.column_dimensions[col].width = w
    info = wb.create_sheet("说明")
    for line in [f"W{a.week}：手工版 {Path(a.psi).name} 和原始周报（工具）逐格对比。",
                 f"没比对的页（周报没收齐）：{'、'.join(skipped) or '无'}。",
                 "单元格是 PSI 里的真实位置（W{} 在这一列）。'可能原因'是自动判断，请核实后在最后一列写处理结论。".format(a.week)]:
        info.append([line])
    wb.save(out)
    print(out)
    print(pd.DataFrame(rows).groupby(["sheet"]).size().to_string() if rows else "全部一致")


if __name__ == "__main__":
    main()
