"""AB POCO 历史周补录：W45 起 AB 的 POCO 表才进每周流程，之前几周的 POCO 补成 PSI DB 页要追加的行。

用法：.venv/bin/python -m tools.poco_backfill 文件夹 [--weeks 42 43 44] [--out 前期资料/分析报告/AB_POCO补录.xlsx]
文件夹里放各周的 xiaomi_poco_sell_through_* / xiaomi_poco_stock_*（或 AB 的周报邮件 .msg），工具按文件名里的日期认周。
同一个文件里如果有人工加的透视页（Sheet1），会拿它的 Grand Total 和明细合计对一下。
"""
import argparse
import shutil
import tempfile
from pathlib import Path

import pandas as pd

from psi_tool.db import build_db, find_files
from psi_tool.paths import PSI_FILE, ROOT

COLS_ST = {"A": "A 周", "B": "B DB", "C": "C 产品名", "I": "I KA Name", "J": "J ST qty", "K": "K AB reference",
           "L": "L IM&ALSO reference", "P": "P ST time"}
COLS_STOCK = {"A": "A 周", "B": "B DB", "C": "C 产品名", "I": "I STOCK", "J": "J reference"}


def pivot_total(path):
    try:
        x = pd.read_excel(path, sheet_name="Sheet1", header=None, dtype=object)
    except ValueError:
        return None
    hit = x[x.iloc[:, 0].astype(str).str.strip() == "Grand Total"]
    return None if hit.empty else pd.to_numeric(hit.iloc[0, -1], errors="coerce")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--weeks", type=int, nargs="+", default=[42, 43, 44])
    ap.add_argument("--out", default=str(ROOT / "前期资料/分析报告/AB_POCO补录_W42-W44.xlsx"))
    a = ap.parse_args()
    sts, stocks, check = [], [], []
    for w in a.weeks:
        with tempfile.TemporaryDirectory() as tmp:
            inbox = Path(tmp) / f"W{w}"
            (inbox / "DB/AB").mkdir(parents=True)
            for key in ("AB_POCO_ST", "AB_POCO_STOCK"):
                for p in find_files(a.folder, key, w):
                    shutil.copy(p, inbox / "DB/AB" / p.name)
            st, stock, notes = build_db(w, str(inbox), PSI_FILE)
            st_files = find_files(inbox, "AB_POCO_ST", w)
            piv = pivot_total(st_files[0]) if st_files else None
        st = st[st["B"] == "AB"] if len(st) else st
        stock = stock[stock["B"] == "AB"] if len(stock) else stock
        no_code = int(st["代码（预查）"].isna().sum() + stock["代码（预查）"].isna().sum()) if len(st) or len(stock) else 0
        check.append({"周": f"W{w}", "ST 行数": len(st), "ST 合计": int(st["J"].sum()) if len(st) else 0,
                      "你的透视 Grand Total": piv, "一致": "" if piv is None else ("✅" if len(st) and piv == st["J"].sum() else "❌"),
                      "库存行数": len(stock), "库存合计": int(stock["I"].sum()) if len(stock) else 0,
                      "查不到代码的行": no_code})
        sts.append(st)
        stocks.append(stock)
    st, stock = pd.concat(sts), pd.concat(stocks)
    with pd.ExcelWriter(a.out) as xw:
        pd.DataFrame(check).to_excel(xw, sheet_name="核对", index=False)
        st[list(COLS_ST) + ["代码（预查）"]].rename(columns=COLS_ST).to_excel(xw, sheet_name="2026 DB ST 补录", index=False)
        stock[list(COLS_STOCK) + ["代码（预查）"]].rename(columns=COLS_STOCK).to_excel(xw, sheet_name="2026 DB STOCK 补录", index=False)
        pd.DataFrame({"说明": [
            "AB 的 POCO 表从 W45 起进每周流程，这里是 W42–W44 要补进 PSI 的行。",
            "粘贴方式和每周 DB 一样：追加到 2026 DB ST / 2026 DB STOCK 最后，A–C、I–L、P 列直接贴，D–H、M–O 从上一行拉公式。",
            "B 列按经销商写 AB（和 smart 表一样）。补进去后，OM 各页 W42–W44 的 ST 公式会自动带上这些 POCO 的量。",
            "代码（预查）列是工具按 工具表1 先查的结果，空着说明 工具表1 里没有这个物料码。",
        ]}).to_excel(xw, sheet_name="说明", index=False)
    print(pd.DataFrame(check).to_string(index=False))
    print(a.out)


if __name__ == "__main__":
    main()
