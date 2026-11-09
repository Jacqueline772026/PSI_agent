"""每周跑一次：原始周报 → 待粘贴表 + 核对报告。

用法：
  .venv/bin/python -m psi_tool.run --week 44 --inbox "W44测试/原始文件" --out "W44测试/工具结果"
  加 --compare 会把结果和 PSI 导出里该周已经填好的数逐格比对（用来验证工具，或粘贴后复核）。
"""
import argparse
import collections
import datetime as dt
from pathlib import Path

import pandas as pd

from checks.structure_check import COL_C, COL_E, COL_G, COL_LABEL, OTHERS, is_number, load as load_psi
from checks.identity import OM as OM_SHEETS, OP as OP_SHEETS, DbSt, check as identity_check, report_lines as identity_lines
from checks.tmo_guard import check_tmo, report_lines
from psi_tool.db import build_db, compare_db
from psi_tool.mapping import Mapper, latest_ttl, resolve
from psi_tool.paths import CONFIG, MAPPING, OUTPUT, PSI_FILE, TTL_DIR
from psi_tool.sources import load_config, load_source, week_range

ROOT = Path(__file__).resolve().parent.parent
ACTUAL = "Channel Inventory Actual"


def psi_layout(sheet):
    """PSI 某张 sheet 的粘贴顺序：[(代码, 产品名, SO 行号, Actual 行号)]，第一项是 Others。"""
    def actual_row(r, code):
        for x in range(r + 1, r + 8):
            if sheet.get(x, COL_LABEL) == ACTUAL and (code is None or sheet.code(x) == code):
                return x
        return None
    others_so = next(r for r in sheet.others_rows if sheet.get(r, COL_LABEL) == "SO")
    rows = [("Others", "Others", others_so, actual_row(others_so, None))]
    for r in sheet.product_rows("SO"):
        code = sheet.code(r)
        rows.append((code, sheet.get(r, COL_E), r, actual_row(r, code)))
    return rows


def map_rows(df, mapper, sheet_codes, projects):
    out = []
    for _, r in df.iterrows():
        code, cat, via = mapper.lookup(r["key_type"], r["key"])
        if cat == "exclude":
            status = "excluded"
        elif code is None:
            status = "excluded" if r.get("drop_unmatched") else "unmatched"
        else:
            code = resolve(code, r["name"], sheet_codes, projects)
            status = "psi" if code in sheet_codes else "others"
        out.append((code, via, status))
    df = df.copy()
    df[["psi_code", "via", "status"]] = pd.DataFrame(out, index=df.index)
    return df


def build(week, inbox, psi_path, config_path, mapping_path, ttl_root):
    cfg = load_config(config_path)
    psi = load_psi(psi_path)
    mapper = Mapper(latest_ttl(ttl_root), mapping_path)
    sources = {name: load_source(inbox, name, c, week) for name, c in cfg.items()}

    sheets = {}
    for name, res in sources.items():
        sheets.setdefault(res.psi_sheet, []).append(res)
    results = {}
    for sheet_name, parts in sheets.items():
        sheet = psi[sheet_name]
        layout = psi_layout(sheet)
        sheet_codes = {c for c, *_ in layout[1:]}
        projects = collections.defaultdict(list)
        for r in sheet.product_rows("SO"):
            projects[str(sheet.get(r, COL_G)).strip()].append(sheet.code(r))
        frames = [p.rows for p in parts if p.rows is not None and len(p.rows)]
        rows = map_rows(pd.concat(frames, ignore_index=True), mapper, sheet_codes, projects) if frames else None
        has_others = all(cfg[p.name].get("others", True) for p in parts)
        fields = {f for p in parts for f in ("so", "stock") if cfg[p.name].get(f)}  # 这张 sheet 的文件里提供了哪些指标
        results[sheet_name] = dict(sheet=sheet, layout=layout, parts=parts, rows=rows, has_others=has_others,
                                   fields=fields)
    return results, mapper


def summarize(res):
    rows = res["rows"]
    if rows is None:
        return None
    by_code = rows[rows["status"] == "psi"].groupby("psi_code")[["so", "stock"]].sum()
    to_others = rows[rows["status"].isin(["others", "unmatched"])] if res["has_others"] else rows.iloc[0:0]
    table = []
    for code, name, so_row, act_row in res["layout"]:
        if code == "Others":
            if not res["has_others"]:
                continue
            so, stock = to_others["so"].sum(), to_others["stock"].sum()
        else:
            so, stock = (by_code.loc[code, "so"], by_code.loc[code, "stock"]) if code in by_code.index else (0, 0)
        table.append(dict(code=code, name=name, so_row=so_row, actual_row=act_row, so=int(so), stock=int(stock)))
    return pd.DataFrame(table)


def compare(res, table, week):
    """和 PSI 导出里该周的手填值比对；PSI 里是公式（还没填）的格子跳过。"""
    sheet = res["sheet"]
    col = sheet.week_col(week)
    diffs, checked = [], 0
    for _, t in table.iterrows():
        for field, r in (("so", t["so_row"]), ("stock", t["actual_row"])):
            if r is None or field not in res["fields"]:
                continue
            v = sheet.get(r, col)
            if not is_number(v):
                continue
            checked += 1
            if int(round(v)) != t[field]:
                diffs.append(dict(code=t["code"], field="SO" if field == "so" else "库存", psi=int(round(v)), ours=t[field]))
    return checked, diffs


CHANGE_ABS, CHANGE_PCT = 50, 0.5  # 和上周比：变化超过 50 台且超过 50% 才提示
APPEAR = 10  # 从 0 变成至少 10 台（或反过来）才提示，过滤 1、2 台的零星波动


def hand_check(cur, prev):
    """给人工抽查用的提示：新出现 / 归零 / 变化大。prev 是 PSI 里上周已填的值（没填为 None）。"""
    if prev is None:
        return ""
    if prev == 0 and cur >= APPEAR:
        return "上周 0，本周有数"
    if prev >= APPEAR and cur == 0:
        return "上周有数，本周 0"
    if abs(cur - prev) > max(CHANGE_ABS, CHANGE_PCT * abs(prev)):
        return f"{cur - prev:+d}"
    return ""


def write_paste(path, results, tables, week):
    """每张 PSI sheet 一页，行顺序 = PSI 里筛选 SO（或 Channel Inventory Actual）后的顺序。
    OM 第一行是 Others；运营商没有 Others。两类都可以整列粘贴：
    运营商周报里没出现的型号填 0（和 ST 的规则一致：BI 查不到就改成 0）。
    右边几列是上周 PSI 的数和提示，方便一边粘贴一边人工抽查；有提示的行标黄。"""
    from openpyxl.styles import PatternFill
    yellow = PatternFill("solid", fgColor="FFF2CC")
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        if all(t is None for t in tables.values()):  # 一个文件都还没收到
            pd.DataFrame({"": ["还没有找到任何当周文件，见核对报告"]}).to_excel(xw, sheet_name="说明", index=False)
        for sheet_name, table in tables.items():
            if table is None:
                continue
            sheet = results[sheet_name]["sheet"]
            prev_col = sheet.week_col(week - 1)
            out = []
            for _, t in table.iterrows():
                prev = {f: sheet.get(r, prev_col) if r else None for f, r in (("so", t["so_row"]), ("stock", t["actual_row"]))}
                prev = {f: int(round(v)) if is_number(v) else None for f, v in prev.items()}
                out.append({
                    "代码": t["code"], "产品": t["name"],
                    f"W{week} SO": t["so"], f"W{week} 库存": t["stock"],
                    f"W{week - 1} SO（PSI）": prev["so"], f"W{week - 1} 库存（PSI）": prev["stock"],
                    "SO 提示": hand_check(t["so"], prev["so"]),
                    "库存提示": hand_check(t["stock"], prev["stock"]),
                    "PSI SO 行": t["so_row"], "PSI 库存行": t["actual_row"],
                })
            out = pd.DataFrame(out)
            for c in ("PSI SO 行", "PSI 库存行", f"W{week - 1} SO（PSI）", f"W{week - 1} 库存（PSI）"):
                out[c] = out[c].astype("Int64")
            total = pd.DataFrame([{"代码": "合计", f"W{week} SO": out[f"W{week} SO"].sum(),
                                   f"W{week} 库存": out[f"W{week} 库存"].sum()}])
            pd.concat([out, total], ignore_index=True).to_excel(xw, sheet_name=sheet_name[:31], index=False)
            ws = xw.sheets[sheet_name[:31]]
            for i, row in enumerate(out.itertuples(index=False), start=2):
                if row[6] or row[7]:
                    for cell in ws[i]:
                        cell.fill = yellow
            ws.column_dimensions["B"].width = 28
            ws.freeze_panes = "C2"


DB_ORDER = {"AB": 0, "ALSO": 1, "IM": 2}


def write_db_paste(path, st, stock):
    """追加到 PSI 的 DB 行。只放要粘贴的列，列名带上 PSI 的列字母；D–H、M–O 是公式，从上一周拉下来。"""
    from openpyxl import load_workbook
    sheets = {
        "2026 DB ST": (st, {"A": "A 周", "B": "B DB", "C": "C 产品名", "I": "I KA Name", "J": "J ST qty",
                            "K": "K AB reference", "L": "L IM&ALSO reference", "P": "P ST time"}),
        "2026 DB STOCK": (stock, {"A": "A 周", "B": "B DB", "C": "C 产品名", "I": "I STOCK", "J": "J reference"}),
    }
    with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
        for name, (df, cols) in sheets.items():
            if not len(df):
                continue
            df = df.sort_values("B", key=lambda b: b.map(DB_ORDER), kind="stable")
            out = df[list(cols) + ["代码（预查）"]].rename(columns=cols)
            out.to_excel(xw, sheet_name=name, index=False)
    wb = load_workbook(path)
    for name in reversed(list(sheets)):
        if name in wb.sheetnames:  # DB 两页放在最前面（周一先做 DB）
            wb.move_sheet(name, offset=-wb.sheetnames.index(name))
    wb.save(path)


def db_report(week, st, stock, notes, cmp):
    L = ["## DB（AB / ALSO / IM）", "", *notes, ""]
    rows = []
    for name, df, q in (("ST", st, "J"), ("库存", stock, "I")):
        if len(df):
            for db, n in df.groupby("B")[q].agg(["count", "sum"]).iterrows():
                rows.append([name, db, n["count"], n["sum"]])
    L += [md_table(["", "DB", "行数", "合计"], rows), ""]
    for name, (n, diff, totals) in (cmp or {}).items():
        L.append(f"**和 PSI「{name}」已贴的 W{week} 比对**：{n} 个 DB+识别码组合，"
                 + ("✅ 全部一致" if not len(diff) else f"⚠️ {len(diff)} 处不同"))
        if len(diff):
            L += ["", md_table(list(diff.columns), diff.values.tolist())]
        L.append("")
    return L


def md_table(headers, rows):
    lines = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
    return "\n".join(lines + ["| " + " | ".join(str(x) for x in r) + " |" for r in rows])


def write_report(path, week, results, tables, comparisons, mapper):
    start, end = week_range(week)
    L = [f"# W{week} 周报处理报告", "",
         f"- 周：2026-W{week}（{start} ~ {end}）",
         f"- 运行时间：{dt.datetime.now():%Y-%m-%d %H:%M}",
         f"- TTL ID 表：`{mapper.ttl_path.name}`；人工映射表：`{mapper.mapping_path.name}`（{len(mapper.manual)} 条）", "",
         "## 总览", ""]
    rows = []
    for sheet_name, res in results.items():
        t = tables[sheet_name]
        got = [p for p in res["parts"] if p.files]
        miss = [p.name for p in res["parts"] if not p.files]
        r = res["rows"]
        unmatched = 0 if r is None else int((r["status"] == "unmatched").sum())
        cmp = comparisons.get(sheet_name)
        cmp_txt = "—" if not cmp else (f"✅ {cmp[0]} 格一致" if not cmp[1] else f"⚠️ {len(cmp[1])}/{cmp[0]} 格不同")
        status = "❌ 未收到" if not got else ("⚠️ 缺 " + "、".join(miss) if miss else "✅")
        rows.append([sheet_name, status, "—" if t is None else t["so"].sum(), "—" if t is None else t["stock"].sum(),
                     unmatched, cmp_txt])
    L += [md_table(["PSI sheet", "文件", "SO 合计", "库存合计", "未匹配行", "和 PSI 比对"], rows), ""]

    for sheet_name, res in results.items():
        L += [f"## {sheet_name}", ""]
        for p in res["parts"]:
            files = "、".join(f"`{Path(f).name}`" for f in p.files) or "**没找到**"
            per = f"，时间段 {p.period[0]} ~ {p.period[1]}" if p.period and None not in p.period else ""
            L.append(f"- **{p.name}**：{files}{per}")
            for w in p.warnings:
                L.append(f"  - ⚠️ {w}")
        r, t = res["rows"], tables[sheet_name]
        if r is None:
            L += [""]
            continue
        agg = r.groupby("status")[["so", "stock"]].sum()
        g = lambda s, c: int(agg.loc[s, c]) if s in agg.index else 0
        L += ["", md_table(["", "SO", "库存"], [
            ["筛选后原始合计", int(r["so"].sum()), int(r["stock"].sum())],
            ["→ 匹配到 PSI 行", g("psi", "so"), g("psi", "stock")],
            ["→ 有代码但 PSI 没这一行（计入 Others）" if res["has_others"] else "→ 有代码但 PSI 没这一行（不填）",
             g("others", "so"), g("others", "stock")],
            ["→ 未匹配（计入 Others，需确认）" if res["has_others"] else "→ 未匹配（需确认）",
             g("unmatched", "so"), g("unmatched", "stock")],
            ["→ 映射表标记排除", g("excluded", "so"), g("excluded", "stock")],
        ]), ""]
        check = t["so"].sum(), t["stock"].sum()
        expect = (g("psi", "so") + (g("others", "so") + g("unmatched", "so") if res["has_others"] else 0),
                  g("psi", "stock") + (g("others", "stock") + g("unmatched", "stock") if res["has_others"] else 0))
        L.append(f"待粘贴表合计 SO {check[0]} / 库存 {check[1]}，" + ("✅ 和上表一致" if check == expect else f"❌ 应为 {expect}"))
        um = r[r["status"] == "unmatched"].groupby(["key", "name"], as_index=False)[["so", "stock"]].sum()
        if len(um):
            L += ["", "**未匹配**（是手机/平板就在 mapping.xlsx 补一行；不是就补一行 category=exclude）", "",
                  md_table(["识别码", "产品名", "SO", "库存"], [[k, n[:50], int(a), int(b)] for k, n, a, b in um.values])]
        oth = r[r["status"] == "others"].groupby("psi_code", as_index=False)[["so", "stock"]].sum()
        oth = oth[(oth["so"] != 0) | (oth["stock"] != 0)]
        if len(oth):
            L += ["", "**有代码但 PSI 这张 sheet 没有对应行**（老品正常；新品要考虑加行）", "",
                  md_table(["代码", "SO", "库存"], [[c, int(a), int(b)] for c, a, b in oth.values])]
        if sheet_name == "TMO":
            excluded = [p.name_excluded for p in res["parts"] if p.name_excluded is not None]
            L += report_lines(check_tmo(r, pd.concat(excluded) if excluded else None), md_table)
        cmp = comparisons.get(sheet_name)
        if cmp and cmp[1]:
            L += ["", f"**和 PSI 已填的 W{week} 不一致的格子**", "",
                  md_table(["代码", "指标", "PSI", "工具"], [[d["code"], d["field"], d["psi"], d["ours"]] for d in cmp[1]])]
        L.append("")
    Path(path).write_text("\n".join(L), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="PSI 周报处理")
    ap.add_argument("--week", type=int, required=True)
    ap.add_argument("--inbox", required=True, help="当周原始文件所在文件夹（会递归查找）")
    ap.add_argument("--psi", default=str(PSI_FILE))
    ap.add_argument("--config", default=str(CONFIG))
    ap.add_argument("--mapping", default=str(MAPPING))
    ap.add_argument("--ttl-dir", default=str(TTL_DIR))
    ap.add_argument("--out", default=str(OUTPUT))
    ap.add_argument("--compare", action="store_true", help="和 PSI 导出里该周已填的数比对")
    a = ap.parse_args()

    results, mapper = build(a.week, a.inbox, a.psi, a.config, a.mapping, a.ttl_dir)
    tables = {k: summarize(v) for k, v in results.items()}
    comparisons = {k: compare(results[k], t, a.week) for k, t in tables.items() if a.compare and t is not None}
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    paste, report = out / f"W{a.week}_paste.xlsx", out / f"W{a.week}_report.md"
    write_paste(paste, results, tables, a.week)
    write_report(report, a.week, results, tables, comparisons, mapper)
    st, stock, notes = build_db(a.week, a.inbox, a.psi)
    db_cmp = compare_db(a.psi, a.week, st, stock) if a.compare else None
    write_db_paste(paste, st, stock)
    # 要找客户 / 同事确认的问题：联系人 + 英文邮件草稿
    missing = [p.name for res in results.values() for p in res["parts"] if not p.files]
    from psi_tool.db import im_crosscheck, mark_im_files
    from psi_tool.inquiry import write as write_inquiry
    diffs = im_crosscheck(a.inbox, a.week) or []
    marked = mark_im_files(a.inbox, a.week, diffs, out) if diffs else []
    write_inquiry(out / f"W{a.week}_待确认邮件.md", a.week, missing, diffs, marked)
    text = report.read_text(encoding="utf-8").split("\n")
    at = text.index("## 总览")  # DB 放在 OM/OP 前面
    # 库存恒等式：上周库存和本周 ST 来自 PSI 导出，本周 SO / 库存用工具结果（粘贴前就能看到）
    current = {k: {x.code: (x.so, x.stock) for x in t.itertuples()} for k, t in tables.items() if t is not None}
    db = DbSt(a.psi, {a.week})
    ident = identity_check({k: results[k]["sheet"] for k in current}, db, a.week, current,
                           sheets=[k for k in OM_SHEETS + OP_SHEETS if k in current])
    text += [""] + identity_lines(ident, a.week, db, md_table, "工具结果")
    report.write_text("\n".join(text[:at] + db_report(a.week, st, stock, notes, db_cmp) + text[at:]), encoding="utf-8")
    print(f"待粘贴表：{paste}\n核对报告：{report}")
    for k, (n, d, _) in (db_cmp or {}).items():
        print(f"  {k}: 比对 {n} 组，不一致 {len(d)}")
    for k, (n, d) in comparisons.items():
        print(f"  {k}: 比对 {n} 格，不一致 {len(d)}")


if __name__ == "__main__":
    main()
