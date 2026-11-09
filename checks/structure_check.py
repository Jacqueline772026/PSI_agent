"""PSI 表结构校验。

读飞书导出的 PSI xlsx（只读公式和手填值，不依赖公式计算结果），检查：
  1. 汇总漏行：KA sheet 里的代码在 OM / OP / ALL 里找不到
  2. 小计区间：SUMIFS 的区间漏掉 Others 或产品行、或者重复计数
  3. 重复：同一张 sheet 里同一个 代码+指标 出现多次
  4. 代码书写：G/H 列前后空格；F 列手写且和 G&H 不一致
  5. ST 公式末尾手动加的 BI 数（+N，MEX、Max）：复制到下一周前要删
  6. 库存恒等式：本周 Actual 与 上周 Actual + ST − SO 偏差
  7. 每张 sheet 已填到第几周

用法：
  .venv/bin/python -m checks.structure_check [数据/Poland PSI DB+Channel.xlsx] [--out output/]
"""
import argparse
import collections
import datetime as dt
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

OM_SHEETS = ["EURO", "MEX", "X-KOM", "MSH", "Michain", "OM Others"]
OP_SHEETS = ["Orange", "Play", "TMO", "Plus"]
AGG_SHEETS = {"OM": OM_SHEETS, "OP": OP_SHEETS, "📌ALL": ["OM", "OP"]}
ALL_SHEETS = OM_SHEETS + OP_SHEETS + list(AGG_SHEETS)

COL_LABEL, COL_A, COL_B, COL_C, COL_E, COL_F, COL_G, COL_H = 8, 0, 1, 2, 4, 5, 6, 7
OTHERS = ("Others", "OTHERS")
VOLUME_LABELS = ("ST", "SO", "Channel Inventory Actual")

IDENTITY_MIN_ABS = 50  # 库存恒等式偏差的绝对阈值（台）
IDENTITY_MIN_PCT = 0.2  # 以及相对上周库存的比例阈值
IDENTITY_TOP_N = 10

SUMIFS_RANGE = re.compile(r"SUMIFS\(\$?[A-Z]+\$(\d+):\$?[A-Z]+\$(\d+),\$([A-Z])\$")
TRAILING_PLUS = re.compile(r"\)\s*\+\s*\d+(\.\d+)?\s*$")


def cell_text(x):
    return getattr(x, "text", x)  # ArrayFormula -> 公式文本


def is_formula(x):
    return isinstance(x, str) and x.startswith("=")


def is_number(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


class Sheet:
    def __init__(self, name, rows):
        self.name = name
        self.rows = rows
        self._locate_weeks()
        self._locate_blocks()

    def get(self, r, c):
        """r 是 1 起的行号，c 是 0 起的列号。"""
        row = self.rows[r - 1] if r - 1 < len(self.rows) else ()
        return row[c] if c < len(row) else None

    def _locate_weeks(self):
        # 表头某行写着 "2026 Q1"，它所在列就是 2026 W1；周数行是该列 == 1 的那行
        self.w1_col = None
        for r in range(1, 12):
            for c, x in enumerate(self.rows[r - 1]):
                if isinstance(x, str) and x.strip().startswith("2026 Q1"):
                    self.w1_col = c
                    break
            if self.w1_col is not None:
                break
        if self.w1_col is None:
            raise ValueError(f"{self.name}: 找不到 '2026 Q1' 表头")
        self.week_row = next(r for r in range(1, 15) if self.get(r, self.w1_col) == 1)

    def week_col(self, w):
        return self.w1_col + w - 1

    def _locate_blocks(self):
        self.others_rows = [r for r in range(1, len(self.rows) + 1)
                            if self.get(r, COL_C) in OTHERS and self.get(r, COL_LABEL)]
        if not self.others_rows:
            raise ValueError(f"{self.name}: 找不到 Others 行")
        self.others_first = self.others_rows[0]
        self.others_last = self.others_rows[-1]
        self.first_product = next(
            r for r in range(self.others_last + 1, len(self.rows) + 1)
            if self.get(r, COL_LABEL) and self.get(r, COL_G) not in (None, *OTHERS))
        self.last_label_row = max(r for r in range(1, len(self.rows) + 1) if self.get(r, COL_LABEL))

    def gh(self, r):
        """G&H 拼出来的代码；按 Excel 的 & 规则，布尔值写成 TRUE / FALSE。"""
        def text(x):
            return str(x).upper() if isinstance(x, bool) else ("" if x is None else str(x))
        return text(self.get(r, COL_G)) + text(self.get(r, COL_H))

    def code(self, r):
        """汇总时真正用来匹配的 F 列代码。Promo Price 行引用相邻行，返回 None。"""
        f = self.get(r, COL_F)
        if is_formula(self.get(r, COL_G)):
            return None
        if is_formula(f):
            return self.gh(r) if re.fullmatch(rf"=G{r}&H{r}", f.replace("$", "")) else None
        return f

    def product_rows(self, label=None):
        for r in range(self.first_product, self.last_label_row + 1):
            lab = self.get(r, COL_LABEL)
            if lab and (label is None or lab == label) and self.get(r, COL_C) not in OTHERS:
                yield r


def load(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    sheets = {}
    for name in ALL_SHEETS:
        ws = wb[name]
        ws.reset_dimensions()  # 飞书导出不写 dimension，不重置只能读到 1 行
        rows = [[cell_text(x) for x in row] for row in ws.iter_rows(values_only=True)]
        sheets[name] = Sheet(name, rows)
    return sheets


def last_filled_week(s):
    """SO 行里手填数值多于公式的最后一周。"""
    last = 0
    for w in range(1, 53):
        c = s.week_col(w)
        vals = [s.get(r, c) for r in s.product_rows("SO")]
        if sum(map(is_number, vals)) > sum(map(is_formula, vals)):
            last = w
    return last


def ytd(s, r, upto):
    return sum(x for w in range(1, upto + 1) for x in [s.get(r, s.week_col(w))] if is_number(x))


def check_orphans(sheets, upto):
    out = []
    for agg, members in AGG_SHEETS.items():
        agg_codes = {sheets[agg].code(r) for r in sheets[agg].product_rows("SO")}
        for m in members:
            s = sheets[m]
            missing = collections.defaultdict(dict)
            for r in s.product_rows():
                code = s.code(r)
                lab = s.get(r, COL_LABEL)
                if code is None or code in agg_codes or "Others" in str(code):
                    continue
                missing[code].setdefault("name", s.get(r, COL_E))
                if lab in VOLUME_LABELS:
                    missing[code][lab] = ytd(s, r, upto)
                    if lab == "Channel Inventory Actual":
                        last = s.get(r, s.week_col(upto))
                        missing[code]["stock_now"] = last if is_number(last) else 0
            for code, info in missing.items():
                vol = info.get("ST", 0) + info.get("SO", 0) + info.get("stock_now", 0)
                out.append(dict(sheet=m, agg=agg, code=code, name=info["name"],
                                st=info.get("ST", 0), so=info.get("SO", 0),
                                stock=info.get("stock_now", 0), has_volume=vol != 0))
    return out


def check_subtotals(s):
    """小计 SUMIFS 的区间是否覆盖了它该覆盖的行。
    对第 r 行的小计（按 B 或 C 列 + I 列指标匹配），"该覆盖"的是 Others 起往下、B/C 和指标都相同的行；
    区间漏掉其中任何一行就是漏数。区间若同时包含小计区和产品区里匹配的行，就会重复计数。
    只汇总小计行（区间不进产品区）的写法不检查，由各品牌小计负责。"""
    issues = collections.defaultdict(lambda: {"cols": [], "example": None})
    for r in range(1, s.others_first):
        lab = s.get(r, COL_LABEL)
        if not lab:
            continue
        for w in range(1, 53):
            c = s.week_col(w)
            f = s.get(r, c)
            m = SUMIFS_RANGE.search(f) if is_formula(f) else None
            if not m:
                continue
            start, end = int(m.group(1)), int(m.group(2))
            key = COL_B if m.group(3) == "B" else COL_C
            val = s.get(r, key)
            if end < s.first_product:
                continue
            match = lambda x: x != r and s.get(x, key) == val and s.get(x, COL_LABEL) == lab
            expected = [x for x in range(s.others_first, s.last_label_row + 1) if match(x)]
            missed = [x for x in expected if not start <= x <= end]
            doubled = [x for x in range(start, s.others_first) if match(x)]
            problems = []
            if missed:
                n_oth = sum(s.get(x, COL_C) in OTHERS for x in missed)
                parts = [f"{n_oth} 行 Others"] if n_oth else []
                prods = [x for x in missed if s.get(x, COL_C) not in OTHERS]
                if prods:
                    parts.append(f"{len(prods)} 行产品（如第 {prods[0]} 行 {s.get(prods[0], COL_E)}）")
                problems.append(f"区间 {start}–{end} 漏掉 " + "、".join(parts))
            if doubled:
                problems.append(f"区间 {start}–{end} 同时包含小计行（如第 {doubled[0]} 行）和产品行，重复计数")
            for p in problems:
                k = (r, lab, s.get(r, COL_B), s.get(r, COL_C), p)
                issues[k]["cols"].append(w)
                issues[k]["example"] = issues[k]["example"] or f"{get_column_letter(c + 1)}{r}: {f[:120]}"
    return [dict(row=k[0], label=k[1], b=k[2], c=k[3], problem=k[4],
                 weeks=v["cols"], example=v["example"]) for k, v in issues.items()]


def check_duplicates(s):
    seen = collections.defaultdict(list)
    for r in s.product_rows():
        code = s.code(r)
        if code:
            seen[(code, s.get(r, COL_LABEL))].append(r)
    dups = collections.defaultdict(list)
    for (code, lab), rows in seen.items():
        if len(rows) > 1:
            dups[code].append((lab, rows))
    out = []
    for c, v in dups.items():
        rows = sorted({r for _, rs in v for r in rs})
        names = sorted({str(s.get(r, COL_E)) for r in rows})
        out.append(dict(code=c, labels=sorted({lab for lab, _ in v}), names=names,
                        rows=rows))
    return out


def check_code_text(s):
    out = []
    for r in s.product_rows():
        for col, colname in ((COL_G, "G"), (COL_H, "H")):
            x = s.get(r, col)
            if isinstance(x, str) and x != x.strip():
                out.append(dict(row=r, label=s.get(r, COL_LABEL), issue=f"{colname} 列有多余空格: {x!r}"))
        f = s.get(r, COL_F)
        if f is not None and not is_formula(f) and str(f) != s.gh(r):
            out.append(dict(row=r, label=s.get(r, COL_LABEL),
                            issue=f"F 列手写 {f!r}，与 G&H = {s.gh(r)!r} 不一致，这一行会按 {f!r} 取数"))
    return out


def check_live_st(s):
    """ST 行里还是 SUMIFS 公式的周。
    最后一个公式周是"当周"，下周复制前要确认末尾没有残留 +N；更早的周按 SOP 应已转成数值。"""
    weeks = collections.defaultdict(list)
    for r in s.product_rows("ST"):
        for w in range(1, 53):
            f = s.get(r, s.week_col(w))
            if is_formula(f) and "DB ST'!" in f:
                weeks[w].append((r, f))
    if not weeks:
        return None
    current = max(weeks)
    plus = [dict(row=r, code=s.code(r), formula=f[-60:])
            for r, f in weeks[current] if TRAILING_PLUS.search(f)]
    return dict(current=current, plus=plus, past=sorted(w for w in weeks if w != current),
                past_plus=sum(1 for w in weeks if w != current for _, f in weeks[w] if TRAILING_PLUS.search(f)))


def check_identity(s, upto):
    """找最近一周 ST / SO / Actual（本周和上周）都是手填数值、且可比对代码数不少于峰值一半的周，比对库存恒等式。"""
    blocks = {}
    for r in s.product_rows():
        code = s.code(r)
        if code:
            blocks.setdefault(code, {})[s.get(r, COL_LABEL)] = r
    results = {}
    for w in range(upto, 1, -1):
        c, p = s.week_col(w), s.week_col(w - 1)
        rows, complete = [], 0
        for code, b in blocks.items():
            if not all(k in b for k in VOLUME_LABELS):
                continue
            st, so = s.get(b["ST"], c), s.get(b["SO"], c)
            act, prev = s.get(b["Channel Inventory Actual"], c), s.get(b["Channel Inventory Actual"], p)
            if not all(map(is_number, (st, so, act, prev))):
                continue
            complete += 1
            diff = act - (prev + st - so)
            if abs(diff) > max(IDENTITY_MIN_ABS, IDENTITY_MIN_PCT * abs(prev)):
                rows.append(dict(code=code, prev=prev, st=st, so=so, actual=act, diff=diff))
        rows.sort(key=lambda x: -abs(x["diff"]))
        results[w] = (w, complete, rows)
    peak = max((r[1] for r in results.values()), default=0)
    for w in sorted(results, reverse=True):
        if peak and results[w][1] >= peak / 2:
            return results[w]
    return None, 0, []


def md_table(headers, rows):
    lines = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
    lines += ["| " + " | ".join(str(x) for x in row) + " |" for row in rows]
    return "\n".join(lines)


def fmt_weeks(ws):
    return f"W{ws[0]}–W{ws[-1]}（{len(ws)} 周）" if len(ws) > 1 else f"W{ws[0]}"


def build_report(path, sheets):
    filled = {n: last_filled_week(sheets[n]) for n in OM_SHEETS + OP_SHEETS}
    upto = max(filled.values())
    orphans = check_orphans(sheets, upto)
    per_sheet = {n: dict(sub=check_subtotals(sheets[n]), dup=check_duplicates(sheets[n]),
                         text=check_code_text(sheets[n]), live=check_live_st(sheets[n]))
                 for n in ALL_SHEETS}
    identity = {n: check_identity(sheets[n], upto) for n in OM_SHEETS + OP_SHEETS}

    loss = [o for o in orphans if o["has_volume"]]
    risk_orphans = [o for o in orphans if not o["has_volume"]]
    n_sub = sum(len(v["sub"]) for v in per_sheet.values())
    n_dup = sum(len(v["dup"]) for v in per_sheet.values())
    n_text = sum(len(v["text"]) for v in per_sheet.values())
    n_plus = sum(len(v["live"]["plus"]) for v in per_sheet.values() if v["live"])
    n_unfrozen = sum(len(v["live"]["past"]) for v in per_sheet.values() if v["live"])

    L = [f"# PSI 结构校验报告",
         "",
         f"- 文件：`{Path(path).name}`",
         f"- 运行时间：{dt.datetime.now():%Y-%m-%d %H:%M}",
         f"- SO 已填到的最晚一周：W{upto}（累计数按 2026 W1–W{upto} 计）",
         "",
         "## 总览",
         "",
         md_table(["级别", "检查项", "数量"], [
             ["🔴", "有量的代码没进汇总（数据丢失）", len(loss)],
             ["🔴", "小计区间有误（按行 × 问题计）", n_sub],
             ["🟠", "没进汇总、目前还没量的代码", len(risk_orphans)],
             ["🟠", "重复代码", n_dup],
             ["🟠", "代码书写问题（空格 / F 列手写不一致）", n_text],
             ["🟠", "当周 ST 公式末尾带 +N（复制到下周前要删）", n_plus],
             ["🟠", "历史周 ST 还是公式、没转数值（按 sheet × 周计）", n_unfrozen],
         ]),
         "",
         "## 1. 汇总漏行",
         "",
         "KA sheet 里有、但汇总 sheet 里没有对应行的代码，它们的数不会进入汇总。",
         ""]
    if orphans:
        L.append(md_table(["级别", "KA sheet", "缺在", "代码", "产品", f"ST 累计", f"SO 累计", f"W{upto} 库存"],
                          [["🔴" if o["has_volume"] else "🟠", o["sheet"], o["agg"], f"`{o['code']}`",
                            o["name"], o["st"], o["so"], o["stock"]]
                           for o in sorted(orphans, key=lambda o: (not o["has_volume"], o["agg"], o["sheet"]))]))
    else:
        L.append("✅ 没有发现。")

    L += ["", "## 2. 小计区间", ""]
    rows = [[n, i["row"], f"{i['b']} / {i['c']} / {i['label']}", i["problem"], fmt_weeks(i["weeks"]), f"`{i['example']}`"]
            for n, v in per_sheet.items() for i in v["sub"]]
    L.append(md_table(["sheet", "行", "B / C / 指标", "问题", "涉及周", "示例"], rows) if rows else "✅ 没有发现。")

    L += ["", "## 3. 重复代码", ""]
    L += ["同一代码 + 指标出现多次，汇总时会加在一起。产品名不同通常是复制产品块时 G/H 列没改。", ""]
    rows = [[n, f"`{d['code']}`", " / ".join(d["names"]), ", ".join(d["labels"]), ", ".join(map(str, d["rows"]))]
            for n, v in per_sheet.items() for d in v["dup"]]
    L.append(md_table(["sheet", "代码", "产品名", "重复的指标", "行号"], rows) if rows else "✅ 没有发现。")

    L += ["", "## 4. 代码书写", ""]
    rows = [[n, t["row"], t["label"], t["issue"]] for n, v in per_sheet.items() for t in v["text"]]
    L.append(md_table(["sheet", "行", "指标", "问题"], rows) if rows else "✅ 没有发现。")

    L += ["", "## 5. ST 公式状态", "",
          "按 SOP，每周把 SUMIFS 公式复制到新一周，再把上一周转成数值。MEX、Max（OM Others）会在公式末尾手动加 BI 查到的数（+N），"
          "复制到下一周前要先删掉。历史周如果一直是公式，以后改工具表或 DB ST 的旧行，历史数会跟着变。", ""]
    rows = []
    for n, v in per_sheet.items():
        live = v["live"]
        if not live:
            continue
        past = f"{fmt_weeks(live['past'])}，其中 {live['past_plus']} 格带 +N" if live["past"] else "—"
        rows.append([n, f"W{live['current']}", len(live["plus"]), past])
    L.append(md_table(["sheet", "当周（最后一个公式周）", "当周带 +N 的格数", "更早的周还没转数值"], rows)
             if rows else "没有 SUMIFS 公式的 ST 行。")
    plus_rows = [[n, p["row"], f"`{p['code']}`", f"`…{p['formula']}`"]
                 for n, v in per_sheet.items() if v["live"] for p in v["live"]["plus"]]
    if plus_rows:
        L += ["", md_table(["sheet", "行", "代码", "公式结尾"], plus_rows)]

    L += ["", "## 6. 库存恒等式（参考）", "",
          f"本周 Actual 与（上周 Actual + ST − SO）的偏差超过 max({IDENTITY_MIN_ABS} 台, 上周库存 × {IDENTITY_MIN_PCT:.0%})。"
          "偏差可能来自退货、调拨、周报口径变化或映射错位，不一定是错，供复核。"
          f"每张 sheet 列出偏差最大的 {IDENTITY_TOP_N} 个。", ""]
    for n, (w, complete, rows) in identity.items():
        if w is None:
            L.append(f"- **{n}**：没有可比对的周（ST 当周还是公式，或数据不全）")
            continue
        L.append(f"- **{n}** W{w}：可比对 {complete} 个代码，{len(rows)} 个超阈值")
        if rows:
            L += ["", md_table(["代码", "上周库存", "ST", "SO", "本周库存", "偏差"],
                               [[f"`{x['code']}`", x["prev"], x["st"], x["so"], x["actual"], f"{x['diff']:+}"]
                                for x in rows[:IDENTITY_TOP_N]]), ""]

    L += ["", "## 7. 填写进度", "",
          md_table(["sheet", "SO 已填到"], [[n, f"W{w}" if w else "—"] for n, w in filled.items()]), ""]
    return "\n".join(L), dict(loss=len(loss), sub=n_sub)


def main():
    ap = argparse.ArgumentParser(description="PSI 表结构校验")
    ap.add_argument("xlsx", nargs="?", default=str(Path(__file__).resolve().parent.parent / "数据" / "Poland PSI DB+Channel.xlsx"))
    ap.add_argument("--out", default="output")
    args = ap.parse_args()

    sheets = load(args.xlsx)
    report, severe = build_report(args.xlsx, sheets)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"structure_check_{dt.date.today():%Y%m%d}.md"
    path.write_text(report, encoding="utf-8")
    print(f"报告：{path}")
    print(f"🔴 数据丢失 {severe['loss']} 个代码，小计区间问题 {severe['sub']} 处")
    return 1 if any(severe.values()) else 0


if __name__ == "__main__":
    sys.exit(main())
