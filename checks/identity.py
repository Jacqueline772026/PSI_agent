"""PSI 库存恒等式：本周 Channel Inventory Actual ≈ 上周 Actual + 本周 ST − 本周 SO。

难点是 ST：飞书导出的 xlsx 里没有公式的计算结果，这里照着 PSI 的公式自己算一遍：
  - 手填数字（运营商从 BI 查过、标紫色的格子）直接用
  - OM 各页：SUMIFS('2026 DB ST'!J, 代码, 客户, 周)，DB ST 页的 代码 / 客户 也照 PSI 公式查：
      代码 = 工具表1 按 69 码（L 列）或 AB 物料码（K 列）查 项目 + 配置去掉 G（'4G+128G' → '4+128'）
      客户 = 工具表2 按客户全名（I 列）查简称
    Max（OM Others）等公式末尾手加的 +N 是从 BI 加的数，一起算进去
  - '=DZ61' 这种引用（运营商没填 BI 时 ST 显示的是 ST FCST）：照引用取值，并标记"仍是预测值"

用法：
  .venv/bin/python -m checks.identity --week 44 [--psi 文件] [--validate 43]
  --validate：用当周的公式模板去算 W43，和 W43 已经转成数值的 ST 逐格比对，检查这里的复刻和飞书是否一致
"""
import argparse
import collections
import re

import openpyxl
from openpyxl.utils import column_index_from_string

from checks.structure_check import COL_LABEL, is_number, load
from psi_tool.paths import PSI_FILE

DB_ST, SKU_TABLE, CHANNEL_TABLE = "2026 DB ST", "工具表1_SKUID对应表 - GDP 平台导出", "工具表2_渠道建成对应表"
OM = ["EURO", "MEX", "X-KOM", "MSH", "Michain", "OM Others"]
OP = ["Orange", "Play", "TMO", "Plus"]
MIN_ABS, MIN_PCT = 50, 0.2  # 偏差超过 max(50 台, 上周库存 × 20%) 才列出
REF = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")


def text(v):
    return getattr(v, "text", v)  # ArrayFormula → 公式文本


def key(v):
    """XLOOKUP 精确匹配：不分大小写；数字和文本的 69 码统一成文本。"""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).casefold()


class DbSt:
    """'2026 DB ST' 页按 PSI 公式算出 (周, 代码, 客户简称) → 数量。"""

    def __init__(self, path, weeks):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
        by_ean, by_sku = {}, {}
        ws = wb[SKU_TABLE]
        ws.reset_dimensions()
        for r in ws.iter_rows(min_row=2, values_only=True):
            proj, cfg, sku, ean = r[1], r[2], r[6], r[12]
            by_ean.setdefault(key(ean), (proj, cfg))  # XLOOKUP 取第一个匹配
            by_sku.setdefault(key(sku), (proj, cfg))
        channel = {}
        ws = wb[CHANNEL_TABLE]
        ws.reset_dimensions()
        for r in ws.iter_rows(min_row=2, values_only=True):
            if len(r) > 4 and r[4] is not None:
                channel.setdefault(key(r[4]), r[3])
        self.qty = collections.Counter()
        self.rows = collections.Counter()  # 周 → DB ST 页这一周有几行（0 = 还没贴，OM 的 ST 会全是 0）
        self.unmatched = collections.Counter()  # (周, 原因) → 数量：查不到代码或客户的行，PSI 里也是 0
        ws = wb[DB_ST]
        ws.reset_dimensions()
        for r in ws.iter_rows(min_row=3, values_only=True):
            if r[0] not in weeks or not is_number(r[9]):
                continue
            self.rows[r[0]] += 1
            d = text(r[3])
            if isinstance(d, str) and d.startswith("="):
                look = by_sku.get(key(r[10])) if "!G:G" in d else by_ean.get(key(r[11]))
            else:
                look = (d, None)  # 直接贴的值
            ka = channel.get(key(r[8]))
            if not look or not look[0] or not ka:
                self.unmatched[(r[0], "代码" if not look or not look[0] else "客户")] += r[9]
                continue
            proj, cfg = look
            code = f"{proj}{str(cfg).replace('G', '') if cfg else ''}"
            self.qty[(r[0], code, ka)] += r[9]


def split_terms(f):
    """'=A+B-C' 按最外层的 + - 拆开，保留符号；括号和引号里的不拆。"""
    out, depth, quote, cur, sign = [], 0, False, "", 1
    for ch in f.lstrip("="):
        if ch == '"':
            quote = not quote
        elif not quote and ch == "(":
            depth += 1
        elif not quote and ch == ")":
            depth -= 1
        if not quote and depth == 0 and ch in "+-" and cur.strip():
            out.append((sign, cur.strip()))
            sign, cur = (1 if ch == "+" else -1), ""
            continue
        cur += ch
    if cur.strip():
        out.append((sign, cur.strip()))
    return out


def split_args(s):
    out, depth, quote, cur = [], 0, False, ""
    for ch in s:
        if ch == '"':
            quote = not quote
        elif not quote and ch == "(":
            depth += 1
        elif not quote and ch == ")":
            depth -= 1
        if ch == "," and depth == 0 and not quote:
            out.append(cur.strip())
            cur = ""
        else:
            cur += ch
    out.append(cur.strip())
    return out


class StEvaluator:
    def __init__(self, sheet, db):
        self.s, self.db = sheet, db

    def ref(self, a1):
        m = REF.match(a1.replace("$", "") if a1 else "")
        if not m:
            return None
        return int(m.group(2)), column_index_from_string(m.group(1)) - 1  # get() 的列号比 Excel 少 1

    def cell(self, a1, week, depth):
        rc = self.ref(a1)
        if rc is None:
            return None, "未知"
        r, c = rc
        v = self.s.get(r, c)
        label = self.s.get(r, COL_LABEL)
        if v is None:
            return 0, "FCST" if label == "ST FCST" else "空"
        if is_number(v):
            return v, "FCST" if label == "ST FCST" else "手填"
        if isinstance(v, str) and v.startswith("=") and depth < 4:
            return self.formula(v, week, depth + 1, r)
        return None, "未知"

    def sumifs(self, args, week, row):
        crit = {}
        for rng, val in zip(args[1::2], args[2::2]):
            col = rng.split("!")[-1].replace("$", "").split(":")[0]
            if val.startswith('"'):
                crit[col] = val.strip('"')
            elif col == "A":
                crit[col] = week  # 周的条件（DZ$5）用要算的那一周
            elif col == "F":
                crit[col] = self.s.code(row)
            else:
                rc = self.ref(val)
                crit[col] = None if rc is None else self.s.get(*rc)
        if set(crit) != {"A", "F", "H"} or DB_ST not in args[0]:
            return None
        return self.db.qty.get((crit["A"], crit["F"], crit["H"]), 0)

    def formula(self, f, week, depth=0, row=None):
        total, kinds = 0, set()
        for sign, term in split_terms(f):
            if term.upper().startswith("SUMIFS("):
                v = self.sumifs(split_args(term[7:-1]), week, row)
                kind = "DB"
            elif is_number(term) or re.fullmatch(r"\d+(\.\d+)?", term):
                v, kind = float(term), "手加"
            else:
                v, kind = self.cell(term, week, depth)
            if v is None:
                return None, "未知"
            total += sign * v
            kinds.add(kind)
        return total, "+".join(sorted(kinds))

    def st(self, row, week, template_week=None):
        """这一行这一周的 ST。template_week：用那一周的公式（只换周）来算，验证时用。"""
        v = self.s.get(row, self.s.week_col(template_week or week))
        if is_number(v) and not template_week:
            return v, "手填"
        if isinstance(v, str) and v.startswith("="):
            return self.formula(v, week, 0, row)
        return (0, "空") if v is None else (None, "未知")


def blocks(sheet):
    out = {}
    for r in sheet.product_rows():
        code = sheet.code(r)
        if code:
            out.setdefault(code, {})[sheet.get(r, COL_LABEL)] = r
    return out


def check(psi, db, week, current=None, sheets=OM + OP):
    """current：{sheet: {code: (so, stock)}}（工具结果）；不给就用 PSI 里已填的数。
    返回 {sheet: dict(checked, flagged=[...], fcst=[...], unknown=n)}。"""
    out = {}
    for name in sheets:
        s = psi[name]
        ev = StEvaluator(s, db)
        c, p = s.week_col(week), s.week_col(week - 1)
        res = dict(checked=0, flagged=[], fcst=[], unknown=0)
        for code, b in blocks(s).items():
            if not all(k in b for k in ("ST", "SO", "Channel Inventory Actual")):
                continue
            st, kind = ev.st(b["ST"], week)
            if "FCST" in kind and st:
                res["fcst"].append(dict(code=code, st=st))
            if current is not None:
                so, act = current.get(name, {}).get(code, (0, 0))
            else:
                so, act = s.get(b["SO"], c), s.get(b["Channel Inventory Actual"], c)
            prev = s.get(b["Channel Inventory Actual"], p)
            if st is None:
                res["unknown"] += 1
                continue
            if not all(map(is_number, (so, act, prev))):
                continue
            res["checked"] += 1
            diff = act - (prev + st - so)
            limit = max(MIN_ABS, MIN_PCT * abs(prev))
            if abs(diff) > limit:
                if so == 0 and act == 0 and prev > 0:
                    why = "本周 SO 和库存都是 0：可能缺文件（如 EURO 平板）或这一行没填"
                elif "FCST" in kind:
                    why = "ST 是预测值，先换成 BI 的数"
                elif st > 0 and abs(act - (prev - so)) <= limit:
                    why = "库存像没收到这批 ST：在途 / 时间差，下周再看"
                elif st == 0 and diff > 0:
                    why = "库存涨了但没有 ST：ST 漏了、调拨，或填错行"
                else:
                    why = "要查：填错行、SO 和库存填反、退货"
                res["flagged"].append(dict(code=code, prev=prev, st=st, so=so, actual=act, diff=diff, st_from=kind,
                                           why=why))
        res["flagged"].sort(key=lambda x: -abs(x["diff"]))
        out[name] = res
    return out


def validate(psi, db, week, template_week):
    """用 template_week 的公式算 week 的 ST，和 week 已经转成数值的 ST 比。"""
    out = {}
    for name in OM:
        s = psi[name]
        ev = StEvaluator(s, db)
        same = diff = 0
        rows = []
        for code, b in blocks(s).items():
            if "ST" not in b:
                continue
            real = s.get(b["ST"], s.week_col(week))
            if not is_number(real):
                continue
            calc, _ = ev.st(b["ST"], week, template_week=template_week)
            if calc is None:
                continue
            if round(calc) == round(real):
                same += 1
            else:
                diff += 1
                rows.append((code, real, calc))
        out[name] = (same, diff, rows)
    return out


def report_lines(res, week, db, md_table, source):
    L = [f"## 库存恒等式（W{week}）", "",
         f"本周库存 ≈ 上周库存（PSI）+ 本周 ST（按 PSI 公式算）− 本周 SO（{source}）。"
         f"偏差超过 max({MIN_ABS} 台, 上周库存 × {int(MIN_PCT * 100)}%) 的列出来：可能是填错行、SO 和库存填反、退货调拨，或 ST 还没更新。", ""]
    if not db.rows[week]:
        L += [f"- ⚠️ PSI 导出里 '2026 DB ST' 还没有 W{week} 的行，OM 各页的 ST 都会按 0 算。贴完 DB 页、"
              "运营商 BI 数标紫后重新导出 PSI 再跑，这一节才准", ""]
    miss = {k[1]: v for k, v in db.unmatched.items() if k[0] == week}
    if miss:
        L += [f"- DB ST 页 W{week} 有 {sum(miss.values()):.0f} 台查不到"
              + "、".join(f"{k}（{v:.0f} 台）" for k, v in miss.items()) + "，PSI 公式里同样算不进 ST", ""]
    rows = []
    for name, r in res.items():
        state = "✅" if not r["flagged"] else f"⚠️ {len(r['flagged'])} 个"
        rows.append([name, r["checked"], state, len(r["fcst"]) or "", r["unknown"] or ""])
    L += [md_table(["sheet", "可比对代码", "超阈值", "ST 仍是预测值", "ST 算不出"], rows), ""]
    for name, r in res.items():
        if r["flagged"]:
            L += [f"**{name}**", "", md_table(["代码", "上周库存", "ST", "ST 来源", "SO", "本周库存", "偏差", "可能原因"],
                                               [[x["code"], int(x["prev"]), int(x["st"]), x["st_from"], int(x["so"]),
                                                 int(x["actual"]), f"{x['diff']:+.0f}", x["why"]]
                                                for x in r["flagged"][:15]]), ""]
        if r["fcst"]:
            L += [f"**{name}：ST 还是 ST FCST（预测值），没有换成 BI 查的数**："
                  + "、".join(f"{x['code']}（{x['st']:.0f}）" for x in r["fcst"]), ""]
    return L


def main():
    from psi_tool.run import md_table
    ap = argparse.ArgumentParser()
    ap.add_argument("--week", type=int, required=True)
    ap.add_argument("--psi", default=str(PSI_FILE))
    ap.add_argument("--validate", type=int, help="用 --week 的公式模板算这一周，和已转成数值的 ST 比对")
    a = ap.parse_args()
    psi = load(a.psi)
    weeks = {a.week} | ({a.validate} if a.validate else set())
    db = DbSt(a.psi, weeks)
    if a.validate:
        print(f"== 用 W{a.week} 的公式算 W{a.validate} 的 ST，和 PSI 里 W{a.validate} 的数值比")
        for name, (same, diff, rows) in validate(psi, db, a.validate, a.week).items():
            print(f"{name}: 一致 {same}，不一致 {diff}", rows[:10])
    print("\n".join(report_lines(check(psi, db, a.week), a.week, db, md_table, "PSI 已填")))


if __name__ == "__main__":
    main()
