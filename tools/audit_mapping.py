"""对照表独立核对：不看历史写法，只按“周报原始名称 + TTL ID 表 + PSI 行”重新推一遍代码，再和 mapping.xlsx 比。

每个识别码得到一个结论：
  一致        名称推出的代码和映射表相同
  冲突        名称能推出唯一代码，但映射表写的不一样 → 映射表很可能错
  漏映射      名称能推出代码、这张 PSI 也有这一行，但映射表里没有 → 这行的数没有填进 PSI
  人工确认    胡佳琪手工确认过；如果和名称推出的不一样，依据列会写出来
  无法验证    TTL 和 PSI 都查不到这个型号（新品、代号、老品），需要人确认
  排除待查    映射表写了不计入，但名称看起来是正常手机平板

用法：
  .venv/bin/python -m tools.audit_mapping                         全量核对，打印汇总（完整结果在 tools.export_mapping 导出的对照表里）
  .venv/bin/python -m tools.audit_mapping --out 结果.xlsx          另存一份完整核对结果
  .venv/bin/python -m tools.audit_mapping --name "Redmi 15C 5G 4/128GB Czarny" --sheet TMO   查一个新型号该填什么
"""
import argparse
import re
from collections import defaultdict
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from checks.structure_check import COL_E, COL_G, COL_H, load
from psi_tool.mapping import latest_ttl, resolve
from psi_tool.paths import CONFIG, MAPPING, PSI_FILE, ROOT, TTL_DIR
from psi_tool.sources import load_config, load_source

SHEET_OF = {"TMO_MATERIAL": "TMO", "PLAY_SKU": "Play", "PLUS_NAME": "Plus", "ORANGE_NAME": "Orange",
            "ORANGE_JET_NAME": "Orange"}
MEMORY = re.compile(r"(\d{1,2})\s*(?:GB?)?\s*[/+]\s*(\d{2,4}|1\s*T)\s*(?:GB?|B)?", re.I)
ROM_ONLY = re.compile(r"\b(32|64|128|256|512)\s*GB?\b", re.I)
# 不计入的版本标记；名称被截断时只留前几个字母
NOT_COUNTED = re.compile(r"DEMO|SAMPLE|OUTL|RESTA|\bEXPO\b|\bUZZ\b|^E\.", re.I)
# 代码是 项目号+内存 连写，P196+256 的内存是 6+256，所以只认实际存在的运存档位
CODE_MEMORY = re.compile(r"((?:2|3|4|6|8|12|16)\+\d{2,4})$")
CR_TAIL = re.compile(r"\bCR\s*$")  # 只认大写结尾的 CR：bi、cz 这类小写尾缀是颜色，不排除
NOT_DEVICE = re.compile(r"BUDS|WATCH|BAND|SCOOTER|HULAJ|\bTV\b|MI-TV|VACUUM|ODKURZ|ROBOT|ADAPTER|ŁADOW|LADOW|POWER (?:BANK|ADAPTER)|PROJECTOR|KAMERA|"
                        r"CAMERA|LAMP|\bTAG\b|GŁOŚN|SPEAKER|ROUTER|MIFI|WIFI RANGE|RANGEEXT|EARBUDS|EARPHONES|SŁUCH|"
                        r"PURIFIER|FAN\b|FRYER|FEEDER|FOUNTAIN|KETTLE|SCALE|ETUI|SZKŁO", re.I)
BUNDLE = re.compile(r"\bKB\b|\+ ?PEN|\bPEN\b|INBOX|BUNDLE|\+Ł|\+ŁAD|COVER", re.I)
EXPANSIONS = [(r"XIAOMI\s*RED\.?\s*NO(?:TE)?", "REDMI NOTE "), (r"RED\.?\s*NO(?:TE)?\b", "REDMI NOTE "),
              (r"REDMI\s*NO(?:TE)?\b", "REDMI NOTE "), (r"REDMINOTE", "REDMI NOTE "), (r"XIAOMIREDMI", "REDMI "),
              (r"\bRM\s+NOTE\b", "REDMI NOTE"), (r"\bRM\b", "REDMI"), (r"\bRED\.?\b(?!MI)", "REDMI"), (r"REDMI(?=\d)", "REDMI ")]


def memory_of(name):
    m = MEMORY.search(str(name).replace("\xa0", " "))
    if m:
        rom = m.group(2).upper().replace(" ", "")
        return f"{int(m.group(1))}+{1024 if rom == '1T' else int(rom)}"
    return None


def rom_of(name):
    m = ROM_ONLY.search(str(name))
    return m.group(1) if m else None


def model_of(name):
    """原始名称 → 统一的型号写法（去品牌、颜色、内存、尾缀）。REDMI / POCO 保留 5G，Xiaomi 数字系列的 5G 没有区分意义。"""
    s = str(name).replace("\xa0", " ").upper().strip()
    if s.startswith("E."):
        s = s[2:]
    m = MEMORY.search(s) or ROM_ONLY.search(s)
    tail_5g = bool(m and re.search(r"\b5G\b", s[m.end():]))  # Tablet ... 6GB/128GB 5G：5G 写在内存后面
    if m:
        s = s[:m.start()]
    if tail_5g:
        s += " 5G"
    for pat, rep in EXPANSIONS:
        s = re.sub(pat, rep, s)
    s = re.sub(r"\(.*?\)|\(.*$|\bTABLET\b|\bDS\b|\bNFC\b|\bLTE\b|\b4G\b|\bTBC\b|\bWIFI\b|- ?BUNDLE ONLY", " ", s)
    s = re.sub(r"(?<=\d)(?=PRO|ULTRA|LITE|MAX|5G\b)|(?<=PRO)(?=\d)|(?<=NOTE)(?=\d)|(?<=PRO)(?=5G)|(?<=\+)(?=5G)", " ", s)
    s = re.sub(r"\bXIAOMI\b|\bMI\b(?= \d)", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" -")
    if not s.startswith(("REDMI", "POCO")):
        s = re.sub(r"\b5G\b", "", s).strip()
    return re.sub(r"\s+", " ", s)


def ttl_models(ttl):
    """TTL 的 Short Name 新品常常是空的：用同一代码下各颜色 Name 的共同前缀当型号。"""
    prefix = {}
    for crr, g in ttl.groupby("CRR1"):
        names = [re.split(r"\s+\d{1,2}\s*GB?\s*RAM|\s+\d{1,2}/\d+", str(n), maxsplit=1)[0].split() for n in g["Name"]]
        common = []
        for words in zip(*names):
            if len(set(w.upper() for w in words)) > 1:
                break
            common.append(words[0])
        if len(g) == 1 or len(common) < 2:  # 只有一个颜色时去掉颜色词
            colour = str(g["Color"].iloc[0]).upper()
            common = [w for w in names[0] if w.upper() != colour][:-1 if str(g["Color"].iloc[0]) != "nan" else None] or names[0]
        prefix[crr] = " ".join(common)
    def pick(r):
        short = r["Short Name"]
        return str(short) if isinstance(short, str) and short.strip() else prefix.get(r["CRR1"], "")
    return ttl.apply(pick, axis=1)


class Evidence:
    """TTL 和 PSI 两边的 型号 + 内存 → 代码。"""

    def __init__(self, psi_path, ttl_path):
        self.psi = load(psi_path)
        self.values = load_workbook(psi_path, read_only=False, data_only=True)  # 项目名称有的是公式，要取计算结果
        self.sheets = {}
        for name in set(SHEET_OF.values()):
            sheet = self.psi[name]
            codes, projects, by_model, labels = set(), defaultdict(list), defaultdict(set), {}
            for r in sheet.product_rows("SO"):
                code, mem = sheet.code(r), str(sheet.get(r, COL_H)).strip()
                label = self.label(name, r)
                codes.add(code)
                projects[str(sheet.get(r, COL_G)).strip()].append(code)
                if label:
                    labels[code] = label
                    by_model[(model_of(label), mem)].add(code)
            self.sheets[name] = dict(codes=codes, projects=projects, by_model=by_model, labels=labels)
        ttl = pd.read_excel(ttl_path, sheet_name="TTL ID", dtype=object)
        self.ttl = ttl[ttl["CRR1"].notna()].copy()
        self.ttl["型号"] = ttl_models(self.ttl)
        self.by_ttl = defaultdict(set)
        for _, r in self.ttl.iterrows():
            crr = str(r["CRR1"]).strip()
            mem = CODE_MEMORY.search(crr)
            self.by_ttl[(model_of(r["型号"]), mem.group(1) if mem else "")].add(crr)

    def label(self, sheet_name, row):
        v = self.values[sheet_name].cell(row, COL_E + 1).value
        return v.strip() if isinstance(v, str) and v.strip() else None

    def candidates(self, name, sheet):
        """返回 (TTL 推出的代码集合, 这张 PSI 表按标签推出的代码集合)。"""
        model, mem = model_of(name), memory_of(name)
        info = self.sheets[sheet]
        return self.by_ttl.get((model, mem), set()), set(info["by_model"].get((model, mem), set()))


def explicit_project(name, info):
    """Jet 表把项目号写在括号里：Pad Pro Wifi (P83)。"""
    m = re.search(r"\(([A-Z0-9]+)\)", str(name).upper())
    return set(info["projects"].get(m.group(1), [])) if m else set()


def audit_entry(ev, source, key, name, code, category, added_by=""):
    sheet = SHEET_OF[source]
    info = ev.sheets[sheet]
    mapped = resolve(code, name, info["codes"], info["projects"]) if code else None
    is_others = bool(mapped) and mapped.upper() == "OTHERS"
    filled = bool(mapped) and mapped in info["codes"] and not is_others
    base = dict(来源=source, PSI表=sheet, 识别码=key, 原始名称=name, 型号=model_of(name), 内存=memory_of(name) or "",
                映射表代码=mapped or "", 类别=category)

    def out(verdict, why, suggestion=""):
        return dict(base, 结论=verdict, 建议代码=suggestion, 依据=why)

    if re.search(r"\bCOAR\b", name):
        return out("无法验证", "待办：COAR 的含义和是否计入待胡佳琪问负责人，暂时保持原来的处理")
    has_memory = bool(memory_of(name))
    if NOT_COUNTED.search(name) or CR_TAIL.search(name):
        return out("一致", "样机 / 翻新 / Outlet / 特殊版本：工具按名称规则或映射表排除，不计入")
    if NOT_DEVICE.search(name) and not has_memory:
        return out("一致" if not filled else "冲突", "名称是非手机平板，不计入")
    ttl, psi = ev.candidates(name, sheet)
    want = (ttl & info["codes"]) or psi or explicit_project(name, info)
    if not want and not has_memory:  # 名称不带内存：这张 PSI 里该型号只有一行时才能确定
        rows = {c for (m, _), cs in info["by_model"].items() if m == base["型号"] for c in cs}
        want = rows if len(rows) == 1 else set()
    human = "胡佳琪" in added_by
    if category == "exclude":
        return out("排除待查", "映射表写不计入，但名称看起来是正常手机平板", "、".join(sorted(want)))
    if human:
        if mapped in want or (is_others and not want):
            return out("人工确认", f"{added_by}；和名称推出的一致")
        return out("人工确认", f"{added_by}；名称推出 {'、'.join(sorted(want or ttl)) or '无'}，以人工为准，建议再看一眼",
                   "、".join(sorted(want)))
    if mapped in want:
        return out("一致", f"{'TTL + PSI' if ttl & psi else ('TTL' if ttl else 'PSI 标签')} 都指向 {mapped}", mapped)
    if not want:
        if ttl:
            why = f"TTL 是 {'、'.join(sorted(ttl))}，这张 PSI 没有这一行，不用填"
            return out("冲突", why + f"；映射表却填到 {mapped}") if filled else out("一致", why)
        if not filled:
            return out("一致", "PSI 没有这个型号这一行，不用填；如果是新品，要考虑在 PSI 加行")
        return out("无法验证", "TTL 和 PSI 都查不到这个型号+内存", mapped)
    if len(want) > 1:
        return out("无法验证", "名称对应多个代码，需要人工确认", "、".join(sorted(want)))
    target = next(iter(want))
    why = f"名称推出 {target}（TTL：{'、'.join(sorted(ttl)) or '无'}；PSI 标签：{'、'.join(sorted(psi)) or '无'}）"
    if not code:
        return out("漏映射", why, target)
    return out("冲突", why + f"；映射表是 {mapped}", target)


def audit_ttl(ev):
    """TTL 表自身：同型号同内存应只有一个代码；Code 和 CODE 列应该一样；代码应该在 PSI 里出现过。"""
    rows = []
    t = ev.ttl.copy()
    t["内存"] = t["CRR1"].astype(str).str.extract(CODE_MEMORY.pattern, expand=False)
    for (short, mem), g in t.groupby(["型号", "内存"]):
        codes = sorted(set(g["CRR1"].astype(str).str.strip()))
        if len(codes) > 1:
            rows.append(dict(问题="同型号同内存有多个代码", 型号=short, 内存=mem, 代码="、".join(codes), 物料数=len(g), 示例=g["Name"].iloc[0]))
    diff = t[t["Code"].astype(str).str.strip() != t["CODE"].astype(str).str.strip()]
    for (a, b), g in diff.groupby(["Code", "CODE"]):
        rows.append(dict(问题="Code 和 CODE 列不一致", 型号=g["Short Name"].iloc[0], 内存="", 代码=f"{a} / {b}", 物料数=len(g), 示例=g["Name"].iloc[0]))
    for _, r in t.iterrows():
        name, crr = str(r["Name"]), str(r["CRR1"]).strip()
        if NOT_DEVICE.search(name) and re.search(r"\d+\+\d+$", crr):
            rows.append(dict(问题="名称不是手机平板却有机型代码", 型号=r["Short Name"], 内存=r["Unnamed: 14"], 代码=crr, 物料数=1, 示例=f"{r['EAN']} {name}"))
        mem = re.search(r"(\d{1,2})G(?:B)? RAM (\d{2,4})G(?:B)? ROM", name)
        col = str(r["Unnamed: 14"]).strip()
        if mem and col and col != "nan" and f"{mem.group(1)}+{mem.group(2)}" != col:
            rows.append(dict(问题="名称内存和代码内存不一致", 型号=r["Short Name"], 内存=r["Unnamed: 14"], 代码=crr, 物料数=1, 示例=f"{r['EAN']} {name}"))
    dup = t[t["EAN"].notna()].groupby("EAN")["CRR1"].nunique()
    for ean in dup[dup > 1].index:
        g = t[t["EAN"] == ean]
        rows.append(dict(问题="同一 EAN 有多个代码", 型号=g["Short Name"].iloc[0], 内存="", 代码="、".join(sorted(set(g["CRR1"].astype(str)))), 物料数=len(g), 示例=str(ean)))
    return pd.DataFrame(rows, columns=["问题", "型号", "内存", "代码", "物料数", "示例"])


def audit_psi_labels(ev):
    """同一个代码在不同 PSI 表里的项目名称应该一样；也要和 TTL 的型号一致。"""
    labels = defaultdict(dict)
    for sheet in ev.psi:
        s = ev.psi[sheet]
        if not hasattr(s, "product_rows") or sheet in ("OM", "OP", "📌ALL"):
            continue
        for r in s.product_rows("SO"):
            e = ev.label(sheet, r)
            if e and s.code(r):
                labels[s.code(r)][sheet] = e
    ttl_model = defaultdict(set)
    for _, r in ev.ttl.iterrows():
        ttl_model[str(r["CRR1"]).strip()].add(str(r["型号"]).strip())
    rows = []
    for code, by_sheet in sorted(labels.items()):
        names = {model_of(v) for v in by_sheet.values()}
        ttl = {model_of(v) for v in ttl_model.get(code, set())}
        # TTL 里新品的型号常常只有开发代号（xx-EU）或半截名称，比不出结果，不算问题
        noisy = any(re.search(r"-EU|[^\x00-\x7f]|^TABLET|PEN|^PAD\b", t) for t in ttl) or code.startswith(("P82", "P81", "O16P"))
        if len(names) > 1 or (ttl and not names & ttl and not noisy):
            rows.append(dict(代码=code, PSI里的名称="；".join(f"{k}={v}" for k, v in sorted(by_sheet.items())),
                             TTL型号="、".join(sorted(ttl_model.get(code, set()))) or "（TTL 没有）",
                             问题="各表名称不一致" if len(names) > 1 else "和 TTL 型号不一致"))
    return pd.DataFrame(rows, columns=["代码", "PSI里的名称", "TTL型号", "问题"])


def report_rows(week_folders, config_path):
    """实际周报里出现过的 识别码 → (名称, 最后一周, SO, 库存)；映射表没有的型号也能查到。"""
    cfg = load_config(config_path)
    seen = {}
    for week, inbox in sorted(week_folders.items()):
        for name, c in cfg.items():
            source = c["key"]["type"]
            if source not in SHEET_OF or not Path(inbox).exists():
                continue
            res = load_source(inbox, name, c, week)
            if res.rows is None:
                continue
            for _, r in res.rows.iterrows():
                seen[(source, r["key"])] = dict(name=r["name"], week=week, so=r["so"], stock=r["stock"])
    return seen


def run_audit(psi_path, mapping_path, week_folders=None, config_path=CONFIG):
    ev = Evidence(psi_path, latest_ttl(TTL_DIR))
    m = pd.read_excel(mapping_path, dtype=str).fillna("")
    m = m[m["source"].isin(SHEET_OF)]
    seen = report_rows(week_folders or {}, config_path)
    table = {(r["source"], r["key"].upper() if r["source"].endswith("_NAME") else r["key"]): r for _, r in m.iterrows()}
    rows = []
    for key in sorted(set(table) | set(seen), key=str):
        r = table.get(key, {})
        info = seen.get(key, {})
        name = str(r.get("product_name") or info.get("name") or key[1])
        row = audit_entry(ev, key[0], key[1], name, r.get("psi_code", ""), r.get("category", ""), r.get("added_by", ""))
        row.update(最后一周=f"W{info['week']}" if info else "", 周报SO=info.get("so", ""), 周报库存=info.get("stock", ""),
                   来历=r.get("added_by", ""))
        rows.append(row)
    return pd.DataFrame(rows), audit_ttl(ev), audit_psi_labels(ev), ev


def write(df, ttl, labels, path):
    order = {"冲突": 0, "漏映射": 1, "排除待查": 2, "无法验证": 3, "人工确认": 4, "一致": 5}
    df = df.assign(_o=df["结论"].map(order)).sort_values(["_o", "PSI表", "型号"]).drop(columns="_o")
    summary = pd.crosstab(df["PSI表"], df["结论"]).reset_index()
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="汇总", index=False)
        df.to_excel(xw, sheet_name="逐条核对", index=False)
        ttl.to_excel(xw, sheet_name="TTL表的问题", index=False)
        labels.to_excel(xw, sheet_name="PSI名称的问题", index=False)
    wb = load_workbook(path)
    fills = {"冲突": "F8CBAD", "漏映射": "F8CBAD", "排除待查": "FFE699", "无法验证": "FFF2CC"}
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for c in ws[1]:
            c.font = Font(bold=True)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(60, max(10, max(len(str(c.value or "")) for c in col[:80]) + 2))
    ws = wb["逐条核对"]
    head = [c.value for c in ws[1]].index("结论")
    for row in ws.iter_rows(min_row=2):
        color = fills.get(row[head].value)
        if color:
            for c in row:
                c.fill = PatternFill("solid", fgColor=color)
    wb.save(path)


def suggest(name, sheet, psi_path):
    ev = Evidence(psi_path, latest_ttl(TTL_DIR))
    ttl, psi = ev.candidates(name, sheet)
    print(f"原始名称：{name}\n型号：{model_of(name)}　内存：{memory_of(name)}")
    print(f"TTL 推出：{'、'.join(sorted(ttl)) or '无'}\n{sheet} 表标签推出：{'、'.join(sorted(psi)) or '无'}")
    info = ev.sheets[sheet]
    want = (ttl & info["codes"]) or psi
    if len(want) == 1:
        print(f"建议代码：{next(iter(want))}")
    elif ttl and not want:
        print("TTL 有这个型号，但这张 PSI 没有这一行：不用填")
    else:
        print("拿不准：请核对 EAN / SKU，或问负责人")
    print("提示：套装、B2B、特殊版本要再看一眼尾缀")


def main():
    ap = argparse.ArgumentParser(description="对照表独立核对")
    ap.add_argument("--psi", default=str(PSI_FILE))
    ap.add_argument("--mapping", default=str(MAPPING))
    ap.add_argument("--out")
    ap.add_argument("--week", type=int, action="append", default=[], help="把这一周的原始周报也纳入（可多次），默认 W44")
    ap.add_argument("--inbox", action="append", default=[], help="和 --week 一一对应的文件夹")
    ap.add_argument("--name", help="查一个型号该填什么")
    ap.add_argument("--sheet", default="TMO", help="配合 --name：PSI 表名")
    a = ap.parse_args()
    if a.name:
        suggest(a.name, a.sheet, a.psi)
        return
    weeks = dict(zip(a.week, a.inbox)) or {44: ROOT / "W44测试" / "W44原始文件"}
    df, ttl, labels, _ = run_audit(a.psi, a.mapping, weeks)
    print(pd.crosstab(df["PSI表"], df["结论"]).to_string())
    print(f"TTL 问题 {len(ttl)} 条，PSI 名称问题 {len(labels)} 条")
    if a.out:
        write(df, ttl, labels, Path(a.out))
        print(f"结果：{a.out}")


if __name__ == "__main__":
    main()
