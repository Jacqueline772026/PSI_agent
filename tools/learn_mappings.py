"""从过去手工处理过的运营商周报里学习 识别码 → 代码，写进 mapping/mapping.xlsx。

Play：'Stock and sales - week N.xlsx' 里人工加的 code 列（SKU → code）
TMO ：'kw_cpfr_main_Xiaomi.xlsx' 里人工加的 code 列（Materiał → code）
同一个识别码在不同周写法不同时，取最近一周的写法，冲突写进 learn_conflicts.md。
已经在 mapping.xlsx 里的 (source, key) 不会被覆盖。

用法：.venv/bin/python -m tools.learn_mappings            （默认读 数据/历史周报/Operators 运营商）
"""
import re
import sys
from pathlib import Path

import pandas as pd

from psi_tool.mapping import append_mappings
from psi_tool.paths import MAPPING, OPERATORS_HISTORY, OUTPUT, PSI_FILE

EXCLUDE = {"AIOT", "OUTLET", "DEMOS", "DEMO", "RESTART", "EXPO"}  # 非手机平板、样机、翻新机
SOURCES = [
    # (mapping 里的 source, 文件名正则, 识别码列, 周数正则)
    ("PLAY_SKU", r".*Stock and sales.*week \d+.*\.xlsx", "SKU", r"week (\d+)"),
    ("TMO_MATERIAL", r".*kw_cpfr_main_Xiaomi.*\.xlsx", "Materiał", r"[Ww](\d+)"),
]


def code_column(df):
    for c in df.columns:
        if str(c).strip().lower() == "code":
            return c
    return None


def learn(folder):
    rows, conflicts = [], []
    for source, file_rx, key_col, week_rx in SOURCES:
        seen = {}
        files = [p for p in Path(folder).rglob("*.xlsx") if re.fullmatch(file_rx, p.name)]
        for p in files:
            m = re.search(week_rx, p.name)
            week = int(m.group(1)) if m else 0
            for sheet, df in pd.read_excel(p, sheet_name=None, dtype=str).items():
                c = code_column(df)
                if c is None or key_col not in df.columns:
                    continue
                name_col = "Model" if "Model" in df.columns else "Opis" if "Opis" in df.columns else key_col
                for _, r in df[[key_col, c, name_col]].dropna(subset=[key_col, c]).iterrows():
                    # 代码里的空格保留（PSI 里就是 'P83 inbox6+128'），匹配 PSI 时 resolve() 会忽略空格比较
                    key, code = str(r[key_col]).strip(), re.sub(r"\s+", " ", str(r[c]).strip())
                    if not key or code.lower() == "nan":
                        continue
                    old = seen.get(key)
                    if old and old[1] != code:
                        conflicts.append((source, key, old[1], old[0], code, week))
                    if not old or week >= old[0]:
                        seen[key] = (week, code, str(r[name_col]))
        for key, (week, code, name) in seen.items():
            exclude = code.upper() in EXCLUDE
            rows.append(dict(source=source, key=key, psi_code="" if exclude else code,
                             category="exclude" if exclude else ("others" if code.upper() == "OTHERS" else "phone/pad"),
                             product_name=name))
    return rows, conflicts


MEM = re.compile(r"(\d{1,2})\s*(?:GB)?\s*[/+]\s*(\d{2,4})\s*(?:GB|G)?", re.I)
DROP = re.compile(r"\b(xiaomi|redmi|ds|dual sim|nfc|wifi|lte|4g|gb|kb|pen|klawiatura|rm)\b")


def name_key(name):
    """'Redmi Note 15 Pro+ 5G 8/256GB DS' 和 PSI 里的 'RM Note 15 Pro+ 5G' + '8+256' 归一成同一个键。"""
    s = re.sub(r"^e\.", "", str(name).lower()).replace("red.", "redmi ")  # Orange 的 E. 前缀、RED. 缩写
    m = MEM.search(s)
    mem = f"{int(m.group(1))}+{int(m.group(2))}" if m else None
    s = s[:m.start()] if m else s  # 内存后面是颜色、DS、配件，不参与匹配
    s = s.replace("+", " plus ")
    s = re.sub(r"(\d)(pro|ultra|max)", r"\1 \2", s)  # 14Pro → 14 pro
    s = DROP.sub(" ", re.sub(r"[^a-z0-9+ ]+", " ", s))
    return " ".join(s.split()), mem


def name_candidates(psi_path, sheet_name, names, source):
    """按 PSI 这张 sheet 自己的 项目名称 + 内存 去匹配运营商周报里的型号名。只收唯一命中的。"""
    from checks.structure_check import COL_E, COL_H, load
    sheet = load(psi_path)[sheet_name]
    index = {}
    for r in sheet.product_rows("SO"):
        e, h = sheet.get(r, COL_E), sheet.get(r, COL_H)
        if isinstance(e, str) and not e.startswith("="):
            index.setdefault((name_key(e)[0], str(h).strip()), set()).add(sheet.code(r))
    rows, misses = [], []
    for name in sorted(set(names)):
        key, mem = name_key(name)
        paren = re.search(r"\(([A-Z0-9]+)\)", name)  # "Pad Pro Wifi (P83)"：括号里直接给了项目代码
        if paren:
            hit = {c for (k, m), v in index.items() for c in v if c and c.startswith(paren.group(1))
                   and not c[len(paren.group(1)):][:1].isalpha()}
        elif mem is None and len(key.split()) >= 2:  # "Redmi A7 Pro"：没写内存，只要这张 sheet 里这个名字只有一行就用它
            hit = set().union(*[v for (k, m), v in index.items() if k == key] or [set()])
        else:
            hit = index.get((key, mem), set())
        if not hit:  # Plus 写 "15T 5G"，PSI 写 "Xiaomi 15T"：去掉 5G 再试，仍要求唯一
            drop5 = lambda k: " ".join(w for w in k.split() if w != "5g")
            hit = set().union(*[v for (k, m), v in index.items() if m == mem and drop5(k) == drop5(key)] or [set()])
        if len(hit) == 1:
            rows.append(dict(source=source, key=re.sub(r"\s+", " ", name.strip()).upper(), psi_code=next(iter(hit)),
                             category="phone/pad", product_name=name))
        else:
            misses.append(name)
    return rows, misses


def model_names(folder, part, name_col, must_have):
    names = set()
    for p in Path(folder).rglob("*.xlsx"):
        if part not in p.as_posix():
            continue
        for df in pd.read_excel(p, sheet_name=None, dtype=str).values():
            if name_col in df.columns and must_have in df.columns:
                names |= set(df[name_col].dropna().astype(str))
    return names


def jet_models(folder):
    names = set()
    for p in Path(folder).rglob("OPL_Stock_Sales*.xlsx"):
        df = pd.read_excel(p, sheet_name="Main", header=None, dtype=str)
        names |= {x for x in df.iloc[:, 0].dropna().astype(str) if x.strip() and x.strip() != "Model"}
    return names


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else str(OPERATORS_HISTORY)
    rows, conflicts = learn(folder)
    n = append_mappings(MAPPING, rows, "learned:历史周报 code 列")
    psi = PSI_FILE
    jet = jet_models(folder)
    for sheet, source, names in (("Plus", "PLUS_NAME", model_names(folder, "Plus/", "Model", "Sprzedaż PLK")),
                                 ("Orange", "ORANGE_NAME", model_names(folder, "Orange/", "Nazwa materiału", "Zapas_POS")),
                                 ("Orange", "ORANGE_JET_NAME", jet)):
        found, misses = name_candidates(psi, sheet, names, source)
        n += append_mappings(MAPPING, found, "auto:按 PSI 项目名+内存匹配（待确认）")
        print(f"{sheet} 型号名：自动匹配 {len(found)} 个，没匹配上 {len(misses)} 个（大多是非手机或老型号）")
    print(f"学到 {len(rows)} 条，新增 {n} 条，写法冲突 {len(conflicts)} 处")
    if conflicts:
        lines = ["| source | 识别码 | 旧写法 | 周 | 新写法 | 周 |", "|---|---|---|---|---|---|"]
        lines += [f"| {s} | {k} | {a} | W{wa} | {b} | W{wb} |" for s, k, a, wa, b, wb in conflicts]
        OUTPUT.mkdir(exist_ok=True)
        (OUTPUT / "learn_conflicts.md").write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    main()
