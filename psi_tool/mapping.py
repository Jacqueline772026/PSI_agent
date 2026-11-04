"""识别码 → PSI 代码。

查找顺序：
  1. mapping/mapping.xlsx（人工确认过的，优先级最高；category=exclude 的行直接丢掉）
  2. TTL ID 表（EAN 列、ID 列 → CRR1 列 = 代码+内存），自动用文件夹里最新的一份

运营商（Play / TMO）常常只给项目代码不带内存，例如 P16。
resolve() 会在目标 sheet 里找这个项目代码，只有一行就用那一行，多行就用产品名里的内存来区分。
"""
import datetime as dt
import re
from pathlib import Path

import pandas as pd

MAPPING_COLUMNS = ["source", "key", "psi_code", "category", "product_name", "added_on", "added_by"]
# 同一个产品的新旧代码：PSI 和 TTL 用右边，BOC、促销表和早期周报还在用左边
CODE_ALIASES = {"Q16U": "BRUSSELS", "Q16P": "MOSCOWP", "Q16": "MOSCOW"}
ALIAS_RX = re.compile(r"^(Q16U|Q16P|Q16)(?=\d|$)")
MEMORY = re.compile(r"(\d{1,2})\s*(?:GB)?\s*[/+]\s*(\d{2,4})\s*(?:GB|G)?", re.I)


def latest_ttl(root):
    files = sorted(Path(root).glob("TTL ID*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not files:
        raise FileNotFoundError("找不到 TTL ID 工具表（TTL ID *.xlsx）")
    return files[-1]


def norm_id(x):
    if pd.isna(x):
        return None
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    return re.sub(r"\.0$", "", str(x).strip())


class Mapper:
    def __init__(self, ttl_path, mapping_path):
        self.ttl_path, self.mapping_path = Path(ttl_path), Path(mapping_path)
        ttl = pd.read_excel(ttl_path, sheet_name="TTL ID", dtype=object)
        ttl = ttl[ttl["CRR1"].notna()]
        self.ttl = {"EAN": {}, "TTL_ID": {}}
        for _, r in ttl.iterrows():
            code = str(r["CRR1"]).strip()
            for kind, col in (("EAN", "EAN"), ("TTL_ID", "ID")):
                k = norm_id(r[col])
                if k:
                    self.ttl[kind].setdefault(k, code)
        if self.mapping_path.exists():
            m = pd.read_excel(self.mapping_path, dtype=str).fillna("")
        else:
            m = pd.DataFrame(columns=MAPPING_COLUMNS)
        self.manual = {(r["source"], r["key"].upper() if r["source"].endswith("_NAME") else r["key"]):
                       (r["psi_code"], r["category"]) for _, r in m.iterrows()}

    def lookup(self, key_type, key):
        """返回 (psi_code, category, 来源)；查不到 psi_code 为 None。"""
        hit = self.manual.get((key_type, key))
        if hit:
            return hit[0] or None, hit[1] or "phone/pad", "mapping.xlsx"
        if key_type in self.ttl and key in self.ttl[key_type]:
            return self.ttl[key_type][key], "phone/pad", "TTL ID"
        return None, None, None


def memory_from_name(name):
    m = MEMORY.search(str(name))
    return f"{int(m.group(1))}+{int(m.group(2))}" if m else None


def canonical(code):
    """旧代码换成 PSI 用的新代码：Q166+256 → MOSCOW6+256。"""
    return ALIAS_RX.sub(lambda m: CODE_ALIASES[m.group(1)], str(code).strip()) if code else code


def legacy_code(code):
    """PSI 代码对应的旧代码（BOC、促销表还在用）：MOSCOW6+256 → Q166+256；没有旧代码的返回空字符串。"""
    for old, new in CODE_ALIASES.items():
        m = re.fullmatch(rf"{new}((?:2|3|4|6|8|12|16)\+\d{{2,4}})?", str(code or ""))
        if m:
            return old + (m.group(1) or "")
    return ""


def resolve(code, name, sheet_codes, project_rows):
    """把 'P16' 这类只有项目代码的结果补成 PSI 里的完整代码。

    sheet_codes: 目标 sheet 的完整代码集合；project_rows: 项目代码 → 该 sheet 里的完整代码列表。
    """
    code = canonical(code)
    if code is None or code in sheet_codes:
        return code
    squeezed = code.replace(" ", "")
    for c in sheet_codes:  # 'LEEDS 4+128' = 'LEEDS4+128'，'P83inbox6+128' = 'P83 inbox6+128'：比较时忽略空格，结果用 PSI 的写法
        if c and str(c).replace(" ", "") == squeezed:
            return c
    options = project_rows.get(squeezed, [])
    if len(options) == 1:
        return options[0]
    mem = memory_from_name(name)
    if mem and f"{squeezed}{mem}" in sheet_codes:
        return f"{squeezed}{mem}"
    return code


def append_mappings(mapping_path, rows, added_by):
    """rows: [{source, key, psi_code, category, product_name}]，已有的 (source, key) 不覆盖。"""
    path = Path(mapping_path)
    old = pd.read_excel(path, dtype=str).fillna("") if path.exists() else pd.DataFrame(columns=MAPPING_COLUMNS)
    seen = set(zip(old["source"], old["key"]))
    new = [dict(r, added_on=dt.date.today().isoformat(), added_by=added_by)
           for r in rows if (r["source"], r["key"]) not in seen]
    out = pd.concat([old, pd.DataFrame(new, columns=MAPPING_COLUMNS)], ignore_index=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    out.to_excel(path, index=False)
    return len(new)
