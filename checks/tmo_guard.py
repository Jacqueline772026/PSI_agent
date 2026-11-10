"""TMO 防串检查：只提示，不改数字。

TMO 的产品名最多约 40 个字符，经常被截断（biały → bi，Czarny → Cza），所以名称尾缀不能直接当规则用。
更可靠的依据是物料号：XA-H-T117040 = 型号 T117 + 颜色 / 版本 040，同一型号应该对应同一个 PSI 代码。

查四件事：
  1. 被名称规则排除的行：尾缀是颜色缩写 → 规则可能误伤（❌）；其余列出来留痕（ℹ️）
  2. 名称结尾的短词既不是颜色，也不是已知的版本标记 → 要人确认（⚠️）
  3. 同型号代码不一致、没匹配但同型号其他颜色有代码 → 可能串了或漏了（⚠️）
  4. 名称里的内存和代码里的内存对不上（⚠️）

用法：报告里自动带；对历史周报回测：.venv/bin/python -m checks.tmo_guard
"""
import re
from collections import defaultdict

from psi_tool.mapping import memory_from_name

# 波兰语颜色。名称可能只留前几个字母，所以按"是不是某个颜色的开头"判断（bi → biały，nieb → niebieski）
COLOURS = ("czarny", "czarne", "biały", "biała", "białe", "bialy", "niebieski", "niebieskie", "zielony", "zielone",
           "fioletowy", "fioletowe", "szary", "szare", "złoty", "zloty", "srebrny", "srebrne", "brązowy", "brazowy",
           "purpurowy", "lawendowy", "lavender", "tytan", "titan", "czerwony", "różowy", "rozowy", "pomarańczowy",
           "żółty", "zolty", "granatowy", "grafitowy", "mocha", "jasnoniebieski", "jasnofioletowy",
           "ciemnyniebieski", "ciemnyfioletowy")
# 已知的版本标记（大写）。计不计入由 config 的 exclude_name 决定，这里只表示"认识"
KNOWN_TAGS = {"CR", "B2B", "COAR", "OS", "HR", "TAPP", "DEMO", "EŚ", "EXPO", "OUTL", "OUTLET", "NPS", "RESTAR", "RESTART"}
MATERIAL = re.compile(r"XA-H-T(\d+)(\d{3})")
CODE_MEMORY = re.compile(r"(\d{1,2}\+\d{2,4})$")


def tail_of(name):
    words = str(name).split()
    return words[-1] if words else ""


def is_colour(word):
    w = word.lower()
    return len(w) >= 2 and any(c.startswith(w) for c in COLOURS)


def family_of(key):
    m = MATERIAL.fullmatch(str(key))
    return m.group(1) if m else None


def check_tmo(rows, name_excluded=None):
    """rows：map_rows 之后的表（key, name, so, stock, psi_code, status）；name_excluded：被 exclude_name 丢掉的行。
    返回 [(级别, 类别, 物料号, 名称, SO, 库存, 说明)]，级别 ❌ / ⚠️ / ℹ️。"""
    out = []

    def add(level, kind, r, msg):
        out.append((level, kind, r["key"], r["name"], int(r["so"]), int(r["stock"]), msg))

    if name_excluded is not None:
        for _, r in name_excluded.iterrows():
            t = tail_of(r["name"])
            if is_colour(t) and not re.search(r"(?i)demo|sample|outl|restar", str(r["name"])):
                add("❌", "规则误伤", r, f"尾缀 {t} 是颜色缩写，不应该因为它被排除")
            elif t == "CR":
                add("ℹ️", "按规则排除", r, "特殊版本 CR")

    families = defaultdict(set)
    for _, r in rows.iterrows():
        fam = family_of(r["key"])
        if fam and r["status"] in ("psi", "others") and r["psi_code"]:
            families[fam].add(str(r["psi_code"]))

    for _, r in rows.iterrows():
        fam = family_of(r["key"])
        if not fam:
            continue
        t = tail_of(r["name"])
        if len(t) <= 4 and not re.search(r"\d", t) and not is_colour(t) and t.upper() not in KNOWN_TAGS:
            add("⚠️", "认不出的尾缀", r, f"尾缀 {t} 既不是颜色也不是已知版本，请确认计入方式")
        codes = families.get(fam, set())
        if r["status"] == "unmatched" and codes:
            add("⚠️", "同型号有代码", r, f"同型号其他物料是 {'、'.join(sorted(codes))}，确认后补进 mapping.xlsx")
        if r["status"] in ("psi", "others") and r["psi_code"]:
            code = str(r["psi_code"])
            if len(codes) > 1 and code.upper() != "OTHERS":
                add("⚠️", "同型号代码不一致", r, f"这一行是 {code}，同型号（T{fam}）一共有 {'、'.join(sorted(codes))}")
            nm = memory_from_name(r["name"])  # 代码是 项目号+内存 连写（P168+256），只能看结尾
            if CODE_MEMORY.search(code) and nm and not code.endswith(nm):
                add("⚠️", "内存对不上", r, f"名称里的内存是 {nm}，代码是 {code}")
    order = {"❌": 0, "⚠️": 1, "ℹ️": 2}
    return sorted(out, key=lambda x: (order[x[0]], x[1], x[2]))


def report_lines(findings, md_table):
    problems = [f for f in findings if f[0] != "ℹ️"]
    head = "✅ 没有发现问题" if not problems else f"⚠️ {len(problems)} 条要看一下"
    lines = ["", f"**TMO 防串检查**：{head}（说明见 使用说明.md）"]
    if findings:
        lines += ["", md_table(["", "类别", "物料号", "名称", "SO", "库存", "说明"],
                               [[a, b, c, str(d)[:44], e, f, g] for a, b, c, d, e, f, g in findings])]
    return lines


def main():
    """把所有历史 TMO 周报跑一遍，用现在的 mapping 和规则，看每周会报什么。"""
    import argparse
    argparse.ArgumentParser(description="TMO 防串检查：对 数据/历史周报 里所有 TMO 周报回测（只读，不改任何文件）").parse_args()
    from pathlib import Path

    from psi_tool.mapping import Mapper, latest_ttl
    from psi_tool.paths import CONFIG, HISTORY, MAPPING, PSI_FILE, TTL_DIR
    from psi_tool.run import load_psi, map_rows, psi_layout
    from psi_tool.sources import load_config, load_source
    import collections
    import shutil
    import tempfile
    from checks.structure_check import COL_G
    cfg = load_config(CONFIG)["TMO"]
    mapper = Mapper(latest_ttl(TTL_DIR), MAPPING)
    sheet = load_psi(PSI_FILE)["TMO"]
    codes = {c for c, *_ in psi_layout(sheet)[1:]}
    projects = collections.defaultdict(list)
    for r in sheet.product_rows("SO"):
        projects[str(sheet.get(r, COL_G)).strip()].append(sheet.code(r))
    seen = collections.Counter()
    files = sorted(Path(HISTORY).rglob("*kw_cpfr_main_Xiaomi*.xlsx"))
    for p in files:
        m = re.search(r"[Ww](\d+)", p.name) or re.search(r"/W(\d+)/", p.as_posix())
        if not m:
            continue
        week = int(m.group(1))
        with tempfile.TemporaryDirectory() as tmp:
            inbox = Path(tmp) / f"W{week}"
            inbox.mkdir()
            shutil.copy(p, inbox / "kw_cpfr_main_Xiaomi.xlsx")
            res = load_source(inbox, "TMO", cfg, week)
        if res.rows is None:
            print(f"W{week}: 读不到（{res.warnings}）")
            continue
        for f in check_tmo(map_rows(res.rows, mapper, codes, projects), res.name_excluded):
            seen[f[:4] + (f[6],)] += 1
    print(f"回测 {len(files)} 个文件")
    for (level, kind, key, name, msg), n in sorted(seen.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2])):
        print(f"{level} {kind} | {key} | {str(name)[:42]} | {msg} | 出现 {n} 周")


if __name__ == "__main__":
    main()
