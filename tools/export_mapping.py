"""导出一份给 Excel 用的 PSI 码对照表：mapping/PSI码对照表.xlsx

每个运营商一页：识别码 → PSI 代码 + 是否计入。在原始周报旁边加一列 XLOOKUP 就能用，
和工具（psi_tool.run）用的是同一份映射，所以手工结果和工具结果可以直接比。

识别码收集自历史周报里出现过的所有型号；代码来自 mapping/mapping.xlsx，
只写项目代码的（如 P16）按 PSI 该 sheet 的行补成完整代码（P168+256）。
每一行都按原始名称、TTL 和 PSI 独立核对过（tools.audit_mapping），不依赖上周写法；结论写在“核对结论”“依据”两列。

用法：.venv/bin/python -m tools.export_mapping
"""
import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from checks.structure_check import COL_E, COL_G, COL_H, load
from psi_tool.mapping import CODE_ALIASES, Mapper, latest_ttl, legacy_code, resolve
from psi_tool.paths import HISTORY, MAPPING, MAPPING_EXCEL, OPERATORS_HISTORY, PSI_FILE, ROOT, TTL_DIR
from tools.audit_mapping import audit_entry, run_audit
from tools.learn_mappings import name_key

REPORTS = OPERATORS_HISTORY
NOT_PHONE = re.compile(r"SŁUCH|BUDS|WATCH|BAND|SCOOTER|HULAJ|TV|VACUUM|ODKURZ|ROBOT|AIR |ADAPTER|ŁADOW|POWER|"
                       r"KAMERA|CAMERA|LAMP|TAG |GŁOŚN|SPEAKER|ROUTER|ETUI|CASE|SZKŁO", re.I)

SOURCES = {
    # 页名: (mapping source, PSI sheet, 文件匹配, sheet 名, 识别码列, 名称列, 原始文件里识别码在哪列, 名称在哪列)
    "TMO": ("TMO_MATERIAL", "TMO", r"kw_cpfr_main_Xiaomi\.xlsx$", "kw_cpfr_main_Xiaomi", "Materiał", "Opis", "B", "C"),
    "Play": ("PLAY_SKU", "Play", r"Stock and sales - week \d+\.xlsx$", None, "SKU", "Model", "C", "D"),
    "Plus": ("PLUS_NAME", "Plus", r"Xiaomi.*\.xlsx$", None, "Model", "Model", "B", "B"),
    "Orange": ("ORANGE_NAME", "Orange", r"Zapas_oraz_dystrybucja_XIAOMI_W\d+\.xlsx$", None, "Nazwa materiału",
               "Nazwa materiału", "B", "B"),
    "Orange（Jet表）": ("ORANGE_JET_NAME", "Orange", r"OPL_Stock_Sales.*\.xlsx$", "Main", "Model", "Model", "A", "A"),
}


def week_of(path):
    m = re.search(r"W(?:eek )?\s*0?(\d{1,2})\b", path.name, re.I) or re.search(r"week (\d+)", path.name, re.I)
    return int(m.group(1)) if m else 0


def collect(file_rx, sheet, key_col, name_col, part):
    """历史文件里出现过的 (识别码, 名称, 最后出现的周)。"""
    seen = {}
    for p in REPORTS.rglob("*.xlsx"):
        if not re.search(file_rx, p.name) or part not in p.as_posix():
            continue
        try:
            frames = pd.read_excel(p, sheet_name=sheet if sheet else None, dtype=str,
                                   header=None if sheet == "Main" else 0)
        except ValueError:
            continue
        frames = frames if isinstance(frames, dict) else {sheet: frames}
        for df in frames.values():
            if sheet == "Main":  # Jet 的表：第一列就是型号
                df = pd.DataFrame({key_col: df.iloc[:, 0]})
            if key_col not in df.columns:
                continue
            names = df[name_col] if name_col in df.columns else df[key_col]
            for k, n in zip(df[key_col], names):
                if pd.isna(k) or not str(k).strip() or str(k).strip() == key_col:
                    continue
                k = re.sub(r"\s+", " ", str(k).replace("\xa0", " ").strip())
                w = week_of(p)
                if k not in seen or w >= seen[k][1]:
                    seen[k] = (str(n).strip(), w)
    return seen


def psi_index(sheet):
    codes, projects, names = set(), {}, {}
    for r in sheet.product_rows("SO"):
        c = sheet.code(r)
        codes.add(c)
        projects.setdefault(str(sheet.get(r, COL_G)).strip(), []).append(c)
        e = sheet.get(r, COL_E)
        if isinstance(e, str) and not e.startswith("="):
            names.setdefault((name_key(e)[0], str(sheet.get(r, COL_H)).strip()), set()).add(c)
    return codes, projects, names


def classify(source, name, code, category, sheet_codes, key=None):
    up = name.upper()
    kind = re.match(r"^XA-[A-Z]-([DSR])", key or "") if source == "TMO_MATERIAL" else None
    if kind:  # 物料号第 6 位：名称被截断（Outl、reSTAR）时也认得出
        return {"D": "不计入：DEMO / EXPO 展示机（物料号 D）", "S": "不计入：SAMPLE 样机（物料号 S）",
                "R": "不计入：Outlet / reSTART（物料号 R）"}[kind.group(1)]
    if source == "TMO_MATERIAL" and "DEMO" in up:
        return "不计入：DEMO 样机"
    if source == "TMO_MATERIAL" and "SAMPLE" in up:
        return "不计入：SAMPLE 样机"
    if source == "TMO_MATERIAL" and "RESTART" in up:
        return "不计入：reSTART 翻新机"
    if source == "TMO_MATERIAL" and "OUTLET" in up:
        return "不计入：Outlet"
    if source == "TMO_MATERIAL" and re.search(r"\bCR\s*$", name):
        return "不计入：特殊版本 CR"
    if source == "ORANGE_NAME" and up.startswith("E."):
        return "不计入：门店展示机（E.）"
    if category == "exclude":
        return "不计入：非手机 / 翻新 / 样机"
    if code is None:
        return "不计入：非手机平板" if NOT_PHONE.search(up) else "⚠️ 没有代码（老型号或新品，请确认）"
    if code not in sheet_codes:
        return "不计入：PSI 这张 sheet 没有这一行"
    return "计入"


def audit_weeks():
    """把最近几周的原始周报也纳入核对，映射表里没有的新型号才查得到。"""
    weeks = {42: HISTORY / "W42-43" / "W42", 43: HISTORY / "W42-43" / "W43", 44: ROOT / "W44测试" / "W44原始文件"}
    for path in ROOT.glob("W[0-9][0-9]/W*原始文件"):
        weeks[int(path.parent.name[1:])] = path
    return {w: p for w, p in weeks.items() if p.exists()}


def build(df, ev):
    psi = load(PSI_FILE)
    mapper = Mapper(latest_ttl(TTL_DIR), MAPPING)
    pages = {}
    for page, (source, sheet_name, file_rx, sheet, key_col, name_col, _, _) in SOURCES.items():
        part = {"TMO": "T-Mobile", "Play": "Play", "Plus": "Plus"}.get(page, "")
        codes, projects, names = psi_index(psi[sheet_name])
        rows = []
        seen = collect(file_rx, sheet, key_col, name_col, part)
        known = {k.upper() for k in seen}
        recent = df[df["来源"] == source]  # 历史文件夹之外的周报、映射表里的识别码也要列出来
        for _, r in recent.iterrows():
            if str(r["识别码"]).upper() not in known:
                seen[r["识别码"]] = (r["原始名称"], int(r["最后一周"][1:]) if r["最后一周"] else 0)
        verdicts = {(str(r["识别码"]).upper()): r for _, r in recent.iterrows()}
        for key, (name, week) in sorted(seen.items()):
            lookup_key = key.upper() if source.endswith("_NAME") else key
            code, category, _ = mapper.lookup(source, lookup_key)
            code = resolve(code, name, codes, projects) if code else None
            status = classify(source, name, code, category, codes, key)
            audited = verdicts.get(key.upper())
            if audited is None:
                audited = audit_entry(ev, source, lookup_key, name, code or "", category or "")
            if status.startswith("⚠️") and audited["结论"] == "一致":
                status = "不计入：PSI 没有这个型号这一行"
            # 用名称独立再配一次，和历史写法不一样就提示
            nk, mem = name_key(name)
            by_name = names.get((nk, mem), set())
            check = ""
            if status == "计入" and len(by_name) == 1 and code not in by_name:
                check = f"按名称应为 {next(iter(by_name))}，请确认"
            rows.append({"识别码": key, "名称（最近一次）": name, "PSI 代码": code if status == "计入" else "",
                         "是否计入": status, "历史写法": (code or "") if status == "计入" else (f"（曾写成 {code}，不要用）" if code else ""), "最后出现": f"W{week}" if week else "",
                         "核对提示": check, "核对结论": audited["结论"], "依据": audited["依据"],
                         "旧代码": legacy_code(code) if status == "计入" else ""})
        pages[page] = pd.DataFrame(rows)
    return pages


GUIDE = [
    ("这份表是做什么的", "运营商周报（TMO / Play / Plus / Orange）原始文件里的识别码 → PSI 代码。"
     "在原始文件旁边加一列 XLOOKUP 取 'PSI 代码'，再透视求和，就是要填进 PSI 的数。和工具用的是同一份映射。"),
    ("'是否计入' 列", "只有'计入'的行有 PSI 代码。DEMO、Outlet、Orange 的 E. 展示机、非手机、PSI 没有这一行的型号都不计入，"
     "XLOOKUP 会返回空，透视时自然不算。"),
    ("以后怎么用", "每周直接按这份表匹配，不要沿用上周文件的 CODE 列：上周写错的会一起带过来。"
     "这份表的代码按原始名称 + TTL + PSI 独立核对过，“核对结论”“依据”两列说明了每行怎么来的。"),
    ("黄色、红色行", "黄色：名称和历史写法不一致或无法独立验证，请过一眼。红色：核对发现冲突，查清前不要用。"),
    ("颜色尾缀", "TMO 名称被截断时 bi、cz、fi、nie 这些都是颜色（biały、czarny、fioletowy、niebieski），不影响代码，也不排除。"
     "想确认颜色时看 TTL ID 表的 Color 列。"),
    ("代码别名", "Q16 = MOSCOW，Q16U = BRUSSELS，Q16P = MOSCOWP。PSI 和“PSI 代码”列都用新代码，最后一列“旧代码”对应 BOC、促销表的写法，在表里搜哪个都能找到。填 PSI 用新代码。"),
    ("其他 sheet", "“待你确认”：拿不准的。“TTL表的问题”“PSI名称的问题”：源表本身的错误。核对方法见 docs/对照表核对方法.md。"),
    ("TMO 公式：原始表 B 列 = Materiał，C 列 = Opis。物料号第 6 位是 D / S / R（DEMO、EXPO、SAMPLE、Outlet、reSTART）"
     "的直接不计入，名称被截断成 Outl、reSTAR 也认得出；名称以大写 CR 结尾的特殊版本也不计入",
     '=IF(OR(ISNUMBER(SEARCH(MID(B2,6,1),"DSR")),ISNUMBER(SEARCH("DEMO",C2)),ISNUMBER(SEARCH("SAMPLE",C2)),'
     'ISNUMBER(SEARCH("OUTL",C2)),ISNUMBER(SEARCH("RESTA",C2)),EXACT(RIGHT(TRIM(C2),3)," CR")),"",'
     'XLOOKUP(B2,TMO!A:A,TMO!C:C,"#N/A"))'),
    ("TMO 单独标出不计入的行（想先看样机、Outlet 有多少时用；筛选这一列 = 不计入）",
     '=IF(ISNUMBER(SEARCH(MID(B2,6,1),"DSR")),"不计入","")'),
    ("Play 公式（原始表 Arkusz1：C 列 = SKU）", '=XLOOKUP(C2,Play!A:A,Play!C:C,"#N/A")'),
    ("Plus 公式（原始表 B 列 = Model）", '=XLOOKUP(TRIM(B2),Plus!A:A,Plus!C:C,"#N/A")'),
    ("Orange 公式（原始表 B 列 = Nazwa materiału；库存 = Zapas_CWH + Zapas_POS + Zapas_B2B）",
     '=XLOOKUP(TRIM(SUBSTITUTE(B2,CHAR(160)," ")),Orange!A:A,Orange!C:C,"#N/A")'),
    ("Orange（Jet 表）公式（Main sheet A 列 = Model）", '=XLOOKUP(TRIM(A5),\'Orange（Jet表）\'!A:A,\'Orange（Jet表）\'!C:C,"")'),
    ("返回 #N/A 怎么办", "说明是表里没有的新型号：手机平板就在 mapping/mapping.xlsx 补一行（工具下次自动生效），"
     "再重新导出这份表；不是手机平板就忽略。返回空 = 不计入。"),
    ("和工具对比", "工具的结果在 W{周}/工具结果/W{周}_paste.xlsx，同一张 sheet、同一行顺序。两边透视结果按代码对一下即可。"),
]


def write(pages, path, extra):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame(GUIDE, columns=["", "说明"]).to_excel(xw, sheet_name="说明", index=False)
        for page, df in pages.items():  # XLOOKUP 不分大小写，识别码保持原样
            df.to_excel(xw, sheet_name=page, index=False)
        for name, df in extra.items():
            df.to_excel(xw, sheet_name=name, index=False)
    wb = load_workbook(path)
    yellow, grey = PatternFill("solid", fgColor="FFF2CC"), Font(color="999999")
    red = PatternFill("solid", fgColor="F8CBAD")
    for ws in wb.worksheets:
        ws.freeze_panes = "B2"
        for cell in ws[1]:
            cell.font = Font(bold=True)
        if ws.title in extra:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(70, max(10, max(len(str(c.value or "")) for c in col) + 2))
            continue
        if ws.title == "说明":
            ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 42, 120
            for row in ws.iter_rows(min_row=2):
                row[1].alignment = Alignment(wrap_text=True, vertical="top")
                if isinstance(row[1].value, str) and row[1].value.startswith("="):
                    row[1].data_type = "s"  # 公式要以文字显示，方便复制；否则 Excel 会去算它，格子看起来是空的
            continue
        for col, w in zip("ABCDEFGHIJ", (26, 46, 16, 30, 16, 10, 30, 10, 60, 14)):
            ws.column_dimensions[col].width = w
        for row in ws.iter_rows(min_row=2):
            if row[7].value == "冲突":
                for c in row:
                    c.fill = red
            elif row[6].value or row[7].value == "无法验证":
                for c in row:
                    c.fill = yellow
            elif row[3].value != "计入":
                for c in row:
                    c.font = grey
    wb.save(path)


def main():
    import argparse
    argparse.ArgumentParser(description="按 mapping.xlsx 重新导出给 Excel 用的 mapping/PSI码对照表.xlsx").parse_args()
    df, ttl, labels, ev = run_audit(PSI_FILE, MAPPING, audit_weeks())
    pages = build(df, ev)
    pending = df[df["结论"].isin(["冲突", "漏映射", "排除待查", "无法验证"])
                 | (df["结论"].eq("人工确认") & df["依据"].str.contains("再看一眼") & df["来历"].str.contains("手工匹配"))]
    pending = pending[["结论", "PSI表", "识别码", "原始名称", "映射表代码", "建议代码", "依据", "最后一周", "周报SO", "周报库存"]]
    aliases = pd.DataFrame([dict(旧代码=old, PSI代码=new) for old, new in CODE_ALIASES.items()])
    aliases["产品"] = aliases["PSI代码"].map({"BRUSSELS": "Redmi Note 17 Pro Max 5G", "MOSCOW": "Redmi Note 17 Pro 5G", "MOSCOWP": "POCO X8"})
    aliases["例子"] = aliases["PSI代码"].map({"BRUSSELS": "Q16U8+256 → BRUSSELS8+256", "MOSCOW": "Q166+256 → MOSCOW6+256；Q168+512 → MOSCOW8+512", "MOSCOWP": "Q16P8+256 → MOSCOWP8+256"})
    aliases["在哪里见到旧代码"] = "BOC、促销表、早期手工表"
    extra = {"代码别名": aliases, "待你确认": pending, "TTL表的问题": ttl, "PSI名称的问题": labels}
    out = MAPPING_EXCEL
    write(pages, out, extra)
    for page, df in pages.items():
        print(f"{page}: {len(df)} 个识别码，计入 {int((df['是否计入'] == '计入').sum())}，"
              f"待确认 {int(df['是否计入'].str.startswith('⚠️').sum())}，名称核对提示 {int((df['核对提示'] != '').sum())}")
    print(out)


if __name__ == "__main__":
    main()
