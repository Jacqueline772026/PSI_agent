"""导出一份给 Excel 用的 PSI 码对照表：mapping/PSI码对照表.xlsx

每个运营商一页：识别码 → PSI 代码 + 是否计入。在原始周报旁边加一列 XLOOKUP 就能用，
和工具（psi_tool.run）用的是同一份映射，所以手工结果和工具结果可以直接比。

识别码收集自历史周报里出现过的所有型号；代码来自 mapping/mapping.xlsx，
只写项目代码的（如 P16）按 PSI 该 sheet 的行补成完整代码（P168+256）。
另外用产品名再独立匹配一次，和历史写法不一致的行标黄，请人工确认。

用法：.venv/bin/python -m tools.export_mapping
"""
import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from checks.structure_check import COL_E, COL_G, COL_H, load
from psi_tool.mapping import Mapper, latest_ttl, resolve
from psi_tool.paths import MAPPING, MAPPING_EXCEL, OPERATORS_HISTORY, PSI_FILE, TTL_DIR
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


def classify(source, name, code, category, sheet_codes):
    up = name.upper()
    if source == "TMO_MATERIAL" and "DEMO" in up:
        return "不计入：DEMO 样机"
    if source == "TMO_MATERIAL" and "SAMPLE" in up:
        return "不计入：SAMPLE 样机"
    if source == "TMO_MATERIAL" and "RESTART" in up:
        return "不计入：reSTART 翻新机"
    if source == "TMO_MATERIAL" and "OUTLET" in up:
        return "不计入：Outlet"
    if source == "ORANGE_NAME" and up.startswith("E."):
        return "不计入：门店展示机（E.）"
    if category == "exclude":
        return "不计入：非手机 / 翻新 / 样机"
    if code is None:
        return "不计入：非手机平板" if NOT_PHONE.search(up) else "⚠️ 没有代码（老型号或新品，请确认）"
    if code not in sheet_codes:
        return "不计入：PSI 这张 sheet 没有这一行"
    return "计入"


def build():
    psi = load(PSI_FILE)
    mapper = Mapper(latest_ttl(TTL_DIR), MAPPING)
    pages = {}
    for page, (source, sheet_name, file_rx, sheet, key_col, name_col, _, _) in SOURCES.items():
        part = {"TMO": "T-Mobile", "Play": "Play", "Plus": "Plus"}.get(page, "")
        codes, projects, names = psi_index(psi[sheet_name])
        rows = []
        for key, (name, week) in sorted(collect(file_rx, sheet, key_col, name_col, part).items()):
            lookup_key = key.upper() if source.endswith("_NAME") else key
            code, category, _ = mapper.lookup(source, lookup_key)
            code = resolve(code, name, codes, projects) if code else None
            status = classify(source, name, code, category, codes)
            # 用名称独立再配一次，和历史写法不一样就提示
            nk, mem = name_key(name)
            by_name = names.get((nk, mem), set())
            check = ""
            if status == "计入" and len(by_name) == 1 and code not in by_name:
                check = f"按名称应为 {next(iter(by_name))}，请确认"
            rows.append({"识别码": key, "名称（最近一次）": name, "PSI 代码": code if status == "计入" else "",
                         "是否计入": status, "历史写法": code or "", "最后出现": f"W{week}" if week else "",
                         "核对提示": check})
        pages[page] = pd.DataFrame(rows)
    return pages


GUIDE = [
    ("这份表是做什么的", "运营商周报（TMO / Play / Plus / Orange）原始文件里的识别码 → PSI 代码。"
     "在原始文件旁边加一列 XLOOKUP 取 'PSI 代码'，再透视求和，就是要填进 PSI 的数。和工具用的是同一份映射。"),
    ("'是否计入' 列", "只有'计入'的行有 PSI 代码。DEMO、Outlet、Orange 的 E. 展示机、非手机、PSI 没有这一行的型号都不计入，"
     "XLOOKUP 会返回空，透视时自然不算。"),
    ("黄色行", "历史上手工写的代码和产品名对不上（例如 Redmi 15C 5G 被写成 Redmi 15 5G 的 O19），已经按名称改正。请过一眼。"),
    ("TMO 公式（原始表 B 列 = Materiał，C 列 = Opis）",
     '=IF(OR(ISNUMBER(SEARCH("DEMO",C2)),ISNUMBER(SEARCH("SAMPLE",C2)),ISNUMBER(SEARCH("OUTLET",C2)),'
     'ISNUMBER(SEARCH("RESTART",C2))),"",XLOOKUP(B2,TMO!A:A,TMO!C:C,"#N/A"))'),
    ("TMO 只去 DEMO 的单独公式（想先单独看样机有多少时用）",
     '=IF(OR(ISNUMBER(SEARCH("DEMO",C2)),ISNUMBER(SEARCH("SAMPLE",C2))),"样机","")   '
     '← 筛选这一列 = 样机 的行就是要去掉的 DEMO / SAMPLE。这个公式不分大小写'),
    ("Play 公式（原始表 Arkusz1：C 列 = SKU）", '=XLOOKUP(C2,Play!A:A,Play!C:C,"#N/A")'),
    ("Plus 公式（原始表 B 列 = Model）", '=XLOOKUP(TRIM(B2),Plus!A:A,Plus!C:C,"#N/A")'),
    ("Orange 公式（原始表 B 列 = Nazwa materiału）",
     '=XLOOKUP(TRIM(SUBSTITUTE(B2,CHAR(160)," ")),Orange!A:A,Orange!C:C,"#N/A")   '
     '库存 = Zapas_CWH + Zapas_POS + Zapas_B2B'),
    ("Orange（Jet 表）公式（Main sheet A 列 = Model）", '=XLOOKUP(TRIM(A5),\'Orange（Jet表）\'!A:A,\'Orange（Jet表）\'!C:C,"")'),
    ("返回 #N/A 怎么办", "说明是表里没有的新型号：手机平板就在 mapping/mapping.xlsx 补一行（工具下次自动生效），"
     "再重新导出这份表；不是手机平板就忽略。返回空 = 不计入。"),
    ("和工具对比", "工具的结果在 W{周}/工具结果/W{周}_paste.xlsx，同一张 sheet、同一行顺序。两边透视结果按代码对一下即可。"),
]


def write(pages, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame(GUIDE, columns=["", "说明"]).to_excel(xw, sheet_name="说明", index=False)
        for page, df in pages.items():  # XLOOKUP 不分大小写，识别码保持原样
            df.to_excel(xw, sheet_name=page, index=False)
    wb = load_workbook(path)
    yellow, grey = PatternFill("solid", fgColor="FFF2CC"), Font(color="999999")
    for ws in wb.worksheets:
        ws.freeze_panes = "B2"
        for cell in ws[1]:
            cell.font = Font(bold=True)
        if ws.title == "说明":
            ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 42, 120
            for row in ws.iter_rows(min_row=2):
                row[1].alignment = Alignment(wrap_text=True, vertical="top")
            continue
        for col, w in zip("ABCDEFG", (26, 46, 16, 30, 16, 10, 30)):
            ws.column_dimensions[col].width = w
        for row in ws.iter_rows(min_row=2):
            if row[6].value:
                for c in row:
                    c.fill = yellow
            elif row[3].value != "计入":
                for c in row:
                    c.font = grey
    wb.save(path)


def main():
    pages = build()
    out = MAPPING_EXCEL
    write(pages, out)
    for page, df in pages.items():
        print(f"{page}: {len(df)} 个识别码，计入 {int((df['是否计入'] == '计入').sum())}，"
              f"待确认 {int(df['是否计入'].str.startswith('⚠️').sum())}，名称核对提示 {int((df['核对提示'] != '').sum())}")
    print(out)


if __name__ == "__main__":
    main()
