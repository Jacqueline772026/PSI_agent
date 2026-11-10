"""把 DB 的待粘贴行直接写进 PSI 导出的 xlsx（追加到 2026 DB ST / 2026 DB STOCK 最后）。

为什么不用 openpyxl 直接改：飞书导出的 PSI 里有单元格图片、线程批注（@同事的讨论）、动态数组公式的元数据，
openpyxl 读进来再保存会把这些丢掉。这里只在两页的 XML 里追加 <row>，文件里其他部分一个字节都不动。

新行的公式照同一页最后一行同类（AB 按物料码 K 列查；ALSO / IM 按 69 码 L 列查）的公式写，只换行号；
值写成内联字符串，不改 sharedStrings。不会覆盖原文件，另存一份。

用法：
  .venv/bin/python -m tools.write_db_to_psi --psi PSI.xlsx --out 新文件.xlsx 待粘贴表1.xlsx [待粘贴表2.xlsx ...]
待粘贴表：psi_tool.run 的 W{周}_paste.xlsx（页名 2026 DB ST / 2026 DB STOCK），或 tools.poco_backfill 的补录表（页名带"补录"）。
"""
import argparse
import datetime as dt
import re
import zipfile
from html import escape

import openpyxl
import pandas as pd

SHEETS = {"2026 DB ST": {"A": "A 周", "B": "B DB", "C": "C 产品名", "I": "I KA Name", "J": "J ST qty",
                         "K": "K AB reference", "L": "L IM&ALSO reference", "P": "P ST time"},
          "2026 DB STOCK": {"A": "A 周", "B": "B DB", "C": "C 产品名", "I": "I STOCK", "J": "J reference"}}
NUMERIC = {"A", "J", "L"}  # 数字列（ST 的 J = 数量、L = 69 码；STOCK 的 J = 69 码、I = 库存）
ROW_RX = re.compile(r'<row\b[^>]*?\br="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
CELL_RX = re.compile(r'<c\b[^>]*?\br="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', re.S)


def sheet_paths(z):
    wb = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    ids = {}
    for rel in re.findall(r"<Relationship\b[^>]*>", rels):
        i, t = re.search(r'Id="([^"]+)"', rel).group(1), re.search(r'Target="([^"]+)"', rel).group(1)
        ids[i] = "xl/" + t.lstrip("/").removeprefix("xl/")
    out = {}
    for s in re.findall(r"<sheet\b[^>]*>", wb):
        name = re.search(r'name="([^"]+)"', s).group(1)
        out[name.replace("&amp;", "&")] = ids[re.search(r'r:id="([^"]+)"', s).group(1)]
    return out


def load_rows(paste_files):
    out = {name: [] for name in SHEETS}
    for f in paste_files:
        for sheet in pd.ExcelFile(f).sheet_names:
            for name, cols in SHEETS.items():
                if sheet == name or sheet == f"{name} 补录":
                    df = pd.read_excel(f, sheet_name=sheet, dtype=object)
                    df = df[[c for c in cols.values() if c in df.columns]].rename(columns={v: k for k, v in cols.items()})
                    out[name].append(df)
    return {k: pd.concat(v, ignore_index=True) if v else pd.DataFrame() for k, v in out.items()}


def templates(path, name):
    """每种 DB 最后一行的行号和公式（openpyxl 会把共享公式展开成完整公式）。"""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[name]
    ws.reset_dimensions()
    last, per_db = 0, {}
    for row in ws.iter_rows(min_row=3):
        a, b = row[0].value, row[1].value
        if isinstance(a, (int, float)) and b:
            last = row[0].row
            per_db[b] = {c.column_letter: getattr(c.value, "text", c.value) for c in row
                         if isinstance(getattr(c.value, "text", c.value), str) and str(getattr(c.value, "text", c.value)).startswith("=")}
            per_db[b]["_row"] = row[0].row
    return last, per_db


def shift(formula, old, new):
    return re.sub(rf"(?<![A-Za-z_$])([A-Z]{{1,3}}){old}(?![0-9])", rf"\g<1>{new}", formula)


def cell_xml(col, r, value, style, numeric):
    s = f' s="{style}"' if style else ""
    if value is None or (isinstance(value, float) and pd.isna(value)) or value == "":
        return ""
    if isinstance(value, (dt.date, dt.datetime, pd.Timestamp)):
        value = (pd.Timestamp(value).normalize() - pd.Timestamp("1899-12-30")).days
        numeric = True
    if numeric:
        try:
            v = float(value)
            return f'<c r="{col}{r}"{s}><v>{int(v) if v.is_integer() else v}</v></c>'
        except (TypeError, ValueError):
            pass
    return f'<c r="{col}{r}"{s} t="inlineStr"><is><t>{escape(str(value), quote=False)}</t></is></c>'


def build_row(r, data, tpl_formulas, tpl_cells):
    """tpl_cells：模板行 XML 里每列的属性（样式、是否数组公式）。"""
    cells = []
    cols = sorted(set(tpl_cells) | set(data), key=lambda c: (len(c), c))
    for col in cols:
        attrs = tpl_cells.get(col, {})
        if col in tpl_formulas:
            f = shift(tpl_formulas[col], tpl_formulas["_row"], r).lstrip("=")
            s = f' s="{attrs["s"]}"' if attrs.get("s") else ""
            if attrs.get("array"):
                cells.append(f'<c r="{col}{r}"{s} cm="1"><f t="array" ref="{col}{r}:{col}{r}">{escape(f, quote=False)}</f></c>')
            else:
                cells.append(f'<c r="{col}{r}"{s}><f>{escape(f, quote=False)}</f></c>')
        elif col in data:
            cells.append(cell_xml(col, r, data[col], attrs.get("s"), col in NUMERIC or col == "I" and "K" not in data and "P" not in data))
    return f'<row r="{r}">' + "".join(cells) + "</row>"


def template_cells(xml, row):
    m = re.search(rf'<row\b[^>]*?\br="{row}"[^>]*?>.*?</row>', xml, re.S)
    out = {}
    for c in re.finditer(r'<c\b([^>]*?)(?:/>|>(.*?)</c>)', m.group(0), re.S):
        attr = c.group(1)
        col = re.search(r'r="([A-Z]+)\d+"', attr).group(1)
        s = re.search(r'\bs="(\d+)"', attr)
        out[col] = {"s": s.group(1) if s else None, "array": 't="array"' in (c.group(2) or "")}
    return out


def append(xml, new_rows):
    """把新行放进 sheetData：行号已存在（空的带格式行）就替换，否则按顺序插入。"""
    start, end = xml.index("<sheetData"), xml.index("</sheetData>")
    head = xml[:xml.index(">", start) + 1]
    body = xml[len(head):end]
    rows = {int(m.group(1)): m.group(0) for m in ROW_RX.finditer(body)}
    rows.update(new_rows)
    return head + "".join(rows[k] for k in sorted(rows)) + xml[end:]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--psi", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("paste", nargs="+")
    a = ap.parse_args()
    data = load_rows(a.paste)
    zin = zipfile.ZipFile(a.psi)
    paths = sheet_paths(zin)
    changed, summary = {}, []
    for name in SHEETS:
        df = data[name]
        if not len(df):
            continue
        last, per_db = templates(a.psi, name)
        xml = zin.read(paths[name]).decode("utf-8")
        new, r = {}, last
        for _, row in df.iterrows():
            db = row["B"]
            tpl = per_db.get(db) or per_db.get("IM")
            r += 1
            values = {k: (None if pd.isna(v) else v) for k, v in row.items()}
            new[r] = build_row(r, values, tpl, template_cells(xml, tpl["_row"]))
        changed[paths[name]] = append(xml, new).encode("utf-8")
        summary.append((name, last + 1, r, len(df)))
    with zipfile.ZipFile(a.out, "w") as zout:
        for info in zin.infolist():
            zout.writestr(info, changed.get(info.filename, zin.read(info.filename)), compress_type=info.compress_type)
    for name, first, lastr, n in summary:
        print(f"{name}：追加 {n} 行（第 {first}–{lastr} 行）")
    print(a.out)


if __name__ == "__main__":
    main()
