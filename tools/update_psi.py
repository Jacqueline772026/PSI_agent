"""在上一周的 PSI 导出上，生成"工具版"的新一周 PSI，并导出 DB 可粘贴表。

  1. DB：把当周 2026 DB ST / STOCK 的行追加进去（公式照同类最后一行往下写，见 tools.write_db_to_psi）
  2. OM / 运营商各页：把当周 SO 和 Channel Inventory Actual 写进对应格子（只写数字，格式不动）
     - 这一页的周报没收齐（比如 EURO 缺平板文件）或缺某个指标（比如缺 Orange 的 SO），就整页 / 整个指标跳过，不写 0
  3. 另存一份 DB 可粘贴表：列位置和 PSI 完全一致，公式列空着，贴完从上一行拉公式

文件里其他部分（单元格图片、批注、其他 sheet）都原样保留，不会覆盖原文件。
等你从飞书下载手工做好的这一周 PSI，用 --compare 或 tools.compare_psi 就能逐格对比工具版和手工版。

用法：
  .venv/bin/python -m tools.update_psi --week 45 --inbox "W45/W45原始文件" --psi 上周PSI.xlsx --out "W45/PSI_W45_工具版.xlsx"
      [--extra 补录表.xlsx ...]   另外要追加进 DB 的行（如 POCO 补录）
"""
import argparse
import re
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl.utils import column_index_from_string, get_column_letter

from psi_tool.db import build_db
from psi_tool.paths import CONFIG, MAPPING, TTL_DIR
from psi_tool.run import DB_ORDER, build, summarize
from tools.write_db_to_psi import SHEETS, append, build_row, load_rows, sheet_paths, template_cells, templates

DB_LAYOUT = {  # PSI 第 2 行的表头；空字符串 = 公式列
    "2026 DB ST": ["WEEK", "DB", "PRODUCT NAME", "", "", "", "", "", "KA Name", "ST qty", "AB reference",
                   "IM&ALSO reference", "", "", "", "ST time"],
    "2026 DB STOCK": ["WEEK", "DB", "Product", "", "", "", "", "", "STOCK", "reference"],
}


def db_frames(week, inbox, psi):
    st, stock, notes = build_db(week, inbox, psi)
    out = {}
    for name, df in (("2026 DB ST", st), ("2026 DB STOCK", stock)):
        if len(df):
            df = df.sort_values("B", key=lambda b: b.map(DB_ORDER), kind="stable")
            out[name] = df[[c for c in SHEETS[name] if c in df.columns]].reset_index(drop=True)
    return out, notes


def write_pastable(frames, path, start_rows):
    with pd.ExcelWriter(path) as xw:
        info = []
        for name, df in frames.items():
            cols = DB_LAYOUT[name]
            grid = pd.DataFrame({get_column_letter(i + 1): [None] * len(df) for i in range(len(cols))})
            for c in df.columns:
                if c in grid:
                    grid[c] = df[c].values
            grid.columns = [h or f"（{get_column_letter(i + 1)} 列：拉公式）" for i, h in enumerate(cols)]
            grid.to_excel(xw, sheet_name=name, index=False)
            info.append(f"{name}：{len(df)} 行。贴到 PSI 这一页的第一个空行（按上周导出算是第 {start_rows[name]} 行；"
                        "如果之后又贴过别的行，以实际第一个空行为准），A 列对 A 列整块贴")
        pd.DataFrame({"说明": info + [
            "列位置和 PSI 完全一致：选中数据区（不含表头）整块贴到 PSI 的 A 列即可。",
            "标着'拉公式'的列是空的：贴完后选中上一行这些列的公式，往下拉到最后一行。",
            "2026 DB STOCK 另外还有 L、M、Q 三列公式（Value、Total value、Spec），也要从上一行拉下来。",
            "B 列按经销商写 AB / ALSO / IM（POCO 也按经销商写）。",
        ]}).to_excel(xw, sheet_name="说明", index=False)


def set_cell(xml, ref, value):
    """把单元格换成数字，保留样式；单元格不存在就按列顺序插进这一行。"""
    m = re.search(rf'<c r="{ref}"([^>]*?)(?:/>|>.*?</c>)', xml, re.S)
    if m:
        s = re.search(r'\bs="(\d+)"', m.group(1))
        new = f'<c r="{ref}"{f" s=\"{s.group(1)}\"" if s else ""}><v>{value}</v></c>'
        return xml[:m.start()] + new + xml[m.end():], True
    col, row = re.match(r"([A-Z]+)(\d+)", ref).groups()
    rm = re.search(rf'<row\b[^>]*?\br="{row}"[^>]*?(/>|>(.*?)</row>)', xml, re.S)
    new = f'<c r="{ref}"><v>{value}</v></c>'
    if rm is None:
        return xml, False
    if rm.group(1) == "/>":
        open_tag = rm.group(0)[:-2] + ">"
        return xml[:rm.start()] + open_tag + new + "</row>" + xml[rm.end():], True
    body_start = rm.start(2)
    for c in re.finditer(r'<c r="([A-Z]+)\d+"', rm.group(2)):
        if column_index_from_string(c.group(1)) > column_index_from_string(col):
            pos = body_start + c.start()
            return xml[:pos] + new + xml[pos:], True
    pos = rm.end(2)
    return xml[:pos] + new + xml[pos:], True


def ka_edits(week, inbox, psi):
    """{sheet: [(单元格, 数)]}，以及跳过的说明。"""
    results, _ = build(week, inbox, psi, CONFIG, MAPPING, TTL_DIR)
    edits, skipped = {}, []
    for sheet, res in results.items():
        t = summarize(res)
        missing = [p.name for p in res["parts"] if not p.files]
        if t is None or missing:
            skipped.append(f"{sheet}：缺 {'、'.join(missing) or '全部文件'}，整页不写")
            continue
        col = get_column_letter(res["sheet"].week_col(week) + 1)  # get() 的列号比 Excel 少 1
        cells = []
        for _, x in t.iterrows():
            if "so" in res["fields"] and x.so_row:
                cells.append((f"{col}{int(x.so_row)}", int(x.so)))
            if "stock" in res["fields"] and x.actual_row:
                cells.append((f"{col}{int(x.actual_row)}", int(x.stock)))
        for f in ("so", "stock"):
            if f not in res["fields"]:
                skipped.append(f"{sheet}：周报里没有 {'SO' if f == 'so' else '库存'}，这一列不写")
        edits[sheet] = cells
    return edits, skipped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--week", type=int, required=True)
    ap.add_argument("--inbox", required=True)
    ap.add_argument("--psi", required=True, help="上一周（已经填好）的 PSI 导出")
    ap.add_argument("--out", required=True)
    ap.add_argument("--extra", nargs="*", default=[], help="另外要追加进 DB 的表（如 POCO 补录）")
    ap.add_argument("--only-ka", action="store_true", help="只写 OM / 运营商各页，不动 DB（验证用）")
    a = ap.parse_args()
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    frames = {} if a.only_ka else db_frames(a.week, a.inbox, a.psi)[0]
    extra = load_rows(a.extra) if a.extra and not a.only_ka else {}
    zin = zipfile.ZipFile(a.psi)
    paths = sheet_paths(zin)
    changed, starts, lines = {}, {}, []
    for name in SHEETS:
        parts = [df for df in (extra.get(name), frames.get(name)) if df is not None and len(df)]
        if not parts:
            continue
        df = pd.concat(parts, ignore_index=True)
        last, per_db = templates(a.psi, name)
        xml = zin.read(paths[name]).decode("utf-8")
        new, r = {}, last
        for _, row in df.iterrows():
            tpl = per_db.get(row["B"]) or per_db.get("IM")
            r += 1
            new[r] = build_row(r, {k: (None if pd.isna(v) else v) for k, v in row.items()}, tpl, template_cells(xml, tpl["_row"]))
        changed[paths[name]] = append(xml, new)
        n_extra = len(extra[name]) if extra.get(name) is not None else 0
        starts[name] = last + 1 + n_extra  # 可粘贴表只放当周的行，从补录行之后开始
        lines.append(f"{name}：追加 {len(df)} 行（第 {last + 1}–{r} 行）")

    edits, skipped = ka_edits(a.week, a.inbox, a.psi)
    for sheet, cells in edits.items():
        p = paths[sheet]
        xml = changed.get(p) or zin.read(p).decode("utf-8")
        ok = 0
        for ref, v in cells:
            xml, done = set_cell(xml, ref, v)
            ok += done
        changed[p] = xml
        lines.append(f"{sheet}：写入 W{a.week} {ok} 格")

    with zipfile.ZipFile(out, "w") as zout:
        for info in zin.infolist():
            data = changed.get(info.filename)
            zout.writestr(info, data.encode("utf-8") if data is not None else zin.read(info.filename),
                          compress_type=info.compress_type)
    if frames:
        pastable = out.parent / f"W{a.week}_DB可粘贴.xlsx"
        write_pastable(frames, pastable, starts)
        lines.append(f"DB 可粘贴表：{pastable}")
    print("\n".join(lines + [f"跳过：{s}" for s in skipped] + [str(out)]))


if __name__ == "__main__":
    main()
