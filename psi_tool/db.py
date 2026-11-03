"""DB（AB / ALSO / IM）：原始周报 → 往 PSI「2026 DB ST」「2026 DB STOCK」追加的行。

PSI 里 D–H、M–O 列是公式（按 69 码或 AB 物料码去 工具表1 查代码），所以这里只产出要粘贴的列：
  DB ST    ：A 周 / B 经销商 / C 产品名 / I 客户名 / J 数量 / K AB 物料码 / L ALSO·IM 的 EAN / P 日期
  DB STOCK ：A 周 / B 经销商 / C 产品名 / I 库存 / J EAN
另外用 工具表1 预先查一遍识别码：查不到的行，PSI 公式会显示 0，报告里会提前列出来。
"""
import datetime as dt
import re
from pathlib import Path

import openpyxl
import pandas as pd

from psi_tool.sources import candidates, week_range

SKU_TABLE = "工具表1_SKUID对应表 - GDP 平台导出"
# 配件、生态链产品：工具表1 查不到且名称命中这些词的行不贴（SOP：手环、手机壳、平板笔之类直接删掉）
ACCESSORY = re.compile(r"cover|keyboard|klawiatur|\bpen\b|etui|case|filter|filtr|brush|strap|buds|watch|band|lamp|vacuum|"
                       r"camera|kamera|charg|ładow|adapter|stand|cable|kabel|purifier|scooter|tv\b|speaker|router", re.I)

# 每家的文件：文件名正则（{week} 会替换）+ 用来核对周的日期规则
FILES = {
    "AB_ST": {"file": r"xiaomi_smart_sell_through_(\d{8})_(\d{8})\.xlsx", "dates": "start_end"},
    "AB_STOCK": {"file": r"xiaomi_smart_stock_(\d{8})\.xlsx", "dates": "end"},
    "ALSO": {"file": r"Xiaomi_Stock_Sellout_(\d{8})-(\d{8})\.xls", "dates": "start_end"},
    "IM_ST": {"file": r"Mobility - Xiaomi -Sales Phones Wk{week}\.xlsx"},
    # IM 库存文件按"发出那周"命名：W41 的库存叫 Internal W42
    "IM_STOCK": {"file": r"Internal W{next} 2026 Ingram SP split\.xlsx"},
}


def find_files(root, key, week):
    spec = FILES[key]
    rx = spec["file"].replace("{week}", f"0?{week}").replace("{next}", f"0?{week + 1}")
    found = []
    for p in candidates(root, {"file": rx}, week):
        if spec.get("dates"):
            m = re.fullmatch(rx, p.name, re.I)
            dates = [dt.datetime.strptime(g, "%Y%m%d").date() for g in m.groups()]
            start, end = week_range(week)
            if spec["dates"] == "start_end" and (dates[0], dates[1]) != (start, end):
                continue
            if spec["dates"] == "end" and dates[0] != end:
                continue
        found.append(p)
    # 邮件附件和另存的文件可能都在：同名只留一份
    uniq = {}
    for p in found:
        uniq.setdefault(p.name, p)
    return list(uniq.values())


def ean(x):
    if pd.isna(x):
        return None
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    s = re.sub(r"\.0$", "", str(x).strip())
    return s or None


def to_date(x):
    try:
        return pd.Timestamp(x).date()
    except (ValueError, TypeError):
        return None


def ab_st(path):
    """AB 原始数据一行一台（带序列号），冲正是一正一负两行。按 日期+产品+客户+物料码 对 Ilość 求和。"""
    df = pd.read_excel(path, sheet_name="Worksheet")
    # 同一客户大小写写法不一（MEDIA MARKT POLSKA SP. Z O.O / Sp. z o.o），Excel 透视表会合并，这里也合并，保留第一种写法
    df["_cust"] = df["Nazwa klienta"].astype(str).str.strip().str.casefold()
    first_name = df.groupby("_cust")["Nazwa klienta"].first()
    df["Nazwa klienta"] = df["_cust"].map(first_name)
    g = df.groupby(["Data faktury", "Nazwa produktu", "Nazwa klienta", "Kod produktu dostawcy"], dropna=False)
    out = g["Ilość"].agg(["sum", "count"]).reset_index()
    return pd.DataFrame({
        "C": out["Nazwa produktu"], "I": out["Nazwa klienta"], "J": out["sum"].astype(int),
        "K": out["Kod produktu dostawcy"], "L": None, "P": out["Data faktury"].map(to_date),
        "_count": out["count"],  # 计数，用来提示"计数和求和不一样"的行
    })


def ab_stock(path):
    df = pd.read_excel(path, sheet_name=0)
    return pd.DataFrame({"C": df["Opis"], "I": pd.to_numeric(df["Stan"], errors="coerce").fillna(0).astype(int),
                         "J": df["EAN"].map(ean)})


def also_st(path):
    """Sellout_PL 是按 客户+EAN 汇总的，没有日期；日期取 IMEI_Sellout 明细里同一 客户+EAN 的第一条
    （和原来的 XLOOKUP 一样，PSI 按周汇总所以不影响）。同时用明细条数核对 Suma，并检查退货。"""
    sell = fill_ean_from_part(pd.read_excel(path, sheet_name="Sellout_PL", dtype={"EAN": str}))
    imei = pd.read_excel(path, sheet_name="IMEI_Sellout", dtype={"EAN": str})
    returns = pd.read_excel(path, sheet_name="IMEI_Returns")
    imei["EAN"] = imei["EAN"].map(ean)
    first = imei.groupby(["Customer name", "EAN"])["Delivery date"].min()
    counts = imei.groupby(["Customer name", "EAN"]).size()
    sell["EAN"] = sell["EAN"].map(ean)
    keys = list(zip(sell["CustomerName"], sell["EAN"]))
    out = pd.DataFrame({
        "C": sell["ProductDescription"], "I": sell["CustomerName"], "J": sell["Suma"].astype(int),
        "K": None, "L": sell["EAN"], "P": [to_date(first.get(k)) for k in keys],
        "_imei": [int(counts.get(k, 0)) for k in keys],
    })
    return out, len(returns)


def fill_ean_from_part(df):
    """ALSO 有些行料号带 _P 后缀（如 82266_P）且 EAN 是空的：用去掉后缀的同料号那一行的 EAN 补上（和手工做法一致）。"""
    df = df.copy()
    df["EAN"] = df["EAN"].map(ean)
    part = df["VendorPartNumber2"].astype(str).str.strip()
    known = dict(zip(part, df["EAN"]))
    missing = df["EAN"].isna()
    df.loc[missing, "EAN"] = [known.get(p.split("_")[0]) for p in part[missing]]
    return df


def also_stock(path):
    df = fill_ean_from_part(pd.read_excel(path, sheet_name="Stock_PL", dtype={"EAN": str}))
    return pd.DataFrame({"C": df["ProductDescription"], "I": pd.to_numeric(df["Qty"], errors="coerce").fillna(0).astype(int),
                         "J": df["EAN"].map(ean)})


def im_st(path):
    raw = pd.read_excel(path, sheet_name=0, header=None, dtype=object)
    h = next(i for i in range(10) if "ITEMNAME" in [str(x).strip() for x in raw.iloc[i]])  # 有的周第一行是合计
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(x).strip() for x in raw.iloc[h]]
    df = df[df["ITEMNAME"].notna()]
    return pd.DataFrame({"C": df["ITEMNAME"], "I": df["NAME"], "J": pd.to_numeric(df["QTY"]).astype(int), "K": None,
                         "L": df["EAN"].map(ean), "P": df["INVOICEDATE"].map(to_date)})


def im_stock(path):
    raw = pd.read_excel(path, sheet_name=0, header=None, dtype=object)
    h = next(i for i in range(10) if str(raw.iat[i, 0]).strip() == "Row Labels")
    df = raw.iloc[h + 1:]
    cols = [str(x).strip() for x in raw.iloc[h]]
    df.columns = cols
    df = df[df["EAN"].notna()]
    return pd.DataFrame({"C": df["Row Labels"], "I": pd.to_numeric(df["Stock"], errors="coerce").fillna(0).astype(int),
                         "J": df["EAN"].map(ean)})


def sku_table(psi_path):
    """PSI 里 工具表1：69 码（M 列）和 SKU（G 列，AB 物料码）→ (项目, 配置)。"""
    wb = openpyxl.load_workbook(psi_path, read_only=True)
    ws = wb[SKU_TABLE]
    ws.reset_dimensions()
    by_ean, by_sku = {}, {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[1]:
            continue
        val = (r[1], r[2])
        if r[12]:
            by_ean.setdefault(ean(r[12]), val)
        if r[6]:
            by_sku.setdefault(str(r[6]).strip(), val)
    return by_ean, by_sku


def build_db(week, inbox, psi_path):
    by_ean, by_sku = sku_table(psi_path)
    st, stock, notes = [], [], []

    def add(rows, db, frame):
        frame = frame.copy()
        frame.insert(0, "B", db)
        frame.insert(0, "A", week)
        rows.append(frame)

    readers = [("AB", "AB_ST", ab_st, st), ("AB", "AB_STOCK", ab_stock, stock),
               ("IM", "IM_ST", im_st, st), ("IM", "IM_STOCK", im_stock, stock)]
    for db, key, fn, target in readers:
        files = find_files(inbox, key, week)
        if not files:
            notes.append(f"❌ {db} {'ST' if target is st else '库存'}：没找到当周文件")
            continue
        notes.append(f"- {db} {'ST' if target is st else '库存'}：`{files[0].name}`")
        add(target, db, fn(files[0]))
    files = find_files(inbox, "ALSO", week)
    if files:
        notes.append(f"- ALSO ST + 库存：`{files[0].name}`")
        frame, n_ret = also_st(files[0])
        if n_ret:
            notes.append(f"  - ⚠️ IMEI_Returns 里有 {n_ret} 条退货，Suma 是否已扣除需要确认")
        add(st, "ALSO", frame)
        add(stock, "ALSO", also_stock(files[0]))
    else:
        notes.append("❌ ALSO：没找到当周文件")

    st = pd.concat(st, ignore_index=True) if st else pd.DataFrame()
    stock = pd.concat(stock, ignore_index=True) if stock else pd.DataFrame()

    def lookup(row, ref_cols):
        for c in ref_cols:
            v = row.get(c)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            v = str(v).strip()
            hit = by_sku.get(v) if c == "K" else by_ean.get(v)
            if not hit and c == "K" and v.startswith("MZBO"):  # AB 物料码偶尔把数字 0 打成字母 O
                hit = by_sku.get("MZB0" + v[4:])
                if hit:
                    row["K"] = "MZB0" + v[4:]
            if hit:
                return f"{hit[0]}{str(hit[1]).replace('G', '')}"
        return None

    if len(st):
        rows = [r.to_dict() for _, r in st.iterrows()]
        st["代码（预查）"] = [lookup(r, ["K", "L"]) for r in rows]
        st["K"] = [r["K"] for r in rows]  # lookup 可能修正了物料码
        st, dropped, new = split_unknown(st)
        if len(dropped):
            notes.append(f"  - ST 里 {len(dropped)} 行是配件，没有放进待粘贴表："
                         + "；".join(f"{b} {c} → {i}（{q}）" for b, c, i, q in dropped[["B", "C", "I", "J"]].values))
        if len(new):
            notes.append(f"  - ⚠️ ST 里 {len(new)} 行 工具表1 查不到、但不像配件（疑似新品，已放进待粘贴表，粘贴前先补 工具表1）："
                         + "；".join(f"{b} {c}（{q}）" for b, c, q in new[["B", "C", "J"]].values))
        diff = st[st.get("_count", pd.Series(dtype=float)).notna() & (st.get("_count") != st["J"])] if "_count" in st else st.iloc[0:0]
        for _, r in diff.iterrows():
            notes.append(f"  - ⚠️ AB {r['C']} → {r['I']}：{int(r['_count'])} 行明细，Ilość 合计 {r['J']}（有冲正，按合计填；"
                         f"透视表若用'计数'会填成 {int(r['_count'])}）")
    if len(stock):
        stock["代码（预查）"] = [lookup(r, ["J"]) for _, r in stock.iterrows()]
        stock, dropped, new = split_unknown(stock)
        if len(dropped):
            notes.append(f"  - 库存里 {len(dropped)} 行是配件，没有放进待粘贴表："
                         + "；".join(f"{b} {c}（{q}）" for b, c, q in dropped[["B", "C", "I"]].values))
        if len(new):
            notes.append(f"  - ⚠️ 库存里 {len(new)} 行 工具表1 查不到、但不像配件（疑似新品，已放进待粘贴表，粘贴前先补 工具表1）："
                         + "；".join(f"{b} {c}（{q}）" for b, c, q in new[["B", "C", "I"]].values))
    return st, stock, notes


def split_unknown(df):
    """返回 (要贴的行, 丢掉的配件, 其中工具表1 查不到的疑似新品)。"""
    unknown = df["代码（预查）"].isna()
    accessory = unknown & df["C"].astype(str).str.contains(ACCESSORY)
    keep = df[~accessory]
    return keep, df[accessory], keep[keep["代码（预查）"].isna()]


def compare_db(psi_path, week, st, stock):
    """和 PSI 里已经贴好的该周行比对：按 经销商 + 识别码 + 客户（ST）/ 经销商 + EAN（库存）汇总数量。"""
    wb = openpyxl.load_workbook(psi_path, read_only=True)
    out = {}
    for name, ours, keys, qty in (("2026 DB ST", st, ["B", "ref", "I"], 9), ("2026 DB STOCK", stock, ["B", "J"], 8)):
        ws = wb[name]
        ws.reset_dimensions()
        rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if r and r[0] == week]
        if not rows or not len(ours):
            continue
        if name.endswith("ST"):
            psi = pd.DataFrame({"B": [r[1] for r in rows], "I": [str(r[8]).strip() for r in rows],
                                "ref": [str(r[10]).strip() if r[1] == "AB" else ean(r[11]) for r in rows],
                                "qty": [r[9] or 0 for r in rows]})
            mine = pd.DataFrame({"B": ours["B"], "I": ours["I"].astype(str).str.strip(),
                                 "ref": [str(k).strip() if b == "AB" else str(l) for b, k, l in zip(ours["B"], ours["K"], ours["L"])],
                                 "qty": ours["J"]})
        else:
            psi = pd.DataFrame({"B": [r[1] for r in rows], "J": [ean(r[9]) for r in rows], "qty": [r[8] or 0 for r in rows]})
            mine = pd.DataFrame({"B": ours["B"], "J": ours["J"].astype(str), "qty": ours["I"]})
        for col in ("I",):
            if col in psi:
                psi[col] = psi[col].str.split().str.join(" ")
                mine[col] = mine[col].str.split().str.join(" ")
        a = psi.groupby(keys)["qty"].sum().rename("PSI")
        b = mine.groupby(keys)["qty"].sum().rename("工具")
        j = pd.concat([a, b], axis=1).fillna(0).astype(int)
        diff = j[j["PSI"] != j["工具"]].reset_index()
        totals = psi.groupby("B")["qty"].sum().to_dict(), mine.groupby("B")["qty"].sum().to_dict()
        out[name] = (len(j), diff, totals)
    return out
