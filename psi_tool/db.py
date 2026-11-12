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
POCO_FROM_WEEK = 45  # 从这一周起 AB 的 POCO 表要一起算，没收到会提示
# 配件、生态链产品：工具表1 查不到且名称命中这些词的行不贴（SOP：手环、手机壳、平板笔之类直接删掉）
ACCESSORY = re.compile(r"cover|keyboard|klawiatur|\bpen\b|etui|case|filter|filtr|brush|strap|buds|watch|band|lamp|vacuum|"
                       r"camera|kamera|charg|ładow|adapter|stand|cable|kabel|purifier|scooter|tv\b|speaker|router", re.I)

# 每家的文件：文件名正则（{week} 会替换）+ 用来核对周的日期规则
FILES = {
    # 日期后面允许多出别的字（下载后改名，如 "——AB"）
    "AB_ST": {"file": r"xiaomi_smart_sell_through_(\d{8})_(\d{8})\D.*\.xlsx|xiaomi_smart_sell_through_(\d{8})_(\d{8})\.xlsx", "dates": "start_end"},
    "AB_STOCK": {"file": r"xiaomi_smart_stock_(\d{8})\D.*\.xlsx|xiaomi_smart_stock_(\d{8})\.xlsx", "dates": "end"},
    "ALSO": {"file": r"Xiaomi_Stock_Sellout_(\d{8})-(\d{8})\D.*\.xls|Xiaomi_Stock_Sellout_(\d{8})-(\d{8})\.xls", "dates": "start_end"},
    # 2026 W45 起 AB 另发 POCO 的表；还没见过样本，先按和 smart 表同样的命名和格式读，格式不一样会在报告里提示
    "AB_POCO_ST": {"file": r"xiaomi_poco_sell_through_(\d{8})_(\d{8})\D.*\.xlsx|xiaomi_poco_sell_through_(\d{8})_(\d{8})\.xlsx", "dates": "start_end"},
    "AB_POCO_STOCK": {"file": r"xiaomi_poco_stock_(\d{8})\D.*\.xlsx|xiaomi_poco_stock_(\d{8})\.xlsx", "dates": "end"},
    "ALSO_POCO": {"file": r"ALSO_Reporting_POCO_(\d{8})_(\d{8})\D.*\.xlsx?|ALSO_Reporting_POCO_(\d{8})_(\d{8})\.xlsx?", "dates": "start_end"},
    "IM_ST": {"file": r"Mobility - Xiaomi -Sales Phones Wk{week}.*\.xlsx"},
    # 2026 W44 起 IM 的 ST 也可能是串号明细：一行一台，按 型号 + 客户 + EAN + 日期 汇总
    "IM_IMEI": {"file": r"IMEI List Wk{week}\D.*\.xlsx|IMEI List Wk{week}\.xlsx"},
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
            dates = [dt.datetime.strptime(g, "%Y%m%d").date() for g in m.groups() if g]
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


def also_poco_st(path):
    """ALSO 的 POCO 表（W45 起）：Sellout 页一行一笔，按 日期 + 产品 + 客户 + EAN 汇总 Quantity。"""
    df = pd.read_excel(path, sheet_name="Sellout", dtype={"EAN": str})
    df["EAN"] = df["EAN"].map(ean)
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    g = df.groupby(["InvoiceDate", "ProductDescription", "CustomerName", "EAN"], dropna=False)["Quantity"].sum().reset_index()
    return pd.DataFrame({"C": g["ProductDescription"], "I": g["CustomerName"], "J": g["Quantity"].astype(int), "K": None,
                         "L": g["EAN"], "P": g["InvoiceDate"].map(to_date)})


def also_poco_stock(path):
    df = pd.read_excel(path, sheet_name="Inventory", dtype={"EAN": str})
    return pd.DataFrame({"C": df["ProductDescription"],
                         "I": pd.to_numeric(df["StockQuantity"], errors="coerce").fillna(0).astype(int),
                         "J": df["EAN"].map(ean)})


def im_st(path):
    raw = pd.read_excel(path, sheet_name=0, header=None, dtype=object)
    h = next(i for i in range(10) if "ITEMNAME" in [str(x).strip() for x in raw.iloc[i]])  # 有的周第一行是合计
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(x).strip() for x in raw.iloc[h]]
    df = df[df["ITEMNAME"].notna()]
    return pd.DataFrame({"C": df["ITEMNAME"], "I": df["NAME"], "J": pd.to_numeric(df["QTY"]).astype(int), "K": None,
                         "L": df["EAN"].map(ean), "P": df["INVOICEDATE"].map(to_date)})


def im_imei(path):
    """IMEI List：一行一个串号（Q SOLD 多为 1，退货为负），按 型号 + 客户 + EAN + 日期 汇总成和 Mobility 表一样的行。"""
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    df = df[df["model"].notna()].copy()
    for c in ("model", "CUSTOMER NAME"):
        df[c] = df[c].astype(str).str.split().str.join(" ")  # 原表里有双空格，PSI 里是单空格（工具表2 按单空格查）
    df["Q SOLD"] = pd.to_numeric(df["Q SOLD"], errors="coerce").fillna(0)
    df["EAN"] = df["MFR PART"].map(ean)
    df["DAY"] = df["INVOICE DAY"].map(to_date)
    g = df.groupby(["model", "CUSTOMER NAME", "EAN", "DAY"], as_index=False, dropna=False)["Q SOLD"].sum()
    g = g[g["Q SOLD"] != 0]
    return pd.DataFrame({"C": g["model"], "I": g["CUSTOMER NAME"], "J": g["Q SOLD"].astype(int), "K": None,
                         "L": g["EAN"], "P": g["DAY"]})


def im_crosscheck(inbox, week):
    """Mobility 和 IMEI List 都在时，按 EAN + 客户比两份的数量；返回不一致的 [dict]，只有一份时返回 None。"""
    mob, imei = find_files(inbox, "IM_ST", week), find_files(inbox, "IM_IMEI", week)
    if not (mob and imei):
        return None

    def by_key(df):
        return df.assign(I=df["I"].astype(str).str.split().str.join(" "), L=df["L"].astype(str)).groupby(["L", "I"])["J"].sum()
    a, b = by_key(im_st(mob[0])), by_key(im_imei(imei[0]))
    j = pd.concat([a.rename("m"), b.rename("i")], axis=1).fillna(0)
    j = j[j["m"] != j["i"]]
    raw = pd.read_excel(imei[0], sheet_name=0, dtype=object)
    raw["EAN"] = raw["MFR PART"].map(ean).astype(str)
    raw["CUSTOMER NAME"] = raw["CUSTOMER NAME"].astype(str).str.split().str.join(" ")
    out = []
    for (l, i), (m, n) in j.iterrows():
        rows = raw[(raw["EAN"] == l) & (raw["CUSTOMER NAME"] == i)]
        out.append(dict(ean=l, customer=i, mobility=int(m), imei=int(n),
                        model=str(rows["model"].iloc[0]).strip() if len(rows) else l,
                        invoices=sorted(rows["INVOICE NBR"].dropna().astype(str).unique()),
                        days=sorted({str(to_date(x)) for x in rows["INVOICE DAY"]}),
                        files=(mob[0].name, imei[0].name)))
    return out


def mark_im_files(inbox, week, diffs, out_dir):
    """把 Mobility 和 IMEI List 里对不上的行标红，存成副本（原始文件不动），可以直接截图或作为邮件附件。
    每份副本最前面加一页 Difference，用英文列出差异，方便发给 IM。返回副本路径列表。"""
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    red, bold = PatternFill("solid", fgColor="FFC7CE"), Font(bold=True, color="9C0006")
    keys = {(d["ean"], d["customer"]) for d in diffs}
    norm = lambda x: " ".join(str(x).split())
    out = []
    for key, ean_col, cust_col in (("IM_ST", "EAN", "NAME"), ("IM_IMEI", "MFR PART", "CUSTOMER NAME")):
        files = find_files(inbox, key, week)
        if not files:
            continue
        wb = load_workbook(files[0])
        ws = wb.worksheets[0]
        head = next(r for r in range(1, 12) if any(str(c.value).strip() == ean_col for c in ws[r]))
        cols = {str(c.value).strip(): c.column for c in ws[head] if c.value is not None}
        n = 0
        for r in range(head + 1, ws.max_row + 1):
            if (ean(ws.cell(r, cols[ean_col]).value), norm(ws.cell(r, cols[cust_col]).value)) in keys:
                for c in ws[r]:
                    c.fill, c.font = red, bold
                n += 1
        ws.freeze_panes = ws.cell(head + 1, 1)
        diff = wb.create_sheet("Difference", 0)
        diff.append(["Product", "EAN", "Customer", "Invoice", "Invoice date", "Mobility report", "IMEI list", "Difference"])
        for d in diffs:
            diff.append([d["model"], d["ean"], d["customer"], ", ".join(d["invoices"]), ", ".join(d["days"]),
                         d["mobility"], d["imei"], d["imei"] - d["mobility"]])
        for c in diff[1]:
            c.font = Font(bold=True)
        for row in diff.iter_rows(min_row=2):
            for c in row:
                c.fill, c.font = red, bold
        for col, w in zip("ABCDEFGH", (46, 16, 36, 14, 14, 16, 12, 12)):
            diff.column_dimensions[col].width = w
        diff.cell(len(diffs) + 3, 1, f"Highlighted rows in sheet '{ws.title}': {n}")
        wb.active = 0
        target = Path(out_dir) / f"[Highlighted] {files[0].name}"
        wb.save(target)
        out.append(target)
    return out


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
        # 原表里的名称常有双空格、不间断空格（IM 的 "Euro-net Sp.  z o.o."）；工具表2 按单空格查客户，
        # 手工粘贴时这些空格会被去掉，这里也统一压成一个空格，否则 PSI 公式查不到客户、OM 的 ST 会漏
        for col in ("C", "I"):
            if col in frame and frame[col].dtype == object:
                frame[col] = frame[col].map(lambda v: " ".join(v.split()) if isinstance(v, str) else v)
        frame.insert(0, "B", db)
        frame.insert(0, "A", week)
        rows.append(frame)

    def main_has_poco(rows, db, poco):
        """POCO 单独发的表：先看这家的主表里有没有 POCO（名称带 POCO，或同一个物料码 / EAN）。
        有 → 不再合并，免得算两次；没有 → 才把 POCO 表加进去（2026-11-12 胡佳琪确认的规则）。"""
        main = [f for f in rows if len(f) and f["B"].iat[0] == db]
        if not main or not len(poco):
            return False
        m = pd.concat(main)
        if m["C"].astype(str).str.contains(r"\bPOCO\b", case=False, regex=True).any():
            return True
        for c in (("K", "L") if rows is st else ("J",)):  # ST 的 J 是数量，库存的 J 才是 EAN
            if c in m and c in poco:
                a = set(m[c].dropna().astype(str).str.strip()) - {"", "None", "nan"}
                b = set(poco[c].dropna().astype(str).str.strip()) - {"", "None", "nan"}
                if a & b:
                    return True
        return False

    readers = [("AB", "AB_ST", ab_st, st), ("AB", "AB_STOCK", ab_stock, stock),
               ("AB", "AB_POCO_ST", ab_st, st), ("AB", "AB_POCO_STOCK", ab_stock, stock),
               ("IM", "IM_ST", im_st, st), ("IM", "IM_STOCK", im_stock, stock)]
    for db, key, fn, target in readers:
        files = find_files(inbox, key, week)
        if not files and key == "IM_ST":
            files, fn = find_files(inbox, "IM_IMEI", week), im_imei
        label = f"{db}{' POCO' if 'POCO' in key else ''} {'ST' if target is st else '库存'}"
        if not files:
            if "POCO" not in key or week >= POCO_FROM_WEEK:
                notes.append(f"❌ {label}：没找到当周文件")
            continue
        notes.append(f"- {label}：`{files[0].name}`")
        try:
            frame = fn(files[0])
            if "POCO" in key and not len(frame):
                notes.append("  - 这份是空表（只有表头），本周没有数")
            if "POCO" in key and main_has_poco(target, db, frame):
                notes.append("  - 主表里已经有 POCO 型号，这份 POCO 表不再合并（避免算两次）")
                continue
            add(target, db, frame)
        except (KeyError, ValueError) as e:  # 新表格式和预想的不一样：不算进去，提示人看
            notes.append(f"  - ⚠️ {files[0].name} 的格式和 AB smart 表不一样（{e}），这份没有算进去，把文件发给工具维护人改读取方式")
    # IM 同时给了 Mobility 和 IMEI List：以 Mobility 为准（SOP），两份按 EAN + 客户对一下，不一致的列出来去问 IM
    diffs = im_crosscheck(inbox, week)
    if diffs is not None and len(diffs):
        # 2026-11-09 已和 IM 确认：IMEI List 按发货日期、Mobility 按发票日期算周，账期口径不同，差额是时间差；以 Mobility 为准
        notes.append(f"  - ℹ️ IM 的 Mobility（PSI 用这份）和 IMEI List 有 {len(diffs)} 处不同，属于口径差异"
                     "（IMEI 按发货日期、Mobility 按发票日期），不用处理："
                     + "；".join(f"{d['model']} → {d['customer']}：Mobility {d['mobility']} / IMEI {d['imei']}" for d in diffs))
    elif diffs is not None:
        notes.append("  - IM 的 Mobility 和 IMEI List 按 EAN + 客户完全一致")
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
    files = find_files(inbox, "ALSO_POCO", week)  # 放在主表之后：要先知道主表里有没有 POCO
    if files:
        notes.append(f"- ALSO POCO ST + 库存：`{files[0].name}`")
        try:
            pst, pstock = also_poco_st(files[0]), also_poco_stock(files[0])
            if not len(pst) and not len(pstock):
                notes.append("  - 这份是空表（只有表头），本周 ALSO 的 POCO 没有销量和库存")
            for rows, frame, what in ((st, pst, "ST"), (stock, pstock, "库存")):
                if main_has_poco(rows, "ALSO", frame):
                    notes.append(f"  - 主表里已经有 POCO 型号，这份 POCO 表的{what}不再合并（避免算两次）")
                else:
                    add(rows, "ALSO", frame)
        except (KeyError, ValueError) as e:
            notes.append(f"  - ⚠️ {files[0].name} 的格式和预想的不一样（{e}），这份没有算进去")
    elif week >= POCO_FROM_WEEK:
        notes.append("❌ ALSO POCO：没找到当周文件")

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
