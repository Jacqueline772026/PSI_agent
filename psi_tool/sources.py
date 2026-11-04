"""按 config/sources.yaml 找到当周的原始文件，读成统一格式的明细表。

统一格式的列：source, psi_sheet, file, key_type, key, name, so, stock
"""
import datetime as dt
import io
import re
import warnings
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import yaml

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

DATE = re.compile(r"\d{4}-\d{2}-\d{2}|\d{2}[./]\d{2}[./]\d{4}")


def load_config(path):
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def week_range(week, year=2026):
    return dt.date.fromisocalendar(year, week, 1), dt.date.fromisocalendar(year, week, 7)


def fmt(value, week):
    """把配置里的 {week} 换成周数。正则里用 0?41 兼容 W041 / W41。"""
    if isinstance(value, str):
        return value.replace("{week}", str(week))
    if isinstance(value, list):
        return [fmt(v, week) for v in value]
    if isinstance(value, dict):
        return {k: fmt(v, week) for k, v in value.items()}
    return value


def week_regex(value, week):
    return value.replace("{week}", f"0?{week}")


def parse_date(x):
    if isinstance(x, (dt.date, dt.datetime, pd.Timestamp)):
        return pd.Timestamp(x).date()
    s = str(x).strip()
    for f in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s[:19] if " " in s else s, f).date()
        except ValueError:
            pass
    return None


@dataclass
class SourceResult:
    name: str
    psi_sheet: str
    files: list = field(default_factory=list)
    rows: pd.DataFrame = None
    raw_total_so: float = 0  # 品类筛选之前（文件里全部行）
    raw_total_stock: float = 0
    period: tuple = None  # (from, to)
    warnings: list = field(default_factory=list)
    name_excluded: pd.DataFrame = None  # 被 exclude_name 丢掉的行（给防串检查用）


def read_sheet(path, sheet):
    """读成没有表头的 DataFrame；sheet 不存在返回 None。"""
    try:
        return pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    except ValueError:
        return None  # 没有这个 sheet


def find_header(raw, wanted):
    wanted = [str(w).strip() for w in wanted]
    for i in range(min(40, len(raw))):
        cells = {str(x).strip() for x in raw.iloc[i] if pd.notna(x)}
        if all(w in cells for w in wanted):
            return i
    return None


def column(df_raw, header_row, spec):
    """spec 可以是列名，也可以是列序号（表头是空格子时用序号）。"""
    if isinstance(spec, int):
        return spec
    names = [str(x).strip() if pd.notna(x) else "" for x in df_raw.iloc[header_row]]
    if str(spec).strip() not in names:
        raise KeyError(spec)
    return names.index(str(spec).strip())  # 重名列取第一个（后面的通常是人工加的）


def flatten_weekly(raw, h, week, metrics):
    """Jet Wu 的 Orange 表这类"周横着排"的表：第 h 行是周（W01…），第 h+1 行是指标名（Delivery/Stock/Sales）。
    取当周那一段，拼成普通的竖表：第 0 行表头 = [第一列名, 指标…]。
    同一个周号出现多次时（表里有去年的周），取最后一个 W01 之后的那一段。"""
    weeks = raw.iloc[h].ffill()
    names = raw.iloc[h + 1].astype(str).str.strip()
    labels = [str(w).strip().upper() for w in weeks]
    start = max((i for i, w in enumerate(labels) if w == "W01"), default=0)
    target = f"W{week:02d}"
    cols = {}
    for i in range(start, raw.shape[1]):
        if labels[i] == target and names.iloc[i] in metrics:
            cols[names.iloc[i]] = i
    if len(cols) < len(metrics):
        return None
    flat = raw.iloc[h + 2:, [0] + [cols[m] for m in metrics]].copy()
    header = pd.DataFrame([[raw.iloc[h, 0]] + metrics], columns=flat.columns)
    return pd.concat([header, flat], ignore_index=True)


def detect_period(raw, header_row, spec):
    if not spec:
        return None
    if "cell_regex" in spec:
        rx = re.compile(spec["cell_regex"])
        for i in range(min(15, len(raw))):
            for x in raw.iloc[i]:
                m = rx.search(str(x)) if pd.notna(x) else None
                if m:
                    return parse_date(m.group(1)), parse_date(m.group(2))
    if "label_from" in spec:
        found = {}
        for i in range(min(15, len(raw))):
            row = list(raw.iloc[i])
            for j, x in enumerate(row):
                for k in ("label_from", "label_to"):
                    if pd.notna(x) and str(x).strip() == spec[k] and j + 1 < len(row):
                        found[k] = parse_date(row[j + 1])
        if len(found) == 2:
            return found["label_from"], found["label_to"]
    if "column" in spec and header_row is not None:
        c = column(raw, header_row, spec["column"])
        dates = [d for d in (parse_date(x) for x in raw.iloc[header_row + 1:, c] if pd.notna(x)) if d]
        if dates:
            return min(dates), max(dates)
    return None


def period_ok(period, week):
    if not period or None in period:
        return None
    start, end = week_range(week)
    return start <= period[0] and period[1] <= end


MSG_CACHE = Path(__file__).resolve().parent.parent / "cache" / "msg"


def extract_msg_attachments(root):
    """把 root 下所有 .msg 邮件的附件解到 cache/msg/<相对路径>/ 里（已解过的跳过）。
    很多周报只存了邮件没另存附件，比如 EURO 的 'RE_ tsma w40.msg'。"""
    import extract_msg
    out = []
    for m in Path(root).rglob("*.msg"):
        target = MSG_CACHE / m.relative_to(root).with_suffix("")
        if not target.exists():
            try:
                msg = extract_msg.Message(str(m))
                target.mkdir(parents=True, exist_ok=True)
                for att in msg.attachments:
                    name = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None)
                    name = (name or "").replace("\x00", "").strip()
                    if name and name.lower().endswith((".xlsx", ".xls", ".csv")):
                        (target / Path(name).name).write_bytes(att.data)
                    elif name.lower().endswith(".zip"):  # GG 的周报是 OneDrive 打包的 zip
                        with zipfile.ZipFile(io.BytesIO(att.data)) as z:
                            for info in z.infolist():
                                n = Path(info.filename).name
                                if n.lower().endswith((".xlsx", ".xls")) and not n.startswith(("~$", ".")):
                                    (target / n).write_bytes(z.read(info))
                msg.close()
            except Exception:  # 损坏的邮件或特殊附件，跳过
                target.mkdir(parents=True, exist_ok=True)
        out.append((target, m.parent))
    return out


def candidates(root, cfg, week):
    file_rx = re.compile(week_regex(cfg["file"], week), re.I)
    path_rx = re.compile(week_regex(cfg["week_in_path"], week), re.I) if cfg.get("week_in_path") else None
    out = []
    # 邮件附件只在邮件所在文件夹里没有另存同名文件时才用（另存的那份可能是人工改过的）
    # 路径规则也看 --inbox 那一层的名字，这样 inbox 文件夹本身叫 W44原始文件 也算"路径里有 W44"
    roots = [(Path(root), Path(root), None)] + [(d, MSG_CACHE, folder) for d, folder in extract_msg_attachments(root)]
    seen = set()
    for base, rel_base, msg_folder in roots:
        for p in sorted(base.rglob("*")):
            if not p.is_file() or p.name.startswith("~$") or not file_rx.fullmatch(p.name):
                continue
            if msg_folder is not None and (msg_folder / p.name).exists():
                continue
            rel = f"{Path(root).resolve().name}/{p.resolve().relative_to(rel_base.resolve()).as_posix()}"
            if path_rx and not path_rx.search(rel):
                continue
            if (p.name, p.stat().st_size) in seen:  # 邮件附件和另存的文件是同一份
                continue
            seen.add((p.name, p.stat().st_size))
            out.append(p)
    return out



def to_number(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


def normalize_key(x, key_type, add_6=False):
    if pd.isna(x):
        return None
    if key_type in ("EAN", "TTL_ID"):
        if isinstance(x, float) and x.is_integer():
            x = int(x)
        s = re.sub(r"\.0$", "", str(x).strip())
        if not s.isdigit():
            return None  # 表头重复行、"Suma:" 合计行等
        if add_6 and not s.startswith("6"):
            s = "6" + s  # KTR 的 EAN 开头少了 6
        return s
    s = re.sub(r"\s+", " ", str(x).strip())
    return s.upper() if key_type.endswith("_NAME") else s


def load_source(root, name, cfg, week):
    # 列名里的 {week} 直接替换；正则里的 {week} 由 week_regex 处理
    cfg = {**cfg, **{k: fmt(cfg[k], week) for k in ("header", "so", "stock") if k in cfg}}
    res = SourceResult(name, cfg["psi_sheet"])
    frames = []
    for path in candidates(root, cfg, week):
        if cfg.get("sheet") == "any":
            names = pd.ExcelFile(path).sheet_names
        else:
            names = [cfg.get("sheet", 0)]
        raw, h = None, None
        for sheet_name in names:
            raw = read_sheet(path, sheet_name)
            h = None if raw is None else find_header(raw, cfg["header"])
            if h is not None:
                break
        alt = False
        if h is None and cfg.get("header_alt"):  # 客户换了格式
            for sheet_name in names:
                raw = read_sheet(path, sheet_name)
                h = None if raw is None else find_header(raw, cfg["header_alt"])
                if h is not None:
                    alt = True
                    break
        if raw is None:
            continue
        if h is None and week < 10:  # KTR 的列名个位数周带前导 0：stan W05
            h = find_header(raw, [x.replace(f"W{week}", f"W{week:02d}") for x in cfg["header"]])
            if h is not None:
                cfg = {k: ([x.replace(f"W{week}", f"W{week:02d}") for x in v] if isinstance(v, list)
                           else v.replace(f"W{week}", f"W{week:02d}"))
                       if k in ("header", "so", "stock") else v for k, v in cfg.items()}
        if h is None:
            continue
        if cfg.get("layout") == "weekly_columns":
            metrics = [m for m in (cfg.get("so"), cfg.get("stock")) if m]
            raw = flatten_weekly(raw, h, week, metrics)
            if raw is None:
                res.warnings.append(f"{path.name} 里没有 W{week:02d} 这一周")
                continue
            h = 0
        period = detect_period(raw, h, cfg.get("period"))
        ok = period_ok(period, week)
        if ok is False:
            continue  # 别的周的文件
        frames.append((path, raw, h, period, alt))
    if not frames:
        res.warnings.append("没找到当周文件")
        return res
    if len(frames) > 1 and not cfg.get("multiple"):
        newest = max(frames, key=lambda f: f[0].stat().st_mtime)  # 邮件附件是刚解出来的，通常是原件
        res.warnings.append(f"找到 {len(frames)} 个候选文件，用了最新的 {newest[0].name}："
                            + "、".join(f[0].name for f in frames))
        frames = [newest]

    parts = []
    for path, raw, h, period, alt in frames:
        if alt:
            res.warnings.append(f"{path.name} 是新格式（{'、'.join(cfg['header_alt'])}），没有品类列：只保留查得到代码的行")
        res.files.append(str(path))
        if period:
            res.period = period
            if period_ok(period, week) is None:
                res.warnings.append(f"{path.name} 读不出完整时间段")
        elif cfg.get("period"):
            res.warnings.append(f"{path.name} 没读到时间段，无法核对是不是 W{week}")
        body = raw.iloc[h + 1:]
        try:
            c_key = column(raw, h, cfg["key"]["column"])
            c_name = column(raw, h, cfg["name"])
            c_so = column(raw, h, cfg["so"]) if cfg.get("so") else None
            # stock 可以是一列，也可以是几列相加（Orange = 中央仓 + 门店 + B2B）
            stock_specs = cfg.get("stock") or []
            c_stock = [column(raw, h, x) for x in (stock_specs if isinstance(stock_specs, list) else [stock_specs])]
        except KeyError as e:
            res.warnings.append(f"{path.name} 缺少列 {e}，文件格式可能变了")
            continue
        keys = [normalize_key(x, cfg["key"]["type"], cfg.get("ean_add_6")) for x in body.iloc[:, c_key]]
        if cfg["key"].get("fallback"):
            c_fb = column(raw, h, cfg["key"]["fallback"])
            keys = [k or normalize_key(x, cfg["key"]["type"], cfg.get("ean_add_6"))
                    for k, x in zip(keys, body.iloc[:, c_fb])]
        df = pd.DataFrame({
            "key": keys,
            # pandas 3 的 astype(str) 会保留空值，统一成空字符串，否则没有名称的汇总行（如占比合计）会被当成产品
            "name": body.iloc[:, c_name].astype(str).str.strip().fillna(""),
            "so": to_number(body.iloc[:, c_so]) if c_so is not None else 0,
            "stock": sum(to_number(body.iloc[:, c]) for c in c_stock) if c_stock else 0,
        })
        # 识别码是空的或不是数字：有数量、有产品名的行保留下来记为未匹配，其余（空行、重复表头）丢掉
        raw_key = body.iloc[:, c_key].astype(str).str.strip().fillna("nan")
        has_qty = (df["so"] != 0) | (df["stock"] != 0)
        named = ~df["name"].isin(["", "nan"])
        keep = df["key"].notna() | (has_qty & named)
        df.loc[df["key"].isna(), "key"] = raw_key.where(raw_key != "nan", "（空）")
        if cfg.get("week_filter"):
            wf = cfg["week_filter"]
            c = column(raw, h, wf["column"])
            keep &= body.iloc[:, c].loc[df.index].astype(str).str.strip().str.fullmatch(week_regex(wf["regex"], week))
        df = df[keep]
        res.raw_total_so += df["so"].sum()
        res.raw_total_stock += df["stock"].sum()
        keep = pd.Series(True, index=df.index)
        if cfg.get("include") and not alt:
            c = column(raw, h, cfg["include"]["column"])
            vals = body.iloc[:, c].loc[df.index].astype(str)
            keep &= vals.str.contains(cfg["include"]["regex"], regex=True, flags=re.I)
        if alt and cfg.get("alt_exclude_name"):
            keep &= ~df["name"].str.contains(cfg["alt_exclude_name"], regex=True, flags=re.I)
        if cfg.get("exclude_name"):
            hit = df["name"].str.contains(cfg["exclude_name"], regex=True, flags=re.I)
            dropped = df[keep & hit]
            res.name_excluded = dropped if res.name_excluded is None else pd.concat([res.name_excluded, dropped])
            keep &= ~hit
        df = df[keep].copy()
        df["file"] = path.name
        df["drop_unmatched"] = alt
        parts.append(df)
    if parts:
        res.rows = pd.concat(parts, ignore_index=True)
        res.rows["source"] = name
        res.rows["psi_sheet"] = cfg["psi_sheet"]
        res.rows["key_type"] = cfg["key"]["type"]
        res.rows["others"] = cfg.get("others", True)
    return res
