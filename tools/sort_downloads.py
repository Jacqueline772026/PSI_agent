"""整理下载文件夹：把一堆下载下来的周报（xlsx / xls / msg）按 来源 + 周 放进标准文件夹。

  下载文件夹/ 里乱放的文件
        ↓
  W44/W44原始文件/DB/AB/…  OM/EURO/…  OP/TMO/…     ← 当周的
  W44/_其他周/W43/OM/KTR/…                          ← 不是当周的，单独放，不混进来
  下载文件夹里剩下的                                  ← 认不出来的，原地不动，列在整理记录里

怎么认：
  1. 来源：先按文件名（config/sources.yaml 的 file 规则，周数换成任意数字），
     对不上或者同时对上几家时，再打开文件看表头。GG 和 ROLV 文件名一样，按 Customer Name 区分。
  2. 周：文件名里的 W44 / week 44 / Wk44 / T44 / 日期，文件里的统计起止日期；都没有就算目标周，并在记录里注明。
  3. .msg 邮件：取出 Excel 附件（附件是 zip 的先解压）分别归类，邮件本身放在第一个附件旁边。zip 文件同样解压后归类。

文件名不改（工具按原始文件名认文件），同名不同内容的加 (2)。整理记录写在 W{周}原始文件/整理记录.md。

用法：
  .venv/bin/python -m tools.sort_downloads ~/Downloads/PSI --week 44 --dest W44测试/W44原始文件
  加 --copy 只复制不移动（下载文件夹保持原样）；加 --dry-run 只看会怎么放，不动文件。
"""
import argparse
import datetime as dt
import re
import shutil
import tempfile
from pathlib import Path

import pandas as pd

from psi_tool.paths import CONFIG
from psi_tool.sources import detect_period, find_header, load_config, read_sheet

EXCEL = (".xlsx", ".xls", ".xlsm")
# DB 三家的文件规则在 psi_tool/db.py 里是写死的，这里对应一份
DB_RULES = [
    ("DB/AB", r"xiaomi_smart_(sell_through_\d{8}_\d{8}|stock_\d{8})\.xlsx"),
    ("DB/ALSO", r"Xiaomi_Stock_Sellout_\d{8}-\d{8}.*\.xls"),
    ("DB/IM", r"(Mobility - Xiaomi -Sales Phones Wk\d+|Internal W\d+ 2026 Ingram SP split|IMEI List Wk\d+).*\.xlsx"),
]


def iso_week(d):
    return d.isocalendar()[1]


def week_from_name(name, folder):
    """从文件名推周；返回 None 表示文件名里没有。"""
    n = name
    if folder == "DB/IM" and (m := re.search(r"Internal W(\d+)", n)):
        return int(m.group(1)) - 1  # IM 库存文件按发出那周命名：W45 的文件是 W44 的库存
    if m := re.search(r"(\d{8})[_-](\d{8})", n):  # AB / ALSO：起止日期
        return iso_week(dt.datetime.strptime(m.group(1), "%Y%m%d").date())
    if m := re.search(r"stock_(\d{8})", n):  # AB 库存：截止日
        return iso_week(dt.datetime.strptime(m.group(1), "%Y%m%d").date())
    if m := re.search(r"sprzedaz\s*(\d{4}-\d{2}-\d{2})", n):  # Plus：下周一的日期
        return iso_week(dt.date.fromisoformat(m.group(1)) - dt.timedelta(days=1))
    if m := re.search(r"na dzień_(\d{4}-\d{2}-\d{2})", n):  # MEX：周日的日期
        return iso_week(dt.date.fromisoformat(m.group(1)))
    if folder == "OM/MAX" and (m := re.search(r"Xiaomi_(\d+)_20\d\d", n)):
        return int(m.group(1))
    if folder == "OM/EURO" and (m := re.search(r"_(?:tele|TAB)_w?0?(\d{1,2})(?![0-9])", n, re.I)):
        return int(m.group(1))
    if folder == "OM/Auchan" and (m := re.search(r"Xiaomi T(\d+)\b", n)):
        return int(m.group(1))
    if m := re.search(r"(?:\bW|week\s*|Wk|_w)0?(\d{1,2})(?![0-9])", n, re.I):
        return int(m.group(1))
    return None


def any_week(rx):
    return rx.replace("{week}", r"\d{1,2}")


class Classifier:
    def __init__(self, config_path):
        self.cfg = load_config(config_path)

    def by_name(self, name):
        hits = []
        for src, c in self.cfg.items():
            if re.fullmatch(any_week(c["file"]), name, re.I):
                hits.append((c["folder"], src))
        for folder, rx in DB_RULES:
            if re.fullmatch(rx, name, re.I):
                hits.append((folder, None))
        return hits

    def by_content(self, path, candidates):
        """打开文件看表头，确认是哪一家。candidates 为空就试所有来源。"""
        pool = candidates or [(c["folder"], s) for s, c in self.cfg.items()]
        for folder, src in pool:
            if src is None:
                return folder, src  # DB 文件名规则已经足够明确
            c = self.cfg[src]
            raw = read_sheet(path, c.get("sheet", 0))
            if raw is None:
                continue
            # 表头里带 {week} 的列名（KTR 的 "stan W{week}"）只比前缀，其余要完全一样
            wanted = [(h.split("{week}")[0].strip(), "{week}" in h) for h in c["header"]]
            for i in range(min(40, len(raw))):
                cells = [str(x).strip() for x in raw.iloc[i] if pd.notna(x)]
                if all(any(cell.startswith(h) if prefix else cell == h for cell in cells) for h, prefix in wanted):
                    return folder, src
        return None, None

    def michain_customer(self, path):
        try:
            df = pd.read_excel(path, sheet_name=None, nrows=5)
        except ValueError:
            return None
        for d in df.values():
            for col in ("Customer Name", "RMS Customer Name"):
                if col in d.columns and d[col].notna().any():
                    v = str(d[col].dropna().iloc[0]).upper()
                    return "GG" if "GG" in v else "ROLV" if "ROLV" in v else None
        return None

    def week_from_content(self, path, src):
        c = self.cfg[src]
        if not c.get("period"):
            return None
        raw = read_sheet(path, c.get("sheet", 0))
        if raw is None:
            return None
        h = find_header(raw, [x.replace("{week}", "") for x in c["header"]] if "{week}" not in str(c["header"]) else [])
        per = detect_period(raw, h, c["period"])
        return iso_week(per[1]) if per and per[1] else None

    def classify(self, path):
        """返回 (子文件夹, 周, 说明)；认不出来子文件夹为 None。"""
        hits = self.by_name(path.name)
        folder, src = hits[0] if len(hits) == 1 else self.by_content(path, hits)
        if folder is None:
            return None, None, "文件名和表头都认不出是哪一家"
        note = []
        if "{michain}" in folder:
            who = self.michain_customer(path)
            if who is None:
                return None, None, "是 Michain 的文件，但分不出 GG 还是 ROLV"
            folder = folder.replace("{michain}", who)
        week = week_from_name(path.name, folder)
        if week is None and src:
            week = self.week_from_content(path, src)
            if week:
                note.append("周数取自文件里的统计日期")
        if week is None:
            note.append("文件里没有周信息，按目标周归档")
        return folder, week, "；".join(note)


def unique_target(target):
    if not target.exists():
        return target
    i = 2
    while (t := target.with_name(f"{target.stem} ({i}){target.suffix}")).exists():
        i += 1
    return t


CONFLICT = "_同名不同内容"


def place(src_path, target, copy, dry_run):
    """返回实际放到的位置；内容完全相同的重复文件返回 None。
    同名但内容不同（常见：一份邮件原件、一份自己另存改过的）：不放进同一个文件夹，
    否则工具会读两遍重复计数；第二份放到 W{周}/_同名不同内容/ 里等人决定。"""
    if target.exists():
        if target.read_bytes() == Path(src_path).read_bytes():
            return None
        parts = target.parts
        root = next(i for i, x in enumerate(parts) if x.endswith("原始文件") or x == "_其他周")
        target = Path(*parts[:root]) / CONFLICT / Path(*parts[root + 1:])
    target = unique_target(target)
    if not dry_run:
        target.parent.mkdir(parents=True, exist_ok=True)
        (shutil.copy2 if copy else shutil.move)(str(src_path), str(target))
    return target


def unzip(path, tmp):
    """zip 里的 Excel 解到临时文件夹（GG 的周报是 OneDrive 打包的 zip）。"""
    import zipfile
    out = []
    target = Path(tempfile.mkdtemp(dir=tmp))
    with zipfile.ZipFile(path) as z:
        for info in z.infolist():
            name = Path(info.filename).name
            if name.lower().endswith(EXCEL) and not name.startswith(("~$", ".")):
                p = target / name
                p.write_bytes(z.read(info))
                out.append(p)
    return out


def msg_attachments(path, tmp):
    import extract_msg
    out = []
    msg = extract_msg.Message(str(path))
    try:
        for att in msg.attachments:
            name = (getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or "").replace("\x00", "").strip()
            if not isinstance(att.data, bytes) or not name.lower().endswith(EXCEL + (".zip",)):
                continue
            p = Path(tempfile.mkdtemp(dir=tmp)) / Path(name).name
            p.write_bytes(att.data)
            out += unzip(p, tmp) if p.suffix.lower() == ".zip" else [p]
        subject = msg.subject or ""
    finally:
        msg.close()
    return out, subject


def sort_folder(download, week, dest, copy=False, dry_run=False, config=CONFIG):
    clf = Classifier(config)
    download, dest = Path(download), Path(dest)
    other = dest.parent / "_其他周"
    log = []

    def target_dir(folder, wk):
        return dest / folder if wk in (None, week) else other / f"W{wk}" / folder

    def rel(d):
        for root, label in ((dest, dest.name), (other, "_其他周")):
            try:
                return f"{label}/{d.relative_to(root).as_posix()}"
            except ValueError:
                pass
        return str(d)

    def put(path, folder, wk, note, label, always_copy=False):
        d = target_dir(folder, wk)
        t = place(path, d / path.name, copy or always_copy, dry_run)
        if t is not None and CONFLICT in t.parts:
            log.append(("⚠️", label, f"{CONFLICT}/{t.parent.name}", f"{rel(d)} 里已有同名但内容不同的文件（可能一份是原件、"
                        "一份是改过的副本）。工具只读前者，请确认留哪份"))
            return d
        log.append(("✅" if t else "＝", label, rel(d), note if t else "重复：那里已经有完全相同的文件"))
        return d

    # 先处理邮件和压缩包：同名冲突时，邮件里的原件优先留在正式文件夹
    files = sorted((p for p in download.rglob("*") if p.is_file() and not p.name.startswith((".", "~$"))),
                   key=lambda p: (p.suffix.lower() not in (".msg", ".zip"), str(p)))
    with tempfile.TemporaryDirectory() as tmp:
        for p in files:
            if p.suffix.lower() == ".msg":
                atts, subject = msg_attachments(p, tmp)
                first = None
                hint = week_from_name(p.name, None)  # 附件名里没周数时，参考邮件文件名
                for a in atts:
                    folder, wk, note = clf.classify(a)
                    if folder and wk is None and hint:
                        wk, note = hint, "周数取自邮件文件名"
                    if folder is None:
                        log.append(("❓", f"{a.name}（邮件 {p.name} 的附件）", "—", note))
                        continue
                    d = put(a, folder, wk, note, f"{a.name}（邮件 {p.name} 的附件）", always_copy=True)
                    first = first or d
                if first is None:
                    log.append(("❓", p.name, "—", f"邮件里没有能认出的 Excel 附件（标题：{subject[:40]}）"))
                    continue
                t = place(p, first / p.name, copy, dry_run)
                log.append(("✅" if t else "＝", p.name, rel(first), "邮件原件，放在附件旁边"))
                continue
            if p.suffix.lower() == ".zip":
                for a in unzip(p, tmp):
                    folder, wk, note = clf.classify(a)
                    if folder is None:
                        log.append(("❓", f"{a.name}（压缩包 {p.name} 里的）", "—", note))
                        continue
                    put(a, folder, wk, note, f"{a.name}（压缩包 {p.name} 里的）", always_copy=True)
                continue
            if p.suffix.lower() not in EXCEL:
                log.append(("❓", p.name, "—", "不是 Excel、邮件或压缩包"))
                continue
            folder, wk, note = clf.classify(p)
            if folder is None:
                log.append(("❓", p.name, "—", note))
                continue
            put(p, folder, wk, note, p.name)
    write_log(dest, week, download, log, dry_run)
    return log


def write_log(dest, week, download, log, dry_run):
    lines = [f"# W{week} 下载文件整理记录{'（试运行，没有移动文件）' if dry_run else ''}", "",
             f"- 时间：{dt.datetime.now():%Y-%m-%d %H:%M}", f"- 下载文件夹：`{download}`", "",
             "✅ 已归档　＝ 重复文件（已有完全相同的一份）　⚠️ 同名不同内容，需要你确认　❓ 认不出来，留在原处", "",
             "| | 文件 | 放到 | 说明 |", "|---|---|---|---|"]
    lines += [f"| {a} | {b} | {c} | {d} |" for a, b, c, d in log]
    if not dry_run:
        dest.mkdir(parents=True, exist_ok=True)
        (dest / "整理记录.md").write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))


def main():
    ap = argparse.ArgumentParser(description="整理下载的周报文件")
    ap.add_argument("download", help="下载文件夹")
    ap.add_argument("--week", type=int, required=True, help="目标周，例如 44")
    ap.add_argument("--dest", help="默认 W{周}/W{周}原始文件")
    ap.add_argument("--copy", action="store_true", help="只复制，不移动")
    ap.add_argument("--dry-run", action="store_true", help="只显示会怎么放，不动文件")
    a = ap.parse_args()
    dest = a.dest or f"W{a.week}/W{a.week}原始文件"
    sort_folder(a.download, a.week, dest, a.copy, a.dry_run)


if __name__ == "__main__":
    main()
