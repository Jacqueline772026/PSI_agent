"""从周报邮件原件（.msg）里整理每家的联系人：谁发数据（客户方）、小米这边谁在收件 / 抄送。

发现原始数据有问题、要确认口径时，先查这张表：
  - 客户方联系人：最近几周发周报的人，直接回复他们的周报邮件最好（对方有上下文）
  - 小米对接人：一直在收件 / 抄送里的小米同事，一般是这家客户的 KAM 或负责人，发外部邮件前先和他们 / 导师确认

输出：数据/联系人.xlsx（含个人邮箱，属于公司数据，只放本地，已被 .gitignore 排除）
用法：.venv/bin/python -m tools.contacts [文件夹 ...]   默认扫 数据/历史周报 和 W*/ 下所有 .msg
"""
import re
import sys
from collections import defaultdict
from pathlib import Path

import extract_msg
import pandas as pd

from psi_tool.paths import DATA, HISTORY, ROOT
from tools.sort_downloads import TAGS

ADDR = re.compile(r'(?:"?([^"<>;\t]*?)"?\s*)?<?([\w.+\'-]+@[\w-]+(?:\.[\w-]+)+)>?')
GENERIC = re.compile(r"(?i)^(pl-psi|poland-so|noreply|no-reply|vrep|vendor_reports|.*raporty)@")


def source_of(path):
    """按所在文件夹（如 OM/MSH、W42-43/W42 下的文件名标记）判断是哪一家。"""
    for part in reversed(path.parts[:-1]):
        key = re.sub(r"[^A-Z-]", "", part.upper())
        for tag in sorted(TAGS, key=len, reverse=True):
            if key == tag or key == tag.replace("-", ""):
                return TAGS[tag].split("/")[-1]
        if part.upper() in ("T-MOBILE",):
            return "TMO"
    from tools.sort_downloads import tag_folder
    t = tag_folder(path.name)
    return t.split("/")[-1] if t else None


def people(field):
    out = []
    for name, addr in ADDR.findall(field or ""):
        addr = addr.strip("'")
        name = re.sub(r"[\x00-\x1f]", "", name)  # 邮件头里偶尔有控制字符，Excel 写不进去
        name = re.sub(r"\s+", " ", name.replace("'", "").strip(" ,;")) or addr
        out.append((name, addr.lower().strip("'")))
    return out


def scan(folders):
    rows = []
    for folder in folders:
        for p in Path(folder).rglob("*.msg"):
            try:
                m = extract_msg.Message(p)
                date = pd.Timestamp(m.date).tz_localize(None) if m.date else None
                rows.append(dict(source=source_of(p), date=date, subject=re.sub(r"[\x00-\x1f]", "", m.subject or "").strip(),
                                 sender=people(m.sender), to=people(m.to), cc=people(m.cc)))
                m.close()
            except Exception:
                continue
    return rows


def build(rows):
    stat = defaultdict(lambda: dict(n=0, last=None, subject="", role=set(), name=""))
    for r in rows:
        if not r["source"]:
            continue
        for role, plist in (("发件", r["sender"]), ("收件", r["to"]), ("抄送", r["cc"])):
            for name, addr in plist:
                if GENERIC.match(addr) and role != "发件":
                    continue
                k = (r["source"], addr)
                s = stat[k]
                s["n"] += 1
                s["role"].add(role)
                if s["last"] is None or (r["date"] and r["date"] > s["last"]):
                    s["last"], s["subject"] = r["date"], r["subject"]
                if len(name) > len(s["name"]):
                    s["name"] = name
    out = []
    for (src, addr), s in stat.items():
        internal = addr.endswith("@xiaomi.com")
        side = "小米" if internal else "客户方"
        if not internal and "发件" in s["role"]:
            kind = "发数据（首选联系人）" if not GENERIC.match(addr) else "系统 / 公共邮箱（不要直接回复）"
        elif not internal:
            kind = "客户方同事（抄送里常见）"
        else:
            kind = "小米对接人（收件）" if "收件" in s["role"] else "小米同事（抄送）"
        out.append(dict(来源=src, 哪一方=side, 角色=kind, 姓名=s["name"], 邮箱=addr, 出现次数=s["n"],
                        最近一次=s["last"].date() if s["last"] is not None else None, 最近邮件标题=s["subject"]))
    df = pd.DataFrame(out)
    order = {"发数据（首选联系人）": 0, "客户方同事（抄送里常见）": 1, "小米对接人（收件）": 2, "小米同事（抄送）": 3,
             "系统 / 公共邮箱（不要直接回复）": 4}
    df["_o"] = df["角色"].map(order)
    return df.sort_values(["来源", "_o", "最近一次", "出现次数"], ascending=[True, True, False, False]).drop(columns="_o")


def main():
    folders = sys.argv[1:] or [HISTORY] + sorted(p for p in ROOT.glob("W*") if p.is_dir())
    df = build(scan(folders))
    out = DATA / "联系人.xlsx"
    with pd.ExcelWriter(out) as xw:
        df.to_excel(xw, sheet_name="全部", index=False)
        first = df[df["哪一方"] == "客户方"].groupby("来源").head(3)
        inner = df[df["哪一方"] == "小米"].groupby("来源").head(4)
        pd.concat([first, inner]).sort_values("来源").to_excel(xw, sheet_name="每家先找谁", index=False)
    print(out, len(df), "条；来源：", ", ".join(sorted(df["来源"].unique())))


if __name__ == "__main__":
    main()
