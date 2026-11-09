"""发现原始数据有问题时，生成"待确认邮件"：该找谁 + 英文邮件草稿 + 中文说明。

只生成草稿，不会发送。发外部邮件前请先给导师 / 这家客户的小米对接人过目。
联系人来自 数据/联系人.xlsx（tools.contacts 从周报邮件原件里整理），没有这个文件时只给邮件正文。

目前覆盖：
  - 当周文件没收到（客户的周报 / 内部同事的表）
  - IM 的 Mobility 和 IMEI List 对不上
以后发现新的问题类型，在 TEMPLATES 里加一种即可。
"""
from pathlib import Path

import pandas as pd

from psi_tool.paths import DATA
from psi_tool.sources import week_range

CONTACTS = DATA / "联系人.xlsx"
MENTOR = "导师"  # 具体联系人只放本地的 数据/联系人.xlsx，不写进代码

# 来源 → (联系人表里的来源名, 文件说明, 是否内部同事的表)
SOURCES = {
    "EURO_TEL": ("EURO", "phone report (Xiaomi_tele_W{week})", False),
    "EURO_TAB": ("EURO", "tablet report (Xiaomi_TAB_W{week})", False),
    "MEX": ("MEX", "TERG stock & sales report", False),
    "MSH": ("MSH", "weekly sales & stock report (Raport Xiaomi W{week})", False),
    "X-KOM": ("X-KOM", "weekly sales & stock report (XIAOMI - SMARTFONY / TELEFONY X-KOM raport W{week})", False),
    "KTR": ("KTR", "weekly report (Raport Xiaomi 2026.Q4, column stan W{week})", False),
    "MAX": ("MAX", "weekly report (Xiaomi_{week}_2026_GT_Group)", False),
    "AUCHAN": ("Auchan", "weekly stock & sales report (Xiaomi T{week})", False),
    "MICHAIN_SO": ("GG", "SO_collect / Stock_collect files", False),
    "PLAY": ("Play", "Stock and sales - week {week}", False),
    "TMO": ("TMO", "Weekly CPFR Data (kw_cpfr_main_Xiaomi)", False),
    "PLUS": ("Plus", "Stany magazynowe W{week}", False),
    "ORANGE": ("Orange", "Zapas oraz dystrybucja W{week}", False),
    "ORANGE_SO": (None, "Jet 的 OPL_Stock_Sales 表（Orange SO）", True),
}


def contacts():
    try:
        return pd.read_excel(CONTACTS, sheet_name="全部")
    except (FileNotFoundError, ValueError):
        return None


def contact_lines(book, source):
    if book is None or source is None:
        return ["- 联系人：没有 数据/联系人.xlsx，先运行 `.venv/bin/python -m tools.contacts`"]
    d = book[book["来源"] == source]
    pick = lambda role, n: d[d["角色"] == role].sort_values("最近一次", ascending=False).head(n)
    rows = [("收件（客户方，发数据的人）", pick("发数据（首选联系人）", 2)),
            ("抄送（客户方同事）", pick("客户方同事（抄送里常见）", 1)),
            ("抄送（小米对接人）", pick("小米对接人（收件）", 3))]
    out = []
    for label, df in rows:
        if len(df):
            out.append(f"- {label}：" + "；".join(f"{n} <{e}>（最近 {str(t)[:10]}）" for n, e, t in df[["姓名", "邮箱", "最近一次"]].values))
    return out or [f"- 联系人表里没有 {source} 的记录"]


def missing_file(week, source, book):
    target, what, internal = SOURCES.get(source, (None, source, False))
    what = what.replace("{week}", str(week))
    start, end = week_range(week)
    head = [f"### {source}：W{week} 文件没收到", ""]
    if internal:
        return head + [f"内部同事的表，直接在飞书 / 企业微信问即可：", "",
                       f"> 你好，W{week}（{start:%m/%d}–{end:%m/%d}）的 {what} 方便发我一下吗？PSI 这周要用，谢谢！", ""]
    return head + contact_lines(book, target) + ["", "**建议做法**：在这家上周的周报邮件里点 Reply All 发（对方有上下文）。", "",
        "```text",
        f"Subject: RE: W{week} report – {what}",
        "",
        "Hi [Name],",
        "",
        f"Hope you are well. We have not yet received the W{week} ({start:%d.%m}–{end:%d.%m.%Y}) {what}.",
        "Could you please send it when you have a moment? We need it for this week's sell-out and stock update.",
        "If it has already been sent, please let me know which mailbox it went to and I will check on my side.",
        "",
        "Thank you in advance!",
        "",
        "Best regards,",
        "Jiaqi Hu",
        "GTM Intern, Xiaomi Poland",
        "```",
        "", f"中文意思：W{week} 的 {what} 还没收到，麻烦发一下；如果已经发了，请告诉我发到了哪个邮箱。", ""]


def im_mismatch(week, diffs, book, marked=()):
    start, end = week_range(week)
    lines = [f"### IM：Mobility 和 IMEI List 对不上（{len(diffs)} 处）", "",
             "PSI 按 SOP 用 Mobility 填，这一步不用改 PSI；请 IM 确认哪份是对的，对方确认后再按需要调整。", ""]
    lines += contact_lines(book, "IM")
    rows = "\n".join(f"| {d['model']} | {d['ean']} | {d['customer']} | {', '.join(d['invoices']) or '—'} | "
                     f"{', '.join(d['days'])} | {d['mobility']} | {d['imei']} |" for d in diffs)
    mob, imei = diffs[0]["files"]
    if marked:
        lines += ["", "**附件 / 截图**：不一致的行已在副本里标红，第一页 Difference 是英文汇总，可以直接截图或作为附件："]
        lines += [f"- `{Path(m).name}`" for m in marked]
    lines += ["", "**建议做法**：回复 Ingram 这周发 Mobility / IMEI List 的那封邮件（Reply All），先发给导师或 IM 的小米对接人"
              "（见上面'抄送（小米对接人）'）看一眼。", "",
              "```text",
              f"Subject: RE: W{week} – quantity difference between Mobility sales report and IMEI list",
              "",
              "Hi [Name],",
              "",
              f"Thank you for the W{week} reports. While reconciling them, we noticed that the sell-through quantity in "
              f"\"{mob}\" does not match the serial-number list in \"{imei}\" for the line(s) below:",
              "",
              "| Product | EAN | Customer | Invoice | Invoice date | Mobility | IMEI list |",
              "|---|---|---|---|---|---|---|",
              rows,
              "",
              "Could you please help us clarify:",
              "1. Which quantity is correct for W" + f"{week} ({start:%d.%m}–{end:%d.%m.%Y})?",
              "2. If the difference is due to timing (e.g. shipped but invoiced later, or credit notes), "
              "in which week will it appear in the Mobility report?",
              "",
              "We currently use the Mobility report for our weekly sell-through figures, so a short confirmation would help us "
              "keep both reports consistent.",
              "",
              *(["For reference, I have attached both files with the relevant rows highlighted in red "
                 "(summary on the 'Difference' sheet).", ""] if marked else []),
              "Many thanks,",
              "Jiaqi Hu",
              "GTM Intern, Xiaomi Poland",
              "```", "",
              "中文意思：两份表在这几行数量不一致，请确认 W" + f"{week} 以哪个为准；如果是发货和开票时间差，差额会在哪一周的 Mobility 里体现。"
              "我们目前按 Mobility 填周报，确认后两边口径就一致了。", ""]
    return lines


def write(path, week, missing, im_diffs, marked=()):
    book = contacts()
    L = [f"# W{week} 待确认邮件（草稿）", "",
         "工具发现的、要找客户或同事确认的问题。**只是草稿，不会自动发送。**", "",
         f"- 发外部邮件前，先给{MENTOR}或这家的小米对接人过目，尤其是第一次联系的客户。",
         "- 尽量在对方上一封周报邮件里 Reply All，保留对方的上下文和抄送人。",
         "- 语气：先感谢、只说事实和数字、问清楚'以哪个为准 / 什么时候补'，不要写成指责对方出错。", ""]
    if not missing and not im_diffs:
        L += ["这周没有需要对外确认的问题。"]
    for s in missing:
        L += missing_file(week, s, book)
    if im_diffs:
        L += im_mismatch(week, im_diffs, book, marked)
    Path(path).write_text("\n".join(L), encoding="utf-8")
