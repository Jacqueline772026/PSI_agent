#!/bin/bash
# 生成 PSI 看板（Mac）。双击运行。
cd "$(dirname "$0")" || exit 1
if [ ! -x .venv/bin/python ]; then
  echo "还没安装，先自动跑一遍安装 ..."
  bash ./一键安装_Mac.command || exit 1
fi
if [ ! -f "数据/Poland PSI DB+Channel.xlsx" ]; then
  echo "找不到 数据/Poland PSI DB+Channel.xlsx。"
  echo "从飞书导出 PSI 表，改名为  Poland PSI DB+Channel.xlsx  放进 数据 文件夹后再双击。"
  mkdir -p 数据; open 数据; read -r -p "按回车关闭"; exit 1
fi
read -r -p "生成第几周的看板？输入数字，比如 47（直接回车 = 最近填好的一周）：" WEEK
ARGS=(-m tools.dashboard)
[[ "$WEEK" =~ ^[0-9]+$ ]] && ARGS+=(--week "$WEEK")
ERR="$(mktemp)"
echo "正在生成（大约 1 分钟）..."
if ! OUT="$(.venv/bin/python "${ARGS[@]}" 2>"$ERR")"; then
  echo "生成失败，详细信息："; cat "$ERR"; read -r -p "按回车关闭"; exit 1
fi
HTML="$(printf '%s\n' "$OUT" | tail -1)"
echo "生成好了：$HTML"
echo "看板里是 PSI 实数（L3 机密），只在公司渠道传。"
open "$HTML"
read -r -p "按回车关闭"
