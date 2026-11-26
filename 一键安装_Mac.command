#!/bin/bash
# PSI agent 一键安装（Mac）。双击运行；第一次如果提示“无法打开”，右键 → 打开。
# 重复运行没关系，已经装好的步骤会跳过。
cd "$(dirname "$0")" || exit 1
ROOT="$(pwd)"
fail() { printf '\n\033[31m[没装好] %s\n把这个窗口截图发给负责人。\033[0m\n' "$1"; read -r -p "按回车关闭"; exit 1; }
ok() { printf '\033[32m      %s\033[0m\n' "$1"; }

echo "== PSI agent 一键安装 =="
echo "安装位置：$ROOT"; echo

echo "[1/4] 检查 Python ..."
CHECK='import sys; assert sys.version_info >= (3, 11)'
if [ -x .venv/bin/python ] && ! .venv/bin/python -c "$CHECK" 2>/dev/null; then
  echo "      旧的 .venv 不能用，重建。"; rm -rf .venv
fi
if [ ! -x .venv/bin/python ]; then
  PY=""
  for c in python3.13 python3.12 python3.11 python3 /opt/homebrew/bin/python3 /usr/local/bin/python3; do
    if command -v "$c" >/dev/null 2>&1 && "$c" -c "$CHECK" 2>/dev/null; then PY="$(command -v "$c")"; break; fi
  done
  if [ -z "$PY" ]; then
    open "https://www.python.org/downloads/macos/"
    fail "没找到 Python 3.11 以上。已打开下载页：装 Python 3.13 的 macOS installer，装完再双击一次 一键安装_Mac.command。"
  fi
  ok "用这个 Python：$PY"
else
  ok ".venv 已经有了，跳过。"
fi

echo "[2/4] 建虚拟环境 .venv ..."
[ -x .venv/bin/python ] || "$PY" -m venv .venv || fail "建 .venv 失败。"
ok "好了。"

echo "[3/4] 安装依赖（第一次大约 1–3 分钟）..."
.venv/bin/python -m pip install --upgrade pip -q 2>/dev/null
if ! .venv/bin/python -m pip install -r requirements.txt -q; then
  echo "      官方源没装上，换清华镜像再试 ..."
  .venv/bin/python -m pip install -r requirements.txt -q -i https://pypi.tuna.tsinghua.edu.cn/simple || fail "依赖没装上，多半是网络问题。"
fi
.venv/bin/python -c 'import pandas, openpyxl, yaml, xlrd' || fail "依赖装完但导入失败。"
ok "好了。"

echo "[4/4] 准备 数据 文件夹 ..."
mkdir -p 数据
echo; printf '\033[32m== 安装完成 ==\033[0m\n'
if [ -f "数据/Poland PSI DB+Channel.xlsx" ]; then
  echo "数据/Poland PSI DB+Channel.xlsx 已经在了，可以直接双击 生成看板_Mac.command。"
else
  echo "还差一步：从飞书导出 PSI 表，改名为  Poland PSI DB+Channel.xlsx  放进刚打开的 数据 文件夹。"
  echo "然后双击 生成看板_Mac.command 就能出看板。"
  open 数据
fi
read -r -p "按回车关闭"
