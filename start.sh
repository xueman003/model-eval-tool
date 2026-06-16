#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "未检测到 Python 3。请先安装 Python 3.9 或更高版本。"
  exit 1
fi

if ! python3 - <<'PY' >/dev/null 2>&1
import pandas
import openpyxl
PY
then
  echo "首次运行需要安装依赖，正在自动安装..."
  python3 -m pip install -r requirements.txt
fi

python3 app.py
