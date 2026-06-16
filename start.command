#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "未检测到 Python 3。请先安装 Python 3.9 或更高版本。"
  echo "按回车键关闭窗口。"
  read -r
  exit 1
fi

if ! python3 - <<'PY' >/dev/null 2>&1
import pandas
import openpyxl
PY
then
  echo "首次运行需要安装依赖，正在自动安装..."
  if ! python3 -m pip install -r requirements.txt; then
    echo "依赖安装失败。请检查网络或 Python/pip 是否可用。"
    echo "按回车键关闭窗口。"
    read -r
    exit 1
  fi
fi

if curl -fsS "http://127.0.0.1:8765" >/dev/null 2>&1; then
  echo "本地工具已在运行，正在打开页面..."
  open "http://127.0.0.1:8765" >/dev/null 2>&1 || true
  exit 0
fi

port="$(python3 - <<'PY'
import socket

for port in range(8765, 8786):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        try:
            sock.bind(("127.0.0.1", port))
        except OSError:
            continue
        print(port)
        break
else:
    raise SystemExit("没有找到可用端口，请关闭其他本地服务后重试。")
PY
)"

echo "正在启动本地工具：http://127.0.0.1:${port}"
PORT="$port" python3 app.py &
server_pid=$!

for _ in {1..30}; do
  if curl -fsS "http://127.0.0.1:${port}" >/dev/null 2>&1; then
    open "http://127.0.0.1:${port}" >/dev/null 2>&1 || true
    wait "$server_pid"
    exit $?
  fi
  sleep 0.3
done

echo "本地服务启动失败。"
echo "请检查："
echo "1. 是否已安装依赖：pip install -r requirements.txt"
echo "2. 是否已有其他程序占用 8765 端口"
echo "3. macOS 是否拦截了终端/脚本的本地网络权限"
set +e
wait "$server_pid"
echo "按回车键关闭窗口。"
read -r
