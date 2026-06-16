@echo off
cd /d "%~dp0"

python --version >nul 2>nul
if errorlevel 1 (
  echo 未检测到 Python。请先安装 Python 3.9 或更高版本。
  pause
  exit /b 1
)

python -c "import pandas, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo 首次运行需要安装依赖，正在自动安装...
  python -m pip install -r requirements.txt
  if errorlevel 1 (
    echo 依赖安装失败。请检查网络或 Python/pip 是否可用。
    pause
    exit /b 1
  )
)

python -c "import socket; s=socket.socket(); raise SystemExit(0 if s.connect_ex(('127.0.0.1',8765))==0 else 1)" >nul 2>nul
if not errorlevel 1 (
  echo 本地工具已在运行，正在打开页面...
  start "" "http://127.0.0.1:8765"
  exit /b 0
)

start "Model Eval Tool" cmd /k "python app.py"
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765"
exit /b 0
