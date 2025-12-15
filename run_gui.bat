@echo off
setlocal
title 中文合同脱敏工具

cd /d "%~dp0"

rem 优先使用 venv，再回落到系统 Python
set "PY_EXE=python"
if exist "venv\Scripts\python.exe" set "PY_EXE=venv\Scripts\python.exe"

%PY_EXE% -V >nul 2>nul
if errorlevel 1 (
  echo 未检测到 Python，可先从 https://www.python.org/ 下载並勾選「Add python.exe to PATH」。
  pause
  exit /b 1
)

echo 正在启动图形界面（请保持联网或已安装好依赖 tkinterdnd2、python-docx）...
%PY_EXE% redact_gui.py

if errorlevel 1 (
  echo 如提示缺少依赖，可在此目录 Shift+右鍵「在此处打开 PowerShell」，再运行：pip install -r requirements.txt
)

pause
