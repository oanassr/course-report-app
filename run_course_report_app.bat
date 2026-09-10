@echo off
setlocal
set PYTHONIOENCODING=utf-8
set "BASE=%~dp0"
set "BOOTSTRAP=C:\Users\MSI\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "PYTHON=%BASE%.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
  echo Installing local dependencies for the first run...
  "%BOOTSTRAP%" -m venv "%BASE%.venv"
  if errorlevel 1 exit /b 1
  "%PYTHON%" -m pip install -r "%BASE%requirements.txt"
  if errorlevel 1 exit /b 1
  "%PYTHON%" "%BASE%course_report_app\preload_ocr.py"
  if errorlevel 1 exit /b 1
)

"%PYTHON%" "%BASE%course_report_app\server.py"
