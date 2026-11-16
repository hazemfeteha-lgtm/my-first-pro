@echo off
cd /d "%~dp0"
if not exist center_management.py (
  echo center_management.py was not found.
  echo Extract the ZIP first, then run this file from the extracted folder.
  pause
  exit /b
)
set PY=
py -3 --version >nul 2>nul && set PY=py -3
if not defined PY python --version >nul 2>nul && set PY=python
if not defined PY (
  echo Python is not installed.
  echo Download it from https://www.python.org/downloads/
  echo and tick "Add Python to PATH" during setup, then run this file again.
  pause
  exit /b
)
%PY% center_management.py
if errorlevel 1 (
  echo.
  echo The program stopped with an error. Send a photo of the message above.
  pause
)
