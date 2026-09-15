@echo off
setlocal
cd /d "%~dp0"

where py.exe >nul 2>nul
if not errorlevel 1 (
  py -3 InstaladorRemoto.py
  exit /b %errorlevel%
)

where python.exe >nul 2>nul
if not errorlevel 1 (
  python InstaladorRemoto.py
  exit /b %errorlevel%
)

echo.
echo ERRO: Python 3 nao foi encontrado.
echo Instale o Python para Windows marcando "Add Python to PATH".
echo https://www.python.org/downloads/windows/
echo.
pause
exit /b 1
