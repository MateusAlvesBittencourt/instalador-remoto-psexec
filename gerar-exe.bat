@echo off
setlocal
cd /d "%~dp0"

where py.exe >nul 2>nul
if errorlevel 1 (
  echo ERRO: Python 3 nao encontrado.
  echo Instale em https://www.python.org/downloads/windows/
  pause
  exit /b 1
)

py -3 -m pip show pyinstaller >nul 2>nul
if errorlevel 1 (
  echo Instalando PyInstaller para gerar o executavel...
  py -3 -m pip install --user pyinstaller
  if errorlevel 1 (
    echo Falha ao instalar o PyInstaller.
    pause
    exit /b 1
  )
)

py -3 -m PyInstaller --noconfirm --clean --onefile --windowed --uac-admin --name InstaladorRemoto --add-data "RemoteInstallerParallel.ps1;." InstaladorRemoto.py
if errorlevel 1 (
  echo Falha ao gerar o executavel.
  pause
  exit /b 1
)

echo.
echo Executavel criado em: dist\InstaladorRemoto.exe
pause
