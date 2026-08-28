@echo off
setlocal EnableExtensions
title Claviculario ECL v0.9 - Gerar Instalador
cd /d "%~dp0"

echo ============================================================
echo CLAVICULARIO ECL v0.9 - GERADOR DO INSTALADOR
echo ============================================================
echo.

set "PYEXE="

if exist "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe" (
  "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe" -m pip --version >nul 2>&1
  if not errorlevel 1 set "PYEXE=%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe"
)
if not defined PYEXE if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
  "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" -m pip --version >nul 2>&1
  if not errorlevel 1 set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
)
if not defined PYEXE (
  python -m pip --version >nul 2>&1
  if not errorlevel 1 for /f "delims=" %%P in ('where python 2^>nul') do if not defined PYEXE set "PYEXE=%%P"
)

if not defined PYEXE (
  echo ERRO: Nao encontrei Python funcional com PIP neste computador.
  pause
  exit /b 1
)

echo Python encontrado:
echo "%PYEXE%"
echo.

"%PYEXE%" -m pip install -r requirements.txt
"%PYEXE%" -m pip install pyinstaller
if errorlevel 1 goto :erro

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist installer_build rmdir /s /q installer_build
if exist installer_dist rmdir /s /q installer_dist
if exist INSTALADOR_PRONTO rmdir /s /q INSTALADOR_PRONTO
for %%F in (*.spec) do del /q "%%F"

echo [1/4] Gerando programa...
"%PYEXE%" -m PyInstaller --noconfirm --clean --onedir --windowed ^
 --name "Claviculario_ECL" ^
 --icon "%~dp0icone.ico" ^
 --collect-all reportlab ^
 --collect-all PIL ^
 app.py
if errorlevel 1 goto :erro

echo [2/4] Montando estrutura...
if exist assets xcopy "assets" "dist\Claviculario_ECL\assets\" /E /I /Y >nul
for %%D in (data backups protocolos relatorios logs) do if not exist "dist\Claviculario_ECL\%%D" mkdir "dist\Claviculario_ECL\%%D"
if exist "data\Claviculario_ECL.xlsx" copy /Y "data\Claviculario_ECL.xlsx" "dist\Claviculario_ECL\data\Claviculario_ECL.xlsx" >nul

echo [3/4] Gerando desinstalador...
"%PYEXE%" -m PyInstaller --noconfirm --clean --onefile --windowed --uac-admin ^
 --name "Desinstalar_Claviculario_ECL" ^
 --icon "%~dp0icone.ico" ^
 --distpath "dist\Claviculario_ECL" ^
 uninstaller_builder.py
if errorlevel 1 goto :erro

echo [4/4] Gerando Setup profissional...
"%PYEXE%" -m PyInstaller --noconfirm --clean --onefile --windowed --uac-admin ^
 --name "Setup_Claviculario_ECL" ^
 --icon "%~dp0icone.ico" ^
 --distpath "installer_dist" ^
 --workpath "installer_build" ^
 --add-data "dist\Claviculario_ECL;payload" ^
 installer_builder.py
if errorlevel 1 goto :erro

mkdir INSTALADOR_PRONTO >nul 2>&1
copy /Y "installer_dist\Setup_Claviculario_ECL.exe" "INSTALADOR_PRONTO\Setup_Claviculario_ECL.exe" >nul

echo.
echo ============================================================
echo INSTALADOR GERADO COM SUCESSO!
echo ============================================================
echo.
echo Coloque no pendrive:
echo %~dp0INSTALADOR_PRONTO\Setup_Claviculario_ECL.exe
echo.
echo Ele instala em C:\Claviculario_ECL
echo e aparece em Aplicativos instalados do Windows.
echo ============================================================
pause
exit /b 0

:erro
echo.
echo ERRO AO GERAR. Tire uma foto das ultimas linhas.
pause
exit /b 1
