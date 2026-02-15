@echo off
setlocal enabledelayedexpansion

set PYTHON_DIR=C:\Users\Administrator\AppData\Local\Programs\Python\Python312\
set TCL_LIBRARY=%PYTHON_DIR%tcl\tcl8.6
set TK_LIBRARY=%PYTHON_DIR%tcl\tk8.6

set "APP_DIR=%~dp0"
cd /d "%APP_DIR%"

echo ====================================
echo Building Prescription System
echo ====================================
echo.

echo [1/4] Installing dependencies...
%PYTHON_DIR%python.exe -m pip install --quiet pyinstaller pywin32 pillow
echo Done.
echo.

echo [2/4] Checking icon file...
if not exist "xw.ico" (
    echo WARNING: xw.ico NOT FOUND! Building without icon.
    set "ICON_PARAM="
) else (
    echo Found xw.ico
    set "ICON_PARAM=--icon=xw.ico"
)
echo.

echo [3/4] Cleaning old build...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
if exist *.spec del /q *.spec
echo Done.
echo.

echo [4/4] Building executable...
%PYTHON_DIR%python.exe -m PyInstaller --onefile --windowed --name "PrescriptionSystem" --collect-all tkinter %ICON_PARAM% main.py
echo.

if exist build rmdir /s /q build
if exist __pycache__ rmdir /s /q __pycache__
if exist *.spec del /q *.spec

if exist "dist\PrescriptionSystem.exe" (
    echo.
    echo ====================================
    echo BUILD SUCCESS!
    echo File: %APP_DIR%dist\PrescriptionSystem.exe
    echo ====================================
) else (
    echo.
    echo ====================================
    echo BUILD FAILED!
    echo ====================================
)
pause
