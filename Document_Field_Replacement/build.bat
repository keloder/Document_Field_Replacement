@echo off
chcp 65001 >nul
title DocFieldReplacer - Build
echo ==========================================
echo   DocFieldReplacer - Build EXE
echo ==========================================
echo.

set PYTHON=D:\Conda\envs\py3.8\python.exe

if not exist "%PYTHON%" (
    echo Error: Python not found at %PYTHON%
    echo Please check the Python path in this script
    pause
    exit /b 1
)

:CHOICE
echo Choose build mode:
echo 1. onedir - Directory mode (multiple files, faster startup)
echo 2. onefile - Single file mode (single exe, slower startup)
echo.
set /p MODE="Enter choice (1 or 2): "

if "%MODE%"=="1" (
    set BUILD_MODE=onedir
    goto :BUILD
)
if "%MODE%"=="2" (
    set BUILD_MODE=onefile
    goto :BUILD
)
echo Invalid choice. Please enter 1 or 2.
echo.
goto :CHOICE

:BUILD
echo.
echo [1/4] Installing dependencies...
"%PYTHON%" -m pip install --no-user flask python-docx lxml "pyinstaller==5.13.2" "pyinstaller-hooks-contrib==2023.2" --force-reinstall --no-deps -q
"%PYTHON%" -m pip install --no-user psutil -q

echo [2/4] Cleaning old build...
if exist "dist" rd /s /q dist
if exist "build" rd /s /q build
del /q *.spec 2>nul

echo [3/4] Building EXE (%BUILD_MODE% mode, please wait)...
"%PYTHON%" build_exe.py %BUILD_MODE%

echo.
echo ==========================================

if "%MODE%"=="1" goto :CHECK_ONEDIR
goto :CHECK_ONEFILE

:CHECK_ONEDIR
if exist "dist\DocFieldReplacer\DocFieldReplacer.exe" goto :SUCCESS_ONEDIR
goto :FAIL

:CHECK_ONEFILE
if exist "dist\DocFieldReplacer.exe" goto :SUCCESS_ONEFILE
goto :FAIL

:SUCCESS_ONEDIR
echo   Build SUCCESS!
echo.
echo   Output: dist\DocFieldReplacer\
echo.
echo   Directory Structure:
echo   - DocFieldReplacer.exe (main program)
echo   - python3*.dll (Python runtime)
echo   - *.dll (dependency libraries)
echo   - base_library.zip (Python standard library)
echo   - *.pyd (Python extension modules)
echo.
echo   Usage:
echo   1. Double-click DocFieldReplacer.exe
echo   2. Browser opens at http://127.0.0.1:5000
echo   3. Enjoy!
echo ==========================================
explorer dist
goto :END

:SUCCESS_ONEFILE
echo   Build SUCCESS!
echo.
echo   Output: dist\DocFieldReplacer.exe
echo.
echo   Single executable file
echo.
echo   Usage:
echo   1. Double-click DocFieldReplacer.exe
echo   2. Browser opens at http://127.0.0.1:5000
echo   3. Enjoy!
echo ==========================================
explorer dist
goto :END

:FAIL
echo   Build FAILED. Output file not found.
echo ==========================================
goto :END

:END
pause