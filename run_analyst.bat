@echo off
setlocal EnableDelayedExpansion
title Energy Data Analyst
color 0A

echo.
echo ================================================================
echo          Energy Data Analyst  -  Launcher
echo ================================================================
echo   Setting up environment, please wait...
echo   (First launch may take a few minutes)
echo ================================================================
echo.

:: Change to the script's own directory
cd /d "%~dp0"

:: ================================================================
:: STEP 1 — Locate or install uv (Python package manager)
:: ================================================================
set "UV_CMD="

:: Check if uv is on PATH
where uv >nul 2>&1
if %ERRORLEVEL%==0 (
    set "UV_CMD=uv"
    echo [OK] uv found on PATH.
    goto :HAS_UV
)

:: Check common install locations
if exist "%USERPROFILE%\.local\bin\uv.exe" (
    set "UV_CMD=%USERPROFILE%\.local\bin\uv.exe"
    echo [OK] uv found at %UV_CMD%.
    goto :HAS_UV
)
if exist "%CARGO_HOME%\bin\uv.exe" (
    set "UV_CMD=%CARGO_HOME%\bin\uv.exe"
    echo [OK] uv found at %UV_CMD%.
    goto :HAS_UV
)
if exist "%USERPROFILE%\.cargo\bin\uv.exe" (
    set "UV_CMD=%USERPROFILE%\.cargo\bin\uv.exe"
    echo [OK] uv found at %UV_CMD%.
    goto :HAS_UV
)

:: uv not found — install it automatically
echo [..] uv not found. Installing uv package manager...
echo     (This is a one-time setup)
echo.
powershell -ExecutionPolicy ByPass -NoProfile -Command "irm https://astral.sh/uv/install.ps1 | iex" 2>nul

:: After install, try to find uv again
where uv >nul 2>&1
if %ERRORLEVEL%==0 (
    set "UV_CMD=uv"
    echo [OK] uv installed successfully.
    goto :HAS_UV
)
if exist "%USERPROFILE%\.local\bin\uv.exe" (
    set "UV_CMD=%USERPROFILE%\.local\bin\uv.exe"
    echo [OK] uv installed successfully.
    goto :HAS_UV
)
if exist "%CARGO_HOME%\bin\uv.exe" (
    set "UV_CMD=%CARGO_HOME%\bin\uv.exe"
    echo [OK] uv installed successfully.
    goto :HAS_UV
)
if exist "%USERPROFILE%\.cargo\bin\uv.exe" (
    set "UV_CMD=%USERPROFILE%\.cargo\bin\uv.exe"
    echo [OK] uv installed successfully.
    goto :HAS_UV
)

:: Still not found — give up
echo.
echo [ERROR] Could not install uv automatically.
echo         Please ask your IT administrator for help, or
echo         visit https://docs.astral.sh/uv/ to install manually.
echo.
pause
exit /b 1

:: ================================================================
:: STEP 2 — Install Python (if not already available via uv)
:: ================================================================
:HAS_UV
echo.
echo [..] Ensuring Python 3.13 is available...
"%UV_CMD%" python install 3.13 --quiet 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [..] Retrying Python install...
    "%UV_CMD%" python install 3.13
)
echo [OK] Python is ready.

:: ================================================================
:: STEP 3 — Install / sync project dependencies
:: ================================================================
echo.
echo [..] Installing dependencies (first run may take a moment)...
"%UV_CMD%" sync --quiet
if %ERRORLEVEL% NEQ 0 (
    echo [WARN] uv sync had issues, retrying...
    "%UV_CMD%" sync
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ERROR] Could not install dependencies.
        echo         Please check your internet connection and try again.
        echo.
        pause
        exit /b 1
    )
)
echo [OK] All dependencies ready.

:: ================================================================
:: STEP 4 — Launch the application
:: ================================================================
echo.
echo ================================================================
echo   Starting the Analyst application...
echo   A browser window will open automatically.
echo.
echo   >> CLOSE THIS WINDOW to stop the application.
echo ================================================================
echo.

"%UV_CMD%" run python run_analyst.py

echo.
echo ================================================================
echo   Application stopped.
echo ================================================================
pause
