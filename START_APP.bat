@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "app.py" (
    echo [ERROR] app.py was not found beside this BAT file.
    echo Extract the FULL ZIP archive and run this file inside pnu_auto_app.
    pause
    exit /b 1
)

set "VENV_PY=.venv\Scripts\python.exe"

if not exist "%VENV_PY%" (
    echo [SETUP] Preparing a Python environment for this PC...
    py -3.13 --version >nul 2>&1
    if not errorlevel 1 (
        py -3.13 -m venv .venv
    ) else (
        py -3.12 --version >nul 2>&1
        if not errorlevel 1 (
            py -3.12 -m venv .venv
        ) else (
            python --version >nul 2>&1
            if not errorlevel 1 (
                python -m venv .venv
            ) else (
                echo [ERROR] Python 3.12 or 3.13 is not installed or not on PATH.
                echo Install Python, then start this BAT again.
                pause
                exit /b 1
            )
        )
    )
    if errorlevel 1 (
        echo [ERROR] Creating the environment failed.
        pause
        exit /b 1
    )
)

if not exist "%VENV_PY%" (
    echo [ERROR] Cannot find the Python executable inside .venv.
    pause
    exit /b 1
)

"%VENV_PY%" -c "import sys" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] The existing .venv is invalid, possibly copied from another PC.
    echo Close this window, rename the .venv folder to .venv_old, and restart.
    pause
    exit /b 1
)

"%VENV_PY%" -c "import streamlit, openpyxl, geopandas, pyogrio, shapely" >nul 2>&1
if errorlevel 1 (
    echo [SETUP] Installing required packages. This may take several minutes.
    "%VENV_PY%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Package installation failed. Check the error above.
        echo Internet access may be required for first-time installation.
        pause
        exit /b 1
    )
)

echo [START] Opening the environmental charge calculation app...
"%VENV_PY%" -m streamlit run app.py
if errorlevel 1 echo [ERROR] The app stopped. Review the messages above.
pause
