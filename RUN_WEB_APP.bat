@echo off
title Lab Dashboard — Web App
color 1F
cd /d "%~dp0"

echo.
echo  ============================================================
echo   Lab Planning Dashboard — Web App
echo  ============================================================
echo.

:: Check if streamlit is installed
streamlit --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [INFO] Streamlit not found. Installing...
    pip install streamlit plotly pandas openpyxl xlrd
    echo.
)

echo  [STARTING] Opening web browser at http://localhost:8501
echo  [INFO] Press Ctrl+C in this window to stop the server.
echo.
streamlit run app.py --server.port 8501 --server.headless false
if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Failed to start. Make sure Python and Streamlit are installed.
    pause
)
