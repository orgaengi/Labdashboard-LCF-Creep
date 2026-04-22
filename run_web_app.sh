#!/bin/bash
# Lab Planning Dashboard — Web App Launcher
cd "$(dirname "$0")"

echo ""
echo " ============================================================"
echo "  Lab Planning Dashboard — Web App"
echo " ============================================================"
echo ""

# Check streamlit
if ! command -v streamlit &>/dev/null; then
    echo " [INFO] Installing Streamlit..."
    pip3 install streamlit plotly pandas openpyxl xlrd
fi

echo " [STARTING] Opening http://localhost:8501"
echo " [INFO] Press Ctrl+C to stop."
echo ""
streamlit run app.py --server.port 8501
