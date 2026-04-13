@echo off
cd /d "%~dp0\.."
python -m streamlit run ".\streamlit_ui\app.py" --server.headless=true --server.port=8501 --browser.gatherUsageStats=false
pause
