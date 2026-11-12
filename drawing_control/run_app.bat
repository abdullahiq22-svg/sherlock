@echo off
cd /d "%~dp0"
echo Installing requirements (first time only)...
py -m pip install -q -r requirements.txt || python -m pip install -q -r requirements.txt
echo Starting the app...
py -m streamlit run app.py || python -m streamlit run app.py
pause
