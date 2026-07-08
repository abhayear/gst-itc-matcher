@echo off
cd /d "%~dp0"
echo Starting GST ITC Matcher...
echo Open the URL shown below in your browser, upload 2 Excel files, and download the report.
python -m streamlit run app.py
pause
