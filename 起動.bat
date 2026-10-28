@echo off
cd /d "%~dp0"
python --version >/dev/null 2>&1
if %errorlevel% neq 0 (echo Python not found. & pause & exit /b 1)
python -c "import pdfplumber" >/dev/null 2>&1
if %errorlevel% neq 0 (echo Installing pdfplumber... & python -m pip install pdfplumber)
python -c "import reportlab" >/dev/null 2>&1
if %errorlevel% neq 0 (echo Installing reportlab... & python -m pip install reportlab)
python -c "import win32com.client" >/dev/null 2>&1
if %errorlevel% neq 0 (echo Installing pywin32... & python -m pip install pywin32)
python -c "import google.auth" >/dev/null 2>&1
if %errorlevel% neq 0 (echo Installing google-auth... & python -m pip install google-auth google-auth-oauthlib google-api-python-client)
start "" pythonw onishi_schedule_app.py
