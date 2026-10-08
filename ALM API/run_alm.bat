@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 goto python_missing

rem Reuse the existing app without PowerShell's Invoke-WebRequest security prompt.
py -3 -c "import sys,urllib.request; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:18766/_stcore/health',timeout=1).read().decode().strip()=='ok' else 1)" >nul 2>nul
if not errorlevel 1 goto app_already_running

rem Prefer an already installed system Python so a network install is not required unnecessarily.
py -3 -c "import streamlit, openai, openpyxl, dotenv" >nul 2>nul
if not errorlevel 1 goto run_system_python

".venv\Scripts\python.exe" -c "import streamlit, openai, openpyxl, dotenv" >nul 2>nul
if not errorlevel 1 goto run_venv

if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
if errorlevel 1 goto startup_error

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto startup_error

:run_venv
set "APP_PYTHON=.venv\Scripts\python.exe"
goto launch_app

:run_system_python
set "APP_PYTHON=py -3"
goto launch_app

:launch_app
if not exist ".env" echo .env not found. Copy .env.example to .env and set OPENAI_API_KEY before using the app.
echo Starting the Complete Sales Diagnosis Streamlit app at http://127.0.0.1:18766
%APP_PYTHON% -m streamlit run streamlit_app.py --server.address 127.0.0.1 --server.port 18766 --server.maxUploadSize 25 --browser.gatherUsageStats false
if errorlevel 1 goto startup_error
exit /b 0

:app_already_running
echo Streamlit is already running at http://127.0.0.1:18766
start "" "http://127.0.0.1:18766"
exit /b 0

:python_missing
echo Python Launcher (py) was not found. Install Python 3.10 or later first.
goto wait_to_close

:startup_error
echo.
echo Startup failed. Read the error message above and keep this window open for details.

:wait_to_close
pause
exit /b 1
