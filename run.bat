@echo off
rem First run installs dependencies into a local venv, then starts Windows App Updater.
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" (
    echo Setting up Windows App Updater for the first time...
    py -3 -m venv .venv || python -m venv .venv
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt || (pause & exit /b 1)
)
start "" ".venv\Scripts\pythonw.exe" "%~dp0App-Updater.py" %*
