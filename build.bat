@echo off
rem Builds dist\Windows-App-Updater.exe (the same way the GitHub release workflow does)
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    py -3 -m venv .venv || python -m venv .venv
)
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller || (pause & exit /b 1)
".venv\Scripts\pyinstaller.exe" --noconfirm --onefile --windowed --name "Windows-App-Updater" ^
  --icon "windows-updater.ico" --add-data "success.wav;." --add-data "kuwait.png;." --add-data "windows-updater.ico;." ^
  App-Updater.py || (pause & exit /b 1)
echo.
echo Built: dist\Windows-App-Updater.exe
pause
