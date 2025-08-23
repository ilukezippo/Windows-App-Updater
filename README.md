# Windows App Updater

GUI (Tkinter) to check and update installed apps via **winget**.

## Run
- Double-click \App-Updater.pyw\ (no console), or:
\\\powershell
pythonw App-Updater.pyw
\\\

## Build EXE (optional)
\\\powershell
pyinstaller --onefile --windowed --icon=windows-updater.ico App-Updater.py
\\\

## Notes
- Turn on **Include unknown apps** to scan with \--include-unknown\.
- Click **Run as Admin** for silent installs.
- Plays a success sound after updates.
- Made by **BoYaqoub** – ilukezippo@gmail.com
