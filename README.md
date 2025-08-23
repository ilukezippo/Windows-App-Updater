# Windows App Updater

GUI tool (Tkinter) to check and update installed apps via **winget**.

## Features
- Flat list with checkboxes (☐ / ☑), Select All / Select None
- **Run as Admin** button (relaunches elevated for silent installs)
- **Include unknown apps** toggle (`--include-unknown`)
- **Loading screen** while scanning
- **Cancel** during updates
- **YIFY-style** single progress bar with counts
- Live log with both vertical & horizontal scrollbars
- App icon + Kuwait flag branding
- Plays a success sound when all updates finish

## Requirements
- Windows 10/11 with **winget** (App Installer) available
- Python 3.10+ (tested with 3.11/3.12/3.13)
- Optional: `Pillow` if you want to load `kuwait.ico` without converting to PNG

```bash
pip install pillow
