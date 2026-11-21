# Windows App Updater privacy policy

_Last updated: November 2026 (Windows App Updater 2.4)_

Windows App Updater collects **no personal data**, no usage statistics and no telemetry. It has no accounts.

It only goes online for these things:

| What | When | What is contacted | What is sent |
|---|---|---|---|
| **Checking for and installing app updates** | When you click *Check for updates* or *Update selected* | Microsoft's **winget**, which talks to its own sources (the winget community repository and the Microsoft Store) and downloads each installer from the app publisher's own website | What winget normally sends. Windows App Updater adds nothing |
| **Checking for a new Windows App Updater** | Once when it starts, and when you click *Check for update* in About | `api.github.com` (this project's releases) | A standard HTTPS request. No files or personal data |
| **Downloading a new Windows App Updater** | Only when you agree to update | GitHub Releases of this project | — |
| **Buy me a coffee / contact / website** | Only when you click them | Your browser or mail app opens the link | — |

## Files on your computer

- Your choices (theme, switches, whether the log is shown) are saved in `%APPDATA%\WindowsAppUpdater\settings.json`.
- The activity log stays in the window; it is only saved when you click *Export*, to a file you choose.
- *Clean temp files* only removes leftover installers in winget's own temporary folder.

## Contact

Questions: open an issue at https://github.com/ilukezippo/Windows-App-Updater/issues
