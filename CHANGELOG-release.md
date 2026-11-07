## Windows App Updater v2.3

### ✨ New features

- **Live download percentage** — the download bar now shows real progress (5%, 12%, 30% …) as it happens, instead of jumping straight from 0% to 100%.
- **Separate Download and Install phases** — the progress bar switches from *Downloading* to *Installing* when winget moves to the install step, so you can see which phase an app is in. The install phase shows an animated bar since most installers don't report a percentage.
- **Per-app "close running app" prompt** — when an app is open and blocking its own update, you now get a dialog naming the app (*"Updating AppName — close it?"*) with **Yes / No** buttons and a **"Remember my choice for the rest of this batch"** checkbox, instead of having to find and close the app yourself.
- **"Auto-close blocking apps" option** — a checkbox to force-close blocking apps automatically with no prompts (confirmed once up front), for a hands-off run.
- **Stall timeout** — if an app hangs with no activity for 3 minutes, it's automatically aborted and marked **⚠ Stalled** so the batch continues instead of freezing. Slow-but-active downloads are never affected.

### 🔧 Improvements

- **Accurate success/failure reporting** — results are now decided by winget's actual exit code instead of scanning its text output. Fixes apps being wrongly marked failed when their output merely contained words like *"error"* or *"0x"*, and correctly distinguishes *"no update needed"* and *"app in use"* from real failures.
- **Skip button fixed** — pressing **Skip** now correctly marks the app as **⏭ Skipped** instead of sometimes showing **✅ Success**.
- **Filter out fake updates** — apps showing a pointless update where the versions are actually the same (e.g. `2.4.0-release` vs `2.4.0`) are now hidden.
- **Equal-size progress bars** — the two progress bars are now exactly the same width.
- **Responsive Skip and Cancel** — both now respond immediately even while an app is mid-extraction, instead of waiting for the next output line.

### 🐛 Fixes

- Declining the UAC prompt no longer silently closes the app — it now stays running without admin.
- Faster temp-folder scanning during updates (scoped to winget's own folder instead of all of `%TEMP%`).
- Cleaner output log — winget's progress-bar fragments no longer flood it.
- Removed dead/duplicate code; future-proofed temp cleanup for newer Python versions.