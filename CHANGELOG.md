# Changelog

## 2.4 — November 2026

### New
- **A new, modern look** with **light and dark themes**: it follows Windows' setting, and the ☀ / ☾ button switches between them
- Clearer layout: the update list, the two progress bars and the activity log each sit in their own panel

### Fixed
- Bug fixes across checking for and installing updates

## 2.3 — June 2026

### New

- **Live download percentage** — the download bar now shows real progress (5%, 12%, 30% …) as it happens, instead of jumping straight from 0% to 100%.
- **Separate Download and Install phases** — the progress bar switches from *Downloading* to *Installing* when winget moves to the install step, so you can see which phase an app is in. The install phase shows an animated bar since most installers don't report a percentage.
- **Per-app "close running app" prompt** — when an app is open and blocking its own update, you now get a dialog naming the app (*"Updating AppName — close it?"*) with **Yes / No** buttons and a **"Remember my choice for the rest of this batch"** checkbox, instead of having to find and close the app yourself.
- **"Auto-close blocking apps" option** — a checkbox to force-close blocking apps automatically with no prompts (confirmed once up front), for a hands-off run.
- **Stall timeout** — if an app hangs with no activity for 3 minutes, it's automatically aborted and marked **⚠ Stalled** so the batch continues instead of freezing. Slow-but-active downloads are never affected.

### Improved

- **Accurate success/failure reporting** — results are now decided by winget's actual exit code instead of scanning its text output. Fixes apps being wrongly marked failed when their output merely contained words like *"error"* or *"0x"*, and correctly distinguishes *"no update needed"* and *"app in use"* from real failures.
- **Skip button fixed** — pressing **Skip** now correctly marks the app as **⏭ Skipped** instead of sometimes showing **✅ Success**.
- **Filter out fake updates** — apps showing a pointless update where the versions are actually the same (e.g. `2.4.0-release` vs `2.4.0`) are now hidden.
- **Equal-size progress bars** — the two progress bars are now exactly the same width.
- **Responsive Skip and Cancel** — both now respond immediately even while an app is mid-extraction, instead of waiting for the next output line.

### Fixed

- Declining the UAC prompt no longer silently closes the app — it now stays running without admin.
- Faster temp-folder scanning during updates (scoped to winget's own folder instead of all of `%TEMP%`).
- Cleaner output log — winget's progress-bar fragments no longer flood it.
- Removed dead/duplicate code; future-proofed temp cleanup for newer Python versions.

## Earlier versions

- **2.2.1** (January 2026): new app logo
- **2.2** (October 2025): *Save / Export log* button, better log show/hide
- **2.1** (September 2025): improved self-update, About window, text encoding fixes
- **2.0** (September 2025): modern interface with bigger buttons
- **1.2** (September 2025): log spinner and encoding fixes, columns fit after a check, failed apps are skipped instead of stopping the batch
- **1.1** (August 2025): interface updates and column fixes
- **1.0** (August 2025): first release
