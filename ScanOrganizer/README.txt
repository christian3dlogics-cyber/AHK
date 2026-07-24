# ScanOrganizer — Setup & How to Change the Watched Folder

This tool watches a folder (by default, Documents) and automatically moves
any scanned PDF that shows up there into a dated folder, like:

    Documents\scans\07-23-2026\my_scan.pdf

## First-time setup

1. Put all four files in the same folder, e.g. `C:\Scripts\ScanOrganizer\`:
   - ScanOrganizer.ahk
   - ScanOrganizerSettings.ahk
   - ScanOrganizer.ini
   - Install-ScanOrganizerTask.ps1

2. Right-click **Install-ScanOrganizerTask.ps1** → **Run with PowerShell**.
   - It finds AutoHotkey automatically, registers the task, and starts it
     right away — you should see "ScanOrganizer is running."
   - It will also start automatically every time you log in to Windows,
     and restart itself if it's ever force-closed or crashes.

## How to change which folder it watches

**Double-click ScanOrganizerSettings.ahk.** A small window opens with two
boxes:

- **Watch this folder for new scanned PDFs** — where your scanner drops
  new PDFs.
- **Save organized scans into this folder** — where the dated folders get
  created.

Click **Browse...** next to either box to pick a folder instead of typing
a path. Click **Save** — it updates the settings and automatically restarts
the background watcher for you. No further steps needed.

(If you'd rather edit the settings file by hand instead of using the box,
it's ScanOrganizer.ini — open it in Notepad, edit the paths after each
`=` sign, save, then restart the task with `Restart-ScheduledTask -TaskName
'ScanOrganizer'` in PowerShell.)

## Things to know

- Only PDFs sitting loose in the watched folder get moved (not ones already
  inside subfolders).
- It checks for new PDFs about once a minute.
- **A new scan usually takes about 2–3 minutes to move**, not instantly.
  It deliberately waits until the file's size has stopped changing across
  two checks in a row (three checks total, roughly one minute apart) before
  touching it — this avoids grabbing a scan while it's still being written,
  which could corrupt it. Large multi-page scans on a slow scanner may take
  a little longer still, since any change in size resets the wait.
- If two files end up with the same name on the same day, the second one is
  saved as `filename_1.pdf` instead of overwriting the first.
- A log file, **ScanOrganizer.log**, appears next to the script — open it in
  Notepad any time to see what it's done. It logs a heartbeat line every
  5 minutes, so if the log stops updating, something's wrong — check that
  the watched folder still exists (e.g. a network drive that got
  disconnected) and try restarting the task.
