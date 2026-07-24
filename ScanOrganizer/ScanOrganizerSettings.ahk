#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================
; ScanOrganizerSettings.ahk
; Simple GUI to set the watched scan folder and destination
; folder. Saves to ScanOrganizer.ini — the background watcher
; (ScanOrganizer.ahk) reads that file automatically.
;
; Just double-click this file to open the settings box.
; ============================================================

INI_FILE := A_ScriptDir "\ScanOrganizer.ini"
TASK_NAME := "ScanOrganizer"

; Load current values (blank if ini doesn't exist yet)
currentWatch := FileExist(INI_FILE) ? IniRead(INI_FILE, "General", "WatchFolder", "") : ""
currentDest  := FileExist(INI_FILE) ? IniRead(INI_FILE, "General", "DestinationFolder", "") : ""

; Sensible defaults if nothing set yet
if (currentWatch = "")
    currentWatch := EnvGet("USERPROFILE") "\Documents"
if (currentDest = "")
    currentDest := EnvGet("USERPROFILE") "\Documents\scans"

; ------------------------- Build GUI -------------------------
myGui := Gui("+AlwaysOnTop", "ScanOrganizer Settings")
myGui.SetFont("s10", "Segoe UI")
myGui.MarginX := 15
myGui.MarginY := 15

myGui.Add("Text", , "Watch this folder for new scanned PDFs:")
watchEdit := myGui.Add("Edit", "w420 vWatchPath", currentWatch)
myGui.Add("Button", "x+8 yp-2 w90", "Browse...").OnEvent("Click", BrowseWatch)

myGui.Add("Text", "xm y+15", "Save organized scans into this folder:")
destEdit := myGui.Add("Edit", "w420 vDestPath", currentDest)
myGui.Add("Button", "x+8 yp-2 w90", "Browse...").OnEvent("Click", BrowseDest)

myGui.Add("Text", "xm y+15 cGray", "Scans will be sorted into dated subfolders here, e.g.:")
previewText := myGui.Add("Text", "xm y+2 cGray")

saveBtn := myGui.Add("Button", "xm y+20 w120 h32 Default", "Save")
saveBtn.OnEvent("Click", SaveSettings)

cancelBtn := myGui.Add("Button", "x+10 yp w120 h32", "Cancel")
cancelBtn.OnEvent("Click", (*) => myGui.Destroy())

statusText := myGui.Add("Text", "xm y+15 w420 cGreen", "")

myGui.OnEvent("Close", (*) => myGui.Destroy())
myGui.Show()
UpdatePreview()

watchEdit.OnEvent("Change", UpdatePreview)
destEdit.OnEvent("Change", UpdatePreview)

; ------------------------- Handlers -------------------------

BrowseWatch(*) {
    folder := DirSelect(watchEdit.Value, 3, "Select the folder your scanner saves PDFs into")
    if (folder != "")
        watchEdit.Value := folder
    UpdatePreview()
}

BrowseDest(*) {
    folder := DirSelect(destEdit.Value, 3, "Select where dated scan folders should be created")
    if (folder != "")
        destEdit.Value := folder
    UpdatePreview()
}

UpdatePreview(*) {
    exampleDate := FormatTime(, "MM-dd-yyyy")
    previewText.Text := "  " destEdit.Value "\" exampleDate "\"
}

SaveSettings(*) {
    watchVal := Trim(watchEdit.Value, "\ `t")
    destVal := Trim(destEdit.Value, "\ `t")

    if (watchVal = "" || destVal = "") {
        MsgBox("Both folders are required.", "ScanOrganizer Settings", "Iconx")
        return
    }

    if !DirExist(watchVal) {
        result := MsgBox("This folder doesn't exist yet:`n`n" watchVal "`n`nSave anyway?", "Folder Not Found", "YesNo Icon!")
        if (result = "No")
            return
    }

    ; Create destination folder if it doesn't exist — harmless to pre-create
    if !DirExist(destVal) {
        try DirCreate(destVal)
    }

    IniWrite(watchVal, INI_FILE, "General", "WatchFolder")
    IniWrite(destVal, INI_FILE, "General", "DestinationFolder")

    statusText.Text := "Saving..."
    RestartTask()
}

RestartTask() {
    global TASK_NAME

    checkCmd := 'powershell.exe -NoProfile -Command "if (Get-ScheduledTask -TaskName \"' TASK_NAME '\" -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"'
    checkExit := RunWait(checkCmd, , "Hide")

    if (checkExit = 1) {
        statusText.Text := "Saved. The ScanOrganizer task isn't installed yet — run Install-ScanOrganizerTask.ps1 once, then it'll pick up these settings automatically."
        return
    }

    restartCmd := 'powershell.exe -NoProfile -Command "try { Stop-ScheduledTask -TaskName \"' TASK_NAME '\" -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 1500; Start-ScheduledTask -TaskName \"' TASK_NAME '\" -ErrorAction Stop; exit 0 } catch { exit 1 }"'

    try {
        restartExit := RunWait(restartCmd, , "Hide")
        if (restartExit = 0)
            statusText.Text := "Saved and applied. ScanOrganizer is now watching the new folder."
        else
            statusText.Text := "Saved, but the task didn't restart cleanly. Log off/on, or run Start-ScheduledTask -TaskName 'ScanOrganizer' manually."
    } catch {
        statusText.Text := "Saved, but could not auto-restart the task. Log off/on, or restart it manually."
    }
}
