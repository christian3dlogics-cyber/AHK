# Install-ScanOrganizerTask.ps1
# Registers ScanOrganizer.ahk to run hidden at logon via Task Scheduler.
# Run this once, as your normal user (elevation not required).

$TaskName = "ScanOrganizer"
$ScriptPath = "$PSScriptRoot\ScanOrganizer.ahk"

# Try common AutoHotkey v2 install locations before giving up.
$CandidatePaths = @(
    "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe",
    "C:\Program Files\AutoHotkey\v2\AutoHotkey32.exe",
    "C:\Program Files (x86)\AutoHotkey\v2\AutoHotkey64.exe",
    "$env:LOCALAPPDATA\Programs\AutoHotkey\v2\AutoHotkey64.exe"
)
$AhkExe = $CandidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $AhkExe) {
    Write-Warning "AutoHotkey v2 was not found in any common install location."
    Write-Warning "Install it from https://www.autohotkey.com/ (choose the v2 download), then re-run this installer."
    exit 1
}

if (!(Test-Path $ScriptPath)) {
    Write-Warning "ScanOrganizer.ahk not found next to this installer script. Place them in the same folder."
    exit 1
}

if (!(Test-Path "$PSScriptRoot\ScanOrganizer.ini")) {
    Write-Warning "ScanOrganizer.ini not found next to this installer script. Place ScanOrganizer.ahk, ScanOrganizerSettings.ahk, ScanOrganizer.ini, and this installer all in the same folder."
    exit 1
}

$Action = New-ScheduledTaskAction -Execute $AhkExe -Argument "`"$ScriptPath`""
$Trigger = New-ScheduledTaskTrigger -AtLogOn

# -RestartCount / -RestartInterval: if the watcher ever crashes or gets
# force-closed, Windows relaunches it automatically instead of leaving scans
# unorganized until the next logon.
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -Hidden `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable

$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description "Moves scanned PDFs from Documents into date-stamped subfolders." | Out-Null

Write-Host "Task '$TaskName' registered using AutoHotkey at: $AhkExe"
Write-Host "It will start automatically at every logon, and restart itself if it ever stops unexpectedly."
Write-Host ""
Write-Host "Starting it now..."
Start-ScheduledTask -TaskName $TaskName
Start-Sleep -Seconds 2

$taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
if ($taskInfo.LastTaskResult -eq 0 -or $taskInfo.LastTaskResult -eq 267009) {
    Write-Host "ScanOrganizer is running."
} else {
    Write-Warning "The task started but reported an unusual result code ($($taskInfo.LastTaskResult))."
    Write-Warning "Check ScanOrganizer.log next to the script for details, or open ScanOrganizerSettings.ahk to confirm your folder paths."
}
