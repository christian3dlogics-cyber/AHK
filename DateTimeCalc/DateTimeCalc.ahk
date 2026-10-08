#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================
; DateTimeCalc.ahk - personal Date & Time Calculator
;
;   Mode 1  Build Finish   start + run duration = finish (with chaining)
;   Mode 2  Add / Subtract base +/- duration
;   Mode 3  Difference     elapsed time between two date/times
;
; Global hotkey (default Ctrl+Alt+D) toggles the window.
; Settings live in DateTimeCalc.ini next to this script.
;
; Sections:  1 Constants   2 Settings   3 Parsing   4 Time math
;            5 Formatting  6 Chain logic   7 GUI   8 Tray/startup/main
; Sections 3-6 are pure functions (no GUI) and are exercised by tests.ahk,
; which #Includes this file. The GUI only starts when this file is the
; script being run (see the guard on the last line).
; ============================================================

; ------------------------------------------------------------
; 1. CONSTANTS
; ------------------------------------------------------------
APP_NAME := "Date & Time Calc"
INI_FILE := A_ScriptDir "\DateTimeCalc.ini"
DEFAULT_HOTKEY := "^!d"
MAX_SECS := 3153600000          ; 100 years - sanity cap on any duration
LIVE_RESET_MS := 60000          ; hidden longer than this -> Start/Base follow the clock again

; Catppuccin Mocha
C_BASE := "1e1e2e", C_MANTLE := "181825", C_SURF := "313244"
C_TEXT := "cdd6f4", C_SUB := "a6adc8", C_DIM := "7f849c"
C_GREEN := "a6e3a1", C_BLUE := "89b4fa", C_LAV := "b4befe"
C_PEACH := "fab387", C_RED := "f38ba8"

DOW_NAMES := ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
MON_NAMES := ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

Cfg := Map()                    ; settings, filled by LoadSettings()
UI := Object()                  ; GUI controls / widgets
S := Object()                   ; runtime state
S.mode := 1
S.sign := 1
S.shown := false
S.hiddenAt := 0
S.rows := []
S.copy := ""
S.curB := ""
S.subColor := ""
S.statusColor := ""

; ------------------------------------------------------------
; 2. SETTINGS
; ------------------------------------------------------------
SettingDefaults() {
    return Map("Hotkey", DEFAULT_HOTKEY, "AlwaysOnTop", "1", "X", "", "Y", "", "LastMode", "1"
        , "ChangeoverMinutes", "0", "BareNumberUnit", "h", "CopyFormat", "")
}

LoadSettings() {
    global Cfg
    EnsureIni()
    Cfg := Map()
    for k, v in SettingDefaults() {
        val := v
        try val := IniRead(INI_FILE, "Settings", k, v)
        Cfg[k] := val
    }
    ; Silently repair anything missing or corrupt.
    Cfg["Hotkey"] := Trim(Cfg["Hotkey"])
    if (Cfg["Hotkey"] = "")
        Cfg["Hotkey"] := DEFAULT_HOTKEY
    if !RegExMatch(Cfg["AlwaysOnTop"], "^[01]$")
        Cfg["AlwaysOnTop"] := "1"
    for k in ["X", "Y"]
        if !RegExMatch(Cfg[k], "^-?\d{1,6}$")
            Cfg[k] := ""
    if !RegExMatch(Cfg["LastMode"], "^[123]$")
        Cfg["LastMode"] := "1"
    Cfg["ChangeoverMinutes"] := Trim(Cfg["ChangeoverMinutes"])
    if !RegExMatch(Cfg["ChangeoverMinutes"], "^\d{1,5}$")
        Cfg["ChangeoverMinutes"] := "0"
    Cfg["BareNumberUnit"] := StrLower(Trim(Cfg["BareNumberUnit"]))
    if !RegExMatch(Cfg["BareNumberUnit"], "^[wdhms]$")
        Cfg["BareNumberUnit"] := "h"
}

EnsureIni() {
    if FileExist(INI_FILE)
        return
    txt := "; DateTimeCalc settings (edit, then use the tray menu > Reload)`r`n"
        . "[Settings]`r`n"
        . "; Global show/hide hotkey, AutoHotkey syntax (^ Ctrl, ! Alt, + Shift, # Win)`r`n"
        . "Hotkey=" DEFAULT_HOTKEY "`r`n"
        . "; 1 keeps the window on top (Ctrl+T toggles), 0 does not`r`n"
        . "AlwaysOnTop=1`r`n"
        . "; Remembered window position (blank means centered)`r`n"
        . "X=`r`nY=`r`n"
        . "; Remembered mode: 1 Build Finish, 2 Add/Subtract, 3 Difference`r`n"
        . "LastMode=1`r`n"
        . "; Gap in minutes between chained builds`r`n"
        . "ChangeoverMinutes=0`r`n"
        . "; What a bare number means in a duration field: w d h m s`r`n"
        . "BareNumberUnit=h`r`n"
        . "; Optional copy template. Tokens: {label} {ddd} {date} {time12} {time24}`r`n"
        . "; Blank uses the built-in format, e.g. Finish: Fri 10/09/2026 3:45 PM (15:45)`r`n"
        . "CopyFormat=`r`n"
    try FileAppend(txt, INI_FILE, "UTF-8-RAW")
}

SaveSetting(key, val) {
    Cfg[key] := String(val)
    try IniWrite(String(val), INI_FILE, "Settings", key)
}

; ------------------------------------------------------------
; 3. PARSING (pure)
; ------------------------------------------------------------
UnitSecs(u) {
    switch SubStr(u, 1, 1) {
        case "w": return 604800
        case "d": return 86400
        case "h": return 3600
        case "m": return 60
        default: return 1
    }
}

; Returns {ok, secs, err}. secs is a signed integer number of elapsed seconds.
; strict = true: a unit is required (no bare numbers, no hh:mm) - used for "+2h" style offsets.
ParseDuration(text, bareUnit := "h", strict := false) {
    r := {ok: false, secs: 0, err: "can't parse"}
    s := RegExReplace(StrLower(Trim(text)), "\s+", "")
    if (s = "") {
        r.err := "empty"
        return r
    }
    neg := false
    if (SubStr(s, 1, 1) = "-") {
        neg := true
        s := SubStr(s, 2)
    } else if (SubStr(s, 1, 1) = "+")
        s := SubStr(s, 2)
    if (s = "")
        return r
    total := 0
    if (!strict && RegExMatch(s, "^(\d+):(\d{1,2})(?::(\d{1,2}))?$", &m)) {
        mi := Integer(m[2])
        sc := (m[3] = "") ? 0 : Integer(m[3])
        if (mi > 59 || sc > 59)
            return r
        total := Integer(m[1]) * 3600 + mi * 60 + sc
    } else if RegExMatch(s, "^(?:\d+(?:\.\d*)?|\.\d+)$") {
        if strict {
            r.err := "unit needed"
            return r
        }
        total := Number(s) * UnitSecs(bareUnit)
    } else {
        pos := 1
        len := StrLen(s)
        while (pos <= len) {
            if !RegExMatch(s, "A)(\d+(?:\.\d*)?|\.\d+)(weeks?|wks?|w|days?|d|hours?|hrs?|h|minutes?|mins?|m|seconds?|secs?|s)", &m, pos)
                return r
            total += Number(m[1]) * UnitSecs(m[2])
            pos += m.Len
        }
    }
    total := Round(total)
    if (total > MAX_SECS) {
        r.err := "too large"
        return r
    }
    r.ok := true
    r.secs := neg ? -total : total
    return r
}

IsLeap(y) => (Mod(y, 4) = 0 && Mod(y, 100) != 0) || Mod(y, 400) = 0

DaysIn(y, m) {
    static dim := [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    return (m = 2 && IsLeap(y)) ? 29 : dim[m]
}

; Build a 14-digit AHK timestamp, or "" if the parts are not a real date/time.
MakeStamp(y, mo, d, h := 0, mi := 0, sc := 0) {
    if (y < 1601 || y > 9999 || mo < 1 || mo > 12 || d < 1 || d > DaysIn(y, mo)
        || h < 0 || h > 23 || mi < 0 || mi > 59 || sc < 0 || sc > 59)
        return ""
    return Format("{:04}{:02}{:02}{:02}{:02}{:02}", y, mo, d, h, mi, sc)
}

Cut(s, m) => SubStr(s, 1, m.Pos - 1) " " SubStr(s, m.Pos + m.Len)

MonNum(name) => (InStr("janfebmaraprmayjunjulaugsepoctnovdec", SubStr(name, 1, 3)) - 1) // 3 + 1

; Parse a flexible date/time string. `now` is the local "now" stamp (callers pass it
; floored to the minute). Returns {ok, stamp, err} with stamp in local time.
ParseDateTime(text, now) {
    r := {ok: false, stamp: "", err: "can't parse"}
    s := Trim(RegExReplace(StrLower(Trim(text)), "[,\s]+", " "))
    if (s = "" || s = "now" || s = "n") {
        r.ok := true
        r.stamp := now
        return r
    }
    ; Relative to now: +2h, -1d 3h
    if RegExMatch(s, "^[+-]") {
        d := ParseDuration(s, "h", true)
        if !d.ok
            return r
        x := LocalPlusElapsed(now, d.secs)
        if (x = "")
            return r
        r.ok := true
        r.stamp := x
        return r
    }
    s := RegExReplace(s, "(\d)t(\d)", "$1 $2")              ; ISO 2026-10-09T18:00
    s := RegExReplace(s, "(^| )(at|@)(?= |$)", " ")
    nowY := Integer(SubStr(now, 1, 4))

    hasRel := false, relOff := 0
    wday := 0, wnext := false
    hasDate := false, y := nowY, mo := 0, dd := 0

    if RegExMatch(s, "\b(today|tomorrow|tmrw|yesterday|yday)\b", &m) {
        hasRel := true
        relOff := (m[1] = "today") ? 0 : (m[1] = "yesterday" || m[1] = "yday") ? -1 : 1
        s := Cut(s, m)
    }
    if RegExMatch(s, "\bnext\b") {
        wnext := true
        s := RegExReplace(s, "\bnext\b", " ")
    }
    if RegExMatch(s, "\b(sun(?:day)?|mon(?:day)?|tue(?:s(?:day)?)?|wed(?:nesday)?|thu(?:r(?:s(?:day)?)?)?|fri(?:day)?|sat(?:urday)?)\b", &m) {
        wday := (InStr("sunmontuewedthufrisat", SubStr(m[1], 1, 3)) - 1) // 3 + 1
        s := Cut(s, m)
    }

    static monRx := "(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
    if RegExMatch(s, "(?<![\d/.:-])(\d{4})-(\d{1,2})-(\d{1,2})(?![\d/-])", &m) {
        hasDate := true
        y := Integer(m[1]), mo := Integer(m[2]), dd := Integer(m[3])
        s := Cut(s, m)
    } else if RegExMatch(s, "(?<![\d/.:-])(\d{1,2})/(\d{1,2})(?:/(\d{4}|\d{2}))?(?![\d/:])", &m) {
        hasDate := true
        mo := Integer(m[1]), dd := Integer(m[2])
        if (m[3] != "")
            y := (StrLen(m[3]) = 2) ? 2000 + Integer(m[3]) : Integer(m[3])
        s := Cut(s, m)
    } else if RegExMatch(s, "\b" monRx "\.?\s*(\d{1,2})(?:st|nd|rd|th)?\b(?![:\d])(?:\s+((?:19|20)\d\d)\b(?!:))?", &m) {
        hasDate := true
        mo := MonNum(m[1]), dd := Integer(m[2])
        if (m[3] != "")
            y := Integer(m[3])
        s := Cut(s, m)
    } else if RegExMatch(s, "\b(\d{1,2})(?:st|nd|rd|th)?\s*" monRx "\b\.?(?:\s+((?:19|20)\d\d)\b(?!:))?", &m) {
        hasDate := true
        dd := Integer(m[1]), mo := MonNum(m[2])
        if (m[3] != "")
            y := Integer(m[3])
        s := Cut(s, m)
    }

    ; Time of day
    hasTime := false, hh := 0, mi := 0, ss := 0, ap := ""
    rest := Trim(s)
    if RegExMatch(rest, "(?<![\d:])(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(am|pm|a|p)\b)?", &m) {
        hasTime := true
        hh := Integer(m[1]), mi := Integer(m[2])
        ss := (m[3] = "") ? 0 : Integer(m[3])
        ap := m[4]
        rest := Cut(rest, m)
    } else if RegExMatch(rest, "(?<![\d:])(\d{3,4})\s*(am|pm|a|p)\b", &m) {
        hasTime := true
        n := m[1]
        hh := Integer(SubStr(n, 1, StrLen(n) - 2)), mi := Integer(SubStr(n, -2))
        ap := m[2]
        rest := Cut(rest, m)
    } else if RegExMatch(rest, "(?<![\d:])(\d{1,2})\s*(am|pm|a|p)\b", &m) {
        hasTime := true
        hh := Integer(m[1])
        ap := m[2]
        rest := Cut(rest, m)
    } else if RegExMatch(rest, "^(\d{3,4})$", &m) {
        hasTime := true
        n := m[1]
        hh := Integer(SubStr(n, 1, StrLen(n) - 2)), mi := Integer(SubStr(n, -2))
        rest := ""
    } else if RegExMatch(rest, "^(\d{1,2})$", &m) {
        hasTime := true                                    ; bare hour = AM hour (or 24h if > 12)
        hh := Integer(m[1])
        rest := ""
    }
    if (Trim(rest) != "")
        return r
    if (ap != "") {
        if (hh < 1 || hh > 12)
            return r
        hh := Mod(hh, 12) + (SubStr(ap, 1, 1) = "p" ? 12 : 0)
    }
    if (hh > 23 || mi > 59 || ss > 59)
        return r

    if (!hasTime && !hasDate && !hasRel && wday = 0)
        return r
    if ((hasDate ? 1 : 0) + (hasRel ? 1 : 0) + (wday > 0 ? 1 : 0) > 1)
        return r

    nowDate := SubStr(now, 1, 8)
    if hasDate {
        if (MakeStamp(y, mo, dd) = "")
            return r
        dateStr := Format("{:04}{:02}{:02}", y, mo, dd)
    } else if hasRel {
        dateStr := SubStr(DateAdd(nowDate "000000", relOff, "Days"), 1, 8)
    } else if (wday > 0) {
        todayW := Integer(FormatTime(nowDate "000000", "WDay"))
        diff := Mod(wday - todayW + 7, 7)
        if (diff = 0) {
            ; Same weekday: today if the time is still ahead (or no time given), else next week.
            if (wnext || (hasTime && Integer(SubStr(now, 9, 6)) >= hh * 10000 + mi * 100 + ss))
                diff := 7
        }
        dateStr := SubStr(DateAdd(nowDate "000000", diff, "Days"), 1, 8)
    } else
        dateStr := nowDate

    r.ok := true
    r.stamp := dateStr Format("{:02}{:02}{:02}", hh, mi, ss)
    return r
}

; ------------------------------------------------------------
; 4. TIME MATH (pure; local <-> UTC through the Windows time zone APIs)
; Stamps are 14-digit AHK timestamps (YYYYMMDDHH24MISS). "Local" stamps are wall-clock
; time in the Windows time zone; "UTC" stamps are used for all arithmetic.
; ------------------------------------------------------------
StampToST(stamp, buf) {
    NumPut("UShort", Integer(SubStr(stamp, 1, 4)), buf, 0)
    NumPut("UShort", Integer(SubStr(stamp, 5, 2)), buf, 2)
    NumPut("UShort", 0, buf, 4)
    NumPut("UShort", Integer(SubStr(stamp, 7, 2)), buf, 6)
    NumPut("UShort", Integer(SubStr(stamp, 9, 2)), buf, 8)
    NumPut("UShort", Integer(SubStr(stamp, 11, 2)), buf, 10)
    NumPut("UShort", Integer(SubStr(stamp, 13, 2)), buf, 12)
    NumPut("UShort", 0, buf, 14)
}

STToStamp(buf) {
    return Format("{:04}{:02}{:02}{:02}{:02}{:02}", NumGet(buf, 0, "UShort"), NumGet(buf, 2, "UShort")
        , NumGet(buf, 6, "UShort"), NumGet(buf, 8, "UShort"), NumGet(buf, 10, "UShort"), NumGet(buf, 12, "UShort"))
}

; Windows' own answer for local -> UTC (its choice on ambiguous/missing times is not relied on).
RawLocalToUtc(localStamp) {
    inB := Buffer(16, 0), outB := Buffer(16, 0)
    StampToST(localStamp, inB)
    if !DllCall("TzSpecificLocalTimeToSystemTime", "Ptr", 0, "Ptr", inB, "Ptr", outB, "Int")
        return ""
    return STToStamp(outB)
}

UtcToLocal(utcStamp) {
    inB := Buffer(16, 0), outB := Buffer(16, 0)
    try {
        StampToST(utcStamp, inB)
        if !DllCall("SystemTimeToTzSpecificLocalTime", "Ptr", 0, "Ptr", inB, "Ptr", outB, "Int")
            return ""
    } catch
        return ""
    return STToStamp(outB)
}

; Local wall-clock -> UTC with predictable handling of DST edge cases.
; Returns {ok, utc, loc, status}; status = "ok" | "ambiguous" (fall-back hour: the FIRST
; occurrence is used) | "gap" (spring-forward gap: the time is moved forward by the gap, loc
; holds the time actually used).
LocalToUtc(localStamp) {
    r := {ok: false, utc: "", loc: localStamp, status: "ok"}
    try {
        u0 := RawLocalToUtc(localStamp)
        if (u0 = "")
            return r
        best := ""
        count := 0
        for d in [0, -3600, 3600, -1800, 1800] {
            u := DateAdd(u0, d, "Seconds")
            if (UtcToLocal(u) = localStamp) {
                count++
                if (best = "" || StrCompare(u, best) < 0)
                    best := u
            }
        }
        if (count = 0) {
            shifted := DateAdd(localStamp, 3600, "Seconds")
            r2 := LocalToUtc(shifted)
            if !r2.ok
                return r
            r.ok := true
            r.utc := r2.utc
            r.loc := shifted
            r.status := "gap"
            return r
        }
        r.ok := true
        r.utc := best
        r.status := (count > 1) ? "ambiguous" : "ok"
    } catch
        r.ok := false
    return r
}

OffsetSecs(utcStamp) {
    loc := UtcToLocal(utcStamp)
    return (loc = "") ? 0 : DateDiff(loc, utcStamp, "Seconds")
}

; UTC-offset changes (in seconds) between two UTC stamps, in order. Negative = fall back.
; Samples every 20 days, which is far finer than the gap between DST transitions.
DstChanges(utc1, utc2) {
    changes := []
    if (StrCompare(utc1, utc2) > 0) {
        t := utc1, utc1 := utc2, utc2 := t
    }
    prevU := utc1
    prevO := OffsetSecs(utc1)
    try {
        loop {
            nextU := DateAdd(prevU, 20 * 86400, "Seconds")
            if (StrCompare(nextU, utc2) > 0)
                nextU := utc2
            o := OffsetSecs(nextU)
            if (o != prevO)
                changes.Push(o - prevO)
            prevU := nextU, prevO := o
            if (nextU = utc2)
                break
        }
    }
    return changes
}

; local stamp + elapsed seconds -> local stamp ("" on failure)
LocalPlusElapsed(localStamp, secs) {
    s := LocalToUtc(localStamp)
    if !s.ok
        return ""
    try return UtcToLocal(DateAdd(s.utc, secs, "Seconds"))
    return ""
}

; Add (or subtract) elapsed seconds to a local start time.
; Returns {ok, err, startLoc, startUtc, endUtc, endLoc, dst[], notes[]}
AddElapsed(localStart, secs) {
    res := {ok: false, err: "", startLoc: localStart, startUtc: "", endUtc: "", endLoc: "", dst: [], notes: []}
    s := LocalToUtc(localStart)
    if !s.ok {
        res.err := "invalid start time"
        return res
    }
    try {
        endUtc := DateAdd(s.utc, secs, "Seconds")
        endLoc := UtcToLocal(endUtc)
    } catch {
        res.err := "result out of range"
        return res
    }
    if (endLoc = "") {
        res.err := "result out of range"
        return res
    }
    res.ok := true
    res.startLoc := s.loc
    res.startUtc := s.utc
    res.endUtc := endUtc
    res.endLoc := endLoc
    res.dst := DstChanges(s.utc, endUtc)
    AddDstNotes(res.notes, res.dst)
    AddStatusNote(res.notes, "Start", s)
    e := LocalToUtc(endLoc)
    if (e.ok && e.status = "ambiguous")
        res.notes.Push("Result falls in the repeated fall-back hour (" (e.utc = endUtc ? "1st" : "2nd") " pass)")
    return res
}

; Elapsed time from one local time to another. Returns {ok, err, secs, dst[], notes[]}
DiffElapsed(stampA, stampB) {
    res := {ok: false, err: "", secs: 0, dst: [], notes: []}
    a := LocalToUtc(stampA)
    b := LocalToUtc(stampB)
    if !(a.ok && b.ok) {
        res.err := "invalid date/time"
        return res
    }
    res.ok := true
    res.secs := DateDiff(b.utc, a.utc, "Seconds")
    res.dst := DstChanges(a.utc, b.utc)
    AddDstNotes(res.notes, res.dst)
    AddStatusNote(res.notes, "Start", a)
    AddStatusNote(res.notes, "End", b)
    return res
}

AddStatusNote(notes, label, st) {
    if (st.status = "ambiguous")
        notes.Push(label " is in the repeated fall-back hour - using the first occurrence")
    else if (st.status = "gap")
        notes.Push(label " doesn't exist (spring-forward gap) - using " Fmt12(st.loc))
}

AddDstNotes(notes, changes) {
    for c in changes {
        kind := (c < 0) ? "fall back, +" : "spring forward, -"
        notes.Push("⚠ crosses DST change (" kind FmtHM(Abs(c)) ")")
    }
}

; ------------------------------------------------------------
; 5. FORMATTING (pure)
; ------------------------------------------------------------
StampDow(stamp) => DOW_NAMES[Integer(FormatTime(SubStr(stamp, 1, 8) "000000", "WDay"))]

Fmt12(stamp) {
    h := Integer(SubStr(stamp, 9, 2)), mi := SubStr(stamp, 11, 2), ss := Integer(SubStr(stamp, 13, 2))
    h12 := Mod(h, 12) = 0 ? 12 : Mod(h, 12)
    return h12 ":" mi (ss ? Format(":{:02}", ss) : "") " " (h < 12 ? "AM" : "PM")
}

Fmt24(stamp) {
    ss := Integer(SubStr(stamp, 13, 2))
    return SubStr(stamp, 9, 2) ":" SubStr(stamp, 11, 2) (ss ? Format(":{:02}", ss) : "")
}

; Fri Oct 9, 2026 · 3:45 PM (15:45)
FmtDT(stamp) {
    return StampDow(stamp) " " MON_NAMES[Integer(SubStr(stamp, 5, 2))] " " Integer(SubStr(stamp, 7, 2)) ", "
        . SubStr(stamp, 1, 4) " · " Fmt12(stamp) " (" Fmt24(stamp) ")"
}

; Fri 10/09/2026 3:45 PM (15:45)
FmtCopyDT(stamp) {
    return StampDow(stamp) " " SubStr(stamp, 5, 2) "/" SubStr(stamp, 7, 2) "/" SubStr(stamp, 1, 4)
        . " " Fmt12(stamp) " (" Fmt24(stamp) ")"
}

; Fri 10/9 6:00 AM
FmtShort(stamp) {
    return StampDow(stamp) " " Integer(SubStr(stamp, 5, 2)) "/" Integer(SubStr(stamp, 7, 2)) " " Fmt12(stamp)
}

; Text for a live field; parses back through ParseDateTime.
FmtLive(stamp) {
    return Integer(SubStr(stamp, 5, 2)) "/" Integer(SubStr(stamp, 7, 2)) "/" SubStr(stamp, 1, 4) " " Fmt12(stamp)
}

; Unambiguous text (ISO) used to hand a computed time to a field.
FmtIso(stamp) {
    return SubStr(stamp, 1, 4) "-" SubStr(stamp, 5, 2) "-" SubStr(stamp, 7, 2) " " Fmt24(stamp)
}

CopyLine(label, stamp) {
    tpl := Cfg.Has("CopyFormat") ? Cfg["CopyFormat"] : ""
    if (tpl = "")
        return label ": " FmtCopyDT(stamp)
    tpl := StrReplace(tpl, "{label}", label)
    tpl := StrReplace(tpl, "{ddd}", StampDow(stamp))
    tpl := StrReplace(tpl, "{date}", SubStr(stamp, 5, 2) "/" SubStr(stamp, 7, 2) "/" SubStr(stamp, 1, 4))
    tpl := StrReplace(tpl, "{time12}", Fmt12(stamp))
    tpl := StrReplace(tpl, "{time24}", Fmt24(stamp))
    return tpl
}

AddCommas(n) => RegExReplace(String(n), "\B(?=(\d{3})+(?!\d))", ",")

; 2d 4h 12m  (zero units omitted; seconds only when non-zero and showSecs)
FmtDur(secs, showSecs := true) {
    neg := secs < 0
    a := Abs(secs)
    if !showSecs
        a -= Mod(a, 60)
    d := a // 86400, h := Mod(a // 3600, 24), m := Mod(a // 60, 60), sc := Mod(a, 60)
    parts := []
    if d
        parts.Push(d "d")
    if h
        parts.Push(h "h")
    if m
        parts.Push(m "m")
    if sc
        parts.Push(sc "s")
    out := ""
    for p in parts
        out .= (out = "" ? "" : " ") p
    if (out = "")
        out := "0m"
    return neg ? "-" out : out
}

FmtHM(secs) {
    h := secs // 3600, m := Mod(secs // 60, 60)
    return (h ? h "h" : "") (m ? (h ? " " : "") m "m" : "")
}

FmtDecHours(secs) => Format("{:.2f} h", secs / 3600)

FmtMinutes(secs) {
    if (Mod(secs, 60) = 0)
        return AddCommas(secs // 60) " min"
    return Format("{:.1f}", secs / 60) " min"
}

; 1w 0d 4h 12m
FmtWeeks(secs) {
    neg := secs < 0
    a := Abs(secs)
    out := Format("{}w {}d {}h {}m", a // 604800, Mod(a // 86400, 7), Mod(a // 3600, 24), Mod(a // 60, 60))
    if Mod(a, 60)
        out .= " " Mod(a, 60) "s"
    return neg ? "-" out : out
}

; All duration formats at once.
DurAll(secs) {
    return {main: FmtDur(secs), dec: FmtDecHours(secs), mins: FmtMinutes(secs), wk: FmtWeeks(secs)}
}

; Relative info for a finish time: secs = finish - now.
FmtRel(secs) {
    if (Abs(secs) < 60)
        return "now"
    secs := (secs < 0) ? -(Abs(secs) // 60 * 60) : secs // 60 * 60
    return (secs > 0) ? "in " FmtDur(secs, false) : FmtDur(-secs, false) " ago"
}

; ------------------------------------------------------------
; 6. CHAIN LOGIC (pure)
; ------------------------------------------------------------
; Start of the next build = previous finish (UTC) + changeover, returned as a local stamp.
ChainNextStart(prevEndUtc, changeoverMin) {
    try return UtcToLocal(DateAdd(prevEndUtc, changeoverMin * 60, "Seconds"))
    return ""
}

ChainText(rows) {
    out := "#`tStart`tRun`tFinish"
    for i, row in rows
        out .= "`r`n" i "`t" FmtCopyDT(row.startLoc) "`t" FmtDur(row.secs, false) "`t" FmtCopyDT(row.endLoc)
    return out
}

; ------------------------------------------------------------
; 7. GUI
; ------------------------------------------------------------
NowMin() => SubStr(A_Now, 1, 12) "00"

RGB2BGR(hex) {
    return Integer("0x" SubStr(hex, 1, 2)) | (Integer("0x" SubStr(hex, 3, 2)) << 8) | (Integer("0x" SubStr(hex, 5, 2)) << 16)
}

SetTxt(ctrl, t) {
    if (ctrl.Text != t)
        ctrl.Text := t
}

Font(size, color := "", face := "Segoe UI") {
    UI.g.SetFont("s" size " c" (color = "" ? C_TEXT : color), face)
}

DarkTitleBar(hwnd) {
    val := Buffer(4, 0)
    NumPut("Int", 1, val)
    try {
        if (DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "Int", 20, "Ptr", val, "Int", 4, "Int") != 0)
            DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "Int", 19, "Ptr", val, "Int", 4, "Int")
    }
}

DarkTheme(ctrl, name := "DarkMode_Explorer") {
    try DllCall("uxtheme\SetWindowTheme", "Ptr", ctrl.Hwnd, "Str", name, "Ptr", 0)
}

; Clickable Text control styled as a flat button.
MakeBtn(label, x, y, w, h, handler, color := "") {
    Font(9, color)
    b := UI.g.Add("Text", Format("x{} y{} w{} h{} Center +0x200 +0x100 Background{} c{}", x, y, w, h, C_SURF, color = "" ? C_TEXT : color), label)
    b.OnEvent("Click", handler)
    return b
}

; Label + framed Edit (+ optional parsed-echo line). Returns {lbl, frame, edit, echo, ctrls[]}
MakeInput(label, x, y, w, withEcho := true) {
    g := UI.g
    o := Object()
    Font(9, C_SUB)
    o.lbl := g.Add("Text", Format("x{} y{} w{} h16", x, y, w), label)
    o.frame := g.Add("Text", Format("x{} y{} w{} h30 Background{}", x, y + 16, w, C_SURF))
    Font(10, C_TEXT)
    o.edit := g.Add("Edit", Format("x{} y{} w{} r1 -E0x200 Background{} c{}", x + 8, y + 22, w - 16, C_SURF, C_TEXT))
    o.ctrls := [o.lbl, o.frame, o.edit]
    o.echo := ""
    if withEcho {
        Font(9, C_DIM)
        o.echo := g.Add("Text", Format("x{} y{} w{} h16", x, y + 48, w), "")
        o.ctrls.Push(o.echo)
    }
    return o
}

class FieldBase {
    SetEcho(txt, bad) {
        SetTxt(this.echo, txt)
        if (bad != this.bad) {
            this.bad := bad
            this.echo.SetFont("s9 c" (bad ? C_RED : C_DIM), "Segoe UI")
        }
    }
    SetText(t) {
        this.lastSet := t
        this.edit.Value := t
    }
    Focus() {
        this.edit.Focus()
        SendMessage(0xB1, 0, -1, this.edit)         ; EM_SETSEL: select all
    }
}

; Date/time field with smart parsing, echo line and optional live "follow the clock".
class DTField extends FieldBase {
    __New(label, x, y, w, hasLive := false) {
        this.hasLive := hasLive
        this.live := hasLive
        this.lastSet := ""
        this.bad := false
        this.shown := false
        inp := MakeInput(label, x, y, w)
        this.edit := inp.edit
        this.echo := inp.echo
        this.ctrls := inp.ctrls
        if hasLive {
            Font(9, C_GREEN)
            this.liveTxt := UI.g.Add("Text", Format("x{} y{} w108 h16 Right c{}", x + w + 8, y, C_GREEN), "● live")
            this.nowBtn := MakeBtn("Now  (Ctrl+N)", x + w + 8, y + 16, 108, 30, NowClicked)
        }
        this.edit.OnEvent("Change", ObjBindMethod(this, "OnChange"))
    }
    OnChange(*) {
        if (this.edit.Value == this.lastSet)         ; programmatic change, not the user
            return
        if this.live {
            this.live := false
            this.UpdateLive()
        }
        ScheduleCalc()
    }
    UpdateLive() {
        if this.hasLive
            this.liveTxt.Visible := this.shown && this.live
    }
    SetNow() {
        this.live := this.hasLive
        this.UpdateLive()
        this.SetText(FmtLive(NowMin()))
    }
    Show(on) {
        this.shown := on
        for c in this.ctrls
            c.Visible := on
        if this.hasLive
            this.nowBtn.Visible := on
        this.UpdateLive()
    }
    Eval(now) {
        r := ParseDateTime(this.edit.Value, now)
        if r.ok
            this.SetEcho("→ " FmtDT(r.stamp), false)
        else
            this.SetEcho("can't parse", true)
        return r
    }
}

; Duration field.
class DurField extends FieldBase {
    __New(label, x, y, w) {
        this.lastSet := ""
        this.bad := false
        inp := MakeInput(label, x, y, w)
        this.edit := inp.edit
        this.echo := inp.echo
        this.ctrls := inp.ctrls
        this.edit.OnEvent("Change", (*) => ScheduleCalc())
    }
    Show(on) {
        for c in this.ctrls
            c.Visible := on
    }
    Eval() {
        r := ParseDuration(this.edit.Value, Cfg["BareNumberUnit"])
        if r.ok
            this.SetEcho("→ " FmtDur(r.secs) " (" Format("{:.2f}", r.secs / 3600) " h)", false)
        else if (r.err = "empty")
            this.SetEcho("→ enter a duration, e.g. 14h 30m", false)
        else
            this.SetEcho(r.err = "can't parse" ? "can't parse" : r.err, true)
        return r
    }
}

; Plain numeric field (changeover minutes).
class NumField extends FieldBase {
    __New(label, x, y, w) {
        this.lastSet := ""
        inp := MakeInput(label, x, y, w, false)
        this.edit := inp.edit
        this.ctrls := inp.ctrls
        this.edit.OnEvent("Change", ObjBindMethod(this, "OnChange"))
    }
    OnChange(*) {
        v := Trim(this.edit.Value)
        if RegExMatch(v, "^\d{1,5}$")
            SaveSetting("ChangeoverMinutes", Integer(v))
        ScheduleCalc()
    }
    Value() {
        v := Trim(this.edit.Value)
        return RegExMatch(v, "^\d{1,5}$") ? Integer(v) : 0
    }
    Show(on) {
        for c in this.ctrls
            c.Visible := on
    }
}

BuildGui() {
    UI.mono := FileExist(A_WinDir "\Fonts\CascadiaMono.ttf") ? "Cascadia Mono" : "Consolas"
    g := Gui("-MinimizeBox -MaximizeBox", APP_NAME)
    UI.g := g
    g.BackColor := C_BASE
    g.MarginX := 0, g.MarginY := 0
    if (Cfg["AlwaysOnTop"] = "1")
        g.Opt("+AlwaysOnTop")
    DarkTitleBar(g.Hwnd)
    Font(9)

    ; --- tabs (Tab3 can't be themed dark, so: three clickable Text controls)
    UI.tabs := []
    names := ["1   Build Finish", "2   Add / Subtract", "3   Difference"]
    loop 3 {
        i := A_Index
        t := g.Add("Text", Format("x{} y10 w148 h28 Center +0x200 +0x100 Background{} c{}", 12 + (i - 1) * 154, C_SURF, C_SUB), names[i])
        t.OnEvent("Click", SetMode.Bind(i))
        UI.tabs.Push(t)
    }

    UI.groups := [[], [], []]

    ; --- mode 1: build finish
    UI.f1s := DTField("Start", 12, 48, 340, true)
    UI.f1r := DurField("Run duration", 12, 118, 340)
    UI.f1c := NumField("Changeover (min)", 360, 118, 108)
    UI.f1c.SetText(Cfg["ChangeoverMinutes"])
    UI.groups[1].Push(UI.f1s, UI.f1r, UI.f1c)

    ; --- mode 2: add / subtract
    UI.f2b := DTField("Base", 12, 48, 340, true)
    Font(9, C_SUB)
    UI.signLbl := g.Add("Text", "x12 y118 w60 h16", "Sign")
    UI.signBtn := MakeBtn("+", 12, 134, 34, 30, SignClicked, C_GREEN)
    UI.f2d := DurField("Duration  (+ / - keys set the sign)", 54, 118, 414)
    UI.groups[2].Push(UI.f2b, UI.signLbl, UI.signBtn, UI.f2d)

    ; --- mode 3: difference
    UI.f3s := DTField("Start", 12, 48, 340, true)
    UI.f3e := DTField("End  (blank = now)", 12, 118, 456)
    UI.groups[3].Push(UI.f3s, UI.f3e)

    ; --- shared result panel
    UI.panel := g.Add("Text", Format("x12 y188 w456 h102 Background{}", C_MANTLE))
    Font(13, C_GREEN, UI.mono)
    UI.resBig := g.Add("Text", Format("x20 y193 w440 h28 +0x200 Background{} c{}", C_MANTLE, C_GREEN))
    Font(9, C_LAV, UI.mono)
    UI.resSub := g.Add("Text", Format("x20 y224 w440 h18 +0x200 Background{} c{}", C_MANTLE, C_LAV))
    Font(9, C_PEACH)
    UI.resNote := g.Add("Text", Format("x20 y244 w440 h44 Background{} c{}", C_MANTLE, C_PEACH))
    S.subColor := C_LAV

    ; --- mode 1: chain controls
    UI.bAdd := MakeBtn("Add build  (Ctrl+Enter)", 12, 298, 180, 26, AddBuild, C_BLUE)
    UI.bUndo := MakeBtn("Remove last  (Ctrl+Bksp)", 200, 298, 148, 26, RemoveLast)
    UI.bClear := MakeBtn("Clear", 356, 298, 112, 26, ClearChain)
    UI.groups[1].Push(UI.bAdd, UI.bUndo, UI.bClear)
    ; Header drawn as Text controls: the native ListView header can't be themed dark reliably.
    UI.lvHdr := []
    Font(9, C_SUB)
    for h in [["#", 16, 24], ["Start", 46, 140], ["Run", 196, 80], ["Finish", 286, 150]]
        UI.lvHdr.Push(g.Add("Text", Format("x{} y328 w{} h16 Background{}", 12 + h[2], h[3], C_BASE), h[1]))
    Font(9, C_TEXT)
    UI.lv := g.Add("ListView", Format("x12 y344 w456 h98 -Multi -Hdr NoSortHdr -E0x200 Background{} c{}", C_MANTLE, C_TEXT), ["#", "Start", "Run", "Finish"])
    UI.lv.ModifyCol(1, 30), UI.lv.ModifyCol(2, 150), UI.lv.ModifyCol(3, 90), UI.lv.ModifyCol(4, 168)
    DarkTheme(UI.lv)
    try {
        SendMessage(0x1001, 0, RGB2BGR(C_MANTLE), UI.lv)    ; LVM_SETBKCOLOR
        SendMessage(0x1026, 0, RGB2BGR(C_MANTLE), UI.lv)    ; LVM_SETTEXTBKCOLOR
        SendMessage(0x1024, 0, RGB2BGR(C_TEXT), UI.lv)      ; LVM_SETTEXTCOLOR
    }
    Font(9, C_SUB, UI.mono)
    UI.span := g.Add("Text", "x12 y446 w456 h18", "")

    ; --- bottom bar (positioned by Layout)
    Font(9, C_SUB)
    UI.status := g.Add("Text", "x12 y0 w340 h22 +0x200", "")
    UI.bCopy := MakeBtn("Copy  (Ctrl+Shift+C)", 308, 0, 160, 24, DoCopy)
    Font(8, C_DIM)
    UI.hint := g.Add("Text", "x12 y0 w456 h16", "")

    g.OnEvent("Escape", (*) => HideWin())
    g.OnEvent("Close", (*) => (HideWin(), true))
}

HintFor(mode) {
    if (mode = 1)
        return "Ctrl+1/2/3 mode · Ctrl+Enter add build · Ctrl+Shift+C copy · Ctrl+T pin · Esc hide"
    if (mode = 2)
        return "Ctrl+1/2/3 mode · + / - sign · Ctrl+N now · Ctrl+Shift+C copy · Ctrl+T pin · Esc hide"
    return "Ctrl+1/2/3 mode · Ctrl+N now · Ctrl+Shift+C copy · Ctrl+T pin · Esc hide"
}

ContentEnd() {
    if (S.mode = 1)
        return S.rows.Length ? 470 : 332
    return 296
}

; Show/hide the control groups for the current mode, then lay out the bottom bar and size the window.
ApplyMode() {
    g := UI.g
    if S.shown
        SendMessage(0xB, 0, 0, g)                   ; WM_SETREDRAW off (no flicker)
    loop 3 {
        on := (A_Index = S.mode)
        for item in UI.groups[A_Index] {
            if (item is Gui.Control)
                item.Visible := on
            else
                item.Show(on)
        }
        t := UI.tabs[A_Index]
        t.Opt("+Background" (on ? C_BLUE : C_SURF))
        t.SetFont("s9 c" (on ? C_BASE : C_SUB), "Segoe UI")
    }
    chain := (S.mode = 1 && S.rows.Length > 0)
    UI.lv.Visible := chain
    for h in UI.lvHdr
        h.Visible := chain
    UI.span.Visible := chain
    SetSubColor(S.mode = 3 ? C_SUB : C_LAV)

    B := ContentEnd()
    UI.status.Move(12, B + 4)
    UI.bCopy.Move(308, B + 2)
    UI.hint.Move(12, B + 36)
    SetTxt(UI.hint, HintFor(S.mode))
    S.height := B + 60
    if S.shown {
        g.Show("NA w480 h" S.height)
        SendMessage(0xB, 1, 0, g)                   ; WM_SETREDRAW on
        DllCall("RedrawWindow", "Ptr", g.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x185)
    }
}

SetSubColor(col) {
    if (S.subColor != col) {
        S.subColor := col
        UI.resSub.SetFont("s9 c" col, UI.mono)
    }
}

SetResult(big, sub, note) {
    SetTxt(UI.resBig, big)
    SetTxt(UI.resSub, sub)
    SetTxt(UI.resNote, note)
}

NotesText(r) {
    out := ""
    for n in r.notes
        out .= (out = "" ? "" : "`n") n
    return out
}

Flash(msg, color := "") {
    color := (color = "") ? C_GREEN : color
    if (S.statusColor != color) {
        S.statusColor := color
        UI.status.SetFont("s9 c" color, "Segoe UI")
    }
    SetTxt(UI.status, msg)
    SetTimer(ClearStatus, -1600)
}

ClearStatus(*) {
    SetTxt(UI.status, "")
}

; --- live recalculation
ScheduleCalc() {
    SetTimer(Recalc, -150)
}

Recalc(*) {
    if !S.shown
        return
    now := NowMin()
    S.copy := ""
    S.curB := ""
    switch S.mode {
        case 1: Calc1(now)
        case 2: Calc2(now)
        case 3: Calc3(now)
    }
}

Calc1(now) {
    st := UI.f1s.Eval(now)
    du := UI.f1r.Eval()
    if !(st.ok && du.ok) {
        SetResult("", "", "")
        return
    }
    r := AddElapsed(st.stamp, du.secs)
    if !r.ok {
        SetResult("", "", "can't compute: " r.err)
        return
    }
    note := NotesText(r)
    if (du.secs < 0)
        note := "Run duration is negative" (note = "" ? "" : "`n" note)
    SetResult(FmtDT(r.endLoc), FmtRel(DateDiff(r.endUtc, A_NowUTC, "Seconds")), note)
    S.copy := CopyLine("Finish", r.endLoc)
    if (du.secs >= 0)
        S.curB := {startLoc: r.startLoc, startUtc: r.startUtc, secs: du.secs, endLoc: r.endLoc, endUtc: r.endUtc}
}

Calc2(now) {
    b := UI.f2b.Eval(now)
    du := UI.f2d.Eval()
    if !(b.ok && du.ok) {
        SetResult("", "", "")
        return
    }
    r := AddElapsed(b.stamp, S.sign * du.secs)
    if !r.ok {
        SetResult("", "", "can't compute: " r.err)
        return
    }
    SetResult(FmtDT(r.endLoc), FmtRel(DateDiff(r.endUtc, A_NowUTC, "Seconds")), NotesText(r))
    S.copy := CopyLine("Result", r.endLoc)
}

Calc3(now) {
    a := UI.f3s.Eval(now)
    b := UI.f3e.Eval(now)
    if !(a.ok && b.ok) {
        SetResult("", "", "")
        return
    }
    d := DiffElapsed(a.stamp, b.stamp)
    if !d.ok {
        SetResult("", "", "can't compute: " d.err)
        return
    }
    da := DurAll(d.secs)
    note := NotesText(d)
    if (d.secs < 0)
        note := "⚠ End is before Start (negative elapsed)" (note = "" ? "" : "`n" note)
    SetResult(da.main, da.dec " · " da.mins " · " da.wk, note)
    S.copy := "Elapsed: " da.main " (" da.dec " · " da.mins ")"
}

; --- once a second while visible: keep live fields on the clock, refresh relative text
Tick(*) {
    if !S.shown
        return
    t := FmtLive(NowMin())
    for f in [UI.f1s, UI.f2b, UI.f3s] {
        if (f.live && t != f.lastSet) {
            f.SetText(t)
            if (DllCall("GetFocus", "Ptr") = f.edit.Hwnd)
                SendMessage(0xB1, 0, -1, f.edit)
        }
    }
    Recalc()
}

; --- mode / sign / now
SetMode(n, *) {
    S.mode := n
    SaveSetting("LastMode", n)
    ApplyMode()
    Recalc()
    FocusFirst()
}

FirstField() {
    return (S.mode = 1) ? UI.f1s : (S.mode = 2) ? UI.f2b : UI.f3s
}

FocusFirst() {
    FirstField().Focus()
}

NowClicked(*) {
    DoNow()
}

DoNow() {
    FirstField().SetNow()
    Recalc()
    FocusFirst()
}

SetSign(v) {
    S.sign := v
    UI.signBtn.Text := (v > 0) ? "+" : "−"
    UI.signBtn.SetFont("s9 c" (v > 0 ? C_GREEN : C_PEACH), "Segoe UI")
    Recalc()
}

SignClicked(*) {
    SetSign(-S.sign)
}

; --- chain actions
AddBuild(*) {
    if (S.mode != 1 || S.curB = "")
        return
    S.rows.Push(S.curB)
    NextBuildStart()
    RefreshChain()
    UI.f1r.Focus()
    Recalc()
}

NextBuildStart() {
    last := S.rows[S.rows.Length]
    nxt := ChainNextStart(last.endUtc, UI.f1c.Value())
    UI.f1s.live := false
    UI.f1s.UpdateLive()
    UI.f1s.SetText(FmtIso(nxt))
}

RemoveLast(*) {
    if (S.mode != 1 || !S.rows.Length)
        return
    S.rows.Pop()
    if S.rows.Length
        NextBuildStart()
    else
        UI.f1s.SetNow()
    RefreshChain()
    Recalc()
}

ClearChain(*) {
    if !S.rows.Length
        return
    S.rows := []
    UI.f1s.SetNow()
    RefreshChain()
    Recalc()
}

RefreshChain() {
    lv := UI.lv
    lv.Delete()
    for i, row in S.rows
        lv.Add("", i, FmtShort(row.startLoc), FmtDur(row.secs, false), FmtShort(row.endLoc))
    if S.rows.Length {
        da := DurAll(DateDiff(S.rows[S.rows.Length].endUtc, S.rows[1].startUtc, "Seconds"))
        SetTxt(UI.span, "Span " da.main " · " da.dec " · " da.mins " · " da.wk)
    } else
        SetTxt(UI.span, "")
    ApplyMode()
}

; --- copy
DoCopy(*) {
    txt := (S.mode = 1 && S.rows.Length) ? ChainText(S.rows) : S.copy
    if (txt = "") {
        Flash("Nothing to copy", C_PEACH)
        return
    }
    A_Clipboard := txt
    Flash("Copied")
}

; --- window show / hide
WinIsActive(*) {
    return S.shown && WinActive("ahk_id " UI.g.Hwnd)
}

ToggleWin(*) {
    if WinIsActive()
        HideWin()
    else
        ShowWin()
}

PositionOK(x, y) {
    loop MonitorGetCount() {
        MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
        if (x + 120 > l && x < r - 120 && y >= t - 10 && y < b - 60)
            return true
    }
    return false
}

ShowWin(*) {
    g := UI.g
    wasShown := S.shown
    if (!wasShown && S.hiddenAt && (A_TickCount - S.hiddenAt) > LIVE_RESET_MS)
        ResetLive()
    S.shown := true
    ApplyMode()
    Recalc()
    opts := "w480 h" S.height
    if (!wasShown && Cfg["X"] != "" && Cfg["Y"] != "" && PositionOK(Integer(Cfg["X"]), Integer(Cfg["Y"])))
        opts .= " x" Cfg["X"] " y" Cfg["Y"]
    g.Show(opts)
    SetTimer(Tick, 1000)
    FocusFirst()
}

HideWin(*) {
    if !S.shown
        return
    SavePos()
    UI.g.Hide()
    S.shown := false
    S.hiddenAt := A_TickCount
    SetTimer(Tick, 0)
    SetTimer(Recalc, 0)
}

SavePos() {
    try {
        WinGetPos(&x, &y, , , "ahk_id " UI.g.Hwnd)
        if (x > -30000 && y > -30000) {
            if (String(x) != Cfg["X"])
                SaveSetting("X", x)
            if (String(y) != Cfg["Y"])
                SaveSetting("Y", y)
        }
    }
}

ResetLive() {
    UI.f1s.SetNow()
    UI.f2b.SetNow()
    UI.f3s.SetNow()
}

ToggleOnTop(*) {
    on := (Cfg["AlwaysOnTop"] != "1")
    SaveSetting("AlwaysOnTop", on ? 1 : 0)
    UI.g.Opt(on ? "+AlwaysOnTop" : "-AlwaysOnTop")
    Flash(on ? "Pinned on top" : "Unpinned")
}

; --- hotkeys that only apply while the calculator is active
Mode1Active(*) => WinIsActive() && S.mode = 1
Mode2Active(*) => WinIsActive() && S.mode = 2
Mode2DurFocus(*) => WinIsActive() && S.mode = 2 && DllCall("GetFocus", "Ptr") = UI.f2d.edit.Hwnd

RegisterWindowHotkeys() {
    HotIf(WinIsActive)
    Hotkey("^1", (*) => SetMode(1))
    Hotkey("^2", (*) => SetMode(2))
    Hotkey("^3", (*) => SetMode(3))
    Hotkey("^+c", (*) => DoCopy())
    Hotkey("^n", (*) => DoNow())
    Hotkey("^t", ToggleOnTop)
    HotIf(Mode1Active)
    Hotkey("^Enter", AddBuild)
    Hotkey("^Backspace", RemoveLast)
    HotIf(Mode2Active)
    Hotkey("^=", (*) => SetSign(1))
    Hotkey("^-", (*) => SetSign(-1))
    HotIf(Mode2DurFocus)
    Hotkey("+=", (*) => SetSign(1))                  ; the "+" key on a US keyboard
    Hotkey("NumpadAdd", (*) => SetSign(1))
    Hotkey("-", (*) => SetSign(-1))
    Hotkey("NumpadSub", (*) => SetSign(-1))
    HotIf()
}

; ------------------------------------------------------------
; 8. TRAY, STARTUP, MAIN
; ------------------------------------------------------------
StartupLink() => A_Startup "\DateTimeCalc.lnk"

StartupEnabled() => FileExist(StartupLink()) != ""

ToggleStartup(name, *) {
    lnk := StartupLink()
    try {
        if StartupEnabled() {
            FileDelete(lnk)
        } else if A_IsCompiled {
            FileCreateShortcut(A_ScriptFullPath, lnk, A_ScriptDir, "/hidden", APP_NAME)
        } else {
            FileCreateShortcut(A_AhkPath, lnk, A_ScriptDir, '"' A_ScriptFullPath '" /hidden', APP_NAME)
        }
    } catch as e {
        MsgBox("Couldn't change the Startup shortcut:`n" e.Message, APP_NAME, "Iconx")
    }
    if StartupEnabled()
        A_TrayMenu.Check(name)
    else
        A_TrayMenu.Uncheck(name)
}

OpenSettings(*) {
    EnsureIni()
    Run('notepad.exe "' INI_FILE '"')
}

BuildTray() {
    A_IconTip := APP_NAME
    m := A_TrayMenu
    m.Delete()
    m.Add("Show/Hide", ToggleWin)
    m.Add("Open settings file", OpenSettings)
    m.Add("Start with Windows", ToggleStartup)
    m.Add()
    m.Add("Reload", (*) => Reload())
    m.Add("Exit", (*) => ExitApp())
    m.Default := "Show/Hide"
    if StartupEnabled()
        m.Check("Start with Windows")
}

; Register the global show/hide hotkey; on failure name the culprit and fall back.
RegisterShowHotkey() {
    hk := Cfg["Hotkey"]
    try {
        Hotkey(hk, ToggleWin)
        return hk
    } catch as e {
        bad := hk
    }
    for fb in [DEFAULT_HOTKEY, "^+!d", "^+!t"] {
        if (fb = bad)
            continue
        try {
            Hotkey(fb, ToggleWin)
            MsgBox("The hotkey '" bad "' (from DateTimeCalc.ini) couldn't be registered:`n" e.Message
                . "`n`nUsing '" fb "' instead. Fix the Hotkey= line in DateTimeCalc.ini and reload.", APP_NAME, "Iconx T20")
            return fb
        }
    }
    MsgBox("Couldn't register the hotkey '" bad "' or any fallback. Use the tray icon to show the window.", APP_NAME, "Iconx T20")
    return ""
}

OnExitHandler(*) {
    if S.shown
        SavePos()
    SetTimer(Tick, 0)
    SetTimer(Recalc, 0)
    SetTimer(ClearStatus, 0)
    return 0
}

Main() {
    LoadSettings()
    S.mode := Integer(Cfg["LastMode"])
    BuildGui()
    BuildTray()
    RegisterWindowHotkeys()
    RegisterShowHotkey()
    OnExit(OnExitHandler)
    ResetLive()
    UI.f2d.edit.Value := ""
    SetSign(1)
    S.hiddenAt := 0
    hidden := false
    for a in A_Args
        if (a = "/hidden")
            hidden := true
    if !hidden
        ShowWin()
}

; Only start the GUI when this file is the script being run (tests.ahk #Includes it).
if (A_LineFile = A_ScriptFullPath)
    Main()
