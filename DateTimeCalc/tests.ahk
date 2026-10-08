#Requires AutoHotkey v2.0
#SingleInstance Force
#Include DateTimeCalc.ahk

; Tests for the pure parsing / time-math / formatting / chain functions.
; Writes test_results.txt next to this script (PASS/FAIL per check, summary at the end).
; Exit code 0 = all passed (SKIPs allowed), 1 = at least one failure.
;
; DST checks need the Windows time zone set to US Pacific; on any other zone
; they are reported as SKIP rather than FAIL.

global Passed := 0, Failed := 0, Skipped := 0, Lines := []

LogLine(s) {
    Lines.Push(s)
}

Check(name, got, expected) {
    global Passed, Failed
    if (got == expected) {
        Passed++
        LogLine("PASS  " name)
    } else {
        Failed++
        LogLine("FAIL  " name "   expected [" expected "] got [" got "]")
    }
}

Skip(name, why) {
    global Skipped
    Skipped++
    LogLine("SKIP  " name "   (" why ")")
}

; ---- helpers ----------------------------------------------------------------
Dur(txt, unit := "h") {
    r := ParseDuration(txt, unit)
    return r.ok ? r.secs : "ERR"
}

DT(txt, now := "20261008143000") {
    r := ParseDateTime(txt, now)
    return r.ok ? r.stamp : "ERR"
}

IsPacific() {
    return OffsetSecs("20260115200000") = -28800 && OffsetSecs("20260715200000") = -25200
}

; ---- duration parsing ---------------------------------------------------------
Check("dur '14h 30m' = 870 min", Dur("14h 30m"), 870 * 60)
Check("dur '36.5h' = 2190 min", Dur("36.5h"), 2190 * 60)
Check("dur '1.5d' = 2160 min", Dur("1.5d"), 2160 * 60)
Check("dur '1w 2d' = 9 days", Dur("1w 2d"), 9 * 86400)
Check("dur '14:30' = 870 min", Dur("14:30"), 870 * 60)
Check("dur '-2h' = -120 min", Dur("-2h"), -120 * 60)
Check("dur bare '8' (default h) = 480 min", Dur("8"), 480 * 60)
Check("dur 'garbage' is an error", Dur("garbage"), "ERR")
Check("dur '2d6h' = 54 h", Dur("2d6h"), 54 * 3600)
Check("dur '90m' = 5400 s", Dur("90m"), 5400)
Check("dur '14:30:15' = 52215 s", Dur("14:30:15"), 52215)
Check("dur '2D 6H' (case-insensitive)", Dur("2D 6H"), 54 * 3600)
Check("dur '1w2d' (no spaces)", Dur("1w2d"), 9 * 86400)
Check("dur bare '90' with unit m = 90 min", Dur("90", "m"), 5400)
Check("dur bare '1.5' = 1.5 h", Dur("1.5"), 5400)
Check("dur '-1d 3h' = -27 h", Dur("-1d 3h"), -27 * 3600)
Check("dur '' is an error", Dur(""), "ERR")
Check("dur '12x' is an error", Dur("12x"), "ERR")
Check("dur '14:75' is an error", Dur("14:75"), "ERR")
Check("dur strict rejects bare number", ParseDuration("8", "h", true).ok, false)

; ---- date/time parsing (now = Thu 2026-10-08 14:30:00) -----------------------------
Check("dt '2026-10-09 18:00'", DT("2026-10-09 18:00"), "20261009180000")
Check("dt '10/9 6pm'", DT("10/9 6pm"), "20261009180000")
Check("dt '10/9/26 18:00'", DT("10/9/26 18:00"), "20261009180000")
Check("dt 'oct 9 6:00 pm'", DT("oct 9 6:00 pm"), "20261009180000")
Check("dt '9 oct 6pm'", DT("9 oct 6pm"), "20261009180000")
Check("dt '6:00 PM' == '18:00'", DT("6:00 PM"), DT("18:00"))
Check("dt '18:00' = today 18:00", DT("18:00"), "20261008180000")
Check("dt '2:30p'", DT("2:30p"), "20261008143000")
Check("dt '2:30 pm'", DT("2:30 pm"), "20261008143000")
Check("dt '230pm'", DT("230pm"), "20261008143000")
Check("dt '1430'", DT("1430"), "20261008143000")
Check("dt '14:30'", DT("14:30"), "20261008143000")
Check("dt 'tomorrow 6am'", DT("tomorrow 6am"), "20261009060000")
Check("dt 'yesterday 14:00'", DT("yesterday 14:00"), "20261007140000")
Check("dt 'tomorrow' = 00:00", DT("tomorrow"), "20261009000000")
Check("dt 'fri 6am' = next Friday", DT("fri 6am"), "20261009060000")
Check("dt 'thu 3pm' = today (still ahead)", DT("thu 3pm"), "20261008150000")
Check("dt 'thu 2pm' = next week (already past)", DT("thu 2pm"), "20261015140000")
Check("dt 'monday 9am'", DT("monday 9am"), "20261012090000")
Check("dt bare '6' = 6 AM today", DT("6"), "20261008060000")
Check("dt '10/9' date only = 00:00", DT("10/9"), "20261009000000")
Check("dt '2026-10-09' date only = 00:00", DT("2026-10-09"), "20261009000000")
Check("dt '12am' = 00:00", DT("12am"), "20261008000000")
Check("dt '12pm' = 12:00", DT("12pm"), "20261008120000")
Check("dt 'now'", DT("now"), "20261008143000")
Check("dt 'n'", DT("n"), "20261008143000")
Check("dt '' = now", DT(""), "20261008143000")
Check("dt '+2h' (no DST in play)", DT("+2h"), "20261008163000")
Check("dt '-1d 3h'", DT("-1d 3h"), "20261007113000")
Check("dt 'feb 30' is an error", DT("feb 30"), "ERR")
Check("dt '25:00' is an error", DT("25:00"), "ERR")
Check("dt '13pm' is an error", DT("13pm"), "ERR")
Check("dt 'banana' is an error", DT("banana"), "ERR")
Check("dt '10/32' is an error", DT("10/32"), "ERR")

; ---- formatting ---------------------------------------------------------------------
Check("FmtDT", FmtDT("20261009154500"), "Fri Oct 9, 2026 · 3:45 PM (15:45)")
Check("FmtDT midnight", FmtDT("20261009000000"), "Fri Oct 9, 2026 · 12:00 AM (00:00)")
Check("FmtDT noon", FmtDT("20261009120500"), "Fri Oct 9, 2026 · 12:05 PM (12:05)")
Check("FmtCopyDT", FmtCopyDT("20261009154500"), "Fri 10/09/2026 3:45 PM (15:45)")
Check("CopyLine Finish", CopyLine("Finish", "20261009154500"), "Finish: Fri 10/09/2026 3:45 PM (15:45)")
Check("FmtShort", FmtShort("20261009060000"), "Fri 10/9 6:00 AM")
Check("FmtDur 187920 s", FmtDur(187920), "2d 4h 12m")
Check("FmtDur negative", FmtDur(-187920), "-2d 4h 12m")
Check("FmtDur zero", FmtDur(0), "0m")
Check("FmtDecHours", FmtDecHours(187920), "52.20 h")
Check("FmtMinutes", FmtMinutes(187920), "3,132 min")
Check("FmtWeeks", FmtWeeks(9 * 86400 + 4 * 3600 + 12 * 60), "1w 2d 4h 12m")
Check("FmtWeeks (spec example shape)", FmtWeeks(7 * 86400 + 4 * 3600 + 12 * 60), "1w 0d 4h 12m")
Check("FmtRel future", FmtRel(14 * 3600 + 20 * 60 + 5), "in 14h 20m")
Check("FmtRel past", FmtRel(-(2 * 3600 + 5 * 60 + 30)), "2h 5m ago")
Check("FmtRel now", FmtRel(20), "now")
Check("FmtLive round-trips through the parser", DT(FmtLive("20261008143100"), "20261008000000"), "20261008143100")
Check("FmtIso round-trips through the parser", DT(FmtIso("20261109073000"), "20261008000000"), "20261109073000")

; ---- difference (no zone dependence: both dates are in PDT) --------------------------------
if true {
    d := DiffElapsed("20261008080000", "20261010121200")
    Check("diff Oct 8 8:00 -> Oct 10 12:12 = 187920 s", d.secs, 187920)
    da := DurAll(d.secs)
    Check("diff main = 2d 4h 12m", da.main, "2d 4h 12m")
    Check("diff dec = 52.20 h", da.dec, "52.20 h")
    Check("diff mins = 3,132 min", da.mins, "3,132 min")
    d2 := DiffElapsed("20261010121200", "20261008080000")
    Check("diff reversed is negative", d2.secs, -187920)
}

; ---- time zone dependent ------------------------------------------------------------------
if !IsPacific() {
    for nm in ["DST fall back (Oct 31 6PM + 12h)", "DST spring forward (Mar 7 6PM + 12h)"
        , "ambiguous 1:30 AM Nov 1", "gap 2:30 AM Mar 8", "chain build 2 start", "diff across fall back"]
        Skip(nm, "Windows time zone is not US Pacific")
} else {
    ; Fall back: Sat Oct 31 2026 6:00 PM PDT + 12 h elapsed = Sun Nov 1 5:00 AM PST
    r := AddElapsed("20261031180000", 12 * 3600)
    Check("DST fall back: finish 5:00 AM Sun Nov 1", r.endLoc, "20261101050000")
    Check("DST fall back: flagged", r.dst.Length, 1)
    Check("DST fall back: delta is -1h (offset change)", r.dst.Length ? r.dst[1] : "", -3600)
    Check("DST fall back: note text", r.notes.Length ? r.notes[1] : "", "⚠ crosses DST change (fall back, +1h)")
    Check("DST fall back: weekday", StampDow(r.endLoc), "Sun")

    ; Spring forward: Sat Mar 7 2026 6:00 PM PST + 12 h = Sun Mar 8 7:00 AM PDT
    r := AddElapsed("20260307180000", 12 * 3600)
    Check("DST spring forward: finish 7:00 AM Sun Mar 8", r.endLoc, "20260308070000")
    Check("DST spring forward: flagged", r.dst.Length, 1)
    Check("DST spring forward: note text", r.notes.Length ? r.notes[1] : "", "⚠ crosses DST change (spring forward, -1h)")

    ; No DST in between -> no flag
    r := AddElapsed("20261008060000", 22 * 3600)
    Check("no DST: finish", r.endLoc, "20261009040000")
    Check("no DST: not flagged", r.dst.Length, 0)

    ; Subtracting across the change
    r := AddElapsed("20261101050000", -12 * 3600)
    Check("subtract back across fall back", r.endLoc, "20261031180000")

    ; Ambiguous local time (fall-back hour): first occurrence = 08:30 UTC
    u := LocalToUtc("20261101013000")
    Check("ambiguous 1:30 AM Nov 1: status", u.status, "ambiguous")
    Check("ambiguous 1:30 AM Nov 1: first occurrence (08:30 UTC)", u.utc, "20261101083000")
    r := AddElapsed("20261101013000", 3600)
    Check("ambiguous start + 1h lands on the 2nd 1:30 AM", r.endLoc, "20261101013000")
    Check("ambiguous start: noted", r.notes.Length >= 1, true)

    ; Nonexistent local time (spring-forward gap)
    u := LocalToUtc("20260308023000")
    Check("gap 2:30 AM Mar 8: status", u.status, "gap")
    Check("gap 2:30 AM Mar 8: moved to 3:30 AM", u.loc, "20260308033000")
    Check("gap 2:30 AM Mar 8: utc = 10:30", u.utc, "20260308103000")

    ; Difference across DST
    d := DiffElapsed("20261031180000", "20261101050000")
    Check("diff across fall back = 12 h elapsed", d.secs, 12 * 3600)
    Check("diff across fall back flagged", d.dst.Length, 1)

    ; Relative offsets are elapsed time, so '+24h' across fall back is 23 wall-clock hours later... check +12h
    Check("dt '+12h' across fall back", DT("+12h", "20261031180000"), "20261101050000")

    ; Chain: build 1 starts 6:00 AM, runs 22 h; changeover 30 min; build 2 = finish + 30 min
    b1 := AddElapsed("20261008060000", 22 * 3600)
    next := ChainNextStart(b1.endUtc, 30)
    b2start := AddElapsed(b1.endLoc, 30 * 60).endLoc
    Check("chain: build 1 finish", b1.endLoc, "20261009040000")
    Check("chain: build 2 start = build 1 finish + 30 min", next, "20261009043000")
    Check("chain: build 2 start matches AddElapsed", next, b2start)
    b2 := AddElapsed(next, 10 * 3600)
    Check("chain: build 2 finish", b2.endLoc, "20261009143000")

    ; Chain across the fall-back change: next start still finish + 30 elapsed minutes
    b1 := AddElapsed("20261031180000", 12 * 3600)
    Check("chain across DST: build 2 start", ChainNextStart(b1.endUtc, 30), "20261101053000")

    ; Chain text
    rows := [{startLoc: "20261008060000", secs: 22 * 3600, endLoc: "20261009040000"}]
    Check("chain copy text", ChainText(rows)
        , "#`tStart`tRun`tFinish`r`n1`tThu 10/08/2026 6:00 AM (06:00)`t22h`tFri 10/09/2026 4:00 AM (04:00)")
}

; ---- chain logic that doesn't depend on the zone -------------------------------------------
Check("chain: zero changeover", ChainNextStart(LocalToUtcSafe("20261009040000"), 0) != "", true)

LocalToUtcSafe(stamp) {
    return LocalToUtc(stamp).utc
}

; ---- write results -------------------------------------------------------------------------------
out := ""
for l in Lines
    out .= l "`r`n"
out .= "`r`n" Passed " passed, " Failed " failed, " Skipped " skipped`r`n"
resFile := A_ScriptDir "\test_results.txt"
try FileDelete(resFile)
FileAppend(out, resFile, "UTF-8")
ExitApp(Failed ? 1 : 0)
