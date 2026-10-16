# P110 Powder Tracker v20 — reference model (test oracle). Pure Python, no Excel.
#
# Physical model (slot-independent):
#   * The machine holds up to two buckets (slot 1, slot 2) feeding one shared funnel.
#   * Removing a bucket (swap or EMPTY) puts its leftover into the funnel, up to funnel capacity.
#   * Powder is consumed: funnel first (oldest leftover first), then the bucket that has been on
#     the machine longest, then the other bucket. Ties are broken by bucket key, never by slot number,
#     so putting a bucket in slot 1 or slot 2 gives identical results.
#
# Verbatim from the v20 build prompt §2.4. Only change: the NBSP literal in norm() is written as
# " " (the prompt's copy rendered it as a plain space, which would make the replace a no-op).

MAX_FUNNEL_LOTS = 4
EPS = 1e-9


def norm(x):
    """Normalize an input cell: None -> "", strip non-printing chars and NBSP, trim, upper-case."""
    if x is None:
        return ""
    s = str(x).replace(" ", " ")
    s = "".join(ch for ch in s if ord(ch) >= 32)
    return s.strip().upper()


def run_machine(rows, full, funnel_cap, roller_total):
    """rows: list of dicts with keys build, date, clean, s1c, s1b, s2c, s2b, usage (raw cell values).
    full: {bucket_key: full weight kg} for filled buckets.
    Returns (rows_out, draws, issues)
      draws  = [(row_index, bucket_key, kg, source)]  source in {"funnel", "slot1", "slot2"}
      issues = [(row_index, "ERROR" | "WARN", text)]"""
    out, draws, issues = [], [], []
    cyc = {"1": "", "2": ""}
    bkt = {"1": "", "2": ""}
    key = {"1": "", "2": ""}
    loaded_at = {"1": None, "2": None}   # row index when the current bucket was put on the machine
    funnel = []                          # FIFO list of [bucket_key, kg], oldest first

    def remaining(b):
        return full.get(b, 0.0) - sum(a for _, bb, a, _ in draws if bb == b)

    for i, row in enumerate(rows):
        active = norm(row["build"]) != ""
        prev_key, prev_loaded = dict(key), dict(loaded_at)

        # 1) Effective slot keys: per-field carry-forward (v19), continues through blank rows.
        for s in ("1", "2"):
            c, b = norm(row["s%sc" % s]), norm(row["s%sb" % s])
            if c == "EMPTY":
                cyc[s], bkt[s] = "", ""
            else:
                if c:
                    cyc[s] = c
                if b:
                    bkt[s] = b
            key[s] = cyc[s] + bkt[s] + "-" if cyc[s] and bkt[s] else ""
            if key[s] != prev_key[s]:
                loaded_at[s] = i if key[s] else None

        # 2) Removed buckets: leftover goes into the shared funnel, longest-on-machine first.
        removed = sorted([s for s in ("1", "2") if prev_key[s] and key[s] != prev_key[s]],
                         key=lambda s: (prev_loaded[s], prev_key[s]))
        for s in removed:
            b = prev_key[s]
            want = max(0.0, remaining(b))
            room = max(0.0, funnel_cap - sum(kg for _, kg in funnel))
            put = min(want, room)
            if want - put > EPS:
                issues.append((i, "WARN", "Funnel full: %.3f kg of %s did not fit (stays on bucket)" % (want - put, b)))
            if put > EPS:
                if len(funnel) >= MAX_FUNNEL_LOTS:
                    issues.append((i, "ERROR", "More than %d bucket leftovers in funnel" % MAX_FUNNEL_LOTS))
                funnel.append([b, put])
        for s in ("1", "2"):
            if key[s] and key[s] != prev_key[s] and any(fb == key[s] for fb, _ in funnel):
                issues.append((i, "ERROR", "%s loaded while its leftover is still in the funnel" % key[s]))

        funnel_before = [tuple(x) for x in funnel]
        avail = {s: (remaining(key[s]) if key[s] else None) for s in ("1", "2")}

        # 3) Consumption for this build.
        u = row["usage"]
        usage = u if isinstance(u, (int, float)) and not isinstance(u, bool) else None
        total, funnel_used, slot_used = None, [], {"1": None, "2": None}
        if active and usage is not None:
            total = usage + (roller_total if norm(row["clean"]) == "Y" else 0.0)
            need = total
            for lot in funnel:                                   # funnel, oldest leftover first
                take = min(lot[1], need)
                if take > 0:
                    funnel_used.append((lot[0], take))
                    lot[1] -= take
                    need -= take
            funnel = [lot for lot in funnel if lot[1] > EPS]
            on_machine = sorted([s for s in ("1", "2") if key[s]], key=lambda s: (loaded_at[s], key[s]))
            slot_used = {"1": 0.0, "2": 0.0}
            if need < EPS:                                       # snap float residue (never record ~0 draws)
                need = 0.0
            for j, s in enumerate(on_machine):                   # longest-on-machine bucket first
                last = j == len(on_machine) - 1
                take = need if last else min(max(0.0, avail[s]), need)
                slot_used[s] = take
                need -= take
                if need < EPS:
                    need = 0.0
            if need > EPS:
                issues.append((i, "ERROR", "%.3f kg used with no bucket loaded" % need))
            if on_machine and avail[on_machine[-1]] - slot_used[on_machine[-1]] < -EPS:
                issues.append((i, "ERROR", "Overdraw: %s goes negative" % key[on_machine[-1]]))
            for b, a in funnel_used:
                draws.append((i, b, a, "funnel"))
            for s in ("1", "2"):
                if key[s] and slot_used[s] > EPS:
                    draws.append((i, key[s], slot_used[s], "slot" + s))

        slot_left = {s: (avail[s] - (slot_used[s] or 0.0)) if key[s] else None for s in ("1", "2")}
        funnel_after = sum(kg for _, kg in funnel)
        machine_left = sum(v for v in slot_left.values() if v is not None) + funnel_after
        out.append(dict(active=active, k1=key["1"], k2=key["2"], total=total,
                        funnel_before=funnel_before, funnel_used=funnel_used,
                        s1u=slot_used["1"], s2u=slot_used["2"],
                        slot1_left=slot_left["1"], slot2_left=slot_left["2"],
                        funnel_after=funnel_after, machine_left=machine_left,
                        left=machine_left if total is not None else None))
    return out, draws, issues
