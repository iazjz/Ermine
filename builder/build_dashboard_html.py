#!/usr/bin/env python3
"""
A phone-sized reading of the live workbook, as one self-contained HTML page.

Why this exists: the workbook is thirty-one sheets and twenty thousand formulas. On a
laptop that is the point; on a phone it is unusable, and "what is happening right now" is
a question Aziz asks far more often than he sits down to plan. So the page carries no
calculation of its own - every figure is lifted out of the recalculated workbook exactly
as the file states it, and the page is a view, not a second source of truth. If a number
here disagrees with the file, the page is wrong and the file is right.

It is a SNAPSHOT. Excel on OneDrive has no API a web page can read, so there is no live
link; the page is rebuilt whenever the workbook is. The build stamp at the top says when,
and says it in hours rather than a timestamp, because "4 hours old" is a judgement you can
act on and "14:32" is one you have to make yourself.

    python3 build_dashboard_html.py [RECALCULATED.xlsx] [OUT.html]
"""
import datetime as dt
import glob
import html
import json
import os
import re
import subprocess
import sys
import tempfile

import openpyxl

import ermine_paths

# Every clock this builder writes is Riyadh time, said as such. The stamps used to be
# naive - fine on Aziz's Mac, whose clock is Riyadh, but three hours behind whenever a
# build ran on a machine set to UTC, because the browser reads a naive time as its own.
try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Asia/Riyadh")
except Exception:
    TZ = dt.timezone(dt.timedelta(hours=3))


def now_local():
    return dt.datetime.now(TZ)  # finds the files by name, wherever Aziz files them


def _ep_one(name):
    """Kept returning a list so the existing `hits[0] if hits else None`
    callers did not all have to change on the same afternoon."""
    p = ermine_paths.find(name)
    return [p] if p else []


HERE = os.path.dirname(os.path.abspath(__file__))


def live_path():
    hits = sorted(_ep_one("Production Plan - LIVE.xlsx"))
    return hits[0] if hits else None


def materials_path():
    hits = sorted(_ep_one("Raw Materials - Ali Nawaz.xlsx"))
    return hits[0] if hits else None


def tasks_path():
    hits = sorted(_ep_one("Tasks - Aziz.xlsx"))
    return hits[0] if hits else None


def warehouse_path():
    """Ali's file, read straight rather than through the Stores sheet.

    Stores carries his totals but not his lines, and the day's feed needs the lines -
    what moved, which box, when. Read-only, and a missing file is not an error: the
    feed simply has the chef's side only.
    """
    hits = sorted(_ep_one("Warehouse - Ali Nawaz.xlsx"))
    return hits[0] if hits else None


def ali_store(path):
    """Ali's own box balances, read from his file rather than from a copy of it.

    The Stores sheet in the production workbook has the same columns and is where the
    page has been reading them from. It is filled by sync_store.py, which WRITES to the
    chef's workbook - so on 19 October, when refreshes were made read-only to stop a
    script overwriting work somebody had open, that sync stopped with them. Nobody
    noticed, because a stale number looks exactly like a current one. Five weeks of
    movements went missing: the page said 380 Mini in the store against Ali's 11, and
    was still quoting a stocktake two weeks older than the one he had done.

    So this reads his file directly and hands back what he actually wrote. The chef's
    side of the same table - what he packed, what sales asked for - stays where it is,
    because those columns are live formulas and were never the problem.

    Ali keeps Half and Fingers split by flavour where the rest of the system treats
    each as one line, so those are added back together here.
    """
    if not path:
        return None
    try:
        ws = openpyxl.load_workbook(path, data_only=True)["Warehouse"]
    except Exception:
        return None

    rows, head = {}, None
    for r in range(8, 16):                       # the header has moved before now
        if str(ws.cell(r, 7).value or "").strip().lower() == "box":
            head = r
            break
    if head is None:
        return None

    # The in/out/waste columns are SUMIFS over the movement log. openpyxl drops the
    # results Excel had cached the moment anything saves the file with it, so reading
    # them gives None and the whole store reads as zero - which is what happened the
    # first time this script touched Ali's file. His raw lines are not formulas and
    # cannot be lost that way, so they are totalled here instead and the cached values
    # are used only to cross-check. Same answer, one fewer thing that can go empty.
    def key_of(nm):
        return ("Half" if nm.startswith("Half") else
                "Fingers" if nm.startswith("Fingers") else nm)

    moved, by_flavour = {}, {}
    for r in range(head + 1, head + 300):
        when, what, box, n = (ws.cell(r, c).value for c in (1, 2, 3, 4))
        if not isinstance(box, str) or not box.strip():
            if not isinstance(when, (dt.date, dt.datetime)):
                continue
        if not isinstance(box, str) or not isinstance(n, (int, float)):
            continue
        w = str(what or "").upper()
        kind = "in" if "IN" in w else "out" if "OUT" in w else "waste" if "WASTE" in w else None
        if not kind:
            continue
        box = box.strip()
        moved.setdefault(key_of(box), {"in": 0, "out": 0, "waste": 0})[kind] += n
        # Half and Fingers are one line everywhere else in the system, because the
        # target, the order book and the plan all treat them as one. But a store is a
        # shelf: "411 Half" does not tell anyone which filling is on it, and sales
        # cannot pick from a total. So the flavour lines are kept alongside the box
        # they roll up into, rather than discarded on the way.
        if box != key_of(box):
            f = by_flavour.setdefault(key_of(box), {}).setdefault(
                box, {"stocktake": 0, "in": 0, "out": 0, "waste": 0})
            f[kind] += n

    for r in range(head + 1, head + 40):
        nm = ws.cell(r, 7).value
        if not isinstance(nm, str) or not nm.strip():
            break
        nm = nm.strip()
        if nm.upper() == "TOTAL":                # the total is a row, not a box
            break
        key = key_of(nm)
        e = rows.setdefault(key, {"stocktake": 0, "in": 0, "out": 0, "waste": 0})
        st = ws.cell(r, 8).value                 # typed by hand, never a formula
        st = st if isinstance(st, (int, float)) else 0
        e["stocktake"] += st
        if nm != key:
            f = by_flavour.setdefault(key, {}).setdefault(
                nm, {"stocktake": 0, "in": 0, "out": 0, "waste": 0})
            f["stocktake"] += st

    for key, m in moved.items():
        e = rows.setdefault(key, {"stocktake": 0, "in": 0, "out": 0, "waste": 0})
        for k in ("in", "out", "waste"):
            e[k] += m[k]

    taken = ws.cell(6, 2).value
    if isinstance(taken, dt.datetime):
        taken = taken.date()
    return {"rows": rows, "by_flavour": by_flavour,
            "stocktake": taken if isinstance(taken, dt.date) else None}


def soffice():
    """Wherever LibreOffice is, if it is anywhere."""
    import shutil as _sh
    for c in ("soffice", "libreoffice",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.isabs(c):
            if os.path.exists(c):
                return c
        elif _sh.which(c):
            return c
    return None


def has_values(path):
    """Does the file already carry its formula results?

    Excel stores the last computed value beside every formula, so a workbook a person
    saved is readable without recalculating anything. A workbook last written by a
    script is not: openpyxl drops those cached values, and every formula cell reads
    empty. This is the difference between the two, asked rather than assumed.
    """
    # A script that patches the XML keeps the OLD cached results beside the new inputs
    # and marks the file fullCalcOnLoad. Those cached results are stale, so a file
    # carrying that mark is recalculated whatever it seems to hold.
    try:
        import zipfile
        with zipfile.ZipFile(path) as z:
            if b'fullCalcOnLoad="1"' in z.read("xl/workbook.xml"):
                return False
    except Exception:
        pass
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["Dashboard"] if "Dashboard" in wb.sheetnames else wb.worksheets[0]
        seen = 0
        for row in ws.iter_rows(min_row=1, max_row=40):
            for c in row:
                if c.value is not None:
                    seen += 1
        return seen > 5
    except Exception:
        return False


def recalc(path):
    """A recalculated copy - or the file itself when it does not need one.

    LibreOffice is the slow, heavy part of every build, and on Aziz's Mac it may not be
    installed at all. It is only needed when the workbook has no cached results, which
    happens after a script writes it. When a person saved it last - the normal case -
    the figures are already in there and this step is skipped entirely.
    """
    if os.environ.get("ERMINE_NO_RECALC") == "1" or has_values(path):
        return path
    exe = soffice()
    if not exe:
        print("  LibreOffice not found and the workbook carries no cached results.")
        print("  Open it in Excel once and save, or install LibreOffice.")
        return None
    out = os.path.join(tempfile.gettempdir(), "dash_recalc")
    subprocess.run(f'rm -rf "{out}" && mkdir -p "{out}" && ulimit -v 2500000; '
                   f'"{exe}" --headless --norestore --convert-to '
                   f'xlsx:"Calc MS Excel 2007 XML" --outdir "{out}" "{path}"',
                   shell=True, capture_output=True)
    got = os.path.join(out, os.path.basename(path))
    return got if os.path.exists(got) else None


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def nums(s):
    """every number in a sentence, in order, commas and decimals understood"""
    return [float(x.replace(",", "")) for x in re.findall(r"-?[\d,]*\.?\d+", str(s or ""))]


def _fin_hm(v):
    if isinstance(v, dt.datetime): return v.strftime("%H:%M")
    if isinstance(v, dt.time): return v.strftime("%H:%M")
    if isinstance(v, (int, float)):
        m = round((v % 1) * 1440); return "%02d:%02d" % (m // 60 % 24, m % 60)
    return None


def _fin_days(v):
    if isinstance(v, dt.datetime) and v.year < 1901:
        return (v.date() - dt.date(1899, 12, 31)).days
    if isinstance(v, (int, float)): return int(v)
    return 0


def code(s, *pairs, default="other"):
    """Classify one of the workbook's English sentences into a token.

    The page is bilingual and the workbook is not: it says "fits - 77 hours spare" and
    "NEEDS OVERTIME - 2.3 extra hours", and switching the page to Arabic left those
    sentences sitting there in English because they are data, not labels. Translating the
    STRING would mean shipping a phrasebook that breaks the moment the workbook rewords
    itself. So the sentence is reduced here to a token and its numbers, and each language
    writes its own sentence from those. The original is kept as a fallback for anything
    this does not recognise, which is better than showing nothing.
    """
    t = str(s or "").upper()
    for token, *words in pairs:
        if any(w.upper() in t for w in words):
            return token
    return default


def read(path):
    """Everything the page shows, lifted out of the recalculated workbook.

    The list is Aziz's, not mine, and it is deliberately short: this page answers "is the
    day going right and what do I do about it", and anything that does not help that
    question was taken off it. Reference material stays in the workbook.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    D, MT, ST, SR, SP, AL, DP, CAL = (wb["Dashboard"], wb["Monthly Targets"], wb["Stock"],
                                      wb["Stores"], wb["Sales Plan"], wb["Actual Log"],
                                      wb["Daily Plan"], wb["Calibration"])
    d = {}

    def find(ws, text, col=1, upto=90):
        for r in range(1, upto):
            v = ws.cell(r, col).value
            if isinstance(v, str) and v.strip().startswith(text):
                return r
        return None

    # ---- 1. the chef's plan for today, and when it should finish ----
    _plan16 = nums(D["C16"].value)          # "9440 pcs to make / 531 boxes to pack"
    _verd = str(D["C17"].value or "")
    _note = str(D["D19"].value or "")       # "(10.3 h of work + break, 25 staff)"
    _n19 = nums(_note)
    # The arithmetic behind "needs overtime", in numbers rather than in a sentence, so
    # the page can show what is being asked for and what declining it would cost.
    _ce = wb["Capacity Engine"]
    _lab = num(_ce["B66"].value)          # labour hours the day needs, all staff
    _std = num(_ce["B7"].value)           # productive hours in a standard day, per head
    _stf = num(_ce["B8"].value)           # people counted for today
    _cap = (_std or 0) * (_stf or 0)
    d["day"] = {
        "labour_h": _lab, "standard_h": _std, "staff_n": _stf,
        "capacity_h": _cap or None,
        "per_person_h": num(_ce["B77"].value),
        "extra_staff_h": round(max(0.0, (_lab or 0) - _cap), 1) if _lab else None,
        # 10.6 SAR per person-hour is the role-level average and the only wage figure in
        # this system; no individual's pay is stored or derivable from it.
        "extra_cost": round(max(0.0, (_lab or 0) - _cap) * 10.6) if _lab else None,
    }
    # where the day's hours actually go, so a refusal can be aimed at something
    steps = []
    for r in range(17, 33):
        nm = _ce.cell(r, 1).value
        hv = num(_ce.cell(r, 15).value)
        if isinstance(nm, str) and nm.strip() and hv and hv > 0:
            steps.append({"step": nm.strip(), "hours": round(hv, 1)})
    steps.sort(key=lambda x: -x["hours"])
    d["steps"] = steps

    d["today"] = {
        "date": ST["B3"].value,
        "committed": str(D["C14"].value or ""),
        # Excel keeps a finish past midnight as 1900-01-01 00:35 (a time over 24 h);
        # read as a date it printed "1900-". Send the clock time and the day offset.
        "finish": _fin_hm(D["C18"].value),
        "finish_days": _fin_days(D["C18"].value),
        "finish_note": _note,
        "verdict": _verd,
        # the same three facts as numbers, so Arabic can say them in Arabic
        "pcs": _plan16[0] if len(_plan16) > 0 else None,
        "boxes": _plan16[1] if len(_plan16) > 1 else None,
        "verdict_code": code(_verd,
                             ("toomuch", "TOO MUCH", "DOES NOT FIT", "OUT OF REACH"),
                             ("overtime", "OVERTIME"),
                             ("ok", "ON TRACK", "FITS", "OK"), default="other"),
        "verdict_h": (nums(_verd) or [None])[0],
        "day_hours": _n19[0] if len(_n19) > 0 else None,
        "day_staff": _n19[1] if len(_n19) > 1 else None,
    }
    # ---- 2. the three suggested plans, and what each costs in wages ----
    RATE = 10.6                                     # one floor person-hour, Sept payroll
    # An overtime hour, same basis. Not a guess: the 21 Sep - 20 Oct payroll run came
    # to 9,439 SAR against 688 overtime person-hours in the Actual Log, which is
    # 13.71 an hour - a 1.29x premium on the 10.60 above. Two numbers from two
    # different places agreeing to within a few percent is the only reason to trust
    # either of them. Like RATE it is a role-level average; no individual's pay is
    # stored here or derivable from it.
    OT_RATE = 13.7
    plans = []
    for i2, label in enumerate(["A", "B", "IDEAL"]):
        r = 7 + i2
        hrs = num(ST.cell(r, 4).value)
        _f = str(ST.cell(r, 6).value or "")      # "fits - 77 hours spare" / "over by 6 hours"
        plans.append({"name": label, "what": str(ST.cell(r, 1).value or ""),
                      "pcs": num(ST.cell(r, 2).value), "boxes": num(ST.cell(r, 3).value),
                      "hours": hrs, "people": num(ST.cell(r, 5).value),
                      "cost": round(hrs * RATE) if hrs else None,
                      "fits": _f,
                      "fits_code": code(_f, ("over", "OVER BY", "DOES NOT FIT"),
                                        ("fits", "FITS"), default="other"),
                      "fits_h": (nums(_f) or [None])[0]})
    d["plans"] = plans
    d["rate"] = RATE
    d["available_h"] = num(ST["D70"].value)      # hours the floor actually has today

    # ---- 2b. the suggestion, product by product -------------------------------
    # The Stock sheet works out A, B and IDEAL per product in its "SUGGESTED FOR TODAY"
    # block. Only the three totals reached the page; the per-product IDEAL - what to
    # make so the orders ship and the month stays on pace, counting stock - did not.
    # Found by its heading and its column names, not by address.
    sg = None
    for r in range(30, 140):
        v = ST.cell(r, 1).value
        if isinstance(v, str) and v.strip().upper().startswith("SUGGESTED FOR TODAY"):
            sg = r
            break
    srows, spack = [], []
    if sg:
        hd = {}
        for c in range(1, 14):
            v = ST.cell(sg + 1, c).value
            if isinstance(v, str) and v.strip():
                hd[v.strip().split("\n")[0].strip().lower()] = c

        def _h(*names):
            for k, c in hd.items():
                if any(k.startswith(n) for n in names):
                    return c
            return None
        cP, cA, cD, cC, cI = (_h("product"), _h("plan a"), _h("what is driving"),
                              _h("chef has planned"), _h("ideal"))
        cB = _h("plan b")
        cBx, cPk = _h("box"), _h("pack today")
        for r in range(sg + 2, sg + 30):
            pv = ST.cell(r, cP).value if cP else None
            bv = ST.cell(r, cBx).value if cBx else None
            if isinstance(pv, str) and pv.strip().lower().startswith("does it fit"):
                break
            if not (isinstance(pv, str) and pv.strip()) and not (isinstance(bv, str) and bv.strip()):
                break
            if isinstance(pv, str) and pv.strip():
                srows.append({"product": pv.strip(),
                              "a": num(ST.cell(r, cA).value) if cA else None,
                              "b": num(ST.cell(r, cB).value) if cB else None,
                              "ideal": num(ST.cell(r, cI).value) if cI else None,
                              "chef": num(ST.cell(r, cC).value) if cC else None,
                              "driver": str(ST.cell(r, cD).value or "").strip() if cD else ""})
            if isinstance(bv, str) and bv.strip():
                spack.append({"box": bv.strip(),
                              "pack": num(ST.cell(r, cPk).value) if cPk else None})
    d["sugg"] = {"rows": srows, "pack": spack}

    # ---- 3. the day that was actually worked, and how far it drifted ----
    log = []
    for r in range(15, 135):
        day = AL.cell(r, 1).value
        if not isinstance(day, (dt.datetime, dt.date)) or not num(AL.cell(r, 2).value):
            continue
        made = sum(v for c in range(11, 21) if (v := num(AL.cell(r, c).value)))
        packed = sum(v for c in range(21, 36) if (v := num(AL.cell(r, c).value)))
        log.append({"date": day, "staff": num(AL.cell(r, 2).value),
                    "hours": num(AL.cell(r, 6).value), "ratio": num(AL.cell(r, 8).value),
                    "reading": str(AL.cell(r, 9).value or ""),
                    "made": made, "packed": packed,
                    "note": str(AL.cell(r, 10).value or ""),
                    # kept in the history so a month can still be reported after the
                    # Actual Log's 120-day calendar has rolled past it
                    "by_prod": {str(AL.cell(14, c).value).strip(): v for c in range(11, 21)
                                if (v := num(AL.cell(r, c).value)) and AL.cell(14, c).value},
                    "by_box": {str(AL.cell(14, c).value).strip(): v for c in range(21, 36)
                               if (v := num(AL.cell(r, c).value)) and AL.cell(14, c).value},
                    "ot": num(AL.cell(r, 61).value)})
    # What each PRODUCT made, day by day. The flavour snapshot records what each
    # FLAVOUR made, and the chef does not always fill it in - so this is what the
    # missing figure gets checked against, and in the unambiguous case recovered from.
    prod_by_day = {}
    for r in range(15, 200):
        dd = AL.cell(r, 1).value
        dd = dd.date() if isinstance(dd, dt.datetime) else dd
        if not isinstance(dd, dt.date) or not AL.cell(r, 2).value:
            continue
        got = {}
        for c in range(11, 21):
            nm = AL.cell(14, c).value
            v = num(AL.cell(r, c).value)
            if isinstance(nm, str) and nm.strip() and v:
                got[nm.strip()] = v
        if got:
            prod_by_day[dd.isoformat()] = got

    d["log"] = log[-14:]          # what the shift page draws
    _log_all = log                # every day the sheet holds, for the pay-run sums
    d["last"] = log[-1] if log else None

    # ---- 3b. delays: what each closed day missed, how far the model was off, and why ----
    # A day is "late" by the file's own rule (Lists BK4601): |made - plan| / plan >= 10%,
    # pieces only, plan as captured into the Actual Log. The reasons are the chef's
    # lines on Delays & Reasons, matched by date. The model check is the Actual Log's
    # ratio (model hours / actual hours) outside 0.85-1.15 - its own "within 15%".
    _dr = {}
    try:
        DRS = wb["Delays & Reasons"]
        for r in range(6, 300):
            dd = DRS.cell(r, 1).value
            dd = dd.date() if isinstance(dd, dt.datetime) else dd
            if not isinstance(dd, dt.date):
                continue
            it = {k: str(DRS.cell(r, c).value).strip() if DRS.cell(r, c).value not in (None, "") else ""
                  for k, c in (("prod", 2), ("where", 3), ("cat", 4), ("reason", 5), ("what", 6))}
            if any(it.values()):
                _dr.setdefault(dd.isoformat(), []).append(it)
    except KeyError:
        pass
    _pn = [str(AL.cell(14, c).value or "").strip() for c in range(11, 21)]
    _today = now_local().date()
    delays = []
    for r in range(15, 135):
        dd = AL.cell(r, 1).value
        dd = dd.date() if isinstance(dd, dt.datetime) else dd
        if not isinstance(dd, dt.date) or dd > _today:
            continue
        k = dd.isoformat()
        rs = _dr.get(k, [])
        if dd == _today and not rs:
            continue                      # today is not closed yet
        plan = {n: (num(AL.cell(r, 25 + i + 11).value) or 0) for i, n in enumerate(_pn) if n}
        made = {n: (num(AL.cell(r, 11 + i).value) or 0) for i, n in enumerate(_pn) if n}
        P, M = sum(plan.values()), sum(made.values())
        rec = bool(num(AL.cell(r, 2).value))
        if not P and not rs and not rec:
            continue
        miss = (abs(M - P) / P) if P else None
        ratio = num(AL.cell(r, 8).value)
        prods = [{"p": n, "plan": plan[n], "made": made[n]} for n in plan
                 if plan[n] or made[n]]
        delays.append({
            "date": k, "plan": P or None, "made": M if rec else None, "recorded": rec,
            "miss": round(miss, 3) if miss is not None else None,
            "late": bool(miss is not None and miss >= 0.10),
            "ratio": round(ratio, 3) if ratio else None,
            "model_off": bool(ratio and (ratio < 0.85 or ratio > 1.15)),
            "hours": num(AL.cell(r, 6).value), "staff": num(AL.cell(r, 2).value),
            "note": str(AL.cell(r, 10).value or "").strip(),
            "prods": prods, "reasons": rs})
    d["delays"] = delays
    _lb = {daykey(x["date"]): x for x in log}
    for x in delays:
        if x["date"] in _lb:
            _lb[x["date"]]["reasons"] = x["reasons"]
            if x["plan"]:
                _lb[x["date"]]["plan_al"] = x["plan"]

    # ---- 3c. the stations, as the Station Sheets lay them out --------------------
    # Seven work areas, each a 40-row block: who the chef put there (AM/PM), the hours
    # the area is busy, its load, and the tasks with product, quantity, hours and crew.
    stations = []
    try:
        SS = wb["Station Sheets"]
        for top in range(1, 400, 40):
            title = SS.cell(top, 1).value
            if not (isinstance(title, str) and title.upper().startswith("STATION")):
                continue
            names = {}
            for half, rr in (("am", top + 2), ("pm", top + 3)):
                raw = str(SS.cell(rr, 2).value or "")
                lst = [] if "nobody" in raw.lower() else [x.strip() for x in raw.split(" / ") if x.strip()]
                names[half] = lst
            tasks, rr = [], top + 6
            while rr < top + 40:
                c = SS.cell(rr, 3).value
                if isinstance(c, str) and c.strip().upper().startswith("TOTAL"):
                    break
                if isinstance(c, str) and c.strip():
                    tasks.append({"group": str(SS.cell(rr, 2).value or ""), "task": c.strip(),
                                  "prod": str(SS.cell(rr, 4).value or ""),
                                  "qty": num(SS.cell(rr, 5).value),
                                  "h": num(SS.cell(rr, 6).value), "crew": num(SS.cell(rr, 7).value)})
                rr += 1
            wait, tied = [], None
            for r2 in range(rr, top + 40):
                c = SS.cell(r2, 3).value
                if isinstance(c, str) and c.strip().lower().startswith("room also tied up"):
                    tied = num(SS.cell(r2, 6).value)
                c = SS.cell(r2, 3).value
                if isinstance(c, str) and c.strip() and c.strip() != "Task" and num(SS.cell(r2, 5).value):
                    wait.append({"task": c.strip(), "prod": str(SS.cell(r2, 4).value or ""),
                                 "min": num(SS.cell(r2, 5).value)})
            stations.append({"title": title.split("-", 1)[-1].strip(), "no": len(stations) + 1,
                             "busy": num(SS.cell(top + 1, 4).value),
                             "load": num(SS.cell(top + 1, 6).value),
                             "am": names["am"], "pm": names["pm"],
                             "tasks": tasks, "wait": wait, "tied": tied})
    except KeyError:
        pass
    d["stations"] = stations

    # ---- 4. what is in the factory ----
    pieces = []
    pr = find(D, "IN THE FACTORY NOW", 2)
    if pr:
        for r in range(pr + 2, pr + 14):
            nm = D.cell(r, 2).value
            if not isinstance(nm, str) or not nm.strip() or nm.strip() == "TOTAL":
                break
            pieces.append({"product": nm.strip(), "pcs": num(D.cell(r, 5).value) or 0})
    d["pieces"] = pieces

    # ---- 5. boxes: the store, what came in from the chef, what went out to sales ----
    boxes, tot = [], {"store": 0, "recv": 0, "out": 0, "asked": 0, "owed": 0,
                      "target": 0, "packed": 0, "left": 0}
    for r in range(13, 28):
        b = MT.cell(r, 1).value
        if not b:
            continue
        row = {"box": str(b),
               "target": num(MT.cell(r, 3).value), "packed": num(MT.cell(r, 4).value),
               "left": num(MT.cell(r, 5).value)}
        boxes.append(row)
    sr = {}
    for r in range(7, 25):
        b = SR.cell(r, 2).value
        if isinstance(b, str) and b.strip() and b.strip() != "TOTAL":
            sr[b.strip()] = {"recv": num(SR.cell(r, 5).value), "gap": num(SR.cell(r, 6).value),
                             "out": num(SR.cell(r, 7).value), "store": num(SR.cell(r, 9).value),
                             "asked": num(SR.cell(r, 10).value), "owed": num(SR.cell(r, 12).value)}

    # Ali's own file, read once and used for both tables that carry his columns. See
    # ali_store: the Stores sheet these came from has been a frozen copy since the
    # sync that filled it was switched off.
    _ali = ali_store(warehouse_path())
    _arows = (_ali or {}).get("rows") or {}

    # The gap is "what the chef packed and the store has not got", and the two sides
    # have to cover the SAME days or it is not a gap at all. Ali's columns restart at
    # his stocktake; the chef's run from the first of the month. Comparing them gave
    # Origin a gap of 5,929 boxes - a number that would have had somebody searching a
    # warehouse for stock that was counted, not lost. So the chef's side is re-totalled
    # from the Actual Log over Ali's window only.
    _since = (_ali or {}).get("stocktake")
    _packed_since = {}
    if _since:
        _pcol = {}
        for c in range(21, 36):
            nm = AL.cell(14, c).value
            if isinstance(nm, str) and nm.strip():
                _pcol[nm.strip()] = c
        for r in range(15, 200):
            dd = AL.cell(r, 1).value
            dd = dd.date() if isinstance(dd, dt.datetime) else dd
            if not isinstance(dd, dt.date) or dd < _since:
                continue
            for nm, c in _pcol.items():
                v = num(AL.cell(r, c).value)
                if v:
                    _packed_since[nm] = _packed_since.get(nm, 0) + v
    for x in boxes:
        x.update(sr.get(x["box"], {}))
        a = _arows.get(x["box"])
        if a:
            x["recv"] = a["in"]
            x["out"] = a["out"]
            x["store"] = a["stocktake"] + a["in"] - a["out"] - a["waste"]
            if _since is not None:
                x["gap"] = _packed_since.get(x["box"], 0) - a["in"]
                x["packed_since"] = _packed_since.get(x["box"], 0)
            elif x.get("packed") is not None:
                x["gap"] = x["packed"] - a["in"]
            if x.get("asked") is not None:
                x["owed"] = x["asked"] - a["out"]
        for k in tot:
            tot[k] += x.get(k) or 0
    d["boxes"] = boxes
    d["tot"] = tot

    # Set from the Stores sheet here and replaced by Ali's own date below, if his
    # file has one. His is the real answer to "when was this counted".
    d["stocktake"] = SR["C3"].value

    # ---- 6. the month ----
    hrs = find(MT, "Labour hours needed")
    _mv = str(MT.cell(hrs + 3, 2).value or "") if hrs else ""
    d["month"] = {"month": MT["B4"].value, "days_left": num(MT["B7"].value),
                  "hours_day": num(MT.cell(hrs, 2).value) if hrs else None,
                  "hours_total": num(MT.cell(hrs, 5).value) if hrs else None,
                  "verdict": _mv,
                  "verdict_code": code(_mv, ("outofreach", "OUT OF REACH", "TOO MUCH"),
                                       ("overtime", "OVERTIME"),
                                       ("ok", "ON TRACK", "FITS", "OK"), default="other"),
                  "verdict_n": (nums(_mv) or [None])[0],
                  "sens": [{"n": num(MT.cell(hrs + 1, c).value),
                            "load": num(MT.cell(hrs + 2, c).value)}
                           for c in range(5, 9)] if hrs else []}

    # ---- 7. who is on the floor today ----
    # The Actual Log records a COUNT, not names. The names are on the chef's own roster,
    # where a person with an area against them is a person he has placed today - which is
    # the nearest thing the file holds to an attendance list, and it is his, not a guess.
    names = []
    rr = find(DP, "Staff Member")
    if rr:
        for r in range(rr + 1, rr + 40):
            nm = DP.cell(r, 1).value
            if not isinstance(nm, str) or not nm.strip():
                break
            where = [str(DP.cell(r, c).value).strip() for c in range(2, 8)
                     if isinstance(DP.cell(r, c).value, str) and DP.cell(r, c).value.strip()]
            if where:
                names.append({"name": nm.strip(),
                              "where": sorted(set(w.split(" - ")[-1] for w in where))})
    d["names"] = names

    # ---- 8. what to put the day into ----
    # Derived, not typed: the file already knows what is furthest behind, and a note
    # somebody has to remember to write is a note that goes stale.
    rates = {}
    for r in range(5, 20):
        prod = CAL.cell(r, 1).value
        if prod:
            rates[str(prod)] = num(CAL.cell(r, 10).value) or 0     # hours per 1,000 pcs
    sizes = {str(MT.cell(r, 1).value): num(MT.cell(r, 2).value) or 1 for r in range(13, 28)
             if MT.cell(r, 1).value}
    # Boxes and loose pieces are totalled apart. Thirteen of these lines are counted in
    # boxes and two - Classic and Kingdom loose - in pieces, because their box holds one
    # piece. Adding them together produced a headline in the tens of thousands that was
    # not a quantity of anything: it was boxes plus pieces. Same fault, and same fix, as
    # the one already corrected on the Sales Plan sheet.
    def _pcs(name):
        return (sizes.get(name) or 1) == 1
    tot["asked_box"] = sum((x.get("asked") or 0) for x in boxes if not _pcs(x["box"]))
    tot["asked_pcs"] = sum((x.get("asked") or 0) for x in boxes if _pcs(x["box"]))
    tot["owed_box"] = sum((x.get("owed") or 0) for x in boxes if not _pcs(x["box"]))
    tot["owed_pcs"] = sum((x.get("owed") or 0) for x in boxes if _pcs(x["box"]))
    for x in boxes:
        x["is_pcs"] = _pcs(x["box"])

    # ---- asked for beyond the month's target -----------------------------------
    # Two different problems, and the first version of this ran them together, which
    # made the table useless: ten thousand loose PIECES with no target at all sat on top
    # of it and buried the four box lines that genuinely overshoot.
    #
    #   over        - a target exists and the asking has passed it. Somebody has to
    #                 approve the extra, or it does not get made.
    #   untargeted  - asked for with no monthly target at all. These are PLANNED IN
    #                 FULL, exactly as asked: with no target there is nothing they can
    #                 be said to exceed, and holding them back would punish sales for a
    #                 gap in finance's sheet. They are listed only so the gap gets
    #                 closed - the action is to ask finance for a target, not to ask
    #                 anyone for an approval.
    #
    # Totals are kept apart by unit for the same reason they are everywhere else: a box
    # and a piece cannot be added together and mean anything.
    over, untargeted = [], []
    for x in boxes:
        a, tg = x.get("asked") or 0, x.get("target") or 0
        if not a:
            continue
        row = {"box": x["box"], "asked": a, "target": tg, "is_pcs": x["is_pcs"],
               "packed": x.get("packed") or 0}
        if tg and a > tg:
            row["planned"], row["over"] = tg, a - tg
            over.append(row)
        elif not tg:
            row["planned"], row["over"] = a, 0
            untargeted.append(row)
    over.sort(key=lambda z: -z["over"])
    untargeted.sort(key=lambda z: -z["asked"])
    d["over"] = over
    d["untargeted"] = untargeted
    d["over_total_box"] = sum(z["over"] for z in over if not z["is_pcs"])
    d["over_total_pcs"] = sum(z["over"] for z in over if z["is_pcs"])

    # ---- how long a thousand pieces actually takes -------------------------
    # This was the constant 15, written early and never checked. Over the 25 worked
    # days of the last payroll run the floor put 5,821 model person-hours into 240,237
    # pieces - 24.2 per thousand, and 23.6 measured against hours actually paid. Two
    # independent readings three percent apart; the 15 was understating every hours
    # figure on the page by nearly forty percent.
    #
    # So it is measured now, over the recent worked days, and moves as the product mix
    # and the floor's speed move. It falls back to the measured 24 rather than to the
    # old 15 when there is not enough history to measure from.
    _h = _p = 0.0
    for _r in (d.get("log") or [])[-20:]:
        if _r.get("made") and _r.get("staff"):
            _mh = (_r.get("hours") or 0) * (_r.get("ratio") or 0)
            if _mh:
                _h += _mh * _r["staff"]
                _p += _r["made"]
    HPK = round(_h / _p * 1000, 1) if (_p and _h) else 24.0
    d["hpk"] = HPK

    focus = []
    worst = sorted([b for b in boxes if (b.get("left") or 0) > 0],
                   key=lambda b: -((b["left"] * sizes.get(b["box"], 1)) / 1000)
                   * max(rates.values() or [1]))[:3]
    if worst:
        focus.append({"kind": "make", "items": [
            {"box": b["box"], "n": b["left"],
             "hours": round(b["left"] * sizes.get(b["box"], 1) / 1000 * HPK, 1)}
            for b in worst]})
    owing = sorted([b for b in boxes if (b.get("owed") or 0) > 0],
                   key=lambda b: -(b["owed"] or 0))[:3]
    if owing:
        focus.append({"kind": "pack",
                      "items": [{"box": b["box"], "n": b["owed"]} for b in owing]})
    d["focus"] = focus

    # ------------------------------------------------------------------ #
    #  The detail behind the four headline cards.                        #
    #                                                                    #
    #  The page used to show a judgement with no way to ask why. Each of #
    #  these is the working behind one of those judgements, read from    #
    #  the sheet that owns it - nothing here is recomputed.              #
    # ------------------------------------------------------------------ #

    def col(ws, row, *labels):
        """a column by its heading, so inserting one does not silently shift the read"""
        want = [l.strip().lower() for l in labels]
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row, c).value
            if isinstance(v, str) and v.strip().lower() in want:
                return c
        return None

    # ---- 1. today's plan, in both units ------------------------------
    SH = 12                                   # the Stock sheet's header row
    c_prod, c_flav = col(ST, SH, "Product"), col(ST, SH, "Flavour")
    c_stk = col(ST, SH, "STOCK NOW (pcs)")
    c_sug = col(ST, SH, "SUGGESTED (pcs)")
    c_pln = col(ST, SH, "PLAN: make today")
    c_act = col(ST, SH, "ACTUAL: made today")
    # the two piece pools the chef keeps: uncoated (STOCK NOW) and coated - separate,
    # so they add up - and what was coated today, for how far the day has got
    c_cst = col(ST, SH, "COATED stock now")
    c_coat = col(ST, SH, "ACTUAL: coated today")
    c_loose = col(ST, SH, "LOOSE: pcs to pack")
    # which box a flavour belongs to, so the box rows can open on their own flavours
    c_set = col(ST, SH, "Box Set")
    c_cnt = col(ST, SH, "Counts as")
    # How many pieces of this flavour go into each of the four big boxes. This is the
    # only honest link between a box and the flavours inside it: "Big 30" is not a Box
    # Set and not a "Counts as", so matching on those left it with no flavours at all
    # and it rendered as a bare line sitting under the heading, looking like a heading.
    c_per = {nm: col(ST, SH, nm) for nm in
             ("Big 30", "Big 24", "Big 12", "Big 6", "Other box")}
    flav, plan_by_prod = [], {}
    for r in range(SH + 1, SH + 80):
        nm = ST.cell(r, c_flav).value if c_flav else None
        if not isinstance(nm, str) or not nm.strip():
            continue
        prod = str(ST.cell(r, c_prod).value or "").strip()
        # Below the flavour table the sheet carries other panels whose own headers sit
        # in these columns ("Product / PLAN A: make (pcs)", "Delivery date / Product").
        # Read as flavours they became two phantom lines on the plan.
        if not prod or prod in ("Product", "Delivery date"):
            continue
        row = {"product": prod, "flavour": nm.strip(),
               "stock": num(ST.cell(r, c_stk).value) if c_stk else None,
               "sugg": num(ST.cell(r, c_sug).value) if c_sug else None,
               "plan": num(ST.cell(r, c_pln).value) if c_pln else None,
               "actual": num(ST.cell(r, c_act).value) if c_act else None,
               "cstock": num(ST.cell(r, c_cst).value) if c_cst else None,
               "coat": num(ST.cell(r, c_coat).value) if c_coat else None,
               "loose": num(ST.cell(r, c_loose).value) if c_loose else None,
               "set": str(ST.cell(r, c_set).value or "").strip() if c_set else "",
               "counts": str(ST.cell(r, c_cnt).value or "").strip() if c_cnt else "",
               "per": {nm: (num(ST.cell(r, cc).value) or 0)
                       for nm, cc in c_per.items() if cc}}
        flav.append(row)
        if row["plan"]:
            plan_by_prod[prod] = plan_by_prod.get(prod, 0) + row["plan"]

    c_box = col(ST, SH, "Box")
    c_bpl = col(ST, SH, "PLAN: boxes to pack")
    c_ask = col(ST, SH, "Aziz asked for")
    c_bac = col(ST, SH, "ACTUAL: boxes packed")
    c_max = col(ST, SH, "Max possible")
    c_lim = col(ST, SH, "Limiting flavour")
    # Stop at the end of the panel. The row after the last box holds the panel's own
    # explanatory caption, and reading on turned that sentence into a box called
    # "PLAN: boxes to pack is what the day is calculated from. [hover for the rest]".
    c_style = col(ST, SH, "Style")
    pboxes, plan_by_box = [], {}
    for r in range(SH + 1, SH + 40):
        nm = ST.cell(r, c_box).value if c_box else None
        if not isinstance(nm, str) or not nm.strip() or nm.strip().upper() == "TOTAL":
            break
        if c_style and not str(ST.cell(r, c_style).value or "").strip():
            break
        row = {"box": nm.strip(),
               "plan": num(ST.cell(r, c_bpl).value) if c_bpl else None,
               "asked": num(ST.cell(r, c_ask).value) if c_ask else None,
               "actual": num(ST.cell(r, c_bac).value) if c_bac else None,
               "max": num(ST.cell(r, c_max).value) if c_max else None,
               "limiting": str(ST.cell(r, c_lim).value or "").strip() if c_lim else ""}
        pboxes.append(row)
        plan_by_box[row["box"]] = row["plan"] or 0
    d["plan"] = {"date": ST["B3"].value, "flavours": flav, "boxes": pboxes}

    # How far the chef has got with the day the sheet is planning: four stages, each
    # either recorded or not. Made and packed fall back to the Actual Log, because the
    # chef records the day there even when the Stock sheet's actual columns stay empty.
    _pdk = daykey(ST["B3"].value) if ST["B3"].value else None
    _lg = next((x for x in log if daykey(x["date"]) == _pdk), None) if _pdk else None
    _lg = _lg or {}

    def _tot(rows, k):
        return sum((r.get(k) or 0) for r in rows) or None
    _mst, _pst = _tot(flav, "actual"), _tot(pboxes, "actual")
    d["stage_today"] = {
        "date": _pdk,
        "plan_pcs": _tot(flav, "plan"), "plan_boxes": _tot(pboxes, "plan"),
        "made": _mst or (_lg.get("made") or None),
        "made_from": "stock" if _mst else ("log" if _lg.get("made") else None),
        "coated": _tot(flav, "coat"),
        "packed": _pst or (_lg.get("packed") or None),
    }

    # ---- 2. the last shift, actual against the morning's plan --------
    # The plan is NOT kept per day: the Stock sheet holds one plan column and the roll
    # clears it each morning. So a planned figure can only be shown when the last logged
    # day IS the day the sheet is currently planning. When it is not, the page says so
    # rather than comparing today's plan with yesterday's output, which would read as a
    # wild miss every single morning.
    AH = 14                                   # the Actual Log's header row

    def band_at(label):
        for c in range(1, AL.max_column + 1):
            v = AL.cell(13, c).value
            if isinstance(v, str) and v.strip().lower().startswith(label.lower()):
                return c
        return None

    a_pcs, a_box = band_at("ACTUAL PCS"), band_at("ACTUAL BOXES")
    p_pcs, p_box = band_at("PLAN pcs"), band_at("PLAN boxes")
    prod_cols = [(c, str(AL.cell(AH, c).value).strip())
                 for c in range(a_pcs or 11, a_box or 21) if AL.cell(AH, c).value]
    unit_cols = [(c, str(AL.cell(AH, c).value).strip())
                 for c in range(a_box or 21, (p_pcs or 36)) if AL.cell(AH, c).value]
    # the captured plan, keyed by name so a column moving cannot shift the comparison
    plan_pcs_col, plan_box_col = {}, {}
    if p_pcs and p_box:
        for c in range(p_pcs, p_box):
            v = AL.cell(AH, c).value
            if isinstance(v, str) and v.strip():
                plan_pcs_col[v.strip()] = c
        for c in range(p_box, p_box + 40):
            v = AL.cell(AH, c).value
            if isinstance(v, str) and v.strip():
                plan_box_col[v.strip()] = c
    shift = None
    if log:
        want = log[-1]["date"]
        want = want.date() if isinstance(want, dt.datetime) else want
        lr = None
        for r in range(15, 135):
            dd = AL.cell(r, 1).value
            dd = dd.date() if isinstance(dd, dt.datetime) else dd
            if dd == want:
                lr = r
                break
        sd = ST["B3"].value
        sd = sd.date() if isinstance(sd, dt.datetime) else sd
        same = (sd == want)

        def planned(nm, cols, fallback):
            """the plan CAPTURED against this day, and only as a last resort today's.

            Reading the Stock sheet's live plan column was the only option before
            capture_plan.py existed, and it is right exactly once a day - when the last
            logged shift happens to be the day being planned. For every other day it
            compares last night's output with this morning's intention. So the captured
            figure wins wherever there is one, and the live column is used only for the
            same day, which is the one case where the two are the same thing.
            """
            c = cols.get(nm)
            if c and lr:
                v = num(AL.cell(lr, c).value)
                if v:
                    return v
            return fallback.get(nm) if same else None

        if lr:
            captured = any(num(AL.cell(lr, c).value)
                           for c in list(plan_pcs_col.values()) + list(plan_box_col.values()))
            shift = {
                "date": log[-1]["date"],
                "same_day_plan": bool(captured or same),
                "from_capture": bool(captured),
                "hours": num(AL.cell(lr, 6).value),
                "model": num(AL.cell(lr, 7).value),
                "ratio": num(AL.cell(lr, 8).value),
                "staff": num(AL.cell(lr, 2).value),
                "note": str(AL.cell(lr, 10).value or ""),
                "products": [{"name": nm, "actual": num(AL.cell(lr, c).value) or 0,
                              "plan": planned(nm, plan_pcs_col, plan_by_prod)}
                             for c, nm in prod_cols],
                "units": [{"name": nm, "actual": num(AL.cell(lr, c).value) or 0,
                           "plan": planned(nm, plan_box_col, plan_by_box)}
                          for c, nm in unit_cols]}
    d["shift"] = shift

    # ---- 2a-bis. overtime --------------------------------------------------
    # Counted in staff-hours: twenty-seven people staying three hours past the standard
    # day is eighty-one hours, which is the figure that costs money. The log keeps a
    # headcount and not names, so this is how much overtime the factory ran, never whose.
    d["overtime"] = {
        # Capacity Engine B7, not Actual Log B7 - the latter is "days within 15%"
        # and reads 30, which would have put a thirty-hour standard day on the page.
        "standard": num(wb["Capacity Engine"]["B7"].value),
        "month": num(AL.cell(6, 10).value),
        "days": num(AL.cell(7, 10).value),
        "last": num(AL.cell(lr, 61).value) if lr else None,
        "cost": None,
    }
    # Per day and per person as well as the total. "431 hours" is a number nobody has
    # any feel for; "4.2 hours each over ten days" is a number you can argue with.
    _ot = d["overtime"]
    _staff = num(MT["B8"].value) or None
    _ot["per_day"] = round(_ot["month"] / _ot["days"], 1) if _ot["month"] and _ot["days"] else None
    _ot["per_person"] = round(_ot["month"] / _staff, 1) if _ot["month"] and _staff else None
    _ot["staff"] = _staff
    if d["overtime"]["month"]:
        # 10.6 SAR per person-hour is the role-level average, the only wage figure kept
        # anywhere in this system; no individual's pay is stored or derivable from it.
        d["overtime"]["cost"] = round(d["overtime"]["month"] * 10.6)

    # ---- 2b. kept promises, day by day ------------------------------
    # The capture turns one comparison into a series: for every logged day that also has
    # a captured plan, what share of the planned pieces actually came out. This is the
    # question the whole exercise was for, and it only becomes askable once the plan
    # stops being thrown away each morning. Days before the capture started simply are
    # not in the list - they are not zeroes, they are unknown, and a chart that draws
    # them as zero would read as a catastrophe that never happened.
    keep = []
    if plan_pcs_col:
        for r in range(15, 135):
            dd = AL.cell(r, 1).value
            dd = dd.date() if isinstance(dd, dt.datetime) else dd
            if not isinstance(dd, dt.date):
                continue
            pl = sum(v for c in plan_pcs_col.values()
                     if (v := num(AL.cell(r, c).value)))
            if not pl:
                continue
            ac = sum(v for c, _nm in prod_cols if (v := num(AL.cell(r, c).value)))
            # A day with a plan and no output yet is today, not a failure. Including it
            # drew the current day at 0% every morning and dragged the whole line down.
            if not ac:
                continue
            keep.append({"date": dd, "plan": pl, "actual": ac, "hit": ac / pl})
    d["kept"] = keep[-14:]

    # ---- 3. the store, box by box ------------------------------------
    SRH = 6
    sc = {k: col(SR, SRH, k) for k in
          ("Box", "Stocktake", "Chef packed", "Ali received", "Gap", "Sales took",
           "Waste", "IN STORE NOW", "Sales asked for", "Still to make", "STILL OWED")}
    # STOP at the end of the table, do not skip past it. Skipping carried the scan on
    # into the next block of the sheet and pulled in its headings as if they were boxes -
    # "IS THE STORE FILLING UP?" appeared in the list of box types, and the duplicate
    # names underneath it quietly doubled several lines.
    store = []
    for r in range(SRH + 1, SRH + 25):
        nm = SR.cell(r, sc["Box"]).value if sc["Box"] else None
        if not isinstance(nm, str) or not nm.strip() or nm.strip().upper() == "TOTAL":
            break
        store.append({"box": nm.strip(), "pcs": sizes.get(nm.strip()),
                      **{k.lower().replace(" ", "_"):
                         (num(SR.cell(r, v).value) if v else None)
                         for k, v in sc.items() if k != "Box"}})
    # Ali's file is the authority on his own columns; the Stores sheet is a copy of
    # them that stopped being refreshed. Overlay what he wrote, and recompute the two
    # figures that are differences of the others so they cannot disagree with it.
    if _arows:
        for row in store:
            a = _arows.get(row["box"])
            if not a:
                continue
            row["stocktake"] = a["stocktake"]
            row["ali_received"] = a["in"]
            row["sales_took"] = a["out"]
            row["waste"] = a["waste"]
            row["in_store_now"] = a["stocktake"] + a["in"] - a["out"] - a["waste"]
            fl = ((_ali or {}).get("by_flavour") or {}).get(row["box"])
            if fl:
                row["by_flavour"] = sorted(
                    ({"name": k,
                      "now": v["stocktake"] + v["in"] - v["out"] - v["waste"],
                      **v} for k, v in fl.items()),
                    key=lambda x: -x["now"])
            if _since is not None:
                row["gap"] = _packed_since.get(row["box"], 0) - a["in"]
                row["packed_since"] = _packed_since.get(row["box"], 0)
            elif row.get("chef_packed") is not None:
                row["gap"] = row["chef_packed"] - a["in"]
            if row.get("sales_asked_for") is not None:
                row["still_owed"] = row["sales_asked_for"] - a["out"]
        if (_ali or {}).get("stocktake"):
            d["stocktake"] = _ali["stocktake"]
    d["gap_since"] = _since.isoformat() if _since else None
    d["store"] = store

    # ---- 4. the month, box by box, against the pace it needs ---------
    # "over target" needs a yardstick, and the only honest one is how much of the month
    # has gone: 190 boxes by the 31st is not the same promise as 190 by the 18th. Working
    # days, not calendar days, because the sheet plans in working days and so does the
    # floor.
    dt_total = num(MT["B6"].value) or 0
    dt_left = num(MT["B7"].value) or 0
    gone = max(0, dt_total - dt_left)
    mrows = []
    for r in range(13, 28):
        nm = MT.cell(r, 1).value
        if not isinstance(nm, str) or not nm.strip():
            continue
        tgt = num(MT.cell(r, 3).value) or 0
        pkd = num(MT.cell(r, 4).value) or 0
        exp = (tgt * gone / dt_total) if dt_total else None
        # vs_target is packed minus target - the whole month's promise, not the pace.
        # It is the figure finance keeps and the one the sheet's MAX(0,...) throws away,
        # so without it the two documents cannot be read against each other.
        mrows.append({"box": nm.strip(), "pcs": num(MT.cell(r, 2).value),
                      "target": tgt, "packed": pkd,
                      "left": num(MT.cell(r, 5).value),
                      "perday": num(MT.cell(r, 7).value),
                      "pct": (pkd / tgt) if tgt else None,
                      "expected": exp,
                      "delta": (pkd - exp) if exp is not None else None,
                      "vs_target": (pkd - tgt) if tgt else None,
                      "price": num(MT.cell(r, 9).value),
                      "value": num(MT.cell(r, 10).value),
                      "value_target": num(MT.cell(r, 11).value)})
    # Prices can be kept OUT of the shared workbook (the chef reads Monthly Targets) in
    # prices.json beside this script. The sheet's own price wins while it still has one.
    try:
        with open(os.path.join(HERE, "prices.json"), encoding="utf-8") as f:
            _pj = json.load(f).get("box", {})
    except Exception:
        _pj = {}
    for x in mrows:
        if x["price"] is None and _pj.get(x["box"]) is not None:
            x["price"] = _pj[x["box"]]
        if x["price"] is not None:
            if x["value"] is None:
                x["value"] = (x["packed"] or 0) * x["price"]
            if x["value_target"] is None:
                x["value_target"] = (x["target"] or 0) * x["price"]
    # ---- the same month in PIECES, which the workbook already works out ----------
    # "PIECES THE BOX TARGET NEEDS" has sat in Monthly Targets since the file was
    # built and has never been on a page. It is the question the chef actually works
    # to - boxes are what sales receive, pieces are what the floor makes - and it is
    # where Mini sitting at 2% of its month shows up.
    pr0 = find(MT, "PIECES THE BOX TARGET NEEDS", 1)
    prows = []
    if pr0:
        for r in range(pr0 + 2, pr0 + 20):
            nm = MT.cell(r, 1).value
            if not isinstance(nm, str) or not nm.strip() or nm.strip().upper().startswith("TOTAL"):
                break
            need = num(MT.cell(r, 2).value)
            made = num(MT.cell(r, 3).value)
            if need is None and made is None:
                continue
            prows.append({"product": nm.strip(), "need": need, "made": made,
                          "left": num(MT.cell(r, 4).value),
                          "perday": num(MT.cell(r, 5).value),
                          "pct": num(MT.cell(r, 6).value),
                          "safety": num(MT.cell(r, 7).value)})

    # ---- and per FLAVOUR, if Aziz has set one ------------------------------------
    # Nothing in the workbook carries a flavour target today. When he adds the block
    # this finds it by its heading and its column names rather than by address, so
    # moving it a row or renaming a column does not silently empty the page. Until it
    # exists, frows is empty and the page says nothing rather than showing zeroes.
    fr0 = find(MT, "MONTHLY TARGET - FLAVOUR", 1) or find(MT, "MONTHLY TARGET - FLAVOURS", 1)
    frows = []
    if fr0:
        hdr = fr0 + 1
        cols = {}
        for c in range(1, 15):
            v = MT.cell(hdr, c).value
            if isinstance(v, str) and v.strip():
                cols[v.strip().lower()] = c
        c_name = cols.get("flavour") or cols.get("النكهة") or 1
        c_tgt = next((c for k, c in cols.items() if k.startswith("target")), 2)
        for r in range(hdr + 1, hdr + 80):
            nm = MT.cell(r, c_name).value
            if not isinstance(nm, str) or not nm.strip() or nm.strip().upper().startswith("TOTAL"):
                break
            tgt = num(MT.cell(r, c_tgt).value)
            if tgt:
                frows.append({"flavour": nm.strip(), "target": tgt})

    d["mtarget"] = {"days_total": dt_total, "days_left": dt_left, "days_gone": gone,
                    "month": MT["B4"].value, "rows": mrows, "pieces": prows,
                    "flavours": frows,
                    "value": sum(r["value"] or 0 for r in mrows),
                    "value_target": sum(r["value_target"] or 0 for r in mrows)}

    # ---- 5. how many days the store covers ---------------------------
    # Stock on its own says nothing: 41 boxes is comfortable for one line and a crisis
    # for another. What matters is how long it lasts against what is already ordered, so
    # the Order Book is walked forward day by day until the shelf goes negative.
    #
    # Deliberately counts ONLY what is on the shelf - not what the chef plans to pack.
    # A warning that quietly assumes the plan will be met is a warning that disappears
    # exactly when the plan starts slipping, which is when it is needed.
    OB = wb["Order Book"]
    obc = {OB.cell(6, c).value: c for c in range(1, OB.max_column + 1) if OB.cell(6, c).value}
    today = ST["B3"].value
    today = today.date() if isinstance(today, dt.datetime) else today
    demand = {}
    if isinstance(today, dt.date):
        for r in range(7, 157):
            dd = OB.cell(r, 1).value
            dd = dd.date() if isinstance(dd, dt.datetime) else dd
            if not isinstance(dd, dt.date) or dd < today:
                continue
            off = (dd - today).days
            if off > 30:
                continue
            for nm, c in obc.items():
                v = num(OB.cell(r, c).value)
                if v:
                    demand.setdefault(nm, {})[off] = demand.get(nm, {}).get(off, 0) + v
    cover = []
    for b in store:
        nm = b["box"]
        have = b.get("in_store_now") or 0
        dm = demand.get(nm, {})
        total = sum(dm.values())
        out_on = None
        run = have
        for off in range(0, 31):
            run -= dm.get(off, 0)
            if run < 0:
                out_on = off
                break
        # Cover alone turned out to be a useless alarm: this factory packs to order, so
        # eight boxes out of ten sit at zero days every morning and a list that is always
        # red is a list nobody reads. The sharp question is narrower - is today's order
        # bigger than the shelf AND today's planned packing together? That is a shortfall
        # somebody has to do something about before the van arrives. Cover stays on the
        # store page, where it is reference rather than warning.
        today_order = dm.get(0, 0)
        planned_today = plan_by_box.get(nm, 0) or 0
        cover.append({"box": nm, "store": have, "ordered30": total,
                      "days": out_on, "first": min(dm) if dm else None,
                      "today": today_order, "plan": planned_today,
                      "short": max(0, today_order - have - planned_today)})
    cover.sort(key=lambda x: (x["days"] is None, x["days"] if x["days"] is not None else 99))
    d["cover"] = cover

    # ---- 5b. what is already promised, day by day - for the what-if page -------
    # The Order Book in boxes, from today forward, with each box's size, so the page can
    # weigh a new order against what sales have already booked.
    _sz = {str(MT.cell(r, 1).value).strip(): (num(MT.cell(r, 2).value) or 1)
           for r in range(13, 28) if MT.cell(r, 1).value}
    obdays = []
    if isinstance(today, dt.date):
        for r in range(7, 157):
            dd = OB.cell(r, 1).value
            dd = dd.date() if isinstance(dd, dt.datetime) else dd
            if not isinstance(dd, dt.date) or dd < today or (dd - today).days > 75:
                continue
            by = {nm: num(OB.cell(r, c).value) for nm, c in obc.items()
                  if nm in _sz and num(OB.cell(r, c).value)}
            if by:
                obdays.append({"date": dd.isoformat(), "by": by})
    d["ob"] = {"days": obdays, "pcs": _sz}
    # The workbook's own model rates: hours to MAKE one piece, by product, and hours to
    # PACK one box, by box. Only used to tell how much of an order's time is packing, so
    # the what-if page can price "pieces only, no packing". Same order as the rows they
    # feed: Monthly Targets' pieces block (products) and its box block (boxes).
    try:
        LS = wb["Lists"]
        _prods = [str(MT.cell(r, 1).value).strip() for r in range(32, 42) if MT.cell(r, 1).value]
        _units = [str(MT.cell(r, 1).value).strip() for r in range(13, 28) if MT.cell(r, 1).value]
        d["rates"] = {
            "make": {p: num(LS.cell(301, 2 + j).value) for j, p in enumerate(_prods)},
            "pack": {u: num(LS.cell(303, 2 + j).value) for j, u in enumerate(_units)}}
    except Exception:
        d["rates"] = {"make": {}, "pack": {}}
    d["dayoff"] = [str(MT["B5"].value or "").strip(), str(MT["D5"].value or "").strip()]

    # ---- 5c. what sales asked for, by filling --------------------------------------
    # The Order Book keeps Half, Fingers and Large Bites as one box total, so the filling
    # sales asked for is lost on the way in. Their requests are kept by filling in
    # sales_flavours.json beside this script, and set against the Order Book here: any
    # boxes booked with no filling behind them are said out loud as needing it.
    reqs = []
    try:
        with open(os.path.join(HERE, "sales_flavours.json"), encoding="utf-8") as f:
            reqs = json.load(f).get("requests", []) or []
    except Exception:
        reqs = []
    fill, horizon = {}, None
    if isinstance(today, dt.date):
        horizon = today + dt.timedelta(days=14)
        for bx in ("Half", "Fingers", "Large Bites"):
            ordered = sum((x["by"].get(bx) or 0) for x in obdays
                          if dt.date.fromisoformat(x["date"]) <= horizon)
            byf = {}
            for rq in reqs:
                try:
                    f0, f1 = dt.date.fromisoformat(rq["from"]), dt.date.fromisoformat(rq["to"])
                except Exception:
                    continue
                lo, hi = max(f0, today), min(f1, horizon)
                if hi < lo:
                    continue
                share = ((hi - lo).days + 1) / ((f1 - f0).days + 1)
                for ln in rq.get("lines", []):
                    if ln.get("box") == bx and ln.get("boxes"):
                        byf[ln["flavour"]] = byf.get(ln["flavour"], 0) + ln["boxes"] * share
            if ordered or byf:
                fill[bx] = {"ordered": ordered, "flav": byf,
                            "unknown": max(0, ordered - sum(byf.values()))}
    d["fill"] = {"boxes": fill, "until": horizon.isoformat() if horizon else None,
                 "requests": [{"from": r.get("from"), "to": r.get("to")} for r in reqs]}

    # ---- 6. what changed today ---------------------------------------
    # Two people touch the system and neither sees the other's work. This is the day in
    # one list: what the chef recorded, and every line Ali wrote, newest first.
    feed = []
    if log:
        lastd = log[-1]["date"]
        lastd = lastd.date() if isinstance(lastd, dt.datetime) else lastd
        if lastd and (today - lastd).days <= 1:
            feed.append({"who": "chef", "date": lastd, "what": "made",
                         "n": log[-1]["made"], "unit": "pcs"})
            feed.append({"who": "chef", "date": lastd, "what": "packed",
                         "n": log[-1]["packed"], "unit": "boxes"})
    wh = warehouse_path()
    if wh:
        try:
            aw = openpyxl.load_workbook(wh, data_only=True)["Warehouse"]
            for r in range(11, 311):
                dd = aw.cell(r, 1).value
                dd = dd.date() if isinstance(dd, dt.datetime) else dd
                kind, box = aw.cell(r, 2).value, aw.cell(r, 3).value
                qty = num(aw.cell(r, 4).value)
                if not (isinstance(dd, dt.date) and isinstance(kind, str) and qty):
                    continue
                if isinstance(today, dt.date) and (today - dd).days > 2:
                    continue
                feed.append({"who": "ali", "date": dd,
                             "what": ("in" if kind.upper().startswith("IN")
                                      else "out" if kind.upper().startswith("OUT") else "waste"),
                             "box": str(box or "").strip(), "n": qty, "unit": "boxes"})
        except Exception:
            pass
    feed.sort(key=lambda x: (x["date"], x["who"]), reverse=True)
    d["feed"] = feed[:14]

    # ---- 7. is the data itself trustworthy? --------------------------
    # Every figure above rests on two people remembering to write things down. When one
    # stops, nothing breaks loudly - the numbers simply go quietly stale, which is the
    # failure that has actually happened here more than once.
    def ago(x):
        x = x.date() if isinstance(x, dt.datetime) else x
        return (today - x).days if (isinstance(x, dt.date) and isinstance(today, dt.date)) else None

    last_chef = log[-1]["date"] if log else None
    last_ali = max((f["date"] for f in feed if f["who"] == "ali"), default=None)
    if last_ali is None and wh:
        try:
            aw = openpyxl.load_workbook(wh, data_only=True)["Warehouse"]
            ds = [aw.cell(r, 1).value for r in range(11, 311)]
            ds = [x.date() if isinstance(x, dt.datetime) else x for x in ds]
            last_ali = max((x for x in ds if isinstance(x, dt.date)), default=None)
        except Exception:
            pass
    missing = 0
    if isinstance(today, dt.date):
        have = {(x["date"].date() if isinstance(x["date"], dt.datetime) else x["date"])
                for x in log}
        for i in range(1, 15):
            if (today - dt.timedelta(days=i)) not in have:
                missing += 1
    # ---- 8. the task list ---------------------------------------------
    # A separate file, read straight. It is Aziz's own and nobody else writes to it, so
    # there is no sync to do here - the dashboard is a window on it, the way it is a
    # window on everything else.
    tasks = []
    tp = tasks_path()
    if tp:
        try:
            tw = openpyxl.load_workbook(tp, data_only=True)["Tasks"]
            for r in range(4, tw.max_row + 1):
                title = tw.cell(r, 2).value
                if not isinstance(title, str) or not title.strip():
                    continue
                due = tw.cell(r, 5).value
                due = due.date() if isinstance(due, dt.datetime) else due
                st = str(tw.cell(r, 7).value or "").strip()
                tasks.append({
                    "id": str(tw.cell(r, 1).value or ""),
                    "title": title.strip(),
                    "project": str(tw.cell(r, 3).value or "").strip(),
                    "owner": str(tw.cell(r, 4).value or "").strip(),
                    "due": due if isinstance(due, dt.date) else None,
                    "priority": str(tw.cell(r, 6).value or "").strip(),
                    "status": st,
                    "repeat": str(tw.cell(r, 8).value or "").strip(),
                    "notes": str(tw.cell(r, 9).value or "").strip(),
                    "overdue": bool(isinstance(due, dt.date) and st != "Done"
                                    and isinstance(today, dt.date) and due < today),
                })
        except Exception:
            pass
    d["tasks"] = tasks

    # ---- 9. raw materials, consumables, packaging -----------------------
    # Ali's own daily report, read straight. Only the lines that need an answer are
    # carried across: a negative balance, a line at or below its request level, or a
    # row he has marked. The file has 175 items and showing all of them would bury the
    # six that matter on any given day.
    # Read from a RECALCULATED copy. The stock and the Make Request flag are formulas,
    # and a workbook last written by a script carries formulas with no results, so a
    # plain read returns None for exactly the two columns this page is about.
    mats = []
    mp = materials_path()
    if mp:
        mp = recalc(mp) or mp
    if mp:
        try:
            mwb = openpyxl.load_workbook(mp, data_only=True)
            for tab in ("Raw Materials", "Consumables", "Packaging"):
                if tab not in mwb.sheetnames:
                    continue
                mws = mwb[tab]
                # the header row moves between tabs, so find it rather than assume
                # The columns are DECLARED, not sniffed. Raw Materials lost its header
                # row somewhere along the way, and on Consumables the Action values sit
                # one column to the right of the Action heading. Reading by heading
                # therefore finds nothing on one tab and the wrong cell on another.
                LAYOUT = {
                    "Raw Materials": dict(first=6, item=3, unit=4, left=11, need=12, act=13),
                    "Consumables":   dict(first=4, item=3, unit=4, left=10, need=11, act=12),
                    "Packaging":     dict(first=2, item=3, unit=4, left=20, need=None, act=None),
                }
                lay = LAYOUT.get(tab)
                if not lay:
                    continue
                hr = lay["first"] - 1
                c_item, c_unit = lay["item"], lay["unit"]
                c_left, c_need, c_act = lay["left"], lay["need"], lay["act"]
                for r in range(hr + 1, mws.max_row + 1):
                    nm = mws.cell(r, c_item).value
                    if not isinstance(nm, str) or not nm.strip():
                        continue
                    left = num(mws.cell(r, c_left).value)
                    need = num(mws.cell(r, c_need).value) if c_need else None
                    act = str(mws.cell(r, c_act).value or "").strip() if c_act else ""
                    # The Action cell is a formula on left vs request level, and its
                    # cached result goes stale when a script fills the inputs (LibreOffice
                    # does not redo it). So it is worked out here from the same two numbers.
                    if c_act and left is not None and need is not None:
                        act = "Make Request" if left <= need else "Available"
                    short = (left is not None and left < 0)
                    low = (left is not None and need is not None and need > 0
                           and left <= need)
                    if not (short or low or act.lower().startswith(("make", "need"))):
                        continue
                    # How far past the line it is, as a share of the request level -
                    # so a line 40 short of a 50 level outranks one 40 short of a 500
                    # level. Without this, big-volume items crowd out the ones that
                    # actually stop production tomorrow.
                    if short:
                        rank = 0
                        gap = abs(left)
                    else:
                        rank = 1
                        gap = (need - left) / need if need else 0
                    mats.append({"tab": tab, "item": nm.strip(),
                                 "unit": str(mws.cell(r, c_unit).value or "").strip()
                                         if c_unit else "",
                                 "left": left, "need": need, "action": act,
                                 "negative": bool(short),
                                 "rank": rank, "gap": round(gap, 3)})
        except Exception:
            pass
    # Worst first, always: anything negative, then whatever has eaten most of its
    # request level. The grouping by kind happens in the page, not here, so that the
    # order inside each group is the same order of danger.
    mats.sort(key=lambda x: (x["rank"], -x["gap"]))
    d["mat_groups"] = [t for t in ("Raw Materials", "Consumables", "Packaging")
                       if any(m["tab"] == t for m in mats)]
    d["materials"] = mats
    # Packaging is counted weekly, not daily, so the page shows it apart and under the
    # date of the count. The date is read from a column headed "date" on that tab (Ali's
    # weekly sheet carries one) - the dated issue columns along row 1 are not it.
    d["pack_count_date"] = None
    try:
        if mp and "Packaging" in mwb.sheetnames:
            pws, best = mwb["Packaging"], None
            hdr = {c: str(pws.cell(1, c).value or "").strip().lower()
                   for c in range(1, pws.max_column + 1)}
            for c, hname in hdr.items():
                if "date" not in hname:
                    continue
                for r in range(2, pws.max_row + 1):
                    v = pws.cell(r, c).value
                    if isinstance(v, dt.datetime): v = v.date()
                    if isinstance(v, dt.date) and (best is None or v > best):
                        best = v
            d["pack_count_date"] = best.isoformat() if best else None
    except Exception:
        pass

    # ---- 9b. the flavour detail, kept day by day ---------------------------
    # NOT named `flav` - that name is already taken further up by the suggestion rows,
    # and borrowing it shadowed them. The health count below reads f["sugg"], which these
    # rows do not have; it survived until a row turned up with no plan, because `and`
    # stops early. A crash months later would have been cheaper than a wrong number, but
    # only just.
    day, flav_rows = flavour_snapshot(wb)
    # every logged day, not the 14 the shift page draws: the history is what monthly
    # reports are built from, and 30 days of it was being thrown away each month
    hist = history_update(day, flav_rows, log=_log_all,
                          plans=plan_snapshot(wb))
    # ---- 9b-bis. give the last shift back its planned column ----------------
    #
    # The Last shift page draws "planned / actual / difference" for every product and
    # box. Since 19 October the planned side has been empty on every line, because the
    # only thing that ever wrote it - capture_plan.py, into the chef's workbook -
    # stopped when refreshes were made read-only. The page went on drawing the columns
    # against nothing, and a flag called same_day_plan went on claiming a plan was
    # there, which is worse than a blank: it says the comparison was made.
    #
    # The snapshot above has held that plan all along, in the repository, flavour by
    # flavour. So it is read back here for the day the shift belongs to.
    if d.get("shift"):
        sd = d["shift"]["date"]
        sd = sd.isoformat()[:10] if hasattr(sd, "isoformat") else str(sd)[:10]
        snap = hist.get(sd)
        pp = (snap or {}).get("plan_prod") or {} if isinstance(snap, dict) else {}
        pb = (snap or {}).get("plan_box") or {} if isinstance(snap, dict) else {}

        # Days snapshotted before the plan totals were stored still have the flavour
        # lines, and which product a flavour counts towards is a fact about the
        # catalogue rather than about that day - so the total can be rebuilt rather
        # than written off. The box plan cannot: it was never in the rows, and is
        # blank for those days rather than guessed.
        if not pp and isinstance(snap, dict):
            counts_of = {}
            for r in range(13, 60):
                nm = ST.cell(r, 2).value
                cn = ST.cell(r, 20).value
                if isinstance(nm, str) and isinstance(cn, str) and nm.strip() and cn.strip():
                    counts_of[nm.strip()] = cn.strip()
            for row in hist_rows(snap):
                cn = row.get("counts") or counts_of.get(row.get("name"))
                if cn and row.get("plan"):
                    pp[cn] = pp.get(cn, 0) + row["plan"]
        if pp or pb:
            for row in d["shift"]["products"]:
                if row.get("plan") is None:
                    row["plan"] = pp.get(row["name"])
            for row in d["shift"]["units"]:
                if row.get("plan") is None:
                    row["plan"] = pb.get(row["name"])
            d["shift"]["from_snapshot"] = True
        # The flag now means what it says: a plan exists for this day, from somewhere.
        d["shift"]["same_day_plan"] = any(
            (r.get("plan") or 0) for r in d["shift"]["products"] + d["shift"]["units"])

    # which product each flavour counts towards - a fact about the catalogue, read once
    counts_of = {}
    for r in range(13, 60):
        nm, cn = ST.cell(r, 2).value, ST.cell(r, 20).value
        if isinstance(nm, str) and isinstance(cn, str) and nm.strip() and cn.strip():
            counts_of[nm.strip()] = cn.strip()
    d["counts_of"] = counts_of

    d["flav_today"] = {"date": day,
                       "rows": infer_made(flav_rows, prod_by_day.get(day, {}), counts_of)}
    prev = [k for k in sorted(hist) if k != day]
    d["flav_last"] = ({"date": prev[-1],
                       "rows": infer_made(hist_rows(hist[prev[-1]]),
                                          prod_by_day.get(prev[-1], {}), counts_of)}
                      if prev else None)
    d["flav_days"] = sum(1 for v in hist.values() if hist_rows(v))

    # ---- what each flavour has made this month ------------------------------------
    # There is no cumulative per-flavour figure anywhere in the workbook - the chef's
    # product tabs are time standards, not production - so the only source is the
    # snapshot history, which started on 20 October. That makes this month's figure
    # PARTIAL, and the page has to say so: a target scored against half a month's
    # output would read as a failure that never happened. From 1 November it covers
    # the whole month.
    mon = None
    if isinstance(d.get("mtarget", {}).get("month"), (dt.date, dt.datetime)):
        m = d["mtarget"]["month"]
        mon = (m.date() if isinstance(m, dt.datetime) else m).strftime("%Y-%m")
    made_flav, first_day = {}, None
    for k in sorted(hist):
        if mon and not k.startswith(mon):
            continue
        rows = infer_made(hist_rows(hist[k]), prod_by_day.get(k, {}), counts_of)
        if rows:
            first_day = first_day or k
        for r in rows:
            if r.get("made"):
                made_flav[r["name"]] = made_flav.get(r["name"], 0) + r["made"]

    fr = d.get("mtarget", {}).get("flavours") or []
    for row in fr:
        row["made"] = made_flav.get(row["flavour"], 0)
        row["left"] = max(0, (row["target"] or 0) - row["made"])
        row["pct"] = (row["made"] / row["target"]) if row.get("target") else None
        row["counts"] = counts_of.get(row["flavour"], "")
    if fr:
        d["mtarget"]["flav_from"] = first_day
        d["mtarget"]["flav_partial"] = bool(
            first_day and mon and first_day > mon + "-01")
    d["made_flav"] = made_flav

    # ---- 9c. the last seven days, plan against actual ----------------------
    # Two sources, deliberately. The PLAN comes from the snapshot, because the Stock
    # sheet keeps one plan column and clears it each morning - after today, the
    # snapshot is the only thing that ever knew. The ACTUAL comes from the Actual Log,
    # which is final and which the chef corrects. Where they disagree the log wins,
    # and the flavour lines are labelled with the time they were taken rather than
    # quietly reconciled.
    kept_by_day = {daykey(k["date"]): k for k in d.get("kept", [])}
    days = []
    for k in sorted(hist)[-7:]:
        e = hist[k] if isinstance(hist[k], dict) else {"rows": hist[k]}
        rws = hist_rows(hist[k])
        plan = e.get("plan_pcs")
        if not plan:
            # Days written before the entry carried its own total: the plan is still
            # there, one flavour at a time.
            plan = sum(r.get("plan") or 0 for r in rws) or None
        if not plan and k in kept_by_day:
            plan = kept_by_day[k]["plan"]          # the days capture_plan still covered
        made = e.get("made")
        if made is None:
            made = sum(r.get("made") or 0 for r in rws) or None
        days.append({
            "date": k, "plan": plan or None, "made": made,
            "packed": e.get("packed"), "hours": e.get("hours"), "staff": e.get("staff"),
            "ratio": e.get("ratio"), "reading": e.get("reading"), "note": e.get("note"),
            "hit": (made / plan) if (plan and made) else None,
            "rows_at": e.get("rows_at"),
            # the two stages the totals above do not carry
            "coated": sum(r.get("coated") or 0 for r in rws) or None,
            "plan_boxes": sum((e.get("plan_box") or {}).values()) or None,
            "rows": sorted(infer_made(rws, prod_by_day.get(k, {}), counts_of),
                           key=lambda r: -((r.get("made") or 0) + (r.get("plan") or 0))),
            # The snapshot's own total, kept so the page can say when the flavour
            # breakdown does not add up to the day - rather than letting the reader
            # discover it by adding the column themselves.
            "rows_made": sum(r.get("made") or 0 for r in rws),
        })
    d["days"] = days

    # ---- 9d. delays: the Actual Log stopped carrying the captured plan when refreshes
    # went read-only, so recent days take their plan from the snapshot history instead.
    for x in d.get("delays", []):
        if x["plan"]:
            x["plan_src"] = "log"
            continue
        e = hist.get(x["date"])
        if not isinstance(e, dict):
            continue
        pp = e.get("plan_prod") or {}
        P = e.get("plan_pcs") or sum(pp.values())
        if not P:
            continue
        mk = {q["p"]: q["made"] for q in x["prods"]}
        x["prods"] = [{"p": n, "plan": pp.get(n, 0) or 0, "made": mk.get(n, 0) or 0}
                      for n in sorted(set(pp) | set(mk)) if pp.get(n) or mk.get(n)]
        x["plan"], x["plan_src"] = P, "snapshot"
        x["miss"] = round(abs((x["made"] or 0) - P) / P, 3)
        x["late"] = x["miss"] >= 0.10
    d["delays"] = [x for x in d.get("delays", []) if x["plan"] or x["reasons"]]

    # ---- 9e. every closed day the history holds, compact, for the reports page ----
    rep = []
    for k in sorted(hist):
        e = hist[k]
        if not isinstance(e, dict) or not e.get("staff"):
            continue
        pp = e.get("plan_prod") or {}
        rep.append({"d": k, "m": e.get("made"), "k": e.get("packed"), "s": e.get("staff"),
                    "h": e.get("hours"), "o": round(e.get("ot") or 0, 2), "r": e.get("ratio"),
                    "p": e.get("plan_al") or e.get("plan_pcs") or (sum(pp.values()) or None),
                    "bp": e.get("by_prod") or {}, "bb": e.get("by_box") or {},
                    "rs": [{"c": z.get("cat", ""), "r": z.get("reason", ""), "w": z.get("what", ""),
                            "p": z.get("prod", "")} for z in (e.get("reasons") or [])]})
    d["rep"] = rep
    d["months"] = month_targets(d)

    # The last day the CHEF has actually filled in. Unlike the file's timestamp this
    # cannot be moved by a script writing to the workbook, so it is the honest answer
    # to "when did he last touch this", and the two are shown side by side.
    # .date() first: a datetime's isoformat carries a T00:00:00 the page then fails to
    # parse, and it showed as "Invalid Date" on the home screen. datetime is a subclass
    # of date, so the isinstance check alone never caught it.
    d["chef_last_day"] = (last_chef.date() if isinstance(last_chef, dt.datetime)
                          else last_chef).isoformat() if isinstance(last_chef, dt.date) else None

    d["health"] = {
        "chef_days": ago(last_chef) if last_chef else None,
        "ali_days": ago(last_ali) if last_ali else None,
        "stocktake_days": ago(d.get("stocktake")) if d.get("stocktake") else None,
        "missing14": missing,
        "unplanned": sum(1 for f in flav if not f["plan"] and (f["sugg"] or 0) > 0),
    }

    # ---- what the order book will cost in overtime, before it is worked -----
    #
    # The overtime block above is a bill: 688 hours, already worked, already owed. This
    # is the same arithmetic done one window earlier, while it is still a decision -
    # which was the whole point of asking whether last month's 9,439 could have been
    # foreseen. It could: the work was there to be counted.
    #
    # Wages run from the 21st to the 20th, not the 1st to the 31st, so the window has
    # to match payroll or the answer is about a month nobody gets paid for. Friday is
    # the day off, read from the gaps in the log rather than assumed.
    #
    # Two demands, deliberately side by side and not reconciled. What SALES have asked
    # for and not yet been given, and what the month TARGET still calls for. When they
    # disagree - and they disagree by a factor of four right now - that gap is the
    # finding, and averaging it away would destroy the only number worth showing.
    today = dt.date.today()
    w0 = (dt.date(today.year, today.month, 21) if today.day >= 21
          else (dt.date(today.year, today.month, 1) - dt.timedelta(days=1)).replace(day=21))
    w1 = dt.date(w0.year + (w0.month == 12), (w0.month % 12) + 1, 20)
    span = [w0 + dt.timedelta(n) for n in range((w1 - w0).days + 1)]
    days_left = sum(1 for x in span if x > today and x.weekday() != 4)   # 4 = Friday
    days_gone = sum(1 for x in span if x <= today and x.weekday() != 4)

    crew = d["overtime"].get("staff") or 25
    cap_left = days_left * crew * 8

    # overtime already run inside THIS window, so the page shows the running total for
    # the run Aziz will actually be billed for, not for a calendar month
    std = d["overtime"].get("standard") or 8
    ot_sofar = sum(max(0.0, (r.get("hours") or 0) - std) * (r.get("staff") or 0)
                   for r in _log_all
                   if isinstance(r.get("date"), dt.date) and w0 <= (
                       r["date"].date() if isinstance(r["date"], dt.datetime) else r["date"]) <= w1)

    def _ot(pieces):
        hours = pieces / 1000.0 * HPK
        over = max(0.0, hours - cap_left)
        return {"pieces": round(pieces), "hours": round(hours),
                "ot": round(over), "cost": round(over * OT_RATE),
                "people": round(over / (8.0 * days_left), 1) if days_left else None}

    owed_pcs = sum(max(0, (r.get("still_owed") or 0)) * (r.get("pcs") or 1)
                   for r in d.get("store", []))
    target_pcs = sum(max(0, (r.get("left") or 0)) * (r.get("pcs") or 1)
                     for r in d.get("mtarget", {}).get("rows", []))

    # The pay run that just closed, as actually worked. Aziz reads his overtime off a
    # payroll line for 21 Sep - 20 Oct; a figure here for "this calendar month" is a
    # different period and invites him to compare two numbers that were never the same
    # question. So the closed run is computed on its own terms and shown beside the
    # current one.
    p0 = (dt.date(w0.year, w0.month, 1) - dt.timedelta(days=1)).replace(day=21)
    p1 = w0 - dt.timedelta(days=1)
    prev_ot = prev_days = 0
    for r in _log_all:
        rd = r.get("date")
        rd = rd.date() if isinstance(rd, dt.datetime) else rd
        if isinstance(rd, dt.date) and p0 <= rd <= p1:
            prev_days += 1
            prev_ot += max(0.0, (r.get("hours") or 0) - std) * (r.get("staff") or 0)

    d["fc"] = {
        "prev": {"from": p0.isoformat(), "to": p1.isoformat(),
                 "ot": round(prev_ot), "cost": round(prev_ot * OT_RATE),
                 "days": prev_days,
                 # "688 hours" is a number nobody has a feel for. Per person, and per
                 # person per day, are numbers you can argue with.
                 "each": round(prev_ot / crew, 1) if crew else None,
                 "each_day": round(prev_ot / crew / prev_days, 2)
                             if (crew and prev_days) else None,
                 # The log only reaches back fourteen days, so a run that closed
                 # earlier than that is partly outside it. Say so rather than quoting
                 # a total that is quietly missing its first week.
                 # True only if the sheet itself does not reach back far enough -
                 # not merely because the page draws fourteen days.
                 "partial": bool(_log_all and
                                 min((x["date"].date() if isinstance(x["date"], dt.datetime)
                                      else x["date"]) for x in _log_all) > p0)},
        "from": w0.isoformat(), "to": w1.isoformat(),
        "days_left": days_left, "days_gone": days_gone,
        "crew": crew, "cap_left": round(cap_left), "hpk": HPK,
        "ot_sofar": round(ot_sofar), "cost_sofar": round(ot_sofar * OT_RATE),
        "rate": OT_RATE, "normal": RATE,
        "orders": _ot(owed_pcs),
        "target": _ot(target_pcs),
        # The headcount question, answered against BOTH demands rather than only the
        # target. Asked against the target alone it reported 21% of capacity while the
        # block above it was asking for 484 overtime hours - the same page giving two
        # opposite answers because two blocks were quietly measuring different things.
        # frozen below, after the dict exists
        "opening": None,
        "sens": [{"n": n,
                  "cap": n * 8 * days_left,
                  "orders": round(_ot(owed_pcs)["hours"] / (n * 8 * days_left), 3)
                            if days_left else None,
                  "target": round(_ot(target_pcs)["hours"] / (n * 8 * days_left), 3)
                            if days_left else None}
                 for n in (18, 20, 22, crew)],
    }

    # ---- the forecast as it stood on day one of this pay run --------------
    #
    # The forecast is recomputed from scratch every build, so by the 20th there is
    # nothing left to compare the payroll line against: the number that was on the
    # page on the 21st is gone. Aziz wants three readings - before, during, and a
    # reconciliation at the close - and only the first has to be remembered.
    #
    # Written once per run and never touched again. A forecast you are allowed to
    # revise after the fact is not a forecast, it is a description.
    d["fc"]["opening"] = opening_forecast(w0.isoformat(), d["fc"])
    return d


def month_targets(d):
    """The month's box target and what was packed against it, one line per month.

    The Monthly Targets sheet is overwritten each month, so last month's target exists
    nowhere once November is entered. Kept beside the flavour history and refreshed on
    every build while the month is current; once the month is over it stops changing."""
    if not HISTORY_PATH:
        return {}
    path = os.path.join(os.path.dirname(HISTORY_PATH), "months.json")
    mm = {}
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                mm = json.load(f) or {}
        except Exception:
            mm = {}
    m = str((d.get("mtarget") or {}).get("month") or "")[:7]
    t = d.get("tot") or {}
    if m and t.get("target"):
        mm[m] = {"target_box": t.get("target"), "packed_box": t.get("packed"),
                 "value_target": (d.get("mtarget") or {}).get("value_target"),
                 "value": (d.get("mtarget") or {}).get("value"),
                 "days_total": (d.get("mtarget") or {}).get("days_total"),
                 # per box and per product, so a closed month can be read line by line
                 "rows": [{k: x.get(k) for k in ("box", "pcs", "target", "packed", "price",
                                                 "value", "value_target")}
                          for x in ((d.get("mtarget") or {}).get("rows") or [])],
                 "pieces": [{k: x.get(k) for k in ("product", "need", "made")}
                            for x in ((d.get("mtarget") or {}).get("pieces") or [])],
                 "at": now_local().strftime("%Y-%m-%dT%H:%M")}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(mm, f, ensure_ascii=False, indent=1, sort_keys=True)
    return mm


def opening_forecast(run_start, fc, keep=12):
    """Freeze this run's first forecast, and hand back whatever was frozen."""
    if not HISTORY_PATH:
        return None
    path = os.path.join(os.path.dirname(HISTORY_PATH), "runs.json")
    runs = {}
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                runs = json.load(f) or {}
        except Exception:
            runs = {}

    if run_start not in runs:
        o = fc.get("orders") or {}
        t = fc.get("target") or {}
        runs[run_start] = {
            "frozen_at": now_local().strftime("%Y-%m-%dT%H:%M"),
            "to": fc.get("to"),
            "days_left": fc.get("days_left"),
            "cap_left": fc.get("cap_left"),
            "hpk": fc.get("hpk"),
            # the overtime already run when the freeze was taken, so the comparison at
            # the close is like for like
            "ot_sofar": fc.get("ot_sofar"),
            "orders": {"ot": o.get("ot"), "cost": o.get("cost"), "hours": o.get("hours")},
            "target": {"ot": t.get("ot"), "cost": t.get("cost"), "hours": t.get("hours")},
        }
        for old in sorted(runs)[:-keep]:
            del runs[old]
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(runs, f, ensure_ascii=False, indent=1, sort_keys=True)
        except Exception:
            return None

    out = dict(runs[run_start])
    out["from"] = run_start
    # how far the picture has moved since the freeze - the whole point of keeping it
    o = fc.get("orders") or {}
    if out.get("orders", {}).get("cost") is not None and o.get("cost") is not None:
        was = (out["orders"]["cost"] or 0) + (out.get("ot_sofar") or 0) * fc.get("rate", 13.7)
        now = (fc.get("cost_sofar") or 0) + (o.get("cost") or 0)
        out["drift_cost"] = round(now - was)
    return out


def j(o):
    def enc(x):
        # the predicted finish is a TIME, not a date - Excel stores it as one and the
        # first encoder only knew about dates, so the whole build died on it
        if isinstance(x, dt.datetime):
            return x.strftime("%Y-%m-%d")
        if isinstance(x, dt.date):
            return x.strftime("%Y-%m-%d")
        if isinstance(x, dt.time):
            return x.strftime("%H:%M")
        raise TypeError(type(x))
    return json.dumps(o, default=enc, ensure_ascii=False)


BRAND_DIR = os.path.join(HERE, "brand")


def logo(name, width):
    """A brand PNG as a data: URL, scaled to the size the page actually draws it.

    Inlined rather than shipped beside the page because the dashboard is opened three
    ways - from GitHub Pages, from the folder on Aziz's Mac, and from the home-screen
    icon - and a relative <img src> that resolves in one of those can fail in another.
    A logo that sometimes appears is worse than no logo on a page meant to look finished.
    Scaled here too: the source wordmark is 4,500px wide and nothing on this page draws
    it past 300.
    """
    path = os.path.join(BRAND_DIR, name)
    if not os.path.exists(path):
        return ""
    try:
        from PIL import Image
        import base64
        import io
        im = Image.open(path).convert("RGBA")
        im = im.resize((width, max(1, round(im.height * width / im.width))), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, "PNG", optimize=True)
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return ""


def font_face(family, file, weight="400"):
    """One @font-face with the font inlined.

    The brand faces are subset to the characters this page can produce - Arabic, Latin,
    digits, the handful of symbols - which takes Zarid Display from 251 KB to 56 KB and
    each Dubai weight to about 30 KB. Inlined for the same reason as the logo: the page
    is opened from GitHub Pages, from a folder on a Mac, and from a home-screen icon, and
    a relative font URL that resolves in one of those fails in the others, leaving the
    page in a fallback face with no sign that anything is wrong.
    """
    path = os.path.join(BRAND_DIR, file)
    if not os.path.exists(path):
        return ""
    import base64
    b = base64.b64encode(open(path, "rb").read()).decode()
    return (f"@font-face{{font-family:'{family}';font-style:normal;font-weight:{weight};"
            f"font-display:swap;src:url(data:font/woff2;base64,{b}) format('woff2');}}")


# Dubai is free to use and redistribute. 29LT Zarid Display is a commercial licence held
# by the design team, and an embedded webfont on a page anyone can open is redistribution
# whether or not that was the intent. So it ships only when this is switched on, which
# should be once the repository is private or the foundry licence covers web use. The
# page falls back to a system serif, which is not the brand face but does not pretend to
# be: the wordmark itself is an image of the real logo and is unaffected either way.
EMBED_DISPLAY_FONT = "--with-zarid" in sys.argv


def fonts():
    return "\n".join(filter(None, [
        font_face("Zarid Display", "zarid-display.woff2", "400")
        if EMBED_DISPLAY_FONT else "",
        font_face("Dubai", "dubai-regular.woff2", "400"),
        font_face("Dubai", "dubai-medium.woff2", "500"),
        font_face("Dubai", "dubai-bold.woff2", "700"),
    ]))


HISTORY_PATH = os.environ.get("ERMINE_HISTORY", "")


def flavour_snapshot(wb):
    """Today's flavour lines, exactly as the Stock sheet has them right now.

    The Stock sheet is a single day. It is overwritten every morning, so the detail of
    what was actually made - which fillings, how many of each - exists only while that
    day is the current one and is gone by the next. The product families do get kept in
    the Actual Log, but 'Classic 2,400' does not answer the question that matters, which
    is whether that was passion fruit or dark chocolate.

    So each build takes a copy. Nothing is reconstructed: yesterday, before the first
    copy was taken, is genuinely unrecoverable and is shown as absent rather than
    guessed at."""
    st = wb["Stock"]
    day = st["B3"].value
    if isinstance(day, dt.datetime):
        day = day.date()
    if not isinstance(day, dt.date):
        return None, []
    rows = []
    for r in range(13, 60):
        name = st.cell(r, 2).value                      # B - Flavour
        if not isinstance(name, str) or not name.strip():
            continue
        plan = num(st.cell(r, 12).value)                # L - PLAN: make today
        made = num(st.cell(r, 13).value)                # M - ACTUAL: made today
        coat = num(st.cell(r, 15).value)                # O - ACTUAL: coated today
        if not (plan or made or coat):
            continue
        rows.append({"name": name.strip(),
                     "group": str(st.cell(r, 4).value or "").strip(),
                     # The product family the flavour counts towards. Kept so the
                     # snapshot can answer "what was Coins planned to do" without the
                     # Stock sheet, which is overwritten every morning.
                     "counts": str(st.cell(r, 20).value or "").strip(),
                     "plan": plan or 0, "made": made or 0, "coated": coat or 0})
    return day.isoformat(), rows


def daykey(v):
    """A date as the ten characters yyyy-mm-dd, whatever shape it arrived in.

    The Actual Log hands back datetimes, so isoformat() produced
    "2026-10-21T00:00:00" while the snapshot had written "2026-10-21". Both are the
    same Tuesday; the history file ended up with two of it, one holding the flavours
    and the other the totals, and the page listed the day twice with half the
    numbers in each. Same mistake as the quality sheet's date column, two weeks apart:
    read the value, do not assume its type.
    """
    if hasattr(v, "isoformat"):
        return v.isoformat()[:10]
    return str(v or "")[:10]


def infer_made(rows, made_by_product, counts_of):
    """Fill in a flavour's made figure where only one answer is possible.

    The chef records what was MADE in the Actual Log, by product. The flavour detail
    lives in a different column of a different sheet, and on 26 and 27 October he did
    not fill it - so the page showed six flavours at zero against a day that made 6,668
    pieces, and a line called "the rest" carrying the whole day.

    Where a product had exactly ONE flavour planned, the product's total IS that
    flavour's total; there is nowhere else it could have gone. That case is filled in
    and marked, so the figure is never mistaken for something the chef typed. Where two
    or more flavours of a product were planned the answer is genuinely unknown, and it
    stays blank - a split invented here would be indistinguishable from a measurement.
    """
    out = [dict(r) for r in rows]
    planned = {}
    for r in out:
        prod = r.get("counts") or counts_of.get(r.get("name"), "")
        if prod and r.get("plan"):
            planned.setdefault(prod, []).append(r)
    for prod, group in planned.items():
        total = made_by_product.get(prod)
        if not total:
            continue
        if any(g.get("made") for g in group):     # he did fill it in; leave it alone
            continue
        if len(group) == 1:
            group[0]["made"] = total
            group[0]["made_inferred"] = True
        else:
            for g in group:
                g["made_ambiguous"] = prod
    return out


def hist_rows(entry):
    """The flavour lines out of a history entry, old shape or new.

    The file started life as {date: [rows]}. It is now {date: {rows, totals...}} so a
    day can carry its own totals. Both shapes are read, because the days already on
    disk are the only record of themselves and rewriting them to fit a newer idea
    would be inventing history."""
    if isinstance(entry, dict):
        return entry.get("rows") or []
    return entry or []


def plan_snapshot(wb):
    """The morning's plan for today, per product family and per box.

    The flavour lines are copied already; these are the same plan summed the two ways
    the rest of the system asks for it. Kept because the Stock sheet holds ONE plan
    column and clears it every morning - so by the time anyone looks at yesterday's
    shift, the plan it was judged against no longer exists anywhere.

    That is what emptied the planned column on the Last shift page: capture_plan.py
    wrote it into the chef's workbook, and it stopped the day refreshes were made
    read-only. This writes to the repository instead, so nothing is written to anyone's
    file and nothing can stop it again for the same reason.
    """
    st = wb["Stock"]
    by_prod, by_box = {}, {}
    for r in range(13, 60):
        nm = st.cell(r, 2).value
        if not isinstance(nm, str) or not nm.strip():
            continue
        plan = num(st.cell(r, 12).value)
        counts = str(st.cell(r, 20).value or "").strip()
        if plan and counts:
            by_prod[counts] = by_prod.get(counts, 0) + plan
    # col 24 = Box, col 28 = PLAN: boxes to pack
    for r in range(13, 40):
        nm = st.cell(r, 24).value
        if not isinstance(nm, str) or not nm.strip() or nm.strip().upper() == "TOTAL":
            break
        plan = num(st.cell(r, 28).value)
        if plan:
            by_box[nm.strip()] = by_box.get(nm.strip(), 0) + plan
    return by_prod, by_box


def history_update(day, rows, keep=800, log=None, plans=None):
    """Keep the snapshots in one small JSON, newest day last, and hand back the lot.

    Today's entry is REPLACED on every build, because a day is rewritten many times as
    the chef fills it in and only the final state is true.

    Closed days used to be frozen at whatever the last build of that day happened to
    catch. On 21 October the last build ran before the chef finished, so the day was
    recorded as 5,680 pieces against a true 10,521 - the file preserved a moment and
    called it a day. So the day TOTALS are now re-read from the Actual Log on every
    build, for every day the log still reaches, and a closed day corrects itself the
    next time anything runs. The flavour lines stay as captured, because the Stock
    sheet no longer holds them and there is nothing to correct them from; they carry
    `rows_at` so the page can say when the detail was taken."""
    hist = {}
    if HISTORY_PATH and os.path.exists(HISTORY_PATH):
        try:
            with open(HISTORY_PATH, encoding="utf-8") as f:
                hist = json.load(f) or {}
        except Exception:
            hist = {}

    if day and rows:
        prev = hist.get(day)
        entry = dict(prev) if isinstance(prev, dict) else {}
        entry["rows"] = rows
        if plans:
            entry["plan_prod"], entry["plan_box"] = plans
        entry["rows_at"] = now_local().strftime("%Y-%m-%dT%H:%M")
        # The plan is the one figure only the snapshot has: the Stock sheet's plan
        # column is cleared every morning, so after today nothing else remembers it.
        entry["plan_pcs"] = sum(r.get("plan") or 0 for r in rows)
        hist[day] = entry

    for lg in (log or []):
        k = daykey(lg["date"])
        prev = hist.get(k)
        entry = dict(prev) if isinstance(prev, dict) else {"rows": prev or []}
        entry.update({"made": lg.get("made"), "packed": lg.get("packed"),
                      "hours": lg.get("hours"), "staff": lg.get("staff"),
                      "ratio": lg.get("ratio"), "reading": lg.get("reading"),
                      "note": lg.get("note"), "by_prod": lg.get("by_prod"),
                      "by_box": lg.get("by_box"), "ot": lg.get("ot")})
        if lg.get("reasons") is not None:
            entry["reasons"] = lg["reasons"]
        if lg.get("plan_al"):
            entry["plan_al"] = lg["plan_al"]
        hist[k] = entry

    # Corrections the chef sends after a day has closed. The flavour lines of a closed
    # day cannot be re-read from the Stock sheet, so a late figure ("I forgot to fill in
    # Mini Coffee") is written here: {"YYYY-MM-DD": {"Flavour name": pieces made}}.
    try:
        with open(os.path.join(HERE, "history_fixes.json"), encoding="utf-8") as f:
            fixes = json.load(f) or {}
    except Exception:
        fixes = {}
    for k, fl in fixes.items():
        e = hist.get(k)
        if not isinstance(e, dict):
            continue
        for r in e.get("rows") or []:
            if r.get("name") in fl:
                r["made"] = fl[r["name"]]
    for old in sorted(hist)[:-keep]:
        del hist[old]
    if HISTORY_PATH:
        os.makedirs(os.path.dirname(HISTORY_PATH), exist_ok=True)
        with open(HISTORY_PATH, "w", encoding="utf-8") as f:
            json.dump(hist, f, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    return hist


def qc_api():
    """The QC web app's address, so the dashboard can ask it for reports.

    Read from a file rather than hard-coded because the address changes whenever the
    Apps Script is redeployed. It is NOT a secret - it already sits in qc/config.js in
    a public repository - and it grants nothing on its own: every action behind it
    demands a code the page asks the viewer for and that never touches this build.
    """
    import ermine_env
    _plan = ermine_env.plan_dir()
    for d in (os.path.join(HERE, "qc_api.txt"),
              os.path.join(_plan, "w Claude", "qc_api.txt") if _plan else "",
              *glob.glob("/sessions/*/mnt/*/w Claude/qc_api.txt")):
        if not d:
            continue
        if os.path.exists(d):
            return open(d, encoding="utf-8").read().strip()
    return ""


def build(data, built_at):
    return (template()
            .replace("__DATA__", j(data))
            .replace("__BUILT__", built_at.isoformat())
            .replace("/*__FONTS__*/", fonts())
            .replace("__WORDMARK__", logo("wordmark-white.png", 300))
            .replace("__SYMBOL__", logo("symbol-white.png", 180))
            .replace("__WORDMARK_GREEN__", logo("wordmark-green.png", 260))
            .replace("__QC_API__", qc_api()))


TEMPLATE_PATH = os.path.join(HERE, "dashboard_template.html")


def template():
    """The page lives in its own .html file, not in a python string. It grew past six
    hundred lines of markup, CSS and script, and every edit inside a triple-quoted string
    meant escaping braces and losing syntax highlighting - which is how a stray brace once
    silently swallowed a whole block."""
    with open(TEMPLATE_PATH, encoding="utf-8") as f:
        return f.read()




def main(src=None, out=None):
    src = src or live_path()
    if not src or not os.path.exists(src):
        raise SystemExit("could not find the live workbook")
    print("recalculating", os.path.basename(src))
    got = recalc(src)
    if not got:
        raise SystemExit("LibreOffice could not recalculate the workbook")
    data = read(got)

    # When the file itself was last written, which is a different fact from when this
    # page was last built. The page can be rebuilt ten times over a workbook nobody has
    # touched since Tuesday - and "updated 2 minutes ago" would then be true of the page
    # and badly misleading about the numbers on it. Taken from the ORIGINAL on OneDrive,
    # not the recalculated copy, whose timestamp is only ever a few seconds old.
    def saved(path):
        try:
            return dt.datetime.fromtimestamp(os.path.getmtime(path), TZ).isoformat(timespec="seconds")
        except Exception:
            return None
    wh = ermine_paths.find(ermine_paths.WAREHOUSE)
    mp = materials_path()
    data["saved"] = {"plan": saved(src), "store": saved(wh) if wh else None,
                     "mat": saved(mp) if mp else None}
    out = out or os.path.join(HERE, "ermine-dashboard.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(build(data, now_local()))
    print(f"saved {out}  ({os.path.getsize(out):,} bytes)")
    print(f"  {len(data['boxes'])} boxes, {len(data['names'])} on the floor, "
          f"{len(data['pieces'])} products, {len(data['log'])} recorded days")
    return out


if __name__ == "__main__":
    main(*sys.argv[1:])
