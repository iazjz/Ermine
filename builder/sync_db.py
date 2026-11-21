"""Bring the Wardiya database in step with Production Plan - LIVE (read-only on Excel).

Runs after every dashboard build (publish_dashboard.py, Mac and cloud). Sends one JSON to
the database function sync_from_excel (app/supabase/003_sync.sql) with the SERVER key:
  env ERMINE_SB_SECRET, or the file  w Claude/supabase.secret  (never printed, never in the repo).
No key -> skipped, the publish carries on. Any error -> one line, the publish carries on.
Same reading rules as app/tools/make_master_sql.py.  usage: python3 sync_db.py [--dry-run]
"""
import datetime as dt, json, os, sys, tempfile, urllib.request
import openpyxl, ermine_paths

URL = "https://txqrrnymaglkilowgvee.supabase.co"
HERE = os.path.dirname(os.path.abspath(__file__))
AR = {"Kitchen": "المطبخ", "Chocolate Room": "غرفة الشوكولاتة", "Hallway": "الممر",
      "Dishwashing Area": "منطقة غسيل الأطباق", "Entrance": "المدخل", "Chiller Room": "غرفة التبريد",
      "Freezer Room": "غرفة التجميد", "Store": "المخزن"}
OF = {"Big 30": "Classic", "Big 12": "Classic", "Big 6": "Classic", "Large Bites": "Classic",
      "Classic pcs loose": "Classic", "Big 24": "Kingdom", "Kingdom pcs loose": "Kingdom",
      "Mini Raspberry loose": "Mini", "Crown 30": "Crown", "Crown 21": "Crown", "Melted": "Winter"}
BY_FLAVOUR = ("Crown 30", "Fingers", "Large Bites")


def secret():
    k = os.environ.get("ERMINE_SB_SECRET", "").strip()
    if k: return k
    f = os.path.join(HERE, "..", "supabase.secret")
    return open(f, encoding="utf-8").read().strip() if os.path.exists(f) else ""


def workbook():
    live = ermine_paths.find(ermine_paths.LIVE)
    if not live: return None
    rc = os.path.join(tempfile.gettempdir(), "dash_recalc", os.path.basename(live))
    # the builder's recalculated copy, when the live file had no cached results
    use = rc if os.path.exists(rc) and os.path.getmtime(rc) >= os.path.getmtime(live) else live
    return openpyxl.load_workbook(use, data_only=True)


def payload(wb):
    mt, st, ob, ar, ls = wb["Monthly Targets"], wb["Stock"], wb["Order Book"], wb["Areas"], wb["Lists"]
    prices = json.load(open(os.path.join(HERE, "prices.json"), encoding="utf-8"))["box"]
    stations = []   # the 10 numbered areas, named as the Daily Plan names them ("1 - Kitchen")
    for r in range(5, 40):
        code, a = ar.cell(r, 1).value, ar.cell(r, 2).value
        if isinstance(code, (int, float)) and a:
            stations.append({"en": f"{int(code)} - {a}", "ar": f"{int(code)} - {AR.get(a, a)}"})
    top = next(r for r in range(1, 120) if mt.cell(r, 1).value == "Product" and mt.cell(r, 2).value == "Pcs needed")
    prods = []
    for r in range(top + 1, top + 30):
        v = mt.cell(r, 1).value
        if not v or v == "TOTAL pcs": break
        prods.append((v, mt.cell(r, 2).value))
    pnames = [p for p, _ in prods]
    H = {st.cell(12, c).value: c for c in range(1, 60) if st.cell(12, c).value}
    cc = next(c for h, c in H.items() if str(h).lower().startswith("counts"))
    flv = []
    for r in range(13, 140):
        nm, pr = st.cell(r, 2).value, st.cell(r, cc).value
        if not nm:
            if r > 20: break
            continue
        if str(st.cell(r, 1).value or "") in ("Product", "Delivery date") or not pr: continue
        pr = str(pr).strip()
        if pr in pnames and {"name": str(nm).strip(), "product": pr} not in flv:
            flv.append({"name": str(nm).strip(), "product": pr})
    hdr = next(r for r in range(1, 40) if mt.cell(r, 1).value == "Box")
    boxes = []
    for r in range(hdr + 1, hdr + 40):
        b = mt.cell(r, 1).value
        if not b or b == "TOTAL boxes": break
        boxes.append({"code": b, "size": int(mt.cell(r, 2).value or 1), "by_flavour": b in BY_FLAVOUR,
                      "sort": len(boxes), "product": OF.get(b, b), "price": prices.get(b),
                      "_t": mt.cell(r, 3).value})
    month = mt["B4"].value
    month = (month.date() if isinstance(month, dt.datetime) else month).replace(day=1)
    reasons = [{"text": ls.cell(r, 62).value, "cat": ls.cell(r, 63).value or "Other",
                "note": "outside" in str(ls.cell(r, 62).value).lower()}
               for r in range(4401, 4418) if ls.cell(r, 62).value]
    t_flav = []
    ft = next((r for r in range(1, mt.max_row + 1) if mt.cell(r, 1).value == "MONTHLY TARGET - FLAVOUR"), None)
    if ft:
        for r in range(ft + 1, ft + 60):
            a, b = mt.cell(r, 1).value, mt.cell(r, 2).value
            if a == "TOTAL pcs": break
            if isinstance(b, (int, float)) and b > 0: t_flav.append({"flavour": a, "qty": int(round(b))})
    OH = {ob.cell(6, c).value: c for c in range(1, 40) if ob.cell(6, c).value}
    orders = []
    for r in range(7, 400):
        d = ob.cell(r, 1).value
        if not isinstance(d, dt.datetime) or d.date() < month: continue
        for b in boxes:
            c = OH.get(b["code"]); v = ob.cell(r, c).value if c else None
            if isinstance(v, (int, float)) and v > 0:
                orders.append({"day": d.date().isoformat(), "box": b["code"], "qty": int(round(v))})
    return {"factory": "ermine", "month": month.isoformat(),
            "working_days": mt["B6"].value, "staff": mt["B8"].value,
            "stations": stations,
            "products": [{"code": p, "coated": p in ("Mini", "Ice Cream"), "sort": i} for i, p in enumerate(pnames)],
            "flavours": flv, "boxes": [{k: v for k, v in b.items() if k != "_t"} for b in boxes],
            "reasons": reasons,
            "t_box": [{"box": b["code"], "qty": int(round(b["_t"]))} for b in boxes if b["_t"] is not None],
            "t_prod": [{"product": p, "qty": int(round(n))} for p, n in prods if n is not None],
            "t_flav": t_flav, "orders_from": month.isoformat(), "orders": orders}


def ref_payload(wb, base):
    """Figures the plan screen shows beside each line, read from the workbook as it stands."""
    mt, st, ls, ss = wb["Monthly Targets"], wb["Stock"], wb["Lists"], wb["Staff & Skills"]
    steps = []
    for r in range(1, 60):
        v = ls.cell(r, 1).value
        if not v or str(v).startswith("labour hours"): break
        steps.append({"name": str(v).strip(), "sort": len(steps)})
    staff = []
    for r in range(5, 80):
        v = ss.cell(r, 1).value
        if not v: break
        staff.append({"name": str(v).strip(), "sort": len(staff)})
    prods = [x["code"] for x in base["products"]]
    boxes = [x["code"] for x in base["boxes"]]
    num = lambda v: v if isinstance(v, (int, float)) and not isinstance(v, bool) else None
    make = [{"code": p, "h": num(ls.cell(301, 2 + j).value)} for j, p in enumerate(prods) if num(ls.cell(301, 2 + j).value) is not None]
    pack = [{"code": b, "h": num(ls.cell(303, 2 + j).value)} for j, b in enumerate(boxes) if num(ls.cell(303, 2 + j).value) is not None]
    H = {st.cell(12, c).value: c for c in range(1, 60) if st.cell(12, c).value}
    targets = {x["flavour"]: x["qty"] for x in base["t_flav"]}
    made = {}
    hist = os.environ.get("ERMINE_HISTORY")
    if hist and os.path.exists(hist):
        mon = base["month"][:7]
        for day, v in json.load(open(hist, encoding="utf-8")).items():
            if day[:7] == mon and isinstance(v, dict):
                for r in v.get("rows") or []:
                    made[r.get("name")] = made.get(r.get("name"), 0) + (r.get("made") or 0)
    fl = []
    for r in range(13, 140):
        nm = st.cell(r, 2).value
        if not nm:
            if r > 20: break
            continue
        nm = str(nm).strip()
        g = lambda h: num(st.cell(r, H[h]).value) if h in H else None
        fl.append({"name": nm, "stock": round(g("STOCK NOW (pcs)") or 0), "coated": round(g("COATED stock now") or 0),
                   "sugg": round(g("SUGGESTED (pcs)") or 0), "target": targets.get(nm),
                   "made": made.get(nm) if hist else None})
    bx = []
    cb, cm = H.get("Box"), H.get("Max possible")
    left = {}
    hdr = next(r for r in range(1, 40) if mt.cell(r, 1).value == "Box")
    for r in range(hdr + 1, hdr + 40):
        b = mt.cell(r, 1).value
        if not b or b == "TOTAL boxes": break
        left[b] = num(mt.cell(r, 5).value)
    for r in range(13, 60):
        b = st.cell(r, cb).value if cb else None
        if not isinstance(b, str) or not b.strip() or b.strip().upper() == "TOTAL": break
        bx.append({"code": b.strip(), "max": round(num(st.cell(r, cm).value) or 0) if cm else None, "left": left.get(b.strip())})
    return {"factory": "ermine", "stations": base["stations"], "steps": steps, "staff": staff,
            "make": make, "pack": pack, "fl": fl, "bx": bx}


SET_BOX = {"mini": "Mini", "miniloose": "Mini Raspberry loose", "coins": "Coins", "origin": "Origin",
           "icecream": "Ice Cream", "melted": "Melted", "crown3": "Crown 21", "mcake": "Marshmallow Cake"}
SINGLE_BOX = {"Crown": "Crown 30", "Fingers": "Fingers", "Cake": "Cake"}


def box_flavour_rows(wb):
    """Pieces of each flavour one box takes: Stock sheet columns Big 30 / 24 / 12 / 6 and
    'Other box' (the box of the flavour's own set)."""
    st = wb["Stock"]
    H = {st.cell(12, c).value: c for c in range(1, 60) if st.cell(12, c).value}
    cc = next(c for h, c in H.items() if str(h).lower().startswith("counts"))
    rows = []
    for r in range(13, 140):
        nm = st.cell(r, 2).value
        if not nm:
            if r > 20: break
            continue
        nm, bset, prod = str(nm).strip(), str(st.cell(r, 4).value or "").strip(), str(st.cell(r, cc).value or "").strip()
        for b in ("Big 30", "Big 24", "Big 12", "Big 6"):
            v = st.cell(r, H[b]).value if b in H else None
            if isinstance(v, (int, float)) and v > 0: rows.append({"box": b, "flavour": nm, "pcs": v})
        v = st.cell(r, H["Other box"]).value if "Other box" in H else None
        box = SET_BOX.get(bset) or (SINGLE_BOX.get(prod) if bset == "single" else None)
        if box and isinstance(v, (int, float)) and v > 0: rows.append({"box": box, "flavour": nm, "pcs": v})
    return rows


def ali_book(name):
    """One of Ali's files, the builder's recalculated copy when it is newer (his files are
    script-written and carry no cached results)."""
    path = ermine_paths.find(name)
    if not path: return None
    rc = os.path.join(tempfile.gettempdir(), "dash_recalc", os.path.basename(path))
    use = rc if os.path.exists(rc) and os.path.getmtime(rc) >= os.path.getmtime(path) else path
    return openpyxl.load_workbook(use, data_only=True)


def store_payload(base):
    """Ali's store balance (Warehouse) and his materials (Raw Materials), mapped to Wardiya's
    boxes and flavours. Crown and Fingers are kept by filling ("Crown Dates", "Fingers Coconut")."""
    boxes = {b["code"] for b in base["boxes"]}
    flv = {f["name"]: f["product"] for f in base["flavours"]}
    store = []
    wh = ali_book("Warehouse - Ali Nawaz.xlsx")
    if wh:
        ws = wh.worksheets[0]
        for r in range(11, 200):
            nm = ws.cell(r, 7).value
            if not isinstance(nm, str) or not nm.strip(): continue
            nm = nm.strip()
            if nm.upper() == "TOTAL": break
            q = ws.cell(r, 12).value
            q = int(round(q)) if isinstance(q, (int, float)) else None
            if nm in boxes:
                store.append({"name": nm, "box": nm, "flavour": None, "qty": q})
            elif nm.startswith("Crown ") and flv.get(nm) == "Crown":
                store.append({"name": nm, "box": "Crown 30", "flavour": nm, "qty": q})
            elif nm.startswith("Fingers ") and flv.get(nm[8:] + " Fingers") == "Fingers":
                store.append({"name": nm, "box": "Fingers", "flavour": nm[8:] + " Fingers", "qty": q})
    mats = []
    rm = ali_book("Raw Materials - Ali Nawaz.xlsx")
    LAY = {"Raw Materials": ("raw", 6, 11, 12), "Consumables": ("consumable", 4, 10, 11), "Packaging": ("packaging", 2, 20, None)}
    if rm:
        seen = set()
        for tab, (kind, first, c_left, c_min) in LAY.items():
            if tab not in rm.sheetnames: continue
            ws = rm[tab]
            for r in range(first, ws.max_row + 1):
                code, nm = ws.cell(r, 2).value, ws.cell(r, 3).value
                if not isinstance(nm, str) or not nm.strip() or not code: continue
                code = str(code).strip()
                if code in seen: continue
                seen.add(code)
                num = lambda v: v if isinstance(v, (int, float)) and not isinstance(v, bool) else None
                mats.append({"code": code, "name": nm.strip(), "unit": str(ws.cell(r, 4).value or "").strip(),
                             "kind": kind, "left": num(ws.cell(r, c_left).value),
                             "min": num(ws.cell(r, c_min).value) if c_min else None})
    return {"factory": "ermine", "store": store, "mats": mats}


def engine_payload(base):
    """The workbook's routing (rows 9..38 of each product sheet) with each step's areas resolved
    the way the Capacity Engine resolves them. Checked against the workbook on 21 Nov 2026: every
    step's hours, the total labour (253.69 h) and all 10 area loads matched exactly
    (app/tools/engine_check.py)."""
    import re
    live = ermine_paths.find(ermine_paths.LIVE)
    wf = openpyxl.load_workbook(live)                 # formulas: which pack cell a row counts
    wv = workbook()
    num = lambda v: v if isinstance(v, (int, float)) and not isinstance(v, bool) else None
    sp, ar, cal = wv["Steps x Products"], wv["Areas"], wv["Calibration"]
    DEF = {sp.cell(r, 1).value: (num(sp.cell(r, 2).value), num(sp.cell(r, 3).value), num(sp.cell(r, 4).value))
           for r in range(5, 21) if sp.cell(r, 1).value}
    code = lambda a: int(str(a).split("-")[0].strip()) if a and str(a).split("-")[0].strip().isdigit() else 0
    steps = []
    for pc in [x["code"] for x in base["products"]]:
        if pc not in wf.sheetnames: continue
        f, v = wf[pc], wv[pc]
        hdr = next(r for r in range(1, 20) if f.cell(r, 1).value == "Sub-Step")
        units = {}
        for c in range(2, 14, 2):
            lab = f.cell(6, c).value
            if isinstance(lab, str) and lab.strip(): units[f.cell(6, c + 1).coordinate] = lab.strip()
        for i, r in enumerate(range(hdr + 1, hdr + 31)):
            sub = f.cell(r, 1).value
            if not sub: continue
            m = re.search(r"\(G\d+/H\d+\)\*\$([A-Z])\$6", str(f.cell(r, 12).value or ""))
            main, a1s, a2s, pct = f.cell(r, 2).value, f.cell(r, 3).value, f.cell(r, 4).value, num(v.cell(r, 5).value)
            d1, d2, dp = DEF.get(main, (None, None, None))
            if a1s: a1, w1 = code(a1s), (pct if pct is not None else 1)
            else: a1, w1 = int(d1 or 0), (1 if d2 is None else (0.5 if dp is None else dp))
            if a2s: a2, w2 = code(a2s), 1 - (pct if pct is not None else 1)
            else: a2, w2 = int(d2 or 0), 1 - (1 if d2 is None else (0.5 if dp is None else dp))
            steps.append({"product": pc, "seq": i, "sub": str(sub), "main": main, "kind": str(f.cell(r, 6).value or ""),
                          "time": num(v.cell(r, 7).value), "qty": num(v.cell(r, 8).value), "crew": num(v.cell(r, 9).value),
                          "busy": num(v.cell(r, 14).value), "passive": v.cell(r, 10).value == "Yes", "batch": num(v.cell(r, 11).value),
                          "unit": units.get(f"{m.group(1)}6") if m else None, "a1": a1, "w1": w1, "a2": a2, "w2": w2})
    cals = [{"code": cal.cell(r, 1).value, "cal": num(cal.cell(r, 2).value)} for r in range(5, 30)
            if cal.cell(r, 1).value and num(cal.cell(r, 2).value) is not None]
    areas = [{"name": f"{int(ar.cell(r, 1).value)} - {ar.cell(r, 2).value}", "code": int(ar.cell(r, 1).value),
              "parallel": num(ar.cell(r, 7).value)} for r in range(5, 40) if isinstance(ar.cell(r, 1).value, (int, float))]
    return {"factory": "ermine", "steps": steps, "cal": cals, "areas": areas}


def engine_ref_payload():
    """The Stock sheet's suggestion (K, A/B/IDEAL/COAT rows 62-72) and max-boxes (Z) formulas,
    translated to JS by xl2js, with the sheet's fixed cells. The app evaluates them on its own
    stock, orders and month figures. Checked 21 Nov 2026: 96 formulas, all equal to Excel."""
    import datetime as _dt, xl2js
    live = ermine_paths.find(ermine_paths.LIVE)
    wf, wv = openpyxl.load_workbook(live), workbook()
    st, stv, ob, mt = wf["Stock"], wv["Stock"], wv["Order Book"], wv["Monthly Targets"]
    ser = lambda d: (d - _dt.datetime(1899, 12, 30)).days if isinstance(d, _dt.datetime) else d
    statics = {}
    for row in stv.iter_rows(min_row=1, max_row=90, max_col=32):
        for c in row:
            if c.value is None: continue
            if c.column_letter in ("J", "N") and 13 <= c.row <= 56: continue      # stock: the app's own
            statics[f"{c.column_letter}{c.row}"] = ser(c.value) if isinstance(c.value, _dt.datetime) else c.value
    jobs = []
    def add(k, cat, key):
        f = st[k].value
        if isinstance(f, str) and f.startswith("=") and key:
            jobs.append({"k": k, "cat": cat, "key": str(key).strip(), "js": xl2js.to_js(f)})
    for r in range(13, 57): add(f"K{r}", "flavour", st.cell(r, 2).value)
    for r in range(62, 73):
        for c, cat in (("B", "A"), ("F", "B"), ("G", "IDEAL"), ("C", "COAT")): add(f"{c}{r}", cat, st.cell(r, 1).value)
    for r in range(13, 40): add(f"Z{r}", "maxbox", st.cell(r, 24).value)
    obcols = {ob.cell(6, c).column_letter: str(ob.cell(6, c).value).strip() for c in range(2, 40) if ob.cell(6, c).value}
    mtv = {f"{c.column_letter}{c.row}": c.value for row in mt.iter_rows(min_row=1, max_row=60, max_col=8)
           for c in row if isinstance(c.value, (int, float)) and not isinstance(c.value, bool)}
    rowfl = {str(r): str(st.cell(r, 2).value).strip() for r in range(13, 57) if st.cell(r, 2).value}
    return {"factory": "ermine", "data": {"statics": statics, "jobs": jobs, "obcols": obcols, "mt": mtv,
            "rowfl": rowfl, "helpers": xl2js.HELPERS}}


def post(key, fn, body):
    req = urllib.request.Request(URL + "/rest/v1/rpc/" + fn, method="POST", data=json.dumps({"p": body}).encode(),
                                 headers={"apikey": key, "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)


def main():
    key = secret()
    if not key:
        print("db sync skipped: no server key"); return
    try:
        wb = workbook()
        if wb is None:
            print("db sync skipped: workbook not found"); return
        p = payload(wb)
        if "--dry-run" in sys.argv:
            q = ref_payload(wb, p)
            print("db sync dry run:", len(p["orders"]), "order lines,", sum(x["qty"] for x in p["t_box"]), "target boxes,",
                  len(p["stations"]), "areas,", len(q["steps"]), "steps,", len(q["staff"]), "staff,", len(q["make"]), "make rates,",
                  len(q["pack"]), "pack rates,", len(q["fl"]), "flavours,", len(q["bx"]), "boxes"); 
            z = store_payload(p); print("store items", len(z["store"]), [x["name"] for x in z["store"]][:30], "materials", len(z["mats"]),
                  {k: sum(1 for m in z["mats"] if m["kind"] == k) for k in ("raw", "consumable", "packaging")}); return
        res = post(key, "sync_from_excel", p)
        res.update(post(key, "sync_ref", ref_payload(wb, p)))
        res.update(post(key, "sync_box_flavours", {"factory": "ermine", "rows": box_flavour_rows(wb)}))
        res.update(post(key, "sync_store", store_payload(p)))
        res.update(post(key, "sync_engine", engine_payload(p)))
        res.update(post(key, "sync_engine_ref", engine_ref_payload()))
        hist = os.environ.get("ERMINE_HISTORY")              # the dashboard's day history (repo data/flavours.json)
        if hist and os.path.exists(hist):
            H = json.load(open(hist, encoding="utf-8"))
            res.update(post(key, "sync_history", {"factory": "ermine",
                       "days": [{"day": k, "data": v} for k, v in H.items() if isinstance(v, dict)]}))
        print("db synced:", ", ".join(f"{k} {v}" for k, v in res.items()))
    except Exception as e:
        msg = str(e).replace(key, "***")
        print("db sync failed (publish continues):", msg[:200])


if __name__ == "__main__":
    main()
