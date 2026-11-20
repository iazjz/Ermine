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
                  len(q["pack"]), "pack rates,", len(q["fl"]), "flavours,", len(q["bx"]), "boxes"); return
        res = post(key, "sync_from_excel", p)
        res.update(post(key, "sync_ref", ref_payload(wb, p)))
        res.update(post(key, "sync_box_flavours", {"factory": "ermine", "rows": box_flavour_rows(wb)}))
        print("db synced:", ", ".join(f"{k} {v}" for k, v in res.items()))
    except Exception as e:
        msg = str(e).replace(key, "***")
        print("db sync failed (publish continues):", msg[:200])


if __name__ == "__main__":
    main()
