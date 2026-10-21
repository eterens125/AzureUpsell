#!/usr/bin/env python3
"""Build the Azure Upsell Leaderboard page from the Domino's "UPSELL BY TM" reports.

Usage:  python3 build.py [--rollup]   # reads ./data, writes ./dist/index.html

Inputs (drop new files in ./data):
  * a multi-sheet period workbook  - one sheet per coupon code (sheet name = code)
  * single-day reports in data/daily/ - one file per coupon, named  YYYY-MM-DD_<code>.xlsx
    (single sheet "repUpsell"; the coupon code is taken from the file name)
Every report carries its own Begin/End Date in the header; overlapping dates are refused
so a day can't be counted twice.
"""
import glob, json, os, re, sys, datetime as dt
import openpyxl
import pandas as pd

ROOT = os.path.dirname(os.path.abspath(__file__))
# Domino's 4-week fiscal periods. A report must fall wholly inside one of them.
PERIOD_DEFS = [
    dict(id="P10", name="Period 10", start="2026-09-07", end="2026-10-04"),
    dict(id="P11", name="Period 11", start="2026-10-05", end="2026-11-02"),
]
MIN_COUPON_ORDER = ["917526", "8255", "1126", "5385", "4342", "8210", "8253", "8225", "5073", "8666", "DIP0226"]


def nice(n):
    """'Mills, Timothy' -> 'Timothy M.' (first name + last initial only)."""
    if "," in n:
        last, first = [x.strip() for x in n.split(",", 1)]
        return f"{first} {last[:1].upper()}."
    return n


def header_dates(ws):
    text = " ".join(str(c) for r in ws.iter_rows(min_row=1, max_row=2, values_only=True) for c in r if c)
    b = re.search(r"Begin Date:\s*(\d+/\d+/\d+)", text)
    e = re.search(r"End Date:\s*(\d+/\d+/\d+)", text)
    if not (b and e):
        raise ValueError("no Begin/End Date in header")
    p = lambda s: dt.datetime.strptime(s, "%m/%d/%Y").date()
    return p(b.group(1)), p(e.group(1))


def parse_sheet(ws, coupon):
    rows, store = [], None
    for r in ws.iter_rows(min_row=5, values_only=True):
        name, code, prod, qty, ordr, tm = r[1], r[3], r[5], r[6], r[7], r[8]
        if name and prod in ("", None) and qty in ("", None):
            store = name                      # store header row
            continue
        if name in ("", None):
            continue                          # store total row (checked below)
        if name and prod:
            rows.append(dict(coupon=coupon, store=store, name=name, code=code or "",
                             ord=float(ordr or 0), tm=float(tm or 0), internet=(name == "Internet")))
    return rows


def load_sources(through=None):
    """Parse every report in ./data. `through` is the last day already folded into
    baseline.json: reports ending on/before it are skipped, one that straddles it is refused."""
    files = sorted(glob.glob(os.path.join(ROOT, "data", "*.xlsx")) + glob.glob(os.path.join(ROOT, "data", "daily", "*.xlsx")))
    if not files:
        return pd.DataFrame(), []
    # pass 1: read each sheet's own Begin/End Date header (workbooks kept open for pass 2)
    meta = []
    for f in files:
        wb = openpyxl.load_workbook(f, data_only=True)
        for ws in wb.worksheets:
            coupon = (re.split(r"[_-]", os.path.splitext(os.path.basename(f))[0])[-1] if ws.title == "repUpsell" else ws.title).upper()
            start, end = header_dates(ws)
            meta.append(dict(f=f, ws=ws, coupon=coupon, start=start, end=end))
    # Domino's reports sometimes share a boundary day: one report's End Date is the next
    # report's Begin Date for the same coupon. Rather than refuse (or double-count that
    # day), nudge the later report's effective start to the day after - its full totals
    # still can't be split by day, so that one day's count comes from the earlier report.
    by_coupon_end = {}
    for m in sorted(meta, key=lambda m: (m["coupon"], m["start"])):
        c = m["coupon"]
        prev_end = by_coupon_end.get(c, through)
        if prev_end is not None and m["start"] == prev_end:
            m["start"] = m["start"] + dt.timedelta(days=1)
        by_coupon_end[c] = max(by_coupon_end.get(c, m["end"]), m["end"])
    if through is not None:
        skipped = {os.path.basename(m["f"]) for m in meta if m["end"] <= through}
        if skipped:
            print(f"  skipping {len(skipped)} report file(s) already in baseline.json (through {through})")
        straddle = [m for m in meta if m["end"] > through >= m["start"]]
        if straddle:
            m = straddle[0]
            sys.exit(f"{os.path.basename(m['f'])} ({m['start']}..{m['end']}) overlaps days already in baseline.json (through {through}); a day would be counted twice.")
        meta = [m for m in meta if m["end"] > through]
    # pass 2: parse rows using the (possibly nudged) start date
    rows, sources = [], []
    for m in meta:
        f, ws, coupon, start, end = m["f"], m["ws"], m["coupon"], m["start"], m["end"]
        part = parse_sheet(ws, coupon)
        # sanity: rows must equal the report's own store total rows
        tot = {}
        store = None
        for r in ws.iter_rows(min_row=5, values_only=True):
            if r[1] and r[5] in ("", None) and r[6] in ("", None):
                store = r[1]
            elif r[1] in ("", None) and r[6] not in ("", None):
                tot[store] = (float(r[7] or 0), float(r[8] or 0))
        for s, (o, t) in tot.items():
            po = sum(x["ord"] for x in part if x["store"] == s)
            pt = sum(x["tm"] for x in part if x["store"] == s)
            if (po, pt) != (o, t):
                raise ValueError(f"{os.path.basename(f)} [{coupon}] store {s}: rows ({po},{pt}) != report total ({o},{t})")
        for x in part:
            x["start"], x["end"] = start, end
        rows += part
        sources.append((start, end, os.path.basename(f), coupon))
    return pd.DataFrame(rows), sources


def merge_ranges(ranges):
    out = []
    for s, e in sorted(set(ranges)):
        if out and s <= out[-1][1] + dt.timedelta(days=1):
            out[-1][1] = max(out[-1][1], e)
        else:
            out.append([s, e])
    return out


def check_overlap(sources):
    per = {}
    for s, e, f, c in sources:
        per.setdefault(c, []).append((s, e, f))
    for c, lst in per.items():
        lst.sort()
        for a, b in zip(lst, lst[1:]):
            if b[0] <= a[1]:
                sys.exit(f"Coupon {c}: {a[2]} ({a[0]}..{a[1]}) overlaps {b[2]} ({b[0]}..{b[1]}); a day would be counted twice.")


def fmt_range(s, e):
    d = lambda x: x.strftime("%b ") + str(x.day)
    if s == e:
        return d(s)
    return f"{d(s)} &ndash; {d(e)}" if s.month != e.month else f"{d(s)} &ndash; {e.day}"


def fmt_range_txt(s, e):
    """Like fmt_range but a plain '-' character for JS chart labels (no HTML entity)."""
    d = lambda x: x.strftime("%b ") + str(x.day)
    if s == e:
        return d(s)
    return f"{d(s)} – {d(e)}" if s.month != e.month else f"{d(s)} – {e.day}"


def build_trend(df):
    """One data point per report period (each distinct Begin/End Date range ingested),
    in chronological order, so the page can chart progression update-by-update."""
    tm = df[~df.internet]
    stores = sorted(df.store.unique())
    periods = sorted(set(zip(tm.start, tm.end)))
    labels = [fmt_range_txt(s, e) for s, e in periods]
    azure, per_store = [], {s: [] for s in stores}
    for s0, e0 in periods:
        sub = tm[(tm.start == s0) & (tm.end == e0)]
        coupon_set = set(sub.coupon)
        ref = "5073" if "5073" in coupon_set else sorted(coupon_set)[0]
        a_ups = a_tm = 0
        for st in stores:
            t = sub[sub.store == st]
            ups, tmo = int(t.ord.sum()), int(t[t.coupon == ref].tm.sum())
            per_store[st].append(dict(ups=ups, tmOrd=tmo))
            a_ups += ups
            a_tm += tmo
        azure.append(dict(ups=a_ups, tmOrd=a_tm))
    return dict(labels=labels, azure=azure, stores=per_store)


def build_data(df, with_codes=False):
    coupons = MIN_COUPON_ORDER + sorted(set(df.coupon) - set(MIN_COUPON_ORDER))
    stores = sorted(df.store.unique())
    tm, net = df[~df.internet], df[df.internet]
    ref = "5073" if "5073" in set(df.coupon) else sorted(set(df.coupon))[0]  # every sheet repeats each person's TM Ord
    S = []
    for s in stores:
        t, n = tm[tm.store == s], net[net.store == s]
        by = {}
        for c in coupons:
            g = t[(t.coupon == c) & (t.ord > 0)].groupby("name").ord.sum().reset_index().sort_values(["ord", "name"], ascending=[False, True])
            by[c] = dict(tm=int(t[t.coupon == c].ord.sum()), net=int(n[n.coupon == c].ord.sum()),
                         top=[dict(n=nice(r["name"]), v=int(r["ord"])) for _, r in g.iterrows()])
        team = []
        for (name, code), g in t.groupby(["name", "code"]):
            tmo = int(g[g.coupon == ref].tm.sum())
            if tmo > 0:
                team.append(dict(n=nice(name), o=int(g.ord.sum()), t=tmo))
        st = dict(id=s, tmOrd=int(t[t.coupon == ref].tm.sum()), ups=int(t.ord.sum()), netOrd=int(n.ord.sum()),
                  netTm=int(n[n.coupon == ref].tm.sum()), by=by, team=team)
        assert sum(x["o"] for x in team) == st["ups"] and sum(x["t"] for x in team) == st["tmOrd"], s
        S.append(st)
    people = []
    for (name, code), g in tm.groupby(["name", "code"]):
        tmo = int(g[g.coupon == ref].tm.sum())
        if tmo > 0:
            p = dict(name=nice(name), stores=sorted(g.store.unique()), ord=int(g.ord.sum()), tm=tmo)
            if with_codes:
                p["code"] = code
            people.append(p)
    return dict(stores=S, coupons=coupons, people=people, trend=build_trend(df))


def merge_data(base, new, notes):
    """Add the DATA of newly ingested reports (`new`, people carry employee `code`) onto the
    aggregate baseline DATA. Everything is keyed by 'First L.'; ambiguous people matches are
    appended to `notes` for a human to look at."""
    out = json.loads(json.dumps(base))
    coupons = list(out["coupons"]) + [c for c in new["coupons"] if c not in out["coupons"]]
    out["coupons"] = coupons
    nstores = {s["id"]: s for s in new["stores"]}
    for bs in out["stores"]:
        ns = nstores.pop(bs["id"], None)
        if ns is None:
            continue
        for k in ("tmOrd", "ups", "netOrd", "netTm"):
            bs[k] += ns[k]
        for c in coupons:
            b = bs["by"].setdefault(c, dict(tm=0, net=0, top=[]))
            n = ns["by"].get(c, dict(tm=0, net=0, top=[]))
            b["tm"] += n["tm"]; b["net"] += n["net"]
            tot = {}
            for x in b["top"] + n["top"]:
                tot[x["n"]] = tot.get(x["n"], 0) + x["v"]
            b["top"] = [dict(n=k, v=v) for k, v in sorted(tot.items(), key=lambda kv: (-kv[1], kv[0]))]
        team = {x["n"]: dict(x) for x in bs["team"]}
        for x in ns["team"]:
            if x["n"] in team:
                team[x["n"]]["o"] += x["o"]; team[x["n"]]["t"] += x["t"]
            else:
                team[x["n"]] = dict(x)
        bs["team"] = list(team.values())
    if nstores:
        sys.exit(f"new reports contain store(s) {sorted(nstores)} that are not in baseline.json")
    # people: the baseline keeps no employee codes, so a new record is merged into an existing
    # "First L." only when they share a store; same name with no shared store becomes its own
    # entry (and is reported), the same way two same-named people are already kept apart.
    people = [dict(p) for p in out["people"]]
    claimed = {}
    for np_ in new["people"]:
        same = [p for p in people if p["name"] == np_["name"]]
        hit = [p for p in same if set(p["stores"]) & set(np_["stores"])]
        if not hit:
            if same:
                notes.append(f"{np_['name']} ({','.join(np_['stores'])}, {np_['ord']}/{np_['tm']}) has the same name as an existing person "
                             f"at other store(s) - kept as a separate person")
            people.append({k: v for k, v in np_.items() if k != "code"})
            continue
        if len(hit) > 1:
            notes.append(f"{np_['name']} ({','.join(np_['stores'])}) shares a store with {len(hit)} existing people - merged into the largest")
        p = max(hit, key=lambda p: p["tm"])
        if id(p) in claimed and claimed[id(p)] != np_["code"]:
            notes.append(f"{np_['name']}: two different employee codes in the new reports share a store with one existing person - check")
        claimed[id(p)] = np_["code"]
        p["ord"] += np_["ord"]; p["tm"] += np_["tm"]
        p["stores"] = sorted(set(p["stores"]) | set(np_["stores"]))
    out["people"] = people
    t, nt = out["trend"], new["trend"]
    t["labels"] += nt["labels"]
    t["azure"] += nt["azure"]
    for sid, pts in nt["stores"].items():
        t["stores"][sid] += pts
    # invariants
    for s in out["stores"]:
        assert sum(x["o"] for x in s["team"]) == s["ups"] and sum(x["t"] for x in s["team"]) == s["tmOrd"], s["id"]
        assert sum(x["v"] for x in s["by"][coupons[0]]["top"]) == s["by"][coupons[0]]["tm"], s["id"]
    assert sum(p["ord"] for p in out["people"]) == sum(s["ups"] for s in out["stores"])
    assert sum(p["tm"] for p in out["people"]) == sum(s["tmOrd"] for s in out["stores"])
    assert len(t["labels"]) == len(t["azure"]) == len(next(iter(t["stores"].values())))
    assert sum(a["ups"] for a in t["azure"]) == sum(s["ups"] for s in out["stores"])
    return out


def sources_text(spans):
    multi = sorted({(s, e) for s, e in spans if s != e})
    single = sorted({s for s, e in spans if s == e})
    parts = []
    if multi:
        parts.append("the UPSELL BY TM report (" + ", ".join(fmt_range(s, e) for s, e in multi) + ")")
    if single:
        parts.append("single-day coupon reports for " + ", ".join(fmt_range(d, d) for d in single))
    return " plus ".join(parts)


def load_baseline():
    """data/baseline.json holds, per period, only 'First L.' aggregates (no names, no employee codes)."""
    path = os.path.join(ROOT, "data", "baseline.json")
    saved = {p["id"]: p for p in json.load(open(path, encoding="utf-8"))["periods"]} if os.path.exists(path) else {}
    periods = []
    for d in PERIOD_DEFS:
        p = saved.get(d["id"], {})
        day = dt.date.fromisoformat
        periods.append(dict(d, data=p.get("data"),
                            spans=[(day(a), day(b)) for a, b in p.get("spans", [])],
                            through=day(p["through"]) if p.get("through") else day(d["start"]) - dt.timedelta(days=1)))
    return path, periods


def main():
    args = sys.argv[1:]
    rollup = "--rollup" in args                      # fold current reports into data/baseline.json
    frag = args[args.index("--fragment") + 1] if "--fragment" in args else None
    bpath, periods = load_baseline()
    through = max(p["through"] for p in periods)
    df, sources = load_sources(through)
    check_overlap(sources)
    notes = []
    for p in periods:
        lo, hi = dt.date.fromisoformat(p["start"]), dt.date.fromisoformat(p["end"])
        if len(df):
            sub = df[(df.start >= lo) & (df.start <= hi)]
            if len(sub) and (sub.end > hi).any():
                bad = sub[sub.end > hi].iloc[0]
                sys.exit(f"A report ({bad.start}..{bad.end}) starts in {p['name']} but runs past its end ({hi}); reports must sit inside one period.")
        else:
            sub = df
        if len(sub):
            if p["data"]:
                p["data"] = merge_data(p["data"], build_data(sub, with_codes=True), notes)
            else:
                p["data"] = build_data(sub)
            p["spans"] += sorted({(s, e) for s, e, _, _ in sources if lo <= s <= hi})
            p["through"] = max(p["through"], max(e for s, e, _, _ in sources if lo <= s <= hi))
    covered = len(df) and sum(len(df[(df.start >= dt.date.fromisoformat(p["start"])) & (df.start <= dt.date.fromisoformat(p["end"]))]) for p in periods)
    if len(df) and covered != len(df):
        sys.exit("some reports fall outside every defined period (see PERIOD_DEFS in build.py)")
    live = [p for p in periods if p["data"]]
    if not live:
        sys.exit("no reports in ./data and no baseline.json")
    for p in live:
        for x in p["data"]["people"]:
            x.pop("code", None)                      # employee codes never leave the raw reports
    page = [dict(id=p["id"], name=p["name"], start=p["start"], end=p["end"], final=p["through"] >= dt.date.fromisoformat(p["end"]),
                 sources=sources_text(p["spans"]), data=p["data"]) for p in live]
    html = open(os.path.join(ROOT, "template.html"), encoding="utf-8").read()
    assert "/*DATA*/" in html
    html = html.replace("/*DATA*/", json.dumps(page, separators=(",", ":")))
    if frag:                                          # bare fragment for the Claude Artifact publish flow
        open(frag, "w", encoding="utf-8").write(html)
    # template.html is a bare content fragment (no doctype/html/head/body) by design,
    # since the Claude Artifact publish flow supplies its own page skeleton with a
    # UTF-8 charset at publish time. This standalone GitHub build has no such wrapper,
    # so without an explicit <meta charset> a browser opening dist/index.html directly
    # (or via a host that doesn't send a charset header) can guess the wrong encoding
    # and mangle multi-byte characters like the em dash into "â€”" mojibake. Wrap it in
    # a real document here so the charset is always declared.
    html = (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n<head>\n<meta charset="UTF-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f"{html}\n</head>\n<body></body>\n</html>\n"
    )
    out = os.path.join(ROOT, "dist", "index.html")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    open(out, "w", encoding="utf-8").write(html)
    if rollup and len(df):
        json.dump(dict(periods=[dict(id=p["id"], through=p["through"].isoformat(), data=p["data"],
                                     spans=[[a.isoformat(), b.isoformat()] for a, b in p["spans"]]) for p in periods if p["data"]]),
                  open(bpath, "w", encoding="utf-8"), separators=(",", ":"))
        print("  baseline.json updated: " + ", ".join(f"{p['id']} through {p['through']}" for p in periods if p["data"]))
    print(f"Built {out}")
    for p in live:
        d = p["data"]
        ups = sum(s["ups"] for s in d["stores"]); tmo = sum(s["tmOrd"] for s in d["stores"])
        print(f"  {p['name']}: through {p['through']}, team-member upsell orders {ups} of {tmo} ({ups / tmo:.1%})")
    for n in notes:
        print("  NOTE:", n)


if __name__ == "__main__":
    main()
