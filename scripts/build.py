"""Build site/data.js from the team workbook and the processed data snapshot.

Usage:
    python scripts/build.py            # real data if data/processed/catalog.json exists, else demo data
    python scripts/build.py --demo     # force synthetic demo data (clearly flagged on the board)

Inputs
    content/team.xlsx                  team workbook (row 3 = headers, data from row 4;
                                       IDs starting EX = examples, Status 'Rejected' = hidden)
    data/processed/catalog.json        list of series: {id, label, unit, source_id, frequency, file}
    data/processed/<file>.csv          columns: period (YYYY-MM), value, flag
Output
    site/data.js                       window.BOARD = {...}; loaded by site/index.html (works from file://)
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import math
import pathlib
import random
import sys

from openpyxl import load_workbook

ROOT = pathlib.Path(__file__).resolve().parents[1]
CONTENT = ROOT / "content" / "team.xlsx"
PROC = ROOT / "data" / "processed"
OUT = ROOT / "site" / "data.js"

TARGET_END = "2027-12"  # forecast to the end of next year (decision horizon)
FIT_WINDOW = 48     # trend methods are fitted on the last 4 years (market structure changed after 2021)
HOLDOUT = 12        # months hidden for validation
Z80 = 1.2816        # indicative 80% band multiplier


# ----------------------------------------------------------------- workbook
def _norm(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    return v


def read_table(wb, sheet: str) -> list[dict]:
    """Rows from a content sheet, skipping blanks, EX examples and Rejected rows."""
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    headers = [_norm(c.value) for c in ws[3]]
    rows = []
    for raw in ws.iter_rows(min_row=4, values_only=True):
        vals = [_norm(v) for v in raw]
        if all(v is None for v in vals):
            continue
        first = str(vals[0] or "")
        if first.upper().startswith("EX"):
            continue
        row = {h: v for h, v in zip(headers, vals) if h}
        if str(row.get("Status") or "").lower() == "rejected":
            continue
        rows.append(row)
    return rows


def num(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def read_content() -> dict:
    wb = load_workbook(CONTENT, data_only=True)
    t = lambda s: read_table(wb, s)

    brief = {r["Field"]: r.get("Value") for r in t("Brief") if r.get("Field")}

    targets = [r for r in t("Forecast target") if str(r.get("ID") or "").startswith("F")]
    decision, reason = None, None
    if "Forecast target" in wb.sheetnames:
        ws = wb["Forecast target"]
        for row in ws.iter_rows(min_row=4):
            if row[0].value == "Team decision":
                decision = _norm(row[1].value)
            if row[0].value == "Reason":
                reason = _norm(row[1].value)

    scen = []
    for r in t("Scenarios"):
        if not r.get("ID"):
            continue
        scen.append({
            "id": r["ID"], "name": r.get("Name") or r["ID"], "story": r.get("Story (2-3 sentences)") or "",
            "u1": r.get("U1 position"), "u2": r.get("U2 position"),
            "d": num(r.get("Export volume vs baseline (% per yr)")),
            "p": num(r.get("Export price (% per yr)")),
            "m": num(r.get("Margin change (pp)")),
            "why": r.get("Reasoning and evidence for the numbers") or "", "src": r.get("Source IDs") or "",
            "status": r.get("Status"),
        })

    comp = {r["ID"]: r for r in t("Company inputs") if r.get("ID")}
    def cin(i, default, lo, hi):
        r = comp.get(i, {})
        return {"value": num(r.get("Value"), default), "min": num(r.get("Slider min"), lo),
                "max": num(r.get("Slider max"), hi), "label": r.get("Parameter") or i,
                "unit": r.get("Unit") or "", "basis": r.get("Basis / source") or "", "status": r.get("Status")}
    company = {"m0": cin("C1", 3, -10, 20), "share": cin("C2", 50, 0, 100), "beta": cin("C3", 1, 0.5, 2)}

    signals = []
    for r in t("Signals"):
        if not r.get("ID"):
            continue
        eff = {s["id"]: {"d": num(r.get(f"{s['id']} volume")), "p": num(r.get(f"{s['id']} price")),
                         "m": num(r.get(f"{s['id']} margin"))} for s in scen}
        signals.append({
            "id": r["ID"], "title": r.get("Headline") or "", "publisher": r.get("Publisher") or "",
            "date": r.get("Date") or "", "url": r.get("URL") or "", "pestle": r.get("PESTLE factor") or "",
            "kind": r.get("Kind") or "Reported news", "fact": r.get("Fact as stated") or "",
            "note": r.get("Our judgement") or "", "read": r.get("Read status") or "", "eff": eff,
            "status": r.get("Status"),
        })

    options = []
    for r in t("Options"):
        if not r.get("ID"):
            continue
        sc = {}
        for s in scen:
            v = r.get(f"Score {s['id']}")
            sc[s["id"]] = None if v is None else int(num(v))
        options.append({"id": r["ID"], "name": r.get("Option") or "", "what": r.get("What it means") or "",
                        "scores": sc, "sig": r.get("Signpost or signal to act / change course") or "",
                        "why": r.get("Reason for scores") or "", "bench": r["ID"] == "O0", "status": r.get("Status")})

    return {
        "brief": brief,
        "targets": targets, "targetDecision": decision, "targetReason": reason,
        "sources": t("Sources"), "checks": t("Data checks"),
        "pestle": t("PESTLE"), "diamond": t("Diamond"),
        "uncertainties": t("Uncertainties"), "scenarios": scen, "company": company,
        "signals": signals, "signposts": t("Signposts"), "options": options,
        "log": t("Log"), "usertest": t("User test"), "ai": t("AI use"), "contrib": t("Contributions"),
    }


# ----------------------------------------------------------------- data
def read_series_real() -> list[dict] | None:
    cat = PROC / "catalog.json"
    if not cat.exists():
        return None
    out = []
    for meta in json.loads(cat.read_text(encoding="utf-8")):
        time, val, flag = [], [], []
        with open(PROC / meta["file"], newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                if row.get("value") in (None, ""):
                    continue
                time.append(row["period"][:7])
                val.append(float(row["value"]))
                flag.append(row.get("flag") or "")
        time, val, flag = fill_gaps(time, val, flag)
        out.append({**meta, "time": time, "val": val, "flag": flag, "role": meta.get("role", "series")})
    return out


def fill_gaps(time, val, flag):
    """Insert missing months between first and last observation by linear interpolation (flag 'i').
    Needed because China publishes January and February combined, so January is missing in Comtrade."""
    if not time:
        return time, val, flag
    have = dict(zip(time, zip(val, flag)))
    t, v, f = [], [], []
    p = time[0]
    while p <= time[-1]:
        t.append(p)
        if p in have:
            v.append(have[p][0]); f.append(have[p][1])
        else:
            v.append(None); f.append("i")
        p = add_months(p, 1)
    for i, x in enumerate(v):
        if x is None:
            lo = max(j for j in range(i) if v[j] is not None)
            hi = min(j for j in range(i + 1, len(v)) if v[j] is not None)
            v[i] = v[lo] + (v[hi] - v[lo]) * (i - lo) / (hi - lo)
    return t, v, f


def add_months(p: str, k: int) -> str:
    y, m = map(int, p.split("-"))
    m = m - 1 + k
    return f"{y + m // 12}-{m % 12 + 1:02d}"


def demo_series() -> list[dict]:
    """Synthetic series so the frames render before real data exists. Never present as evidence."""
    rnd = random.Random(7)
    time = [add_months("2022-01", k) for k in range(57)]  # to 2026-09
    vol, uv = [], []
    for k in range(len(time)):
        season = 1 + 0.08 * math.sin(2 * math.pi * (k - 2) / 12)
        vol.append(round((18 + 0.25 * k) * season * (1 + rnd.uniform(-0.04, 0.04)), 2))
        uv.append(round(max(0.07, 0.22 * math.exp(-0.025 * k)) * (1 + rnd.uniform(-0.03, 0.03)), 4))
    return [
        {"id": "demo_export_volume", "label": "DEMO: China PV module exports, volume", "unit": "GW (demo)",
         "source_id": "demo", "frequency": "Monthly", "time": time, "val": vol, "flag": [""] * len(time)},
        {"id": "demo_unit_value", "label": "DEMO: China PV module export unit value", "unit": "USD/W (demo)",
         "source_id": "demo", "frequency": "Monthly", "time": time, "val": uv, "flag": [""] * len(time)},
    ]


# ----------------------------------------------------------------- forecasting
def mean(a):
    return sum(a) / len(a)


def linfit(y):
    n = len(y)
    x = list(range(1, n + 1))
    xm, ym = mean(x), mean(y)
    b = sum((xi - xm) * (yi - ym) for xi, yi in zip(x, y)) / sum((xi - xm) ** 2 for xi in x)
    return ym - b * xm, b  # intercept, slope (x = 1..n)


def f_snaive(y, h):
    out = []
    hist = list(y)
    for _ in range(h):
        out.append(hist[-12])
        hist.append(hist[-12])
    return out


def f_ma12(y, h):
    return [mean(y[-12:])] * h


def f_trend(y, h):
    y = y[-FIT_WINDOW:]
    a, b = linfit(y)
    n = len(y)
    return [a + b * (n + k) for k in range(1, h + 1)]


def f_trend_season(y, h):
    """Linear trend + additive monthly seasonal index from detrended residuals (needs >= 24 points)."""
    y = y[-FIT_WINDOW:]
    a, b = linfit(y)
    n = len(y)
    resid = [y[i] - (a + b * (i + 1)) for i in range(n)]
    idx = [mean(resid[m::12]) for m in range(12)]
    adj = mean(idx)
    idx = [v - adj for v in idx]
    return [a + b * (n + k) + idx[(n + k - 1) % 12] for k in range(1, h + 1)]


METHODS = {
    "snaive": ("Seasonal naive (benchmark)", f_snaive),
    "ma12": ("12-month moving average", f_ma12),
    "trend": ("Linear trend", f_trend),
    "trend_season": ("Linear trend + seasonal index", f_trend_season),
}


def months_between(a: str, b: str) -> int:
    ya, ma = map(int, a.split("-")); yb, mb = map(int, b.split("-"))
    return (yb - ya) * 12 + (mb - ma)


def forecast_series(s: dict) -> dict:
    y = s["val"]
    n = len(y)
    HORIZON = max(12, months_between(s["time"][-1], TARGET_END))
    H = HOLDOUT if n >= 36 else max(6, n // 4)
    train, test = y[:-H], y[-H:]
    res = {}
    for key, (label, fn) in METHODS.items():
        if len(train) < 24:
            continue
        hf = fn(train, H)
        ape = [abs(f - a) / abs(a) for f, a in zip(hf, test) if a]
        se = [(f - a) ** 2 for f, a in zip(hf, test)]
        rmse = math.sqrt(mean(se))
        ahead = fn(y, HORIZON)
        w = [Z80 * rmse * math.sqrt(max(k, 1) / H) for k in range(1, HORIZON + 1)]
        res[key] = {"label": label, "mape": mean(ape) * 100, "rmse": rmse,
                    "holdoutFc": hf, "ahead": ahead,
                    "lo": [f - wi for f, wi in zip(ahead, w)], "hi": [f + wi for f, wi in zip(ahead, w)]}
    bench = res.get("snaive", {}).get("mape", float("inf"))
    better = [k for k in res if k != "snaive" and res[k]["mape"] < bench]
    default = min(better, key=lambda k: res[k]["mape"]) if better else "snaive"
    return {"holdout": H, "holdoutStart": s["time"][-H], "methods": res, "default": default, "horizon": HORIZON, "fitWindow": FIT_WINDOW,
            "aheadTime": [add_months(s["time"][-1], k) for k in range(1, HORIZON + 1)]}


def add_annual_check(series):
    """Out-of-sample check: compare each method's 12-month forecast total with the annual total
    published later (data/raw/comtrade/annual_*.json). Only for flow series (kg, usd)."""
    files = sorted((ROOT / "data" / "raw" / "comtrade").glob("annual_*.json"))
    if not files:
        return
    a = json.loads(files[-1].read_text(encoding="utf-8"))
    for s in series:
        key, scale = {"cn_pv_exports_kg": ("kg", 1e6), "cn_pv_exports_usd": ("usd", 1e9)}.get(s["id"], (None, 0))
        if not key:
            continue
        year = s["fc"]["aheadTime"][0][:4]
        if not all(year in a[c] for c in ("854142", "854143")) or s["fc"]["aheadTime"][-1][:4] != year:
            continue
        actual = sum(a[c][year][key] for c in ("854142", "854143")) / scale
        s["annualCheck"] = {"year": year, "actual": actual, "source": a.get("source", ""),
                            "methods": {k: {"label": m["label"], "forecast": sum(m["ahead"]),
                                            "error": sum(m["ahead"]) / actual - 1} for k, m in s["fc"]["methods"].items()}}


def crosscheck(series):
    """Ember (China customs) vs UN Comtrade (China's report to the UN) for overlapping months and 2025."""
    kg = next((s for s in series if s["id"] == "cn_pv_exports_kg"), None)
    usd = next((s for s in series if s["id"] == "cn_pv_exports_usd"), None)
    raw = ROOT / "data" / "raw" / "comtrade"
    mf, af = raw / "2026-11-24.json", raw / "annual_2026-11-24.json"
    if not (kg and usd and mf.exists()):
        return None
    rows = json.loads(mf.read_text(encoding="utf-8"))
    ct = {}
    for r in rows:
        p = f"{r['period'][:4]}-{r['period'][4:]}"
        o = ct.setdefault(p, {"kg": 0, "usd": 0}); o["kg"] += r["netWgt"] / 1e6; o["usd"] += r["primaryValue"] / 1e9
    ek, eu = dict(zip(kg["time"], kg["val"])), dict(zip(usd["time"], usd["val"]))
    months = sorted(p for p in ct if p in ek)
    diff = lambda a, b: (a / b - 1) * 100
    out = {"months": len(months), "from": months[0], "to": months[-1],
           "kg": {"ember": sum(ek[p] for p in months), "comtrade": sum(ct[p]["kg"] for p in months)},
           "usd": {"ember": sum(eu[p] for p in months), "comtrade": sum(ct[p]["usd"] for p in months)},
           "maxMonthlyDiffPct": max(max(abs(diff(ek[p], ct[p]["kg"])), abs(diff(eu[p], ct[p]["usd"]))) for p in months),
           "source": "UN Comtrade public preview API C/M/HS, reporter 156, partner 0, flow X, HS 854142+854143, captured 2026-11-24"}
    for k in ("kg", "usd"):
        out[k]["diffPct"] = diff(out[k]["ember"], out[k]["comtrade"])
    if af.exists():
        a = json.loads(af.read_text(encoding="utf-8"))
        y = "2025"
        if all(y in a[c] for c in ("854142", "854143")) and all(f"{y}-{m:02d}" in ek for m in range(1, 13)):
            ck = sum(a[c][y]["kg"] for c in ("854142", "854143")) / 1e6
            cu = sum(a[c][y]["usd"] for c in ("854142", "854143")) / 1e9
            sk = sum(ek[f"{y}-{m:02d}"] for m in range(1, 13)); su = sum(eu[f"{y}-{m:02d}"] for m in range(1, 13))
            out["annual"] = {"year": y, "kg": {"ember": sk, "comtrade": ck, "diffPct": diff(sk, ck)},
                             "usd": {"ember": su, "comtrade": cu, "diffPct": diff(su, cu)}}
    return out


# ----------------------------------------------------------------- main
def main():
    force_demo = "--demo" in sys.argv
    series = None if force_demo else read_series_real()
    demo = series is None
    if demo:
        series = demo_series()
    markets = [s for s in series if s.get("role") == "market"]
    series = [s for s in series if s.get("role") != "market"]
    for s in series:
        s["fc"] = forecast_series(s)
    xcheck = None if demo else crosscheck(series)
    board = {
        "meta": {"built": dt.datetime.now().strftime("%Y-%m-%d %H:%M"), "demo": demo,
                 "snapshot": None if demo else json.loads((PROC / "catalog.json").read_text(encoding="utf-8"))[0].get("accessed"),
                 "horizon": series[0]["fc"]["horizon"] if series else 12},
        "content": read_content(),
        "series": series,
        "markets": markets,
        "crosscheck": xcheck,
    }
    OUT.write_text("window.BOARD = " + json.dumps(board, ensure_ascii=False, indent=1) + ";\n", encoding="utf-8")
    print(f"wrote {OUT.relative_to(ROOT)}  demo={demo}  series={len(series)}")
    for s in series:
        f = s["fc"]
        print(f"  {s['id']}: n={len(s['val'])} holdout={f['holdout']} default={f['default']} " +
              ", ".join(f"{k} MAPE {v['mape']:.2f}%" for k, v in f["methods"].items()))


if __name__ == "__main__":
    main()
