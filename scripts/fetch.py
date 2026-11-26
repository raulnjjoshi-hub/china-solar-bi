"""Download public data into data/raw (dated) and write tidy series to data/processed.

STATUS: Ember is the main source (China customs exports by destination + China solar generation),\nrefreshed by download. Comtrade (preview API) is kept for the cross-check. Eurostat and NBS are TODO.
Until catalog.json exists, build.py runs in demo mode.

    python scripts/fetch.py                       # fetch from the API
    python scripts/fetch.py --from-raw FILE.json  # rebuild from a saved Comtrade response

Output contract (read by build.py)
    data/processed/catalog.json   [{id, label, unit, source_id, frequency, file, accessed}]
    data/processed/<file>.csv     period,value,flag    (period = YYYY-MM, oldest first, no gaps)

Rules
    - Never alter source values. Record every transformation (unit change, Jan-Feb split, sums)
      in team.xlsx > Sources > "Transformations applied".
    - Keep the raw response in data/raw/<source>/<YYYY-MM-DD>.json so a reviewer can trace values.
    - Secrets (e.g. COMTRADE_KEY) come from environment variables, never from the repo.
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import json
import os
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
PROC = ROOT / "data" / "processed"
TODAY = dt.date.today().isoformat()


COMTRADE_PREVIEW = "https://comtradeapi.un.org/public/v1/preview/C/M/HS"
COMTRADE_FULL = "https://comtradeapi.un.org/data/v1/get/C/M/HS"
HS_CODES = {"854142": "cells", "854143": "modules"}


def _months(start="2022-01"):
    y, m = map(int, start.split("-"))
    today = dt.date.today()
    while (y, m) <= (today.year, today.month):
        yield f"{y}{m:02d}"
        m += 1
        if m > 12:
            y, m = y + 1, 1


def fetch_comtrade_exports() -> list[dict]:
    """China (156) exports to World (0) of HS 854142 cells and 854143 modules, monthly from 2022-01.
    Uses the free preview endpoint (no key, one period per call, ~1 call / 3 s). If COMTRADE_KEY is
    set, uses the keyed endpoint instead. Raw responses are saved to data/raw/comtrade/<date>.json.
    China publishes January and February combined, so January (and sometimes February) is missing:
    those months are left as gaps and filled in build.py by linear interpolation (flag 'i')."""
    import time
    import requests

    key = os.environ.get("COMTRADE_KEY")
    url = COMTRADE_FULL if key else COMTRADE_PREVIEW
    headers = {"Ocp-Apim-Subscription-Key": key} if key else {}
    rows = []
    for p in _months():
        for attempt in range(5):
            r = requests.get(url, params={"reporterCode": 156, "partnerCode": 0, "flowCode": "X",
                                          "cmdCode": ",".join(HS_CODES), "period": p}, headers=headers, timeout=60)
            if r.status_code == 429:
                time.sleep(4 + attempt * 2)
                continue
            r.raise_for_status()
            rows += r.json().get("data") or []
            break
        time.sleep(3.2)
    if len({str(d.get("period")) for d in rows}) < 24:
        raise RuntimeError(f"only {len(rows)} rows returned")
    return write_comtrade(rows)


def write_comtrade(rows: list[dict], accessed: str | None = None) -> list[dict]:
    """Save raw rows and write tidy CSVs + catalog entries. Shared by the API fetch and by a
    browser-captured JSON (data/raw/comtrade/<date>.json) when CI cannot reach Comtrade."""
    accessed = accessed or TODAY
    (RAW / "comtrade").mkdir(parents=True, exist_ok=True)
    raw_path = RAW / "comtrade" / f"{accessed}.json"
    if not raw_path.exists():
        raw_path.write_text(json.dumps(rows, indent=1), encoding="utf-8")
    by = {}
    for d in rows:
        per = str(d.get("period"))
        by.setdefault(per, {})[str(d.get("cmdCode"))] = d
    periods = sorted(by)
    out = []
    specs = [
        ("cn_pv_exports_kg", "China PV cell + module exports, net weight", "thousand tonnes",
         lambda d: (d.get("netWgt") or 0) / 1e6),
        ("cn_pv_exports_usd", "China PV cell + module exports, value", "USD billion",
         lambda d: (d.get("primaryValue") or 0) / 1e9),
        ("cn_pv_export_unit_value", "China PV cell + module exports, unit value (value / weight)", "USD per kg", None),
    ]
    for sid, label, unit, fn in specs:
        path = PROC / f"{sid}.csv"
        with open(path, "w", encoding="utf-8", newline="") as fh:
            fh.write("period,value,flag\n")
            for per in periods:
                parts = by[per]
                if not all(c in parts for c in HS_CODES):
                    continue
                v = sum(fn(parts[c]) for c in HS_CODES) if fn else (
                    sum(parts[c].get("primaryValue") or 0 for c in HS_CODES) /
                    sum(parts[c].get("netWgt") or 0 for c in HS_CODES))
                fh.write(f"{per[:4]}-{per[4:]},{v:.6f},\n")
        out.append({"id": sid, "label": label, "unit": unit, "source_id": "S1",
                    "frequency": "Monthly", "file": path.name, "accessed": accessed,
                    "note": "UN Comtrade, reporter China (156), partner World (0), flow X, HS 854142+854143"})
    return out



# ----------------------------------------------------------------- Ember (main source since 26.11.2026)
EMBER_EXPORTS = "https://files.ember-energy.org/public-downloads/customs/outputs/mart_solar_exports_full_release_monthly.csv"
EMBER_GEN = "https://files.ember-energy.org/public-downloads/generation/outputs/release_generation_monthly_global.csv"
EMBER_AREAS = {"world": ("World", "Region"), "eu": ("EU", "Region"), "us": ("United States", "Country or economy"),
               "me": ("Middle East", "Region"), "latam": ("Latin America and Caribbean", "Region"),
               "africa": ("Africa", "Region"), "asean": ("ASEAN", "Region")}
AGG_COLS = ["world_kg", "world_usd", "world_mw", "eu_kg", "us_kg", "me_kg", "latam_kg", "africa_kg", "asean_kg"]


def aggregate_ember_exports(text: str) -> list[dict]:
    """Monthly totals of Panels + Cells (Ember China Solar Exports, from China customs).
    World and regional rows are Ember's own aggregates; the US is a country row."""
    rows = csv.reader(io.StringIO(text))
    next(rows)
    m = {}
    for r in rows:
        if len(r) < 12 or r[5] not in ("Panels", "Cells"):
            continue
        per = r[1][:7]
        for k, (area, typ) in EMBER_AREAS.items():
            if r[0] == area and r[2] == typ:
                o = m.setdefault(per, {})
                o[k + "_kg"] = o.get(k + "_kg", 0) + float(r[7] or 0)
                if k == "world":
                    o["world_usd"] = o.get("world_usd", 0) + float(r[8] or 0)
                    o["world_mw"] = o.get("world_mw", 0) + float(r[11] or 0)
    return [{"period": p, **{c: round(m[p].get(c, 0)) for c in AGG_COLS}} for p in sorted(m)]


def china_solar_generation(text: str) -> list[dict]:
    out = []
    for r in csv.reader(io.StringIO(text)):
        if len(r) > 6 and r[0] == "China" and r[4] == "Solar":
            out.append({"period": r[2][:7], "twh": float(r[6])})
    return out


def fetch_ember() -> list[dict]:
    """Download both Ember files, save dated aggregates to data/raw/ember, then process."""
    import requests
    ex = requests.get(EMBER_EXPORTS, timeout=120); ex.raise_for_status()
    gen = requests.get(EMBER_GEN, timeout=300); gen.raise_for_status()
    agg, g = aggregate_ember_exports(ex.text), china_solar_generation(gen.text)
    if len(agg) < 60 or len(g) < 60:
        raise RuntimeError("Ember files look incomplete")
    d = RAW / "ember"; d.mkdir(parents=True, exist_ok=True)
    with open(d / f"solar_exports_agg_{TODAY}.csv", "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, ["period"] + AGG_COLS); w.writeheader(); w.writerows(agg)
    with open(d / f"china_solar_generation_{TODAY}.csv", "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, ["period", "twh"]); w.writeheader(); w.writerows(g)
    return process_ember()


def _latest(pattern):
    files = sorted((RAW / "ember").glob(pattern))
    return files[-1] if files else None


def process_ember() -> list[dict]:
    """Turn the latest saved Ember aggregates into tidy series + catalog entries."""
    fe, fg = _latest("solar_exports_agg_*.csv"), _latest("china_solar_generation_*.csv")
    if not fe or not fg:
        return []
    acc = fe.stem.split("_")[-1]
    ag = list(csv.DictReader(open(fe, encoding="utf-8")))
    num = lambda r, c: float(r[c])

    def write(sid, rows):
        with open(PROC / f"{sid}.csv", "w", newline="", encoding="utf-8") as fh:
            fh.write("period,value,flag\n")
            for per, v in rows:
                fh.write(f"{per},{v:.6f},\n")
        return sid + ".csv"

    note_ex = "Ember China Solar Exports (China customs), World, Panels + Cells, monthly"
    cat = []
    specs = [
        ("cn_pv_exports_kg", "China PV cell + module exports, net weight", "thousand tonnes", lambda r: num(r, "world_kg") / 1e6, "series"),
        ("cn_pv_exports_usd", "China PV cell + module exports, value", "USD billion", lambda r: num(r, "world_usd") / 1e9, "series"),
        ("cn_pv_export_unit_value", "China PV cell + module exports, unit value (value / weight)", "USD per kg", lambda r: num(r, "world_usd") / num(r, "world_kg"), "series"),
        ("cn_pv_exports_gw", "China PV cell + module exports, capacity (Ember estimate)", "GW", lambda r: num(r, "world_mw") / 1000, "series"),
        ("cn_pv_share_eu_us", "Share of China PV exports going to the EU and US", "% of weight", lambda r: (num(r, "eu_kg") + num(r, "us_kg")) / num(r, "world_kg") * 100, "market"),
        ("cn_pv_share_emerging", "Share of China PV exports going to the Middle East, Latin America and Africa", "% of weight", lambda r: (num(r, "me_kg") + num(r, "latam_kg") + num(r, "africa_kg")) / num(r, "world_kg") * 100, "market"),
        ("cn_pv_exports_eu_kt", "China PV exports to the EU, net weight", "thousand tonnes", lambda r: num(r, "eu_kg") / 1e6, "market"),
    ]
    for sid, label, unit, fn, role in specs:
        f = write(sid, [(r["period"], fn(r)) for r in ag])
        cat.append({"id": sid, "label": label, "unit": unit, "source_id": "S11", "frequency": "Monthly",
                    "file": f, "accessed": acc, "role": role, "note": note_ex})
    g = list(csv.DictReader(open(fg, encoding="utf-8")))
    f = write("cn_solar_generation_twh", [(r["period"], float(r["twh"])) for r in g])
    cat.append({"id": "cn_solar_generation_twh", "label": "China solar power generation", "unit": "TWh",
                "source_id": "S6", "frequency": "Monthly", "file": f, "accessed": fg.stem.split("_")[-1], "role": "series",
                "note": "Ember Monthly Electricity Data, China, Solar, generation (TWh)"})
    return cat


def fetch_eurostat_comext() -> list[dict]:
    """TODO: Eurostat Comext DS-045409, partner CN, CN8 codes under HS 854142 / 854143, imports, monthly."""
    return []


def fetch_nbs_solar_cells() -> list[dict]:
    """TODO: NBS monthly output of solar cells. If the site cannot be reached from CI, download by hand
    into data/raw/nbs/ and parse here. Jan-Feb are published combined: document how they are handled."""
    return []


def main():
    import sys
    PROC.mkdir(parents=True, exist_ok=True)
    if "--from-raw" in sys.argv:  # legacy: rebuild Comtrade series from a saved response
        f = pathlib.Path(sys.argv[sys.argv.index("--from-raw") + 1])
        cat = write_comtrade(json.loads(f.read_text(encoding="utf-8")), accessed=f.stem)
        (PROC / "catalog.json").write_text(json.dumps(cat, indent=1), encoding="utf-8")
        print(f"catalog from {f.name}: {len(cat)} series")
        return
    if "--offline" in sys.argv:  # process the latest saved Ember files without downloading
        catalog = process_ember()
    else:
        try:
            catalog = fetch_ember()
        except Exception as e:  # network blocked or file changed: keep the committed snapshot
            print(f"fetch_ember failed ({e}); processing the latest saved Ember files")
            catalog = process_ember()
    if catalog:
        (PROC / "catalog.json").write_text(json.dumps(catalog, indent=1), encoding="utf-8")
        print(f"catalog: {len(catalog)} series, accessed {catalog[0]['accessed']}")
    else:
        print("no series; existing data/processed snapshot is kept")


if __name__ == "__main__":
    main()
