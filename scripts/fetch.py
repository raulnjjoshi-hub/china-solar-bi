"""Download public data into data/raw (dated) and write tidy series to data/processed.

STATUS: Comtrade implemented (preview endpoint, no key needed). Eurostat and NBS are TODO.
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

import datetime as dt
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
    if "--from-raw" in sys.argv:
        f = pathlib.Path(sys.argv[sys.argv.index("--from-raw") + 1])
        cat = write_comtrade(json.loads(f.read_text(encoding="utf-8")), accessed=f.stem)
        (PROC / "catalog.json").write_text(json.dumps(cat, indent=1), encoding="utf-8")
        print(f"catalog from {f.name}: {len(cat)} series")
        return
    catalog = []
    for fn in (fetch_comtrade_exports, fetch_eurostat_comext, fetch_nbs_solar_cells):
        try:
            catalog += fn()
        except Exception as e:  # network blocked, API change: keep the committed snapshot
            print(f"{fn.__name__} failed ({e}); keeping existing snapshot")
    if catalog:
        for c in catalog:
            c.setdefault("accessed", TODAY)
        (PROC / "catalog.json").write_text(json.dumps(catalog, indent=1), encoding="utf-8")
        print(f"catalog: {len(catalog)} series")
    else:
        print("no series fetched; existing data/processed snapshot is kept")


if __name__ == "__main__":
    main()
