"""Download public data into data/raw (dated) and write tidy series to data/processed.

STATUS: skeleton. Each fetcher below is a TODO to implement once the team has picked and
verified the sources (team.xlsx > Sources). Until catalog.json exists, build.py runs in demo mode.

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


def fetch_comtrade_exports() -> list[dict]:
    """TODO: UN Comtrade API, reporter China (156), flow X, HS 854142 and 854143, monthly, all partners.
    Needs COMTRADE_KEY (free tier). Save raw JSON, aggregate to world total + key partner groups
    (EU27, US, India, rest). Write volume (net weight or quantity) and value series."""
    key = os.environ.get("COMTRADE_KEY")
    if not key:
        print("comtrade: COMTRADE_KEY not set, skipped")
        return []
    raise NotImplementedError


def fetch_eurostat_comext() -> list[dict]:
    """TODO: Eurostat Comext DS-045409, partner CN, CN8 codes under HS 854142 / 854143, imports, monthly."""
    return []


def fetch_nbs_solar_cells() -> list[dict]:
    """TODO: NBS monthly output of solar cells. If the site cannot be reached from CI, download by hand
    into data/raw/nbs/ and parse here. Jan-Feb are published combined: document how they are handled."""
    return []


def main():
    PROC.mkdir(parents=True, exist_ok=True)
    catalog = []
    for fn in (fetch_comtrade_exports, fetch_eurostat_comext, fetch_nbs_solar_cells):
        catalog += fn()
    if catalog:
        for c in catalog:
            c.setdefault("accessed", TODAY)
        (PROC / "catalog.json").write_text(json.dumps(catalog, indent=1), encoding="utf-8")
        print(f"catalog: {len(catalog)} series")
    else:
        print("no series fetched; build.py will stay in demo mode")


if __name__ == "__main__":
    main()
