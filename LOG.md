# Development log (technical side)

Team-level log lives in `content/team.xlsx > Log`. This file records build changes with links to commits
and screenshots so the two can be cross-referenced.

## Stage 1: Brief and design (23.11.2026)
- **Issue:** needed a user, decision and forecastable indicator for "China, renewable energy".
- **Changes:** focused on PV manufacturing; user = private Chinese PV manufacturer; five candidate targets
  listed with pros and cons; static board + scheduled GitHub Action; team workbook as content source.
- **Verification:** workbook Forecast sheet tested against an independent numpy calculation on synthetic
  data (all methods matched). Board forecasts checked the same way (trend + seasonal and seasonal naive MAPE matched to 4 dp).
- **Found and fixed:** Options sheet proposal text was shifted one column (formula column skipped in the
  template generator). Fixed before team use.
- **Screenshots:** docs/screenshots/v0-frames-*.png (demo data).

## Stage 2: Prototype with real data (24.11.2026)
- **Issue:** shells (cloud, local VM) blocked from Comtrade; free preview API allows one period per call.
- **Changes:** captured China monthly exports (HS 854142+854143, partner World) via the desktop browser,
  saved raw JSON to data/raw/comtrade/2026-11-24.json, processed with `fetch.py --from-raw`. Three series:
  net weight, value, unit value. Team scenarios (base/challenging/optimistic) and three starter signals added.
- **Finding:** Comtrade has China monthly data only to 2024-12 (2025 annual total exists, 2025 months empty).
- **Verification:** browser vs file checksums identical (72 rows; sums of value, weight, qty match).
  Monthly 2024 sum = 2024 annual total (0.00% difference). Out-of-sample: 2025 forecast vs 2025 annual actual:
  volume trend -0.1%, seasonal naive -6.2%; value trend -20.7% (price fall slowed in 2025), naive +8.5%.
- **Effect:** board runs on real, traceable data; the forecast is validated twice (holdout + later actual).
- **Remaining:** series ends 2024-12, ~2 years before the snapshot. Options: mirror data (partners' imports
  from China, available into 2025) or China Customs monthly releases.

## Stage 3: User test and revision (planned 26-27.11)
