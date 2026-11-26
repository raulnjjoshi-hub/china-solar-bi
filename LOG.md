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

## Stage 3: Condensed structure and UI polish (24-25.11.2026)
- **Issue:** 9 tabs; forecast, scenarios and news on separate pages made the outlook hard to follow; generic styling.
- **Changes:** restyle (design tokens, bundled IBM Plex Sans, KPI row, shared chart theme, colour-blind-safe scenario
  colours), then cut to 5 tabs: Overview, Outlook (forecast + scenarios + news in one view), Strategy, Context, Evidence.
  Auto-written bottom line and act/prepare/watch summary on Overview.
- **Verification:** data.js byte-identical; every number shown on the old Market/Scenarios/Signals tabs found on the new
  Outlook/Evidence tabs (scripted comparison, also after moving a slider). Rendered at 1440/1280/768 px, no errors.
  Found and fixed a bug: "Candidate - not verified" sources were counted as verified (10 instead of 2-3).
- **Screenshots:** docs/screenshots/ui-before-*, ui-after-*, v2-*.

## Stage 4: Data to 2026 and forecast to 2027 (26.11.2026)
- **Issue:** the memo spoke about 2025, but the decision is about 2027. Comtrade monthly data for China ends 2024-12.
- **Changes:** switched the main source to Ember China Solar Exports (China customs, monthly to Oct 2026, by destination)
  and added Ember China solar generation as a domestic-demand proxy (the team's case sells in China). Forecast now runs
  to Dec 2027 (14 months); trend methods fitted on the last 48 months. New: domestic demand view, destination-share chart,
  2027 memo, signpost readings from data. Comtrade kept as an independent cross-check.
- **Verification:** browser vs file checksums identical. Ember vs Comtrade 2022-2024: weight +0.001%, value +0.003%,
  largest single month 0.06%; 2025 annual within 0.012%. Holdout: exports 12-month average 18.52% vs benchmark 20.90%
  (trend 30.59%, rejected); domestic trend + season 7.31% vs 17.01%. Rendered at 1440/1280/768 px, no errors.
- **Effect:** the memo and scenarios now speak about 2027 with data only one month old.

## Stage 5: User test and revision (planned 27.11)
