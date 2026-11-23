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

## Stage 2: Prototype with real data (planned 24-25.11)

## Stage 3: User test and revision (planned 26-27.11)
