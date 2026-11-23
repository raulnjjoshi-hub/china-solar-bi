# CLAUDE.md: China Solar PV Export Board (Group 7)

Team assignment, TSE course "Business Intelligence and the Global Business Environment", autumn 2026.
**Due Sun 29.11.2026.** Raul is tech lead and the only developer; three teammates write all content
(sources, forecast check in Excel, PESTLE, Diamond, scenarios, options, signals, log) in `content/team.xlsx`.

## What the grader rewards (weights)
Iteration and reflection 20 · data traceability 15 · forecast method and validation 15 · scenarios 15 ·
frameworks 10 · strategic usefulness 10 · functionality and communication 10 · teamwork 5.
Polish alone earns little. Every displayed number must trace to a source table. Observed, modelled and
assumed must be visibly distinct. LLM-generated numbers are never evidence.

## Decisions made (23.11)
- Country/industry assigned: China, renewable energy. Focus: solar PV manufacturing (cells, modules).
- User: privately owned, export-oriented Chinese PV manufacturer. Decision: which export markets to
  prioritise and how much capacity to run, next 12-24 months.
- Forecast target (team confirms in Forecast target sheet): proposal F1 = China PV exports, HS 854142
  (cells) and 854143 (modules), UN Comtrade, monthly, volume with value alongside. HS 2022 split means
  clean series from 2022. Second series: NBS monthly solar cell output.
- Horizons: forecast 12 months; scenarios to 2030 (15th Five-Year Plan).
- Scenario axes: U1 Western market access (Open/Closed) × U2 China supply discipline (Strong/Weak).
  A Price war behind walls (Closed/Weak), B Managed consolidation (Closed/Strong), C Open-door volume race (Open/Weak).
- Architecture: static page + GitHub Action (monthly cron + on push) → GitHub Pages. Zip backup for submission.
- Team edits Excel on UTU SharePoint (no API access); Raul downloads it to content/team.xlsx and pushes.

## Workbook contract (content/team.xlsx)
- Content sheets: row 1 title, row 2 instruction, **row 3 headers, data from row 4**.
- Skip rows whose first cell starts with `EX` (examples) and rows with Status `Rejected`.
- Header strings are the keys. `build.py > read_content()` maps them; if a header changes, update both
  `scripts/make_template.py` and `build.py`.
- `Forecast` sheet is the team's independent Excel forecast (not parsed). Its methods 1-3 (seasonal naive,
  12-mo MA, linear trend) must match build.py exactly so MAPE values can be compared.
- `Forecast target` decision is in the cell right of "Team decision" below the table.

## Repo
```
content/team.xlsx  scripts/{fetch.py,build.py,make_template.py}  data/{raw,processed}
site/{index.html,data.js}  .github/workflows/build.yml  README.md  LOG.md  docs/screenshots/
```
- `build.py`: demo mode (synthetic series, red banner) until `data/processed/catalog.json` exists.
- Forecast methods: snaive (benchmark), ma12, trend, trend_season. Holdout 12. Default = lowest MAPE that
  beats snaive. Band = ±1.2816 × holdout RMSE × sqrt(step/H), labelled indicative.
- Scenario engine (index.html `runScen`): volume drift d, price drift p, margin change m, company inputs
  m0/share/beta; signals add to d/p/m per scenario. Profit shown per 100 of today's revenue (avoids index
  blow-up with thin margins).
- `index.html` is dependency-free; works from file:// because data is a JS file, not fetched JSON.

## Next steps (in order)
1. Push repo to GitHub, enable Pages (source: GitHub Actions), add `COMTRADE_KEY` secret.
2. Implement `fetch_comtrade_exports` (world total + EU27/US/India/rest) → catalog.json. Verify a few
   values by hand against the Comtrade UI; team logs them in Data checks.
3. Wire signposts to series (add a `series`/`calc` column or a mapping in build.py) so readings and
   tripped state show.
4. Add second series (NBS or Eurostat) once the team verifies the source.
5. Screenshot each version into docs/screenshots/ (v1 prototype, v2 after user test) for the log.
6. Thursday user test → fixes → retest. Keep LOG.md in sync with the Log sheet.

## Conventions
- No external CDNs or fonts (offline backup must work). Plain HTML/CSS/JS.
- Keep facts and judgement separate in the UI. Show source IDs next to evidence.
- Commit messages describe the analytical or usability change; they double as log evidence.
