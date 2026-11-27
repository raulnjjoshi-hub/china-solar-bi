# China Solar PV Export Board (Group 7)

Business Intelligence and the Global Business Environment, TSE, autumn 2026. Assignment 2.

A BI board for Company C, a privately owned Chinese solar panel manufacturer that sells only in China, to
Chinese companies and private individuals. Management must decide how to meet growing demand while preparing for
sudden change: how much capacity to run and whether to start exporting. Public customs and electricity data feed a
validated forecast to December 2027 and three scenarios (Base, Challenging, Optimistic); options are tested against them.

**Live board:** https://raulnjjoshi-hub.github.io/china-solar-bi/
**Backup:** `site/` folder in the submission zip; open `site/index.html` in any browser (works offline).

## How to open and use

1. Open the link (or `site/index.html`). No account needed.
2. **Overview:** key figures, the management memo for 2027, scenario outlook and what to do.
3. **Outlook:** forecast and scenarios in one chart. Pick a measure (exports, domestic demand, company revenue or
   profit), open a scenario to move its sliders, switch news items on and off.
4. **Strategy:** where China's exports go, option scores (start now / prepare, then trigger) and signposts.
5. **Context:** why the context matters, PESTLE and Porter's Diamond, each row with evidence and a source ID.
6. **Evidence:** data and holdout test, cross-check, sources, method, development log, AI use, limits.
7. Slider and score changes are kept only in your browser.

Tags on the board: `observed` = public data, `modelled` = forecast, `assumption` = team choice,
`proposal`/`draft` = not yet confirmed by the team.

## Live or snapshot

Dated snapshot. A GitHub Action (`.github/workflows/build.yml`) re-fetches the public data on the
3rd of each month, rebuilds, commits the snapshot and redeploys. It also rebuilds on every push.
The snapshot date is shown top right of the board.

## How it is built

```
content/team.xlsx        all team-written content (sources, PESTLE, scenarios, options, log...)
scripts/fetch.py         downloads public data -> data/raw (dated) and data/processed (tidy CSV)
scripts/build.py         reads team.xlsx + data/processed, computes forecasts + holdout -> site/data.js
site/index.html          the board (plain HTML/JS, no external libraries)
scripts/make_template.py regenerates an empty team workbook (structure reference)
```

Run locally:

```bash
pip install -r requirements.txt
python scripts/fetch.py      # downloads Ember data; --offline reprocesses the saved files in data/raw/ember
python scripts/build.py      # --demo forces synthetic placeholder data
# open site/index.html
```

## Refresh procedure

- **Data:** automatic monthly, or Actions tab > "Refresh data and deploy board" > Run workflow.
- **Team content:** edit `content/team.xlsx` (the team's narrative is in `docs/Business_Intelligence_final.docx`),
  push. The board rebuilds.

## Data sources

Main data (all monthly, dated snapshot, raw files in `data/raw/ember/`):
- **S11** Ember, China Solar Exports (from China customs): panels and cells, weight, value, destination, 2017-01 to 2026-10.
- **S6** Ember, monthly electricity data: China solar generation (TWh), 2016-01 to 2026-10. Proxy for domestic demand.
- **S1** UN Comtrade, HS 854142 + 854143, China exports: independent cross-check (Ember matches within 0.003%).

The **Sources** table on the Evidence tab (generated from `team.xlsx > Sources`) lists all 20 sources with table-level
links, variables, units, coverage, access dates and transformations.

## Tools used

Claude (Anthropic) for planning, code, data retrieval, source search and editing; [other LLMs used by the team]. Python 3.12, openpyxl,
GitHub Actions, GitHub Pages. See the AI use sheet / Method and log tab for what was accepted,
revised or rejected.

## Limitations

Company figures are illustrative. Signal effects and option scores are team judgement. Forecast
ranges are indicative (±1.28 × holdout RMSE), not statistical intervals. Ember republishes customs data, so recent months may be revised. Domestic demand is judged from solar generation,
which follows the installed stock and likely overstates panel demand in 2026, when new installations fell. Export GW
is Ember's estimate. Full list on the Evidence tab.
