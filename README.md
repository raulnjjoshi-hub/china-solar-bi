# China Solar PV Export Board (Group 7)

Business Intelligence and the Global Business Environment, TSE, autumn 2026. Assignment 2.

A BI board for a privately owned, export-oriented Chinese PV manufacturer deciding which export
markets to prioritise and how much capacity to run. Public trade and production data plus news feed a
validated forecast and three scenarios; options are tested against them.

**Live board:** https://raulnjjoshi-hub.github.io/china-solar-bi/
**Backup:** `site/` folder in the submission zip; open `site/index.html` in any browser (works offline).

## How to open and use

1. Open the link (or `site/index.html`). No account needed.
2. Start on **Overview**, then **Market and forecast**, **Scenarios** (move the sliders),
   **Options and signposts**, **Signals**.
3. Your changes are kept only in your browser. **README** tab has a reset button.

Chips on the board: `observed` = public data, `modelled` = forecast, `assumption` = team choice,
`proposal` = not yet reviewed by the team.

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
python scripts/fetch.py      # needs COMTRADE_KEY env var for Comtrade
python scripts/build.py      # --demo forces synthetic placeholder data
# open site/index.html
```

## Refresh procedure

- **Data:** automatic monthly, or Actions tab > "Refresh data and deploy board" > Run workflow.
- **Team content:** edit the shared Excel, download it to `content/team.xlsx`, push. The board rebuilds.

## Data sources

See the **Sources** tab (generated from `team.xlsx > Sources`): dataset, table code, variables,
units, coverage, access date, missing values, transformations.

## Tools used

Claude (Anthropic) for planning and code; [other LLMs used by the team]. Python 3.12, openpyxl,
GitHub Actions, GitHub Pages. See the AI use sheet / Method and log tab for what was accepted,
revised or rejected.

## Limitations

Company figures are illustrative. Signal effects and option scores are team judgement. Forecast
ranges are indicative (±1.28 × holdout RMSE), not statistical intervals. [Add data limitations.]
