"""Build Group7 China Solar PV team workbook (content contract for the dashboard build)."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1] if len(sys.argv) > 1 else "Group7_China_SolarPV_Team.xlsx"
TEST = len(sys.argv) > 2 and sys.argv[2] == "test"

F = "Arial"
f_title = Font(name=F, size=14, bold=True)
f_instr = Font(name=F, size=10, italic=True, color="404040")
f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
f_body = Font(name=F, size=10)
f_prop = Font(name=F, size=10, color="0000FF")          # proposal from tech lead: review
f_ex = Font(name=F, size=10, italic=True, color="808080")  # example row
f_bold = Font(name=F, size=10, bold=True)
fill_head = PatternFill("solid", fgColor="1F4E5F")
fill_in = PatternFill("solid", fgColor="FFF9DB")     # team fills in
fill_calc = PatternFill("solid", fgColor="E7E9EC")   # formula, do not edit
fill_ex = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="C9CFD4")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")

wb = Workbook()
wb.remove(wb.active)

# ---------------- lists (hidden) ----------------
LISTS = {
    "Status": ["Proposal", "Draft", "Verified", "Rejected"],
    "PESTLE": ["Political", "Economic", "Social", "Technological", "Legal", "Environmental"],
    "Driver": ["Export volume", "Export price", "Domestic demand", "Western market access",
               "Supply discipline", "Cost / margin", "Technology"],
    "Direction": ["Helps manufacturer", "Hurts manufacturer", "Mixed", "Unknown"],
    "SrcStatus": ["Candidate - not verified", "Verified", "Used in board", "Context only", "Rejected"],
    "Kind": ["Reported news", "Hypothetical test"],
    "ReadStatus": ["Read", "Queued - not read"],
    "Score": [-2, -1, 0, 1, 2],
    "AboveBelow": ["above", "below"],
    "YN": ["Yes", "No", "Partly"],
    "Severity": ["High", "Medium", "Low"],
    "AIOutcome": ["Accepted", "Revised", "Rejected"],
    "ForecastOpt": ["F1", "F2", "F3", "F4", "F5", "Other"],
    "LowHigh": ["Open", "Closed", "Strong", "Weak"],
}
ls = wb.create_sheet("Lists")
for j, (k, vals) in enumerate(LISTS.items(), start=1):
    ls.cell(1, j, k).font = f_bold
    for i, v in enumerate(vals, start=2):
        ls.cell(i, j, v).font = f_body
LREF = {k: f"Lists!${L(j)}$2:${L(j)}${len(v)+1}" for j, (k, v) in enumerate(LISTS.items(), start=1)}


def dv(ws, listname, rng):
    d = DataValidation(type="list", formula1=f"={LREF[listname]}", allow_blank=True)
    d.add(rng)
    ws.add_data_validation(d)


def table(name, title, instr, cols, rows=(), n=40, owner=""):
    """cols: list of dicts {h, w, list?, f? (formula template with {r}), note?}
    Layout contract: row1 title, row2 instruction, row3 headers, data from row4."""
    ws = wb.create_sheet(name)
    ws["A1"] = title + (f"   |   Owner: {owner}" if owner else "")
    ws["A1"].font = f_title
    ws["A2"] = instr
    ws["A2"].font = f_instr
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(len(cols), 6))
    ws.row_dimensions[2].height = 15 * max(2, len(instr) // 150 + 1)
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(3, j, c["h"])
        cell.font, cell.fill, cell.border = f_head, fill_head, box
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[L(j)].width = c.get("w", 18)
        if c.get("note"):
            cell.comment = Comment(c["note"], "Tech lead")
    ws.row_dimensions[3].height = 32
    last = 3 + max(n, len(rows))
    for i in range(4, last + 1):
        data = rows[i - 4] if i - 4 < len(rows) else None
        kind = data[0] if data else None  # 'ex' | 'prop' | 'fixed'
        vals = data[1] if data else []
        for j, c in enumerate(cols, start=1):
            cell = ws.cell(i, j)
            cell.border = box
            cell.alignment = wrap
            if "f" in c:
                cell.value = c["f"].format(r=i)
                cell.fill, cell.font = fill_calc, f_body
                if c.get("fmt"):
                    cell.number_format = c["fmt"]
                continue
            if j - 1 < len(vals) and vals[j - 1] is not None:
                cell.value = vals[j - 1]
            if kind == "ex":
                cell.font, cell.fill = f_ex, fill_ex
            elif kind == "prop":
                cell.font, cell.fill = f_prop, fill_in
            elif kind == "fixed":
                cell.font, cell.fill = f_bold if j == 1 else f_body, fill_in
            else:
                cell.font, cell.fill = f_body, fill_in
            if c.get("fmt"):
                cell.number_format = c["fmt"]
    for j, c in enumerate(cols, start=1):
        if c.get("list"):
            dv(ws, c["list"], f"{L(j)}4:{L(j)}{last}")
    ws.freeze_panes = "B4"
    return ws


# ---------------- Start here ----------------
st = wb.create_sheet("Start here", 0)
st.column_dimensions["A"].width = 30
st.column_dimensions["B"].width = 95
r = 1
def put(a, b="", fa=f_body, fb=f_body, fill=None):
    global r
    st.cell(r, 1, a).font = fa
    st.cell(r, 2, b).font = fb
    st.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    st.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        st.cell(r, 1).fill = fill
        st.cell(r, 2).fill = fill
    r += 1
put("Group 7 | China | Renewable energy: solar PV manufacturing", fa=f_title)
put("Team workbook for the BI dashboard (Assignment 2, due 29.11.2026)", fa=f_instr)
r += 1
put("What this file is", fa=f_bold)
put("", "Everything the team writes goes here. The dashboard reads its text, tables and assumptions straight from this file, so nobody else needs to touch code. When a sheet is ready, tell Raul; he copies this file into the project and the board rebuilds.")
r += 1
put("Rules (the build depends on them)", fa=f_bold)
for a, b in [
    ("1", "Do not rename sheets or column headers, and do not insert rows above row 4. Add your rows below the existing ones."),
    ("2", "Give every row a unique ID (S1, S2... for sources, P1... for PESTLE, and so on). Refer to sources by their ID from the Sources sheet."),
    ("3", "Facts only from pages or tables you actually opened. Put the fact in the fact column and your interpretation in the judgement column. Numbers from an AI chat are not evidence until checked at the source."),
    ("4", "Set Status on every row. Proposal = idea from the tech lead, review it. Draft = written, not checked. Verified = checked against the source by a second person. Rejected = kept for the log, hidden from the board."),
    ("5", "Rows whose ID starts with EX are format examples. Leave them; the board ignores them."),
    ("6", "No personal data (names of testers, emails) in this file. Use roles, e.g. 'Tester 1, logistics student'."),
]:
    put(a, b)
r += 1
put("Colour legend", fa=f_bold)
put("Light yellow cell", "You fill this in.", fill=fill_in)
put("Blue text", "Proposal from the tech lead. Review, change or reject, then change Status.", fb=f_prop)
put("Grey cell", "Formula. Do not type here.", fill=fill_calc)
put("Grey italic", "Example row (ID starts with EX).", fb=f_ex)
r += 1
put("Who owns what", fa=f_bold)
for a, b in [
    ("Raul (tech lead)", "Dashboard build, data pipeline, Brief (with all), final checks."),
    ("Member A: data and forecast", "Forecast target, Sources, Data checks, Forecast."),
    ("Member B: frameworks and news", "PESTLE, Diamond, Signals."),
    ("Member C: scenarios and strategy", "Uncertainties, Scenarios, Company inputs, Signposts, Options, User test."),
    ("Everyone", "Log, AI use, Contributions. Write your own entries as you go."),
]:
    put(a, b)
r += 1
put("Timeline", fa=f_bold)
for a, b in [
    ("Mon 23.11", "Kick-off. Agree Brief and forecast target. Start Sources."),
    ("Tue 24.11", "Sources verified, data pulled, PESTLE and Diamond drafted."),
    ("Wed 25.11", "Forecast in Excel, scenarios and options drafted. Prototype v1 reviewed together."),
    ("Thu 26.11", "User test with someone outside the team (User test sheet)."),
    ("Fri 27.11", "Fixes, second test, reflection draft."),
    ("Sat 28.11", "README, log, reflection final."),
    ("Sun 29.11", "Final checks and submission: Group7_China_RenewableEnergy."),
]:
    put(a, b)
r += 1
put("Sheets", fa=f_bold)
for a, b in [
    ("Brief", "User, decision, horizons. Shown on the board's first tab."),
    ("Forecast target", "Options for the indicator we forecast. Team picks one."),
    ("Sources", "Every dataset and article, at table level, with access date, units, coverage, gaps."),
    ("Data checks", "Board values checked against the source by hand."),
    ("Forecast", "The team's own forecast and holdout test in Excel. Used to check the board's Python numbers."),
    ("PESTLE / Diamond", "Framework rows linked to indicators, sources and model drivers."),
    ("Uncertainties / Scenarios / Company inputs", "The scenario logic and the slider defaults."),
    ("Signals", "News items: fact, judgement and the effect on each scenario."),
    ("Signposts", "Indicators with thresholds that tell us which future is arriving."),
    ("Options", "Strategic moves scored in each scenario. Verdict is calculated."),
    ("Log / User test / AI use / Contributions", "Evidence for iteration, testing, AI judgement and teamwork."),
]:
    put(a, b)

# ---------------- Brief ----------------
table("Brief", "Brief", "One value per row. The board's first tab is written from this sheet. Proposals are in blue: agree or change them at the kick-off.",
      [dict(h="Field", w=28), dict(h="Value", w=70), dict(h="Status", w=12, list="Status"), dict(h="Notes / evidence needed", w=55)],
      [("fixed", ["Country", "China", "Verified", "Assigned"]),
       ("fixed", ["Industry", "Renewable energy, focus on solar PV manufacturing (cells and modules)", "Draft", "Assigned: renewable energy. Focus chosen by the team."]),
       ("fixed", ["Group", "Group 7", "Verified", ""]),
       ("prop", ["Intended user", "A privately owned, export-oriented Chinese PV cell and module manufacturer (strategy and sales leadership)", "Proposal", "Why this user: the leading producers are mostly private and export a large share of output. Name 2-3 real examples with a source."]),
       ("prop", ["Decision", "Which export markets to prioritise, and how much capacity to run, over the next 12-24 months", "Proposal", ""]),
       ("prop", ["Key question", "Where will demand for Chinese PV exports come from, at what prices, and which moves are safe under every plausible future?", "Proposal", ""]),
       ("prop", ["Forecast horizon", "12 months ahead, monthly data", "Proposal", "Justify: export data is monthly and shifts quickly with prices and trade rules; longer point forecasts are not credible."]),
       ("prop", ["Scenario horizon", "To 2030", "Proposal", "Justify: factory build and capacity cycles take a few years; matches China's 15th Five-Year Plan (2026-2030)."]),
       ("prop", ["Why this context matters", "", "Proposal", "Write 3-4 sentences with sources, e.g. China's share of each stage of the global PV supply chain (IEA), share of production exported, recent price falls."]),
       ("prop", ["Data mode", "Dated snapshot, refreshed monthly by a scheduled GitHub Action; the date is shown on the board", "Proposal", "Tech lead handles this."]),
       ("fixed", ["Submission name", "Group7_China_RenewableEnergy", "Verified", "Handout format: GroupX_Country_Industry"]),
       ("fixed", ["Deadline", "29.11.2026", "Verified", ""]),
       ], n=16, owner="All, led by Raul")

# ---------------- Forecast target ----------------
ft = table("Forecast target", "Forecast target options",
      "Pick the indicator we forecast. It must be monthly (so a 12-month holdout test is possible), public and relevant to the user's decision. Record the choice and reason in the two yellow cells below the table.",
      [dict(h="ID", w=6), dict(h="Indicator", w=34), dict(h="Source and dataset", w=34), dict(h="Frequency / unit", w=20),
       dict(h="Why it fits the user", w=40), dict(h="Problems to handle", w=48), dict(h="Tech lead view", w=26)],
      [("prop", ["F1", "China exports of PV cells and modules, volume (and value)", "UN Comtrade, HS 854142 (cells) and 854143 (modules), reporter China, by partner", "Monthly; kg or units, USD", "Directly matches the manufacturer's export decision; can be split by destination market", "HS codes split in 2022 (before that, code 854140 also included LEDs), so the clean series starts 2022. Value fell with prices: show volume and value together.", "Recommended main target"]),
       ("prop", ["F2", "China output of solar cells", "National Bureau of Statistics (NBS), industrial output of major products", "Monthly; 10,000 kW", "Matches the production focus and the capacity decision", "January and February are reported together (Spring Festival). Site can be unreliable from outside China.", "Good second series"]),
       ("prop", ["F3", "EU imports of PV cells and modules from China", "Eurostat Comext, DS-045409, CN codes for 854142 and 854143, partner CN", "Monthly; EUR, kg", "Shows demand in one key export market; cleanest data", "One market only", "Easiest to refresh"]),
       ("prop", ["F4", "China new PV installations", "National Energy Administration (NEA) monthly releases", "Monthly; GW", "Home-market demand absorbing domestic output", "Very volatile from policy-driven install rushes; published in press releases, not a database", "Context, not target"]),
       ("prop", ["F5", "China installed PV capacity", "IRENA Renewable Capacity Statistics", "Annual; MW", "Long-run demand trend", "Annual only: too few points for a good holdout test", "Context, not target"]),
       ], n=6, owner="Member A")
rr = 3 + 6 + 2
ft.cell(rr, 1, "Team decision").font = f_bold
ft.cell(rr, 2).fill = fill_in
ft.cell(rr, 2).border = box
dv(ft, "ForecastOpt", f"B{rr}")
ft.cell(rr + 1, 1, "Reason").font = f_bold
ft.merge_cells(start_row=rr + 1, start_column=2, end_row=rr + 1, end_column=6)
ft.cell(rr + 1, 2).fill = fill_in
ft.cell(rr + 1, 2).alignment = wrap
ft.row_dimensions[rr + 1].height = 45

# ---------------- Sources ----------------
table("Sources", "Sources (dataset level)",
      "One row per dataset table or article. The handout requires: dataset name, link, variables, units, coverage, access date, missing values and transformations. Use the exact table, not the institution homepage. Candidate rows are ideas: open them, fill the gaps and set Status.",
      [dict(h="ID", w=6), dict(h="Name (exact table or article title)", w=38), dict(h="Provider / publisher", w=18),
       dict(h="Dataset code / table", w=20), dict(h="URL (table-level link)", w=38), dict(h="Variables and filters", w=34),
       dict(h="Unit", w=14), dict(h="Frequency", w=11), dict(h="Coverage (from-to)", w=18), dict(h="Access date", w=12, fmt="yyyy-mm-dd"),
       dict(h="Missing values and flags", w=30), dict(h="Transformations applied", w=28), dict(h="Used for", w=26),
       dict(h="Status", w=20, list="SrcStatus"), dict(h="Checked by", w=12)],
      [("ex", ["EX-1", "GDP growth (annual %)", "World Bank", "WDI, NY.GDP.MKTP.KD.ZG", "https://databank.worldbank.org/source/world-development-indicators", "Country = China", "% per year", "Annual", "2000-2025", "2026-11-24", "None for China", "None", "PESTLE: Economic", "Verified", "Member A"]),
       ("prop", ["S1", "China exports of PV cells and modules by partner", "UN Comtrade", "HS 854142, 854143", "https://comtradeplus.un.org/ (add query link)", "Reporter China, flow export, all partners", "USD; kg / units", "Monthly", "", None, "Series break at HS 2022 change", "", "Forecast target (if F1)", "Candidate - not verified", ""]),
       ("prop", ["S2", "China output of solar cells", "National Bureau of Statistics of China", "Industrial output of major products", "https://data.stats.gov.cn/english/ (add table link)", "Solar cells, current month and cumulative", "10,000 kW", "Monthly", "", None, "Jan-Feb reported combined", "", "Second series / production context", "Candidate - not verified", ""]),
       ("prop", ["S3", "EU trade by CN8, imports from China", "Eurostat (Comext)", "DS-045409", "(add table link)", "Partner CN, CN8 for HS 854142 and 854143, imports", "EUR; kg", "Monthly", "", None, "", "", "EU market exposure", "Candidate - not verified", ""]),
       ("prop", ["S4", "Monthly new PV installations", "National Energy Administration (NEA)", "Press releases", "https://www.nea.gov.cn/ (add release links)", "New solar PV capacity added", "GW", "Monthly", "", None, "Not a database; collect by hand", "", "Domestic demand context", "Candidate - not verified", ""]),
       ("prop", ["S5", "Renewable Capacity Statistics", "IRENA", "IRENASTAT", "https://www.irena.org/Data (add table link)", "China, solar PV, installed capacity", "MW", "Annual", "", None, "", "", "Long-run demand context", "Candidate - not verified", ""]),
       ("prop", ["S6", "Monthly electricity data", "Ember", "", "https://ember-energy.org/data/ (add dataset link)", "China, solar generation", "TWh", "Monthly", "", None, "", "", "Demand cross-check", "Candidate - not verified", ""]),
       ("prop", ["S7", "World Development Indicators", "World Bank", "WDI (pick indicators)", "https://databank.worldbank.org/source/world-development-indicators", "China: GDP, electricity, exports", "various", "Annual", "", None, "", "", "PESTLE context", "Candidate - not verified", ""]),
       ("prop", ["S8", "Solar PV Global Supply Chains (special report)", "IEA", "", "(add link)", "China's share of each supply-chain stage", "%", "One-off", "", None, "", "", "Why this context matters", "Candidate - not verified", ""]),
       ], n=40, owner="Member A")

# ---------------- Data checks ----------------
dc = table("Data checks", "Data checks (board value vs source)",
      "Pick values shown on the board and check them by hand at the source. Record at least 10 checks across sources. Mismatches go into the Log. Tolerance for a match is set in cell K1.",
      [dict(h="ID", w=6), dict(h="Indicator", w=34), dict(h="Period", w=11), dict(h="Value on board", w=14),
       dict(h="Value at source", w=14), dict(h="Match?", w=11, f='=IF(OR(D{r}="",E{r}=""),"",IF(ABS(D{r}-E{r})<=$K$1,"OK","MISMATCH"))'),
       dict(h="Source ID", w=10), dict(h="Checked by", w=12), dict(h="Date", w=12, fmt="yyyy-mm-dd"), dict(h="Note", w=40)],
      [("ex", ["EX-1", "China PV module exports, value", "2026-03", 1234.5, 1234.5, None, "S1", "Member A", "2026-11-25", "Example only"])],
      n=30, owner="Member A")
dc["J1"] = "Tolerance"
dc["J1"].font = f_bold
dc["K1"] = 0.05
dc["K1"].font, dc["K1"].fill = f_prop, fill_in

# ---------------- Forecast ----------------
fs = wb.create_sheet("Forecast")
fs["A1"] = "Forecast and holdout test (team's independent check)   |   Owner: Member A"
fs["A1"].font = f_title
fs["A2"] = ("Paste the monthly series in column C from row 14, oldest first, with no gaps (fix missing months before pasting and note how in Sources). "
            "Holdout months = 12 hides the last 12 months, forecasts them from the earlier data and scores each method. "
            "Set holdout to 0 to get the final forecast from all data. Then type the board's MAPE values into column H to check that Excel and the board agree.")
fs["A2"].font = f_instr
fs["A2"].alignment = Alignment(wrap_text=True, vertical="top")
fs.merge_cells("A2:N2")
fs.row_dimensions[2].height = 48
labels = [("Series name", None), ("Unit", None), ("Source ID", None), ("First month (date)", None),
          ("Holdout months (0-12)", 12), ("Months to forecast", 12)]
for i, (lab, v) in enumerate(labels, start=4):
    fs.cell(i, 1, lab).font = f_bold
    c = fs.cell(i, 2, v)
    c.fill, c.border = fill_in, box
    c.font = f_prop if v is not None else f_body
fs["B7"].number_format = "yyyy-mm"
hv = DataValidation(type="whole", operator="between", formula1="0", formula2="12")
hv.add("B8"); fs.add_data_validation(hv)
mv = DataValidation(type="whole", operator="between", formula1="1", formula2="24")
mv.add("B9"); fs.add_data_validation(mv)
FR0, FR1 = 14, 14 + 143  # 144 months
C = f"$C${FR0}:$C${FR1}"
calc = [("Data points entered", f"=COUNT({C})", "0"),
        ("Training ends at month n", '=IF(B10<24,"",B10-B8)', "0"),
        ("12-month average at end of training", f'=IF(B11="","",AVERAGE(INDEX({C},B11-11):INDEX({C},B11)))', "#,##0.00")]
for i, (lab, f, fmt) in enumerate(calc, start=10):
    fs.cell(i, 1, lab).font = f_bold
    c = fs.cell(i, 2, f)
    c.fill, c.border, c.font, c.number_format = fill_calc, box, f_body, fmt
fs.column_dimensions["A"].width = 34
fs.column_dimensions["B"].width = 14

# results block
res_h = ["Method", "Holdout MAPE %", "Holdout RMSE", "Board MAPE % (type in)", "Excel = board?"]
for j, h in enumerate(res_h, start=4):
    c = fs.cell(3, j, h)
    c.font, c.fill, c.border = f_head, fill_head, box
    c.alignment = Alignment(wrap_text=True, vertical="center")
methods = [("Seasonal naive (benchmark)", "K", "O"), ("12-month moving average", "L", "P"),
           ("Linear trend", "M", "Q"), ("Own method (column H)", "N", "R")]
for i, (m, ape, se) in enumerate(methods, start=4):
    fs.cell(i, 4, m).font = f_body
    a = f"${ape}${FR0}:${ape}${FR1}"
    s = f"${se}${FR0}:${se}${FR1}"
    fs.cell(i, 5, f'=IF(COUNT({a})=0,"",AVERAGE({a})*100)').number_format = "0.00"
    fs.cell(i, 6, f'=IF(COUNT({s})=0,"",SQRT(AVERAGE({s})))').number_format = "#,##0.00"
    fs.cell(i, 7).fill = fill_in
    fs.cell(i, 7).number_format = "0.00"
    fs.cell(i, 8, f'=IF(OR(E{i}="",G{i}=""),"",IF(ABS(E{i}-G{i})<=0.05,"OK","CHECK"))')
    for j in (4, 5, 6, 7, 8):
        fs.cell(i, j).border = box
        if j in (5, 6, 8):
            fs.cell(i, j).fill, fs.cell(i, j).font = fill_calc, f_body
fs.cell(8, 4, "Lowest MAPE wins only if it beats the seasonal naive benchmark. Agreement tolerance: 0.05 percentage points.").font = f_instr
for col, w in zip("DEFGH", (30, 14, 14, 16, 13)):
    fs.column_dimensions[col].width = w

heads = ["n", "Month", "Actual (paste)", "Set", "Seasonal naive", "12-mo moving avg", "Linear trend", "Own method (optional)",
         "Train x", "Train y", "APE seasonal naive", "APE moving avg", "APE trend", "APE own",
         "Sq err seasonal naive", "Sq err moving avg", "Sq err trend", "Sq err own"]
for j, h in enumerate(heads, start=1):
    c = fs.cell(13, j, h)
    c.font, c.fill, c.border = f_head, fill_head, box
    c.alignment = Alignment(wrap_text=True, vertical="center")
fs.row_dimensions[13].height = 32
for j in range(9, 19):
    fs.column_dimensions[L(j)].width = 11
fs.column_dimensions["C"].width = 14
fs.column_dimensions["H"].width = 13

for k in range(144):
    r_ = FR0 + k
    n = k + 1
    row = {
        1: n,
        2: f'=IF($B$7="","",EDATE($B$7,A{r_}-1))',
        4: f'=IF($B$11="","",IF(A{r_}<=$B$11,"Train",IF(A{r_}<=$B$10,"Holdout",IF(A{r_}<=$B$10+$B$9,"Forecast",""))))',
        5: (f'=IF(OR(D{r_}="",D{r_}="Train"),"",IF(D{r_-12}="Train",C{r_-12},E{r_-12}))' if n > 12 else None),
        6: f'=IF(OR(D{r_}="",D{r_}="Train"),"",$B$12)',
        7: f'=IF(OR(D{r_}="",D{r_}="Train"),"",FORECAST(A{r_},$J${FR0}:$J${FR1},$I${FR0}:$I${FR1}))',
        9: f'=IF(D{r_}="Train",A{r_},"")',
        10: f'=IF(D{r_}="Train",C{r_},"")',
    }
    for jj, (pc, apec, sec) in enumerate([("E", "K", "O"), ("F", "L", "P"), ("G", "M", "Q"), ("H", "N", "R")]):
        cond = f'AND(D{r_}="Holdout",ISNUMBER({pc}{r_}),ISNUMBER(C{r_}))'
        row[11 + jj] = f'=IF({cond},IF(C{r_}=0,"",ABS({pc}{r_}-C{r_})/ABS(C{r_})),"")'
        row[15 + jj] = f'=IF({cond},({pc}{r_}-C{r_})^2,"")'
    for j in range(1, 19):
        c = fs.cell(r_, j)
        c.border = box
        c.font = f_body
        if j in row and row[j] is not None:
            c.value = row[j]
        if j in (3, 8):
            c.fill = fill_in
        elif j != 1:
            c.fill = fill_calc
        if j == 2:
            c.number_format = "yyyy-mm"
        elif j in (3, 5, 6, 7, 8, 10):
            c.number_format = "#,##0.00"
        elif 11 <= j <= 14:
            c.number_format = "0.0%"
        elif j >= 15:
            c.number_format = "#,##0.00"
fs.freeze_panes = "C14"

if TEST:  # synthetic data only for formula testing; never shipped
    import math, datetime
    fs["B7"] = datetime.date(2019, 1, 1)
    for k in range(60):
        fs.cell(FR0 + k, 3, round(100 + 2 * k + 10 * math.sin(2 * math.pi * k / 12), 2))
    fs["H"+str(FR0+48)] = 200

# ---------------- PESTLE ----------------
pest_rows = [
    ("ex", ["EX-1", "Economic", "Global module prices", "[Paste the stated fact with date, e.g. a price index level and change]", "S#", "Module price index (USD/W)", "Lower prices cut margins but widen demand", "Mixed", "Export price", "A, C", "Member B", "Draft"]),
    ("prop", ["P1", "Political", "Trade defence against Chinese PV in the US, EU and India (tariffs, anti-dumping, local-content rules)", "", "", "Share of exports going to US and EU", "", "Hurts manufacturer", "Western market access", "A, B, C", "Member B", "Proposal"]),
    ("prop", ["P2", "Political", "China industrial policy: 15th Five-Year Plan, support for manufacturing, capacity controls", "", "", "", "", "Mixed", "Supply discipline", "B", "Member B", "Proposal"]),
    ("prop", ["P3", "Economic", "Overcapacity and price war along the supply chain (polysilicon to modules)", "", "", "Export unit value (value / volume)", "", "Hurts manufacturer", "Export price", "A, C", "Member B", "Proposal"]),
    ("prop", ["P4", "Economic", "Demand growth in emerging markets (Middle East, Africa, South and Southeast Asia, Latin America)", "", "S1", "Exports by destination", "", "Helps manufacturer", "Export volume", "A, C", "Member B", "Proposal"]),
    ("prop", ["P5", "Social", "Public and buyer concern over forced labour in the polysilicon supply chain", "", "", "", "", "Hurts manufacturer", "Western market access", "A, B", "Member B", "Proposal"]),
    ("prop", ["P6", "Technological", "Shift in cell technology (e.g. TOPCon, back-contact, perovskite tandem) making older lines obsolete", "", "", "", "", "Mixed", "Technology", "B", "Member B", "Proposal"]),
    ("prop", ["P7", "Legal", "Supply-chain due diligence laws (e.g. US UFLPA, EU Forced Labour Regulation, EU Net-Zero Industry Act criteria)", "", "", "", "", "Hurts manufacturer", "Western market access", "A, B", "Member B", "Proposal"]),
    ("prop", ["P8", "Legal", "China domestic renewable pricing reform and its effect on home installations", "", "S4", "Monthly new installations", "", "Mixed", "Domestic demand", "A, B, C", "Member B", "Proposal"]),
    ("prop", ["P9", "Environmental", "Carbon-footprint requirements in buyer markets; China's climate targets", "", "", "", "", "Mixed", "Cost / margin", "A, B", "Member B", "Proposal"]),
]
table("PESTLE", "PESTLE linked to indicators and model drivers",
      "Each row: a driver, the evidence you hold (fact as stated, with date), the indicator that measures it, and the model driver it moves. 'No indicator yet' is allowed but must be said. Proposals are candidate drivers only: verify every fact and add the source ID.",
      [dict(h="ID", w=6), dict(h="Factor", w=13, list="PESTLE"), dict(h="Driver (what we watch)", w=40),
       dict(h="Evidence: fact as stated, with date", w=44), dict(h="Source ID", w=9), dict(h="Indicator", w=26),
       dict(h="Effect on the manufacturer (our judgement)", w=38), dict(h="Direction", w=16, list="Direction"),
       dict(h="Model driver", w=20, list="Driver"), dict(h="Scenarios affected", w=12), dict(h="Owner", w=11), dict(h="Status", w=11, list="Status")],
      pest_rows, n=30, owner="Member B")

# ---------------- Diamond ----------------
dia = [("Factor conditions", "Does China keep its cost lead (energy, labour, scale, skills, capital)?"),
       ("Demand conditions", "Is the home market large and demanding enough to absorb output and push quality?"),
       ("Related and supporting industries", "How complete is the domestic supply chain (polysilicon, wafers, glass, equipment)?"),
       ("Firm strategy, structure and rivalry", "How intense is rivalry among Chinese producers, and is consolidation happening?"),
       ("Government", "How do Chinese and foreign governments shape costs and market access?"),
       ("Chance", "Which shocks could change the picture (trade war escalation, technology breakthrough, energy crisis)?")]
table("Diamond", "Porter's Diamond for China's PV manufacturing",
      "One row per determinant (add more rows if needed, IDs D7...). Same rule as PESTLE: evidence only where something was read or measured.",
      [dict(h="ID", w=6), dict(h="Determinant", w=26), dict(h="Question", w=44), dict(h="Evidence: fact as stated, with date", w=44),
       dict(h="Source ID", w=9), dict(h="Indicator", w=24), dict(h="Implication for the manufacturer", w=38),
       dict(h="Model driver", w=20, list="Driver"), dict(h="Owner", w=11), dict(h="Status", w=11, list="Status")],
      [("prop", [f"D{i}", d, q, "", "", "", "", None, "Member B", "Proposal"]) for i, (d, q) in enumerate(dia, start=1)],
      n=12, owner="Member B")

# ---------------- Uncertainties ----------------
table("Uncertainties", "Critical uncertainties (scenario axes)",
      "Two uncertainties that matter most to the user and are genuinely unknown. Each scenario on the next sheet sits at one end of each axis. Back each axis with evidence.",
      [dict(h="ID", w=6), dict(h="Uncertainty", w=34), dict(h="One end", w=24), dict(h="Other end", w=24),
       dict(h="Why it is uncertain (evidence)", w=50), dict(h="Source IDs", w=12), dict(h="Status", w=11, list="Status")],
      [("prop", ["U1", "Access to Western markets (US, EU, India)", "Open", "Closed", "", "", "Proposal"]),
       ("prop", ["U2", "Supply discipline in China's PV industry", "Strong", "Weak", "", "", "Proposal"])],
      n=4, owner="Member C")

# ---------------- Scenarios ----------------
table("Scenarios", "Scenarios and slider defaults",
      "2-3 distinct futures. The three number columns set the board's slider defaults; they are assumptions, so write the reasoning and evidence next to them. Volume and price are % per year relative to the baseline forecast. Margin change is percentage points by the end of the forecast horizon. No probabilities.",
      [dict(h="ID", w=6), dict(h="Name", w=24), dict(h="Story (2-3 sentences)", w=50),
       dict(h="U1 position", w=11, list="LowHigh"), dict(h="U2 position", w=11, list="LowHigh"),
       dict(h="Export volume vs baseline (% per yr)", w=14, fmt="0.0"), dict(h="Export price (% per yr)", w=12, fmt="0.0"),
       dict(h="Margin change (pp)", w=11, fmt="0.0"),
       dict(h="Reasoning and evidence for the numbers", w=48), dict(h="Source IDs", w=11), dict(h="Status", w=11, list="Status")],
      [("prop", ["A", "Price war behind walls", "Western markets keep raising barriers while Chinese capacity keeps growing. Surplus output is pushed into the remaining markets and prices keep falling.", "Closed", "Weak", -4, -8, -2.0, "Illustrative starting values. Replace with reasoning from sources.", "", "Proposal"]),
       ("prop", ["B", "Managed consolidation", "Barriers stay, but Beijing and the industry cut capacity. Volumes shrink a little, prices recover and survivors rebuild margins.", "Closed", "Strong", -6, 6, 1.5, "Illustrative starting values. Replace with reasoning from sources.", "", "Proposal"]),
       ("prop", ["C", "Open-door volume race", "Importing countries prioritise cheap deployment over protection and barriers soften, but capacity discipline fails. Volumes surge, prices fall further.", "Open", "Weak", 6, -10, -1.5, "Illustrative starting values. Replace with reasoning from sources.", "", "Proposal"]),
       ], n=5, owner="Member C")

# ---------------- Company inputs ----------------
table("Company inputs", "Illustrative company inputs (the manufacturer)",
      "The board turns scenarios into the manufacturer's revenue and profit. We have no real company data, so these are illustrative and labelled as such. If you find a real listed producer's annual report figures, cite them here.",
      [dict(h="ID", w=6), dict(h="Parameter", w=40), dict(h="Value", w=10), dict(h="Unit", w=14), dict(h="Slider min", w=10),
       dict(h="Slider max", w=10), dict(h="Basis / source", w=50), dict(h="Status", w=11, list="Status")],
      [("prop", ["C1", "Starting operating margin", 3, "%", -10, 20, "Illustrative. PV makers' margins have been thin or negative during the price war: check a listed producer's report.", "Proposal"]),
       ("prop", ["C2", "Share of the company's sales that are exports", 50, "%", 0, 100, "Illustrative.", "Proposal"]),
       ("prop", ["C3", "Company volume sensitivity to market export volume", 1.0, "x", 0.5, 2, "1.0 = company moves with the market.", "Proposal"])],
      n=6, owner="Member C")

# ---------------- Signals ----------------
table("Signals", "Signals (news that changes an assumption)",
      "Only from articles you opened and read. Fact = what the source states. Judgement = what we think it does to each scenario, as numbers in the same units as the Scenarios sheet. Hypothetical tests are allowed if labelled.",
      [dict(h="ID", w=6), dict(h="Headline", w=40), dict(h="Publisher", w=16), dict(h="Date", w=11, fmt="yyyy-mm-dd"),
       dict(h="URL", w=30), dict(h="PESTLE factor", w=12, list="PESTLE"), dict(h="Kind", w=14, list="Kind"),
       dict(h="Fact as stated", w=44), dict(h="Our judgement", w=40),
       dict(h="A volume", w=8, fmt="0.0"), dict(h="A price", w=8, fmt="0.0"), dict(h="A margin", w=8, fmt="0.0"),
       dict(h="B volume", w=8, fmt="0.0"), dict(h="B price", w=8, fmt="0.0"), dict(h="B margin", w=8, fmt="0.0"),
       dict(h="C volume", w=8, fmt="0.0"), dict(h="C price", w=8, fmt="0.0"), dict(h="C margin", w=8, fmt="0.0"),
       dict(h="Read status", w=14, list="ReadStatus"), dict(h="Owner", w=10), dict(h="Status", w=11, list="Status")],
      [("ex", ["EX-1", "[Exact headline]", "[Publisher]", "2026-11-20", "https://...", "Legal", "Reported news",
               "[Measure, rate and start date as the article states them]", "[Why this moves volume, price or margin in each scenario]",
               -1, 0, 0, -0.5, 0, 0, 0, 0, 0, "Read", "Member B", "Draft"])],
      n=30, owner="Member B")

# ---------------- Signposts ----------------
table("Signposts", "Signposts (when to change course)",
      "Indicators we can measure from our data, with a threshold. When a threshold is crossed, it points to a scenario and an action. Thresholds are proposals: justify them.",
      [dict(h="ID", w=6), dict(h="Indicator", w=44), dict(h="Source ID", w=9), dict(h="Unit", w=10),
       dict(h="Threshold", w=10), dict(h="Tripped when", w=11, list="AboveBelow"), dict(h="Points to", w=12),
       dict(h="Action if tripped", w=44), dict(h="Why this threshold", w=40), dict(h="Owner", w=10), dict(h="Status", w=11, list="Status")],
      [("prop", ["G1", "China PV export unit value (value / volume), 12-month change", "S1", "%", -15, "below", "A or C", "Protect cash; stop expanding capacity", "", "Member C", "Proposal"]),
       ("prop", ["G2", "Share of China PV exports going to the EU and US", "S1", "%", 30, "below", "A or B", "Accelerate emerging-market sales and overseas assembly", "", "Member C", "Proposal"]),
       ("prop", ["G3", "China solar cell output, 12-month change", "S2", "%", 0, "below", "B", "Join consolidation; shift to higher-efficiency lines", "", "Member C", "Proposal"]),
       ("prop", ["G4", "EU imports of PV from China, 12-month change", "S3", "%", 10, "above", "C", "Scale EU sales and compliance work", "", "Member C", "Proposal"])],
      n=10, owner="Member C")

# ---------------- Options ----------------
table("Options", "Strategic options for the manufacturer",
      "Score each option from -2 (hurts) to +2 (helps) in each scenario. Verdict is calculated: zero or better in all three = start now (robust); negative in any = prepare and wait for a signpost (contingent). Scores are judgement: write the reason.",
      [dict(h="ID", w=6), dict(h="Option", w=40), dict(h="What it means", w=44),
       dict(h="Score A", w=8, list="Score"), dict(h="Score B", w=8, list="Score"), dict(h="Score C", w=8, list="Score"),
       dict(h="Verdict", w=18, f='=IF(B{r}="","",IF(A{r}="O0","Benchmark",IF(COUNT(D{r}:F{r})<3,"Score all three",IF(MIN(D{r}:F{r})>=0,"Start now","Prepare, then trigger"))))'),
       dict(h="Signpost or signal to act / change course", w=40), dict(h="Reason for scores", w=40), dict(h="Owner", w=10), dict(h="Status", w=11, list="Status")],
      [("prop", ["O0", "Keep the current market mix (benchmark)", "No new move; shows what each future does on its own", None, None, None, None, "Not a recommendation", "", "Member C", "Proposal"]),
       ("prop", ["O1", "Diversify sales to emerging markets", "Middle East, Africa, South and Southeast Asia, Latin America", None, None, None, None, "G2", "", "Member C", "Proposal"]),
       ("prop", ["O2", "Build assembly capacity outside China", "Keeps access to markets with tariffs or local-content rules", None, None, None, None, "G2", "", "Member C", "Proposal"]),
       ("prop", ["O3", "Cut or pause marginal capacity", "Join industry consolidation instead of fighting on price", None, None, None, None, "G1, G3", "", "Member C", "Proposal"]),
       ("prop", ["O4", "Invest in next-generation cell technology", "Higher efficiency to escape the commodity price war", None, None, None, None, "", "", "Member C", "Proposal"]),
       ("prop", ["O5", "Supply-chain traceability and carbon-footprint certification", "Meet due-diligence and carbon rules in EU and US", None, None, None, None, "", "", "Member C", "Proposal"]),
       ("prop", ["O6", "Move downstream", "Storage bundles, project development, services", None, None, None, None, "", "", "Member C", "Proposal"])],
      n=12, owner="Member C")

# ---------------- Log ----------------
table("Log", "Development log",
      "At least three stages (handout). For each: issue or feedback, changes, prompt or code example, verification, effect on quality, screenshot or version link. Select examples where the team questioned or corrected AI output; no raw transcripts.",
      [dict(h="Stage", w=8), dict(h="Date", w=11, fmt="yyyy-mm-dd"), dict(h="Title", w=26), dict(h="Issue or feedback", w=40),
       dict(h="Changes made", w=40), dict(h="Prompt or code example", w=40), dict(h="Verification performed", w=36),
       dict(h="Effect on quality", w=32), dict(h="Screenshot / version link", w=24), dict(h="Owner", w=10)],
      [("prop", [1, "2026-11-23", "Brief and design", "Needed a user, decision and forecastable indicator for China renewable energy",
                 "Focused on solar PV manufacturing; chose a private Chinese manufacturer as user; listed five candidate forecast targets; chose a static board with scheduled data refresh and this workbook as the content source",
                 "Planning with Claude: compared targets on frequency, coverage and fit to the user's decision", "Checked candidate datasets are public and monthly (in progress)",
                 "Clear scope before building", "", "Raul"])],
      n=20, owner="Everyone")

# ---------------- User test ----------------
table("User test", "User test (Thursday)",
      "Test with someone outside the team. Give the tasks below without help, watch, and record what went wrong. No names: use a role.",
      [dict(h="ID", w=6), dict(h="Tester (role only)", w=22), dict(h="Date", w=11, fmt="yyyy-mm-dd"), dict(h="Task", w=46),
       dict(h="Completed?", w=11, list="YN"), dict(h="Problem observed", w=40), dict(h="Severity", w=10, list="Severity"),
       dict(h="Fix", w=36), dict(h="Fixed in version", w=10)],
      [("prop", ["T1", "", None, "Find the latest export figure, its unit and its source", None, "", None, "", ""]),
       ("prop", ["T2", "", None, "Explain what the forecast shows and how uncertain it is", None, "", None, "", ""]),
       ("prop", ["T3", "", None, "Tell observed data, forecast and assumptions apart on the main chart", None, "", None, "", ""]),
       ("prop", ["T4", "", None, "Change one scenario assumption and say what happens to the manufacturer's profit", None, "", None, "", ""]),
       ("prop", ["T5", "", None, "Name the actions that are safe in every scenario and one that depends on a signpost", None, "", None, "", ""]),
       ("prop", ["T6", "", None, "Find out how to refresh the data", None, "", None, "", ""])],
      n=20, owner="Member C")

# ---------------- AI use ----------------
table("AI use", "AI use (accepted, revised, rejected)",
      "The handout asks us to name the LLMs used and what we accepted, rejected or revised. Log the cases that matter, especially wrong numbers or claims you caught.",
      [dict(h="ID", w=6), dict(h="Date", w=11, fmt="yyyy-mm-dd"), dict(h="Member", w=12), dict(h="Tool and model", w=18),
       dict(h="Purpose", w=34), dict(h="What it produced", w=40), dict(h="Outcome", w=11, list="AIOutcome"),
       dict(h="Why / how we checked", w=44)],
      [("ex", ["EX-1", "2026-11-24", "Member B", "ChatGPT (free)", "Summarise EU rules affecting Chinese PV", "A list of measures with dates", "Revised", "Two dates were wrong when checked against the official text; corrected and cited the source"])],
      n=30, owner="Everyone")

# ---------------- Contributions ----------------
table("Contributions", "Contributions",
      "Each member writes their own row and keeps it updated. Used for the reflection and the 5% teamwork criterion.",
      [dict(h="Member", w=16), dict(h="Role", w=26), dict(h="Sheets / parts owned", w=34), dict(h="What I did (concrete)", w=60),
       dict(h="Reviewed whose work", w=24)],
      [("fixed", ["Raul Joshi", "Tech lead", "Dashboard, pipeline, Brief", "", ""]),
       ("fixed", ["Member A", "Data and forecast", "Forecast target, Sources, Data checks, Forecast", "", ""]),
       ("fixed", ["Member B", "Frameworks and news", "PESTLE, Diamond, Signals", "", ""]),
       ("fixed", ["Member C", "Scenarios and strategy", "Uncertainties, Scenarios, Company inputs, Signposts, Options, User test", "", ""])],
      n=4, owner="Everyone")

# move Lists to end and hide
wb.move_sheet("Lists", offset=len(wb.sheetnames))
ls.sheet_state = "hidden"
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 100
wb.active = 0
wb.save(OUT)
print("saved", OUT)
