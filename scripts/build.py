"""Generates mini-ats.xlsx and the CSV files in data/

Run from the repository root:  python scripts/build.py
"""
import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent

# ---------------------------------------------------------------- sample data
STAGES = ["New", "Screening", "Interview", "Offer", "Hired", "Rejected"]
ROLES = ["Junior Data Analyst", "Administrative Assistant"]

candidates = [
    # id, name, role, stage, applied, closed, source, email, notes
    ("C001", "Ana Ribeiro", ROLES[0], "Hired", date(2026, 10, 3), date(2026, 11, 4), "LinkedIn", "ana.ribeiro@example.com", "Starts 1 Dec."),
    ("C002", "Tiago Matos", ROLES[0], "Rejected", date(2026, 10, 2), date(2026, 10, 7), "ITJobs.pt", "tiago.matos@example.com", "No SQL."),
    ("C003", "Inês Carvalho", ROLES[0], "Offer", date(2026, 10, 6), None, "Referral", "ines.carvalho@example.com", "Offer sent 18/11, waiting."),
    ("C004", "Rui Ferreira", ROLES[0], "Interview", date(2026, 10, 21), None, "LinkedIn", "rui.ferreira@example.com", ""),
    ("C005", "Sofia Lopes", ROLES[0], "Screening", date(2026, 11, 10), None, "Landing.jobs", "sofia.lopes@example.com", "Call her back after 15h"),
    ("C006", "Marta Sousa", ROLES[1], "Hired", date(2026, 9, 29), date(2026, 10, 17), "Net-Empregos", "marta.sousa@example.com", "Started 3/11"),
    ("C007", "João Pinto", ROLES[1], "Rejected", date(2026, 10, 6), date(2026, 10, 20), "Indeed", "joao.pinto@example.com", "Only wants remote, role is hybrid"),
    ("C008", "Beatriz Costa", ROLES[1], "Interview", date(2026, 10, 28), None, "LinkedIn", "beatriz.costa@example.com", ""),
    ("C009", "Pedro Almeida", ROLES[1], "New", date(2026, 11, 18), None, "Indeed", "pedro.almeida@example.com", ""),
    ("C010", "Carla Mendes", ROLES[0], "New", date(2026, 11, 20), None, "Referral", "carla.mendes@example.com", "Referred by Ana R."),
    ("C011", "Diogo Nunes", ROLES[0], "Rejected", date(2026, 10, 4), date(2026, 10, 5), "LinkedIn", "diogo.nunes@example.com", "Lives in Brazil, no visa"),
    ("C012", "Mariana Teixeira", ROLES[0], "Rejected", date(2026, 10, 9), date(2026, 10, 30), "ITJobs.pt", "mariana.teixeira@example.com", "Good interview but salary way above budget"),
    ("C013", "Hugo Martins", ROLES[1], "Hired", date(2026, 9, 22), date(2026, 11, 8), "Indeed", "hugo.martins@example.com", "Took long because of 30 day notice period"),
    ("C014", "Rita Gomes", ROLES[1], "Rejected", date(2026, 10, 1), date(2026, 10, 2), "Indeed", "rita.gomes@example.com", ""),
    ("C015", "Filipe Rocha", ROLES[1], "Rejected", date(2026, 10, 12), date(2026, 10, 13), "Net-Empregos", "filipe.rocha@example.com", ""),
    ("C016", "Cláudia Pires", ROLES[0], "Screening", date(2026, 11, 16), None, "LinkedIn", "claudia.pires@example.com", ""),
    ("C017", "André Silva", ROLES[1], "Rejected", date(2026, 10, 15), date(2026, 11, 3), "LinkedIn", "andre.silva@example.com", "Withdrew - accepted another offer"),
    ("C018", "Joana Moreira", ROLES[0], "Interview", date(2026, 11, 2), None, "Referral", "joana.moreira@example.com", "2nd interview to book"),
    ("C019", "Nuno Barbosa", ROLES[1], "New", date(2026, 11, 21), None, "", "nuno.barbosa@example.com", "CV came by email, source?"),
    ("C020", "Vera Antunes", ROLES[0], "Rejected", date(2026, 10, 19), date(2026, 10, 20), "Landing.jobs", "vera.antunes@example.com", ""),
    ("C021", "Luís Correia", ROLES[1], "Screening", date(2026, 11, 19), None, "Indeed", "luis.correia@example.com", ""),
]

requirements = [
    (ROLES[0], "Excel (pivot tables, lookups)", "SQL basics", "English B2+", "Power BI or Tableau", "Data/IT degree or course"),
    (ROLES[1], "Excel / Office", "Portuguese + English", "1+ year admin experience", "CRM or ERP experience", "Hybrid in Porto"),
]

screening = {  # id -> Req 1..5 (1-3 mandatory, 4-5 nice to have); None = not screened yet
    "C001": (1, 1, 1, 1, 0), "C002": (1, 0, 1, 0, 0), "C003": (1, 1, 1, 1, 1),
    "C004": (1, 1, 1, 0, 1), "C005": (1, 1, 1, 0, 0), "C006": (1, 1, 1, 1, 0),
    "C007": (1, 1, 1, 0, 0), "C008": (1, 1, 1, 1, 0), "C009": None, "C010": None,
    "C011": (1, 1, 1, 0, 1), "C012": (1, 1, 1, 1, 1), "C013": (1, 1, 1, 0, 0),
    "C014": (1, 0, 0, 0, 0), "C015": (0, 1, 1, 0, 0), "C016": (1, 1, 0, 1, 0),
    "C017": (1, 1, 1, 1, 0), "C018": (1, 1, 1, 0, 0), "C019": None,
    "C020": (0, 0, 1, 0, 0), "C021": (1, 1, 1, 0, 0),
}

stage_since = {  # last time the candidate changed stage (open candidates only)
    "C003": date(2026, 11, 18), "C004": date(2026, 11, 14), "C005": date(2026, 11, 12),
    "C008": date(2026, 11, 20), "C009": date(2026, 11, 18), "C010": date(2026, 11, 20),
    "C016": date(2026, 11, 17), "C018": date(2026, 11, 12), "C019": date(2026, 11, 21),
    "C021": date(2026, 11, 19),
}

interviews = [
    ("C001", date(2026, 10, 13), "10:00", "Video", "Hiring manager", 4, "Strong Excel, explained a KPI dashboard well", "Second interview"),
    ("C001", date(2026, 10, 20), "14:30", "In person", "Head of Data", 5, "SQL case ok. good fit with team", "Make offer"),
    ("C012", date(2026, 10, 22), "17:00", "Video", "Hiring manager", 4, "Asks 38k, budget is 28k", "Reject"),
    ("C003", date(2026, 10, 23), "11:15", "Video", "Hiring manager", 3, "A bit nervous. Power BI project on GitHub is good", "Second interview"),
    ("C003", date(2026, 11, 5), "15:00", "In person", "Head of Data", 4, "Much better, good business questions", "Make offer"),
    ("C006", date(2026, 10, 8), "09:30", "In person", "Office manager", 5, "Very organised, knows PHC", "Make offer"),
    ("C013", date(2026, 10, 9), "16:00", "Phone", "Office manager", 4, "Has 30 days notice", "Make offer"),
    ("C007", date(2026, 10, 16), "12:00", "Phone", "Office manager", 2, "Doesn't want hybrid", "Reject"),
    ("C017", date(2026, 10, 27), "10:00", "In person", "Office manager", 4, "", "Second interview"),
    ("C018", date(2026, 11, 12), "18:00", "Video", "Hiring manager", 4, "Only free after 18h", "Second interview"),
    ("C004", date(2026, 11, 25), "10:30", "Video", "Hiring manager", None, "", "To decide"),
    ("C008", date(2026, 11, 27), "11:00", "In person", "Office manager", None, "", "To decide"),
]

# ---------------------------------------------------------------- styles
FONT = "Calibri"
INK = "333333"
MUTED = "777777"
HEADER_FILL = PatternFill("solid", fgColor="F2F2F2")
LINE = Side(style="thin", color="DDDDDD")
base_font = Font(name=FONT, size=11, color=INK)
header_font = Font(name=FONT, size=11, bold=True, color=INK)
DATE_FMT = "dd/mm/yyyy"
ROWS = 200  # rows prepared for new entries


def style_header(ws, row, first_col, last_col):
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = HEADER_FILL
        cell.border = Border(bottom=LINE)
        cell.alignment = Alignment(vertical="center")


def base_style(ws):
    for row in ws.iter_rows():
        for cell in row:
            if cell.font.name != FONT:  # cells still in the default font
                f = cell.font
                cell.font = Font(name=FONT, size=f.size if f.size and f.size != 11 else 11, bold=f.bold, italic=f.italic,
                                 color=f.color if f.color and f.color.rgb not in (None, "FF000000") else INK)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def lookup(id_cell, col):
    """Name/role from the Candidates sheet. INDEX/MATCH works in every Excel version."""
    return (f'=IF({id_cell}="","",IFERROR(INDEX(Candidates!${col}$2:${col}${ROWS},'
            f'MATCH({id_cell},Candidates!$A$2:$A${ROWS},0)),""))')


wb = Workbook()

# ---------------------------------------------------------------- Dashboard
dash = wb.active
dash.title = "Dashboard"

# ---------------------------------------------------------------- Candidates
cand = wb.create_sheet("Candidates")
cand_head = ["ID", "Candidate", "Role", "Stage", "In stage since", "Days in stage", "Applied on", "Closed on",
             "Days to close", "Screening score", "Interview rating", "Source", "Email", "Notes"]
cand.append(cand_head)
style_header(cand, 1, 1, len(cand_head))
for i, (cid, name, role, stage, applied, closed, source, email, notes) in enumerate(candidates, start=2):
    cand.append([cid, name, role, stage, stage_since.get(cid),
                 f'=IF(OR(E{i}="",H{i}<>""),"",TODAY()-E{i})',
                 applied, closed,
                 f'=IF(OR(G{i}="",H{i}=""),"",H{i}-G{i})',
                 f'=IF(A{i}="","",IFERROR(INDEX(\'CV Screening\'!$I$11:$I${ROWS},MATCH(A{i},\'CV Screening\'!$A$11:$A${ROWS},0)),""))',
                 f'=IF(COUNTIFS(Interviews!$A$2:$A${ROWS},A{i},Interviews!$H$2:$H${ROWS},">0")=0,"",AVERAGEIFS(Interviews!$H$2:$H${ROWS},Interviews!$A$2:$A${ROWS},A{i}))',
                 source or None, email, notes or None])
for r in range(2, ROWS + 1):
    for col in "EGH":
        cand[f"{col}{r}"].number_format = DATE_FMT
    cand[f"F{r}"].number_format = "0"
    cand[f"I{r}"].number_format = "0"
    cand[f"J{r}"].number_format = "0%"
    cand[f"K{r}"].number_format = "0.0"
    cand[f"N{r}"].alignment = Alignment(wrap_text=True, vertical="top")
for rng, opts in ((f"D2:D{ROWS}", ",".join(STAGES)), (f"C2:C{ROWS}", ",".join(ROLES)),
                  (f"L2:L{ROWS}", "LinkedIn,Indeed,Referral,ITJobs.pt,Net-Empregos,Landing.jobs,Other")):
    dv = DataValidation(type="list", formula1=f'"{opts}"', allow_blank=True)
    cand.add_data_validation(dv)
    dv.add(rng)
stage_fill = {"Hired": "E6F2EA", "Rejected": "F8EAEA", "Offer": "FBF5E3"}
for stage, color in stage_fill.items():
    cand.conditional_formatting.add(
        f"D2:D{ROWS}", FormulaRule(formula=[f'$D2="{stage}"'], fill=PatternFill("solid", bgColor=color)))
# soft amber when someone has been waiting more than 5 days
cand.conditional_formatting.add(
    f"F2:F{ROWS}", FormulaRule(formula=['AND(ISNUMBER($F2),$F2>5)'],
                               fill=PatternFill("solid", bgColor="FBEBD4"), font=Font(color="9A5B13")))
cand.freeze_panes = "C2"
widths(cand, {"A": 7, "B": 17, "C": 24, "D": 11, "E": 14, "F": 13, "G": 12, "H": 12, "I": 13,
              "J": 15, "K": 15, "L": 14, "M": 28, "N": 40})

# ---------------------------------------------------------------- CV Screening
scr = wb.create_sheet("CV Screening")
scr["A1"] = "1 = has it, 0 = doesn't. Req 1-3 are mandatory."
scr["A1"].font = Font(name=FONT, size=11, color=MUTED)
scr["A3"] = "Requirements by role"
scr["A3"].font = header_font
req_head = ["Role", "Req 1 (must)", "Req 2 (must)", "Req 3 (must)", "Req 4 (nice)", "Req 5 (nice)"]
for c, v in enumerate(req_head, start=1):
    scr.cell(row=4, column=c, value=v)
style_header(scr, 4, 1, 6)
for r, row in enumerate(requirements, start=5):
    for c, v in enumerate(row, start=1):
        scr.cell(row=r, column=c, value=v).alignment = Alignment(wrap_text=True, vertical="top")

scr["A9"] = "Screening checklist"
scr["A9"].font = header_font
chk_head = ["ID", "Candidate", "Role", "Req 1", "Req 2", "Req 3", "Req 4", "Req 5", "Score", "Decision"]
for c, v in enumerate(chk_head, start=1):
    scr.cell(row=10, column=c, value=v)
style_header(scr, 10, 1, 10)
for r, (cid, _, *rest) in enumerate(candidates, start=11):
    marks = screening[cid]
    scr.cell(row=r, column=1, value=cid)
    scr.cell(row=r, column=2, value=lookup(f"A{r}", "B"))
    scr.cell(row=r, column=3, value=lookup(f"A{r}", "C"))
    for k, m in enumerate(marks or ()):
        scr.cell(row=r, column=4 + k, value=m)
for r in range(11, ROWS + 1):
    scr[f"I{r}"] = f'=IF(COUNT(D{r}:H{r})=0,"",SUM(D{r}:H{r})/5)'
    scr[f"J{r}"] = (f'=IF(COUNT(D{r}:H{r})=0,IF(A{r}="","","Not screened"),'
                    f'IF(SUM(D{r}:F{r})=3,"Advance","Reject"))')
    scr[f"I{r}"].number_format = "0%"
    for col in "DEFGH":
        scr[f"{col}{r}"].alignment = Alignment(horizontal="center")
dv_bin = DataValidation(type="list", formula1='"1,0"', allow_blank=True)
scr.add_data_validation(dv_bin)
dv_bin.add(f"D11:H{ROWS}")
scr.conditional_formatting.add(f"J11:J{ROWS}", FormulaRule(formula=['$J11="Advance"'], font=Font(color="3C7A55")))
scr.conditional_formatting.add(f"J11:J{ROWS}", FormulaRule(formula=['$J11="Reject"'], font=Font(color="A05050")))
scr.freeze_panes = "A11"
widths(scr, {"A": 24, "B": 22, "C": 22, "D": 22, "E": 22, "F": 22, "G": 8, "H": 8, "I": 9, "J": 14})

# ---------------------------------------------------------------- Interviews
itv = wb.create_sheet("Interviews")
itv_head = ["ID", "Candidate", "Role", "Date", "Time", "Type", "Interviewer", "Rating (1-5)", "Notes", "Next step"]
itv.append(itv_head)
style_header(itv, 1, 1, len(itv_head))
for r, (cid, d, t, kind, who, rating, notes, nxt) in enumerate(interviews, start=2):
    hh, mm = map(int, t.split(":"))
    itv.append([cid, lookup(f"A{r}", "B"), lookup(f"A{r}", "C"), d, None, kind, who, rating, notes, nxt])
    itv[f"E{r}"] = f"=TIME({hh},{mm},0)"
for r in range(2, ROWS + 1):
    itv[f"D{r}"].number_format = DATE_FMT
    itv[f"E{r}"].number_format = "hh:mm"
    itv[f"I{r}"].alignment = Alignment(wrap_text=True, vertical="top")
for rng, opts in ((f"F2:F{ROWS}", "Phone,Video,In person"),
                  (f"J2:J{ROWS}", "To decide,Second interview,Make offer,Reject"),
                  (f"H2:H{ROWS}", "1,2,3,4,5")):
    dv = DataValidation(type="list", formula1=f'"{opts}"', allow_blank=True)
    itv.add_data_validation(dv)
    dv.add(rng)
itv.freeze_panes = "C2"
widths(itv, {"A": 7, "B": 16, "C": 24, "D": 12, "E": 8, "F": 11, "G": 16, "H": 12, "I": 42, "J": 17})

# ---------------------------------------------------------------- Email templates
tpl = wb.create_sheet("Email Templates")
tpl_rows = [line for line in csv.reader((ROOT / "templates" / "email-templates.csv").open(encoding="utf-8"))]
for row in tpl_rows:
    tpl.append(row)
style_header(tpl, 1, 1, 5)
for r in range(2, len(tpl_rows) + 1):
    for c in range(1, 6):
        tpl.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")
widths(tpl, {"A": 20, "B": 9, "C": 22, "D": 36, "E": 70})

# ---------------------------------------------------------------- Dashboard content
SOURCES = ["LinkedIn", "Indeed", "Referral", "ITJobs.pt", "Net-Empregos", "Landing.jobs"]
C = lambda col: f"Candidates!${col}$2:${col}${ROWS}"
dash.column_dimensions["A"].width = 3
dash["B2"] = "Recruitment overview"
dash["B2"].font = Font(name=FONT, size=14, bold=True, color=INK)

# the one number that matters most
dash["B4"] = "Avg days to hire"
dash["B4"].font = Font(name=FONT, size=11, color=MUTED)
dash["B5"] = f'=IFERROR(AVERAGEIFS({C("I")},{C("D")},"Hired"),"-")'
dash["B5"].number_format = "0"
dash["B5"].font = Font(name=FONT, size=26, bold=True, color=INK)
dash["B5"].alignment = Alignment(horizontal="left")
dash["C5"] = f'="from "&COUNTIF({C("D")},"Hired")&" hires"'
dash["C5"].font = Font(name=FONT, size=11, color=MUTED)
dash["C5"].alignment = Alignment(vertical="bottom")

# secondary numbers, small
small = [
    ("Candidates", f'=COUNTA({C("A")})'),
    ("Still in process", f'=COUNTIF({C("D")},"New")+COUNTIF({C("D")},"Screening")+COUNTIF({C("D")},"Interview")+COUNTIF({C("D")},"Offer")'),
    ("Waiting > 5 days", f'=COUNTIF({C("F")},">5")'),
    ("Interviews this week", f'=COUNTIFS(Interviews!$D$2:$D${ROWS},">="&TODAY(),Interviews!$D$2:$D${ROWS},"<"&TODAY()+7)'),
]
for k, (label, formula) in enumerate(small):
    dash.cell(row=4, column=5 + k, value=label).font = Font(name=FONT, size=10, color=MUTED)
    v = dash.cell(row=5, column=5 + k, value=formula)
    v.font = Font(name=FONT, size=14, color=INK)
    v.alignment = Alignment(horizontal="left")

def table(title, top, left, headers, rows, fmts):
    dash.cell(row=top, column=left, value=title).font = header_font
    for c, v in enumerate(headers):
        dash.cell(row=top + 1, column=left + c, value=v)
    style_header(dash, top + 1, left, left + len(headers) - 1)
    for i, row in enumerate(rows):
        for c, v in enumerate(row):
            cell = dash.cell(row=top + 2 + i, column=left + c, value=v)
            cell.border = Border(bottom=LINE)
            if fmts[c]:
                cell.number_format = fmts[c]

table("Where candidates are now", 8, 2, ["Stage", "Candidates"],
      [(st, f'=COUNTIF({C("D")},B{10 + i})') for i, st in enumerate(STAGES)], [None, "0"])

src_rows = [(src, f'=COUNTIF({C("L")},E{10 + i})', f'=COUNTIFS({C("L")},E{10 + i},{C("D")},"Hired")')
            for i, src in enumerate(SOURCES)]
src_rows.append(("Not recorded", f'=COUNTIFS({C("A")},"<>",{C("L")},"")', f'=COUNTIFS({C("A")},"<>",{C("L")},"",{C("D")},"Hired")'))
table("Which channels bring hires", 8, 5, ["Source", "Candidates", "Hired"], src_rows, [None, "0", "0"])

table("Time to hire by role", 19, 5, ["Role", "Hired", "Avg days"],
      [(role, f'=COUNTIFS({C("C")},E{21 + i},{C("D")},"Hired")',
        f'=IFERROR(AVERAGEIFS({C("I")},{C("C")},E{21 + i},{C("D")},"Hired"),"-")') for i, role in enumerate(ROLES)],
      [None, "0", "0"])

chart = BarChart()
chart.type = "bar"
chart.title = None
chart.legend = None
chart.roundedCorners = False
chart.x_axis.delete = False
chart.y_axis.delete = False
chart.y_axis.majorGridlines = None
chart.y_axis.majorUnit = 2
chart.x_axis.scaling.orientation = "maxMin"
chart.add_data(Reference(dash, min_col=3, min_row=9, max_row=15), titles_from_data=True)
chart.set_categories(Reference(dash, min_col=2, min_row=10, max_row=15))
chart.series[0].graphicalProperties.solidFill = "8A8F98"
chart.series[0].graphicalProperties.line.solidFill = "8A8F98"
chart.gapWidth = 60
chart.width, chart.height = 6.6, 5.5
dash.add_chart(chart, "B18")
widths(dash, {"B": 20, "C": 12, "D": 4, "E": 24, "F": 16, "G": 16, "H": 18})

for ws in wb.worksheets:
    base_style(ws)
wb.save(ROOT / "mini-ats.xlsx")

# ---------------------------------------------------------------- CSV exports (values only)
with (ROOT / "data" / "candidates.csv").open("w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["id", "candidate", "role", "stage", "applied_on", "closed_on", "source"])
    for cid, name, role, stage, applied, closed, source, *_ in candidates:
        w.writerow([cid, name, role, stage, applied.isoformat(), closed.isoformat() if closed else "", source])
with (ROOT / "data" / "interviews.csv").open("w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["id", "date", "time", "type", "interviewer", "rating", "next_step"])
    for cid, d, t, kind, who, rating, _, nxt in interviews:
        w.writerow([cid, d.isoformat(), t, kind, who, rating if rating else "", nxt])
with (ROOT / "data" / "cv_screening.csv").open("w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["id", "req_1", "req_2", "req_3", "req_4", "req_5"])
    for cid, marks in screening.items():
        w.writerow([cid, *(marks or [""] * 5)])
print("built")
