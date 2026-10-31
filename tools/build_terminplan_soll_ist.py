"""Terminplan P800 mit Soll-Ist-Vergleich und automatischer Verschiebung."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

# Stammdaten (Vorgänge, Rollen) aus dem bestehenden Fertigungsplan übernehmen
src = open("tools/build_fertigungsplan.py").read()
exec(src[: src.index("wb = Workbook()")])

OUT = "/tmp/build/Terminplan_P800_Soll_Ist.xlsx"

# Anpassungen: SLV-Prüfung als Vorgang mit Dauer
A[0][2][1] = ("F02", "SLV", "SLV-Prüfung Testschuss 1 bis Ergebnis", "SA", "02.11.26", "20.11.26", "Proben am 30.10. versandt")

# Vorgänger (Ende → Start). Abstand wird aus dem Soll-Plan berechnet.
PRED = {
    "F02": ["T09"], "F03": ["F02"],
    "T05": ["T01"], "T06": ["T03"], "T07": ["T06"], "T08": ["T07"], "T09": ["T07"],
    "T11": ["T08", "T10"], "T12": ["T11"], "T13": ["T12"], "T14": ["T13"], "T15": ["T14"], "T16": ["T15"], "T17": ["T16"], "T18": ["T17"],
    "A02": ["A01"], "A03": ["A02"], "A04": ["A03"], "A05": ["A04"], "A06": ["A05", "F01"], "A07": ["A06"], "A08": ["A07"],
    "A09": ["A08"], "A10": ["A09"], "A11": ["A08"], "A12": ["A11"], "A13": ["A10"], "A14": ["A13"], "A15": ["A14"], "A16": ["A15"],
    "A17": ["A16", "F03"], "A18": ["A17"], "A19": ["A18"], "A20": ["A19"], "A21": ["A20"], "A22": ["A21"], "A23": ["A22"], "A24": ["A23"],
    "B01": ["A02"], "B02": ["B01"], "B03": ["B02"], "B04": ["B03", "A04"], "B05": ["B04"], "B06": ["B05", "A06"], "B07": ["B06"],
    "B08": ["B07", "A08"], "B09": ["B08"], "B10": ["B09"], "B11": ["B10"], "B12": ["B10"], "B13": ["B11"], "B14": ["B13"],
    "B15": ["B14", "A17"], "B16": ["B15"], "B17": ["B16"], "B18": ["B17"], "B19": ["B18"], "B20": ["B19"], "B21": ["B20"],
    "B22": ["B21", "A23"], "B23": ["B22"],
    "C01": ["A24", "B23"], "C02": ["C01"], "C03": ["C02"], "C04": ["C03"], "C05": ["C04"], "C06": ["C05"], "C07": ["C06"], "C08": ["C07"],
}

HOLS = set(h.date() for h in HOL)
def is_wd(x): return x.weekday() < 5 and x not in HOLS
def netwd(a, b):  # wie NETWORKDAYS
    sgn = 1
    if b < a: a, b, sgn = b, a, -1
    n, x = 0, a
    while x <= b:
        n += is_wd(x); x += dt.timedelta(1)
    return n * sgn
def lag_of(end_pred, start):
    """Abstand in AT: 0 = gleicher Tag, 1 = nächster Arbeitstag, negativ = Überlappung. Puffer wird nicht übernommen (max. 1)."""
    if start >= end_pred: g = netwd(end_pred, start) - 1
    else: g = -(netwd(start, end_pred) - 1)
    return min(g, 1)

# Kalender für Balkenplan (Arbeitstage)
days = []
x = dt.date(2026, 10, 19)
while x <= dt.date(2027, 3, 31):
    if is_wd(x): days.append(x)
    x += dt.timedelta(1)

ADARK = {"Freigaben": "B54A0F", "Testschuss 1": "5E6B78", "Schuss 301": "1F78AE", "Schuss 302": "3F7F2E", "Rundnaht 301+302": NAVY}
INPUT = "FFF4CC"

wb = Workbook()
wsU = wb.active; wsU.title = "Übersicht"
wsP = wb.create_sheet("Terminplan")
wsV = wb.create_sheet("Verantwortliche")
wsA = wb.create_sheet("Anleitung")
wsL = wb.create_sheet("Listen")

# ---------------- Listen
wsL["A1"] = "Betriebsruhe / Feiertage"
for i, h in enumerate(HOL): wsL.cell(2 + i, 1, h).number_format = "DD.MM.YYYY"
URS = ["Personal", "Material", "Anlage / Vorrichtung", "Qualität / Nacharbeit", "Freigabe / WPS / SLV", "Vorgänger verspätet", "Transport", "Sonstiges"]
wsL["C1"] = "Ursachen"
for i, u in enumerate(URS): wsL.cell(2 + i, 3, u)
wsL.sheet_state = "hidden"
wb.defined_names["Feiertage"] = DefinedName("Feiertage", attr_text=f"Listen!$A$2:$A${1+len(HOL)}")

# ---------------- Terminplan
P = wsP
P.sheet_view.showGridLines = False
P.sheet_view.zoomScale = 90
COLS = ["ID", "Arbeitsschritt", "Station", "Verantwortlich", "Vorgänger",
        "Soll-Start", "Soll-Ende", "Dauer AT",
        "Ist-Start", "Ist-Ende", "Fertig %", "Mehr-aufwand AT",
        "Prognose Start", "Prognose Ende", "Verzug Ende AT", "Status",
        "Ursache", "Bemerkung / Maßnahme"]
WID = [6, 38, 10, 24, 11, 10, 10, 7, 10, 10, 8, 9, 10, 10, 9, 19, 19, 30]
# Hilfsspalten S..V (ausgeblendet): Soll erledigt AT, Ist erledigt AT, Abschnitt, Verzug Start AT
G0 = 24  # X = erste Kalenderspalte
GL = G0 + len(days) - 1
gl = CL(GL)
P["A1"] = "Terminplan P800 · Soll-Ist-Vergleich"; P["A1"].font = font(size=16, bold=True, color=NAVY)
P["A2"] = '="Stichtag: "&TEXT(Stichtag,"DD.MM.YYYY")&"   ·   Gelbe Spalten ausfüllen, alles andere rechnet automatisch. Verschiebt sich ein Vorgang, verschieben sich alle Nachfolger."'
P["A2"].font = font(size=10, color=GREY)
# Gruppenköpfe Zeile 4
groups = [(1, 5, "VORGANG", NAVY), (6, 8, "SOLL (Basisplan, fest)", "5E6B78"), (9, 12, "IST (eintragen)", "B8860B"), (13, 16, "PROGNOSE / AUSWIRKUNG", ORANGE), (17, 18, "KLÄRUNG (eintragen)", "B8860B")]
for a_, b_, t, c_ in groups:
    P.merge_cells(start_row=4, start_column=a_, end_row=4, end_column=b_)
    cc = P.cell(4, a_, t); cc.font = font(bold=True, color="FFFFFF", size=9); cc.alignment = CENTER
    for k in range(a_, b_ + 1): P.cell(4, k).fill = fill(c_)
for j, h in enumerate(COLS, 1):
    P.merge_cells(start_row=5, start_column=j, end_row=7, end_column=j)
    c = P.cell(5, j, h); c.font = font(bold=True, color="FFFFFF", size=9)
    c.alignment = Alignment(horizontal="left" if j in (2, 18) else "center", vertical="center", wrap_text=True)
    for r_ in (5, 6, 7):
        P.cell(r_, j).fill = fill("8C6A00" if 9 <= j <= 12 or j >= 17 else NAVY); P.cell(r_, j).border = BOX
    P.column_dimensions[CL(j)].width = WID[j - 1]
for j, h in zip(range(19, 23), ["Soll erl. AT", "Ist erl. AT", "Abschnitt", "Verzug Start AT"]):
    P.cell(7, j, h).font = font(size=8, color=GREY)
    P.column_dimensions[CL(j)].width = 9
    P.column_dimensions[CL(j)].hidden = True
P.column_dimensions["W"].width = 1.5
# Kalenderkopf
WD = ["Mo", "Di", "Mi", "Do", "Fr"]
wk_start = G0
for i, dd in enumerate(days):
    cidx = G0 + i
    c6 = P.cell(6, cidx, dd); c6.number_format = "DD.MM."; c6.font = font(size=7, bold=True, color=NAVY)
    c6.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90); c6.fill = fill(LIGHT); c6.border = BOX
    c7 = P.cell(7, cidx, WD[dd.weekday()][0]); c7.font = font(size=7, color=GREY); c7.alignment = CENTER; c7.fill = fill(LIGHT); c7.border = BOX
    P.column_dimensions[CL(cidx)].width = 2.6
    nxt = days[i + 1] if i + 1 < len(days) else None
    if nxt is None or nxt.isocalendar()[1] != dd.isocalendar()[1]:
        P.merge_cells(start_row=5, start_column=wk_start, end_row=5, end_column=cidx)
        wf = fill(ORANGE if dd.isocalendar()[1] % 2 else "3A4B5C")
        c5 = P.cell(5, wk_start, f"KW {dd.isocalendar()[1]}"); c5.font = font(size=8, bold=True, color="FFFFFF"); c5.alignment = CENTER
        for cc in range(wk_start, cidx + 1): P.cell(5, cc).fill = wf
        wk_start = cidx + 1
    if dd.day <= 7 and (i == 0 or days[i - 1].month != dd.month):
        P.cell(4, cidx, dd.strftime("%b %y").replace("Oct", "Okt").replace("Dec", "Dez").replace("Mar", "Mär").replace("May", "Mai")).font = font(size=9, bold=True, color=NAVY)
P.row_dimensions[6].height = 34

rowof = {}
r = 8
data_rows = []
sec_rows = {}
for name, colr, tasks in A:
    P.merge_cells(start_row=r, start_column=1, end_row=r, end_column=18)
    c = P.cell(r, 1, name.upper()); c.font = font(bold=True, color="FFFFFF", size=10); c.alignment = Alignment(vertical="center", indent=1)
    for cc in range(1, 19): P.cell(r, cc).fill = fill(colr)
    for cc in range(G0, GL + 1): P.cell(r, cc).fill = fill(colr)
    P.row_dimensions[r].height = 18
    r += 1
    sec_rows[name] = []
    for (tid, st, step, role, s, e, hint) in tasks:
        rowof[tid] = r
        sec_rows[name].append(r)
        r += 1
    r += 0
LAST = r - 1
# Werte und Formeln
info = {t[0]: (name, t) for name, _, ts in A for t in ts}
for tid, rr in rowof.items():
    name, (tid_, st, step, role, s, e, hint) = info[tid]
    ss, se = D(s), D(e)
    preds = PRED.get(tid, [])
    vals = {1: tid, 2: step, 3: st, 4: R[role], 5: ", ".join(preds), 6: ss, 7: se}
    for j, v in vals.items():
        c = P.cell(rr, j, v); c.font = font(size=9, bold=(j == 1)); c.border = BOX
        c.alignment = CENTER if j in (1, 3, 5, 6, 7) else LEFT
        if j in (6, 7): c.number_format = "DD.MM.YY"
    if hint:
        P.cell(rr, 2).comment = Comment(hint, "Fertigungssteuerung")
        if "abgeleitet" in hint: P.cell(rr, 2).font = font(size=9, italic=True)
        if "Meilenstein" in hint: P.cell(rr, 2).font = font(size=9, bold=True)
    P.cell(rr, 8, f"=NETWORKDAYS(F{rr},G{rr},Feiertage)")
    # Prognose Start
    terms = [f"F{rr}", "StichtagAT"]
    for p_ in preds:
        pr = rowof[p_]
        lg = lag_of(D(info[p_][1][5]).date(), ss.date())
        terms.append(f"WORKDAY($N${pr},{lg},Feiertage)")
    P.cell(rr, 13, f'=IF(I{rr}<>"",I{rr},MAX({",".join(terms)}))')
    P.cell(rr, 14, f'=IF(J{rr}<>"",J{rr},MAX(WORKDAY(M{rr},H{rr}-1+N(L{rr}),Feiertage),IF(I{rr}<>"",StichtagAT,0)))')
    P.cell(rr, 15, f"=IF(N{rr}>=G{rr},NETWORKDAYS(G{rr},N{rr},Feiertage)-1,-(NETWORKDAYS(N{rr},G{rr},Feiertage)-1))")
    P.cell(rr, 16, f'=IF(J{rr}<>"",IF(O{rr}>0,"Erledigt, verspätet","Erledigt"),IF(I{rr}<>"",IF(O{rr}>0,"Läuft, Verzug","Läuft"),IF(F{rr}<StichtagAT,"Start überfällig",IF(O{rr}>0,"Verschoben","Geplant"))))')
    # Hilfsspalten
    P.cell(rr, 19, f"=IF(Stichtag<=F{rr},0,IF(Stichtag>G{rr},H{rr},NETWORKDAYS(F{rr},Stichtag-1,Feiertage)))")
    P.cell(rr, 20, f'=H{rr}*IF(J{rr}<>"",1,IF(I{rr}="",0,IF(K{rr}="",0.5,MIN(K{rr},0.95))))')
    P.cell(rr, 21, name)
    P.cell(rr, 22, f"=IF(M{rr}>=F{rr},NETWORKDAYS(F{rr},M{rr},Feiertage)-1,-(NETWORKDAYS(M{rr},F{rr},Feiertage)-1))")
    for j in range(8, 19):
        c = P.cell(rr, j); c.border = BOX; c.font = font(size=9, bold=(j in (14, 15, 16)))
        c.alignment = LEFT if j in (17, 18) else CENTER
        if j in (9, 10, 11, 12, 17, 18): c.fill = fill(INPUT)
    for j in (9, 10, 13, 14): P.cell(rr, j).number_format = "DD.MM.YY"
    P.cell(rr, 11).number_format = "0%"
    P.cell(rr, 15).number_format = '+0;-0;"–"'
    for cc in range(G0, GL + 1): P.cell(rr, cc).border = Border(left=Side(style="hair", color="E3E7EB"), right=Side(style="hair", color="E3E7EB"), top=thin, bottom=thin)
    data_rows.append(rr)

# Wochenlinien
for i, dd in enumerate(days):
    if i == 0 or dd.isocalendar()[1] != days[i - 1].isocalendar()[1]:
        for rr in range(6, LAST + 1):
            c = P.cell(rr, G0 + i)
            c.border = Border(left=Side(style="thin", color="8A96A3"), right=c.border.right, top=c.border.top, bottom=c.border.bottom)
# Balkenplan per bedingter Formatierung
GR = f"{CL(G0)}8:{gl}{LAST}"
X = f"{CL(G0)}$6"
cf = P.conditional_formatting
cf.add(GR, FormulaRule(formula=[f'AND({X}=Stichtag,$A8<>"")'], fill=fill(ORANGE), stopIfTrue=True))
cf.add(GR, FormulaRule(formula=[f'AND($I8<>"",{X}>=$I8,{X}<=IF($J8<>"",$J8,StichtagAT))'], fill=fill("5FA84A"), stopIfTrue=True))
cf.add(GR, FormulaRule(formula=[f'AND($A8<>"",{X}>=$M8,{X}<=$N8,OR({X}<$F8,{X}>$G8))'], fill=fill(RED), stopIfTrue=True))
cf.add(GR, FormulaRule(formula=[f'AND($A8<>"",{X}>=$M8,{X}<=$N8)'], fill=fill("2E9BD6"), stopIfTrue=True))
cf.add(GR, FormulaRule(formula=[f'AND($A8<>"",{X}>=$F8,{X}<=$G8)'], fill=fill("CBD3DB"), stopIfTrue=True))
cf.add(f"{CL(G0)}6:{gl}7", FormulaRule(formula=[f"{CL(G0)}$6=Stichtag"], fill=fill(ORANGE), font=Font(name=F, color="FFFFFF", bold=True)))
STATUS_COL = [("Erledigt", "C6E7B9", "1E5631"), ("Erledigt, verspätet", "E2EFDA", "7F6000"), ("Läuft", "D6E6F5", NAVY), ("Läuft, Verzug", "FCE4D6", RED),
              ("Start überfällig", RED, "FFFFFF"), ("Verschoben", "FCE4D6", "C00000"), ("Geplant", "F2F2F2", "595959")]
for txt, bg, fg in STATUS_COL:
    cf.add(f"P8:P{LAST}", FormulaRule(formula=[f'$P8="{txt}"'], fill=fill(bg), font=Font(name=F, color=fg, bold=True)))
cf.add(f"O8:O{LAST}", FormulaRule(formula=['AND($A8<>"",$O8>0)'], font=Font(name=F, color=RED, bold=True), fill=fill("FDE2E4")))
cf.add(f"O8:O{LAST}", FormulaRule(formula=['AND($A8<>"",$O8<=0)'], font=Font(name=F, color="1E5631", bold=True)))
cf.add(f"M8:N{LAST}", FormulaRule(formula=['AND($A8<>"",M8>IF(COLUMN(M8)=13,$F8,$G8))'], font=Font(name=F, color=RED, bold=True)))
# Eingabeprüfungen
dvd = DataValidation(type="date", operator="between", formula1="DATE(2026,9,1)", formula2="DATE(2027,12,31)", allow_blank=True,
                     showErrorMessage=True, error="Bitte ein Datum eintragen (TT.MM.JJ).", promptTitle="Ist-Datum", prompt="Tatsächliches Datum eintragen. Leer lassen, solange nicht begonnen bzw. nicht fertig.", showInputMessage=True)
dvp = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True, showErrorMessage=True,
                     error="Wert zwischen 0 % und 100 %.", promptTitle="Fertig %", prompt="Nur für laufende Vorgänge. Leer = 50 % angenommen.", showInputMessage=True)
dvm = DataValidation(type="whole", operator="between", formula1="0", formula2="60", allow_blank=True, showErrorMessage=True,
                     error="Ganze Arbeitstage 0 bis 60.", promptTitle="Mehraufwand", prompt="Zusätzliche Arbeitstage, z. B. für Reparatur. Verschiebt alle Nachfolger.", showInputMessage=True)
dvu = DataValidation(type="list", formula1=f"=Listen!$C$2:$C${1+len(URS)}", allow_blank=True)
for dv_ in (dvd, dvp, dvm, dvu): P.add_data_validation(dv_)
for rr in data_rows:
    dvd.add(f"I{rr}:J{rr}"); dvp.add(f"K{rr}"); dvm.add(f"L{rr}"); dvu.add(f"Q{rr}")
P.freeze_panes = "C8"
P.auto_filter.ref = f"A7:R{LAST}"
P.page_setup.orientation = "landscape"; P.page_setup.paperSize = 8
P.page_setup.fitToWidth = 1; P.page_setup.fitToHeight = 0
P.sheet_properties.pageSetUpPr.fitToPage = True
P.print_title_rows = "4:7"
# Legende rechts oben im Kalender
lgx = G0
for t, c_ in [("Soll (Basisplan)", "CBD3DB"), ("Prognose im Soll", "2E9BD6"), ("Verschoben / Verzug", RED), ("Ist (begonnen/erledigt)", "5FA84A"), ("Stichtag", ORANGE)]:
    P.cell(2, lgx).fill = fill(c_); P.cell(2, lgx).border = BOX
    P.cell(2, lgx + 1, t).font = font(size=8)
    lgx += 9

# ---------------- Verantwortliche
wsV.sheet_view.showGridLines = False
wsV["A1"] = "Verantwortliche"; wsV["A1"].font = font(size=18, bold=True, color=NAVY)
wsV["A2"] = "Namen in die gelben Felder eintragen. Die Kennzahlen rechnen automatisch aus dem Terminplan."
wsV["A2"].font = font(color=GREY)
hdr = ["Rolle", "Zuständig für", "Name", "Vertretung", "Telefon", "Start überfällig", "Läuft mit Verzug", "Betroffen (verschoben)", "Handlungsbedarf"]
for j, h in enumerate(hdr, 1):
    c = wsV.cell(4, j, h); c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY); c.border = BOX
    c.alignment = Alignment(horizontal="center" if j > 2 else "left", vertical="center", wrap_text=True)
PL_ = f"Terminplan!$D$8:$D${LAST}"; ST_ = f"Terminplan!$P$8:$P${LAST}"; OV_ = f"Terminplan!$O$8:$O${LAST}"
for i, (k, inf) in enumerate(ROLE_INFO):
    rr = 5 + i
    for j, v in enumerate([R[k], inf, None, None, None], 1):
        c = wsV.cell(rr, j, v); c.font = font(bold=(j == 1)); c.border = BOX; c.alignment = LEFT
        if j >= 3: c.fill = fill(INPUT)
    wsV.cell(rr, 6, f'=COUNTIFS({PL_},$A{rr},{ST_},"Start überfällig")')
    wsV.cell(rr, 7, f'=COUNTIFS({PL_},$A{rr},{ST_},"Läuft, Verzug")')
    wsV.cell(rr, 8, f'=COUNTIFS({PL_},$A{rr},{ST_},"Verschoben")')
    wsV.cell(rr, 9, f'=IF(F{rr}+G{rr}=0,IF(H{rr}>0,"informieren","keiner"),"Ansprechen")')
    for j in range(6, 10):
        c = wsV.cell(rr, j); c.alignment = CENTER; c.border = BOX; c.font = font(bold=(j == 9))
lastV = 4 + len(ROLE_INFO)
wsV.conditional_formatting.add(f"I5:I{lastV}", FormulaRule(formula=['I5="Ansprechen"'], fill=fill(RED), font=Font(name=F, color="FFFFFF", bold=True)))
wsV.conditional_formatting.add(f"I5:I{lastV}", FormulaRule(formula=['I5="informieren"'], fill=fill("FFE699"), font=Font(name=F, color="7F6000", bold=True)))
wsV.conditional_formatting.add(f"I5:I{lastV}", FormulaRule(formula=['I5="keiner"'], fill=fill("C6E7B9"), font=Font(name=F, color="1E5631", bold=True)))
for col_, w in zip("ABCDEFGHI", [28, 46, 20, 18, 14, 12, 12, 12, 15]): wsV.column_dimensions[col_].width = w
wsV.row_dimensions[4].height = 30
wsV.page_setup.orientation = "landscape"; wsV.page_setup.paperSize = 9; wsV.page_setup.fitToWidth = 1; wsV.page_setup.fitToHeight = 1
wsV.sheet_properties.pageSetUpPr.fitToPage = True
dvr = DataValidation(type="list", formula1=f"=Verantwortliche!$A$5:$A${lastV}", allow_blank=False)
P.add_data_validation(dvr)
for rr in data_rows: dvr.add(f"D{rr}")

# ---------------- Übersicht
U = wsU
U.sheet_view.showGridLines = False
U.column_dimensions["A"].width = 2
for col_, w in zip("BCDEFGHIJKLM", [20, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 2]): U.column_dimensions[col_].width = w
U["B2"] = "Fertigung P800 · Stand auf einen Blick"; U["B2"].font = font(size=22, bold=True, color=NAVY)
U["B3"] = "Testschuss 1, Schüsse 301 und 302, Rundnaht 301+302 · einschichtig · Soll = Basisplan, Prognose = rechnet Verschiebungen automatisch durch"
U["B3"].font = font(size=10, color=GREY)
U["B5"] = "Stichtag"; U["B5"].font = font(bold=True, size=11)
U["C5"] = "=TODAY()"; U["C5"].number_format = "DD.MM.YYYY"; U["C5"].font = font(bold=True, size=12, color=NAVY); U["C5"].fill = fill(INPUT); U["C5"].alignment = CENTER; U["C5"].border = BOX
U["D5"] = "automatisch heute. Für einen anderen Tag ein Datum eintragen."; U["D5"].font = font(size=9, italic=True, color=GREY)
U["L5"] = "=WORKDAY(C5-1,1,Feiertage)"; U["L5"].number_format = "DD.MM.YY"; U["L5"].font = font(size=8, color="FFFFFF")
wb.defined_names["Stichtag"] = DefinedName("Stichtag", attr_text="'Übersicht'!$C$5")
wb.defined_names["StichtagAT"] = DefinedName("StichtagAT", attr_text="'Übersicht'!$L$5")

SOLL = f"Terminplan!$S$8:$S${LAST}"; IST = f"Terminplan!$T$8:$T${LAST}"; DUR = f"Terminplan!$H$8:$H${LAST}"; SEC = f"Terminplan!$U$8:$U${LAST}"
endrow = rowof["C08"]
# Großer Fortschrittsblock
U.merge_cells("B7:E7"); U["B7"] = "ABGEARBEITET (IST)"; U["B7"].font = font(bold=True, size=11, color="FFFFFF"); U["B7"].fill = fill(NAVY); U["B7"].alignment = CENTER
for k in "CDE": U[f"{k}7"].fill = fill(NAVY)
U.merge_cells("B8:E11"); U["B8"] = f"=IFERROR(SUM({IST})/SUM({DUR}),0)"; U["B8"].number_format = "0%"
U["B8"].font = font(bold=True, size=60, color=NAVY); U["B8"].alignment = CENTER
U.merge_cells("B12:E12"); U["B12"] = '="von "&TEXT(SUM(' + DUR + '),"0")&" geplanten Arbeitstagen"'; U["B12"].font = font(size=9, color=GREY); U["B12"].alignment = CENTER
tiles = [
    ("G", "SOLL HEUTE", f"=IFERROR(SUM({SOLL})/SUM({DUR}),0)", "0%", "laut Basisplan", "5E6B78"),
    ("I", "ABWEICHUNG", "=B8-G8", '+0%;-0%;0%', "Prozentpunkte Ist − Soll", ORANGE),
    ("K", "VERZUG ENDTERMIN", f"=Terminplan!O{endrow}", '+0" AT";-0" AT";"0 AT"', "Fertigmeldung Rundnaht", RED),
]
for col_, t, f_, nf, sub, c_ in tiles:
    c2 = CL(U[col_ + "1"].column + 1)
    U.merge_cells(f"{col_}7:{c2}7"); U[f"{col_}7"] = t; U[f"{col_}7"].font = font(bold=True, size=10, color="FFFFFF"); U[f"{col_}7"].alignment = CENTER
    U[f"{col_}7"].fill = fill(c_); U[f"{c2}7"].fill = fill(c_)
    U.merge_cells(f"{col_}8:{c2}10"); U[f"{col_}8"] = f_; U[f"{col_}8"].number_format = nf; U[f"{col_}8"].font = font(bold=True, size=30, color=c_); U[f"{col_}8"].alignment = CENTER
    U.merge_cells(f"{col_}11:{c2}12"); U[f"{col_}11"] = sub; U[f"{col_}11"].font = font(size=9, color=GREY); U[f"{col_}11"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    for rr in range(7, 13):
        for cc in (col_, c2): U[f"{cc}{rr}"].border = BOX
for rr in range(7, 13):
    for cc in "BCDE": U[f"{cc}{rr}"].border = BOX
U.row_dimensions[8].height = 28
# Fortschrittsbalken Soll vs Ist
U["B14"] = "Fortschritt Ist"; U["B14"].font = font(bold=True)
U["B15"] = "Fortschritt Soll"; U["B15"].font = font(bold=True)
U.merge_cells("C14:L14"); U["C14"] = "=B8"; U.merge_cells("C15:L15"); U["C15"] = "=G8"
for a_ in ("C14", "C15"): U[a_].number_format = "0%"; U[a_].font = font(bold=True); U[a_].alignment = Alignment(horizontal="left", vertical="center")
U.conditional_formatting.add("C14", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="2E9BD6", showValue=True))
U.conditional_formatting.add("C15", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="A6B1BC", showValue=True))
U.row_dimensions[14].height = 22; U.row_dimensions[15].height = 22
U["B16"] = "Wertung: erledigte Vorgänge 100 %, begonnene laut „Fertig %“ (leer = 50 %), Gewicht = Dauer in Arbeitstagen."
U["B16"].font = font(size=8, italic=True, color=GREY)

# Abschnitte
U["B18"] = "Soll-Ist je Abschnitt"; U["B18"].font = font(bold=True, size=13, color=NAVY)
h2 = ["Abschnitt", "Soll %", "Ist %", "Abw. Pkt", "Soll-Ende", "Prognose-Ende", "Verzug AT", "Start überfällig", "Lage"]
for j, h in enumerate(h2):
    c = U.cell(19, 2 + j, h); c.font = font(bold=True, color="FFFFFF", size=9); c.fill = fill(NAVY); c.border = BOX
    c.alignment = Alignment(horizontal="center" if j else "left", vertical="center", wrap_text=True)
U.merge_cells("J19:K19")
secs = [a[0] for a in A]
for i, name in enumerate(secs):
    rr = 20 + i
    last_task = sec_rows[name][-1]
    U.cell(rr, 2, name).font = font(bold=True, color=ADARK[name])
    U.cell(rr, 3, f'=IFERROR(SUMIFS({SOLL},{SEC},$B{rr})/SUMIFS({DUR},{SEC},$B{rr}),0)')
    U.cell(rr, 4, f'=IFERROR(SUMIFS({IST},{SEC},$B{rr})/SUMIFS({DUR},{SEC},$B{rr}),0)')
    U.cell(rr, 5, f"=D{rr}-C{rr}")
    U.cell(rr, 6, f"=_xlfn.MAXIFS(Terminplan!$G$8:$G${LAST},{SEC},$B{rr})")
    U.cell(rr, 7, f"=_xlfn.MAXIFS(Terminplan!$N$8:$N${LAST},{SEC},$B{rr})")
    U.cell(rr, 8, f"=IF(G{rr}>=F{rr},NETWORKDAYS(F{rr},G{rr},Feiertage)-1,-(NETWORKDAYS(G{rr},F{rr},Feiertage)-1))")
    U.cell(rr, 9, f'=COUNTIFS({SEC},$B{rr},Terminplan!$P$8:$P${LAST},"Start überfällig")')
    U.merge_cells(start_row=rr, start_column=10, end_row=rr, end_column=11)
    U.cell(rr, 10, f'=IF(D{rr}>=0.999,"abgeschlossen",IF(OR(H{rr}>0,I{rr}>0),"kritisch",IF(E{rr}<-0.05,"hinter Soll","im Plan")))')
    for j in range(2, 12):
        c = U.cell(rr, j); c.border = BOX
        if j > 2: c.alignment = CENTER; c.font = font(bold=(j in (4, 8, 10)))
    for j in (3, 4): U.cell(rr, j).number_format = "0%"
    U.cell(rr, 5).number_format = '+0%;-0%;0%'
    for j in (6, 7): U.cell(rr, j).number_format = "DD.MM.YY"
    U.cell(rr, 8).number_format = '+0;-0;0'
    U.row_dimensions[rr].height = 20
s1 = 19 + len(secs)
U.conditional_formatting.add(f"D20:D{s1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="2E9BD6", showValue=True))
U.conditional_formatting.add(f"C20:C{s1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="A6B1BC", showValue=True))
U.conditional_formatting.add(f"H20:H{s1}", FormulaRule(formula=["H20>0"], font=Font(name=F, color=RED, bold=True), fill=fill("FDE2E4")))
for txt, bg, fg in [("kritisch", RED, "FFFFFF"), ("hinter Soll", "FCE4D6", "C00000"), ("im Plan", "C6E7B9", "1E5631"), ("abgeschlossen", "D6E6F5", NAVY)]:
    U.conditional_formatting.add(f"J20:J{s1}", FormulaRule(formula=[f'$J20="{txt}"'], fill=fill(bg), font=Font(name=F, color=fg, bold=True)))

# Diagramm Soll vs Ist je Abschnitt
ch = BarChart(); ch.type = "bar"; ch.grouping = "clustered"
ch.title = "Soll und Ist je Abschnitt"; ch.style = 10
ch.add_data(Reference(U, min_col=3, max_col=4, min_row=19, max_row=s1), titles_from_data=True)
ch.set_categories(Reference(U, min_col=2, min_row=20, max_row=s1))
ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1; ch.y_axis.number_format = "0%"; ch.y_axis.majorGridlines = None
ch.x_axis.scaling.orientation = "maxMin"
ch.y_axis.delete = False; ch.x_axis.delete = False
ch.series[0].graphicalProperties.solidFill = "A6B1BC"; ch.series[1].graphicalProperties.solidFill = "2E9BD6"
for s_ in ch.series: s_.graphicalProperties.line.noFill = True
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True; ch.dataLabels.showSerName = False; ch.dataLabels.showCatName = False; ch.dataLabels.showLegendKey = False; ch.dataLabels.showPercent = False; ch.dataLabels.numFmt = "0%"
ch.legend.position = "b"; ch.height = 6.5; ch.width = 16; ch.gapWidth = 60
# Meilensteine
m0 = s1 + 3
U.cell(m0 - 1, 2, "Meilensteine: Soll, Prognose, Auswirkung").font = font(bold=True, size=13, color=NAVY)
h3 = ["ID", "Meilenstein", "", "", "Soll", "Prognose", "Verzug AT", "Status", "", "Verantwortlich", ""]
for a_, b_ in ((3, 5), (9, 10), (11, 12)): U.merge_cells(start_row=m0, start_column=a_, end_row=m0, end_column=b_)
for j, h in enumerate(h3):
    c = U.cell(m0, 2 + j)
    if h: c.value = h
    c.font = font(bold=True, color="FFFFFF", size=9); c.fill = fill(NAVY); c.alignment = CENTER; c.border = BOX
MS = ["F01", "F02", "F03", "B17", "A24", "B23", "C01", "C08"]
for i, mid in enumerate(MS):
    rr = m0 + 1 + i; pr = rowof[mid]
    U.cell(rr, 2, mid)
    U.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=5)
    U.cell(rr, 3, f"=Terminplan!B{pr}")
    U.cell(rr, 6, f"=Terminplan!G{pr}"); U.cell(rr, 7, f"=Terminplan!N{pr}")
    U.cell(rr, 8, f"=Terminplan!O{pr}")
    U.merge_cells(start_row=rr, start_column=9, end_row=rr, end_column=10); U.cell(rr, 9, f"=Terminplan!P{pr}")
    U.merge_cells(start_row=rr, start_column=11, end_row=rr, end_column=12); U.cell(rr, 11, f"=Terminplan!D{pr}")
    for j in range(2, 13):
        c = U.cell(rr, j); c.border = BOX; c.font = font(size=9, bold=(j in (2, 7, 8, 9))); c.alignment = LEFT if j in (3, 11) else CENTER
    for j in (6, 7): U.cell(rr, j).number_format = "DD.MM.YY"
    U.cell(rr, 8).number_format = '+0;-0;"–"'
m1 = m0 + len(MS)
U.conditional_formatting.add(f"H{m0+1}:H{m1}", FormulaRule(formula=[f"H{m0+1}>0"], font=Font(name=F, color=RED, bold=True), fill=fill("FDE2E4")))
U.conditional_formatting.add(f"G{m0+1}:G{m1}", FormulaRule(formula=[f"G{m0+1}>F{m0+1}"], font=Font(name=F, color=RED, bold=True)))
for txt, bg, fg in STATUS_COL:
    U.conditional_formatting.add(f"I{m0+1}:I{m1}", FormulaRule(formula=[f'$I{m0+1}="{txt}"'], fill=fill(bg), font=Font(name=F, color=fg, bold=True)))
# Wer ist gefragt
v0 = m1 + 3
U.cell(v0 - 1, 2, "Wer ist gefragt?").font = font(bold=True, size=13, color=NAVY)
h4 = ["Rolle", "", "Name", "", "Start überfällig", "Läuft mit Verzug", "Betroffen (verschoben)", "Handlungsbedarf", ""]
for a_, b_ in ((2, 3), (4, 5), (9, 10)): U.merge_cells(start_row=v0, start_column=a_, end_row=v0, end_column=b_)
for j, h in enumerate(h4):
    c = U.cell(v0, 2 + j)
    if h: c.value = h
    c.font = font(bold=True, color="FFFFFF", size=9); c.fill = fill(NAVY); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
for i in range(len(ROLE_INFO)):
    rr = v0 + 1 + i; srcr = 5 + i
    U.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3); U.cell(rr, 2, f"=Verantwortliche!A{srcr}")
    U.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=5); U.cell(rr, 4, f'=IF(Verantwortliche!C{srcr}="","(Name eintragen)",Verantwortliche!C{srcr})')
    U.cell(rr, 6, f"=Verantwortliche!F{srcr}"); U.cell(rr, 7, f"=Verantwortliche!G{srcr}"); U.cell(rr, 8, f"=Verantwortliche!H{srcr}")
    U.merge_cells(start_row=rr, start_column=9, end_row=rr, end_column=10); U.cell(rr, 9, f"=Verantwortliche!I{srcr}")
    for j in range(2, 11):
        c = U.cell(rr, j); c.border = BOX; c.font = font(size=9, bold=(j in (2, 9))); c.alignment = LEFT if j in (2, 4) else CENTER
    U.cell(rr, 4).font = font(size=9, italic=True, color=GREY)
v1 = v0 + len(ROLE_INFO)
U.conditional_formatting.add(f"I{v0+1}:I{v1}", FormulaRule(formula=[f'$I{v0+1}="Ansprechen"'], fill=fill(RED), font=Font(name=F, color="FFFFFF", bold=True)))
U.conditional_formatting.add(f"I{v0+1}:I{v1}", FormulaRule(formula=[f'$I{v0+1}="informieren"'], fill=fill("FFE699"), font=Font(name=F, color="7F6000", bold=True)))
U.conditional_formatting.add(f"I{v0+1}:I{v1}", FormulaRule(formula=[f'$I{v0+1}="keiner"'], fill=fill("C6E7B9"), font=Font(name=F, color="1E5631", bold=True)))
for cn in "FG":
    U.conditional_formatting.add(f"{cn}{v0+1}:{cn}{v1}", FormulaRule(formula=[f"{cn}{v0+1}>0"], font=Font(name=F, color=RED, bold=True)))
U.add_chart(ch, f"B{v1+3}")
U.page_setup.orientation = "portrait"; U.page_setup.paperSize = 9; U.page_setup.fitToWidth = 1; U.page_setup.fitToHeight = 1
U.sheet_properties.pageSetUpPr.fitToPage = True
U.print_area = f"A1:M{v1+18}"

# ---------------- Anleitung
N = wsA
N.sheet_view.showGridLines = False
for col_, w in zip("ABCD", [2, 6, 36, 92]): N.column_dimensions[col_].width = w
N["B2"] = "So funktioniert der Terminplan"; N["B2"].font = font(size=18, bold=True, color=NAVY)
N["B3"] = "Soll bleibt fest. Ist wird eingetragen. Die Prognose rechnet jede Verschiebung bis zum Endtermin durch."; N["B3"].font = font(size=11, color=GREY)
steps = [
    ("1", "Ist-Start eintragen", "Im Blatt „Terminplan“ in Spalte I das Datum eintragen, an dem der Vorgang tatsächlich begonnen hat."),
    ("2", "Ist-Ende eintragen", "In Spalte J das Datum eintragen, an dem der Vorgang fertig war. Bei laufenden Vorgängen optional „Fertig %“ (Spalte K)."),
    ("3", "Mehraufwand melden", "Dauert ein Vorgang länger, in Spalte L die zusätzlichen Arbeitstage eintragen. Alle Nachfolger verschieben sich sofort."),
    ("4", "Ursache und Maßnahme", "Bei Verzug in Spalte Q die Ursache wählen und in Spalte R die Maßnahme mit Kürzel des Verantwortlichen notieren."),
    ("5", "Übersicht lesen", "Seite 1 zeigt Fortschritt Ist gegen Soll, Verzug am Endtermin und die Meilensteine. „Ansprechen“ = eigener Vorgang überfällig oder im Verzug, „informieren“ = nur durch Vorgänger verschoben."),
]
for i, (n, h, t) in enumerate(steps):
    rr = 5 + i * 2
    c = N.cell(rr, 2, n); c.font = font(size=16, bold=True, color="FFFFFF"); c.fill = fill(ORANGE); c.alignment = CENTER
    N.cell(rr, 3, h).font = font(size=12, bold=True, color=NAVY); N.cell(rr, 3).alignment = Alignment(vertical="center")
    N.cell(rr, 4, t).font = font(size=10); N.cell(rr, 4).alignment = LEFT
    N.row_dimensions[rr].height = 32
r0 = 16
N.cell(r0, 2, "Wie die Verschiebung rechnet").font = font(size=12, bold=True, color=NAVY)
rules = [
    "Prognose-Start = spätester Wert aus Soll-Start, Prognose-Ende der Vorgänger (Spalte E) und heute, solange der Vorgang nicht begonnen hat.",
    "Prognose-Ende = Prognose-Start + Dauer + Mehraufwand. Ist ein Vorgang begonnen und noch nicht fertig, frühestens heute.",
    "Sobald Ist-Daten eingetragen sind, gelten diese. Ein Vorgang startet nie früher als im Soll; Puffer im Basisplan fangen Verzug auf.",
    "Verzug AT = Arbeitstage zwischen Soll-Ende und Prognose-Ende. Wochenenden und Betriebsruhe 24.12.–01.01. zählen nicht.",
]
for i, t in enumerate(rules):
    rr = r0 + 1 + i
    N.cell(rr, 2, "•").alignment = CENTER
    N.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
    N.cell(rr, 3, t).font = font(size=10); N.cell(rr, 3).alignment = LEFT; N.row_dimensions[rr].height = 26
r1 = r0 + 7
N.cell(r1, 2, "Farben im Balkenplan").font = font(size=12, bold=True, color=NAVY)
for i, (t, c_) in enumerate([("Soll: ursprünglicher Basisplan", "CBD3DB"), ("Prognose innerhalb des Soll-Zeitraums", "2E9BD6"), ("Prognose außerhalb des Soll: Verschiebung / Verzug", RED), ("Ist: begonnen bzw. erledigt", "5FA84A"), ("Stichtag (heute)", ORANGE)]):
    rr = r1 + 1 + i
    N.cell(rr, 2).fill = fill(c_); N.cell(rr, 2).border = BOX
    N.cell(rr, 3, t).font = font(size=10)
r2 = r1 + 8
N.cell(r2, 2, "Status je Vorgang").font = font(size=12, bold=True, color=NAVY)
expl = {"Erledigt": "Ist-Ende eingetragen, im Soll", "Erledigt, verspätet": "Ist-Ende nach Soll-Ende", "Läuft": "begonnen, Prognose im Soll",
        "Läuft, Verzug": "begonnen, Prognose-Ende nach Soll-Ende", "Start überfällig": "Soll-Start vorbei, kein Ist-Start: Rückmeldung oder Klärung nötig",
        "Verschoben": "noch nicht fällig, aber durch Vorgänger nach hinten geschoben", "Geplant": "läuft wie im Soll"}
for i, (t, bg, fg) in enumerate(STATUS_COL):
    rr = r2 + 1 + i
    N.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
    c = N.cell(rr, 2, t); c.fill = fill(bg); c.font = font(bold=True, color=fg); c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    N.cell(rr, 4, expl[t]).font = font(size=10)
r3 = r2 + 9
N.cell(r3, 2, "Hinweise").font = font(size=12, bold=True, color=NAVY)
notes = [
    "Soll-Daten: Blatt „1 Schichtig“ der Datei Kreutzchenplan_Schweißproben_PHPL-Wismar_Ablauf.xlsx.",
    "C04 bis C08 sind im Ursprungsplan nicht terminiert und aus den Dauern von Testschuss 1 abgeleitet (kursiv).",
    "Vorgänger sind fest verknüpft (Spalte E zeigt sie an). Station 12, 15, 16 und Messtrupp sind als Reihenfolge 301 vor 302 hinterlegt.",
    "Gelbe Felder sind Eingabefelder. Alles andere rechnet automatisch. Stichtag in der Übersicht steht auf heute.",
]
for i, t in enumerate(notes):
    rr = r3 + 1 + i
    N.cell(rr, 2, "•").alignment = CENTER
    N.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
    N.cell(rr, 3, t).font = font(size=10); N.cell(rr, 3).alignment = LEFT; N.row_dimensions[rr].height = 22
N.page_setup.orientation = "landscape"; N.page_setup.paperSize = 9; N.page_setup.fitToWidth = 1; N.page_setup.fitToHeight = 1
N.sheet_properties.pageSetUpPr.fitToPage = True

for ws, c_ in ((wsU, ORANGE), (wsP, NAVY), (wsV, GREEN), (wsA, GREY)): ws.sheet_properties.tabColor = c_
wb.active = 0
wb.save(OUT)
print("ok rows", LAST, "days", len(days), gl)
