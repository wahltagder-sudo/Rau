import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment

OUT = "/tmp/build/Fertigungsplan_P800_Abkreuzliste.xlsx"
F = "Arial"
NAVY, ORANGE, BLUE, GREEN, GREY, RED = "22303E", "E8641B", "2E9BD6", "5FA84A", "8A96A3", "C8102E"
LIGHT = "EEF1F4"

def fill(c): return PatternFill("solid", fgColor=c)
def font(**k): k.setdefault("name", F); k.setdefault("size", 10); return Font(**k)
thin = Side(style="thin", color="D5DAE0")
med = Side(style="medium", color="8A96A3")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

D = lambda s: dt.datetime.strptime(s, "%d.%m.%y")
HOL = [D(x) for x in ["24.12.26", "25.12.26", "28.12.26", "29.12.26", "30.12.26", "31.12.26", "01.01.27"]]
days = []
d = D("19.10.26")
while d <= D("19.02.27"):
    if d.weekday() < 5 and d not in HOL: days.append(d)
    d += dt.timedelta(1)

R = {
    "FS": "Fertigungssteuerung",
    "MO": "Meister Montage St. 12",
    "SW": "Meister Schweißen St. 15/16",
    "RN": "Meister Rundnaht St. 18",
    "SA": "Schweißaufsicht",
    "QS": "QS / NDT",
    "MT": "Messtrupp",
    "LO": "Logistik / Transport",
}
ROLE_INFO = [
    ("FS", "Termine, Abgaben, Eskalation, Gesamtstatus"),
    ("MO", "Umbau und Schussmontage Station 12"),
    ("SW", "Längsnähte St. 15, Spante St. 16, Putzplatz, Reparaturen"),
    ("RN", "Einrichten und Rundnaht Station 18"),
    ("SA", "WPS-Freigaben, Proben, Kontakt SLV Rostock"),
    ("QS", "VT, UT, RT, MT, Wartezeiten vor NDT"),
    ("MT", "Vermessen Halle 40, Messen, Spantabstand"),
    ("LO", "Transporte zwischen den Stationen"),
]

# Abschnitte: (Name, Farbe, [(ID, Station, Arbeitsschritt, Rolle, Start, Ende, Hinweis)])
A = [
    ("Freigaben", ORANGE, [
        ("F01", "SA", "WPS 138P/121 Station 15 freigegeben", "SA", "27.10.26", "27.10.26", "Voraussetzung für Längsnähte 301 ab 10.11."),
        ("F02", "SLV", "SLV-Ergebnis Testschuss 1 liegt vor", "SA", "20.11.26", "20.11.26", "Proben am 30.10. versandt"),
        ("F03", "SA", "WPS 121 Station 16 freigegeben", "SA", "03.12.26", "03.12.26", "Voraussetzung für Spante 301 ab 04.12."),
    ]),
    ("Testschuss 1", GREY, [
        ("T01", "Putzplatz", "Putzen Spante", "SW", "19.10.26", "21.10.26", ""),
        ("T02", "Putzplatz", "VT Schuss", "QS", "19.10.26", "20.10.26", ""),
        ("T03", "Putzplatz", "Wartezeit vor NDT", "QS", "19.10.26", "23.10.26", ""),
        ("T04", "St. 16", "Verfahrensprüfung Reparatur Spant 4", "SA", "22.10.26", "22.10.26", ""),
        ("T05", "Transport", "Transport auf Messplatz", "LO", "22.10.26", "22.10.26", ""),
        ("T06", "Putzplatz", "NDT UT/MT", "QS", "26.10.26", "27.10.26", ""),
        ("T07", "St. 16", "Probenentnahme Spant/DK", "SA", "28.10.26", "28.10.26", ""),
        ("T08", "St. 16", "Löcher der Probenentnahme schließen", "SW", "29.10.26", "29.10.26", ""),
        ("T09", "SLV", "Probenversand an SLV Rostock", "SA", "30.10.26", "30.10.26", "Spant 4 und 5"),
        ("T10", "St. 18", "Station 18 einrichten", "RN", "29.10.26", "30.10.26", ""),
        ("T11", "Transport", "Transport Schuss St. 16 auf St. 18", "LO", "02.11.26", "02.11.26", ""),
        ("T12", "St. 18", "Rundnaht verarbeiten", "RN", "02.11.26", "06.11.26", ""),
        ("T13", "St. 18", "Rundnaht messen (Spantabstand)", "MT", "09.11.26", "09.11.26", ""),
        ("T14", "St. 18", "Rundnaht innen schweißen", "RN", "10.11.26", "12.11.26", ""),
        ("T15", "St. 18", "Rundnaht außen fräsen inkl. NDT MT", "RN", "13.11.26", "13.11.26", ""),
        ("T16", "St. 18", "Rundnaht außen schweißen", "RN", "16.11.26", "18.11.26", ""),
        ("T17", "St. 18", "Reserve", "RN", "19.11.26", "19.11.26", ""),
        ("T18", "St. 18", "Decklage außen schleifen", "RN", "20.11.26", "20.11.26", ""),
    ]),
    ("Schuss 301", BLUE, [
        ("A01", "St. 12", "Umbau Vorrichtung P800", "MO", "19.10.26", "23.10.26", ""),
        ("A02", "St. 12", "Schussmontage 301", "MO", "26.10.26", "30.10.26", ""),
        ("A03", "Transport", "Transport Halle 40 zum Vermessen", "LO", "02.11.26", "02.11.26", ""),
        ("A04", "Halle 40", "Schuss vermessen", "MT", "03.11.26", "06.11.26", ""),
        ("A05", "Transport", "Transport auf St. 15", "LO", "09.11.26", "09.11.26", ""),
        ("A06", "St. 15", "Längsnähte innen schweißen (LN 1, LN 2)", "SW", "10.11.26", "13.11.26", "Benötigt F01"),
        ("A07", "St. 15", "Ausfugen inkl. MT", "SW", "16.11.26", "16.11.26", ""),
        ("A08", "St. 15", "Längsnähte außen schweißen (LN 1, LN 2)", "SW", "17.11.26", "20.11.26", ""),
        ("A09", "St. 15", "VT Längsnähte", "QS", "23.11.26", "23.11.26", ""),
        ("A10", "St. 15", "Wartezeit vor NDT", "QS", "23.11.26", "27.11.26", ""),
        ("A11", "St. 15", "Decklage schleifen", "SW", "24.11.26", "24.11.26", ""),
        ("A12", "St. 15", "Restanbindung LN", "SW", "25.11.26", "26.11.26", ""),
        ("A13", "St. 15", "NDT Längsnähte UT/RT/MT", "QS", "30.11.26", "30.11.26", ""),
        ("A14", "St. 15", "Filmausschnitte / Reparatur", "SW", "30.11.26", "01.12.26", ""),
        ("A15", "Transport", "Transport St. 15 auf St. 16", "LO", "03.12.26", "03.12.26", ""),
        ("A16", "St. 16", "Station 16 einrichten", "SW", "03.12.26", "03.12.26", ""),
        ("A17", "St. 16", "Spante schweißen", "SW", "04.12.26", "10.12.26", "Benötigt F03"),
        ("A18", "Transport", "Transport auf Putzplatz", "LO", "11.12.26", "11.12.26", ""),
        ("A19", "Putzplatz", "Putzen Spante", "SW", "14.12.26", "16.12.26", ""),
        ("A20", "Putzplatz", "VT Schuss", "QS", "17.12.26", "17.12.26", ""),
        ("A21", "Putzplatz", "Wartezeit vor NDT", "QS", "17.12.26", "23.12.26", ""),
        ("A22", "Putzplatz", "NDT UT/MT", "QS", "04.01.27", "05.01.27", ""),
        ("A23", "Messplatz", "Messen", "MT", "06.01.27", "11.01.27", ""),
        ("A24", "Messplatz", "Abgabe Schuss 301", "FS", "18.01.27", "18.01.27", "Meilenstein"),
    ]),
    ("Schuss 302", GREEN, [
        ("B01", "St. 12", "Umbau Vorrichtung 302", "MO", "02.11.26", "02.11.26", ""),
        ("B02", "St. 12", "Schussmontage 302 (03./04.11. zweite Schicht)", "MO", "03.11.26", "05.11.26", ""),
        ("B03", "Transport", "Transport Halle 40 zum Vermessen", "LO", "06.11.26", "06.11.26", ""),
        ("B04", "Halle 40", "Schuss vermessen", "MT", "09.11.26", "12.11.26", ""),
        ("B05", "Transport", "Transport auf St. 15", "LO", "16.11.26", "16.11.26", ""),
        ("B06", "St. 15", "Längsnähte innen schweißen (2 Schichten)", "SW", "17.11.26", "18.11.26", ""),
        ("B07", "St. 15", "Ausfugen inkl. MT", "SW", "23.11.26", "23.11.26", ""),
        ("B08", "St. 15", "Längsnähte außen schweißen (2 Schichten)", "SW", "24.11.26", "25.11.26", ""),
        ("B09", "St. 15", "VT Längsnähte", "QS", "26.11.26", "26.11.26", ""),
        ("B10", "St. 15", "Decklage schleifen", "SW", "27.11.26", "27.11.26", ""),
        ("B11", "St. 15", "Wartezeit vor NDT", "QS", "30.11.26", "04.12.26", ""),
        ("B12", "St. 15", "Restanbindung LN (2 Schichten)", "SW", "30.11.26", "30.11.26", ""),
        ("B13", "St. 15", "NDT Längsnähte UT/RT/MT", "QS", "07.12.26", "07.12.26", ""),
        ("B14", "St. 15", "Filmausschnitte / Reparatur (2 Schichten)", "SW", "07.12.26", "07.12.26", ""),
        ("B15", "Transport", "Transport St. 15 auf St. 16", "LO", "11.12.26", "11.12.26", ""),
        ("B16", "St. 16", "Station 16 einrichten", "SW", "11.12.26", "11.12.26", ""),
        ("B17", "St. 16", "Spante 1–6 schweißen (2 Schichten)", "SW", "14.12.26", "16.12.26", "Start nicht vor 14.12."),
        ("B18", "Transport", "Transport auf Putzplatz", "LO", "18.12.26", "18.12.26", ""),
        ("B19", "Putzplatz", "Putzen Spante", "SW", "21.12.26", "22.12.26", ""),
        ("B20", "Putzplatz", "VT Schuss", "QS", "23.12.26", "23.12.26", ""),
        ("B21", "Putzplatz", "NDT UT/MT", "QS", "06.01.27", "07.01.27", "Wartezeit über Betriebsruhe"),
        ("B22", "Messplatz", "Messen", "MT", "12.01.27", "15.01.27", ""),
        ("B23", "Messplatz", "Abgabe Schuss 302", "FS", "21.01.27", "21.01.27", "Meilenstein"),
    ]),
    ("Rundnaht 301+302", NAVY, [
        ("C01", "St. 18", "Station 18 einrichten, Schüsse 301+302", "RN", "22.01.27", "26.01.27", "Benötigt A24 und B23"),
        ("C02", "St. 18", "Rundnaht verarbeiten", "RN", "27.01.27", "01.02.27", ""),
        ("C03", "St. 18", "Rundnaht messen (Spantabstand)", "MT", "02.02.27", "02.02.27", ""),
        ("C04", "St. 18", "Rundnaht innen schweißen", "RN", "03.02.27", "05.02.27", "Planwert abgeleitet"),
        ("C05", "St. 18", "Rundnaht außen fräsen inkl. NDT MT", "RN", "08.02.27", "08.02.27", "Planwert abgeleitet"),
        ("C06", "St. 18", "Rundnaht außen schweißen", "RN", "09.02.27", "11.02.27", "Planwert abgeleitet"),
        ("C07", "St. 18", "Reserve", "RN", "12.02.27", "12.02.27", "Planwert abgeleitet"),
        ("C08", "St. 18", "Decklage außen schleifen, Fertigmeldung", "FS", "15.02.27", "15.02.27", "Planwert abgeleitet"),
    ]),
]

wb = Workbook()
wsA = wb.active; wsA.title = "Anleitung"
wsU = wb.create_sheet("Übersicht")
wsP = wb.create_sheet("Fertigungsplan")
wsD = wb.create_sheet("Abweichungen")
wsV = wb.create_sheet("Verantwortliche")
wsL = wb.create_sheet("Listen")

# ---------------- Listen (Hilfsdaten)
wsL["A1"] = "Betriebsruhe / Feiertage"; wsL["A1"].font = font(bold=True)
for i, h in enumerate(HOL): wsL.cell(2 + i, 1, h).number_format = "DD.MM.YYYY"
wsL["C1"] = "Ursachen"; wsL["C1"].font = font(bold=True)
URS = ["Personal", "Material", "Anlage / Vorrichtung", "Qualität / Nacharbeit", "Freigabe / WPS / SLV", "Vorgänger nicht fertig", "Transport", "Sonstiges"]
for i, u in enumerate(URS): wsL.cell(2 + i, 3, u)
wsL["E1"] = "Status Maßnahme"; wsL["E1"].font = font(bold=True)
for i, u in enumerate(["Offen", "In Arbeit", "Erledigt"]): wsL.cell(2 + i, 5, u)
wsL["G1"] = "Vorgangs-IDs"; wsL["G1"].font = font(bold=True)
ids = [t[0] for _, _, ts in A for t in ts]
for i, x in enumerate(ids): wsL.cell(2 + i, 7, x)
for c in "ACEG": wsL.column_dimensions[c].width = 26
wsL.sheet_state = "hidden"
wb.defined_names["Feiertage"] = DefinedName("Feiertage", attr_text=f"Listen!$A$2:$A${1+len(HOL)}")

# ---------------- Verantwortliche
wsV.sheet_view.showGridLines = False
wsV["A1"] = "Verantwortliche"; wsV["A1"].font = font(size=18, bold=True, color=NAVY)
wsV["A2"] = "Gelb hinterlegte Felder bitte mit Namen ausfüllen. Die Zahlen rechts aktualisieren sich automatisch aus Fertigungsplan und Abweichungen."
wsV["A2"].font = font(color=GREY)
hdr = ["Rolle", "Zuständig für", "Name", "Vertretung", "Telefon", "Rückmeldung fehlt (AT)", "Vorgänge mit Ausfall", "Offene Maßnahmen", "Handlungsbedarf"]
for j, h in enumerate(hdr, 1):
    c = wsV.cell(4, j, h); c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY); c.alignment = CENTER if j > 2 else LEFT; c.border = BOX
for i, (k, info) in enumerate(ROLE_INFO):
    r = 5 + i
    vals = [R[k], info, None, None, None]
    for j, v in enumerate(vals, 1):
        c = wsV.cell(r, j, v); c.font = font(bold=(j == 1)); c.border = BOX; c.alignment = LEFT
        if j in (3, 4, 5): c.fill = fill("FFF4CC")
    wsV.cell(r, 6, f"=SUMIFS(Fertigungsplan!$K:$K,Fertigungsplan!$E:$E,$A{r})")
    wsV.cell(r, 7, f'=COUNTIFS(Fertigungsplan!$E:$E,$A{r},Fertigungsplan!$J:$J,">0")')
    wsV.cell(r, 8, f'=COUNTIFS(Abweichungen!$K:$K,$A{r},Abweichungen!$M:$M,"<>Erledigt",Abweichungen!$A:$A,"<>Beispiel")')
    wsV.cell(r, 9, f'=IF(F{r}+G{r}+H{r}=0,"keiner","Ansprechen")')
    for j in range(6, 10):
        c = wsV.cell(r, j); c.font = font(bold=(j == 9)); c.alignment = CENTER; c.border = BOX
last_role = 4 + len(ROLE_INFO)
wsV.conditional_formatting.add(f"I5:I{last_role}", FormulaRule(formula=['I5="Ansprechen"'], fill=fill("FDE2E4"), font=Font(name=F, color=RED, bold=True)))
wsV.conditional_formatting.add(f"I5:I{last_role}", FormulaRule(formula=['I5="keiner"'], fill=fill("E3F1DC"), font=Font(name=F, color="2E6B1F")))
for col, w in zip("ABCDEFGHI", [28, 46, 20, 18, 14, 14, 14, 14, 15]): wsV.column_dimensions[col].width = w
wsV.row_dimensions[4].height = 32
for j in range(1, 10): wsV.cell(4, j).alignment = Alignment(horizontal="center" if j > 2 else "left", vertical="center", wrap_text=True)
wb.defined_names["Rollen"] = DefinedName("Rollen", attr_text=f"Verantwortliche!$A$5:$A${last_role}")
wsV.freeze_panes = "A5"

# ---------------- Fertigungsplan
P = wsP
P.sheet_view.showGridLines = False
P.sheet_view.zoomScale = 90
G0 = 13  # erste Kalenderspalte (M)
GL = G0 + len(days) - 1
gl = CL(GL)
P["A1"] = "Fertigungsplan P800 · Abkreuzliste je Arbeitstag"
P["A1"].font = font(size=16, bold=True, color=NAVY)
P["A2"] = '="Stichtag: "&TEXT(Stichtag,"DD.MM.YYYY")&"   ·   einschichtig, Station 12 bis 18, Halle 20+40"'
P["A2"].font = font(size=10, color=GREY)
# Legende Zeile 3/4
P["E3"] = "In die Tagesfelder eintragen:"; P["E3"].font = font(bold=True, size=9); P["E3"].alignment = Alignment(horizontal="right", vertical="center")
leg = [("x", "erledigt", "C6E7B9", "1E5631"), ("t", "teilweise", "FFE699", "7F6000"), ("o", "nicht gel.", RED, "FFFFFF"), ("", "geplant", "D6E6F5", NAVY), ("", "fehlt", "F8CBAD", "833C0B")]
for k, (sym, txt, bg, fg) in enumerate(leg):
    cc = 6 + k + (1 if k >= 3 else 0)
    c = P.cell(3, cc, sym); c.fill = fill(bg); c.font = font(bold=True, color=fg); c.alignment = CENTER; c.border = BOX
    t = P.cell(4, cc, txt); t.font = font(size=8, color=GREY); t.alignment = CENTER
P["I3"] = "Farben:"; P["I3"].font = font(bold=True, size=9); P["I3"].alignment = Alignment(horizontal="right", vertical="center")
# Kopf
heads = ["ID", "Abschnitt", "Station", "Arbeitsschritt", "Verantwortlich", "Start", "Ende", "Plan AT", "Erledigt AT", "Nicht gelaufen AT", "Rückmeldung fehlt AT", "Status"]
for j, h in enumerate(heads, 1):
    P.merge_cells(start_row=5, start_column=j, end_row=7, end_column=j)
    c = P.cell(5, j, h); c.font = font(bold=True, color="FFFFFF", size=9); c.fill = fill(NAVY)
    c.alignment = Alignment(horizontal="center" if j != 4 else "left", vertical="center", wrap_text=True)
for j in range(1, 13):
    for r in (5, 6, 7): P.cell(r, j).fill = fill(NAVY); P.cell(r, j).border = BOX
WD = ["Mo", "Di", "Mi", "Do", "Fr"]
wk_start = G0
for i, dd in enumerate(days):
    cidx = G0 + i
    c6 = P.cell(6, cidx, dd); c6.number_format = "DD.MM."; c6.font = font(size=8, bold=True, color=NAVY)
    c6.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90); c6.fill = fill(LIGHT); c6.border = BOX
    c7 = P.cell(7, cidx, WD[dd.weekday()]); c7.font = font(size=8, color=GREY); c7.alignment = CENTER; c7.fill = fill(LIGHT); c7.border = BOX
    P.column_dimensions[CL(cidx)].width = 3.4
    nxt = days[i + 1] if i + 1 < len(days) else None
    if nxt is None or nxt.isocalendar()[1] != dd.isocalendar()[1]:
        P.merge_cells(start_row=5, start_column=wk_start, end_row=5, end_column=cidx)
        wf = fill(ORANGE if dd.isocalendar()[1] % 2 else "3A4B5C")
        c5 = P.cell(5, wk_start, f"KW {dd.isocalendar()[1]}"); c5.font = font(size=9, bold=True, color="FFFFFF"); c5.alignment = CENTER
        for cc in range(wk_start, cidx + 1): P.cell(5, cc).fill = wf
        wk_start = cidx + 1
P.row_dimensions[6].height = 38
for col_, w in zip("ABCDEFGHIJKL", [6, 17, 10, 40, 24, 9, 9, 7, 8, 9, 12, 17]): P.column_dimensions[col_].width = w

r = 8
data_rows = []
first_cols = {}
for name, colr, tasks in A:
    P.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    c = P.cell(r, 1, name.upper()); c.font = font(bold=True, color="FFFFFF", size=10); c.fill = fill(colr); c.alignment = Alignment(vertical="center", indent=1)
    for cc in range(1, 13): P.cell(r, cc).fill = fill(colr)
    for cc in range(G0, GL + 1): P.cell(r, cc).fill = fill(colr)
    P.row_dimensions[r].height = 18
    r += 1
    for (tid, st, step, role, s, e, hint) in tasks:
        vals = [tid, name, st, step, R[role], D(s), D(e)]
        for j, v in enumerate(vals, 1):
            c = P.cell(r, j, v); c.font = font(size=9, bold=(j == 1)); c.border = BOX
            c.alignment = CENTER if j in (1, 3, 6, 7) else LEFT
            if j in (6, 7): c.number_format = "DD.MM.YY"
        if hint:
            P.cell(r, 4).comment = Comment(hint, "Fertigungssteuerung")
            if "abgeleitet" in hint: P.cell(r, 4).font = font(size=9, italic=True)
            if "Meilenstein" in hint: P.cell(r, 4).font = font(size=9, bold=True)
        rng = f"${CL(G0)}{r}:${gl}{r}"
        hdr6 = f"${CL(G0)}$6:${gl}$6"
        P.cell(r, 8, f"=NETWORKDAYS(F{r},G{r},Feiertage)")
        P.cell(r, 9, f'=COUNTIF({rng},"x")+0.5*COUNTIF({rng},"t")')
        P.cell(r, 10, f'=COUNTIF({rng},"o")')
        P.cell(r, 11, f'=SUMPRODUCT(({hdr6}>=$F{r})*({hdr6}<=$G{r})*({hdr6}<Stichtag)*({rng}=""))')
        P.cell(r, 12, f'=IF(I{r}>=H{r},"Erledigt",IF(J{r}>0,"Ausfall",IF(K{r}>0,"Rückmeldung fehlt",IF(Stichtag>G{r},"Überfällig",IF(Stichtag>=F{r},"Läuft","Geplant")))))')
        for j in range(8, 13):
            c = P.cell(r, j); c.font = font(size=9, bold=(j == 12)); c.alignment = CENTER; c.border = BOX
        P.cell(r, 9).number_format = "0.0;-0.0;-"
        for j in (10, 11): P.cell(r, j).number_format = "0;-0;-"
        for cc in range(G0, GL + 1):
            c = P.cell(r, cc); c.border = BOX; c.alignment = CENTER; c.font = font(size=9, bold=True)
        data_rows.append(r)
        r += 1
LAST = r - 1
# Wochenlinien
for i, dd in enumerate(days):
    if i == 0 or dd.isocalendar()[1] != days[i - 1].isocalendar()[1]:
        for rr in range(6, LAST + 1):
            c = P.cell(rr, G0 + i)
            c.border = Border(left=med, right=c.border.right, top=c.border.top, bottom=c.border.bottom)
GR = f"{CL(G0)}8:{gl}{LAST}"
tl = f"{CL(G0)}8"
cf = P.conditional_formatting
cf.add(GR, FormulaRule(formula=[f'LOWER({tl})="x"'], fill=fill("C6E7B9"), font=Font(name=F, color="1E5631", bold=True), stopIfTrue=True))
cf.add(GR, FormulaRule(formula=[f'LOWER({tl})="t"'], fill=fill("FFE699"), font=Font(name=F, color="7F6000", bold=True), stopIfTrue=True))
cf.add(GR, FormulaRule(formula=[f'LOWER({tl})="o"'], fill=fill(RED), font=Font(name=F, color="FFFFFF", bold=True), stopIfTrue=True))
cf.add(GR, FormulaRule(formula=[f'AND($F8<>"",{CL(G0)}$6>=$F8,{CL(G0)}$6<=$G8,{CL(G0)}$6<Stichtag)'], fill=fill("F8CBAD"), stopIfTrue=True))
cf.add(GR, FormulaRule(formula=[f'AND($F8<>"",{CL(G0)}$6>=$F8,{CL(G0)}$6<=$G8)'], fill=fill("D6E6F5"), stopIfTrue=True))
cf.add(f"{CL(G0)}6:{gl}7", FormulaRule(formula=[f'{CL(G0)}$6=Stichtag'], fill=fill(ORANGE), font=Font(name=F, color="FFFFFF", bold=True)))
st = f"L8:L{LAST}"
for txt, bg, fg in [("Erledigt", "C6E7B9", "1E5631"), ("Ausfall", RED, "FFFFFF"), ("Rückmeldung fehlt", "F8CBAD", "833C0B"), ("Überfällig", "FDE2E4", RED), ("Läuft", "D6E6F5", NAVY), ("Geplant", "F2F2F2", "595959")]:
    cf.add(st, FormulaRule(formula=[f'$L8="{txt}"'], fill=fill(bg), font=Font(name=F, color=fg, bold=True)))
cf.add(f"J8:J{LAST}", FormulaRule(formula=["$J8>0"], font=Font(name=F, color=RED, bold=True)))
cf.add(f"K8:K{LAST}", FormulaRule(formula=["$K8>0"], font=Font(name=F, color="833C0B", bold=True)))
dv = DataValidation(type="list", formula1='"x,t,o"', allow_blank=True, showErrorMessage=True,
                    errorTitle="Ungültige Eingabe", error="Bitte nur x (erledigt), t (teilweise) oder o (nicht gelaufen) eintragen.",
                    promptTitle="Rückmeldung", prompt="x = erledigt · t = teilweise · o = nicht gelaufen", showInputMessage=True)
P.add_data_validation(dv)
for rr in data_rows: dv.add(f"{CL(G0)}{rr}:{gl}{rr}")
dvr = DataValidation(type="list", formula1="=Rollen", allow_blank=False)
P.add_data_validation(dvr)
for rr in data_rows: dvr.add(f"E{rr}")
P.freeze_panes = P.cell(8, G0)
P.auto_filter.ref = f"A7:L{LAST}"
P.print_options.gridLines = False
P.page_setup.orientation = "landscape"; P.page_setup.paperSize = 8
P.page_setup.fitToWidth = 1; P.page_setup.fitToHeight = 0
P.sheet_properties.pageSetUpPr.fitToPage = True
P.print_title_rows = "5:7"

# ---------------- Abweichungen
W = wsD
W.sheet_view.showGridLines = False
W["A1"] = "Abweichungen und Maßnahmen"; W["A1"].font = font(size=16, bold=True, color=NAVY)
W["A2"] = "Jedes „o“ und jedes „t“ im Fertigungsplan hier erfassen: was ist nicht gelaufen, warum, wer kümmert sich bis wann. Gelbe Felder ausfüllen, graue Felder rechnen automatisch."
W["A2"].font = font(color=GREY)
hdr = ["Nr", "Datum", "Vorgangs-ID", "Abschnitt", "Arbeitsschritt", "Verantwortlich Vorgang", "Was ist nicht gelaufen?", "Ursache", "Verlust AT", "Maßnahme", "Maßnahme verantwortlich", "Fällig bis", "Status", "Ampel"]
for j, h in enumerate(hdr, 1):
    c = W.cell(4, j, h); c.font = font(bold=True, color="FFFFFF", size=9); c.fill = fill(NAVY); c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
W.row_dimensions[4].height = 30
NROW = 60
for i in range(NROW):
    rr = 5 + i
    for j in range(1, 15):
        c = W.cell(rr, j); c.border = BOX; c.font = font(size=9)
        c.alignment = LEFT if j in (5, 6, 7, 10, 11) else CENTER
        if j in (2, 3, 7, 8, 9, 10, 11, 12, 13): c.fill = fill("FFFBEA")
        if j in (4, 5, 6, 14): c.fill = fill("F2F2F2")
    W.cell(rr, 1, i if i else "Beispiel")
    W.cell(rr, 2).number_format = "DD.MM.YY"; W.cell(rr, 12).number_format = "DD.MM.YY"
    W.cell(rr, 4, f'=IF($C{rr}="","",IFERROR(INDEX(Fertigungsplan!$B:$B,MATCH($C{rr},Fertigungsplan!$A:$A,0)),"ID unbekannt"))')
    W.cell(rr, 5, f'=IF($C{rr}="","",IFERROR(INDEX(Fertigungsplan!$D:$D,MATCH($C{rr},Fertigungsplan!$A:$A,0)),""))')
    W.cell(rr, 6, f'=IF($C{rr}="","",IFERROR(INDEX(Fertigungsplan!$E:$E,MATCH($C{rr},Fertigungsplan!$A:$A,0)),""))')
    W.cell(rr, 14, f'=IF($C{rr}="","",IF($M{rr}="Erledigt","erledigt",IF(AND($L{rr}<>"",$L{rr}<Stichtag),"eskalieren",IF($M{rr}="In Arbeit","in Arbeit","offen"))))')
# Beispielzeile
ex = {2: D("13.11.26"), 3: "A06", 7: "LN 2 innen am 13.11. nicht geschweißt", 8: "Personal", 9: 1, 10: "LN 2 am 16.11. nachholen, Ausfugen auf 17.11.", 11: R["SW"], 12: D("17.11.26"), 13: "In Arbeit"}
for j, v in ex.items(): W.cell(5, j, v)
for j in range(1, 15): W.cell(5, j).font = font(size=9, italic=True, color="7F7F7F")
W.cell(5, 1).comment = Comment("Beispielzeile zur Orientierung. Wird in keiner Auswertung gezählt. Kann gelöscht oder überschrieben werden.", "Fertigungssteuerung")
dvi = DataValidation(type="list", formula1=f"=Listen!$G$2:$G${1+len(ids)}", allow_blank=True); W.add_data_validation(dvi); dvi.add(f"C5:C{4+NROW}")
dvu = DataValidation(type="list", formula1=f"=Listen!$C$2:$C${1+len(URS)}", allow_blank=True); W.add_data_validation(dvu); dvu.add(f"H5:H{4+NROW}")
dvk = DataValidation(type="list", formula1="=Rollen", allow_blank=True); W.add_data_validation(dvk); dvk.add(f"K5:K{4+NROW}")
dvs = DataValidation(type="list", formula1="=Listen!$E$2:$E$4", allow_blank=True); W.add_data_validation(dvs); dvs.add(f"M5:M{4+NROW}")
amp = f"N5:N{4+NROW}"
for txt, bg, fg in [("eskalieren", RED, "FFFFFF"), ("offen", "F8CBAD", "833C0B"), ("in Arbeit", "FFE699", "7F6000"), ("erledigt", "C6E7B9", "1E5631")]:
    W.conditional_formatting.add(amp, FormulaRule(formula=[f'$N5="{txt}"'], fill=fill(bg), font=Font(name=F, color=fg, bold=True)))
for col_, w in zip("ABCDEFGHIJKLMN", [8, 10, 10, 15, 32, 24, 34, 20, 8, 36, 24, 10, 11, 11]): W.column_dimensions[col_].width = w
W.freeze_panes = "A5"
W.auto_filter.ref = f"A4:N{4+NROW}"
W.page_setup.orientation = "landscape"; W.page_setup.paperSize = 9; W.page_setup.fitToWidth = 1; W.page_setup.fitToHeight = 0
W.sheet_properties.pageSetUpPr.fitToPage = True
AB = f"Abweichungen!$A$5:$A${4+NROW}"

# ---------------- Übersicht
U = wsU
U.sheet_view.showGridLines = False
U["B2"] = "Fertigungsplan P800 · Übersicht"; U["B2"].font = font(size=20, bold=True, color=NAVY)
U["B3"] = "Testschuss 1, Schüsse 301 und 302, Rundnaht 301+302 · einschichtig"; U["B3"].font = font(size=11, color=GREY)
U["B5"] = "Stichtag"; U["B5"].font = font(bold=True, size=11)
U["C5"] = "=TODAY()"; U["C5"].number_format = "DD.MM.YYYY"; U["C5"].font = font(bold=True, size=12, color=NAVY); U["C5"].fill = fill("FFF4CC"); U["C5"].alignment = CENTER; U["C5"].border = BOX
U["D5"] = "automatisch heute. Für eine Auswertung zu einem anderen Tag Datum eintragen."; U["D5"].font = font(size=9, italic=True, color=GREY)
wb.defined_names["Stichtag"] = DefinedName("Stichtag", attr_text="'Übersicht'!$C$5")

PL = f"Fertigungsplan!$H$8:$H${LAST}"; ER = f"Fertigungsplan!$I$8:$I${LAST}"; NG = f"Fertigungsplan!$J$8:$J${LAST}"; RF = f"Fertigungsplan!$K$8:$K${LAST}"; STc = f"Fertigungsplan!$L$8:$L${LAST}"; ABc = f"Fertigungsplan!$B$8:$B${LAST}"
kp = [
    ("Fortschritt", f"=IFERROR(SUM({ER})/SUM({PL}),0)", "0%", "erledigte von geplanten Arbeitstagen", NAVY),
    ("Vorgänge erledigt", f'=COUNTIF({STc},"Erledigt")&" / "&COUNTA(Fertigungsplan!$A$8:$A${LAST})-{len(A)}', "@", "abgeschlossene Arbeitsschritte", GREEN),
    ("Nicht gelaufen", f"=SUM({NG})", "0", "Arbeitstage mit „o“", RED),
    ("Rückmeldung fehlt", f"=SUM({RF})", "0", "geplante Tage ohne Eintrag", ORANGE),
    ("Offene Maßnahmen", f'=COUNTIFS(Abweichungen!$C$5:$C${4+NROW},"<>",Abweichungen!$M$5:$M${4+NROW},"<>Erledigt",{AB},"<>Beispiel")', "0", "aus Blatt Abweichungen", BLUE),
]
for i, (t, f_, nf, sub, colr) in enumerate(kp):
    c0 = 2 + i * 2
    U.merge_cells(start_row=7, start_column=c0, end_row=7, end_column=c0 + 1)
    U.merge_cells(start_row=8, start_column=c0, end_row=8, end_column=c0 + 1)
    U.merge_cells(start_row=9, start_column=c0, end_row=9, end_column=c0 + 1)
    a = U.cell(7, c0, t); a.font = font(bold=True, size=10, color="FFFFFF"); a.fill = fill(colr); a.alignment = CENTER
    b = U.cell(8, c0, f_); b.font = font(bold=True, size=24, color=colr); b.alignment = CENTER; b.number_format = nf
    s = U.cell(9, c0, sub); s.font = font(size=8, color=GREY); s.alignment = CENTER
    for rr in (7, 8, 9):
        for cc in (c0, c0 + 1):
            U.cell(rr, cc).border = BOX
            if rr > 7: U.cell(rr, cc).fill = fill("FFFFFF" if rr == 8 else LIGHT)
            else: U.cell(rr, cc).fill = fill(colr)
U.row_dimensions[8].height = 44

# Abschnitte
U["B12"] = "Stand je Abschnitt"; U["B12"].font = font(bold=True, size=13, color=NAVY)
h2 = ["Abschnitt", "Plan AT", "Erledigt AT", "Fortschritt", "Nicht gelaufen", "Rückmeldung fehlt", "Überfällige Vorgänge", "Plan-Ende", "Lage"]
for j, h in enumerate(h2):
    c = U.cell(13, 2 + j, h); c.font = font(bold=True, color="FFFFFF", size=9); c.fill = fill(NAVY); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
U.row_dimensions[13].height = 30
for i, (name, colr, _) in enumerate(A):
    rr = 14 + i
    U.cell(rr, 2, name).font = font(bold=True, color=colr if colr != GREY else "595959")
    U.cell(rr, 3, f'=SUMIFS({PL},{ABc},$B{rr})')
    U.cell(rr, 4, f'=SUMIFS({ER},{ABc},$B{rr})'); U.cell(rr, 4).number_format = "0.0"
    U.cell(rr, 5, f'=IF(C{rr}=0,0,D{rr}/C{rr})'); U.cell(rr, 5).number_format = "0%"
    U.cell(rr, 6, f'=SUMIFS({NG},{ABc},$B{rr})')
    U.cell(rr, 7, f'=SUMIFS({RF},{ABc},$B{rr})')
    U.cell(rr, 8, f'=COUNTIFS({ABc},$B{rr},{STc},"Überfällig")+COUNTIFS({ABc},$B{rr},{STc},"Ausfall")')
    U.cell(rr, 9, f'=_xlfn.MAXIFS(Fertigungsplan!$G$8:$G${LAST},{ABc},$B{rr})'); U.cell(rr, 9).number_format = "DD.MM.YY"
    U.cell(rr, 10, f'=IF(D{rr}>=C{rr},"abgeschlossen",IF(F{rr}+H{rr}>0,"kritisch",IF(G{rr}>0,"Rückmeldung fehlt","im Plan")))')
    for j in range(2, 11):
        c = U.cell(rr, j); c.border = BOX; c.alignment = CENTER if j > 2 else LEFT
        if j > 2: c.font = font(bold=(j == 10))
    U.row_dimensions[rr].height = 22
e2 = 13 + len(A)
U.conditional_formatting.add(f"E14:E{e2}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=GREEN, showValue=True))
for txt, bg, fg in [("kritisch", RED, "FFFFFF"), ("Rückmeldung fehlt", "F8CBAD", "833C0B"), ("im Plan", "C6E7B9", "1E5631"), ("abgeschlossen", "D6E6F5", NAVY)]:
    U.conditional_formatting.add(f"J14:J{e2}", FormulaRule(formula=[f'$J14="{txt}"'], fill=fill(bg), font=Font(name=F, color=fg, bold=True)))

# Meilensteine
m0 = e2 + 3
U.cell(m0 - 1, 2, "Meilensteine und Freigaben").font = font(bold=True, size=13, color=NAVY)
h3 = ["ID", "Meilenstein", "", "", "Plan", "Verantwortlich", "", "Status", ""]
for a_, b_ in ((3, 5), (7, 8), (9, 10)): U.merge_cells(start_row=m0, start_column=a_, end_row=m0, end_column=b_)
for j, h in enumerate(h3):
    c = U.cell(m0, 2 + j)
    if h: c.value = h
    c.font = font(bold=True, color="FFFFFF", size=9); c.fill = fill(NAVY); c.alignment = CENTER; c.border = BOX
MS = ["F01", "F02", "F03", "B17", "A24", "B23", "C01", "C08"]
for i, mid in enumerate(MS):
    rr = m0 + 1 + i
    U.cell(rr, 2, mid)
    U.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=5)
    U.cell(rr, 3, f'=INDEX(Fertigungsplan!$D:$D,MATCH($B{rr},Fertigungsplan!$A:$A,0))')
    U.cell(rr, 6, f'=INDEX(Fertigungsplan!$F:$F,MATCH($B{rr},Fertigungsplan!$A:$A,0))'); U.cell(rr, 6).number_format = "DD.MM.YY"
    U.merge_cells(start_row=rr, start_column=7, end_row=rr, end_column=8)
    U.cell(rr, 7, f'=INDEX(Fertigungsplan!$E:$E,MATCH($B{rr},Fertigungsplan!$A:$A,0))')
    U.merge_cells(start_row=rr, start_column=9, end_row=rr, end_column=10)
    U.cell(rr, 9, f'=INDEX(Fertigungsplan!$L:$L,MATCH($B{rr},Fertigungsplan!$A:$A,0))')
    for j in range(2, 11):
        c = U.cell(rr, j); c.border = BOX; c.font = font(bold=(j in (2, 9))); c.alignment = LEFT if j in (3, 7) else CENTER
m1 = m0 + len(MS)
for txt, bg, fg in [("Erledigt", "C6E7B9", "1E5631"), ("Ausfall", RED, "FFFFFF"), ("Rückmeldung fehlt", "F8CBAD", "833C0B"), ("Überfällig", "FDE2E4", RED), ("Läuft", "D6E6F5", NAVY), ("Geplant", "F2F2F2", "595959")]:
    U.conditional_formatting.add(f"I{m0+1}:I{m1}", FormulaRule(formula=[f'$I{m0+1}="{txt}"'], fill=fill(bg), font=Font(name=F, color=fg, bold=True)))

# Verantwortliche kompakt
v0 = m1 + 3
U.cell(v0 - 1, 2, "Wer ist gefragt?").font = font(bold=True, size=13, color=NAVY)
h4 = ["Rolle", "", "Name", "", "Rückmeldung fehlt (AT)", "Vorgänge mit Ausfall", "Offene Maßnahmen", "Handlungsbedarf", ""]
for a_, b_ in ((2, 3), (4, 5), (9, 10)): U.merge_cells(start_row=v0, start_column=a_, end_row=v0, end_column=b_)
for j, h in enumerate(h4):
    c = U.cell(v0, 2 + j)
    if h: c.value = h
    c.font = font(bold=True, color="FFFFFF", size=9); c.fill = fill(NAVY); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
U.row_dimensions[v0].height = 30
for i in range(len(ROLE_INFO)):
    rr = v0 + 1 + i; src = 5 + i
    U.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
    U.cell(rr, 2, f"=Verantwortliche!A{src}")
    U.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=5)
    U.cell(rr, 4, f'=IF(Verantwortliche!C{src}="","(Name eintragen)",Verantwortliche!C{src})')
    U.cell(rr, 6, f"=Verantwortliche!F{src}"); U.cell(rr, 7, f"=Verantwortliche!G{src}"); U.cell(rr, 8, f"=Verantwortliche!H{src}")
    U.merge_cells(start_row=rr, start_column=9, end_row=rr, end_column=10)
    U.cell(rr, 9, f"=Verantwortliche!I{src}")
    for j in range(2, 11):
        c = U.cell(rr, j); c.border = BOX; c.alignment = LEFT if j in (2, 4) else CENTER; c.font = font(bold=(j in (2, 9)))
    U.cell(rr, 4).font = font(color=GREY, italic=True)
v1 = v0 + len(ROLE_INFO)
U.conditional_formatting.add(f"I{v0+1}:I{v1}", FormulaRule(formula=[f'$I{v0+1}="Ansprechen"'], fill=fill(RED), font=Font(name=F, color="FFFFFF", bold=True)))
U.conditional_formatting.add(f"I{v0+1}:I{v1}", FormulaRule(formula=[f'$I{v0+1}="keiner"'], fill=fill("C6E7B9"), font=Font(name=F, color="1E5631", bold=True)))
for cnum in range(6, 9):
    U.conditional_formatting.add(f"{CL(cnum)}{v0+1}:{CL(cnum)}{v1}", FormulaRule(formula=[f'{CL(cnum)}{v0+1}>0'], font=Font(name=F, color=RED, bold=True)))
U.column_dimensions["A"].width = 2
for col_, w in zip("BCDEFGHIJK", [18, 14, 14, 14, 14, 14, 14, 14, 18, 4]): U.column_dimensions[col_].width = w
U.page_setup.orientation = "portrait"; U.page_setup.paperSize = 9; U.page_setup.fitToWidth = 1; U.page_setup.fitToHeight = 1
U.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- Anleitung
N = wsA
N.sheet_view.showGridLines = False
N.column_dimensions["A"].width = 2; N.column_dimensions["B"].width = 6; N.column_dimensions["C"].width = 36; N.column_dimensions["D"].width = 90
N["B2"] = "Fertigungsplan P800 · So funktioniert die Abkreuzliste"; N["B2"].font = font(size=18, bold=True, color=NAVY)
N["B3"] = "Ein Plan, eine Wahrheit: Jeder Arbeitstag wird abgehakt. Was nicht gelaufen ist, bekommt einen Verantwortlichen und eine Maßnahme."
N["B3"].font = font(size=11, color=GREY)
steps = [
    ("1", "Täglich abkreuzen", "Im Blatt „Fertigungsplan“ in der Zeile des Arbeitsschritts und der Spalte des Tages eintragen: x = erledigt, t = teilweise, o = nicht gelaufen. Blau hinterlegte Felder sind geplant."),
    ("2", "Rote und orange Felder klären", "Orange = geplanter Tag ohne Rückmeldung. Rot = nicht gelaufen. Beides im Blatt „Abweichungen“ erfassen: was, warum, wer, bis wann."),
    ("3", "Verantwortliche ansprechen", "Im Blatt „Übersicht“ steht je Rolle, wie viele Rückmeldungen fehlen und wie viele Maßnahmen offen sind. „Ansprechen“ heißt: heute klären."),
    ("4", "Wöchentlich berichten", "Die „Übersicht“ zeigt Fortschritt, kritische Abschnitte und Meilensteine. Sie passt auf eine Seite und kann so ausgedruckt werden."),
]
for i, (n, h, t) in enumerate(steps):
    rr = 5 + i * 2
    c = N.cell(rr, 2, n); c.font = font(size=16, bold=True, color="FFFFFF"); c.fill = fill(ORANGE); c.alignment = CENTER
    N.cell(rr, 3, h).font = font(size=12, bold=True, color=NAVY); N.cell(rr, 3).alignment = Alignment(vertical="center")
    N.cell(rr, 4, t).font = font(size=10); N.cell(rr, 4).alignment = LEFT
    N.row_dimensions[rr].height = 34
lr0 = 14
N.cell(lr0, 2, "Legende Fertigungsplan").font = font(size=12, bold=True, color=NAVY)
lg = [("x", "C6E7B9", "1E5631", "erledigt, der Tag ist wie geplant gelaufen"), ("t", "FFE699", "7F6000", "teilweise erledigt, zählt als halber Tag, Rest als Abweichung erfassen"),
      ("o", RED, "FFFFFF", "nicht gelaufen, als Abweichung mit Ursache und Maßnahme erfassen"), ("", "D6E6F5", NAVY, "geplanter Arbeitstag, noch offen"),
      ("", "F8CBAD", "833C0B", "geplanter Arbeitstag in der Vergangenheit ohne Eintrag: Rückmeldung fehlt"), ("", ORANGE, "FFFFFF", "Spaltenkopf in Orange = Stichtag (heute)")]
for i, (sym, bg, fg, t) in enumerate(lg):
    rr = lr0 + 1 + i
    c = N.cell(rr, 2, sym); c.fill = fill(bg); c.font = font(bold=True, color=fg); c.alignment = CENTER; c.border = BOX
    N.cell(rr, 3, t).font = font(size=10)
    N.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
s0 = lr0 + 9
N.cell(s0, 2, "Status je Arbeitsschritt (Spalte L)").font = font(size=12, bold=True, color=NAVY)
sts = [("Erledigt", "C6E7B9", "1E5631", "alle geplanten Tage abgehakt"), ("Ausfall", RED, "FFFFFF", "mindestens ein Tag „o“, Abweichung erfassen"),
       ("Rückmeldung fehlt", "F8CBAD", "833C0B", "vergangene Plantage ohne Eintrag"), ("Überfällig", "FDE2E4", RED, "Plan-Ende überschritten, nicht vollständig erledigt"),
       ("Läuft", "D6E6F5", NAVY, "im Zeitraum, bisher alles gemeldet"), ("Geplant", "F2F2F2", "595959", "Start liegt in der Zukunft")]
for i, (t, bg, fg, d_) in enumerate(sts):
    rr = s0 + 1 + i
    N.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
    c = N.cell(rr, 2, t); c.fill = fill(bg); c.font = font(bold=True, color=fg); c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    N.cell(rr, 4, d_).font = font(size=10)
h0 = s0 + 9
N.cell(h0, 2, "Hinweise").font = font(size=12, bold=True, color=NAVY)
notes = [
    "Plandaten: Blatt „1 Schichtig“ der Datei Kreutzchenplan_Schweißproben_PHPL-Wismar_Ablauf.xlsx. Betriebsruhe 24.12.–01.01. ist kein Arbeitstag.",
    "C04 bis C08 (Rundnaht 301+302 ab Innennaht) sind im Ursprungsplan nicht terminiert. Die Planwerte sind aus den Dauern von Testschuss 1 abgeleitet (kursiv).",
    "Felder mit gelbem Hintergrund sind Eingabefelder. Alle Zahlen, Status und Ampeln rechnen automatisch.",
    "Stichtag im Blatt „Übersicht“ steht auf heute. Für einen Rückblick auf einen anderen Tag dort ein Datum eintragen.",
    "Zellkommentare im Arbeitsschritt (rotes Dreieck) nennen Voraussetzungen, z. B. „Benötigt F03“.",
]
for i, t in enumerate(notes):
    rr = h0 + 1 + i
    N.cell(rr, 2, "•").alignment = CENTER
    N.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
    N.cell(rr, 3, t).font = font(size=10); N.cell(rr, 3).alignment = LEFT
    N.row_dimensions[rr].height = 28

for ws in (wsA, wsU, wsP, wsD, wsV):
    ws.sheet_properties.tabColor = {"Anleitung": GREY, "Übersicht": ORANGE, "Fertigungsplan": NAVY, "Abweichungen": RED, "Verantwortliche": GREEN}[ws.title]
for ws_ in (wsA, wsV):
    ws_.page_setup.orientation = "landscape"; ws_.page_setup.paperSize = 9
    ws_.page_setup.fitToWidth = 1; ws_.page_setup.fitToHeight = 1; ws_.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 1
wb.save(OUT)
print("rows", LAST, "days", len(days), "last col", gl)
