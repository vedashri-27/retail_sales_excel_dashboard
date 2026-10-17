"""Build the Retail Sales Performance Dashboard workbook."""
import pandas as pd
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.comments import Comment

df = pd.read_csv(ROOT / "data" / "sales_data.csv", dtype={"Year": str})
df["Order Date"] = pd.to_datetime(df["Order Date"])
N = len(df)
LAST = N + 1  # last data row on Data sheet

FONT = "Arial"
INK = "1F2937"        # primary text
INK2 = "6B7280"       # secondary text
BG = "F3F4F6"         # page background
WHITE = "FFFFFF"
BANNER = "111827"
BORDER = "E5E7EB"
INPUT_FILL = "FEF3C7"  # filter cells (user inputs)
S1, S2, S3 = "2A78D6", "EB6834", "1BAF7A"  # categorical slots 1-3 (validated palette)
GOOD, BAD = "15803D", "B91C1C"

def f(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)

def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)

thin = Side(style="thin", color=BORDER)
box = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
dash = wb.active
dash.title = "Dashboard"
ana = wb.create_sheet("Analysis")
data = wb.create_sheet("Data")
about = wb.create_sheet("About")

# ----------------------------------------------------------------- DATA sheet
cols = list(df.columns) + ["New Order"]
data.append(cols)
for r in df.itertuples(index=False):
    data.append(list(r))
for i in range(2, LAST + 1):
    data.cell(i, 18, f"=IF(A{i}<>A{i-1},1,0)")
widths = [13, 12, 7, 7, 9, 13, 15, 13, 15, 13, 26, 7, 11, 10, 12, 12, 12, 11]
for i, w in enumerate(widths, 1):
    data.column_dimensions[get_column_letter(i)].width = w
fmt = {2: "dd-mmm-yyyy", 13: "₹#,##0", 14: "0%", 15: "₹#,##0.00", 16: "₹#,##0.00", 17: "₹#,##0.00;[Red]-₹#,##0.00"}
for row in data.iter_rows(min_row=1, max_row=LAST):
    for c in row:
        c.font = f(9)
        if c.row > 1 and c.column in fmt:
            c.number_format = fmt[c.column]
tbl = Table(displayName="SalesData", ref=f"A1:R{LAST}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
data.add_table(tbl)
data.freeze_panes = "A2"
data["R1"].comment = Comment("Helper column: flags the first line of each order so distinct orders can be counted with SUMIFS.", "Analyst")

# named ranges over data columns
names = {"dYear": "C", "dMonth": "D", "dRegion": "E", "dChannel": "H", "dCategory": "I",
         "dProduct": "K", "dUnits": "L", "dDiscount": "N", "dRevenue": "O", "dCost": "P",
         "dProfit": "Q", "dNewOrder": "R"}
for n, col in names.items():
    wb.defined_names[n] = DefinedName(n, attr_text=f"Data!${col}$2:${col}${LAST}")

# ------------------------------------------------------------- ANALYSIS sheet
ana.sheet_view.showGridLines = False
ana.column_dimensions["A"].width = 24
for c in "BCDEFGH":
    ana.column_dimensions[c].width = 15
ana["A1"] = "Analysis layer: every number on the Dashboard is calculated here from the Data sheet"
ana["A1"].font = f(13, True)
ana["A2"] = "All formulas reference named ranges (dRevenue, dRegion, ...) and the filter criteria below. No hardcoded results."
ana["A2"].font = f(9, color=INK2, italic=True)

def header(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row, col + i, lab)
        c.font = f(9, True, WHITE)
        c.fill = fill("374151")
        c.alignment = Alignment(horizontal="center" if i else "left", vertical="center")

def section(ws, row, text):
    c = ws.cell(row, 1, text)
    c.font = f(11, True, S1)

# Filter criteria
section(ana, 4, "1. Filter criteria (from Dashboard dropdowns)")
header(ana, 5, ["Filter", "Selected", "SUMIFS criterion"])
filters = [("Year", "Dashboard!$C$5"), ("Region", "Dashboard!$F$5"), ("Channel", "Dashboard!$I$5")]
for i, (lab, ref) in enumerate(filters):
    r = 6 + i
    ana.cell(r, 1, lab)
    ana.cell(r, 2, f"={ref}")
    if lab == "Year":
        ana.cell(r, 3, f'=IF(B{r}="All","*",TEXT(B{r},"0"))')
    else:
        ana.cell(r, 3, f'=IF(B{r}="All","*",B{r})')
ana.cell(9, 1, "Prior year (for YoY)")
ana.cell(9, 3, '=IF(OR(C6="*",C6="2023"),"",TEXT(C6-1,"0"))')
ana.cell(9, 4, "← blank when 'All' or 2023 is selected (no prior year in data)").font = f(8, color=INK2, italic=True)
for n, ref in {"fYear": "C6", "fRegion": "C7", "fChannel": "C8", "fPrior": "C9"}.items():
    wb.defined_names[n] = DefinedName(n, attr_text=f"Analysis!${ref[0]}${ref[1:]}")

# KPI block
section(ana, 11, "2. Headline KPIs (all filters applied)")
header(ana, 12, ["Metric", "Selected period", "Prior year", "YoY change"])
crit = "dYear,fYear,dRegion,fRegion,dChannel,fChannel"
critP = "dYear,fPrior,dRegion,fRegion,dChannel,fChannel"
kpi_rows = {}
def kpi(r, label, cur, prior, yoy, nf):
    ana.cell(r, 1, label)
    ana.cell(r, 2, cur).number_format = nf
    ana.cell(r, 3, prior).number_format = nf
    ana.cell(r, 4, yoy).number_format = "+0.0%;-0.0%;0.0%" if "pts" not in label else '+0.0" pts";-0.0" pts"'
    kpi_rows[label] = r
kpi(13, "Revenue", f"=SUMIFS(dRevenue,{crit})", f'=IF(fPrior="","",SUMIFS(dRevenue,{critP}))',
    '=IF(OR(C13="",C13=0),"",B13/C13-1)', "₹#,##0")
kpi(14, "Profit", f"=SUMIFS(dProfit,{crit})", f'=IF(fPrior="","",SUMIFS(dProfit,{critP}))',
    '=IF(OR(C14="",C14=0),"",B14/C14-1)', "₹#,##0")
kpi(15, "Profit margin (pts)", "=IF(B13=0,0,B14/B13)", '=IF(OR(C13="",C13=0),"",C14/C13)',
    '=IF(C15="","",(B15-C15)*100)', "0.0%")
kpi(16, "Orders", f"=SUMIFS(dNewOrder,{crit})", f'=IF(fPrior="","",SUMIFS(dNewOrder,{critP}))',
    '=IF(OR(C16="",C16=0),"",B16/C16-1)', "#,##0")
kpi(17, "Avg order value", "=IF(B16=0,0,B13/B16)", '=IF(OR(C16="",C16=0),"",C13/C16)',
    '=IF(OR(C17="",C17=0),"",B17/C17-1)', "₹#,##0")
kpi(18, "Units sold", f"=SUMIFS(dUnits,{crit})", f'=IF(fPrior="","",SUMIFS(dUnits,{critP}))',
    '=IF(OR(C18="",C18=0),"",B18/C18-1)', "#,##0")

# Monthly trend
section(ana, 20, "3. Monthly revenue, ₹ Lakh (Region + Channel filters; all years shown for comparison)")
header(ana, 21, ["Month #", "Month", "2023", "2024", "2025", "Selected filters (₹)"])
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
for i, m in enumerate(months):
    r = 22 + i
    ana.cell(r, 1, i + 1)
    ana.cell(r, 2, m)
    for j, y in enumerate(["2023", "2024", "2025"]):
        ana.cell(r, 3 + j, f'=SUMIFS(dRevenue,dYear,"{y}",dMonth,$A{r},dRegion,fRegion,dChannel,fChannel)/100000').number_format = "#,##0.0"
    ana.cell(r, 6, f"=SUMIFS(dRevenue,dMonth,$A{r},{crit})").number_format = "₹#,##0"
MON0, MON1 = 22, 33

# Category
section(ana, 35, "4. By category (all filters)")
header(ana, 36, ["Category", "Revenue (₹ L)", "Profit (₹ L)", "Margin", "Revenue share"])
cats = ["Electronics", "Fashion", "Home & Kitchen", "Grocery"]
for i, c in enumerate(cats):
    r = 37 + i
    ana.cell(r, 1, c)
    ana.cell(r, 2, f"=SUMIFS(dRevenue,dCategory,$A{r},{crit})/100000").number_format = "#,##0.0"
    ana.cell(r, 3, f"=SUMIFS(dProfit,dCategory,$A{r},{crit})/100000").number_format = "#,##0.0"
    ana.cell(r, 4, f"=IF(B{r}=0,0,C{r}/B{r})").number_format = "0.0%"
    ana.cell(r, 5, f"=IF(SUM($B$37:$B$40)=0,0,B{r}/SUM($B$37:$B$40))").number_format = "0.0%"
CAT0, CAT1 = 37, 40

# Region
section(ana, 42, "5. By region (Year + Channel filters; Region filter ignored so all regions compare)")
header(ana, 43, ["Region", "Revenue (₹ L)", "Profit (₹ L)", "Margin", "Revenue share"])
regs = ["North", "South", "East", "West"]
for i, c in enumerate(regs):
    r = 44 + i
    ana.cell(r, 1, c)
    ana.cell(r, 2, f"=SUMIFS(dRevenue,dRegion,$A{r},dYear,fYear,dChannel,fChannel)/100000").number_format = "#,##0.0"
    ana.cell(r, 3, f"=SUMIFS(dProfit,dRegion,$A{r},dYear,fYear,dChannel,fChannel)/100000").number_format = "#,##0.0"
    ana.cell(r, 4, f"=IF(B{r}=0,0,C{r}/B{r})").number_format = "0.0%"
    ana.cell(r, 5, f"=IF(SUM($B$44:$B$47)=0,0,B{r}/SUM($B$44:$B$47))").number_format = "0.0%"
REG0, REG1 = 44, 47

# Channel
section(ana, 49, "6. By channel (Year + Region filters; Channel filter ignored)")
header(ana, 50, ["Channel", "Revenue (₹ L)", "Profit (₹ L)", "Margin", "Revenue share"])
chans = ["Online", "In-Store", "Marketplace"]
for i, c in enumerate(chans):
    r = 51 + i
    ana.cell(r, 1, c)
    ana.cell(r, 2, f"=SUMIFS(dRevenue,dChannel,$A{r},dYear,fYear,dRegion,fRegion)/100000").number_format = "#,##0.0"
    ana.cell(r, 3, f"=SUMIFS(dProfit,dChannel,$A{r},dYear,fYear,dRegion,fRegion)/100000").number_format = "#,##0.0"
    ana.cell(r, 4, f"=IF(B{r}=0,0,C{r}/B{r})").number_format = "0.0%"
    ana.cell(r, 5, f"=IF(SUM($B$51:$B$53)=0,0,B{r}/SUM($B$51:$B$53))").number_format = "0.0%"
CH0, CH1 = 51, 53

# Discount
section(ana, 55, "7. Profit margin by discount level (all filters)")
header(ana, 56, ["Discount", "Label", "Revenue (₹ L)", "Profit (₹ L)", "Margin", "Revenue share"])
discs = [0, 0.05, 0.10, 0.15, 0.20, 0.25, 0.30]
for i, d in enumerate(discs):
    r = 57 + i
    ana.cell(r, 1, d).number_format = "0%"
    ana.cell(r, 2, "None" if d == 0 else f"{int(d*100)}%")
    ana.cell(r, 3, f"=SUMIFS(dRevenue,dDiscount,$A{r},{crit})/100000").number_format = "#,##0.0"
    ana.cell(r, 4, f"=SUMIFS(dProfit,dDiscount,$A{r},{crit})/100000").number_format = "#,##0.0"
    ana.cell(r, 5, f"=IF(C{r}=0,0,D{r}/C{r})").number_format = "0.0%"
    ana.cell(r, 6, f"=IF(SUM($C$57:$C$63)=0,0,C{r}/SUM($C$57:$C$63))").number_format = "0.0%"
DIS0, DIS1 = 57, 63
ana.cell(64, 1, "Margin at 20%+ discount")
ana.cell(64, 5, f'=IFERROR(SUMIFS(dProfit,dDiscount,">=0.2",{crit})/SUMIFS(dRevenue,dDiscount,">=0.2",{crit}),0)').number_format = "0.0%"
ana.cell(65, 1, "Margin at full price")
ana.cell(65, 5, "=E57").number_format = "0.0%"

# Products
section(ana, 67, "8. Product league table (all filters); rank breaks ties by list order")
header(ana, 68, ["Product", "Category", "Units", "Revenue (₹)", "Profit (₹)", "Margin", "Rank"])
prods = (df.groupby("Product")["Category"].first().reset_index()
         .sort_values(["Category", "Product"]).values.tolist())
P0 = 69
P1 = P0 + len(prods) - 1
for i, (p, c) in enumerate(prods):
    r = P0 + i
    ana.cell(r, 1, p)
    ana.cell(r, 2, c)
    ana.cell(r, 3, f"=SUMIFS(dUnits,dProduct,$A{r},{crit})").number_format = "#,##0"
    ana.cell(r, 4, f"=SUMIFS(dRevenue,dProduct,$A{r},{crit})").number_format = "₹#,##0"
    ana.cell(r, 5, f"=SUMIFS(dProfit,dProduct,$A{r},{crit})").number_format = "₹#,##0"
    ana.cell(r, 6, f"=IF(D{r}=0,0,E{r}/D{r})").number_format = "0.0%"
    ana.cell(r, 7, f"=RANK(D{r},$D${P0}:$D${P1},0)+COUNTIF($D${P0}:D{r},D{r})-1")

# Insights (dynamic text)
IN0 = P1 + 2
section(ana, IN0, "9. Auto-generated insights (feed the Dashboard text panel)")
ins = [
    f'=IFERROR(INDEX($A${CAT0}:$A${CAT1},MATCH(MAX($B${CAT0}:$B${CAT1}),$B${CAT0}:$B${CAT1},0))&" is the largest category at "&TEXT(MAX($E${CAT0}:$E${CAT1}),"0%")&" of revenue, but earns only a "&TEXT(INDEX($D${CAT0}:$D${CAT1},MATCH(MAX($B${CAT0}:$B${CAT1}),$B${CAT0}:$B${CAT1},0)),"0%")&" margin.","No data for this selection.")',
    f'=IFERROR(INDEX($A${CAT0}:$A${CAT1},MATCH(MAX($D${CAT0}:$D${CAT1}),$D${CAT0}:$D${CAT1},0))&" is the most profitable category with a "&TEXT(MAX($D${CAT0}:$D${CAT1}),"0.0%")&" margin vs "&TEXT($B$15,"0.0%")&" overall.","No data for this selection.")',
    f'=IFERROR(INDEX($B${MON0}:$B${MON1},MATCH(MAX($F${MON0}:$F${MON1}),$F${MON0}:$F${MON1},0))&" is the peak month, bringing in "&TEXT(MAX($F${MON0}:$F${MON1})/SUM($F${MON0}:$F${MON1}),"0%")&" of revenue for the selection.","No data for this selection.")',
    f'="Lines discounted 20% or more earn a "&TEXT($E$64,"0.0%")&" margin, vs "&TEXT($E$65,"0.0%")&" at full price."',
    f'=IFERROR(INDEX($A${REG0}:$A${REG1},MATCH(MAX($B${REG0}:$B${REG1}),$B${REG0}:$B${REG1},0))&" leads all regions with "&TEXT(MAX($E${REG0}:$E${REG1}),"0%")&" of revenue; "&INDEX($A${REG0}:$A${REG1},MATCH(MIN($B${REG0}:$B${REG1}),$B${REG0}:$B${REG1},0))&" trails at "&TEXT(MIN($E${REG0}:$E${REG1}),"0%")&".","No data for this selection.")',
]
for i, t in enumerate(ins):
    ana.cell(IN0 + 1 + i, 1, f"Insight {i+1}")
    ana.cell(IN0 + 1 + i, 2, t)

for row in ana.iter_rows(min_row=3, max_row=IN0 + 6):
    for c in row:
        if c.font.color is None or c.font.color.rgb in (None, "FF000000"):
            c.font = f(9)
ana.freeze_panes = "A4"

# ------------------------------------------------------------ DASHBOARD sheet
ws = dash
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 90
MAIN, GAP = 12, 2
ws.column_dimensions["A"].width = GAP
for i in range(2, 20):
    col = get_column_letter(i)
    ws.column_dimensions[col].width = GAP if col in "DGJMP" else MAIN
ws.column_dimensions["S"].width = GAP
for r in range(1, 66):
    for c in range(1, 20):
        ws.cell(r, c).fill = fill(BG)

# Banner
for r in (2, 3):
    for c in range(2, 19):
        ws.cell(r, c).fill = fill(BANNER)
ws.row_dimensions[2].height = 30
ws.row_dimensions[3].height = 18
ws["B2"] = "Retail Sales Performance Dashboard"
ws["B2"].font = f(18, True, WHITE)
ws["B2"].alignment = Alignment(vertical="center", indent=1)
ws["B3"] = "FY2023 – FY2025  ·  4 regions · 16 cities · 3 channels · 26 products  ·  Amounts in ₹ (L = Lakh, Cr = Crore)"
ws["B3"].font = f(9, color="D1D5DB")
ws["B3"].alignment = Alignment(vertical="top", indent=1)

# Filters
ws.row_dimensions[5].height = 22
dv_specs = [("B5", "C5", "Year", '"All,2023,2024,2025"', "All"),
            ("E5", "F5", "Region", '"All,North,South,East,West"', "All"),
            ("H5", "I5", "Channel", '"All,Online,In-Store,Marketplace"', "All")]
for lab_cell, val_cell, lab, lst, default in dv_specs:
    ws[lab_cell] = lab + "  ▸"
    ws[lab_cell].font = f(10, True, INK2)
    ws[lab_cell].alignment = Alignment(horizontal="right", vertical="center")
    ws[val_cell] = default
    ws[val_cell].font = f(10, True, INK)
    ws[val_cell].fill = fill(INPUT_FILL)
    ws[val_cell].border = Border(left=Side("thin", "D97706"), right=Side("thin", "D97706"),
                                 top=Side("thin", "D97706"), bottom=Side("thin", "D97706"))
    ws[val_cell].alignment = Alignment(horizontal="center", vertical="center")
    dv = DataValidation(type="list", formula1=lst, allow_blank=False, showDropDown=False)
    dv.error = f"Pick a {lab.lower()} from the list."
    dv.errorTitle = "Invalid filter"
    dv.prompt = f"Filter the whole dashboard by {lab.lower()}"
    dv.promptTitle = lab
    ws.add_data_validation(dv)
    dv.add(val_cell)
ws["K5"] = "← Pick from the yellow dropdowns. Every KPI, chart and insight updates."
ws["K5"].font = f(9, color=INK2, italic=True)
ws["K5"].alignment = Alignment(vertical="center")

# KPI cards
cards = [("B", "C", "TOTAL REVENUE", "=Analysis!B13/10000000", '"₹"0.00" Cr"', 13, "pct"),
         ("E", "F", "TOTAL PROFIT", "=Analysis!B14/10000000", '"₹"0.00" Cr"', 14, "pct"),
         ("H", "I", "PROFIT MARGIN", "=Analysis!B15", "0.0%", 15, "pts"),
         ("K", "L", "ORDERS", "=Analysis!B16", "#,##0", 16, "pct"),
         ("N", "O", "AVG ORDER VALUE", "=Analysis!B17", '"₹"#,##0', 17, "pct"),
         ("Q", "R", "UNITS SOLD", "=Analysis!B18", "#,##0", 18, "pct")]
ws.row_dimensions[7].height = 20
ws.row_dimensions[8].height = 34
ws.row_dimensions[9].height = 20
for c1, c2, label, formula, nf, arow, kind in cards:
    for r in (7, 8, 9):
        ws.merge_cells(f"{c1}{r}:{c2}{r}")
        for cc in (c1, c2):
            ws[f"{cc}{r}"].fill = fill(WHITE)
    # accent top border + outline
    for cc in (c1, c2):
        ws[f"{cc}7"].border = Border(top=Side("thick", S1), left=thin if cc == c1 else None, right=thin if cc == c2 else None)
        ws[f"{cc}8"].border = Border(left=thin if cc == c1 else None, right=thin if cc == c2 else None)
        ws[f"{cc}9"].border = Border(bottom=thin, left=thin if cc == c1 else None, right=thin if cc == c2 else None)
    ws[f"{c1}7"] = label
    ws[f"{c1}7"].font = f(8, True, INK2)
    ws[f"{c1}7"].alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
    ws[f"{c1}8"] = formula
    ws[f"{c1}8"].number_format = nf
    ws[f"{c1}8"].font = f(18, True, INK)
    ws[f"{c1}8"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    if kind == "pts":
        sub = f'=IF(Analysis!D{arow}="",IF(fYear="*","All years combined","No prior year in data"),IF(ROUND(Analysis!D{arow},1)=0,"■ Flat vs "&fPrior,IF(Analysis!D{arow}>0,"▲ ","▼ ")&TEXT(ABS(Analysis!D{arow}),"0.0")&" pts vs "&fPrior))'
    else:
        sub = f'=IF(Analysis!D{arow}="",IF(fYear="*","All years combined","No prior year in data"),IF(ROUND(Analysis!D{arow},3)=0,"■ Flat vs "&fPrior,IF(Analysis!D{arow}>0,"▲ ","▼ ")&TEXT(ABS(Analysis!D{arow}),"0.0%")&" vs "&fPrior))'
    ws[f"{c1}9"] = sub
    ws[f"{c1}9"].font = f(8, color=INK2)
    ws[f"{c1}9"].alignment = Alignment(horizontal="left", vertical="top", indent=1)
    ws.conditional_formatting.add(f"{c1}9", FormulaRule(formula=[f'AND(ISNUMBER(Analysis!$D${arow}),ROUND(Analysis!$D${arow},{1 if kind == "pts" else 3})>0)'], font=Font(name=FONT, size=8, bold=True, color=GOOD)))
    ws.conditional_formatting.add(f"{c1}9", FormulaRule(formula=[f'AND(ISNUMBER(Analysis!$D${arow}),ROUND(Analysis!$D${arow},{1 if kind == "pts" else 3})<0)'], font=Font(name=FONT, size=8, bold=True, color=BAD)))

def panel_header(cell_from, cell_to, title, note):
    ws.merge_cells(f"{cell_from}:{cell_to}")
    c = ws[cell_from]
    c.value = title
    c.font = f(11, True, INK)
    c.alignment = Alignment(vertical="center", indent=1)
    r = c.row
    c0 = ws[cell_from].column
    c1 = ws[cell_to].column
    for cc in range(c0, c1 + 1):
        ws.cell(r, cc).fill = fill(WHITE)
        ws.cell(r + 1, cc).fill = fill(WHITE)
        ws.cell(r, cc).border = Border(top=thin, left=thin if cc == c0 else None, right=thin if cc == c1 else None)
    ws.merge_cells(start_row=r + 1, start_column=c0, end_row=r + 1, end_column=c1)
    n = ws.cell(r + 1, c0, note)
    n.font = f(8, color=INK2, italic=True)
    n.alignment = Alignment(vertical="top", indent=1)
    ws.row_dimensions[r].height = 20

# Row map
panel_header("B11", "L11", "Monthly Revenue Trend", "₹ Lakh by month, one line per year · follows Region + Channel filters")
panel_header("N11", "R11", "Key Insights", "Written by formulas, so they rewrite themselves when you change a filter")
panel_header("B30", "F30", "Revenue & Profit by Category", "₹ Lakh · all filters")
panel_header("H30", "L30", "Revenue & Profit by Region", "₹ Lakh · Year + Channel filters")
panel_header("N30", "R30", "Revenue by Channel", "₹ Lakh · Year + Region filters")
panel_header("B48", "L48", "Top 10 Products by Revenue", "All filters · ranked live from the Analysis sheet")
panel_header("N48", "R48", "Profit Margin by Discount Level", "All filters · deeper discounts erode margin")

# Insights text panel
for i in range(5):
    r0 = 13 + i * 3
    ws.merge_cells(start_row=r0, start_column=14, end_row=r0 + 2, end_column=18)
    c = ws.cell(r0, 14, f"=\"●  \"&Analysis!B{IN0 + 1 + i}")
    c.font = f(10, color=INK)
    c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
for r in range(13, 29):
    for cc in range(14, 19):
        ws.cell(r, cc).fill = fill(WHITE)
        ws.cell(r, cc).border = Border(left=thin if cc == 14 else None, right=thin if cc == 18 else None,
                                       bottom=thin if r == 28 else None)

# Top 10 table
TH = 50
heads = [("B", "B", "#"), ("C", "E", "Product"), ("F", "H", "Category"), ("I", "J", "Units"), ("K", "K", "Revenue"), ("L", "L", "Margin")]
for r in range(TH, TH + 11):
    for cc in range(2, 13):
        ws.cell(r, cc).fill = fill(WHITE)
        ws.cell(r, cc).border = Border(left=thin if cc == 2 else None, right=thin if cc == 12 else None,
                                       bottom=thin if r == TH + 10 else Side("hair", BORDER))
for a, b, h in heads:
    if a != b:
        ws.merge_cells(f"{a}{TH}:{b}{TH}")
    c = ws[f"{a}{TH}"]
    c.value = h
    c.font = f(9, True, WHITE)
for cc in range(2, 13):
    ws.cell(TH, cc).fill = fill("374151")
for k in range(1, 11):
    r = TH + k
    m = f"MATCH({k},Analysis!$G${P0}:$G${P1},0)"
    ws[f"B{r}"] = k
    ws.merge_cells(f"C{r}:E{r}")
    ws[f"C{r}"] = f"=INDEX(Analysis!$A${P0}:$A${P1},{m})"
    ws.merge_cells(f"F{r}:H{r}")
    ws[f"F{r}"] = f"=INDEX(Analysis!$B${P0}:$B${P1},{m})"
    ws.merge_cells(f"I{r}:J{r}")
    ws[f"I{r}"] = f"=INDEX(Analysis!$C${P0}:$C${P1},{m})"
    ws[f"K{r}"] = f"=INDEX(Analysis!$D${P0}:$D${P1},{m})"
    ws[f"L{r}"] = f"=INDEX(Analysis!$F${P0}:$F${P1},{m})"
    ws[f"I{r}"].number_format = "#,##0"
    ws[f"K{r}"].number_format = "₹#,##0"
    ws[f"L{r}"].number_format = "0.0%"
    for col in "BCFIKL":
        ws[f"{col}{r}"].font = f(9, bold=(col == "C"))
    ws[f"B{r}"].font = f(9, True, INK2)
for col in "BIKL":
    for r in range(TH, TH + 11):
        ws[f"{col}{r}"].alignment = Alignment(horizontal="center" if col == "B" else "right", vertical="center")
for r in range(TH, TH + 11):
    for col in "CF":
        ws[f"{col}{r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1 if col == "C" else 0)
ws.conditional_formatting.add(f"K{TH+1}:K{TH+10}", DataBarRule(start_type="num", start_value=0, end_type="max", color="9CC3EE", showValue=True))
ws.conditional_formatting.add(f"L{TH+1}:L{TH+10}", FormulaRule(formula=[f"L{TH+1}<Analysis!$B$15"], font=Font(name=FONT, size=9, color=BAD)))
ws[f"B{TH+11}"] = "Margin in red = below the overall margin for the current selection."
ws[f"B{TH+11}"].font = f(8, color=INK2, italic=True)

# ------------------------------------------------------------------ charts
def px_cm(px):
    return px * 2.54 / 96

COLPX, GAPPX, ROWPX = 89, 19, 20

def grid_line():
    return ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="E5E7EB", w=6350)))

def style_axes(ch, yfmt="#,##0"):
    ch.y_axis.majorGridlines = grid_line()
    ch.y_axis.numFmt = yfmt
    ch.y_axis.number_format = yfmt
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="9CA3AF", w=9525))
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont
    def txt(sz=800):
        cp = CharacterProperties(sz=sz, solidFill="6B7280", latin=DFont(typeface=FONT))
        return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])
    ch.x_axis.txPr = txt()
    ch.y_axis.txPr = txt()
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(solidFill="E5E7EB", w=9525))
    if ch.legend is not None:
        ch.legend.position = "b"
        ch.legend.txPr = txt(800)

def bar_colors(ch, colors):
    for s, c in zip(ch.series, colors):
        s.graphicalProperties.solidFill = c
        s.graphicalProperties.line.noFill = True

# Trend line chart
lc = LineChart()
lc.title = None
lc.height = px_cm(16 * ROWPX) - 0.15
lc.width = px_cm(8 * COLPX + 3 * GAPPX)
lc.add_data(Reference(ana, min_col=3, max_col=5, min_row=21, max_row=MON1), titles_from_data=True)
lc.set_categories(Reference(ana, min_col=2, min_row=MON0, max_row=MON1))
for s, c in zip(lc.series, [S1, S2, S3]):
    s.graphicalProperties.line.solidFill = c
    s.graphicalProperties.line.width = 28575
    s.smooth = False
    s.marker.symbol = "circle"
    s.marker.size = 5
    s.marker.graphicalProperties = GraphicalProperties(solidFill=c, ln=LineProperties(solidFill="FFFFFF"))
style_axes(lc)
ws.add_chart(lc, "B13")

def bar(min_row, max_row, val_cols, colors, horiz=False, yfmt="#,##0", labels=False, catcol=1, sheet=ana, hdr_row=None):
    b = BarChart()
    b.type = "bar" if horiz else "col"
    b.grouping = "clustered"
    b.gapWidth = 60
    b.overlap = -5 if len(val_cols) > 1 else 0
    b.title = None
    for vc in val_cols:
        b.add_data(Reference(sheet, min_col=vc, min_row=hdr_row, max_row=max_row), titles_from_data=True)
    b.set_categories(Reference(sheet, min_col=catcol, min_row=min_row, max_row=max_row))
    bar_colors(b, colors)
    style_axes(b, yfmt)
    if len(val_cols) == 1:
        b.legend = None
    if labels:
        b.dataLabels = DataLabelList()
        b.dataLabels.showVal = True
        b.dataLabels.showSerName = False
        b.dataLabels.showCatName = False
        b.dataLabels.showLegendKey = False
        b.dataLabels.showPercent = False
        b.dataLabels.numFmt = yfmt
        from openpyxl.chart.text import RichText
        from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont
        cp = CharacterProperties(sz=800, b=True, solidFill="1F2937", latin=DFont(typeface=FONT))
        b.dataLabels.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])
    if horiz:
        b.x_axis.scaling.orientation = "maxMin"  # keep list order top-to-bottom
        b.y_axis.delete = True  # values are labelled on the bars
        b.y_axis.majorGridlines = None
    b.width = px_cm(4 * COLPX + GAPPX)
    b.height = px_cm(15 * ROWPX) - 0.15
    return b

ws.add_chart(bar(CAT0, CAT1, [2, 3], [S1, S2], hdr_row=36), "B32")
ws.add_chart(bar(REG0, REG1, [2, 3], [S1, S2], hdr_row=43), "H32")
ws.add_chart(bar(CH0, CH1, [2], [S1], horiz=True, labels=True, yfmt="#,##0", hdr_row=50), "N32")
dc = bar(DIS0, DIS1, [5], [S1], labels=True, yfmt="0%", catcol=2, hdr_row=56)
dc.height = px_cm(12 * ROWPX) - 0.15
ws.add_chart(dc, "N50")

# white fill behind chart areas so panels read as cards
for (r0, r1, c0, c1) in [(13, 28, 2, 12), (32, 46, 2, 6), (32, 46, 8, 12), (32, 46, 14, 18), (50, 61, 14, 18)]:
    for r in range(r0, r1 + 1):
        for cc in range(c0, c1 + 1):
            ws.cell(r, cc).fill = fill(WHITE)

ws.print_area = "A1:S62"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.sheet_properties.tabColor = S1
ana.sheet_properties.tabColor = "6B7280"
data.sheet_properties.tabColor = "9CA3AF"

# ---------------------------------------------------------------- ABOUT sheet
about.sheet_view.showGridLines = False
about.column_dimensions["A"].width = 3
about.column_dimensions["B"].width = 24
about.column_dimensions["C"].width = 95
lines = [
    ("title", "About this project"),
    ("p", "An interactive Excel dashboard analysing three years of retail sales for a multi-region Indian retailer. Built entirely with native formulas (no macros, no PivotTables), so every number is traceable from the Data sheet to the Dashboard."),
    ("h", "Business questions answered"),
    ("kv", ("Growth", "How did revenue, profit and order volume change year over year?")),
    ("kv", ("Mix", "Which categories, regions and channels drive revenue, and which drive profit?")),
    ("kv", ("Seasonality", "When does demand peak, and how strong is the festive-season lift?")),
    ("kv", ("Pricing", "What do discounts cost in margin terms?")),
    ("kv", ("Products", "Which products lead on revenue, and are the top sellers also profitable?")),
    ("h", "How to use"),
    ("kv", ("Filters", "On the Dashboard, pick a Year, Region and Channel from the yellow dropdowns. KPIs, charts, the Top 10 table and the insight sentences all recalculate."),),
    ("kv", ("YoY", "Pick 2024 or 2025 to see year-over-year change on the KPI cards (green ▲ / red ▼).")),
    ("h", "Sheet guide"),
    ("kv", ("Dashboard", "One-page view for stakeholders: 6 KPI cards, 5 charts, a live Top 10 table and auto-written insights.")),
    ("kv", ("Analysis", "The calculation layer: filter criteria, KPI block with prior-year comparison, and aggregate tables that feed every chart.")),
    ("kv", ("Data", f"{N:,} order lines ({df['Order ID'].nunique():,} orders) in an Excel Table, with a helper column to count distinct orders.")),
    ("h", "Techniques demonstrated"),
    ("kv", ("Formulas", "SUMIFS with wildcard criteria for 'All' filters, INDEX/MATCH lookups, RANK + COUNTIF tie-breaking, IFERROR guards, dynamic text with TEXT()")),
    ("kv", ("Modelling", "Named ranges (dRevenue, fYear …), separated data / calculation / presentation layers, no hardcoded results")),
    ("kv", ("Interactivity", "Data-validation dropdowns driving the entire workbook")),
    ("kv", ("Design", "KPI cards, conditional formatting (data bars, ▲▼ colour cues, below-average margin flag), consistent colour per measure, colour-blind-safe palette")),
    ("h", "Data dictionary"),
    ("kv", ("Order ID / Order Date", "Order identifier and date; one order can have several product lines")),
    ("kv", ("Year / Month", "Calendar year (text, so filters can use a wildcard) and month number")),
    ("kv", ("Region / City", "4 regions, 16 cities")),
    ("kv", ("Segment / Channel", "Consumer, Corporate, Small Business · Online, In-Store, Marketplace")),
    ("kv", ("Category / Sub-Category / Product", "4 categories, 13 sub-categories, 24 products")),
    ("kv", ("Units / Unit Price / Discount", "Quantity, list price per unit (₹) and discount rate applied")),
    ("kv", ("Revenue / Cost / Profit", "Revenue = Units × Unit Price × (1 − Discount); Profit = Revenue − Cost (₹)")),
    ("kv", ("New Order", "Helper flag = 1 on the first line of each order")),
    ("h", "Data source"),
    ("p", "Synthetic dataset generated for portfolio demonstration (seeded Python simulation with realistic seasonality, a festive-season peak in Oct–Nov, year-on-year growth and category-specific margins). It does not represent any real company."),
]
r = 2
for kind, content in lines:
    if kind == "title":
        about.cell(r, 2, content).font = f(16, True)
        r += 1
    elif kind == "p":
        about.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = about.cell(r, 2, content)
        c.font = f(10, color=INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        about.row_dimensions[r].height = 30
        r += 1
    elif kind == "h":
        r += 1
        about.cell(r, 2, content).font = f(11, True, S1)
        r += 1
    else:
        k, v = content if isinstance(content, tuple) and len(content) == 2 and isinstance(content[0], str) else content[0]
        about.cell(r, 2, k).font = f(10, True)
        c = about.cell(r, 3, v)
        c.font = f(10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        about.cell(r, 2).alignment = Alignment(vertical="top")
        r += 1

wb.calculation.fullCalcOnLoad = True
wb.active = 0
wb.save(ROOT / "Retail_Sales_Dashboard.xlsx")
print("saved; products", len(prods), "insights at", IN0)
