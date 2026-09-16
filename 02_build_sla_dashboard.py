import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, LineChart
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule

path = "/home/claude/travclan/TravClan_Booking_Ops_SLA_Tracker.xlsx"
wb = openpyxl.load_workbook(path)
raw = wb["Raw Bookings"]
n_rows = raw.max_row  # includes header
data_last = n_rows

suppliers = sorted(set(raw.cell(row=r, column=6).value for r in range(2, n_rows+1)))

dash = wb.create_sheet("Supplier SLA Dashboard")
ARIAL = "Arial"
title_font = Font(name=ARIAL, size=16, bold=True, color="1F4E78")
h_font = Font(name=ARIAL, size=11, bold=True, color="FFFFFF")
h_fill = PatternFill("solid", fgColor="1F4E78")
label_font = Font(name=ARIAL, size=10, bold=True)
normal_font = Font(name=ARIAL, size=10)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

dash["B2"] = "TravClan Operations — Supplier SLA & Voucher Delivery Dashboard"
dash["B2"].font = title_font
dash["B3"] = "Data source: Raw Bookings tab | Auto-updates via formulas"
dash["B3"].font = Font(name=ARIAL, size=9, italic=True, color="808080")

# ---- KPI summary row ----
kpi_row = 5
kpis = [
    ("Total Bookings", f"=COUNTA('Raw Bookings'!A2:A{data_last})"),
    ("Total GMV (INR)", f"=SUM('Raw Bookings'!G2:G{data_last})"),
    ("On-Time Voucher %", f"=COUNTIF('Raw Bookings'!K2:K{data_last},\"On Time\")/COUNTA('Raw Bookings'!K2:K{data_last})"),
    ("Avg Payment Delay (days)", f"=AVERAGE('Raw Bookings'!L2:L{data_last})"),
    ("Pending KYC Agents", f"=COUNTIF('Raw Bookings'!D2:D{data_last},\"Pending\")+COUNTIF('Raw Bookings'!D2:D{data_last},\"Expired\")"),
]
for i, (label, formula) in enumerate(kpis):
    col = 2 + i * 2
    c1 = dash.cell(row=kpi_row, column=col, value=label)
    c1.font = h_font
    c1.fill = h_fill
    c1.alignment = Alignment(horizontal="center", wrap_text=True)
    dash.merge_cells(start_row=kpi_row, start_column=col, end_row=kpi_row, end_column=col+1)
    c2 = dash.cell(row=kpi_row+1, column=col, value=formula)
    c2.font = Font(name=ARIAL, size=13, bold=True, color="1F4E78")
    c2.alignment = Alignment(horizontal="center")
    dash.merge_cells(start_row=kpi_row+1, start_column=col, end_row=kpi_row+1, end_column=col+1)
    if "%" in label:
        c2.number_format = "0.0%"
    elif "GMV" in label:
        c2.number_format = "#,##0"
    elif "Delay" in label:
        c2.number_format = "0.0"

# ---- Supplier-wise SLA breakdown table ----
tbl_start = 9
headers = ["Supplier", "Bookings", "Total GMV (INR)", "Avg Voucher TAT (hrs)",
           "On-Time Voucher %", "SLA Breaches", "Avg Payment Delay (days)", "Ops Risk Flag"]
for j, h in enumerate(headers):
    cell = dash.cell(row=tbl_start, column=2+j, value=h)
    cell.font = h_font
    cell.fill = h_fill
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    cell.border = border

for i, sup in enumerate(suppliers):
    r = tbl_start + 1 + i
    dash.cell(row=r, column=2, value=sup)
    dash.cell(row=r, column=3, value=f"=COUNTIF('Raw Bookings'!$F$2:$F${data_last},$B{r})")
    dash.cell(row=r, column=4, value=f"=SUMIF('Raw Bookings'!$F$2:$F${data_last},$B{r},'Raw Bookings'!$G$2:$G${data_last})")
    dash.cell(row=r, column=5, value=f"=AVERAGEIF('Raw Bookings'!$F$2:$F${data_last},$B{r},'Raw Bookings'!$I$2:$I${data_last})")
    dash.cell(row=r, column=6, value=(
        f"=SUMPRODUCT(('Raw Bookings'!$F$2:$F${data_last}=$B{r})*"
        f"('Raw Bookings'!$K$2:$K${data_last}=\"On Time\"))/$C{r}"
    ))
    dash.cell(row=r, column=7, value=(
        f"=SUMPRODUCT(('Raw Bookings'!$F$2:$F${data_last}=$B{r})*"
        f"('Raw Bookings'!$K$2:$K${data_last}=\"Breach\"))"
    ))
    dash.cell(row=r, column=8, value=f"=AVERAGEIF('Raw Bookings'!$F$2:$F${data_last},$B{r},'Raw Bookings'!$L$2:$L${data_last})")
    dash.cell(row=r, column=9, value=f'=IF(OR($F{r}<0.7,$H{r}>4),"Review Needed","Healthy")')

    for col in range(2, 10):
        cell = dash.cell(row=r, column=col)
        cell.font = normal_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center")
    dash.cell(row=r, column=3).number_format = "0"
    dash.cell(row=r, column=4).number_format = "#,##0"
    dash.cell(row=r, column=5).number_format = "0.0"
    dash.cell(row=r, column=6).number_format = "0.0%"
    dash.cell(row=r, column=7).number_format = "0"
    dash.cell(row=r, column=8).number_format = "0.0"

table_end = tbl_start + len(suppliers)

# Conditional formatting: On-Time % column and Risk Flag column
dash.conditional_formatting.add(
    f"F{tbl_start+1}:F{table_end}",
    ColorScaleRule(start_type="min", start_color="F8696B",
                   mid_type="percentile", mid_value=50, mid_color="FFEB84",
                   end_type="max", end_color="63BE7B")
)
dash.conditional_formatting.add(
    f"I{tbl_start+1}:I{table_end}",
    CellIsRule(operator="equal", formula=['"Review Needed"'], fill=PatternFill("solid", fgColor="F8696B"))
)

for col, width in zip(range(2, 10), [24, 10, 16, 18, 16, 12, 18, 15]):
    dash.column_dimensions[get_column_letter(col)].width = width

# ---- Chart: On-time % by supplier ----
chart = BarChart()
chart.title = "On-Time Voucher Delivery % by Supplier"
chart.y_axis.title = "% On Time"
chart.x_axis.title = "Supplier"
chart.style = 10
data_ref = Reference(dash, min_col=6, min_row=tbl_start, max_row=table_end)
cats_ref = Reference(dash, min_col=2, min_row=tbl_start+1, max_row=table_end)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats_ref)
chart.height = 8
chart.width = 18
dash.add_chart(chart, f"B{table_end+3}")

dash.sheet_view.showGridLines = False
wb.save(path)
print("dashboard sheet added; suppliers:", suppliers)
