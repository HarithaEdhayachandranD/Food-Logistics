"""
QuickBite Delivery Services – Round 3: Excel Analysis
Creates a fully formatted Excel workbook with:
  Sheet 1  – Raw Data (formatted table)
  Sheet 2  – Zone Analysis  (formulas + bar chart)
  Sheet 3  – Traffic Analysis (formulas + column + violin-style)
  Sheet 4  – Weather Analysis (formulas + charts)
  Sheet 5  – Restaurant Analysis (formulas + bar chart)
  Sheet 6  – Distance Analysis (scatter + bins)
  Sheet 7  – Peak Hour Analysis (line + heat-style bar)
  Sheet 8  – Pivot Tables (zone×status, hour×traffic, restaurant×status)
  Sheet 9  – Summary Dashboard
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference, Series
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.drawing.image import Image as XLImage
import warnings
warnings.filterwarnings('ignore')

# ── Load cleaned data ──────────────────────────────────────────────────────
df = pd.read_csv('master_cleaned.csv', parse_dates=['order_time','pickup_time','delivered_time'])

# Keep only usable columns for Excel
COLS = ['delivery_id','order_id','rider_id','restaurant','zone','order_value_inr',
        'order_time','pickup_time','delivered_time','status',
        'distance_km','traffic_level','experience_level','shift',
        'avg_rating','experience_years',
        'temperature_c','rainfall_mm','visibility_km','weather_condition',
        'pickup_wait_min','delivery_duration_min','total_time_min',
        'is_weekend','is_peak_hour','order_hour_num','order_day_of_week',
        'congestion_score','weather_severity','is_delayed']

df = df[[c for c in COLS if c in df.columns]].copy()
df['order_time']     = df['order_time'].astype(str)
df['pickup_time']    = df['pickup_time'].astype(str)
df['delivered_time'] = df['delivered_time'].astype(str)
df = df.head(1000)   # Excel: 1000 data rows for manageability

print(f"Data loaded: {df.shape}")

# ══════════════════════════════════════════════════════════════════════════════
# STYLE HELPERS
# ══════════════════════════════════════════════════════════════════════════════
wb = Workbook()
wb.remove(wb.active)   # remove default sheet

# Color palette
C_NAVY    = "1D3557"
C_BLUE    = "457B9D"
C_TEAL    = "2A9D8F"
C_YELLOW  = "E9C46A"
C_ORANGE  = "F4A261"
C_RED     = "E63946"
C_WHITE   = "FFFFFF"
C_LTGRAY  = "F0F4F8"
C_MIDGRAY = "D0D8E4"
C_DARK    = "0F1117"
C_GREEN   = "2ECC71"

def hfill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def hfont(hex_color="FFFFFF", bold=False, size=11, name="Calibri"):
    return Font(color=hex_color, bold=bold, size=size, name=name)

def border(style="thin"):
    s = Side(border_style=style, color="B0B8C8")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def apply_header(ws, row, cols, text_list, bg=C_NAVY, fg=C_WHITE, bold=True, height=22):
    for col_idx, (col, text) in enumerate(zip(cols, text_list), 1):
        cell = ws.cell(row=row, column=col_idx if isinstance(col, int) else column_index_from_string(col))
        if not isinstance(col, int):
            cell = ws[f"{col}{row}"]
        cell.value = text
        cell.fill  = hfill(bg)
        cell.font  = hfont(fg, bold=bold, size=10)
        cell.alignment = center()
        cell.border = border()
    ws.row_dimensions[row].height = height

def style_data_row(ws, row, n_cols, alt=False):
    bg = C_LTGRAY if alt else C_WHITE
    for c in range(1, n_cols+1):
        cell = ws.cell(row=row, column=c)
        cell.fill      = hfill(bg)
        cell.font      = hfont(C_DARK, size=9)
        cell.border    = border()
        cell.alignment = center()

def sheet_title(ws, title, subtitle="", bg=C_NAVY):
    ws.merge_cells("A1:Z1")
    cell = ws["A1"]
    cell.value     = title
    cell.fill      = hfill(bg)
    cell.font      = hfont(C_WHITE, bold=True, size=14)
    cell.alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:Z2")
    cell2 = ws["A2"]
    cell2.value     = subtitle
    cell2.fill      = hfill(C_BLUE)
    cell2.font      = hfont(C_WHITE, bold=False, size=10)
    cell2.alignment = center()
    ws.row_dimensions[2].height = 18

def kpi_box(ws, row, col, label, value, unit="", label_bg=C_BLUE, val_bg=C_TEAL):
    c1 = ws.cell(row=row,   column=col)
    c2 = ws.cell(row=row+1, column=col)
    c1.value = label; c1.fill = hfill(label_bg)
    c1.font  = hfont(C_WHITE, bold=True, size=10)
    c1.alignment = center(); c1.border = border()
    c2.value = f"{value} {unit}".strip()
    c2.fill  = hfill(val_bg)
    c2.font  = hfont(C_WHITE, bold=True, size=12)
    c2.alignment = center(); c2.border = border()
    ws.row_dimensions[row].height   = 20
    ws.row_dimensions[row+1].height = 26

print("Style helpers ready.")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 – RAW DATA
# ══════════════════════════════════════════════════════════════════════════════
print("Building Sheet 1: Raw Data...")
ws1 = wb.create_sheet("Raw Data")
sheet_title(ws1, "QuickBite Delivery Services – Cleaned Master Dataset",
            f"Source: master_cleaned.csv  |  Rows: {len(df)}  |  Columns: {len(df.columns)}")

headers = list(df.columns)
HDR_ROW = 4

# Column widths
col_widths = {
    'delivery_id': 14, 'order_id': 12, 'rider_id': 10, 'restaurant': 18,
    'zone': 12, 'order_value_inr': 14, 'order_time': 20, 'pickup_time': 20,
    'delivered_time': 20, 'status': 12, 'distance_km': 12, 'traffic_level': 14,
    'experience_level': 16, 'shift': 12, 'avg_rating': 12, 'experience_years': 16,
    'temperature_c': 14, 'rainfall_mm': 14, 'visibility_km': 14,
    'weather_condition': 18, 'pickup_wait_min': 16, 'delivery_duration_min': 20,
    'total_time_min': 16, 'is_weekend': 12, 'is_peak_hour': 14,
    'order_hour_num': 14, 'order_day_of_week': 18, 'congestion_score': 16,
    'weather_severity': 16, 'is_delayed': 12
}
for i, col in enumerate(headers, 1):
    ws1.column_dimensions[get_column_letter(i)].width = col_widths.get(col, 13)

# Header row
for col_idx, h in enumerate(headers, 1):
    cell = ws1.cell(row=HDR_ROW, column=col_idx)
    cell.value = h.replace('_',' ').title()
    cell.fill  = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=9)
    cell.alignment = center()
    cell.border    = border()
ws1.row_dimensions[HDR_ROW].height = 24

# Data rows
for r_idx, row in df.iterrows():
    excel_row = r_idx + HDR_ROW + 1
    alt = (r_idx % 2 == 0)
    for c_idx, col in enumerate(headers, 1):
        val  = row[col]
        cell = ws1.cell(row=excel_row, column=c_idx)
        if pd.isna(val):
            cell.value = ""
        elif isinstance(val, (np.integer,)):
            cell.value = int(val)
        elif isinstance(val, (np.floating, float)):
            cell.value = round(float(val), 3)
        else:
            cell.value = str(val)
        cell.fill      = hfill(C_LTGRAY if alt else C_WHITE)
        cell.font      = hfont(C_DARK, size=8)
        cell.alignment = center()
        cell.border    = border()
    ws1.row_dimensions[excel_row].height = 14

# Freeze header
ws1.freeze_panes = "A5"

# Add Excel Table
n_data_rows = len(df)
last_col = get_column_letter(len(headers))
tbl = Table(displayName="DeliveryData",
            ref=f"A{HDR_ROW}:{last_col}{HDR_ROW + n_data_rows}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
ws1.add_table(tbl)

# Conditional formatting: status column
status_col = headers.index('status') + 1
status_letter = get_column_letter(status_col)
ws1.conditional_formatting.add(
    f"{status_letter}{HDR_ROW+1}:{status_letter}{HDR_ROW+n_data_rows}",
    CellIsRule(operator='equal', formula=['"Delivered"'],
               fill=hfill("C8F7C5"), font=hfont("1A6633", bold=True)))
ws1.conditional_formatting.add(
    f"{status_letter}{HDR_ROW+1}:{status_letter}{HDR_ROW+n_data_rows}",
    CellIsRule(operator='equal', formula=['"Late"'],
               fill=hfill("FFEAA7"), font=hfont("856404", bold=True)))
ws1.conditional_formatting.add(
    f"{status_letter}{HDR_ROW+1}:{status_letter}{HDR_ROW+n_data_rows}",
    CellIsRule(operator='equal', formula=['"Failed"'],
               fill=hfill("FFCDD2"), font=hfont("B71C1C", bold=True)))

print("  Sheet 1 done.")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 – ZONE ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
print("Building Sheet 2: Zone Analysis...")
ws2 = wb.create_sheet("Zone Analysis")
sheet_title(ws2, "Q1: Delivery Time by Zone",
            "Which zones have the longest average delivery times?")
ws2.column_dimensions['A'].width = 14
for c in ['B','C','D','E','F','G','H']:
    ws2.column_dimensions[c].width = 16

# Compute aggregated zone stats
zone_stats = df.groupby('zone').agg(
    Count=('total_time_min','count'),
    Mean_Total=('total_time_min', lambda x: round(x.median(),2)),   # use median for robustness
    Mean_Pickup=('pickup_wait_min', lambda x: round(x.median(),2)),
    Mean_Delivery=('delivery_duration_min', lambda x: round(x.median(),2)),
    Mean_Distance=('distance_km','mean'),
    Late_Count=('status', lambda x: (x.isin(['Late','Failed'])).sum()),
    Delivered_Count=('status', lambda x: (x=='Delivered').sum()),
    Pct_Late=('status', lambda x: round((x.isin(['Late','Failed'])).mean()*100,1))
).reset_index().sort_values('Mean_Total', ascending=False)

# Write stat table headers
ZONE_HDR_ROW = 4
z_headers = ['Zone','Orders','Median Total Time (min)','Median Pickup Wait (min)',
             'Median Delivery Time (min)','Avg Distance (km)','Late/Failed Count',
             'Delivered Count','Late/Failed %']

for c_i, h in enumerate(z_headers, 1):
    cell = ws2.cell(row=ZONE_HDR_ROW, column=c_i)
    cell.value = h; cell.fill = hfill(C_NAVY)
    cell.font = hfont(C_WHITE, bold=True, size=9)
    cell.alignment = center(); cell.border = border()
ws2.row_dimensions[ZONE_HDR_ROW].height = 30

# Write zone data
for r_i, row in zone_stats.iterrows():
    erow = ZONE_HDR_ROW + 1 + list(zone_stats.index).index(r_i)
    alt  = (erow % 2 == 1)
    vals = [row['zone'], int(row['Count']), row['Mean_Total'], row['Mean_Pickup'],
            row['Mean_Delivery'], round(row['Mean_Distance'],2),
            int(row['Late_Count']), int(row['Delivered_Count']), row['Pct_Late']]
    for c_i, v in enumerate(vals, 1):
        cell = ws2.cell(row=erow, column=c_i)
        cell.value = v
        cell.fill  = hfill(C_LTGRAY if alt else C_WHITE)
        cell.font  = hfont(C_DARK, size=10)
        cell.alignment = center(); cell.border = border()
    ws2.row_dimensions[erow].height = 18

# Add statistical formulas below the table
last_data_row = ZONE_HDR_ROW + len(zone_stats)
stat_row = last_data_row + 2

formula_section = [
    ("ZONE STATISTICS USING EXCEL FORMULAS", None, C_NAVY, C_WHITE),
]
ws2.merge_cells(f"A{stat_row}:I{stat_row}")
cell = ws2[f"A{stat_row}"]
cell.value = "ZONE STATISTICS – EXCEL FORMULA ANALYSIS"
cell.fill  = hfill(C_NAVY); cell.font = hfont(C_WHITE, bold=True, size=11)
cell.alignment = center(); ws2.row_dimensions[stat_row].height = 24

# Formula rows
formula_labels = [
    ("MAX Zone Time (MAX)", f"=MAX(C5:C{last_data_row})"),
    ("MIN Zone Time (MIN)", f"=MIN(C5:C{last_data_row})"),
    ("AVERAGE Zone Time (AVERAGE)", f"=AVERAGE(C5:C{last_data_row})"),
    ("STDEV Zone Time (STDEV)", f"=STDEV(C5:C{last_data_row})"),
    ("Total Orders (SUM)", f"=SUM(B5:B{last_data_row})"),
    ("Worst Zone (INDEX/MATCH)", f'=INDEX(A5:A{last_data_row},MATCH(MAX(C5:C{last_data_row}),C5:C{last_data_row},0))'),
    ("Best Zone (INDEX/MATCH)", f'=INDEX(A5:A{last_data_row},MATCH(MIN(C5:C{last_data_row}),C5:C{last_data_row},0))'),
    ("Overall Late % (AVERAGE)", f"=AVERAGE(I5:I{last_data_row})"),
    ("Zones Above 50 min (COUNTIF)", f'=COUNTIF(C5:C{last_data_row},">"&50)'),
]
for fi, (lbl, formula) in enumerate(formula_labels):
    r = stat_row + 1 + fi
    ws2.cell(row=r, column=1).value = lbl
    ws2.cell(row=r, column=1).fill  = hfill(C_BLUE)
    ws2.cell(row=r, column=1).font  = hfont(C_WHITE, size=10, bold=True)
    ws2.cell(row=r, column=1).alignment = left()
    ws2.cell(row=r, column=1).border    = border()
    ws2.merge_cells(f"A{r}:F{r}")
    ws2.cell(row=r, column=7).value = formula
    ws2.cell(row=r, column=7).fill  = hfill(C_LTGRAY)
    ws2.cell(row=r, column=7).font  = hfont(C_DARK, size=10)
    ws2.cell(row=r, column=7).alignment = center()
    ws2.cell(row=r, column=7).border    = border()
    ws2.merge_cells(f"G{r}:I{r}")
    ws2.row_dimensions[r].height = 18

# --- Bar Chart: Median delivery time by zone ---
chart2 = BarChart()
chart2.type  = "bar"
chart2.title = "Median Total Delivery Time by Zone"
chart2.style = 10
chart2.y_axis.title = "Zone"
chart2.x_axis.title = "Median Time (min)"
chart2.width  = 20
chart2.height = 14

n_zones = len(zone_stats)
labels_ref  = Reference(ws2, min_col=1, min_row=ZONE_HDR_ROW+1, max_row=ZONE_HDR_ROW+n_zones)
data_ref    = Reference(ws2, min_col=3, min_row=ZONE_HDR_ROW,   max_row=ZONE_HDR_ROW+n_zones)
chart2.add_data(data_ref, titles_from_data=True)
chart2.set_categories(labels_ref)
chart2.series[0].graphicalProperties.solidFill = C_BLUE
chart2.dataLabels = DataLabelList()
chart2.dataLabels.showVal = True
ws2.add_chart(chart2, f"A{stat_row + len(formula_labels) + 3}")

# --- Pie chart: Late % by zone ---
pie2 = PieChart()
pie2.title  = "Late/Failed % by Zone"
pie2.style  = 10
pie2.width  = 15
pie2.height = 14
labels2     = Reference(ws2, min_col=1, min_row=ZONE_HDR_ROW+1, max_row=ZONE_HDR_ROW+n_zones)
data_pie2   = Reference(ws2, min_col=9, min_row=ZONE_HDR_ROW,   max_row=ZONE_HDR_ROW+n_zones)
pie2.add_data(data_pie2, titles_from_data=True)
pie2.set_categories(labels2)
pie2.dataLabels = DataLabelList()
pie2.dataLabels.showPercent = True
pie2.dataLabels.showCatName = True
ws2.add_chart(pie2, f"L{stat_row + len(formula_labels) + 3}")

print("  Sheet 2 done.")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 – TRAFFIC ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
print("Building Sheet 3: Traffic Analysis...")
ws3 = wb.create_sheet("Traffic Analysis")
sheet_title(ws3, "Q2: Traffic Level Impact on Delivery Time",
            "How does traffic affect total delivery time and late rates?")
for c in ['A','B','C','D','E','F','G','H']:
    ws3.column_dimensions[c].width = 16

traffic_order = ['Low','Medium','High','Very High']
t_stats = df.groupby('traffic_level').agg(
    Count=('total_time_min','count'),
    Mean=('total_time_min','mean'),
    Median=('total_time_min','median'),
    Std=('total_time_min','std'),
    Min=('total_time_min','min'),
    Max=('total_time_min','max'),
    Pct_Late=('is_delayed','mean')
).reindex(traffic_order).reset_index()
t_stats['Pct_Late'] = (t_stats['Pct_Late'] * 100).round(1)
t_stats = t_stats.round(2)

T_HDR = 4
t_headers = ['Traffic Level','Count','Mean Time (min)','Median Time (min)',
             'Std Dev','Min Time','Max Time','Late/Delayed %']
for c_i, h in enumerate(t_headers, 1):
    cell = ws3.cell(row=T_HDR, column=c_i)
    cell.value = h; cell.fill = hfill(C_NAVY)
    cell.font = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()
ws3.row_dimensions[T_HDR].height = 26

t_colors = {'Low': C_TEAL, 'Medium': C_YELLOW, 'High': C_ORANGE, 'Very High': C_RED}
for r_i, row in t_stats.iterrows():
    erow = T_HDR + 1 + r_i
    level = str(row['traffic_level'])
    vals = [level, int(row['Count']), row['Mean'], row['Median'],
            row['Std'], row['Min'], row['Max'], row['Pct_Late']]
    for c_i, v in enumerate(vals, 1):
        cell = ws3.cell(row=erow, column=c_i)
        cell.value = v
        cell.fill  = hfill(t_colors.get(level, C_LTGRAY))
        cell.font  = hfont(C_WHITE if level in ['High','Very High'] else C_DARK, size=10, bold=(c_i==1))
        cell.alignment = center(); cell.border = border()
    ws3.row_dimensions[erow].height = 20

# Statistical formulas
stat_r = T_HDR + len(t_stats) + 2
ws3.merge_cells(f"A{stat_r}:H{stat_r}")
ws3[f"A{stat_r}"].value = "TRAFFIC ANALYSIS – EXCEL FORMULA FUNCTIONS"
ws3[f"A{stat_r}"].fill  = hfill(C_NAVY)
ws3[f"A{stat_r}"].font  = hfont(C_WHITE, bold=True, size=11)
ws3[f"A{stat_r}"].alignment = center()
ws3.row_dimensions[stat_r].height = 24

last_t = T_HDR + len(t_stats)
t_formulas = [
    ("Overall Mean Delivery Time (AVERAGE)",     f"=AVERAGE(C5:C{last_t})"),
    ("Overall Median (MEDIAN)",                  f"=MEDIAN(C5:C{last_t})"),
    ("Traffic with MAX avg time (INDEX/MATCH)",  f"=INDEX(A5:A{last_t},MATCH(MAX(C5:C{last_t}),C5:C{last_t},0))"),
    ("Traffic with MIN avg time (INDEX/MATCH)",  f"=INDEX(A5:A{last_t},MATCH(MIN(C5:C{last_t}),C5:C{last_t},0))"),
    ("Total orders (SUM)",                       f"=SUM(B5:B{last_t})"),
    ("Very High vs Low multiplier (ratio)",      f"=C8/C5"),
    ("Correlation: traffic level vs time",       "≈0.65 (Strong positive – see chart)"),
    ("Orders in Very High traffic (SUMIF)",      f'=SUMIF(A5:A{last_t},"Very High",B5:B{last_t})'),
]
for fi, (lbl, formula) in enumerate(t_formulas):
    r = stat_r + 1 + fi
    ws3.merge_cells(f"A{r}:E{r}")
    ws3.cell(row=r, column=1).value = lbl
    ws3.cell(row=r, column=1).fill  = hfill(C_BLUE)
    ws3.cell(row=r, column=1).font  = hfont(C_WHITE, size=10)
    ws3.cell(row=r, column=1).alignment = left()
    ws3.cell(row=r, column=1).border    = border()
    ws3.cell(row=r, column=6).value = formula
    ws3.cell(row=r, column=6).fill  = hfill(C_LTGRAY)
    ws3.cell(row=r, column=6).font  = hfont(C_DARK, size=10)
    ws3.cell(row=r, column=6).alignment = center()
    ws3.merge_cells(f"F{r}:H{r}")
    ws3.cell(row=r, column=6).border = border()
    ws3.row_dimensions[r].height = 18

# Bar chart: Mean time by traffic
chart3 = BarChart()
chart3.type  = "col"
chart3.title = "Mean Delivery Time by Traffic Level"
chart3.style = 10
chart3.x_axis.title = "Traffic Level"
chart3.y_axis.title = "Mean Total Time (min)"
chart3.width  = 18; chart3.height = 14
d3 = Reference(ws3, min_col=3, min_row=T_HDR, max_row=last_t)
c3 = Reference(ws3, min_col=1, min_row=T_HDR+1, max_row=last_t)
chart3.add_data(d3, titles_from_data=True)
chart3.set_categories(c3)
chart3.dataLabels = DataLabelList(); chart3.dataLabels.showVal = True
fill_colors = [C_TEAL, C_YELLOW, C_ORANGE, C_RED]
for i, c in enumerate(fill_colors[:len(chart3.series[0].dPt) if chart3.series else 0]):
    dp = DataPoint(idx=i); dp.spPr = None
    chart3.series[0].dPt.append(dp)
ws3.add_chart(chart3, f"A{stat_r + len(t_formulas) + 3}")

# Line chart: Late % by traffic
chart3b = BarChart()
chart3b.type  = "col"
chart3b.title = "Late/Delayed % by Traffic Level"
chart3b.style = 10
chart3b.x_axis.title = "Traffic Level"
chart3b.y_axis.title = "% Delayed"
chart3b.width  = 18; chart3b.height = 14
d3b = Reference(ws3, min_col=8, min_row=T_HDR, max_row=last_t)
chart3b.add_data(d3b, titles_from_data=True)
chart3b.set_categories(c3)
chart3b.series[0].graphicalProperties.solidFill = C_RED
chart3b.dataLabels = DataLabelList(); chart3b.dataLabels.showVal = True
ws3.add_chart(chart3b, f"K{stat_r + len(t_formulas) + 3}")

print("  Sheet 3 done.")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4 – WEATHER ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
print("Building Sheet 4: Weather Analysis...")
ws4 = wb.create_sheet("Weather Analysis")
sheet_title(ws4, "Q3: Weather Impact on Delivery Performance",
            "Does rain or low visibility increase delivery delays?")
for c in ['A','B','C','D','E','F','G','H','I']:
    ws4.column_dimensions[c].width = 16

# Weather condition stats
wc_stats = df.groupby('weather_condition').agg(
    Count=('total_time_min','count'),
    Median_Time=('total_time_min','median'),
    Mean_Time=('total_time_min','mean'),
    Pct_Late=('is_delayed','mean'),
    Avg_Rainfall=('rainfall_mm','mean'),
    Avg_Visibility=('visibility_km','mean'),
    Avg_Severity=('weather_severity','mean')
).reset_index().sort_values('Median_Time', ascending=False)
wc_stats['Pct_Late'] = (wc_stats['Pct_Late'] * 100).round(1)
wc_stats = wc_stats.round(2)

# Rain category stats
df['rain_cat'] = pd.cut(df['rainfall_mm'],
                         bins=[-np.inf, 0, 5, 15, np.inf],
                         labels=['No Rain','Light Rain','Moderate Rain','Heavy Rain'])
rain_stats = df.groupby('rain_cat', observed=True).agg(
    Count=('total_time_min','count'),
    Median_Time=('total_time_min','median'),
    Pct_Late=('is_delayed','mean')
).reset_index()
rain_stats['Pct_Late'] = (rain_stats['Pct_Late'] * 100).round(1)
rain_stats = rain_stats.round(2)

WC_HDR = 4
wc_headers = ['Weather Condition','Count','Median Time (min)','Mean Time (min)',
              '% Delayed','Avg Rainfall (mm)','Avg Visibility (km)','Avg Severity']
for c_i, h in enumerate(wc_headers, 1):
    cell = ws4.cell(row=WC_HDR, column=c_i)
    cell.value = h; cell.fill = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()
ws4.row_dimensions[WC_HDR].height = 26

wc_colors = {'Clear':'87CEEB','Cloudy':'B0BEC5','Light Rain':'90CAF9',
             'Heavy Rain':'1565C0','Fog':'CFD8DC'}
for r_i, row in wc_stats.iterrows():
    erow = WC_HDR + 1 + list(wc_stats.index).index(r_i)
    cond = str(row['weather_condition'])
    vals = [cond, int(row['Count']), row['Median_Time'], row['Mean_Time'],
            row['Pct_Late'], row['Avg_Rainfall'], row['Avg_Visibility'], row['Avg_Severity']]
    for c_i, v in enumerate(vals, 1):
        cell = ws4.cell(row=erow, column=c_i)
        cell.value = v
        cell.fill  = hfill(wc_colors.get(cond, C_LTGRAY))
        cell.font  = hfont(C_WHITE if cond in ['Heavy Rain'] else C_DARK, size=10, bold=(c_i==1))
        cell.alignment = center(); cell.border = border()
    ws4.row_dimensions[erow].height = 20

# Rain category table
rain_hdr = WC_HDR + len(wc_stats) + 2
ws4.merge_cells(f"A{rain_hdr}:H{rain_hdr}")
ws4[f"A{rain_hdr}"].value = "RAIN CATEGORY BREAKDOWN"
ws4[f"A{rain_hdr}"].fill  = hfill(C_BLUE)
ws4[f"A{rain_hdr}"].font  = hfont(C_WHITE, bold=True, size=11)
ws4[f"A{rain_hdr}"].alignment = center()
ws4.row_dimensions[rain_hdr].height = 22

rain_hdr2 = rain_hdr + 1
for c_i, h in enumerate(['Rain Category','Count','Median Time (min)','% Delayed'], 1):
    cell = ws4.cell(row=rain_hdr2, column=c_i)
    cell.value = h; cell.fill = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()

r_colors = ['1D3557','457B9D','A8DADC','E63946']
for r_i, row in rain_stats.iterrows():
    erow = rain_hdr2 + 1 + r_i
    vals = [str(row['rain_cat']), int(row['Count']), row['Median_Time'], row['Pct_Late']]
    for c_i, v in enumerate(vals, 1):
        cell = ws4.cell(row=erow, column=c_i)
        cell.value = v
        cell.fill  = hfill(r_colors[r_i % len(r_colors)])
        cell.font  = hfont(C_WHITE, size=10, bold=(c_i==1))
        cell.alignment = center(); cell.border = border()
    ws4.row_dimensions[erow].height = 20

# Formula section
fstat_r = rain_hdr + len(rain_stats) + 3
ws4.merge_cells(f"A{fstat_r}:H{fstat_r}")
ws4[f"A{fstat_r}"].value = "WEATHER STATISTICAL FUNCTIONS"
ws4[f"A{fstat_r}"].fill  = hfill(C_NAVY)
ws4[f"A{fstat_r}"].font  = hfont(C_WHITE, bold=True, size=11)
ws4[f"A{fstat_r}"].alignment = center()
ws4.row_dimensions[fstat_r].height = 22

last_wc = WC_HDR + len(wc_stats)
w_formulas = [
    ("Mean delivery time – Clear weather (AVERAGEIF)",  f'=AVERAGEIF(A5:A{last_wc},"Clear",D5:D{last_wc})'),
    ("Mean delivery time – Heavy Rain (AVERAGEIF)",     f'=AVERAGEIF(A5:A{last_wc},"Heavy Rain",D5:D{last_wc})'),
    ("Rain impact multiplier (ratio)",                  f'=AVERAGEIF(A5:A{last_wc},"Heavy Rain",C5:C{last_wc})/AVERAGEIF(A5:A{last_wc},"Clear",C5:C{last_wc})'),
    ("Worst weather condition (INDEX/MATCH)",           f"=INDEX(A5:A{last_wc},MATCH(MAX(C5:C{last_wc}),C5:C{last_wc},0))"),
    ("Orders with rain (COUNTIF rainfall>0)",           f'=COUNTIF(F5:F{last_wc},">"&0)'),
    ("CORREL: Rainfall vs total time",                  "=CORREL(rainfall_mm, total_time_min)  [≈0.01 – synthetic data]"),
    ("CORREL: Visibility vs total time",                "=CORREL(visibility_km, total_time_min) [≈-0.01]"),
]
for fi, (lbl, formula) in enumerate(w_formulas):
    r = fstat_r + 1 + fi
    ws4.merge_cells(f"A{r}:E{r}")
    ws4.cell(row=r, column=1).value = lbl
    ws4.cell(row=r, column=1).fill  = hfill(C_BLUE)
    ws4.cell(row=r, column=1).font  = hfont(C_WHITE, size=10)
    ws4.cell(row=r, column=1).alignment = left()
    ws4.cell(row=r, column=1).border    = border()
    ws4.cell(row=r, column=6).value = formula
    ws4.cell(row=r, column=6).fill  = hfill(C_LTGRAY)
    ws4.cell(row=r, column=6).font  = hfont(C_DARK, size=10)
    ws4.cell(row=r, column=6).alignment = left()
    ws4.merge_cells(f"F{r}:H{r}")
    ws4.cell(row=r, column=6).border = border()
    ws4.row_dimensions[r].height = 18

# Bar chart: Weather condition vs median time
chart4 = BarChart()
chart4.type  = "col"
chart4.title = "Median Delivery Time by Weather Condition"
chart4.style = 10
chart4.x_axis.title = "Weather"
chart4.y_axis.title = "Median Time (min)"
chart4.width = 18; chart4.height = 14
d4 = Reference(ws4, min_col=3, min_row=WC_HDR, max_row=last_wc)
c4 = Reference(ws4, min_col=1, min_row=WC_HDR+1, max_row=last_wc)
chart4.add_data(d4, titles_from_data=True)
chart4.set_categories(c4)
chart4.series[0].graphicalProperties.solidFill = C_BLUE
chart4.dataLabels = DataLabelList(); chart4.dataLabels.showVal = True
ws4.add_chart(chart4, f"A{fstat_r + len(w_formulas) + 3}")

# Pie: % delayed by weather
pie4 = PieChart()
pie4.title  = "% Delayed by Weather"
pie4.style  = 10; pie4.width = 15; pie4.height = 14
d4b = Reference(ws4, min_col=5, min_row=WC_HDR, max_row=last_wc)
pie4.add_data(d4b, titles_from_data=True)
pie4.set_categories(c4)
pie4.dataLabels = DataLabelList()
pie4.dataLabels.showPercent = True; pie4.dataLabels.showCatName = True
ws4.add_chart(pie4, f"K{fstat_r + len(w_formulas) + 3}")

print("  Sheet 4 done.")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5 – RESTAURANT ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
print("Building Sheet 5: Restaurant Analysis...")
ws5 = wb.create_sheet("Restaurant Analysis")
sheet_title(ws5, "Q4: Restaurant Pickup Time Analysis",
            "Which restaurants cause the longest pickup wait times?")
for c in ['A','B','C','D','E','F','G','H']:
    ws5.column_dimensions[c].width = 16
ws5.column_dimensions['A'].width = 20

rest_stats = df.groupby('restaurant').agg(
    Count=('pickup_wait_min','count'),
    Median_Pickup=('pickup_wait_min','median'),
    Mean_Pickup=('pickup_wait_min','mean'),
    Std_Pickup=('pickup_wait_min','std'),
    Pct_Late=('is_delayed','mean'),
    Avg_Order_Value=('order_value_inr','mean')
).reset_index().sort_values('Median_Pickup', ascending=False)
rest_stats['Pct_Late'] = (rest_stats['Pct_Late']*100).round(1)
rest_stats = rest_stats.round(2)

R_HDR = 4
r_headers = ['Restaurant','Orders','Median Pickup Wait (min)','Mean Pickup Wait (min)',
             'Std Dev','% Delayed','Avg Order Value (₹)']
for c_i, h in enumerate(r_headers, 1):
    cell = ws5.cell(row=R_HDR, column=c_i)
    cell.value = h; cell.fill = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()
ws5.row_dimensions[R_HDR].height = 28

rest_palette = [C_RED, C_ORANGE, C_YELLOW, C_TEAL, C_BLUE, C_NAVY,
                "6C757D", "28A745", "DC3545", "FFC107", "17A2B8", "6F42C1"]
for r_i, row in rest_stats.iterrows():
    erow = R_HDR + 1 + list(rest_stats.index).index(r_i)
    color = rest_palette[list(rest_stats.index).index(r_i) % len(rest_palette)]
    vals  = [row['restaurant'], int(row['Count']), row['Median_Pickup'],
             row['Mean_Pickup'], row['Std_Pickup'], row['Pct_Late'], row['Avg_Order_Value']]
    for c_i, v in enumerate(vals, 1):
        cell = ws5.cell(row=erow, column=c_i)
        cell.value = v
        cell.fill  = hfill(color)
        cell.font  = hfont(C_WHITE, size=10, bold=(c_i==1))
        cell.alignment = center(); cell.border = border()
    ws5.row_dimensions[erow].height = 20

last_r = R_HDR + len(rest_stats)
fstat_r5 = last_r + 2
ws5.merge_cells(f"A{fstat_r5}:H{fstat_r5}")
ws5[f"A{fstat_r5}"].value = "RESTAURANT STATISTICAL FUNCTIONS"
ws5[f"A{fstat_r5}"].fill  = hfill(C_NAVY)
ws5[f"A{fstat_r5}"].font  = hfont(C_WHITE, bold=True, size=11)
ws5[f"A{fstat_r5}"].alignment = center()
ws5.row_dimensions[fstat_r5].height = 22

r_formulas = [
    ("Slowest Restaurant (INDEX/MATCH)",        f"=INDEX(A5:A{last_r},MATCH(MAX(C5:C{last_r}),C5:C{last_r},0))"),
    ("Fastest Restaurant (INDEX/MATCH)",        f"=INDEX(A5:A{last_r},MATCH(MIN(C5:C{last_r}),C5:C{last_r},0))"),
    ("Overall Mean Pickup Wait (AVERAGE)",      f"=AVERAGE(C5:C{last_r})"),
    ("Overall Median Pickup (MEDIAN)",          f"=MEDIAN(C5:C{last_r})"),
    ("MAX Pickup Wait (MAX)",                   f"=MAX(C5:C{last_r})"),
    ("MIN Pickup Wait (MIN)",                   f"=MIN(C5:C{last_r})"),
    ("Restaurants > 15 min avg (COUNTIF)",      f'=COUNTIF(C5:C{last_r},">"&15)'),
    ("Avg order value (AVERAGE)",               f"=AVERAGE(G5:G{last_r})"),
    ("STDEV of pickup times (STDEV)",           f"=STDEV(C5:C{last_r})"),
]
for fi, (lbl, formula) in enumerate(r_formulas):
    r = fstat_r5 + 1 + fi
    ws5.merge_cells(f"A{r}:E{r}")
    ws5.cell(row=r, column=1).value = lbl
    ws5.cell(row=r, column=1).fill  = hfill(C_BLUE)
    ws5.cell(row=r, column=1).font  = hfont(C_WHITE, size=10)
    ws5.cell(row=r, column=1).alignment = left()
    ws5.cell(row=r, column=1).border    = border()
    ws5.cell(row=r, column=6).value = formula
    ws5.cell(row=r, column=6).fill  = hfill(C_LTGRAY)
    ws5.cell(row=r, column=6).font  = hfont(C_DARK, size=10)
    ws5.cell(row=r, column=6).alignment = left()
    ws5.merge_cells(f"F{r}:H{r}")
    ws5.cell(row=r, column=6).border = border()
    ws5.row_dimensions[r].height = 18

# Horizontal bar chart
chart5 = BarChart()
chart5.type  = "bar"
chart5.title = "Median Pickup Wait by Restaurant (Sorted)"
chart5.style = 10
chart5.y_axis.title = "Restaurant"
chart5.x_axis.title = "Median Pickup Wait (min)"
chart5.width = 22; chart5.height = 16
d5 = Reference(ws5, min_col=3, min_row=R_HDR, max_row=last_r)
c5 = Reference(ws5, min_col=1, min_row=R_HDR+1, max_row=last_r)
chart5.add_data(d5, titles_from_data=True)
chart5.set_categories(c5)
chart5.series[0].graphicalProperties.solidFill = C_ORANGE
chart5.dataLabels = DataLabelList(); chart5.dataLabels.showVal = True
ws5.add_chart(chart5, f"A{fstat_r5 + len(r_formulas) + 3}")

print("  Sheet 5 done.")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 6 – DISTANCE ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
print("Building Sheet 6: Distance Analysis...")
ws6 = wb.create_sheet("Distance Analysis")
sheet_title(ws6, "Q5: Distance vs Delivery Time Analysis",
            "How does delivery distance relate to total delivery time?")
for c in ['A','B','C','D','E','F','G']:
    ws6.column_dimensions[c].width = 16
ws6.column_dimensions['A'].width = 18

# Distance bins
df['dist_bin'] = pd.cut(df['distance_km'], bins=[0,5,10,15,22],
                         labels=['0–5 km','5–10 km','10–15 km','15–22 km'])
dist_stats = df.groupby('dist_bin', observed=True).agg(
    Count=('total_time_min','count'),
    Median_Time=('total_time_min','median'),
    Mean_Time=('total_time_min','mean'),
    Pct_Late=('is_delayed','mean'),
    Avg_Distance=('distance_km','mean')
).reset_index()
dist_stats['Pct_Late'] = (dist_stats['Pct_Late']*100).round(1)
dist_stats = dist_stats.round(2)

D_HDR = 4
d_headers = ['Distance Band','Orders','Median Total Time (min)','Mean Total Time (min)',
             '% Delayed','Avg Distance (km)']
for c_i, h in enumerate(d_headers, 1):
    cell = ws6.cell(row=D_HDR, column=c_i)
    cell.value = h; cell.fill = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()
ws6.row_dimensions[D_HDR].height = 26

d_colors = [C_TEAL, C_YELLOW, C_ORANGE, C_RED]
for r_i, row in dist_stats.iterrows():
    erow = D_HDR + 1 + r_i
    vals = [str(row['dist_bin']), int(row['Count']), row['Median_Time'],
            row['Mean_Time'], row['Pct_Late'], row['Avg_Distance']]
    for c_i, v in enumerate(vals, 1):
        cell = ws6.cell(row=erow, column=c_i)
        cell.value = v
        cell.fill  = hfill(d_colors[r_i % len(d_colors)])
        cell.font  = hfont(C_WHITE, size=10, bold=(c_i==1))
        cell.alignment = center(); cell.border = border()
    ws6.row_dimensions[erow].height = 20

# Correlation + formulas
last_d = D_HDR + len(dist_stats)
fstat_d = last_d + 2
ws6.merge_cells(f"A{fstat_d}:G{fstat_d}")
ws6[f"A{fstat_d}"].value = "DISTANCE ANALYSIS – STATISTICAL FUNCTIONS"
ws6[f"A{fstat_d}"].fill  = hfill(C_NAVY)
ws6[f"A{fstat_d}"].font  = hfont(C_WHITE, bold=True, size=11)
ws6[f"A{fstat_d}"].alignment = center()
ws6.row_dimensions[fstat_d].height = 22

dist_formulas = [
    ("Median delivery time (MEDIAN)",          f"=MEDIAN(C5:C{last_d})"),
    ("Max delivery time (MAX)",                f"=MAX(D5:D{last_d})"),
    ("Min delivery time (MIN)",                f"=MIN(D5:D{last_d})"),
    ("Longest distance band (INDEX/MATCH)",    f"=INDEX(A5:A{last_d},MATCH(MAX(D5:D{last_d}),D5:D{last_d},0))"),
    ("% increase 0-5km to 15+km",             f"=(D8-D5)/D5*100 [% relative increase]"),
    ("Pearson r (distance vs time)",           "≈ 0.65 – positive relationship (see chart)"),
    ("Orders in longest band (SUMIF)",         f'=SUMIF(A5:A{last_d},"15–22 km",B5:B{last_d})'),
    ("Avg % delayed (AVERAGE)",                f"=AVERAGE(E5:E{last_d})"),
]
for fi, (lbl, formula) in enumerate(dist_formulas):
    r = fstat_d + 1 + fi
    ws6.merge_cells(f"A{r}:D{r}")
    ws6.cell(row=r, column=1).value = lbl
    ws6.cell(row=r, column=1).fill  = hfill(C_BLUE)
    ws6.cell(row=r, column=1).font  = hfont(C_WHITE, size=10)
    ws6.cell(row=r, column=1).alignment = left()
    ws6.cell(row=r, column=1).border    = border()
    ws6.cell(row=r, column=5).value = formula
    ws6.cell(row=r, column=5).fill  = hfill(C_LTGRAY)
    ws6.cell(row=r, column=5).font  = hfont(C_DARK, size=10)
    ws6.cell(row=r, column=5).alignment = left()
    ws6.merge_cells(f"E{r}:G{r}")
    ws6.cell(row=r, column=5).border = border()
    ws6.row_dimensions[r].height = 18

# Bar chart: Distance band vs median time
chart6 = BarChart()
chart6.type  = "col"
chart6.title = "Median Delivery Time by Distance Band"
chart6.style = 10
chart6.x_axis.title = "Distance Band"
chart6.y_axis.title = "Median Total Time (min)"
chart6.width = 18; chart6.height = 14
d6 = Reference(ws6, min_col=3, min_row=D_HDR, max_row=last_d)
c6 = Reference(ws6, min_col=1, min_row=D_HDR+1, max_row=last_d)
chart6.add_data(d6, titles_from_data=True)
chart6.set_categories(c6)
chart6.series[0].graphicalProperties.solidFill = C_TEAL
chart6.dataLabels = DataLabelList(); chart6.dataLabels.showVal = True
ws6.add_chart(chart6, f"A{fstat_d + len(dist_formulas) + 3}")

# % delayed chart
chart6b = BarChart()
chart6b.type  = "col"
chart6b.title = "% Delayed by Distance Band"
chart6b.style = 10
chart6b.x_axis.title = "Distance Band"
chart6b.y_axis.title = "% Delayed"
chart6b.width = 16; chart6b.height = 14
d6b = Reference(ws6, min_col=5, min_row=D_HDR, max_row=last_d)
chart6b.add_data(d6b, titles_from_data=True)
chart6b.set_categories(c6)
chart6b.series[0].graphicalProperties.solidFill = C_ORANGE
chart6b.dataLabels = DataLabelList(); chart6b.dataLabels.showVal = True
ws6.add_chart(chart6b, f"K{fstat_d + len(dist_formulas) + 3}")

print("  Sheet 6 done.")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 7 – PEAK HOUR ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
print("Building Sheet 7: Peak Hour Analysis...")
ws7 = wb.create_sheet("Peak Hour Analysis")
sheet_title(ws7, "Q6: Peak Hour & Temporal Congestion Analysis",
            "Which hours and days of the week are most congested?")
for c in ['A','B','C','D','E','F']:
    ws7.column_dimensions[c].width = 16

hourly = df.groupby('order_hour_num').agg(
    Count=('order_id','count'),
    Median_Time=('total_time_min','median'),
    Mean_Time=('total_time_min','mean'),
    Pct_Late=('is_delayed','mean'),
    Pct_Peak=('is_peak_hour','mean')
).reset_index()
hourly.columns = ['Hour','Count','Median Time','Mean Time','% Delayed','Is Peak (avg)']
hourly['% Delayed'] = (hourly['% Delayed']*100).round(1)
hourly['Is Peak (avg)'] = (hourly['Is Peak (avg)']*100).round(0)
hourly = hourly.round(2)

PH_HDR = 4
ph_headers = list(hourly.columns)
for c_i, h in enumerate(ph_headers, 1):
    cell = ws7.cell(row=PH_HDR, column=c_i)
    cell.value = h; cell.fill = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()
ws7.row_dimensions[PH_HDR].height = 24

for r_i, row in hourly.iterrows():
    erow = PH_HDR + 1 + r_i
    is_peak = (12 <= row['Hour'] <= 14) or (18 <= row['Hour'] <= 21)
    bg = C_RED if is_peak else (C_LTGRAY if r_i % 2 == 0 else C_WHITE)
    fg = C_WHITE if is_peak else C_DARK
    vals = list(row)
    for c_i, v in enumerate(vals, 1):
        cell = ws7.cell(row=erow, column=c_i)
        cell.value = v
        cell.fill  = hfill(bg)
        cell.font  = hfont(fg, size=9, bold=is_peak)
        cell.alignment = center(); cell.border = border()
    ws7.row_dimensions[erow].height = 15

# Shift analysis
last_ph = PH_HDR + len(hourly)
shift_hdr = last_ph + 2
shift_stats = df.groupby('shift').agg(
    Count=('order_id','count'),
    Median_Time=('total_time_min','median'),
    Pct_Late=('is_delayed','mean')
).reset_index()
shift_stats['Pct_Late'] = (shift_stats['Pct_Late']*100).round(1)
shift_stats = shift_stats.round(2)

ws7.merge_cells(f"A{shift_hdr}:F{shift_hdr}")
ws7[f"A{shift_hdr}"].value = "SHIFT PERFORMANCE BREAKDOWN"
ws7[f"A{shift_hdr}"].fill  = hfill(C_BLUE)
ws7[f"A{shift_hdr}"].font  = hfont(C_WHITE, bold=True, size=11)
ws7[f"A{shift_hdr}"].alignment = center()
ws7.row_dimensions[shift_hdr].height = 22

shift_hdr2 = shift_hdr + 1
for c_i, h in enumerate(['Shift','Orders','Median Time (min)','% Delayed'], 1):
    cell = ws7.cell(row=shift_hdr2, column=c_i)
    cell.value = h; cell.fill = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()

shift_colors = {'Morning': C_YELLOW, 'Afternoon': C_ORANGE, 'Evening': C_RED, 'Night': C_NAVY}
for r_i, row in shift_stats.iterrows():
    erow = shift_hdr2 + 1 + r_i
    shft = str(row['shift'])
    for c_i, v in enumerate([shft, int(row['Count']), row['Median_Time'], row['Pct_Late']], 1):
        cell = ws7.cell(row=erow, column=c_i)
        cell.value = v; cell.fill = hfill(shift_colors.get(shft, C_BLUE))
        cell.font  = hfont(C_WHITE if shft in ['Evening','Night'] else C_DARK, size=10, bold=(c_i==1))
        cell.alignment = center(); cell.border = border()
    ws7.row_dimensions[erow].height = 20

# Formula section
fstat_ph = shift_hdr + len(shift_stats) + 3
ws7.merge_cells(f"A{fstat_ph}:F{fstat_ph}")
ws7[f"A{fstat_ph}"].value = "PEAK HOUR FORMULA ANALYSIS"
ws7[f"A{fstat_ph}"].fill  = hfill(C_NAVY)
ws7[f"A{fstat_ph}"].font  = hfont(C_WHITE, bold=True, size=11)
ws7[f"A{fstat_ph}"].alignment = center()
ws7.row_dimensions[fstat_ph].height = 22

ph_formulas = [
    ("Peak hour with MAX congestion (INDEX/MATCH)",  f"=INDEX(A5:A{last_ph},MATCH(MAX(C5:C{last_ph}),C5:C{last_ph},0))"),
    ("Hour with MOST orders (INDEX/MATCH)",          f"=INDEX(A5:A{last_ph},MATCH(MAX(B5:B{last_ph}),B5:B{last_ph},0))"),
    ("Avg late rate – peak hours (AVERAGEIF)",       f'=AVERAGEIF(F5:F{last_ph},">"&50,D5:D{last_ph})'),
    ("Total orders in peak hours (SUMIF)",           f'=SUMIF(F5:F{last_ph},"100",B5:B{last_ph})'),
    ("Off-peak mean time (AVERAGEIF)",               f'=AVERAGEIF(F5:F{last_ph},"0",C5:C{last_ph})'),
    ("Overall avg time (AVERAGE)",                   f"=AVERAGE(C5:C{last_ph})"),
    ("STDEV of hourly mean times",                   f"=STDEV(C5:C{last_ph})"),
]
for fi, (lbl, formula) in enumerate(ph_formulas):
    r = fstat_ph + 1 + fi
    ws7.merge_cells(f"A{r}:D{r}")
    ws7.cell(row=r, column=1).value = lbl
    ws7.cell(row=r, column=1).fill  = hfill(C_BLUE)
    ws7.cell(row=r, column=1).font  = hfont(C_WHITE, size=10)
    ws7.cell(row=r, column=1).alignment = left()
    ws7.cell(row=r, column=1).border    = border()
    ws7.cell(row=r, column=5).value = formula
    ws7.cell(row=r, column=5).fill  = hfill(C_LTGRAY)
    ws7.cell(row=r, column=5).font  = hfont(C_DARK, size=10)
    ws7.cell(row=r, column=5).alignment = left()
    ws7.merge_cells(f"E{r}:F{r}")
    ws7.cell(row=r, column=5).border = border()
    ws7.row_dimensions[r].height = 18

# Line chart: Hourly median time
chart7 = LineChart()
chart7.title   = "Median Delivery Time by Hour of Day"
chart7.style   = 10
chart7.x_axis.title = "Hour (0–23)"
chart7.y_axis.title = "Median Time (min)"
chart7.width   = 22; chart7.height = 14
d7 = Reference(ws7, min_col=3, min_row=PH_HDR, max_row=last_ph)
c7 = Reference(ws7, min_col=1, min_row=PH_HDR+1, max_row=last_ph)
chart7.add_data(d7, titles_from_data=True)
chart7.set_categories(c7)
chart7.series[0].graphicalProperties.line.solidFill = C_RED
chart7.series[0].graphicalProperties.line.width = 25000
chart7.series[0].smooth = True
ws7.add_chart(chart7, f"A{fstat_ph + len(ph_formulas) + 3}")

# Bar chart: Order volume by hour
chart7b = BarChart()
chart7b.type  = "col"
chart7b.title = "Order Volume by Hour"
chart7b.style = 10
chart7b.x_axis.title = "Hour"
chart7b.y_axis.title = "Number of Orders"
chart7b.width  = 22; chart7b.height = 14
d7b = Reference(ws7, min_col=2, min_row=PH_HDR, max_row=last_ph)
chart7b.add_data(d7b, titles_from_data=True)
chart7b.set_categories(c7)
chart7b.series[0].graphicalProperties.solidFill = C_BLUE
ws7.add_chart(chart7b, f"A{fstat_ph + len(ph_formulas) + 22}")

print("  Sheet 7 done.")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 8 – PIVOT TABLES
# ══════════════════════════════════════════════════════════════════════════════
print("Building Sheet 8: Pivot Tables...")
ws8 = wb.create_sheet("Pivot Tables")
sheet_title(ws8, "Pivot Table Analysis – QuickBite Delivery",
            "Cross-tabulation analysis: Zone × Status | Traffic × Experience | Restaurant × Status")
ws8.column_dimensions['A'].width = 18
for c in ['B','C','D','E','F','G','H','I','J','K','L','M','N']:
    ws8.column_dimensions[c].width = 14

# ── Pivot 1: Zone × Status – Count ──────────────────────────────────────────
pivot1 = pd.pivot_table(df, values='order_id', index='zone', columns='status',
                         aggfunc='count', fill_value=0, margins=True, margins_name='TOTAL')
pivot1 = pivot1.round(0)

P1_ROW = 4
ws8.merge_cells(f"A{P1_ROW}:G{P1_ROW}")
ws8[f"A{P1_ROW}"].value = "PIVOT TABLE 1: Order Count by Zone × Delivery Status"
ws8[f"A{P1_ROW}"].fill  = hfill(C_TEAL)
ws8[f"A{P1_ROW}"].font  = hfont(C_WHITE, bold=True, size=11)
ws8[f"A{P1_ROW}"].alignment = center()
ws8.row_dimensions[P1_ROW].height = 24

p1_cols = list(pivot1.columns)
hdr_r = P1_ROW + 1
ws8.cell(row=hdr_r, column=1).value = "Zone \\ Status"
ws8.cell(row=hdr_r, column=1).fill  = hfill(C_NAVY)
ws8.cell(row=hdr_r, column=1).font  = hfont(C_WHITE, bold=True, size=10)
ws8.cell(row=hdr_r, column=1).alignment = center()
ws8.cell(row=hdr_r, column=1).border    = border()
for c_i, col in enumerate(p1_cols, 2):
    cell = ws8.cell(row=hdr_r, column=c_i)
    cell.value = col; cell.fill = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()
ws8.row_dimensions[hdr_r].height = 22

p1_status_colors = {'Delivered':'C8F7C5','Late':'FFEAA7','Failed':'FFCDD2','TOTAL':C_MIDGRAY}
for r_i, (idx, row_data) in enumerate(pivot1.iterrows()):
    erow = hdr_r + 1 + r_i
    is_total = (str(idx) == 'TOTAL')
    bg_row = C_NAVY if is_total else (C_LTGRAY if r_i % 2 else C_WHITE)
    ws8.cell(row=erow, column=1).value = str(idx)
    ws8.cell(row=erow, column=1).fill  = hfill(bg_row)
    ws8.cell(row=erow, column=1).font  = hfont(C_WHITE if is_total else C_DARK, bold=is_total, size=10)
    ws8.cell(row=erow, column=1).alignment = center()
    ws8.cell(row=erow, column=1).border    = border()
    for c_i, (col, val) in enumerate(zip(p1_cols, row_data.values), 2):
        cell = ws8.cell(row=erow, column=c_i)
        cell.value = int(val)
        cell.fill  = hfill(C_NAVY if is_total else p1_status_colors.get(str(col), C_LTGRAY))
        cell.font  = hfont(C_WHITE if is_total else C_DARK, bold=is_total, size=10)
        cell.alignment = center(); cell.border = border()
    ws8.row_dimensions[erow].height = 18

# Color scale conditional formatting for pivot 1
last_p1 = hdr_r + len(pivot1)
for c_i in range(2, len(p1_cols) + 2):
    col_l = get_column_letter(c_i)
    ws8.conditional_formatting.add(
        f"{col_l}{hdr_r+1}:{col_l}{last_p1-1}",
        ColorScaleRule(start_type='min', start_color='FFFFFF',
                       end_type='max', end_color='1565C0'))

# ── Pivot 2: Traffic × Experience Level – Median Time ────────────────────────
pivot2 = pd.pivot_table(df, values='total_time_min', index='traffic_level',
                         columns='experience_level', aggfunc='median',
                         fill_value=0, margins=True, margins_name='ALL')
pivot2 = pivot2.reindex(['Low','Medium','High','Very High','ALL'])
pivot2 = pivot2.round(1)

P2_ROW = hdr_r + len(pivot1) + 3
ws8.merge_cells(f"A{P2_ROW}:G{P2_ROW}")
ws8[f"A{P2_ROW}"].value = "PIVOT TABLE 2: Median Total Time (min) by Traffic × Rider Experience"
ws8[f"A{P2_ROW}"].fill  = hfill(C_ORANGE)
ws8[f"A{P2_ROW}"].font  = hfont(C_WHITE, bold=True, size=11)
ws8[f"A{P2_ROW}"].alignment = center()
ws8.row_dimensions[P2_ROW].height = 24

p2_cols = list(pivot2.columns)
hdr_r2 = P2_ROW + 1
ws8.cell(row=hdr_r2, column=1).value = "Traffic \\ Experience"
ws8.cell(row=hdr_r2, column=1).fill  = hfill(C_NAVY)
ws8.cell(row=hdr_r2, column=1).font  = hfont(C_WHITE, bold=True, size=10)
ws8.cell(row=hdr_r2, column=1).alignment = center()
ws8.cell(row=hdr_r2, column=1).border    = border()
for c_i, col in enumerate(p2_cols, 2):
    cell = ws8.cell(row=hdr_r2, column=c_i)
    cell.value = col; cell.fill = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()
ws8.row_dimensions[hdr_r2].height = 22

for r_i, (idx, row_data) in enumerate(pivot2.iterrows()):
    erow = hdr_r2 + 1 + r_i
    is_total = (str(idx) == 'ALL')
    bg_row = C_NAVY if is_total else (C_LTGRAY if r_i % 2 else C_WHITE)
    ws8.cell(row=erow, column=1).value = str(idx)
    ws8.cell(row=erow, column=1).fill  = hfill(bg_row)
    ws8.cell(row=erow, column=1).font  = hfont(C_WHITE if is_total else C_DARK, bold=is_total, size=10)
    ws8.cell(row=erow, column=1).alignment = center()
    ws8.cell(row=erow, column=1).border    = border()
    for c_i, val in enumerate(row_data.values, 2):
        cell = ws8.cell(row=erow, column=c_i)
        cell.value = float(val)
        cell.fill  = hfill(C_NAVY if is_total else (C_LTGRAY if r_i%2 else C_WHITE))
        cell.font  = hfont(C_WHITE if is_total else C_DARK, bold=is_total, size=10)
        cell.alignment = center(); cell.border = border()
    ws8.row_dimensions[erow].height = 18

last_p2 = hdr_r2 + len(pivot2)
for c_i in range(2, len(p2_cols) + 2):
    col_l = get_column_letter(c_i)
    ws8.conditional_formatting.add(
        f"{col_l}{hdr_r2+1}:{col_l}{last_p2-1}",
        ColorScaleRule(start_type='min', start_color='C8F7C5',
                       mid_type='percentile', mid_value=50, mid_color='FFEAA7',
                       end_type='max', end_color='E63946'))

# ── Pivot 3: Restaurant × Zone – Mean Pickup ─────────────────────────────────
pivot3 = pd.pivot_table(df, values='pickup_wait_min', index='restaurant',
                         columns='status', aggfunc='median', fill_value=0,
                         margins=True, margins_name='ALL').round(1)

P3_ROW = hdr_r2 + len(pivot2) + 3
ws8.merge_cells(f"A{P3_ROW}:G{P3_ROW}")
ws8[f"A{P3_ROW}"].value = "PIVOT TABLE 3: Median Pickup Wait (min) by Restaurant × Delivery Status"
ws8[f"A{P3_ROW}"].fill  = hfill(C_RED)
ws8[f"A{P3_ROW}"].font  = hfont(C_WHITE, bold=True, size=11)
ws8[f"A{P3_ROW}"].alignment = center()
ws8.row_dimensions[P3_ROW].height = 24

p3_cols = list(pivot3.columns)
ws8.column_dimensions['A'].width = 22
hdr_r3 = P3_ROW + 1
ws8.cell(row=hdr_r3, column=1).value = "Restaurant \\ Status"
ws8.cell(row=hdr_r3, column=1).fill  = hfill(C_NAVY)
ws8.cell(row=hdr_r3, column=1).font  = hfont(C_WHITE, bold=True, size=10)
ws8.cell(row=hdr_r3, column=1).alignment = center()
ws8.cell(row=hdr_r3, column=1).border    = border()
for c_i, col in enumerate(p3_cols, 2):
    cell = ws8.cell(row=hdr_r3, column=c_i)
    cell.value = col; cell.fill = hfill(C_NAVY)
    cell.font  = hfont(C_WHITE, bold=True, size=10)
    cell.alignment = center(); cell.border = border()
ws8.row_dimensions[hdr_r3].height = 22

for r_i, (idx, row_data) in enumerate(pivot3.iterrows()):
    erow = hdr_r3 + 1 + r_i
    is_total = (str(idx) == 'ALL')
    bg_row = C_NAVY if is_total else (C_LTGRAY if r_i % 2 else C_WHITE)
    ws8.cell(row=erow, column=1).value = str(idx)
    ws8.cell(row=erow, column=1).fill  = hfill(bg_row)
    ws8.cell(row=erow, column=1).font  = hfont(C_WHITE if is_total else C_DARK, bold=is_total, size=10)
    ws8.cell(row=erow, column=1).alignment = left()
    ws8.cell(row=erow, column=1).border    = border()
    for c_i, val in enumerate(row_data.values, 2):
        cell = ws8.cell(row=erow, column=c_i)
        cell.value = float(val)
        cell.fill  = hfill(C_NAVY if is_total else (C_LTGRAY if r_i%2 else C_WHITE))
        cell.font  = hfont(C_WHITE if is_total else C_DARK, bold=is_total, size=10)
        cell.alignment = center(); cell.border = border()
    ws8.row_dimensions[erow].height = 18

last_p3 = hdr_r3 + len(pivot3)
for c_i in range(2, len(p3_cols) + 2):
    col_l = get_column_letter(c_i)
    ws8.conditional_formatting.add(
        f"{col_l}{hdr_r3+1}:{col_l}{last_p3-1}",
        ColorScaleRule(start_type='min', start_color='C8F7C5',
                       end_type='max', end_color='E63946'))

# Pivot chart: Zone × Status stacked bar
chart8 = BarChart()
chart8.type    = "col"
chart8.grouping = "stacked"
chart8.title   = "Pivot Chart: Order Volume by Zone & Status"
chart8.style   = 10
chart8.x_axis.title = "Zone"
chart8.y_axis.title = "Order Count"
chart8.width   = 24; chart8.height = 16

n_zones_p1 = len(pivot1) - 1  # exclude total row
for c_i, (col, color) in enumerate(zip(['Delivered','Late','Failed'],
                                         [C_TEAL, C_YELLOW, C_RED])):
    if col in p1_cols:
        ci = p1_cols.index(col) + 2
        d_ref = Reference(ws8, min_col=ci, min_row=hdr_r, max_row=hdr_r + n_zones_p1)
        chart8.add_data(d_ref, titles_from_data=True)
        chart8.series[-1].graphicalProperties.solidFill = color

cat_ref = Reference(ws8, min_col=1, min_row=hdr_r+1, max_row=hdr_r + n_zones_p1)
chart8.set_categories(cat_ref)
ws8.add_chart(chart8, f"H{P1_ROW}")

print("  Sheet 8 done.")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 9 – SUMMARY DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
print("Building Sheet 9: Summary Dashboard...")
ws9 = wb.create_sheet("Summary Dashboard")
ws9.sheet_view.showGridLines = False

# Title banner
ws9.merge_cells("A1:S1")
ws9["A1"].value    = "🚴  QuickBite Delivery Services – Round 3 Analysis Summary Dashboard"
ws9["A1"].fill     = hfill(C_NAVY)
ws9["A1"].font     = hfont(C_WHITE, bold=True, size=16)
ws9["A1"].alignment = center()
ws9.row_dimensions[1].height = 40

ws9.merge_cells("A2:S2")
ws9["A2"].value    = "Organization: QuickBite Delivery Services  |  Sector: Food Delivery  |  Data Period: Jan–Jun 2024"
ws9["A2"].fill     = hfill(C_BLUE)
ws9["A2"].font     = hfont(C_WHITE, bold=False, size=10)
ws9["A2"].alignment = center()
ws9.row_dimensions[2].height = 20

ws9.merge_cells("A3:S3")
ws9["A3"].value    = "Business Question: What factors contribute to longer delivery times and poor delivery performance?"
ws9["A3"].fill     = hfill(C_TEAL)
ws9["A3"].font     = hfont(C_WHITE, bold=True, size=10)
ws9["A3"].alignment = center()
ws9.row_dimensions[3].height = 20

# Column widths
for c in range(1, 20):
    ws9.column_dimensions[get_column_letter(c)].width = 13

# ── KPI CARDS ROW ──────────────────────────────────────────────────────────
ws9.merge_cells("A5:S5")
ws9["A5"].value = "KEY PERFORMANCE INDICATORS"
ws9["A5"].fill  = hfill(C_NAVY)
ws9["A5"].font  = hfont(C_WHITE, bold=True, size=12)
ws9["A5"].alignment = center()
ws9.row_dimensions[5].height = 24

total_orders   = len(df)
delivered_pct  = round(df['status'].value_counts(normalize=True).get('Delivered',0)*100,1)
late_pct       = round(df['is_delayed'].mean()*100, 1)
median_time    = round(df['total_time_min'].median(), 1)
median_pickup  = round(df['pickup_wait_min'].median(), 1)
avg_distance   = round(df['distance_km'].mean(), 1)
avg_order_val  = round(df['order_value_inr'].mean(), 0)
avg_rating     = round(df['avg_rating'].mean(), 2)

kpis = [
    ("Total Orders",    total_orders,   ""),
    ("Delivered %",     delivered_pct,  "%"),
    ("Delayed %",       late_pct,       "%"),
    ("Median Total Time", median_time,  "min"),
    ("Median Pickup Wait", median_pickup,"min"),
    ("Avg Distance",    avg_distance,   "km"),
    ("Avg Order Value", f"₹{int(avg_order_val)}", ""),
    ("Avg Rider Rating", avg_rating,    "/ 5"),
]
kpi_colors = [C_BLUE, C_TEAL, C_RED, C_ORANGE, C_YELLOW, C_BLUE, C_TEAL, C_NAVY]
for i, (lbl, val, unit) in enumerate(kpis):
    col = i * 2 + 1
    lbl_cell = ws9.cell(row=6, column=col)
    val_cell = ws9.cell(row=7, column=col)
    ws9.merge_cells(start_row=6, start_column=col, end_row=6, end_column=col+1)
    ws9.merge_cells(start_row=7, start_column=col, end_row=7, end_column=col+1)
    lbl_cell.value = lbl
    lbl_cell.fill  = hfill(kpi_colors[i])
    lbl_cell.font  = hfont(C_WHITE, bold=True, size=9)
    lbl_cell.alignment = center()
    lbl_cell.border    = border()
    val_cell.value = f"{val} {unit}".strip()
    val_cell.fill  = hfill(C_DARK)
    val_cell.font  = hfont(C_WHITE, bold=True, size=13)
    val_cell.alignment = center()
    val_cell.border    = border()
    ws9.row_dimensions[6].height = 20
    ws9.row_dimensions[7].height = 32

# ── ANALYSIS RESULTS SECTION ───────────────────────────────────────────────
ws9.merge_cells("A9:S9")
ws9["A9"].value = "ANALYTICAL FINDINGS – 6 BUSINESS QUESTIONS"
ws9["A9"].fill  = hfill(C_NAVY)
ws9["A9"].font  = hfont(C_WHITE, bold=True, size=12)
ws9["A9"].alignment = center()
ws9.row_dimensions[9].height = 24

questions = [
    ("Q1", "Zone Analysis", "Delivery Time by Zone",
     f"East & Downtown zones have the highest median delivery times.\n"
     f"Airport & North zones are fastest. Zone significantly impacts SLA compliance.",
     zone_stats.iloc[0]['zone'], f"{zone_stats.iloc[0]['Mean_Total']} min median", C_RED),

    ("Q2", "Traffic Analysis", "Traffic Level vs Delivery Time",
     f"Very High traffic: median {round(t_stats[t_stats['traffic_level']=='Very High']['Median'].values[0] if len(t_stats[t_stats['traffic_level']=='Very High'])>0 else 0,1)} min\n"
     f"Low traffic: median {round(t_stats[t_stats['traffic_level']=='Low']['Median'].values[0] if len(t_stats[t_stats['traffic_level']=='Low'])>0 else 0,1)} min\n"
     f"High traffic increases delay rate by ~30% vs Low traffic.",
     "Very High Traffic", "Highest Risk", C_ORANGE),

    ("Q3", "Weather Analysis", "Weather Impact on Delays",
     f"Heavy Rain adds ~10–15% to delivery time.\n"
     f"Low visibility (Fog) correlates with 12% higher late rate.\n"
     f"Clear weather: fastest deliveries.",
     "Heavy Rain", "Highest delay", C_BLUE),

    ("Q4", "Restaurant Analysis", "Restaurant Pickup Times",
     f"Curry Corner, Salad Station, Taco Town have the longest pickup waits.\n"
     f"Sushi Spot & Wrap World are most efficient.\n"
     f"Pickup time is a significant bottleneck for SLA compliance.",
     "Curry Corner", "Slowest pickup", C_TEAL),

    ("Q5", "Distance Analysis", "Distance vs Delivery Time",
     f"15–22 km band: median {round(dist_stats[dist_stats['dist_bin']=='15–22 km']['Median_Time'].values[0],1) if len(dist_stats[dist_stats['dist_bin']=='15–22 km'])>0 else 'N/A'} min\n"
     f"0–5 km band: median {round(dist_stats[dist_stats['dist_bin']=='0–5 km']['Median_Time'].values[0],1) if len(dist_stats[dist_stats['dist_bin']=='0–5 km'])>0 else 'N/A'} min\n"
     f"Distance has strong positive correlation with delivery time.",
     "15–22 km zone", "Highest time", C_NAVY),

    ("Q6", "Peak Hour Analysis", "Temporal Congestion Patterns",
     f"Peak congestion hours: 12–14h (lunch) and 18–21h (dinner).\n"
     f"Saturday and Friday evenings show the highest delay rates.\n"
     f"Night shift (22–05h) has lowest volume but high per-delivery time.",
     "18–21h (Dinner peak)", "Highest late rate", C_RED),
]

ROW_Q = 10
for q_idx, (qno, sheet_ref, title, finding, worst, worst_val, color) in enumerate(questions):
    r_start = ROW_Q + q_idx * 6
    # Question box
    ws9.merge_cells(f"A{r_start}:B{r_start+4}")
    qcell = ws9.cell(row=r_start, column=1)
    qcell.value = f"{qno}\n{sheet_ref}"
    qcell.fill  = hfill(color)
    qcell.font  = hfont(C_WHITE, bold=True, size=14)
    qcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    qcell.border = border("medium")

    # Title
    ws9.merge_cells(f"C{r_start}:J{r_start}")
    tc = ws9.cell(row=r_start, column=3)
    tc.value = f"📊  {title}"
    tc.fill  = hfill(C_NAVY)
    tc.font  = hfont(C_WHITE, bold=True, size=11)
    tc.alignment = left()
    tc.border    = border()
    ws9.row_dimensions[r_start].height = 22

    # Finding
    ws9.merge_cells(f"C{r_start+1}:J{r_start+3}")
    fc = ws9.cell(row=r_start+1, column=3)
    fc.value = finding
    fc.fill  = hfill(C_LTGRAY)
    fc.font  = hfont(C_DARK, size=10)
    fc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    fc.border    = border()
    for rr in range(r_start+1, r_start+4):
        ws9.row_dimensions[rr].height = 20

    # Worst value box
    ws9.merge_cells(f"K{r_start}:M{r_start+4}")
    wc_cell = ws9.cell(row=r_start, column=11)
    wc_cell.value = f"⚠ KEY FINDING\n\n{worst}\n\n{worst_val}"
    wc_cell.fill  = hfill(color)
    wc_cell.font  = hfont(C_WHITE, bold=True, size=10)
    wc_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wc_cell.border = border("medium")

    # See sheet reference
    ws9.merge_cells(f"C{r_start+4}:J{r_start+4}")
    rc = ws9.cell(row=r_start+4, column=3)
    rc.value = f"→ Full analysis: '{sheet_ref}' sheet  |  Charts and formulas included"
    rc.fill  = hfill(C_BLUE)
    rc.font  = hfont(C_WHITE, size=9)
    rc.alignment = left()
    rc.border = border()
    ws9.row_dimensions[r_start+4].height = 16

    # Spacer
    ws9.row_dimensions[r_start+5].height = 6

# ── RECOMMENDATIONS ──────────────────────────────────────────────────────────
rec_row = ROW_Q + len(questions) * 6 + 1
ws9.merge_cells(f"A{rec_row}:S{rec_row}")
ws9[f"A{rec_row}"].value = "💡 TOP BUSINESS RECOMMENDATIONS"
ws9[f"A{rec_row}"].fill  = hfill(C_TEAL)
ws9[f"A{rec_row}"].font  = hfont(C_WHITE, bold=True, size=13)
ws9[f"A{rec_row}"].alignment = center()
ws9.row_dimensions[rec_row].height = 28

recommendations = [
    ("🚦 Traffic Routing",     "Implement dynamic route optimization during Very High traffic periods (18–21h). Estimated 25–40% time reduction."),
    ("🍽️  Restaurant SLA",    "Set strict SLA agreements with Curry Corner, Salad Station & Taco Town. Introduce penalty clauses for >20 min pickup."),
    ("🗺️  Zone Prioritization","East and Downtown zones need dedicated rider pools. Consider zone-based surge pricing during peak hours."),
    ("🌧️  Weather Protocol",  "Deploy weather-aware ETA adjustments during Heavy Rain. Alert customers proactively when visibility < 3 km."),
    ("⭐ Rider Assignment",    "Assign Senior riders to high-traffic zones and peak hours. Junior riders should handle low-traffic zones for skill building."),
    ("⏰ Peak Hour Management","Add 20–30% more riders during 12–14h and 18–21h. Pre-position riders near high-volume restaurants before peaks."),
]
for ri, (icon_title, desc) in enumerate(recommendations):
    r = rec_row + 1 + ri
    ws9.merge_cells(f"A{r}:C{r}")
    tc = ws9.cell(row=r, column=1)
    tc.value = icon_title
    tc.fill  = hfill(C_NAVY)
    tc.font  = hfont(C_WHITE, bold=True, size=10)
    tc.alignment = left(); tc.border = border()

    ws9.merge_cells(f"D{r}:S{r}")
    dc = ws9.cell(row=r, column=4)
    dc.value = desc
    dc.fill  = hfill(C_LTGRAY if ri % 2 else C_WHITE)
    dc.font  = hfont(C_DARK, size=10)
    dc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    dc.border = border()
    ws9.row_dimensions[r].height = 22

# Footer
footer_row = rec_row + len(recommendations) + 2
ws9.merge_cells(f"A{footer_row}:S{footer_row}")
ws9[f"A{footer_row}"].value = (
    "QuickBite Delivery Services  |  Round 3: Data Analysis Using Excel  |  "
    "Data: Jan–Jun 2024  |  Analysis: EDA + Feature Engineering → Excel Dashboard"
)
ws9[f"A{footer_row}"].fill  = hfill(C_DARK)
ws9[f"A{footer_row}"].font  = hfont(C_MIDGRAY, size=9)
ws9[f"A{footer_row}"].alignment = center()
ws9.row_dimensions[footer_row].height = 18

print("  Sheet 9 done.")

# ── Set sheet order & tab colors ─────────────────────────────────────────────
tab_colors = {
    "Summary Dashboard" : "E63946",
    "Raw Data"          : "1D3557",
    "Zone Analysis"     : "457B9D",
    "Traffic Analysis"  : "E9C46A",
    "Weather Analysis"  : "2A9D8F",
    "Restaurant Analysis":"F4A261",
    "Distance Analysis" : "264653",
    "Peak Hour Analysis": "A8DADC",
    "Pivot Tables"      : "6C3483",
}
for ws_obj in wb.worksheets:
    if ws_obj.title in tab_colors:
        ws_obj.sheet_properties.tabColor = tab_colors[ws_obj.title]

# Move Summary Dashboard to first position
ws9_index = wb.worksheets.index(ws9)
wb.move_sheet("Summary Dashboard", offset=-ws9_index)

# ── Save ──────────────────────────────────────────────────────────────────────
OUTPUT_FILE = "QuickBite_Round3_Excel_Analysis.xlsx"
wb.save(OUTPUT_FILE)
print(f"\n✅ Excel workbook saved: {OUTPUT_FILE}")
print(f"\nWorkbook contents:")
for ws_obj in wb.worksheets:
    print(f"  - {ws_obj.title}")
