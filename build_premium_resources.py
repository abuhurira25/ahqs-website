#!/usr/bin/env python3
"""Build 6 premium AHQS Excel templates with charts, dashboards, and verified formulas."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle, Protection
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule, ColorScaleRule, IconSetRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart import BarChart, PieChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from PIL import Image as PILImage
import os, datetime

# ── Brand constants ──
NAVY = "071F3B"
NAVY2 = "0C3560"
TEAL = "087F83"
TEAL2 = "0A5C5E"
GOLD = "D8A72B"
GOLD2 = "F0CB67"
LIGHT_BG = "F5F8FA"
LIGHT_TEAL = "EAF7F7"
LIGHT_GOLD = "FFF8E8"
BORDER_GRAY = "DCE6EB"
WHITE = "FFFFFF"
RED = "FF6B6B"
ORANGE = "FFA94D"
YELLOW = "FFE066"
GREEN = "C3F6C3"
INK = "173047"
MUTED = "667789"

OUTDIR = "downloads/free-resources"
os.makedirs(OUTDIR, exist_ok=True)

# Resize logo
logo_src = "ahqs-official-logo-hd.png"
logo_small = os.path.join(OUTDIR, "ahqs-logo-small.png")
if os.path.exists(logo_src):
    img = PILImage.open(logo_src)
    img.thumbnail((200, 70), PILImage.Resampling.LANCZOS)
    img.save(logo_small, "PNG")

# ── Reusable style helpers ──
thin = Side(style="thin", color=BORDER_GRAY)
medium = Side(style="medium", color=NAVY)
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
border_box = Border(left=medium, right=medium, top=medium, bottom=medium)

def hdr_cell(ws, row, col, value, fill_color=TEAL, font_color=WHITE, size=11):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Calibri", bold=True, size=size, color=font_color)
    c.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border_all
    return c

def data_cell(ws, row, col, value=None, bold=False, size=10, fill=None, align="left"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Calibri", bold=bold, size=size, color=INK)
    c.alignment = Alignment(vertical="center", horizontal=align, wrap_text=True)
    c.border = border_all
    if fill:
        c.fill = PatternFill(start_color=fill, end_color=fill, fill_type="solid")
    return c

def kpi_card(ws, row, col, label, formula, fill_color=NAVY, value_color=GOLD2):
    """Create a KPI card spanning 2 columns and 3 rows."""
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+1)
    ws.merge_cells(start_row=row+1, start_column=col, end_row=row+2, end_column=col+1)
    lbl = ws.cell(row=row, column=col, value=label)
    lbl.font = Font(name="Calibri", bold=True, size=9, color="D4E2EB")
    lbl.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    lbl.alignment = Alignment(horizontal="center", vertical="center")
    val = ws.cell(row=row+1, column=col, value=formula)
    val.font = Font(name="Calibri", bold=True, size=22, color=value_color)
    val.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    val.alignment = Alignment(horizontal="center", vertical="center")
    # Set row heights
    ws.row_dimensions[row].height = 22
    ws.row_dimensions[row+1].height = 20
    ws.row_dimensions[row+2].height = 20

def add_cover(wb, title, description, tool_id, features_list):
    """Add a premium branded cover sheet."""
    ws = wb.create_sheet("Cover", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 3

    # Logo
    if os.path.exists(logo_small):
        img = XLImage(logo_small)
        img.anchor = "B2"
        ws.add_image(img)

    # Title block
    ws.merge_cells("B9:B9")
    c = ws.cell(row=9, column=2, value=title)
    c.font = Font(name="Calibri", bold=True, size=22, color=NAVY)

    c = ws.cell(row=10, column=2, value="AHQS — Accelerate Healthcare Quality Solutions")
    c.font = Font(name="Calibri", size=11, color=TEAL, italic=True)

    # Divider
    for col in range(2, 4):
        ws.cell(row=12, column=col).border = Border(bottom=Side(style="medium", color=GOLD))

    # Description
    c = ws.cell(row=13, column=2, value=description)
    c.font = Font(name="Calibri", size=11, color=MUTED)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[13].height = 50

    # Key features
    c = ws.cell(row=15, column=2, value="Key Features")
    c.font = Font(name="Calibri", bold=True, size=13, color=NAVY)
    for i, feat in enumerate(features_list):
        c = ws.cell(row=16+i, column=2, value=f"  •  {feat}")
        c.font = Font(name="Calibri", size=10, color=INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    # Instructions
    start = 16 + len(features_list) + 1
    c = ws.cell(row=start, column=2, value="How to Use")
    c.font = Font(name="Calibri", bold=True, size=13, color=NAVY)
    instructions = [
        "1.  Start with the Data Entry sheet — fill in yellow-shaded input cells only",
        "2.  Blue-shaded cells contain formulas — do not overwrite them",
        "3.  Use dropdown menus for standardized entries",
        "4.  Check the Dashboard sheet for auto-calculated metrics and charts",
        "5.  Save a copy for each audit cycle or reporting period",
        "6.  This tool is free to use for internal quality improvement",
    ]
    for i, instr in enumerate(instructions):
        c = ws.cell(row=start+1+i, column=2, value=instr)
        c.font = Font(name="Calibri", size=10, color=INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    # Disclaimer
    disc_row = start + 8
    c = ws.cell(row=disc_row, column=2, value="Disclaimer & License")
    c.font = Font(name="Calibri", bold=True, size=13, color=NAVY)
    disclaimer = ("This is an AHQS-created quality improvement tool. It is not an official CAP, ISO 15189, "
                  "JCI or AABB standards document, and does not reproduce official standards text. "
                  "Always purchase official standards directly from the accrediting bodies. "
                  "This template is provided free of charge for internal quality improvement use. "
                  "© 2026 Accelerate Healthcare Quality Solutions (AHQS). Not for resale.")
    c = ws.cell(row=disc_row+1, column=2, value=disclaimer)
    c.font = Font(name="Calibri", size=9, color=MUTED, italic=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[disc_row+1].height = 65

    # Footer
    c = ws.cell(row=disc_row+3, column=2, value=f"Tool ID: {tool_id}  |  Version: 2.0  |  Created: October 2026  |  AHQS Healthcare Quality Solutions")
    c.font = Font(name="Calibri", size=9, color=MUTED)

    print(f"  Cover sheet: {title}")

def add_settings_sheet(wb, sheet_name, lists_dict):
    """Add a hidden settings sheet with named lists for data validation."""
    ws = wb.create_sheet(sheet_name)
    ws.sheet_state = "hidden"
    ws.sheet_view.showGridLines = False
    col = 1
    for list_name, values in lists_dict.items():
        hdr_cell(ws, 1, col, list_name, fill_color=NAVY)
        for i, v in enumerate(values, 2):
            c = ws.cell(row=i, column=col, value=v)
            c.font = Font(name="Calibri", size=10, color=INK)
        ws.column_dimensions[get_column_letter(col)].width = 25
        col += 1
    print(f"  Settings sheet: {sheet_name} ({len(lists_dict)} lists)")

def add_chart_title(ws, row, col, title, span=6):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+span-1)
    c = ws.cell(row=row, column=col, value=title)
    c.font = Font(name="Calibri", bold=True, size=13, color=NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 24

# ══════════════════════════════════════════════════════════
# 1. PREMIUM AUDIT TOOL TEMPLATE
# ══════════════════════════════════════════════════════════
print("\n[1/6] Building Premium Audit Tool Template...")
wb = openpyxl.Workbook()
add_cover(wb, "Quality Audit Tool Template",
    "A comprehensive internal audit tool with weighted compliance scoring, findings tracking, risk-rated gaps, and a visual dashboard with charts.",
    "AHQS-AUD-002",
    ["Weighted compliance scoring (0-100%)",
     "Multi-domain audit checklist with auto-scoring",
     "Findings register with risk levels and CAPA linkage",
     "Evidence tracking with status",
     "Executive dashboard with KPIs and charts",
     "Cross-reference to applicable standards"])

# Settings
add_settings_sheet(wb, "Settings", {
    "Audit_Areas": ["Document Control", "Competency & Training", "Quality Control", "CAPA", "Equipment Management",
                    "Safety", "Personnel", "Facilities & Environment", "Information Management", "Service Quality"],
    "Risk_Levels": ["Critical", "High", "Medium", "Low", "Observation"],
    "Status": ["Open", "In Progress", "Closed", "Not Applicable"],
    "Finding_Type": ["Nonconformity", "Observation", "Opportunity for Improvement"],
    "Compliance": ["Compliant", "Non-Compliant", "Partially Compliant", "Not Assessed"],
})

# Audit Checklist sheet
ws = wb.create_sheet("Audit Checklist")
ws.sheet_view.showGridLines = False
headers = ["#", "Audit Domain", "Audit Question", "Applicable Reference", "Compliance", "Weight (%)", "Score", "Finding", "Risk Level", "Corrective Action", "Owner", "Due Date", "Status"]
for c, h in enumerate(headers, 1):
    hdr_cell(ws, 1, c, h)
ws.row_dimensions[1].height = 36

widths = [4, 18, 38, 16, 14, 9, 7, 30, 11, 28, 14, 12, 12]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

questions = [
    ("Document Control", "Are SOPs current, reviewed, and approved by authorized personnel?", "Verify against your licensed standard", "Partially Compliant", 10),
    ("Document Control", "Is there a master list of all controlled documents?", "Verify against your licensed standard", "Compliant", 8),
    ("Document Control", "Are obsolete documents removed from points of use?", "Verify against your licensed standard", "Compliant", 7),
    ("Competency & Training", "Are competency assessments completed for all staff at required intervals?", "Verify against your licensed standard", "Non-Compliant", 10),
    ("Competency & Training", "Are training records maintained and up to date?", "Verify against your licensed standard", "Partially Compliant", 8),
    ("Quality Control", "Are QC records reviewed daily by authorized personnel?", "Verify against your licensed standard", "Compliant", 10),
    ("Quality Control", "Are QC failures investigated and documented?", "Verify against your licensed standard", "Partially Compliant", 8),
    ("CAPA", "Are corrective actions tracked to closure within target timelines?", "Verify against your licensed standard", "Non-Compliant", 10),
    ("CAPA", "Are root cause analyses performed for significant findings?", "Verify against your licensed standard", "Compliant", 7),
    ("Equipment Management", "Is equipment calibrated per schedule and documented?", "Verify against your licensed standard", "Compliant", 8),
    ("Equipment Management", "Are preventive maintenance records maintained?", "Verify against your licensed standard", "Partially Compliant", 7),
    ("Safety", "Is PPE available, appropriate, and used by all staff?", "Verify against your licensed standard", "Compliant", 8),
    ("Safety", "Are safety incidents reported, investigated, and trended?", "Verify against your licensed standard", "Compliant", 7),
    ("Personnel", "Are job descriptions current and on file for all staff?", "Verify against your licensed standard", "Not Assessed", 5),
    ("Personnel", "Are continuing education records maintained?", "Verify against your licensed standard", "Partially Compliant", 5),
    ("Facilities & Environment", "Is temperature and humidity monitored and recorded?", "Verify against your licensed standard", "Compliant", 7),
    ("Facilities & Environment", "Is the laboratory clean, organized, and free of clutter?", "Verify against your licensed standard", "Compliant", 5),
    ("Information Management", "Is the LIS backed up regularly with documented recovery testing?", "Verify against your licensed standard", "Partially Compliant", 8),
    ("Information Management", "Are access permissions reviewed periodically?", "Verify against your licensed standard", "Compliant", 5),
    ("Service Quality", "Are turnaround times monitored and reported?", "Verify against your licensed standard", "Compliant", 5),
]

# Data validations
dv_compliance = DataValidation(type="list", formula1='"Compliant,Non-Compliant,Partially Compliant,Not Assessed"', allow_blank=True)
dv_risk = DataValidation(type="list", formula1='"Critical,High,Medium,Low,Observation"', allow_blank=True)
dv_status = DataValidation(type="list", formula1='"Open,In Progress,Closed,Not Applicable"', allow_blank=True)
ws.add_data_validation(dv_compliance)
ws.add_data_validation(dv_risk)
ws.add_data_validation(dv_status)

for i, (domain, question, ref, compliance, weight) in enumerate(questions, 1):
    row = i + 1
    data_cell(ws, row, 1, i, align="center")
    data_cell(ws, row, 2, domain)
    data_cell(ws, row, 3, question)
    data_cell(ws, row, 4, ref)
    data_cell(ws, row, 5, compliance, align="center")
    data_cell(ws, row, 6, weight, align="center")
    # Score: Compliant=1, Partially=0.5, Non-Compliant=0, Not Assessed=0
    data_cell(ws, row, 7, f'=IF(E{row}="Compliant",1,IF(E{row}="Partially Compliant",0.5,0))', align="center", bold=True, fill=LIGHT_TEAL)
    data_cell(ws, row, 8, "")
    data_cell(ws, row, 9, "")
    data_cell(ws, row, 10, "")
    data_cell(ws, row, 11, "")
    data_cell(ws, row, 12, "")
    data_cell(ws, row, 13, "")
    dv_compliance.add(ws.cell(row=row, column=5))
    dv_risk.add(ws.cell(row=row, column=9))
    dv_status.add(ws.cell(row=row, column=13))

# Add empty rows
for row in range(len(questions)+2, len(questions)+32):
    data_cell(ws, row, 1, row-1, align="center")
    data_cell(ws, row, 7, f'=IF(E{row}="Compliant",1,IF(E{row}="Partially Compliant",0.5,0))', align="center", bold=True, fill=LIGHT_TEAL)
    dv_compliance.add(ws.cell(row=row, column=5))
    dv_risk.add(ws.cell(row=row, column=9))
    dv_status.add(ws.cell(row=row, column=13))

last_row = len(questions) + 31

# Conditional formatting on compliance column
ws.conditional_formatting.add(f"E2:E{last_row}",
    CellIsRule(operator="equal", formula=['"Compliant"'], fill=PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")))
ws.conditional_formatting.add(f"E2:E{last_row}",
    CellIsRule(operator="equal", formula=['"Non-Compliant"'], fill=PatternFill(start_color=RED, end_color=RED, fill_type="solid"), font=Font(bold=True, color=WHITE)))
ws.conditional_formatting.add(f"E2:E{last_row}",
    CellIsRule(operator="equal", formula=['"Partially Compliant"'], fill=PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:M{last_row}"

# Findings Register sheet
ws_f = wb.create_sheet("Findings Register")
ws_f.sheet_view.showGridLines = False
f_headers = ["#", "Finding ID", "Domain", "Finding Description", "Risk Level", "Root Cause", "Corrective Action", "Owner", "Due Date", "Status", "Date Closed", "Days Open", "Effectiveness"]
for c, h in enumerate(f_headers, 1):
    hdr_cell(ws_f, 1, c, h)
ws_f.row_dimensions[1].height = 36
f_widths = [4, 12, 16, 30, 11, 25, 28, 14, 12, 12, 12, 10, 22]
for i, w in enumerate(f_widths, 1):
    ws_f.column_dimensions[get_column_letter(i)].width = w

findings = [
    ("F-001", "Competency & Training", "Competency assessments not completed for 2 staff members", "High", "Competency records stored in different locations", "Consolidate competency records into central file", "HR Manager", "2026-11-15", "In Progress", "", "", ""),
    ("F-002", "CAPA", "Corrective actions not closed within 30-day target", "Critical", "CAPA tracking system not consistently used", "Implement centralized CAPA tracking with auto-alerts", "Quality Manager", "2026-11-30", "Open", "", "", ""),
    ("F-003", "Quality Control", "QC failures not consistently investigated", "Medium", "Staff unsure of investigation procedure", "Develop and train QC investigation SOP", "Lab Supervisor", "2026-12-15", "Open", "", "", ""),
    ("F-004", "Equipment Management", "Preventive maintenance records incomplete", "Medium", "Maintenance log not consistently filled", "Implement digital maintenance log with reminders", "Lab Manager", "2026-12-01", "Open", "", "", ""),
    ("F-005", "Information Management", "LIS backup recovery not tested", "High", "No documented recovery test procedure", "Schedule and document quarterly recovery tests", "IT Manager", "2026-11-20", "Open", "", "", ""),
]

for i, (fid, domain, desc, risk, rca, ca, owner, due, status, closed, days, eff) in enumerate(findings, 1):
    row = i + 1
    data_cell(ws_f, row, 1, i, align="center")
    data_cell(ws_f, row, 2, fid, align="center", bold=True)
    data_cell(ws_f, row, 3, domain)
    data_cell(ws_f, row, 4, desc)
    data_cell(ws_f, row, 5, risk, align="center")
    data_cell(ws_f, row, 6, rca)
    data_cell(ws_f, row, 7, ca)
    data_cell(ws_f, row, 8, owner)
    data_cell(ws_f, row, 9, due, align="center")
    data_cell(ws_f, row, 10, status, align="center")
    data_cell(ws_f, row, 11, closed, align="center")
    data_cell(ws_f, row, 12, f'=IF(AND(J{row}<>"Closed",J{row}<>"Not Applicable",I{row}<>"",TODAY()>DATEVALUE(I{row})),TODAY()-DATEVALUE(I{row}),IF(K{row}<>"",DATEVALUE(K{row})-DATEVALUE(I{row}),0))', align="center", bold=True, fill=LIGHT_TEAL)
    data_cell(ws_f, row, 13, eff)

for row in range(len(findings)+2, len(findings)+22):
    data_cell(ws_f, row, 1, row-1, align="center")
    data_cell(ws_f, row, 12, f'=IF(AND(J{row}<>"Closed",J{row}<>"Not Applicable",I{row}<>"",TODAY()>DATEVALUE(I{row})),TODAY()-DATEVALUE(I{row}),IF(K{row}<>"",DATEVALUE(K{row})-DATEVALUE(I{row}),0))', align="center", bold=True, fill=LIGHT_TEAL)

f_last = len(findings) + 21
dv_risk2 = DataValidation(type="list", formula1='"Critical,High,Medium,Low,Observation"', allow_blank=True)
dv_status2 = DataValidation(type="list", formula1='"Open,In Progress,Closed,Not Applicable"', allow_blank=True)
ws_f.add_data_validation(dv_risk2)
ws_f.add_data_validation(dv_status2)
for row in range(2, f_last+1):
    dv_risk2.add(ws_f.cell(row=row, column=5))
    dv_status2.add(ws_f.cell(row=row, column=10))

ws_f.conditional_formatting.add(f"L2:L{f_last}",
    CellIsRule(operator="greaterThan", formula=["30"], fill=PatternFill(start_color=RED, end_color=RED, fill_type="solid"), font=Font(bold=True, color=WHITE)))
ws_f.freeze_panes = "A2"
ws_f.auto_filter.ref = f"A1:M{f_last}"

# Dashboard sheet
ws_d = wb.create_sheet("Dashboard")
ws_d.sheet_view.showGridLines = False
ws_d.column_dimensions["A"].width = 3
for col in "BCDEFGHIJKL":
    ws_d.column_dimensions[col].width = 14

ws_d.merge_cells("B2:L2")
c = ws_d.cell(row=2, column=2, value="Quality Audit Dashboard")
c.font = Font(name="Calibri", bold=True, size=18, color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_d.row_dimensions[2].height = 32

# KPI Cards
kpi_card(ws_d, 4, 2, "Overall Compliance", '=IFERROR(SUMPRODUCT(\'Audit Checklist\'!G2:G21,\'Audit Checklist\'!F2:F21)/SUM(\'Audit Checklist\'!F2:F21)*100,0)', NAVY, GOLD2)
ws_d.cell(row=4, column=4).value = None
ws_d.merge_cells(start_row=4, start_column=4, end_row=4, end_column=5)
ws_d.merge_cells(start_row=5, start_column=4, end_row=6, end_column=5)
lbl = ws_d.cell(row=4, column=4, value="Total Questions")
lbl.font = Font(bold=True, size=9, color="D4E2EB")
lbl.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
lbl.alignment = Alignment(horizontal="center", vertical="center")
val = ws_d.cell(row=5, column=4, value='=COUNTA(\'Audit Checklist\'!B2:B51)')
val.font = Font(bold=True, size=22, color=GOLD2)
val.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
val.alignment = Alignment(horizontal="center", vertical="center")
ws_d.row_dimensions[4].height = 22
ws_d.row_dimensions[5].height = 20
ws_d.row_dimensions[6].height = 20

kpi_card(ws_d, 4, 6, "Open Findings", '=COUNTIF(\'Findings Register\'!J2:J26,"Open")+COUNTIF(\'Findings Register\'!J2:J26,"In Progress")', NAVY, GOLD2)
kpi_card(ws_d, 4, 8, "Critical/High Risks", '=COUNTIF(\'Findings Register\'!E2:E26,"Critical")+COUNTIF(\'Findings Register\'!E2:E26,"High")', NAVY, GOLD2)
kpi_card(ws_d, 4, 10, "Closed Findings", '=COUNTIF(\'Findings Register\'!J2:J26,"Closed")', NAVY, GOLD2)

# Domain compliance table
add_chart_title(ws_d, 9, 2, "Compliance Score by Domain", 6)
ws_d.cell(row=10, column=2, value="Domain")
ws_d.cell(row=10, column=3, value="Compliant")
ws_d.cell(row=10, column=4, value="Partial")
ws_d.cell(row=10, column=5, value="Non-Compliant")
ws_d.cell(row=10, column=6, value="Score %")
for c in range(2, 7):
    hdr_cell(ws_d, 10, c, ws_d.cell(row=10, column=c).value)

domains = ["Document Control", "Competency & Training", "Quality Control", "CAPA", "Equipment Management",
           "Safety", "Personnel", "Facilities & Environment", "Information Management", "Service Quality"]
for i, dom in enumerate(domains, 1):
    row = 10 + i
    data_cell(ws_d, row, 2, dom)
    data_cell(ws_d, row, 3, f'=COUNTIFS(\'Audit Checklist\'!B2:B51,"{dom}",\'Audit Checklist\'!E2:E51,"Compliant")', align="center")
    data_cell(ws_d, row, 4, f'=COUNTIFS(\'Audit Checklist\'!B2:B51,"{dom}",\'Audit Checklist\'!E2:E51,"Partially Compliant")', align="center")
    data_cell(ws_d, row, 5, f'=COUNTIFS(\'Audit Checklist\'!B2:B51,"{dom}",\'Audit Checklist\'!E2:E51,"Non-Compliant")', align="center")
    data_cell(ws_d, row, 6, f'=IFERROR(C{row}/(C{row}+D{row}*0.5+E{row})*100,0)', align="center", bold=True, fill=LIGHT_TEAL)

# Data bars on score column
ws_d.conditional_formatting.add("F11:F20", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color=TEAL))

# Bar chart: compliance by domain
chart1 = BarChart()
chart1.type = "bar"
chart1.style = 10
chart1.title = "Compliance Score by Domain"
chart1.y_axis.title = "Domain"
chart1.x_axis.title = "Score %"
data = Reference(ws_d, min_col=6, min_row=10, max_row=20, max_col=6)
cats = Reference(ws_d, min_col=2, min_row=11, max_row=20)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.height = 10
chart1.width = 18
ws_d.add_chart(chart1, "H9")

# Findings by risk level table
add_chart_title(ws_d, 22, 2, "Findings by Risk Level", 6)
ws_d.cell(row=23, column=2, value="Risk Level")
ws_d.cell(row=23, column=3, value="Count")
for c in range(2, 4):
    hdr_cell(ws_d, 23, c, ws_d.cell(row=23, column=c).value)

risks = ["Critical", "High", "Medium", "Low", "Observation"]
for i, risk in enumerate(risks, 1):
    row = 23 + i
    data_cell(ws_d, row, 2, risk)
    data_cell(ws_d, row, 3, f'=COUNTIF(\'Findings Register\'!E2:E26,"{risk}")', align="center", bold=True)

# Pie chart: findings by risk
chart2 = PieChart()
chart2.title = "Findings by Risk Level"
data2 = Reference(ws_d, min_col=3, min_row=23, max_row=28)
cats2 = Reference(ws_d, min_col=2, min_row=24, max_row=28)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.height = 8
chart2.width = 12
chart2.dataLabels = DataLabelList(showPercent=True)
ws_d.add_chart(chart2, "E22")

wb.save(os.path.join(OUTDIR, "ahqs-audit-tool-template.xlsx"))
print("  Saved: ahqs-audit-tool-template.xlsx")

# ══════════════════════════════════════════════════════════
# 2. PREMIUM RISK REGISTER TEMPLATE
# ══════════════════════════════════════════════════════════
print("\n[2/6] Building Premium Risk Register Template...")
wb = openpyxl.Workbook()
add_cover(wb, "Risk Register Template",
    "A comprehensive risk management tool with inherent and residual risk scoring, 5x5 matrix heatmap, mitigation tracking, and a visual dashboard.",
    "AHQS-RSK-002",
    ["Inherent and residual risk scoring (likelihood x severity)",
     "5x5 risk matrix with color-coded heatmap",
     "Risk category breakdown with counts",
     "Mitigation action tracking with owners and deadlines",
     "Top risks dashboard with charts",
     "Auto-calculated risk reduction (delta)"])

add_settings_sheet(wb, "Settings", {
    "Categories": ["Clinical", "Operational", "Financial", "Compliance", "Safety", "Equipment", "IT/Data", "Personnel", "Facility"],
    "Risk_Levels": ["Low", "Medium", "High", "Critical"],
    "Status": ["Open", "Mitigated", "Closed", "Monitoring", "Escalated"],
    "Owners": ["Lab Manager", "Quality Manager", "Medical Director", "Safety Officer", "IT Manager", "HR Manager"],
})

# Risk Register sheet
ws = wb.create_sheet("Risk Register")
ws.sheet_view.showGridLines = False
headers = ["#", "Risk ID", "Date", "Category", "Risk Description", "Likelihood (1-5)", "Severity (1-5)", "Inherent Score", "Inherent Rating", "Mitigation Action", "Residual Likelihood", "Residual Severity", "Residual Score", "Residual Rating", "Owner", "Target Date", "Status", "Review Date"]
for c, h in enumerate(headers, 1):
    hdr_cell(ws, 1, c, h)
ws.row_dimensions[1].height = 40
widths = [4, 10, 12, 14, 35, 10, 10, 10, 10, 30, 10, 10, 10, 10, 14, 12, 12, 12]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

risks_data = [
    ("R001", "2026-01-15", "Clinical", "Delayed critical value reporting to requesting physician", 3, 5, "Implement automated alert system for critical values", 2, 3),
    ("R002", "2026-01-20", "Compliance", "Expired reagents in use on automated analyzer", 2, 4, "Implement barcode scanning for reagent verification", 1, 2),
    ("R003", "2026-02-01", "Safety", "Inadequate PPE compliance in specimen processing area", 3, 3, "Mandatory PPE training and spot audits", 2, 2),
    ("R004", "2026-02-10", "Equipment", "Analyzer calibration drift affecting QC results", 2, 5, "Implement daily calibration verification and trending", 1, 3),
    ("R005", "2026-02-15", "IT/Data", "LIS downtime during peak testing hours", 3, 4, "Implement redundant server with automatic failover", 1, 3),
    ("R006", "2026-03-01", "Operational", "Insufficient staffing during peak hours", 3, 3, "Cross-train staff and implement flexible scheduling", 2, 2),
    ("R007", "2026-03-10", "Financial", "Budget constraints limiting equipment replacement", 2, 4, "Develop 5-year capital equipment replacement plan", 2, 3),
    ("R008", "2026-03-15", "Personnel", "Key person dependency for specialized testing", 2, 5, "Cross-train minimum 2 staff per specialized test", 1, 3),
    ("R009", "2026-04-01", "Facility", "Temperature excursions in storage areas", 2, 4, "Install continuous temperature monitoring with alarms", 1, 2),
    ("R010", "2026-04-10", "Clinical", "Misidentification of patient samples", 2, 5, "Implement two-patient identifier verification at all steps", 1, 2),
]

dv_cat = DataValidation(type="list", formula1='"Clinical,Operational,Financial,Compliance,Safety,Equipment,IT/Data,Personnel,Facility"', allow_blank=True)
dv_status = DataValidation(type="list", formula1='"Open,Mitigated,Closed,Monitoring,Escalated"', allow_blank=True)
dv_owner = DataValidation(type="list", formula1='"Lab Manager,Quality Manager,Medical Director,Safety Officer,IT Manager,HR Manager"', allow_blank=True)
ws.add_data_validation(dv_cat)
ws.add_data_validation(dv_status)
ws.add_data_validation(dv_owner)

for i, (rid, date, cat, desc, lik, sev, mit, rlik, rsev) in enumerate(risks_data, 1):
    row = i + 1
    data_cell(ws, row, 1, i, align="center")
    data_cell(ws, row, 2, rid, align="center", bold=True)
    data_cell(ws, row, 3, date, align="center")
    data_cell(ws, row, 4, cat)
    data_cell(ws, row, 5, desc)
    data_cell(ws, row, 6, lik, align="center")
    data_cell(ws, row, 7, sev, align="center")
    data_cell(ws, row, 8, f"=F{row}*G{row}", align="center", bold=True, fill=LIGHT_TEAL)
    data_cell(ws, row, 9, f'=IF(H{row}>=15,"Critical",IF(H{row}>=9,"High",IF(H{row}>=4,"Medium","Low")))', align="center", bold=True)
    data_cell(ws, row, 10, mit)
    data_cell(ws, row, 11, rlik, align="center")
    data_cell(ws, row, 12, rsev, align="center")
    data_cell(ws, row, 13, f"=K{row}*L{row}", align="center", bold=True, fill=LIGHT_TEAL)
    data_cell(ws, row, 14, f'=IF(M{row}>=15,"Critical",IF(M{row}>=9,"High",IF(M{row}>=4,"Medium","Low")))', align="center", bold=True)
    data_cell(ws, row, 15, "")
    data_cell(ws, row, 16, "")
    data_cell(ws, row, 17, "")
    data_cell(ws, row, 18, "")
    dv_cat.add(ws.cell(row=row, column=4))
    dv_status.add(ws.cell(row=row, column=17))
    dv_owner.add(ws.cell(row=row, column=15))

# Empty rows with formulas
for row in range(len(risks_data)+2, len(risks_data)+32):
    data_cell(ws, row, 1, row-1, align="center")
    data_cell(ws, row, 8, f"=F{row}*G{row}", align="center", bold=True, fill=LIGHT_TEAL)
    data_cell(ws, row, 9, f'=IF(H{row}>=15,"Critical",IF(H{row}>=9,"High",IF(H{row}>=4,"Medium","Low")))', align="center", bold=True)
    data_cell(ws, row, 13, f"=K{row}*L{row}", align="center", bold=True, fill=LIGHT_TEAL)
    data_cell(ws, row, 14, f'=IF(M{row}>=15,"Critical",IF(M{row}>=9,"High",IF(M{row}>=4,"Medium","Low")))', align="center", bold=True)
    dv_cat.add(ws.cell(row=row, column=4))
    dv_status.add(ws.cell(row=row, column=17))
    dv_owner.add(ws.cell(row=row, column=15))

last_risk = len(risks_data) + 31

# Conditional formatting on inherent and residual scores
for col_letter in ["H", "M"]:
    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{last_risk}",
        CellIsRule(operator="greaterThanOrEqual", formula=["15"], fill=PatternFill(start_color=RED, end_color=RED, fill_type="solid"), font=Font(bold=True, color=WHITE)))
    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{last_risk}",
        CellIsRule(operator="between", formula=["9", "14"], fill=PatternFill(start_color=ORANGE, end_color=ORANGE, fill_type="solid")))
    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{last_risk}",
        CellIsRule(operator="between", formula=["4", "8"], fill=PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")))
    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{last_risk}",
        CellIsRule(operator="lessThan", formula=["4"], fill=PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:R{last_risk}"

# Risk Matrix sheet
ws_m = wb.create_sheet("Risk Matrix")
ws_m.sheet_view.showGridLines = False
ws_m.column_dimensions["A"].width = 3
ws_m.column_dimensions["B"].width = 22
for col in "CDEFG":
    ws_m.column_dimensions[col].width = 16

ws_m.merge_cells("B2:G2")
c = ws_m.cell(row=2, column=2, value="Risk Assessment Matrix (5x5)")
c.font = Font(name="Calibri", bold=True, size=16, color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_m.row_dimensions[2].height = 28

# Matrix headers
ws_m.merge_cells("B4:C4")
c = ws_m.cell(row=4, column=2, value="Likelihood / Severity")
c.font = Font(bold=True, size=10, color=WHITE)
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center")
ws_m.row_dimensions[4].height = 30

sev_labels = ["1 - Rare", "2 - Unlikely", "3 - Possible", "4 - Likely", "5 - Almost Certain"]
lik_labels = ["5 - Almost Certain", "4 - Likely", "3 - Possible", "2 - Unlikely", "1 - Rare"]

for c_idx, label in enumerate(sev_labels):
    hdr_cell(ws_m, 5, 3+c_idx, label, fill_color=TEAL)
ws_m.row_dimensions[5].height = 36

for r_idx, lik_label in enumerate(lik_labels):
    row = 6 + r_idx
    hdr_cell(ws_m, row, 2, lik_label, fill_color=TEAL)
    ws_m.row_dimensions[row].height = 30
    for c_idx in range(5):
        score = (5 - r_idx) * (c_idx + 1)
        cell = data_cell(ws_m, row, 3+c_idx, score, align="center", bold=True, size=14)
        if score >= 15:
            cell.fill = PatternFill(start_color=RED, end_color=RED, fill_type="solid")
            cell.font = Font(size=14, bold=True, color=WHITE)
        elif score >= 9:
            cell.fill = PatternFill(start_color=ORANGE, end_color=ORANGE, fill_type="solid")
        elif score >= 4:
            cell.fill = PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")
        else:
            cell.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")

# Legend
ws_m.cell(row=12, column=2, value="Legend:").font = Font(bold=True, size=10, color=NAVY)
legend = [("Critical (15-25)", RED), ("High (9-14)", ORANGE), ("Medium (4-8)", YELLOW), ("Low (1-3)", GREEN)]
for i, (label, color) in enumerate(legend):
    c = ws_m.cell(row=12, column=3+i, value=label)
    c.font = Font(size=10, bold=True, color=WHITE if color in [RED, ORANGE] else INK)
    c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    c.alignment = Alignment(horizontal="center")

# Dashboard
ws_d = wb.create_sheet("Dashboard")
ws_d.sheet_view.showGridLines = False
ws_d.column_dimensions["A"].width = 3
for col in "BCDEFGHIJKL":
    ws_d.column_dimensions[col].width = 14

ws_d.merge_cells("B2:L2")
c = ws_d.cell(row=2, column=2, value="Risk Management Dashboard")
c.font = Font(name="Calibri", bold=True, size=18, color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_d.row_dimensions[2].height = 32

# KPI cards
kpi_card(ws_d, 4, 2, "Total Risks", f'=COUNTA(\'Risk Register\'!B2:B{last_risk})', NAVY, GOLD2)
kpi_card(ws_d, 4, 4, "Critical Risks", f'=COUNTIF(\'Risk Register\'!I2:I{last_risk},"Critical")', NAVY, GOLD2)
kpi_card(ws_d, 4, 6, "High Risks", f'=COUNTIF(\'Risk Register\'!I2:I{last_risk},"High")', NAVY, GOLD2)
kpi_card(ws_d, 4, 8, "Open Risks", f'=COUNTIF(\'Risk Register\'!Q2:Q{last_risk},"Open")+COUNTIF(\'Risk Register\'!Q2:Q{last_risk},"Escalated")', NAVY, GOLD2)
kpi_card(ws_d, 4, 10, "Risk Reduction %", f'=IFERROR((SUM(\'Risk Register\'!H2:H{last_risk})-SUM(\'Risk Register\'!M2:M{last_risk}))/SUM(\'Risk Register\'!H2:H{last_risk})*100,0)', NAVY, GOLD2)

# Risk by category table
add_chart_title(ws_d, 9, 2, "Risk Count by Category", 6)
ws_d.cell(row=10, column=2, value="Category")
ws_d.cell(row=10, column=3, value="Count")
ws_d.cell(row=10, column=4, value="Critical")
ws_d.cell(row=10, column=5, value="High")
for c in range(2, 6):
    hdr_cell(ws_d, 10, c, ws_d.cell(row=10, column=c).value)

cats = ["Clinical", "Operational", "Financial", "Compliance", "Safety", "Equipment", "IT/Data", "Personnel", "Facility"]
for i, cat in enumerate(cats, 1):
    row = 10 + i
    data_cell(ws_d, row, 2, cat)
    data_cell(ws_d, row, 3, f'=COUNTIF(\'Risk Register\'!D2:D{last_risk},"{cat}")', align="center", bold=True)
    data_cell(ws_d, row, 4, f'=COUNTIFS(\'Risk Register\'!D2:D{last_risk},"{cat}",\'Risk Register\'!I2:I{last_risk},"Critical")', align="center")
    data_cell(ws_d, row, 5, f'=COUNTIFS(\'Risk Register\'!D2:D{last_risk},"{cat}",\'Risk Register\'!I2:I{last_risk},"High")', align="center")

# Bar chart
chart1 = BarChart()
chart1.type = "col"
chart1.style = 10
chart1.title = "Risks by Category"
chart1.y_axis.title = "Count"
data = Reference(ws_d, min_col=3, min_row=10, max_row=19, max_col=5)
cats_ref = Reference(ws_d, min_col=2, min_row=11, max_row=19)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats_ref)
chart1.height = 9
chart1.width = 18
ws_d.add_chart(chart1, "G9")

# Inherent vs Residual
add_chart_title(ws_d, 22, 2, "Inherent vs Residual Risk Score", 6)
for i, risk_id in enumerate(["R001", "R002", "R003", "R004", "R005"]):
    row = 23 + i
    data_cell(ws_d, row, 2, risk_id, align="center", bold=True)
    data_cell(ws_d, row, 3, f'=INDEX(\'Risk Register\'!H2:H{last_risk},MATCH("{risk_id}",\'Risk Register\'!B2:B{last_risk},0))', align="center")
    data_cell(ws_d, row, 4, f'=INDEX(\'Risk Register\'!M2:M{last_risk},MATCH("{risk_id}",\'Risk Register\'!B2:B{last_risk},0))', align="center")

chart2 = BarChart()
chart2.type = "col"
chart2.style = 12
chart2.title = "Inherent vs Residual Risk"
chart2.y_axis.title = "Score"
data2 = Reference(ws_d, min_col=3, min_row=23, max_row=27, max_col=4)
cats2 = Reference(ws_d, min_col=2, min_row=24, max_row=27)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.height = 8
chart2.width = 14
ws_d.add_chart(chart2, "G22")

wb.save(os.path.join(OUTDIR, "ahqs-risk-register-template.xlsx"))
print("  Saved: ahqs-risk-register-template.xlsx")

# ══════════════════════════════════════════════════════════
# 3. PREMIUM CAPA TRACKING TEMPLATE
# ══════════════════════════════════════════════════════════
print("\n[3/6] Building Premium CAPA Tracking Template...")
wb = openpyxl.Workbook()
add_cover(wb, "CAPA Tracking Template",
    "A comprehensive corrective and preventive action tracker with root cause analysis, aging buckets, effectiveness verification, and a closure-rate dashboard.",
    "AHQS-CAPA-002",
    ["Root cause analysis with 5-Why methodology",
     "Aging buckets (0-30, 31-60, 61-90, 90+ days)",
     "Auto-calculated overdue flags",
     "Effectiveness verification tracking",
     "Closure rate dashboard with charts",
     "CAPA source breakdown (audit, incident, complaint)"])

add_settings_sheet(wb, "Settings", {
    "Sources": ["Audit Finding", "Incident", "Complaint", "External Assessment", "Internal Review", "QC Failure", "Safety Event"],
    "Status": ["Open", "In Progress", "Closed", "Overdue", "Cancelled"],
    "RCA_Method": ["5 Whys", "Fishbone", "FMEA", "Root Cause Analysis", "Process Mapping"],
    "Effectiveness": ["Effective", "Partially Effective", "Not Effective", "Monitoring", "Not Yet Assessed"],
})

# CAPA Register
ws = wb.create_sheet("CAPA Register")
ws.sheet_view.showGridLines = False
headers = ["#", "CAPA ID", "Date Identified", "Source", "Issue Description", "Root Cause", "RCA Method", "Corrective Action", "Preventive Action", "Owner", "Due Date", "Status", "Date Closed", "Days Open", "Aging Bucket", "Effectiveness", "Review Date"]
for c, h in enumerate(headers, 1):
    hdr_cell(ws, 1, c, h)
ws.row_dimensions[1].height = 40
widths = [4, 10, 12, 16, 32, 28, 12, 28, 28, 14, 12, 12, 12, 10, 12, 16, 12]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

capas = [
    ("CAPA-001", "2026-01-10", "Audit Finding", "QC records not reviewed by supervisor within 24 hours", "Supervisor workload exceeded capacity during peak hours", "5 Whys", "Reassign QC review to alternate supervisor during peak hours", "Implement automated QC alert system", "Lab Manager", "2026-02-15", "Closed", "2026-02-10", "Effective"),
    ("CAPA-002", "2026-01-20", "Incident", "Expired reagent used in patient testing", "Reagent inventory not checked before use", "5 Whys", "Discard expired reagents, retest affected samples", "Implement barcode scanning for reagent verification", "Quality Officer", "2026-02-28", "In Progress", "", ""),
    ("CAPA-003", "2026-02-05", "External Assessment", "Competency assessment not documented for 2 staff", "Competency records stored in different locations", "Fishbone", "Consolidate all competency records into central file", "Implement quarterly competency review schedule", "HR Manager", "2026-03-15", "Open", "", ""),
    ("CAPA-004", "2026-02-12", "Complaint", "Delayed TAT for routine chemistry panels", "Analyzer bottleneck during morning run", "Process Mapping", "Optimize sample loading sequence", "Evaluate need for additional analyzer capacity", "Lab Supervisor", "2026-03-01", "In Progress", "", ""),
    ("CAPA-005", "2026-03-01", "QC Failure", "Westgard rule violation not acted upon", "Staff not trained on Westgard rules interpretation", "5 Whys", "Retrain all staff on Westgard multi-rule QC", "Implement mandatory annual QC competency", "Quality Manager", "2026-03-20", "Open", "", ""),
    ("CAPA-006", "2026-03-10", "Safety Event", "Chemical spill not reported within required timeframe", "Staff unsure of reporting procedure", "5 Whys", "Post spill response procedure at all workstations", "Conduct safety refresher training for all staff", "Safety Officer", "2026-04-01", "Open", "", ""),
    ("CAPA-007", "2026-03-20", "Internal Review", "Document review cycle not maintained", "No automated reminder system for document reviews", "FMEA", "Create document review calendar with reminders", "Implement QMS document control module", "Quality Manager", "2026-04-30", "In Progress", "", ""),
    ("CAPA-008", "2026-04-01", "Audit Finding", "Equipment maintenance logs incomplete", "Maintenance log not consistently filled", "5 Whys", "Backfill missing maintenance records", "Implement digital maintenance log with mandatory fields", "Lab Manager", "2026-04-15", "Closed", "2026-04-10", "Partially Effective"),
]

dv_source = DataValidation(type="list", formula1='"Audit Finding,Incident,Complaint,External Assessment,Internal Review,QC Failure,Safety Event"', allow_blank=True)
dv_status = DataValidation(type="list", formula1='"Open,In Progress,Closed,Overdue,Cancelled"', allow_blank=True)
dv_rca = DataValidation(type="list", formula1='"5 Whys,Fishbone,FMEA,Root Cause Analysis,Process Mapping"', allow_blank=True)
dv_eff = DataValidation(type="list", formula1='"Effective,Partially Effective,Not Effective,Monitoring,Not Yet Assessed"', allow_blank=True)
ws.add_data_validation(dv_source)
ws.add_data_validation(dv_status)
ws.add_data_validation(dv_rca)
ws.add_data_validation(dv_eff)

for i, (cid, date, source, issue, rca, method, ca, pa, owner, due, status, closed, eff) in enumerate(capas, 1):
    row = i + 1
    data_cell(ws, row, 1, i, align="center")
    data_cell(ws, row, 2, cid, align="center", bold=True)
    data_cell(ws, row, 3, date, align="center")
    data_cell(ws, row, 4, source)
    data_cell(ws, row, 5, issue)
    data_cell(ws, row, 6, rca)
    data_cell(ws, row, 7, method, align="center")
    data_cell(ws, row, 8, ca)
    data_cell(ws, row, 9, pa)
    data_cell(ws, row, 10, owner)
    data_cell(ws, row, 11, due, align="center")
    data_cell(ws, row, 12, status, align="center")
    data_cell(ws, row, 13, closed, align="center")
    data_cell(ws, row, 14, f'=IF(L{row}="Closed",IF(N{row}<>"",DATEVALUE(N{row})-DATEVALUE(C{row}),0),IF(C{row}<>"",TODAY()-DATEVALUE(C{row}),0))', align="center", bold=True, fill=LIGHT_TEAL)
    data_cell(ws, row, 15, f'=IF(N{row}=0,"",IF(N{row}<=30,"0-30 Days",IF(N{row}<=60,"31-60 Days",IF(N{row}<=90,"61-90 Days","90+ Days"))))', align="center")
    data_cell(ws, row, 16, eff, align="center")
    data_cell(ws, row, 17, "")
    dv_source.add(ws.cell(row=row, column=4))
    dv_status.add(ws.cell(row=row, column=12))
    dv_rca.add(ws.cell(row=row, column=7))
    dv_eff.add(ws.cell(row=row, column=16))

for row in range(len(capas)+2, len(capas)+32):
    data_cell(ws, row, 1, row-1, align="center")
    data_cell(ws, row, 14, f'=IF(L{row}="Closed",IF(N{row}<>"",DATEVALUE(N{row})-DATEVALUE(C{row}),0),IF(C{row}<>"",TODAY()-DATEVALUE(C{row}),0))', align="center", bold=True, fill=LIGHT_TEAL)
    data_cell(ws, row, 15, f'=IF(N{row}=0,"",IF(N{row}<=30,"0-30 Days",IF(N{row}<=60,"31-60 Days",IF(N{row}<=90,"61-90 Days","90+ Days"))))', align="center")
    dv_source.add(ws.cell(row=row, column=4))
    dv_status.add(ws.cell(row=row, column=12))
    dv_rca.add(ws.cell(row=row, column=7))
    dv_eff.add(ws.cell(row=row, column=16))

last_capa = len(capas) + 31

# Conditional formatting on days open
ws.conditional_formatting.add(f"N2:N{last_capa}",
    CellIsRule(operator="greaterThan", formula=["30"], fill=PatternFill(start_color=RED, end_color=RED, fill_type="solid"), font=Font(bold=True, color=WHITE)))
ws.conditional_formatting.add(f"N2:N{last_capa}",
    CellIsRule(operator="between", formula=["15", "30"], fill=PatternFill(start_color=ORANGE, end_color=ORANGE, fill_type="solid")))
ws.conditional_formatting.add(f"L2:L{last_capa}",
    CellIsRule(operator="equal", formula=['"Overdue"'], fill=PatternFill(start_color=RED, end_color=RED, fill_type="solid"), font=Font(bold=True, color=WHITE)))
ws.conditional_formatting.add(f"L2:L{last_capa}",
    CellIsRule(operator="equal", formula=['"Open"'], fill=PatternFill(start_color=ORANGE, end_color=ORANGE, fill_type="solid")))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:Q{last_capa}"

# Dashboard
ws_d = wb.create_sheet("Dashboard")
ws_d.sheet_view.showGridLines = False
ws_d.column_dimensions["A"].width = 3
for col in "BCDEFGHIJKL":
    ws_d.column_dimensions[col].width = 14

ws_d.merge_cells("B2:L2")
c = ws_d.cell(row=2, column=2, value="CAPA Management Dashboard")
c.font = Font(name="Calibri", bold=True, size=18, color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_d.row_dimensions[2].height = 32

kpi_card(ws_d, 4, 2, "Total CAPAs", f'=COUNTA(\'CAPA Register\'!B2:B{last_capa})', NAVY, GOLD2)
kpi_card(ws_d, 4, 4, "Open", f'=COUNTIF(\'CAPA Register\'!L2:L{last_capa},"Open")', NAVY, GOLD2)
kpi_card(ws_d, 4, 6, "In Progress", f'=COUNTIF(\'CAPA Register\'!L2:L{last_capa},"In Progress")', NAVY, GOLD2)
kpi_card(ws_d, 4, 8, "Closed", f'=COUNTIF(\'CAPA Register\'!L2:L{last_capa},"Closed")', NAVY, GOLD2)
kpi_card(ws_d, 4, 10, "Closure Rate %", f'=IFERROR(COUNTIF(\'CAPA Register\'!L2:L{last_capa},"Closed")/(COUNTA(\'CAPA Register\'!B2:B{last_capa}))*100,0)', NAVY, GOLD2)

# Status breakdown
add_chart_title(ws_d, 9, 2, "CAPA Status Breakdown", 6)
ws_d.cell(row=10, column=2, value="Status")
ws_d.cell(row=10, column=3, value="Count")
for c in range(2, 4):
    hdr_cell(ws_d, 10, c, ws_d.cell(row=10, column=c).value)

statuses = ["Open", "In Progress", "Closed", "Overdue", "Cancelled"]
for i, status in enumerate(statuses, 1):
    row = 10 + i
    data_cell(ws_d, row, 2, status)
    data_cell(ws_d, row, 3, f'=COUNTIF(\'CAPA Register\'!L2:L{last_capa},"{status}")', align="center", bold=True)

# Pie chart
chart1 = PieChart()
chart1.title = "CAPA Status Breakdown"
data = Reference(ws_d, min_col=3, min_row=10, max_row=15)
cats = Reference(ws_d, min_col=2, min_row=11, max_row=15)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.height = 8
chart1.width = 12
chart1.dataLabels = DataLabelList(showPercent=True)
ws_d.add_chart(chart1, "E9")

# Source breakdown
add_chart_title(ws_d, 22, 2, "CAPA Source Breakdown", 6)
ws_d.cell(row=23, column=2, value="Source")
ws_d.cell(row=23, column=3, value="Count")
for c in range(2, 4):
    hdr_cell(ws_d, 23, c, ws_d.cell(row=23, column=c).value)

sources = ["Audit Finding", "Incident", "Complaint", "External Assessment", "Internal Review", "QC Failure", "Safety Event"]
for i, source in enumerate(sources, 1):
    row = 23 + i
    data_cell(ws_d, row, 2, source)
    data_cell(ws_d, row, 3, f'=COUNTIF(\'CAPA Register\'!D2:D{last_capa},"{source}")', align="center", bold=True)

chart2 = BarChart()
chart2.type = "col"
chart2.style = 10
chart2.title = "CAPA Sources"
chart2.y_axis.title = "Count"
data2 = Reference(ws_d, min_col=3, min_row=23, max_row=30)
cats2 = Reference(ws_d, min_col=2, min_row=24, max_row=30)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.height = 8
chart2.width = 14
ws_d.add_chart(chart2, "E22")

# Aging breakdown
add_chart_title(ws_d, 35, 2, "CAPA Aging Analysis", 6)
ws_d.cell(row=36, column=2, value="Age")
ws_d.cell(row=36, column=3, value="Count")
for c in range(2, 4):
    hdr_cell(ws_d, 36, c, ws_d.cell(row=36, column=c).value)

buckets = ["0-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]
for i, bucket in enumerate(buckets, 1):
    row = 36 + i
    data_cell(ws_d, row, 2, bucket)
    data_cell(ws_d, row, 3, f'=COUNTIF(\'CAPA Register\'!O2:O{last_capa},"{bucket}")', align="center", bold=True)

chart3 = BarChart()
chart3.type = "col"
chart3.style = 12
chart3.title = "CAPA Aging"
chart3.y_axis.title = "Count"
data3 = Reference(ws_d, min_col=3, min_row=36, max_row=40)
cats3 = Reference(ws_d, min_col=2, min_row=37, max_row=40)
chart3.add_data(data3, titles_from_data=True)
chart3.set_categories(cats3)
chart3.height = 8
chart3.width = 14
ws_d.add_chart(chart3, "E35")

wb.save(os.path.join(OUTDIR, "ahqs-capa-tracking-template.xlsx"))
print("  Saved: ahqs-capa-tracking-template.xlsx")

# ══════════════════════════════════════════════════════════
# 4. PREMIUM QUALITY INDICATOR DASHBOARD
# ══════════════════════════════════════════════════════════
print("\n[4/6] Building Premium Quality Indicator Dashboard...")
wb = openpyxl.Workbook()
add_cover(wb, "Quality Indicator Dashboard",
    "A comprehensive KPI tracking tool for laboratory quality indicators with monthly data entry, target comparison, YTD performance, trend charts, and a RAG status dashboard.",
    "AHQS-KPI-002",
    ["8 key laboratory quality indicators with 12-month tracking",
     "Auto-calculated compliance status (Met/Not Met)",
     "YTD performance percentage",
     "Trend line charts for each indicator",
     "RAG (Red/Amber/Green) status dashboard",
     "Target direction support (higher is better / lower is better)"])

# KPI Data sheet
ws = wb.create_sheet("KPI Data")
ws.sheet_view.showGridLines = False
headers = ["#", "Month", "Indicator", "Target", "Actual", "Unit", "Direction", "Status", "Variance", "YTD Avg", "YTD % Met"]
for c, h in enumerate(headers, 1):
    hdr_cell(ws, 1, c, h)
ws.row_dimensions[1].height = 36
widths = [4, 12, 32, 10, 10, 8, 10, 12, 10, 10, 10]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

months = ["Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26"]
indicators = [
    ("Turnaround Time - Routine", 4.0, "hours", "≤"),
    ("Turnaround Time - STAT", 2.0, "hours", "≤"),
    ("Sample Rejection Rate", 2.0, "%", "≤"),
    ("Critical Value Reporting Time", 30.0, "min", "≤"),
    ("CAPA Closure Rate (30 days)", 90.0, "%", "≥"),
    ("QC Pass Rate", 95.0, "%", "≥"),
    ("Proficiency Testing Score", 80.0, "%", "≥"),
    ("Repeat Test Rate", 5.0, "%", "≤"),
]

dv_status = DataValidation(type="list", formula1='"Met,Not Met,N/A"', allow_blank=True)
ws.add_data_validation(dv_status)

row = 2
indicator_start_rows = []
for ind_idx, (indicator, target, unit, direction) in enumerate(indicators):
    start_row = row
    indicator_start_rows.append(start_row)
    for month_idx, month in enumerate(months):
        # Sample data: some months met, some not
        if direction == "≥":
            actual = target + (0.5 if month_idx % 3 == 0 else 1.5)
        else:
            actual = target - (0.5 if month_idx % 3 == 0 else -0.3)
        actual = round(actual, 1)

        data_cell(ws, row, 1, row-1, align="center")
        data_cell(ws, row, 2, month, align="center")
        data_cell(ws, row, 3, indicator)
        data_cell(ws, row, 4, target, align="center")
        data_cell(ws, row, 5, actual, align="center", bold=True)
        data_cell(ws, row, 6, unit, align="center")
        data_cell(ws, row, 7, direction, align="center")
        data_cell(ws, row, 8, f'=IF(E{row}="","",IF(G{row}="≥",IF(E{row}>=D{row},"Met","Not Met"),IF(E{row}<=D{row},"Met","Not Met")))', align="center", bold=True, fill=LIGHT_TEAL)
        data_cell(ws, row, 9, f'=IF(E{row}="","",IF(G{row}="≥",E{row}-D{row},D{row}-E{row}))', align="center")
        data_cell(ws, row, 10, f'=IFERROR(AVERAGEIF($C${start_row}:$C${start_row+11},C{row},$E${start_row}:$E${start_row+11}),"")', align="center")
        data_cell(ws, row, 11, f'=IFERROR(COUNTIFS($C${start_row}:$C${start_row+11},C{row},$H${start_row}:$H${start_row+11},"Met")/COUNTA($H${start_row}:$H${start_row+11})*100,0)', align="center", bold=True)
        dv_status.add(ws.cell(row=row, column=8))
        row += 1

last_kpi = row - 1
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:K{last_kpi}"

# Conditional formatting on status
ws.conditional_formatting.add(f"H2:H{last_kpi}",
    CellIsRule(operator="equal", formula=['"Met"'], fill=PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")))
ws.conditional_formatting.add(f"H2:H{last_kpi}",
    CellIsRule(operator="equal", formula=['"Not Met"'], fill=PatternFill(start_color=RED, end_color=RED, fill_type="solid"), font=Font(bold=True, color=WHITE)))

# Dashboard
ws_d = wb.create_sheet("Dashboard")
ws_d.sheet_view.showGridLines = False
ws_d.column_dimensions["A"].width = 3
for col in "BCDEFGHIJKL":
    ws_d.column_dimensions[col].width = 14

ws_d.merge_cells("B2:L2")
c = ws_d.cell(row=2, column=2, value="Quality Indicator Dashboard")
c.font = Font(name="Calibri", bold=True, size=18, color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_d.row_dimensions[2].height = 32

# KPI summary table
add_chart_title(ws_d, 4, 2, "Indicator Summary", 10)
sum_headers = ["", "Indicator", "Target", "Latest Actual", "Status", "YTD % Met"]
for c, h in enumerate(sum_headers, 1):
    if h:
        hdr_cell(ws_d, 5, c, h)

for i, (indicator, target, unit, direction) in enumerate(indicators):
    r = 6 + i
    start = indicator_start_rows[i]
    end = start + 11
    data_cell(ws_d, r, 2, f"{indicator} ({unit})")
    data_cell(ws_d, r, 3, target, align="center")
    data_cell(ws_d, r, 4, f'=IFERROR(LOOKUP(2,1/(\'KPI Data\'!E{start}:E{end}<>""),\'KPI Data\'!E{start}:E{end}),"-")', align="center", bold=True)
    data_cell(ws_d, r, 5, f'=IF(D{r}="-","",IF("{direction}"="≥",IF(D{r}>={target},"Met","Not Met"),IF(D{r}<={target},"Met","Not Met")))', align="center", bold=True)
    data_cell(ws_d, r, 6, f'=IFERROR(\'KPI Data\'!K{start}/100,"-")', align="center", bold=True)

    ws_d.conditional_formatting.add(f"E{r}",
        CellIsRule(operator="equal", formula=['"Met"'], fill=PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")))
    ws_d.conditional_formatting.add(f"E{r}",
        CellIsRule(operator="equal", formula=['"Not Met"'], fill=PatternFill(start_color=RED, end_color=RED, fill_type="solid"), font=Font(bold=True, color=WHITE)))

# Data bars on YTD
ws_d.conditional_formatting.add(f"F6:F{6+len(indicators)-1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=TEAL))

# Trend chart for first indicator
chart1 = LineChart()
chart1.title = "Turnaround Time - Routine (hours)"
chart1.style = 10
chart1.y_axis.title = "Hours"
chart1.x_axis.title = "Month"
start = indicator_start_rows[0]
end = start + 11
data = Reference(ws, min_col=5, min_row=start, max_row=end)
cats = Reference(ws, min_col=2, min_row=start, max_row=end)
chart1.add_data(data, titles_from_data=False)
chart1.set_categories(cats)
chart1.height = 8
chart1.width = 16
ws_d.add_chart(chart1, "H4")

# Trend chart for CAPA closure
chart2 = LineChart()
chart2.title = "CAPA Closure Rate (%)"
chart2.style = 12
chart2.y_axis.title = "%"
chart2.x_axis.title = "Month"
start = indicator_start_rows[4]
end = start + 11
data2 = Reference(ws, min_col=5, min_row=start, max_row=end)
cats2 = Reference(ws, min_col=2, min_row=start, max_row=end)
chart2.add_data(data2, titles_from_data=False)
chart2.set_categories(cats2)
chart2.height = 8
chart2.width = 16
ws_d.add_chart(chart2, "H22")

# Overall KPI cards
add_chart_title(ws_d, 22, 2, "Overall Performance", 6)
kpi_card(ws_d, 24, 2, "Indicators Tracked", f'=COUNTA(Dashboard!B6:B13)', NAVY, GOLD2)
kpi_card(ws_d, 24, 4, "Met This Month", f'=COUNTIF(Dashboard!E6:E13,"Met")', NAVY, GOLD2)
kpi_card(ws_d, 24, 6, "Not Met", f'=COUNTIF(Dashboard!E6:E13,"Not Met")', NAVY, GOLD2)
kpi_card(ws_d, 24, 8, "Overall Compliance %", f'=IFERROR(COUNTIF(Dashboard!E6:E13,"Met")/(COUNTIF(Dashboard!E6:E13,"Met")+COUNTIF(Dashboard!E6:E13,"Not Met"))*100,0)', NAVY, GOLD2)

wb.save(os.path.join(OUTDIR, "ahqs-quality-indicator-dashboard.xlsx"))
print("  Saved: ahqs-quality-indicator-dashboard.xlsx")

# ══════════════════════════════════════════════════════════
# 5. PREMIUM DOCUMENT CONTROL REGISTER
# ══════════════════════════════════════════════════════════
print("\n[5/6] Building Premium Document Control Register...")
wb = openpyxl.Workbook()
add_cover(wb, "Document Control Register",
    "A comprehensive document management tool with version control, review tracking, approval workflow, distribution lists, and a status dashboard with charts.",
    "AHQS-DOC-002",
    ["Document version control with review cycle tracking",
     "Auto-calculated overdue review flags",
     "Document type and status breakdown",
     "Distribution list tracking",
     "Review calendar with upcoming deadlines",
     "Dashboard with charts and KPIs"])

add_settings_sheet(wb, "Settings", {
    "Doc_Types": ["SOP", "Policy", "Form", "Manual", "Guideline", "Work Instruction", "Record"],
    "Doc_Status": ["Draft", "Under Review", "Approved", "Active", "Superseded", "Archived"],
    "Departments": ["Laboratory", "Quality", "Safety", "Administration", "IT", "HR", "All"],
})

# Document Register
ws = wb.create_sheet("Document Register")
ws.sheet_view.showGridLines = False
headers = ["#", "Doc ID", "Document Title", "Type", "Version", "Status", "Author", "Approved By", "Effective Date", "Next Review", "Department", "Distribution", "Location/Link", "Review Status", "Days to Review"]
for c, h in enumerate(headers, 1):
    hdr_cell(ws, 1, c, h)
ws.row_dimensions[1].height = 40
widths = [4, 10, 30, 10, 8, 12, 14, 14, 12, 12, 12, 18, 22, 12, 12]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

docs = [
    ("DOC-001", "Sample Receipt and Accessioning SOP", "SOP", "3.1", "Active", "Lab Tech", "Lab Manager", "2026-01-01", "2027-01-01", "Laboratory", "All lab staff", "Shared Drive / SOP-001"),
    ("DOC-002", "Critical Value Notification Policy", "Policy", "2.0", "Active", "Quality Officer", "Medical Director", "2026-01-15", "2027-01-15", "Quality", "All staff", "Shared Drive / POL-002"),
    ("DOC-003", "Equipment Calibration Record Form", "Form", "1.5", "Active", "Lab Manager", "Quality Manager", "2026-02-01", "2026-11-01", "Laboratory", "Lab staff", "Shared Drive / FRM-003"),
    ("DOC-004", "Laboratory Safety Manual", "Manual", "4.0", "Active", "Safety Officer", "Lab Director", "2026-01-01", "2027-01-01", "Safety", "All staff", "Shared Drive / MAN-004"),
    ("DOC-005", "CAPA Investigation Form", "Form", "2.1", "Under Review", "Quality Officer", "", "2025-08-01", "2026-08-01", "Quality", "Quality team", "Shared Drive / FRM-005"),
    ("DOC-006", "Quality Management System Manual", "Manual", "5.0", "Active", "Quality Manager", "Lab Director", "2026-03-01", "2027-03-01", "Quality", "All staff", "Shared Drive / MAN-006"),
    ("DOC-007", "Specimen Rejection Criteria", "SOP", "2.3", "Active", "Lab Supervisor", "Lab Manager", "2026-02-15", "2027-02-15", "Laboratory", "Lab staff", "Shared Drive / SOP-007"),
    ("DOC-008", "Competency Assessment Procedure", "SOP", "1.2", "Active", "HR Manager", "Lab Director", "2026-01-20", "2026-10-20", "HR", "All staff", "Shared Drive / SOP-008"),
    ("DOC-009", "Proficiency Testing Procedure", "SOP", "3.0", "Active", "Quality Manager", "Lab Director", "2026-01-10", "2027-01-10", "Laboratory", "Lab staff", "Shared Drive / SOP-009"),
    ("DOC-010", "Temperature Monitoring Log Form", "Form", "1.0", "Draft", "Lab Tech", "", "", "2026-12-01", "Laboratory", "Lab staff", "Shared Drive / FRM-010"),
]

dv_type = DataValidation(type="list", formula1='"SOP,Policy,Form,Manual,Guideline,Work Instruction,Record"', allow_blank=True)
dv_status = DataValidation(type="list", formula1='"Draft,Under Review,Approved,Active,Superseded,Archived"', allow_blank=True)
dv_dept = DataValidation(type="list", formula1='"Laboratory,Quality,Safety,Administration,IT,HR,All"', allow_blank=True)
ws.add_data_validation(dv_type)
ws.add_data_validation(dv_status)
ws.add_data_validation(dv_dept)

for i, (did, title, dtype, ver, status, author, approved, eff, review, dept, dist, loc) in enumerate(docs, 1):
    row = i + 1
    data_cell(ws, row, 1, i, align="center")
    data_cell(ws, row, 2, did, align="center", bold=True)
    data_cell(ws, row, 3, title)
    data_cell(ws, row, 4, dtype, align="center")
    data_cell(ws, row, 5, ver, align="center")
    data_cell(ws, row, 6, status, align="center")
    data_cell(ws, row, 7, author)
    data_cell(ws, row, 8, approved)
    data_cell(ws, row, 9, eff, align="center")
    data_cell(ws, row, 10, review, align="center")
    data_cell(ws, row, 11, dept, align="center")
    data_cell(ws, row, 12, dist)
    data_cell(ws, row, 13, loc)
    data_cell(ws, row, 14, f'=IF(J{row}="","",IF(TODAY()>DATEVALUE(J{row}),"OVERDUE",IF(DATEVALUE(J{row})-TODAY()<=30,"DUE SOON","OK")))', align="center", bold=True)
    data_cell(ws, row, 15, f'=IF(J{row}="","",IF(TODAY()>DATEVALUE(J{row}),TODAY()-DATEVALUE(J{row}),DATEVALUE(J{row})-TODAY()))', align="center", bold=True, fill=LIGHT_TEAL)
    dv_type.add(ws.cell(row=row, column=4))
    dv_status.add(ws.cell(row=row, column=6))
    dv_dept.add(ws.cell(row=row, column=11))

for row in range(len(docs)+2, len(docs)+32):
    data_cell(ws, row, 1, row-1, align="center")
    data_cell(ws, row, 14, f'=IF(J{row}="","",IF(TODAY()>DATEVALUE(J{row}),"OVERDUE",IF(DATEVALUE(J{row})-TODAY()<=30,"DUE SOON","OK")))', align="center", bold=True)
    data_cell(ws, row, 15, f'=IF(J{row}="","",IF(TODAY()>DATEVALUE(J{row}),TODAY()-DATEVALUE(J{row}),DATEVALUE(J{row})-TODAY()))', align="center", bold=True, fill=LIGHT_TEAL)
    dv_type.add(ws.cell(row=row, column=4))
    dv_status.add(ws.cell(row=row, column=6))
    dv_dept.add(ws.cell(row=row, column=11))

last_doc = len(docs) + 31

# Conditional formatting
ws.conditional_formatting.add(f"N2:N{last_doc}",
    CellIsRule(operator="equal", formula=['"OVERDUE"'], fill=PatternFill(start_color=RED, end_color=RED, fill_type="solid"), font=Font(bold=True, color=WHITE)))
ws.conditional_formatting.add(f"N2:N{last_doc}",
    CellIsRule(operator="equal", formula=['"DUE SOON"'], fill=PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")))
ws.conditional_formatting.add(f"O2:O{last_doc}",
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=365, color=GOLD))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:O{last_doc}"

# Dashboard
ws_d = wb.create_sheet("Dashboard")
ws_d.sheet_view.showGridLines = False
ws_d.column_dimensions["A"].width = 3
for col in "BCDEFGHIJKL":
    ws_d.column_dimensions[col].width = 14

ws_d.merge_cells("B2:L2")
c = ws_d.cell(row=2, column=2, value="Document Control Dashboard")
c.font = Font(name="Calibri", bold=True, size=18, color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_d.row_dimensions[2].height = 32

kpi_card(ws_d, 4, 2, "Total Documents", f'=COUNTA(\'Document Register\'!B2:B{last_doc})', NAVY, GOLD2)
kpi_card(ws_d, 4, 4, "Active", f'=COUNTIF(\'Document Register\'!F2:F{last_doc},"Active")', NAVY, GOLD2)
kpi_card(ws_d, 4, 6, "Overdue Reviews", f'=COUNTIF(\'Document Register\'!N2:N{last_doc},"OVERDUE")', NAVY, GOLD2)
kpi_card(ws_d, 4, 8, "Due Soon", f'=COUNTIF(\'Document Register\'!N2:N{last_doc},"DUE SOON")', NAVY, GOLD2)
kpi_card(ws_d, 4, 10, "Drafts", f'=COUNTIF(\'Document Register\'!F2:F{last_doc},"Draft")', NAVY, GOLD2)

# By type
add_chart_title(ws_d, 9, 2, "Documents by Type", 6)
ws_d.cell(row=10, column=2, value="Type")
ws_d.cell(row=10, column=3, value="Count")
for c in range(2, 4):
    hdr_cell(ws_d, 10, c, ws_d.cell(row=10, column=c).value)

doc_types = ["SOP", "Policy", "Form", "Manual", "Guideline", "Work Instruction", "Record"]
for i, dt in enumerate(doc_types, 1):
    row = 10 + i
    data_cell(ws_d, row, 2, dt)
    data_cell(ws_d, row, 3, f'=COUNTIF(\'Document Register\'!D2:D{last_doc},"{dt}")', align="center", bold=True)

chart1 = PieChart()
chart1.title = "Documents by Type"
data = Reference(ws_d, min_col=3, min_row=10, max_row=17)
cats = Reference(ws_d, min_col=2, min_row=11, max_row=17)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.height = 8
chart1.width = 12
chart1.dataLabels = DataLabelList(showPercent=True)
ws_d.add_chart(chart1, "E9")

# By status
add_chart_title(ws_d, 22, 2, "Documents by Status", 6)
ws_d.cell(row=23, column=2, value="Status")
ws_d.cell(row=23, column=3, value="Count")
for c in range(2, 4):
    hdr_cell(ws_d, 23, c, ws_d.cell(row=23, column=c).value)

doc_statuses = ["Draft", "Under Review", "Approved", "Active", "Superseded", "Archived"]
for i, st in enumerate(doc_statuses, 1):
    row = 23 + i
    data_cell(ws_d, row, 2, st)
    data_cell(ws_d, row, 3, f'=COUNTIF(\'Document Register\'!F2:F{last_doc},"{st}")', align="center", bold=True)

chart2 = BarChart()
chart2.type = "col"
chart2.style = 10
chart2.title = "Documents by Status"
chart2.y_axis.title = "Count"
data2 = Reference(ws_d, min_col=3, min_row=23, max_row=29)
cats2 = Reference(ws_d, min_col=2, min_row=24, max_row=29)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.height = 8
chart2.width = 14
ws_d.add_chart(chart2, "E22")

wb.save(os.path.join(OUTDIR, "ahqs-document-control-register.xlsx"))
print("  Saved: ahqs-document-control-register.xlsx")

# ══════════════════════════════════════════════════════════
# 6. PREMIUM COMPETENCY ASSESSMENT FORM
# ══════════════════════════════════════════════════════════
print("\n[6/6] Building Premium Competency Assessment Form...")
wb = openpyxl.Workbook()
add_cover(wb, "Competency Assessment Form",
    "A comprehensive staff competency evaluation tool with staff information, multi-method assessment matrix, competency gap analysis, and a coverage dashboard with charts.",
    "AHQS-COMP-002",
    ["Multi-method assessment (observation, records review, written test, oral exam)",
     "18 competency criteria across 6 domains",
     "Competency gap analysis with action plans",
     "Staff coverage matrix",
     "Renewal due date tracking",
     "Dashboard with competency rates and charts"])

add_settings_sheet(wb, "Settings", {
    "Results": ["Competent", "Needs Improvement", "Not Competent", "N/A"],
    "Methods": ["Direct Observation", "Records Review", "Written Test", "Oral Exam", "Performance Monitoring"],
    "Domains": ["Technical Skills", "Quality Management", "Safety", "Communication", "Knowledge", "Professional Development"],
    "Assessment_Type": ["Initial", "6-Month", "Annual", "Re-assessment"],
})

# Staff Information sheet
ws_info = wb.create_sheet("Staff Information")
ws_info.sheet_view.showGridLines = False
ws_info.column_dimensions["A"].width = 3
ws_info.column_dimensions["B"].width = 25
ws_info.column_dimensions["C"].width = 40

ws_info.merge_cells("B2:C2")
c = ws_info.cell(row=2, column=2, value="Staff Competency Assessment")
c.font = Font(name="Calibri", bold=True, size=18, color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_info.row_dimensions[2].height = 30

fields = [
    ("Staff Name", ""),
    ("Position / Title", ""),
    ("Department / Section", ""),
    ("Employee ID", ""),
    ("Date of Assessment", ""),
    ("Assessment Type", "Initial / 6-Month / Annual"),
    ("Assessor Name", ""),
    ("Assessor Title", ""),
    ("Next Assessment Due", ""),
]

for i, (label, placeholder) in enumerate(fields):
    r = 4 + i * 2
    c = ws_info.cell(row=r, column=2, value=label)
    c.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
    c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
    c.border = border_all
    ws_info.row_dimensions[r].height = 24

    c = ws_info.cell(row=r, column=3, value=placeholder)
    c.font = Font(name="Calibri", size=11, color=MUTED, italic=True)
    c.alignment = Alignment(vertical="center")
    c.fill = PatternFill(start_color=LIGHT_BG, end_color=LIGHT_BG, fill_type="solid")
    c.border = border_all

# Assessment Checklist
ws = wb.create_sheet("Assessment Checklist")
ws.sheet_view.showGridLines = False
headers = ["#", "Domain", "Assessment Method", "Criteria", "Result", "Score", "Comments", "Date Assessed", "Assessor", "Action Required"]
for c, h in enumerate(headers, 1):
    hdr_cell(ws, 1, c, h)
ws.row_dimensions[1].height = 36
widths = [4, 18, 18, 38, 14, 8, 25, 14, 14, 22]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

competencies = [
    ("Technical Skills", "Direct Observation", "Performs test procedures accurately following SOP"),
    ("Technical Skills", "Direct Observation", "Operates equipment correctly and safely"),
    ("Technical Skills", "Records Review", "Documents results accurately and completely"),
    ("Technical Skills", "Direct Observation", "Performs QC and interprets results correctly"),
    ("Technical Skills", "Performance Monitoring", "Troubleshoots instrument issues effectively"),
    ("Quality Management", "Records Review", "Follows CAPA procedures when issues are identified"),
    ("Quality Management", "Direct Observation", "Maintains chain of custody for samples"),
    ("Quality Management", "Records Review", "Completes required documentation within timelines"),
    ("Safety", "Direct Observation", "Uses appropriate PPE consistently"),
    ("Safety", "Direct Observation", "Follows chemical and biohazard safety protocols"),
    ("Safety", "Records Review", "Reports incidents and near-misses promptly"),
    ("Communication", "Oral Exam", "Communicates critical values per policy"),
    ("Communication", "Direct Observation", "Interacts professionally with patients and staff"),
    ("Knowledge", "Written Test", "Demonstrates knowledge of relevant quality standards"),
    ("Knowledge", "Oral Exam", "Understands emergency procedures and responses"),
    ("Knowledge", "Written Test", "Understands laboratory information system functionality"),
    ("Professional Development", "Records Review", "Completes required continuing education"),
    ("Professional Development", "Records Review", "Participates in proficiency testing program"),
    ("Professional Development", "Direct Observation", "Maintains professional conduct and appearance"),
]

dv_result = DataValidation(type="list", formula1='"Competent,Needs Improvement,Not Competent,N/A"', allow_blank=True)
dv_method = DataValidation(type="list", formula1='"Direct Observation,Records Review,Written Test,Oral Exam,Performance Monitoring"', allow_blank=True)
ws.add_data_validation(dv_result)
ws.add_data_validation(dv_method)

for i, (domain, method, criteria) in enumerate(competencies, 1):
    row = i + 1
    data_cell(ws, row, 1, i, align="center")
    data_cell(ws, row, 2, domain)
    data_cell(ws, row, 3, method)
    data_cell(ws, row, 4, criteria)
    data_cell(ws, row, 5, "")
    data_cell(ws, row, 6, f'=IF(E{row}="Competent",1,IF(E{row}="Needs Improvement",0.5,IF(E{row}="Not Competent",0,0)))', align="center", bold=True, fill=LIGHT_TEAL)
    data_cell(ws, row, 7, "")
    data_cell(ws, row, 8, "")
    data_cell(ws, row, 9, "")
    data_cell(ws, row, 10, f'=IF(E{row}="Not Competent","Yes",IF(E{row}="Needs Improvement","Yes","No"))', align="center")
    dv_result.add(ws.cell(row=row, column=5))
    dv_method.add(ws.cell(row=row, column=3))

last_comp = len(competencies) + 1

# Conditional formatting
ws.conditional_formatting.add(f"E2:E{last_comp}",
    CellIsRule(operator="equal", formula=['"Competent"'], fill=PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")))
ws.conditional_formatting.add(f"E2:E{last_comp}",
    CellIsRule(operator="equal", formula=['"Needs Improvement"'], fill=PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")))
ws.conditional_formatting.add(f"E2:E{last_comp}",
    CellIsRule(operator="equal", formula=['"Not Competent"'], fill=PatternFill(start_color=RED, end_color=RED, fill_type="solid"), font=Font(bold=True, color=WHITE)))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:J{last_comp}"

# Summary Dashboard
ws_d = wb.create_sheet("Dashboard")
ws_d.sheet_view.showGridLines = False
ws_d.column_dimensions["A"].width = 3
for col in "BCDEFGHIJKL":
    ws_d.column_dimensions[col].width = 14

ws_d.merge_cells("B2:L2")
c = ws_d.cell(row=2, column=2, value="Competency Assessment Dashboard")
c.font = Font(name="Calibri", bold=True, size=18, color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_d.row_dimensions[2].height = 32

kpi_card(ws_d, 4, 2, "Total Criteria", f'=COUNTA(\'Assessment Checklist\'!B2:B{last_comp})', NAVY, GOLD2)
kpi_card(ws_d, 4, 4, "Competent", f'=COUNTIF(\'Assessment Checklist\'!E2:E{last_comp},"Competent")', NAVY, GOLD2)
kpi_card(ws_d, 4, 6, "Needs Improvement", f'=COUNTIF(\'Assessment Checklist\'!E2:E{last_comp},"Needs Improvement")', NAVY, GOLD2)
kpi_card(ws_d, 4, 8, "Not Competent", f'=COUNTIF(\'Assessment Checklist\'!E2:E{last_comp},"Not Competent")', NAVY, GOLD2)
kpi_card(ws_d, 4, 10, "Overall %", f'=IFERROR(SUM(\'Assessment Checklist\'!F2:F{last_comp})/COUNTA(\'Assessment Checklist\'!B2:B{last_comp})*100,0)', NAVY, GOLD2)

# Domain breakdown
add_chart_title(ws_d, 9, 2, "Competency by Domain", 6)
ws_d.cell(row=10, column=2, value="Domain")
ws_d.cell(row=10, column=3, value="Total")
ws_d.cell(row=10, column=4, value="Competent")
ws_d.cell(row=10, column=5, value="Needs Imp.")
ws_d.cell(row=10, column=6, value="Not Comp.")
ws_d.cell(row=10, column=7, value="% Comp.")
for c in range(2, 8):
    hdr_cell(ws_d, 10, c, ws_d.cell(row=10, column=c).value)

domains = ["Technical Skills", "Quality Management", "Safety", "Communication", "Knowledge", "Professional Development"]
for i, domain in enumerate(domains, 1):
    row = 10 + i
    data_cell(ws_d, row, 2, domain)
    data_cell(ws_d, row, 3, f'=COUNTIF(\'Assessment Checklist\'!B2:B{last_comp},"{domain}")', align="center")
    data_cell(ws_d, row, 4, f'=COUNTIFS(\'Assessment Checklist\'!B2:B{last_comp},"{domain}",\'Assessment Checklist\'!E2:E{last_comp},"Competent")', align="center")
    data_cell(ws_d, row, 5, f'=COUNTIFS(\'Assessment Checklist\'!B2:B{last_comp},"{domain}",\'Assessment Checklist\'!E2:E{last_comp},"Needs Improvement")', align="center")
    data_cell(ws_d, row, 6, f'=COUNTIFS(\'Assessment Checklist\'!B2:B{last_comp},"{domain}",\'Assessment Checklist\'!E2:E{last_comp},"Not Competent")', align="center")
    data_cell(ws_d, row, 7, f'=IFERROR(D{row}/C{row}*100,0)', align="center", bold=True, fill=LIGHT_TEAL)

# Data bars
ws_d.conditional_formatting.add(f"G11:G16", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color=TEAL))

# Bar chart
chart1 = BarChart()
chart1.type = "col"
chart1.style = 10
chart1.title = "Competency Results by Domain"
chart1.y_axis.title = "Count"
data = Reference(ws_d, min_col=4, min_row=10, max_row=16, max_col=6)
cats = Reference(ws_d, min_col=2, min_row=11, max_row=16)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.height = 9
chart1.width = 16
ws_d.add_chart(chart1, "H9")

# Result distribution
add_chart_title(ws_d, 22, 2, "Result Distribution", 6)
ws_d.cell(row=23, column=2, value="Result")
ws_d.cell(row=23, column=3, value="Count")
ws_d.cell(row=23, column=4, value="Percentage")
for c in range(2, 5):
    hdr_cell(ws_d, 23, c, ws_d.cell(row=23, column=c).value)

results = ["Competent", "Needs Improvement", "Not Competent", "N/A"]
for i, result in enumerate(results, 1):
    row = 23 + i
    data_cell(ws_d, row, 2, result)
    data_cell(ws_d, row, 3, f'=COUNTIF(\'Assessment Checklist\'!E2:E{last_comp},"{result}")', align="center", bold=True)
    data_cell(ws_d, row, 4, f'=IFERROR(C{row}/COUNTA(\'Assessment Checklist\'!B2:B{last_comp})*100,0)', align="center", bold=True)

# Pie chart
chart2 = PieChart()
chart2.title = "Result Distribution"
data2 = Reference(ws_d, min_col=3, min_row=23, max_row=27)
cats2 = Reference(ws_d, min_col=2, min_row=24, max_row=27)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.height = 8
chart2.width = 12
chart2.dataLabels = DataLabelList(showPercent=True)
ws_d.add_chart(chart2, "E22")

wb.save(os.path.join(OUTDIR, "ahqs-competency-assessment-form.xlsx"))
print("  Saved: ahqs-competency-assessment-form.xlsx")

# ══════════════════════════════════════════════════════════
print("\n══════════════════════════════════════════════════════════")
print("ALL 6 PREMIUM TEMPLATES BUILT SUCCESSFULLY")
print("══════════════════════════════════════════════════════════")
for f in sorted(os.listdir(OUTDIR)):
    if f.endswith('.xlsx'):
        size = os.path.getsize(os.path.join(OUTDIR, f))
        print(f"  {f} ({size:,} bytes)")
