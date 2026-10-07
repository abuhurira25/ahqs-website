#!/usr/bin/env python3
"""Build 6 branded AHQS free resource Excel templates."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import os

# AHQS brand colors
NAVY = "071F3B"
TEAL = "087F83"
GOLD = "D8A72B"
LIGHT_BG = "F5F8FA"
BORDER_GRAY = "DCE6EB"
WHITE = "FFFFFF"

OUTDIR = "downloads/free-resources"
os.makedirs(OUTDIR, exist_ok=True)

# Create resized logo
logo_src = "ahqs-official-logo-hd.png"
logo_small = os.path.join(OUTDIR, "ahqs-logo-small.png")
if os.path.exists(logo_src):
    img = PILImage.open(logo_src)
    img.thumbnail((180, 60), PILImage.Resampling.LANCZOS)
    img.save(logo_small, "PNG")
    print(f"Logo resized to {logo_small}")


def style_header(ws, row, cols, fill_color=TEAL):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(
            left=Side(style="thin", color=BORDER_GRAY),
            right=Side(style="thin", color=BORDER_GRAY),
            top=Side(style="thin", color=BORDER_GRAY),
            bottom=Side(style="thin", color=BORDER_GRAY),
        )


def style_data_cell(cell, font_size=10):
    cell.font = Font(name="Calibri", size=font_size, color="173047")
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.border = Border(
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
        top=Side(style="thin", color=BORDER_GRAY),
        bottom=Side(style="thin", color=BORDER_GRAY),
    )


def add_cover_sheet(wb, title, description, tool_id):
    ws = wb.create_sheet("Cover & Instructions", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 4

    if os.path.exists(logo_small):
        img = XLImage(logo_small)
        img.anchor = "B2"
        ws.add_image(img)

    cell = ws.cell(row=8, column=2, value=title)
    cell.font = Font(name="Calibri", bold=True, size=20, color=NAVY)
    cell.alignment = Alignment(horizontal="left")

    cell = ws.cell(row=9, column=2, value="AHQS — Accelerate Healthcare Quality Solutions")
    cell.font = Font(name="Calibri", size=11, color=TEAL, italic=True)

    cell = ws.cell(row=11, column=2, value=description)
    cell.font = Font(name="Calibri", size=11, color="667789")
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[11].height = 60

    cell = ws.cell(row=13, column=2, value="How to Use This Tool")
    cell.font = Font(name="Calibri", bold=True, size=13, color=NAVY)

    instructions = [
        "1. Enter your organization name and facility details in the header area of each sheet",
        "2. Fill in the data columns — yellow-shaded cells contain formulas and should not be edited",
        "3. Use the dropdown menus where provided for consistent data entry",
        "4. Review the summary/dashboard sheet for auto-calculated metrics",
        "5. Save a copy for each audit cycle or reporting period",
        "6. This tool is free to use and customize for your organization's internal quality improvement",
    ]
    for i, instr in enumerate(instructions):
        cell = ws.cell(row=14 + i, column=2, value=instr)
        cell.font = Font(name="Calibri", size=10, color="173047")
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    cell = ws.cell(row=21, column=2, value="Disclaimer")
    cell.font = Font(name="Calibri", bold=True, size=13, color=NAVY)

    disclaimer = ("This is an AHQS-created quality improvement tool — not an official CAP, ISO 15189, "
                  "JCI or AABB standards document. Always purchase official standards directly from "
                  "the accrediting bodies. This template is provided free of charge for internal "
                  "quality improvement use. © 2026 Accelerate Healthcare Quality Solutions (AHQS). "
                  "Not for resale.")
    cell = ws.cell(row=22, column=2, value=disclaimer)
    cell.font = Font(name="Calibri", size=9, color="667789", italic=True)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[22].height = 60

    cell = ws.cell(row=24, column=2, value=f"Tool ID: {tool_id} | Version: 1.0 | Created: October 2026")
    cell.font = Font(name="Calibri", size=9, color="667789")
    print(f"  Cover sheet added for {title}")


# ============================================================
# 1. AUDIT TOOL TEMPLATE
# ============================================================
print("\n--- Building Audit Tool Template ---")
wb = openpyxl.Workbook()
add_cover_sheet(wb, "Quality Audit Tool Template",
    "A structured internal audit checklist template for laboratory and healthcare quality audits. Covers document control, competency, QC, CAPA, equipment, and safety.",
    "AHQS-AUD-001")

ws = wb.create_sheet("Audit Checklist")
ws.sheet_view.showGridLines = False

headers = ["#", "Audit Area", "Audit Question", "Reference", "Finding", "Risk Level", "Status", "Corrective Action", "Owner", "Due Date", "Date Closed"]
for c, h in enumerate(headers, 1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(headers))
ws.row_dimensions[1].height = 32

widths = [5, 20, 45, 15, 25, 12, 12, 30, 15, 12, 12]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

audit_areas = [
    ("Document Control", "Are SOPs current, reviewed, and approved?", "ISO 15189 §8.3"),
    ("Document Control", "Is there a master list of controlled documents?", "ISO 15189 §8.3"),
    ("Competency", "Are competency assessments completed for all staff?", "ISO 15189 §7.2"),
    ("Competency", "Are training records maintained and up to date?", "ISO 15189 §7.2"),
    ("Quality Control", "Are QC records reviewed daily?", "ISO 15189 §7.3"),
    ("Quality Control", "Are QC failures investigated and documented?", "ISO 15189 §7.3"),
    ("CAPA", "Are corrective actions tracked to closure?", "ISO 15189 §7.10"),
    ("CAPA", "Are root cause analyses performed for significant findings?", "ISO 15189 §7.10"),
    ("Equipment", "Is equipment calibrated per schedule?", "ISO 15189 §7.1"),
    ("Equipment", "Are maintenance records maintained?", "ISO 15189 §7.1"),
    ("Safety", "Is PPE available and used appropriately?", "ISO 15189 §7.1"),
    ("Safety", "Are safety incidents reported and investigated?", "ISO 15189 §7.1"),
    ("Personnel", "Are job descriptions current?", "ISO 15189 §7.2"),
    ("Personnel", "Are continuing education records maintained?", "ISO 15189 §7.2"),
    ("Facilities", "Is temperature/humidity monitored and recorded?", "ISO 15189 §7.1"),
]

risk_validation = DataValidation(type="list", formula1='"High,Medium,Low"', allow_blank=True)
status_validation = DataValidation(type="list", formula1='"Open,In Progress,Closed,Not Applicable"', allow_blank=True)
ws.add_data_validation(risk_validation)
ws.add_data_validation(status_validation)

for i, (area, question, ref) in enumerate(audit_areas, 1):
    row = i + 1
    ws.cell(row=row, column=1, value=i)
    ws.cell(row=row, column=2, value=area)
    ws.cell(row=row, column=3, value=question)
    ws.cell(row=row, column=4, value=ref)
    for c in range(1, len(headers) + 1):
        style_data_cell(ws.cell(row=row, column=c))
    risk_validation.add(ws.cell(row=row, column=6))
    status_validation.add(ws.cell(row=row, column=7))

# Empty rows for user input
for row in range(len(audit_areas) + 2, len(audit_areas) + 22):
    for c in range(1, len(headers) + 1):
        style_data_cell(ws.cell(row=row, column=c))
    risk_validation.add(ws.cell(row=row, column=6))
    status_validation.add(ws.cell(row=row, column=7))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:K{len(audit_areas)+21}"

# Summary
ws_sum = wb.create_sheet("Summary")
ws_sum.sheet_view.showGridLines = False
ws_sum.column_dimensions["A"].width = 4
ws_sum.column_dimensions["B"].width = 30
ws_sum.column_dimensions["C"].width = 15
ws_sum.column_dimensions["D"].width = 15
ws_sum.column_dimensions["E"].width = 15
ws_sum.column_dimensions["F"].width = 15

ws_sum.cell(row=2, column=2, value="Audit Summary").font = Font(name="Calibri", bold=True, size=16, color=NAVY)

for c, h in enumerate(["", "Status", "Count", "High Risk", "Medium Risk", "Low Risk"], 1):
    ws_sum.cell(row=4, column=c, value=h)
style_header(ws_sum, 4, 6)

statuses = ["Open", "In Progress", "Closed", "Not Applicable"]
for i, status in enumerate(statuses):
    r = 5 + i
    ws_sum.cell(row=r, column=2, value=status)
    ws_sum.cell(row=r, column=3, value=f'=COUNTIF(\'Audit Checklist\'!G:G,"{status}")')
    ws_sum.cell(row=r, column=4, value=f'=COUNTIFS(\'Audit Checklist\'!G:G,"{status}",\'Audit Checklist\'!F:F,"High")')
    ws_sum.cell(row=r, column=5, value=f'=COUNTIFS(\'Audit Checklist\'!G:G,"{status}",\'Audit Checklist\'!F:F,"Medium")')
    ws_sum.cell(row=r, column=6, value=f'=COUNTIFS(\'Audit Checklist\'!G:G,"{status}",\'Audit Checklist\'!F:F,"Low")')
    for c in range(1, 7):
        style_data_cell(ws_sum.cell(row=r, column=c))

r = 9
ws_sum.cell(row=r, column=2, value="TOTAL")
ws_sum.cell(row=r, column=3, value='=SUM(C5:C8)')
for c in range(1, 7):
    cell = ws_sum.cell(row=r, column=c)
    cell.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
    cell.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    cell.alignment = Alignment(horizontal="center")

wb.save(os.path.join(OUTDIR, "ahqs-audit-tool-template.xlsx"))
print("  Saved: ahqs-audit-tool-template.xlsx")

# ============================================================
# 2. RISK REGISTER TEMPLATE
# ============================================================
print("\n--- Building Risk Register Template ---")
wb = openpyxl.Workbook()
add_cover_sheet(wb, "Risk Register Template",
    "A structured risk register with severity and likelihood scoring. Track risks, owners, mitigation actions, and status.",
    "AHQS-RSK-001")

ws = wb.create_sheet("Risk Register")
ws.sheet_view.showGridLines = False

headers = ["#", "Risk ID", "Date Identified", "Risk Category", "Risk Description", "Likelihood (1-5)", "Severity (1-5)", "Risk Score", "Risk Rating", "Risk Owner", "Mitigation Action", "Target Date", "Status", "Review Date"]
for c, h in enumerate(headers, 1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(headers))
ws.row_dimensions[1].height = 32

widths = [5, 12, 14, 18, 40, 14, 14, 12, 14, 15, 35, 14, 14, 14]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

risk_validation = DataValidation(type="list", formula1='"Low,Medium,High,Critical"', allow_blank=True)
status_validation = DataValidation(type="list", formula1='"Open,Mitigated,Closed,Monitoring"', allow_blank=True)
cat_validation = DataValidation(type="list", formula1='"Clinical,Operational,Financial,Compliance,Safety,Equipment,IT/Data,Personnel"', allow_blank=True)
ws.add_data_validation(risk_validation)
ws.add_data_validation(status_validation)
ws.add_data_validation(cat_validation)

sample_risks = [
    ("R001", "2026-01-15", "Clinical", "Delayed critical value reporting to requesting physician", 3, 5),
    ("R002", "2026-01-20", "Compliance", "Expired reagents in use on automated analyzer", 2, 4),
    ("R003", "2026-02-01", "Safety", "Inadequate PPE compliance in specimen processing area", 3, 3),
    ("R004", "2026-02-10", "Equipment", "Analyzer calibration drift affecting QC results", 2, 5),
    ("R005", "2026-02-15", "IT/Data", "LIS downtime during peak testing hours", 3, 4),
]

for i, (rid, date, cat, desc, lik, sev) in enumerate(sample_risks, 1):
    row = i + 1
    ws.cell(row=row, column=1, value=i)
    ws.cell(row=row, column=2, value=rid)
    ws.cell(row=row, column=3, value=date)
    ws.cell(row=row, column=4, value=cat)
    ws.cell(row=row, column=5, value=desc)
    ws.cell(row=row, column=6, value=lik)
    ws.cell(row=row, column=7, value=sev)
    ws.cell(row=row, column=8, value=f"=F{row}*G{row}")
    ws.cell(row=row, column=9, value=f'=IF(H{row}>=15,"Critical",IF(H{row}>=9,"High",IF(H{row}>=4,"Medium","Low")))')
    for c in range(1, len(headers) + 1):
        style_data_cell(ws.cell(row=row, column=c))
    risk_validation.add(ws.cell(row=row, column=9))
    status_validation.add(ws.cell(row=row, column=13))
    cat_validation.add(ws.cell(row=row, column=4))

# Empty rows with formulas
for row in range(len(sample_risks) + 2, len(sample_risks) + 27):
    ws.cell(row=row, column=1, value=row - 1)
    ws.cell(row=row, column=8, value=f"=F{row}*G{row}")
    ws.cell(row=row, column=9, value=f'=IF(H{row}>=15,"Critical",IF(H{row}>=9,"High",IF(H{row}>=4,"Medium","Low")))')
    for c in range(1, len(headers) + 1):
        style_data_cell(ws.cell(row=row, column=c))
    risk_validation.add(ws.cell(row=row, column=9))
    status_validation.add(ws.cell(row=row, column=13))
    cat_validation.add(ws.cell(row=row, column=4))

# Conditional formatting for risk score
ws.conditional_formatting.add(f"H2:H{len(sample_risks)+26}",
    CellIsRule(operator="greaterThanOrEqual", formula=["15"],
               fill=PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")))
ws.conditional_formatting.add(f"H2:H{len(sample_risks)+26}",
    CellIsRule(operator="between", formula=["9", "14"],
               fill=PatternFill(start_color="FFA94D", end_color="FFA94D", fill_type="solid")))
ws.conditional_formatting.add(f"H2:H{len(sample_risks)+26}",
    CellIsRule(operator="between", formula=["4", "8"],
               fill=PatternFill(start_color="FFE066", end_color="FFE066", fill_type="solid")))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:N{len(sample_risks)+26}"

# Risk Matrix sheet
ws_matrix = wb.create_sheet("Risk Matrix")
ws_matrix.sheet_view.showGridLines = False
ws_matrix.cell(row=2, column=2, value="Risk Assessment Matrix").font = Font(name="Calibri", bold=True, size=16, color=NAVY)

ws_matrix.cell(row=4, column=2, value="Likelihood / Severity")
ws_matrix.cell(row=4, column=2).font = Font(bold=True, size=10, color=NAVY)
ws_matrix.cell(row=4, column=2).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

sev_labels = ["1 - Rare", "2 - Unlikely", "3 - Possible", "4 - Likely", "5 - Almost Certain"]
lik_labels = ["1 - Rare", "2 - Unlikely", "3 - Possible", "4 - Likely", "5 - Almost Certain"]

for c in range(5):
    cell = ws_matrix.cell(row=5, column=3 + c, value=sev_labels[c])
    cell.font = Font(bold=True, size=9, color=WHITE)
    cell.fill = PatternFill(start_color=TEAL, end_color=TEAL, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", wrap_text=True)

for r in range(5):
    cell = ws_matrix.cell(row=6 + r, column=2, value=lik_labels[4 - r])
    cell.font = Font(bold=True, size=9, color=WHITE)
    cell.fill = PatternFill(start_color=TEAL, end_color=TEAL, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for c in range(5):
        score = (5 - r) * (c + 1)
        cell = ws_matrix.cell(row=6 + r, column=3 + c, value=score)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(size=11, bold=True)
        if score >= 15:
            cell.fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
        elif score >= 9:
            cell.fill = PatternFill(start_color="FFA94D", end_color="FFA94D", fill_type="solid")
        elif score >= 4:
            cell.fill = PatternFill(start_color="FFE066", end_color="FFE066", fill_type="solid")
        else:
            cell.fill = PatternFill(start_color="C3F6C3", end_color="C3F6C3", fill_type="solid")

ws_matrix.column_dimensions["B"].width = 22
for c in range(5):
    ws_matrix.column_dimensions[get_column_letter(3 + c)].width = 16

wb.save(os.path.join(OUTDIR, "ahqs-risk-register-template.xlsx"))
print("  Saved: ahqs-risk-register-template.xlsx")

# ============================================================
# 3. CAPA TRACKING TEMPLATE
# ============================================================
print("\n--- Building CAPA Tracking Template ---")
wb = openpyxl.Workbook()
add_cover_sheet(wb, "CAPA Tracking Template",
    "Corrective and Preventive Action tracker with root cause analysis, effectiveness checks, and overdue auto-flagging.",
    "AHQS-CAPA-001")

ws = wb.create_sheet("CAPA Register")
ws.sheet_view.showGridLines = False

headers = ["#", "CAPA ID", "Date Identified", "Source", "Issue Description", "Root Cause", "Corrective Action", "Preventive Action", "Owner", "Due Date", "Status", "Date Closed", "Effectiveness Check", "Days Overdue"]
for c, h in enumerate(headers, 1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(headers))
ws.row_dimensions[1].height = 32

widths = [5, 12, 14, 18, 35, 30, 30, 30, 15, 14, 14, 14, 25, 12]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

source_validation = DataValidation(type="list", formula1='"Audit Finding,Incident,Complaint,External Assessment,Internal Review,QC Failure,Safety Event"', allow_blank=True)
status_validation = DataValidation(type="list", formula1='"Open,In Progress,Closed,Overdue,Cancelled"', allow_blank=True)
ws.add_data_validation(source_validation)
ws.add_data_validation(status_validation)

sample_capas = [
    ("CAPA-001", "2026-01-10", "Audit Finding", "QC records not reviewed by supervisor within 24 hours", "Supervisor workload exceeded capacity during peak hours", "Reassigned QC review to alternate supervisor during peak", "Implemented automated QC alert system", "Lab Manager", "2026-02-15", "Closed", "2026-02-10", "Effective - no recurrence in 3 months"),
    ("CAPA-002", "2026-01-20", "Incident", "Expired reagent used in patient testing", "Reagent inventory not checked before use", "Discarded expired reagents, retested affected samples", "Implemented barcode scanning for reagent verification", "Quality Officer", "2026-02-28", "In Progress", "", ""),
    ("CAPA-003", "2026-02-05", "External Assessment", "Competency assessment not documented for 2 staff members", "Competency records stored in different locations", "Consolidated all competency records into central file", "Implemented quarterly competency review schedule", "HR Manager", "2026-03-15", "Open", "", ""),
]

for i, (cid, date, source, issue, rca, ca, pa, owner, due, status, closed, eff) in enumerate(sample_capas, 1):
    row = i + 1
    ws.cell(row=row, column=1, value=i)
    ws.cell(row=row, column=2, value=cid)
    ws.cell(row=row, column=3, value=date)
    ws.cell(row=row, column=4, value=source)
    ws.cell(row=row, column=5, value=issue)
    ws.cell(row=row, column=6, value=rca)
    ws.cell(row=row, column=7, value=ca)
    ws.cell(row=row, column=8, value=pa)
    ws.cell(row=row, column=9, value=owner)
    ws.cell(row=row, column=10, value=due)
    ws.cell(row=row, column=11, value=status)
    ws.cell(row=row, column=12, value=closed)
    ws.cell(row=row, column=13, value=eff)
    ws.cell(row=row, column=14, value=f'=IF(AND(K{row}<>"Closed",K{row}<>"Cancelled",J{row}<>"",TODAY()>DATEVALUE(J{row})),TODAY()-DATEVALUE(J{row}),0)')
    for c in range(1, len(headers) + 1):
        style_data_cell(ws.cell(row=row, column=c))
    source_validation.add(ws.cell(row=row, column=4))
    status_validation.add(ws.cell(row=row, column=11))

# Empty rows
for row in range(len(sample_capas) + 2, len(sample_capas) + 27):
    ws.cell(row=row, column=1, value=row - 1)
    ws.cell(row=row, column=14, value=f'=IF(AND(K{row}<>"Closed",K{row}<>"Cancelled",J{row}<>"",TODAY()>DATEVALUE(J{row})),TODAY()-DATEVALUE(J{row}),0)')
    for c in range(1, len(headers) + 1):
        style_data_cell(ws.cell(row=row, column=c))
    source_validation.add(ws.cell(row=row, column=4))
    status_validation.add(ws.cell(row=row, column=11))

ws.conditional_formatting.add(f"N2:N{len(sample_capas)+26}",
    CellIsRule(operator="greaterThan", formula=["0"],
               fill=PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid"),
               font=Font(bold=True, color=WHITE)))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:N{len(sample_capas)+26}"

# Summary
ws_sum = wb.create_sheet("Summary")
ws_sum.sheet_view.showGridLines = False
ws_sum.column_dimensions["B"].width = 20
ws_sum.column_dimensions["C"].width = 15
ws_sum.column_dimensions["D"].width = 15

ws_sum.cell(row=2, column=2, value="CAPA Summary").font = Font(name="Calibri", bold=True, size=16, color=NAVY)

for c, h in enumerate(["", "Status", "Count", "Overdue"], 1):
    ws_sum.cell(row=4, column=c, value=h)
style_header(ws_sum, 4, 4)

statuses = ["Open", "In Progress", "Closed", "Overdue", "Cancelled"]
for i, status in enumerate(statuses):
    r = 5 + i
    ws_sum.cell(row=r, column=2, value=status)
    ws_sum.cell(row=r, column=3, value=f'=COUNTIF(\'CAPA Register\'!K:K,"{status}")')
    ws_sum.cell(row=r, column=4, value=f'=COUNTIFS(\'CAPA Register\'!K:K,"{status}",\'CAPA Register\'!N:N,">0")')
    for c in range(1, 5):
        style_data_cell(ws_sum.cell(row=r, column=c))

wb.save(os.path.join(OUTDIR, "ahqs-capa-tracking-template.xlsx"))
print("  Saved: ahqs-capa-tracking-template.xlsx")

# ============================================================
# 4. QUALITY INDICATOR DASHBOARD
# ============================================================
print("\n--- Building Quality Indicator Dashboard ---")
wb = openpyxl.Workbook()
add_cover_sheet(wb, "Quality Indicator Dashboard",
    "KPI tracking template for laboratory quality indicators including TAT, rejection rates, critical value reporting, and CAPA closure.",
    "AHQS-KPI-001")

ws = wb.create_sheet("KPI Data")
ws.sheet_view.showGridLines = False

headers = ["#", "Month", "Indicator", "Target", "Actual", "Unit", "Status", "Trend", "Comments"]
for c, h in enumerate(headers, 1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(headers))
ws.row_dimensions[1].height = 32

widths = [5, 12, 35, 12, 12, 10, 12, 10, 30]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

months = ["Jan-2026", "Feb-2026", "Mar-2026", "Apr-2026", "May-2026", "Jun-2026", "Jul-2026", "Aug-2026", "Sep-2026", "Oct-2026", "Nov-2026", "Dec-2026"]
indicators = [
    ("Turnaround Time - Routine", 4, "hours", "≤"),
    ("Turnaround Time - STAT", 2, "hours", "≤"),
    ("Sample Rejection Rate", 2, "%", "≤"),
    ("Critical Value Reporting Time", 30, "minutes", "≤"),
    ("CAPA Closure Rate (30 days)", 90, "%", "≥"),
    ("QC Pass Rate", 95, "%", "≥"),
    ("Proficiency Testing Score", 80, "%", "≥"),
    ("Repeat Test Rate", 5, "%", "≤"),
]

status_validation = DataValidation(type="list", formula1='"Met,Not Met,N/A"', allow_blank=True)
ws.add_data_validation(status_validation)

row = 2
for indicator, target, unit, direction in indicators:
    for month in months:
        ws.cell(row=row, column=1, value=row - 1)
        ws.cell(row=row, column=2, value=month)
        ws.cell(row=row, column=3, value=indicator)
        ws.cell(row=row, column=4, value=target)
        ws.cell(row=row, column=6, value=unit)
        ws.cell(row=row, column=7, value=f'=IF(E{row}="","",IF("{direction}"="≥",IF(E{row}>={target},"Met","Not Met"),IF(E{row}<={target},"Met","Not Met")))')
        for c in range(1, len(headers) + 1):
            style_data_cell(ws.cell(row=row, column=c))
        status_validation.add(ws.cell(row=row, column=7))
        row += 1

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:I{row-1}"

# Summary dashboard
ws_sum = wb.create_sheet("Summary Dashboard")
ws_sum.sheet_view.showGridLines = False
ws_sum.column_dimensions["B"].width = 35
ws_sum.column_dimensions["C"].width = 12
ws_sum.column_dimensions["D"].width = 12
ws_sum.column_dimensions["E"].width = 12
ws_sum.column_dimensions["F"].width = 15

ws_sum.cell(row=2, column=2, value="Quality Indicator Dashboard").font = Font(name="Calibri", bold=True, size=16, color=NAVY)

for c, h in enumerate(["", "Indicator", "Target", "Latest Actual", "Status", "% Met (YTD)"], 1):
    ws_sum.cell(row=4, column=c, value=h)
style_header(ws_sum, 4, 6)

for i, (indicator, target, unit, direction) in enumerate(indicators):
    r = 5 + i
    ws_sum.cell(row=r, column=2, value=f"{indicator} ({unit})")
    ws_sum.cell(row=r, column=3, value=target)
    # Latest actual: lookup last non-empty value
    data_start = 2 + i * 12
    data_end = data_start + 11
    ws_sum.cell(row=r, column=4, value=f"=IFERROR(LOOKUP(2,1/('KPI Data'!E{data_start}:E{data_end}<>""),'KPI Data'!E{data_start}:E{data_end}),\"\")")
    ws_sum.cell(row=r, column=5, value=f'=IF(D{r}="","",IF("{direction}"="≥",IF(D{r}>={target},"Met","Not Met"),IF(D{r}<={target},"Met","Not Met")))')
    ws_sum.cell(row=r, column=6, value=f'=IFERROR(COUNTIF(\'KPI Data\'!G{data_start}:G{data_end},"Met")/MAX(1,COUNTA(\'KPI Data\'!G{data_start}:G{data_end})),0)')
    for c in range(1, 7):
        style_data_cell(ws_sum.cell(row=r, column=c))

wb.save(os.path.join(OUTDIR, "ahqs-quality-indicator-dashboard.xlsx"))
print("  Saved: ahqs-quality-indicator-dashboard.xlsx")

# ============================================================
# 5. DOCUMENT CONTROL REGISTER
# ============================================================
print("\n--- Building Document Control Register ---")
wb = openpyxl.Workbook()
add_cover_sheet(wb, "Document Control Register",
    "Track SOPs, policies, forms, and manuals with version control, review dates, approval status, and distribution tracking.",
    "AHQS-DOC-001")

ws = wb.create_sheet("Document Register")
ws.sheet_view.showGridLines = False

headers = ["#", "Doc ID", "Document Title", "Document Type", "Version", "Status", "Author", "Approved By", "Effective Date", "Next Review", "Distribution", "Location/Link", "Review Status"]
for c, h in enumerate(headers, 1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(headers))
ws.row_dimensions[1].height = 32

widths = [5, 12, 35, 15, 10, 12, 15, 15, 14, 14, 20, 25, 14]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

type_validation = DataValidation(type="list", formula1='"SOP,Policy,Form,Manual,Guideline,Work Instruction,Record"', allow_blank=True)
status_validation = DataValidation(type="list", formula1='"Draft,Under Review,Approved,Active,Superseded,Archived"', allow_blank=True)
ws.add_data_validation(type_validation)
ws.add_data_validation(status_validation)

sample_docs = [
    ("DOC-001", "Sample Receipt and Accessioning SOP", "SOP", "3.1", "Active", "Lab Tech", "Lab Manager", "2026-01-01", "2027-01-01", "All lab staff", "Shared Drive / SOP-001"),
    ("DOC-002", "Critical Value Notification Policy", "Policy", "2.0", "Active", "Quality Officer", "Medical Director", "2026-01-15", "2027-01-15", "All staff", "Shared Drive / POL-002"),
    ("DOC-003", "Equipment Calibration Record Form", "Form", "1.5", "Active", "Lab Manager", "Quality Manager", "2026-02-01", "2027-02-01", "Lab staff", "Shared Drive / FRM-003"),
    ("DOC-004", "Laboratory Safety Manual", "Manual", "4.0", "Active", "Safety Officer", "Lab Director", "2026-01-01", "2027-01-01", "All staff", "Shared Drive / MAN-004"),
    ("DOC-005", "CAPA Investigation Form", "Form", "2.1", "Under Review", "Quality Officer", "", "", "", "Quality team", "Shared Drive / FRM-005"),
]

for i, (did, title, dtype, ver, status, author, approved, eff, review, dist, loc) in enumerate(sample_docs, 1):
    row = i + 1
    ws.cell(row=row, column=1, value=i)
    ws.cell(row=row, column=2, value=did)
    ws.cell(row=row, column=3, value=title)
    ws.cell(row=row, column=4, value=dtype)
    ws.cell(row=row, column=5, value=ver)
    ws.cell(row=row, column=6, value=status)
    ws.cell(row=row, column=7, value=author)
    ws.cell(row=row, column=8, value=approved)
    ws.cell(row=row, column=9, value=eff)
    ws.cell(row=row, column=10, value=review)
    ws.cell(row=row, column=11, value=dist)
    ws.cell(row=row, column=12, value=loc)
    ws.cell(row=row, column=13, value=f'=IF(J{row}="","",IF(TODAY()>DATEVALUE(J{row}),"OVERDUE","OK"))')
    for c in range(1, len(headers) + 1):
        style_data_cell(ws.cell(row=row, column=c))
    type_validation.add(ws.cell(row=row, column=4))
    status_validation.add(ws.cell(row=row, column=6))

# Empty rows
for row in range(len(sample_docs) + 2, len(sample_docs) + 27):
    ws.cell(row=row, column=1, value=row - 1)
    ws.cell(row=row, column=13, value=f'=IF(J{row}="","",IF(TODAY()>DATEVALUE(J{row}),"OVERDUE","OK"))')
    for c in range(1, len(headers) + 1):
        style_data_cell(ws.cell(row=row, column=c))
    type_validation.add(ws.cell(row=row, column=4))
    status_validation.add(ws.cell(row=row, column=6))

ws.conditional_formatting.add(f"M2:M{len(sample_docs)+26}",
    CellIsRule(operator="equal", formula=['"OVERDUE"'],
               fill=PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid"),
               font=Font(bold=True, color=WHITE)))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:M{len(sample_docs)+26}"

# Summary
ws_sum = wb.create_sheet("Summary")
ws_sum.sheet_view.showGridLines = False
ws_sum.column_dimensions["B"].width = 20
ws_sum.column_dimensions["C"].width = 12

ws_sum.cell(row=2, column=2, value="Document Control Summary").font = Font(name="Calibri", bold=True, size=16, color=NAVY)

for c, h in enumerate(["", "Status", "Count"], 1):
    ws_sum.cell(row=4, column=c, value=h)
style_header(ws_sum, 4, 3)

statuses = ["Draft", "Under Review", "Approved", "Active", "Superseded", "Archived"]
for i, status in enumerate(statuses):
    r = 5 + i
    ws_sum.cell(row=r, column=2, value=status)
    ws_sum.cell(row=r, column=3, value=f'=COUNTIF(\'Document Register\'!F:F,"{status}")')
    for c in range(1, 4):
        style_data_cell(ws_sum.cell(row=r, column=c))

r = 11
ws_sum.cell(row=r, column=2, value="Reviews Overdue")
ws_sum.cell(row=r, column=3, value='=COUNTIF(\'Document Register\'!M:M,"OVERDUE")')
ws_sum.cell(row=r, column=2).font = Font(bold=True, color="FF6B6B")
ws_sum.cell(row=r, column=3).font = Font(bold=True, color="FF6B6B")

wb.save(os.path.join(OUTDIR, "ahqs-document-control-register.xlsx"))
print("  Saved: ahqs-document-control-register.xlsx")

# ============================================================
# 6. COMPETENCY ASSESSMENT FORM
# ============================================================
print("\n--- Building Competency Assessment Form ---")
wb = openpyxl.Workbook()
add_cover_sheet(wb, "Competency Assessment Form",
    "Staff competency evaluation template for initial and ongoing assessment. Covers direct observation, records review, and more.",
    "AHQS-COMP-001")

# Staff info sheet
ws_info = wb.create_sheet("Staff Information")
ws_info.sheet_view.showGridLines = False
ws_info.column_dimensions["B"].width = 25
ws_info.column_dimensions["C"].width = 40

ws_info.cell(row=2, column=2, value="Staff Competency Assessment").font = Font(name="Calibri", bold=True, size=16, color=NAVY)

fields = [
    ("Staff Name", ""),
    ("Position/Title", ""),
    ("Department/Section", ""),
    ("Employee ID", ""),
    ("Date of Assessment", ""),
    ("Assessment Type", "Initial / 6-Month / Annual"),
    ("Assessor Name", ""),
    ("Assessor Title", ""),
]

for i, (label, placeholder) in enumerate(fields):
    r = 4 + i * 2
    cell = ws_info.cell(row=r, column=2, value=label)
    cell.font = Font(name="Calibri", bold=True, size=11, color=NAVY)
    cell.fill = PatternFill(start_color=LIGHT_BG, end_color=LIGHT_BG, fill_type="solid")
    cell.alignment = Alignment(vertical="center")
    cell.border = Border(bottom=Side(style="thin", color=BORDER_GRAY))

    cell = ws_info.cell(row=r, column=3, value=placeholder)
    cell.font = Font(name="Calibri", size=11, color="667789", italic=True)
    cell.alignment = Alignment(vertical="center")
    cell.border = Border(bottom=Side(style="thin", color=BORDER_GRAY))

# Assessment checklist
ws = wb.create_sheet("Assessment Checklist")
ws.sheet_view.showGridLines = False

headers = ["#", "Competency Area", "Assessment Method", "Criteria", "Result", "Comments", "Date Assessed", "Assessor"]
for c, h in enumerate(headers, 1):
    ws.cell(row=1, column=c, value=h)
style_header(ws, 1, len(headers))
ws.row_dimensions[1].height = 32

widths = [5, 22, 20, 40, 12, 25, 14, 15]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

result_validation = DataValidation(type="list", formula1='"Competent,Needs Improvement,Not Competent,N/A"', allow_blank=True)
method_validation = DataValidation(type="list", formula1='"Direct Observation,Records Review,Written Test,Oral Exam,Performance Monitoring"', allow_blank=True)
ws.add_data_validation(result_validation)
ws.add_data_validation(method_validation)

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
    ("Knowledge", "Written Test", "Demonstrates knowledge of relevant standards (ISO/CAP/JCI)"),
    ("Knowledge", "Oral Exam", "Understands emergency procedures and responses"),
    ("Professional Development", "Records Review", "Completes required continuing education"),
    ("Professional Development", "Records Review", "Participates in proficiency testing program"),
    ("Professional Development", "Direct Observation", "Maintains professional conduct and appearance"),
]

for i, (area, method, criteria) in enumerate(competencies, 1):
    row = i + 1
    ws.cell(row=row, column=1, value=i)
    ws.cell(row=row, column=2, value=area)
    ws.cell(row=row, column=3, value=method)
    ws.cell(row=row, column=4, value=criteria)
    for c in range(1, len(headers) + 1):
        style_data_cell(ws.cell(row=row, column=c))
    result_validation.add(ws.cell(row=row, column=5))
    method_validation.add(ws.cell(row=row, column=3))

# Conditional formatting
ws.conditional_formatting.add(f"E2:E{len(competencies)+1}",
    CellIsRule(operator="equal", formula=['"Not Competent"'],
               fill=PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid"),
               font=Font(bold=True, color=WHITE)))
ws.conditional_formatting.add(f"E2:E{len(competencies)+1}",
    CellIsRule(operator="equal", formula=['"Needs Improvement"'],
               fill=PatternFill(start_color="FFE066", end_color="FFE066", fill_type="solid")))
ws.conditional_formatting.add(f"E2:E{len(competencies)+1}",
    CellIsRule(operator="equal", formula=['"Competent"'],
               fill=PatternFill(start_color="C3F6C3", end_color="C3F6C3", fill_type="solid")))

ws.freeze_panes = "A2"

# Summary
ws_sum = wb.create_sheet("Summary")
ws_sum.sheet_view.showGridLines = False
ws_sum.column_dimensions["B"].width = 22
ws_sum.column_dimensions["C"].width = 12

ws_sum.cell(row=2, column=2, value="Competency Summary").font = Font(name="Calibri", bold=True, size=16, color=NAVY)

for c, h in enumerate(["", "Result", "Count"], 1):
    ws_sum.cell(row=4, column=c, value=h)
style_header(ws_sum, 4, 3)

results = ["Competent", "Needs Improvement", "Not Competent", "N/A"]
for i, result in enumerate(results):
    r = 5 + i
    ws_sum.cell(row=r, column=2, value=result)
    ws_sum.cell(row=r, column=3, value=f'=COUNTIF(\'Assessment Checklist\'!E:E,"{result}")')
    for c in range(1, 4):
        style_data_cell(ws_sum.cell(row=r, column=c))

r = 9
ws_sum.cell(row=r, column=2, value="Total Assessed")
ws_sum.cell(row=r, column=3, value="=SUM(C5:C8)")
ws_sum.cell(row=r, column=2).font = Font(bold=True, color=NAVY)
ws_sum.cell(row=r, column=3).font = Font(bold=True, color=NAVY)

wb.save(os.path.join(OUTDIR, "ahqs-competency-assessment-form.xlsx"))
print("  Saved: ahqs-competency-assessment-form.xlsx")

# ============================================================
print("\n--- All 6 files created ---")
for f in sorted(os.listdir(OUTDIR)):
    if f.endswith('.xlsx'):
        size = os.path.getsize(os.path.join(OUTDIR, f))
        print(f"  {f} ({size:,} bytes)")
