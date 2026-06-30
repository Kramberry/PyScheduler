from flask import Flask, render_template, request, send_file, redirect, url_for, Response
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from reportlab.lib.pagesizes import LETTER
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import json
import os
import sys
import webbrowser
import logging
APP_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.abspath(".")
EMP_FILE = os.path.join(APP_DIR, "employees.json")
SCHEDULE_FILE = os.path.join(APP_DIR, "last_schedule.json")

logging.basicConfig(filename=os.path.join(APP_DIR, "app.log"), level=logging.INFO)

def resource_path(relative: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, relative)

app = Flask(
    __name__,
    template_folder=resource_path("templates"),
)

def to_24h(time_str):
    if not time_str:
        return ''
    try:
        from datetime import datetime
        return datetime.strptime(time_str.strip(), "%I:%M %p").strftime("%H:%M")
    except Exception:
        return ''

app.jinja_env.filters['to24h'] = to_24h


DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]


# ------------------------------------
# Helpers to load/save employees
# ------------------------------------
def load_employees():
    if not os.path.exists(EMP_FILE):
        # default list the first time
        default_employees = [
            "Nora",
            "Viviana",
            "Sonia",
            "Carlet",
            "Marisol",
            "Patricia",
            "Trinidad",
            "Yessica",
            "Marcia",
            "Brandon",
            "Matthew",
            "Sue",
            "Raishod",
            "Levent",
            "Carolina",
        ]
        save_employees(default_employees)
        return default_employees

    with open(EMP_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_employees(employees_list):
    with open(EMP_FILE, "w", encoding="utf-8") as f:
        json.dump(employees_list, f, indent=2)

def load_last_schedule():
    if not os.path.exists(SCHEDULE_FILE):
        return {}
    with open(SCHEDULE_FILE, "r", encoding="utf-8") as f:
        return json.load(f) 

def save_last_schedule(data):
    with open(SCHEDULE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)      


# ------------------------------------
# Calculate hours from "9:00 - 6:00"
# ------------------------------------
def calculate_hours(cell_value: str) -> float:
    if not cell_value:
        return 0
    if "PTO" in cell_value:
        return 8
    
    try:
        time_line = cell_value.split("\n")[0]  # "09:00 - 6:00pm"
        start_str, end_str = [t.strip() for t in time_line.split("-")]

        def to_minutes_12hr(t):
            time_part, period = t.split()
            hour, minute = map(int, time_part.split(":"))
            if period == "PM" and hour != 12:
                hour += 12
            if period == "AM" and hour == 12:
                hour = 0

            return hour * 60 + minute
        start = to_minutes_12hr(start_str)
        end = to_minutes_12hr(end_str)

        if end < start:
            return 0  # prevents overnight shifts for now
        
        return (end - start) / 60.0
    except Exception:
        return 0
# ------------------------------------
# MAIN SCHEDULE PAGE
# ------------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    employees = load_employees()
    saved = load_last_schedule()

    if request.method == "POST":
        from datetime import datetime, timedelta
        week_start = request.form.get("week_start")

        start_date = datetime.strptime(week_start, "%Y-%m-%d")
        end_date = start_date + timedelta(days=4)

        date_range = f"{start_date.strftime('%b %d, %Y')} - {end_date.strftime('%b %d, %Y')}"
        # Build schedule dict from form data
        schedule = {}

        for emp in employees:
            shifts_for_emp = []
            for day in DAYS:
                roles = request.form.getlist(f"{emp}_{day}_role")
                roles = [r.strip() for r in roles if r.strip()]
                start = request.form.get(f"{emp}_{day}_start", "").strip()
                end = request.form.get(f"{emp}_{day}_end", "").strip()

                role_text = ", ".join(roles)
                if "PTO" in roles:
                    cell_text = "PTO"
                elif start and end:
                    cell_text = f"{start} - {end}\n{role_text}" if role_text else f"{start} - {end}"
                elif roles:
                    cell_text = role_text
                else:
                    cell_text = ""
                    
                shifts_for_emp.append(cell_text)

            schedule[emp] = shifts_for_emp
            
        # Save schedule so it persists
        saved_data = {
            "week_start": week_start,
            "schedule": {}
        }

        for emp in employees:
            saved_data["schedule"][emp] = {}
            for day in DAYS:
                saved_data["schedule"][emp][day] = {
                    "role": [r for r in request.form.getlist(f"{emp}_{day}_role") if r.strip()],
                    "start": request.form.get(f"{emp}_{day}_start", ""),
                    "end": request.form.get(f"{emp}_{day}_end", "")
                }

        save_last_schedule(saved_data)

        total_hours = {
            emp: sum(calculate_hours(cell) for cell in cells)
            for emp, cells in schedule.items()
        }

        df = pd.DataFrame(schedule).T
        df.columns = DAYS
        df["Total Hours"] = df.index.map(total_hours)
        df = df[(df[DAYS] != "").any(axis=1)]

        file_name = f"Schedule_{start_date.strftime('%Y_%m_%d')}.xlsx"
        file_path = os.path.join(APP_DIR, file_name)

        df.to_excel(file_name, index_label="Team Member")

        wb = load_workbook(file_path)
        ws = wb.active
        num_cols = ws.max_column

        ws.insert_rows(1)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)

        ws.column_dimensions["A"].width = 18
        for col in ["B", "C", "D", "E", "F"]:
            ws.column_dimensions[col].width = 26
        ws.column_dimensions["G"].width = 14

        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 22
        for row_idx in range(3, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 45

        title_cell = ws.cell(row=1, column=1)
        title_cell.value = f"Weekly Schedule  ·  {date_range}"
        title_cell.font = Font(bold=True, size=14, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="1E3A5F", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")

        day_colors = {
            2: {"header": "2563EB", "light": "EFF6FF", "band": "DBEAFE"},  # Monday    - blue
            3: {"header": "7C3AED", "light": "F5F3FF", "band": "EDE9FE"},  # Tuesday   - violet
            4: {"header": "0D9488", "light": "F0FDFA", "band": "CCFBF1"},  # Wednesday - teal
            5: {"header": "EA580C", "light": "FFF7ED", "band": "FFEDD5"},  # Thursday  - orange
            6: {"header": "DB2777", "light": "FDF2F8", "band": "FCE7F3"},  # Friday    - pink
        }
        total_header = "16A34A"
        total_bg     = "DCFCE7"
        total_font   = "15803D"

        for col_idx, cell in enumerate(ws[2], start=1):
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if col_idx in day_colors:
                cell.fill = PatternFill(start_color=day_colors[col_idx]["header"], fill_type="solid")
            elif col_idx == num_cols:
                cell.fill = PatternFill(start_color=total_header, fill_type="solid")
            else:
                cell.fill = PatternFill(start_color="1E3A5F", fill_type="solid")

        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        for band_idx, row in enumerate(ws.iter_rows(min_row=3)):
            is_band = band_idx % 2 == 1
            for col_idx, cell in enumerate(row, start=1):
                val = str(cell.value or "").strip()
                cell.border = thin_border
                cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

                if col_idx == 1:  # Team Member
                    cell.font = Font(bold=True, size=10)
                    if is_band:
                        cell.fill = PatternFill(start_color="F8FAFC", fill_type="solid")
                elif col_idx == num_cols:  # Total Hours
                    cell.font = Font(bold=True, size=10, color=total_font)
                    cell.fill = PatternFill(start_color=total_bg, fill_type="solid")
                elif col_idx in day_colors:  # Day columns
                    dc = day_colors[col_idx]
                    if val == "PTO":
                        cell.font = Font(bold=True, size=10, color="92400E")
                        cell.fill = PatternFill(start_color="FEF3C7", fill_type="solid")
                    elif not val:
                        cell.font = Font(size=10)
                        cell.fill = PatternFill(start_color="F1F5F9", fill_type="solid")
                    else:
                        cell.font = Font(size=10)
                        cell.fill = PatternFill(start_color=dc["band"] if is_band else dc["light"], fill_type="solid")

        ws.freeze_panes = "A3"

        wb.save(file_path)
        return send_file(file_path, as_attachment=True, download_name=file_name)

    return render_template(
        "schedule_form.html",
        employees=employees,
        days=DAYS,
        saved=saved
    )

@app.route("/export-pdf", methods=["POST"])
def export_pdf():
    employees = load_employees()
    from datetime import datetime, timedelta
    week_start = request.form.get("week_start")
    # Build Schedule from Current form data
    saved_data = {
        "week_start": week_start,
        "schedule": {}
    }
    for emp in employees:
        saved_data["schedule"][emp] = {}
        for day in DAYS:
            saved_data["schedule"][emp][day] = {
                "role": request.form.getlist(f"{emp}_{day}_role"),
                "start": request.form.get(f"{emp}_{day}_start", ""),
                "end": request.form.get(f"{emp}_{day}_end", "")
            }
    save_last_schedule(saved_data)
    saved = saved_data
    
    date_range_text = ""
    if week_start:
        start_date = datetime.strptime(week_start, "%Y-%m-%d")
        end_date = start_date + timedelta(days=4)
        date_range_text = f"{start_date.strftime('%b %d, %Y')} - {end_date.strftime('%b %d, %Y')}"

    
    if not saved or "schedule" not in saved:
        return "No schedule data to export.", 400
    
    file_name = "Weekly_Schedule.pdf"
    file_path = os.path.join(APP_DIR, file_name)

    doc = SimpleDocTemplate(
        file_path,
        pagesize=LETTER,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36,
    )

    styles = getSampleStyleSheet()
    elements = []

   

    # Title
    elements.append(Paragraph("<b>Weekly Schedule</b>", styles["Title"]))
    if date_range_text:
        elements.append(Paragraph(date_range_text, styles["Normal"]))
    elements.append(Spacer(1, 16))
    

    cell_style = ParagraphStyle('cell', fontSize=8, leading=11)
    pto_style  = ParagraphStyle('pto',  fontSize=8, leading=11,
                                textColor=colors.HexColor('#92400E'), fontName='Helvetica-Bold')

    # Table data — page is 612pt wide, 36pt margins each side = 540pt available
    # col widths: 80 + 81*5 + 55 = 540
    table_data = [["Employee"] + DAYS + ["Total Hours"]]
    for emp, day_data in saved["schedule"].items():
        row = [emp]
        total_hours = 0
        for d in DAYS:
            cell = day_data[d]
            roles = [r for r in (cell["role"] if isinstance(cell["role"], list) else []) if r.strip()]

            if "PTO" in roles:
                row.append(Paragraph("PTO", pto_style))
                total_hours += 8
            elif cell["start"] and cell["end"]:
                role_text = ", ".join(roles)
                para_text = f"{cell['start']} - {cell['end']}"
                if role_text:
                    para_text += f"<br/>{role_text}"
                row.append(Paragraph(para_text, cell_style))
                total_hours += calculate_hours(f"{cell['start']} - {cell['end']}")
            else:
                row.append("")
        row.append(f"{total_hours:.1f}")
        table_data.append(row)
    table_data = [row for row in table_data if any(cell for cell in row)]

    col_widths = [80] + [81] * len(DAYS) + [55]

    table = Table(
        table_data,
        colWidths=col_widths,
        repeatRows=1
    )

    table.setStyle(TableStyle([
        ("GRID",       (0, 0), (-1, -1), 1, colors.black),
        ("BACKGROUND", (0, 0), (-1,  0), colors.lightgrey),
        ("FONT",       (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONT",       (0, 1), ( 0, -1), "Helvetica-Bold"),
        ("VALIGN",     (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
    ]))
   
    elements.append(table)

    doc.build(elements)

    return send_file(file_path, as_attachment=True, download_name=file_name)
  
                                                              
    
# ------------------------------------
# PRINT PREVIEW
# ------------------------------------
@app.route("/print-preview", methods=["POST"])
def print_preview():
    from datetime import datetime, timedelta
    employees = load_employees()
    week_start = request.form.get("week_start")

    saved_data = {
        "week_start": week_start,
        "schedule": {}
    }
    for emp in employees:
        saved_data["schedule"][emp] = {}
        for day in DAYS:
            saved_data["schedule"][emp][day] = {
                "role": request.form.getlist(f"{emp}_{day}_role"),
                "start": request.form.get(f"{emp}_{day}_start", ""),
                "end": request.form.get(f"{emp}_{day}_end", "")
            }
    save_last_schedule(saved_data)

    totals = {}
    for emp in employees:
        total = 0.0
        for day in DAYS:
            cell = saved_data["schedule"][emp][day]
            roles = cell["role"]
            if "PTO" in roles:
                total += 8.0
            elif cell["start"] and cell["end"]:
                total += calculate_hours(f"{cell['start']} - {cell['end']}")
        totals[emp] = total

    role_color_palette = [
        "#2563EB", "#7C3AED", "#0D9488", "#EA580C", "#DB2777",
        "#16A34A", "#D97706", "#DC2626", "#0891B2", "#7C2D12",
    ]
    role_colors = {}
    for emp in employees:
        for day in DAYS:
            for role in saved_data["schedule"][emp][day]["role"]:
                if role and role != "PTO" and role not in role_colors:
                    role_colors[role] = role_color_palette[len(role_colors) % len(role_color_palette)]

    return render_template(
        "print_preview.html",
        employees=employees,
        days=DAYS,
        saved=saved_data,
        totals=totals,
        role_colors=role_colors,
    )


# ------------------------------------
# EMPLOYEE MANAGEMENT PAGE
# ------------------------------------
@app.route("/employees", methods=["GET", "POST"])
def manage_employees():
    employees = load_employees()

    if request.method == "POST":
        # Add new employee
        new_emp = request.form.get("new_employee", "").strip()
        if new_emp and new_emp not in employees:
            employees.append(new_emp)

        # Delete selected employees
        to_delete = request.form.getlist("delete_emp")
        employees = [e for e in employees if e not in to_delete]

        save_employees(employees)
        return redirect(url_for("manage_employees"))

    return render_template("employees.html", employees=employees)


if __name__ == "__main__":
    webbrowser.open("http://127.0.0.1:5000")
    app.run(debug=False)
