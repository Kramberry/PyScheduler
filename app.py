from flask import Flask, render_template, request, send_file, redirect, url_for, Response
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from reportlab.lib.pagesizes import LETTER
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
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
                role = request.form.get(f"{emp}_{day}_role", "").strip()
                start = request.form.get(f"{emp}_{day}_start", "").strip()
                end = request.form.get(f"{emp}_{day}_end", "").strip()

                if role == "PTO":
                    cell_text = "PTO"
                elif role and start and end:
                    cell_text = f"{start} - {end}\n{role}"
                elif role:
                    cell_text = role
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
                    "role": request.form.get(f"{emp}_{day}_role", ""),
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

        file_name = f"Schedule_{start_date.strftime('%Y_%m_%d')}.xlsx"
        file_path = os.path.join(APP_DIR, file_name)

        df.to_excel(file_name, index_label="Team Member")

        # Excel formatting (unchanged)
        wb = load_workbook(file_path)
        ws = wb.active
        ws.insert_rows(1)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
        # Column widths
        ws.column_dimensions["A"].width = 18   # Team Member
        ws.column_dimensions["B"].width = 28   # Monday
        ws.column_dimensions["C"].width = 28   # Tuesday
        ws.column_dimensions["D"].width = 28   # Wednesday
        ws.column_dimensions["E"].width = 28   # Thursday
        ws.column_dimensions["F"].width = 28   # Friday
        ws.column_dimensions["G"].width = 14   # Total Hours   

        # Increase row height for schedule rows
        for row_idx in range(3, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 45

        date_cell = ws.cell(row=1, column=1)
        date_cell.value = f"Weekly Schedule: {date_range}"
        date_cell.font = Font(bold=True, size=14)
        date_cell.alignment = Alignment(horizontal="center")

        header_fill = PatternFill(start_color="D9D9D9", fill_type="solid")
        for cell in ws[2]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ROLE_COLORS = {
            "Printer": "CCFFFF",
            "Sealer": "CCFFCC",
            "Shipper": "FFE5CC",
            "Production Coord.": "E6CCFF",
            "PTO": "FFCCCC",
        }

        for row in ws.iter_rows(min_row=3, min_col=2, max_col=6):
            for cell in row:
                if not cell.value:
                    continue
                for role, color in ROLE_COLORS.items():
                    if role in str(cell.value):
                        cell.fill = PatternFill(start_color=color, fill_type="solid")
                        break

        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(
                    wrap_text=True, horizontal="center", vertical="center"
                )

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
                "role": request.form.get(f"{emp}_{day}_role", ""),
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

    ROLE_COLORS = {
        "Printer": colors.HexColor("#CCFFFF"),
        "Sealer": colors.HexColor("#CCFFCC"),
        "Shipper": colors.HexColor("#FFE5CC"),
        "Production Coord.": colors.HexColor("#E6CCFF"),
        "PTO": colors.HexColor("#FFCCCC"),
    }

    # Title
    elements.append(Paragraph("<b>Weekly Schedule</b>", styles["Title"]))
    if date_range_text:
        elements.append(Paragraph(date_range_text, styles["Normal"]))
    elements.append(Spacer(1, 16))
    

    # Table data
    table_data = [["Employee"] + DAYS]
    for emp, day_data in saved["schedule"].items():
        row = [emp]
        for d in DAYS:
            cell = day_data[d]
            if cell["role"] == "PTO":
                row.append("PTO")
            else:
                row.append(f"{cell['start']} - {cell['end']}\n{cell['role']}" )
        table_data.append(row)

    col_widths = [90] + [85] * len(DAYS)

    table = Table(
        table_data,
        colWidths=col_widths,
        repeatRows=1
    )

    table.setStyle(TableStyle([
    # Grid & header
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONT", (0, 1), (0, -1), "Helvetica-Bold"),

    # Alignment
        ("ALIGN", (1, 1), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),

    # Padding (THIS is the big improvement)
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),

    # Font size for body
        ("FONTSIZE", (0, 1), (-1, -1), 9),

]))
    for row_idx, (emp,day_data) in enumerate(saved["schedule"].items(), start=1):
        for col_idx, day in enumerate(DAYS, start=1):
            role = day_data[day]["role"]
            for role_name, color in ROLE_COLORS.items():
                if role_name in role:
                    table.setStyle(TableStyle([
                        ("BACKGROUND", (col_idx, row_idx), (col_idx, row_idx), color)
                    ]))
                    break

    elements.append(table)

    doc.build(elements)

    return send_file(file_path, as_attachment=True, download_name=file_name)
  
                                                              
    
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
