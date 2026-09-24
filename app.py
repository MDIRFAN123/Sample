from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from copy import copy
from datetime import datetime, timedelta
import pandas as pd
import os
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
MASTER_FILE = "ORB_Employees.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ============================================================
# HELPERS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    return str(value).strip().lower()


def normalize_status(value):
    value = normalize_text(value)

    value = value.replace("-", " ")
    value = re.sub(r"\s+", " ", value)

    if value in [
        "ooo",
        "out of office",
        "out-of-office",
        "outoffice",
        "out_of_office"
    ]:
        return "OOO"

    return value.upper()


def clean_name(value):
    if value is None:
        return ""

    return re.sub(r"\s+", " ", str(value).strip())


def parse_date(value):
    """
    Handles Excel dates and common string date formats.
    """

    if pd.isna(value):
        return None

    if isinstance(value, datetime):
        return value.date()

    # pandas Timestamp
    if isinstance(value, pd.Timestamp):
        return value.date()

    # Python date
    try:
        if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
            return value
    except Exception:
        pass

    text = str(value).strip()

    if not text:
        return None

    formats = [
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m-%d-%Y",
        "%m-%d-%y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    # Last attempt using pandas
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if not pd.isna(parsed):
            return parsed.date()
    except Exception:
        pass

    return None


# ============================================================
# LOAD MASTER EMPLOYEE FILE
# ============================================================

def load_master_employees():

    if not os.path.exists(MASTER_FILE):
        raise FileNotFoundError(
            f"Master employee file '{MASTER_FILE}' was not found."
        )

    df = pd.read_excel(MASTER_FILE)

    # Normalize column names
    df.columns = [
        str(col).strip().lower()
        for col in df.columns
    ]

    name_col = None
    code_col = None

    for col in df.columns:

        if col in [
            "employee name",
            "employee_name",
            "name",
            "employee"
        ]:
            name_col = col

        if col in [
            "employee code",
            "employee_code",
            "code",
            "emp code",
            "emp_code"
        ]:
            code_col = col

    if not name_col or not code_col:
        raise ValueError(
            "Master file must contain 'Employee Name' and 'Employee Code' columns."
        )

    employees = {}

    for _, row in df.iterrows():

        name = clean_name(row[name_col])
        code = clean_name(row[code_col])

        if not name or not code:
            continue

        employees[normalize_text(name)] = {
            "name": name,
            "code": code
        }

    if not employees:
        raise ValueError("No employees found in master file.")

    return employees


# ============================================================
# READ OOO INPUT
# ============================================================

def read_ooo_file(filepath):

    df = pd.read_excel(filepath)

    # Clean column names
    df.columns = [
        str(col).strip().lower()
        for col in df.columns
    ]

    date_col = None
    status_col = None
    employee_col = None

    for col in df.columns:

        if col in [
            "date",
            "ooo date",
            "out of office date"
        ]:
            date_col = col

        if col in [
            "status",
            "ooo status"
        ]:
            status_col = col

        if col in [
            "employee name",
            "employee",
            "name"
        ]:
            employee_col = col

    missing = []

    if not date_col:
        missing.append("Date")

    if not status_col:
        missing.append("Status")

    if not employee_col:
        missing.append("Employee Name")

    if missing:
        raise ValueError(
            "Input file is missing: " + ", ".join(missing)
        )

    records = []

    for _, row in df.iterrows():

        date_value = parse_date(row[date_col])
        status = normalize_status(row[status_col])
        employee_name = clean_name(row[employee_col])

        if not date_value:
            continue

        if not employee_name:
            continue

        # Only OOO records matter
        if status != "OOO":
            continue

        records.append({
            "date": date_value,
            "employee": employee_name
        })

    if not records:
        raise ValueError(
            "No valid OOO records were found in the input file."
        )

    return records


# ============================================================
# FIND WEEK
# ============================================================

def get_week_dates(records):

    dates = [
        r["date"]
        for r in records
        if r.get("date")
    ]

    if not dates:
        raise ValueError("No valid dates found.")

    # Use the earliest date supplied in the input.
    anchor_date = min(dates)

    # Monday = 0
    monday = anchor_date - timedelta(
        days=anchor_date.weekday()
    )

    week_dates = [
        monday + timedelta(days=i)
        for i in range(5)
    ]

    return week_dates


# ============================================================
# CREATE OOO LOOKUP
# ============================================================

def build_ooo_lookup(records, employees):

    ooo_lookup = {}

    for record in records:

        date_value = record["date"]
        employee_name = clean_name(record["employee"])

        employee_key = normalize_text(employee_name)

        if employee_key not in employees:
            # Employee exists in input but not master.
            # Ignore it rather than breaking report.
            continue

        employee_code = employees[employee_key]["code"]

        if date_value not in ooo_lookup:
            ooo_lookup[date_value] = []

        if employee_code not in ooo_lookup[date_value]:
            ooo_lookup[date_value].append(employee_code)

    return ooo_lookup


# ============================================================
# CREATE EXCEL REPORT
# ============================================================

def create_report(records):

    employees = load_master_employees()

    week_dates = get_week_dates(records)

    ooo_lookup = build_ooo_lookup(
        records,
        employees
    )

    # --------------------------------------------------------
    # Workbook
    # --------------------------------------------------------

    wb = load_workbook(
        filename=None
    ) if False else None

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "ORB Schedule"

    # --------------------------------------------------------
    # Colors
    # --------------------------------------------------------

    BLACK = "000000"
    WHITE = "FFFFFF"

    SYNCHRONY_BLUE = "00AEEF"
    OOO_YELLOW = "FFC000"

    LIGHT_GREY = "E7E6E6"

    RED = "C00000"

    thin_gray = Side(
        style="thin",
        color="808080"
    )

    border = Border(
        left=thin_gray,
        right=thin_gray,
        top=thin_gray,
        bottom=thin_gray
    )

    # --------------------------------------------------------
    # Layout
    #
    # A = Labels
    # B-F = Monday-Friday
    # --------------------------------------------------------

    ws.column_dimensions["A"].width = 24

    for col in range(2, 7):
        ws.column_dimensions[
            get_column_letter(col)
        ].width = 18

    # --------------------------------------------------------
    # Day Header
    # --------------------------------------------------------

    day_names = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday"
    ]

    for index, day_name in enumerate(day_names, start=2):

        cell = ws.cell(
            row=1,
            column=index
        )

        cell.value = day_name

        cell.fill = PatternFill(
            "solid",
            fgColor=BLACK
        )

        cell.font = Font(
            name="Calibri",
            size=12,
            bold=True,
            color=WHITE
        )

        cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

    # --------------------------------------------------------
    # Dates
    # --------------------------------------------------------

    for index, date_value in enumerate(
        week_dates,
        start=2
    ):

        cell = ws.cell(
            row=2,
            column=index
        )

        # Excel date
        cell.value = date_value

        cell.number_format = "m/d"

        cell.fill = PatternFill(
            "solid",
            fgColor=BLACK
        )

        cell.font = Font(
            name="Calibri",
            size=12,
            bold=True,
            color=WHITE
        )

        cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

    # --------------------------------------------------------
    # Row Labels
    # --------------------------------------------------------

    ws["A3"] = "Online"
    ws["A4"] = "Out of Office"

    ws["A3"].fill = PatternFill(
        "solid",
        fgColor=SYNCHRONY_BLUE
    )

    ws["A4"].fill = PatternFill(
        "solid",
        fgColor=OOO_YELLOW
    )

    ws["A3"].font = Font(
        name="Calibri",
        size=12,
        bold=True,
        color=WHITE
    )

    ws["A4"].font = Font(
        name="Calibri",
        size=12,
        bold=True,
        color=RED
    )

    ws["A3"].alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    ws["A4"].alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    # --------------------------------------------------------
    # IMPORTANT LOGIC
    #
    # Online = ALL employees - OOO employees
    # --------------------------------------------------------

    all_employee_codes = [
        employee["code"]
        for employee in employees.values()
    ]

    # Keep master order
    all_employee_codes = list(
        dict.fromkeys(all_employee_codes)
    )

    for col_index, date_value in enumerate(
        week_dates,
        start=2
    ):

        # Employees OOO on this date
        ooo_codes = ooo_lookup.get(
            date_value,
            []
        )

        # Online = everyone NOT OOO
        online_codes = [
            code
            for code in all_employee_codes
            if code not in ooo_codes
        ]

        online_cell = ws.cell(
            row=3,
            column=col_index
        )

        ooo_cell = ws.cell(
            row=4,
            column=col_index
        )

        # ----------------------------------------------------
        # ONLINE
        # ----------------------------------------------------

        online_cell.value = "/".join(
            online_codes
        )

        # ----------------------------------------------------
        # OOO
        # ----------------------------------------------------

        ooo_cell.value = "/".join(
            ooo_codes
        )

        # ----------------------------------------------------
        # Formatting
        # ----------------------------------------------------

        for cell in [
            online_cell,
            ooo_cell
        ]:

            cell.fill = PatternFill(
                "solid",
                fgColor=LIGHT_GREY
            )

            cell.border = border

            cell.font = Font(
                name="Calibri",
                size=11,
                color="000000"
            )

            cell.alignment = Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=True
            )

        # OOO text in red
        if ooo_codes:

            ooo_cell.font = Font(
                name="Calibri",
                size=11,
                bold=True,
                color=RED
            )

    # --------------------------------------------------------
    # Borders for date/header area
    # --------------------------------------------------------

    for row in range(1, 5):

        for col in range(2, 7):

            ws.cell(
                row=row,
                column=col
            ).border = border

    # --------------------------------------------------------
    # Operational Rows
    # --------------------------------------------------------

    ws["A5"] = "TG2/TG3 Agenda"
    ws["A6"] = "Rewards IT Queue"
    ws["A7"] = "Remediations"

    for row in range(5, 8):

        ws.cell(
            row=row,
            column=1
        ).font = Font(
            name="Calibri",
            size=12
        )

        ws.cell(
            row=row,
            column=1
        ).alignment = Alignment(
            horizontal="left"
        )

    # --------------------------------------------------------
    # Example assignments can be configured here
    # --------------------------------------------------------
    #
    # If these are fixed responsibilities, uncomment/update.
    #
    # ws["B5"] = "Velina"
    # ws["B6"] = "Cristen"
    # ws["B7"] = "Tracy/Shakima"
    #
    # --------------------------------------------------------

    # Row heights
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 35
    ws.row_dimensions[4].height = 35

    # Freeze panes
    ws.freeze_panes = "B3"

    # --------------------------------------------------------
    # Output
    # --------------------------------------------------------

    first_date = week_dates[0]
    last_date = week_dates[-1]

    filename = (
        f"ORB_OOO_Schedule_"
        f"{first_date.strftime('%Y%m%d')}_"
        f"{last_date.strftime('%Y%m%d')}.xlsx"
    )

    output_path = os.path.join(
        OUTPUT_FOLDER,
        filename
    )

    wb.save(output_path)

    return output_path


# ============================================================
# ROUTES
# ============================================================

@app.route("/", methods=["GET"])
def index():

    return render_template(
        "index.html"
    )


@app.route("/generate", methods=["POST"])
def generate():

    try:

        if "file" not in request.files:

            return render_template(
                "index.html",
                error="Please upload the OOO input file."
            )

        uploaded_file = request.files["file"]

        if uploaded_file.filename == "":

            return render_template(
                "index.html",
                error="Please select an Excel file."
            )

        # Save uploaded file
        input_path = os.path.join(
            UPLOAD_FOLDER,
            uploaded_file.filename
        )

        uploaded_file.save(input_path)

        # Read OOO records
        records = read_ooo_file(
            input_path
        )

        # Generate report
        output_path = create_report(
            records
        )

        # Automatically download
        return send_file(
            output_path,
            as_attachment=True,
            download_name=os.path.basename(
                output_path
            ),
            mimetype=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )

    except Exception as e:

        return render_template(
            "index.html",
            error=str(e)
        )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
