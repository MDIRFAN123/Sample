from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, date, timedelta
import os
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
MASTER_FILE = "ORB_Employee_Master.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ============================================================
# HELPERS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    return str(value).strip()


def normalize_status(value):
    """
    Converts different OOO formats into a standard value.
    """
    text = normalize_text(value).lower()

    text = text.replace("-", " ")
    text = text.replace("_", " ")
    text = re.sub(r"\s+", " ", text)

    if text in [
        "ooo",
        "out of office",
        "outofoffice",
        "out of the office",
        "away"
    ]:
        return "OOO"

    return text.upper()


def normalize_name(value):
    """
    Used to match employee names safely.
    """
    text = normalize_text(value).lower()
    text = re.sub(r"\s+", " ", text)
    return text


def parse_date(value):
    """
    Handles Excel dates and common text date formats.
    """

    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    # Excel serial number
    if isinstance(value, (int, float)):
        try:
            from openpyxl.utils.datetime import from_excel
            result = from_excel(value)
            if isinstance(result, datetime):
                return result.date()
            return result
        except Exception:
            pass

    text = normalize_text(value)

    if not text:
        return None

    formats = [
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m-%d-%Y",
        "%m-%d-%y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%m/%d",
    ]

    for fmt in formats:
        try:
            parsed = datetime.strptime(text, fmt)

            # If only month/day was supplied, use current year
            if fmt == "%m/%d":
                parsed = parsed.replace(year=datetime.now().year)

            return parsed.date()

        except ValueError:
            continue

    return None


# ============================================================
# MASTER EMPLOYEE FILE
# ============================================================

def load_master_employees():
    """
    Reads:

    Employee Name | Employee Code

    Example:

    Velina   | VR
    Cristen  | CR
    Tracy    | TR
    Shakima  | SL
    Holly    | HB
    """

    master_path = os.path.join(os.path.dirname(__file__), MASTER_FILE)

    if not os.path.exists(master_path):
        raise FileNotFoundError(
            f"Master file '{MASTER_FILE}' was not found. "
            f"Place it in the same folder as app.py."
        )

    wb = load_workbook(master_path, data_only=True)

    ws = wb.active

    employees = []

    headers = {}

    for cell in ws[1]:
        if cell.value is not None:
            headers[normalize_text(cell.value).lower()] = cell.column

    name_col = None
    code_col = None

    for key, col in headers.items():

        if key in [
            "employee name",
            "employee",
            "name"
        ]:
            name_col = col

        if key in [
            "employee code",
            "code",
            "emp code",
            "employee id"
        ]:
            code_col = col

    if name_col is None or code_col is None:
        raise ValueError(
            "Master file must contain 'Employee Name' and 'Employee Code' columns."
        )

    for row in range(2, ws.max_row + 1):

        name = normalize_text(ws.cell(row, name_col).value)
        code = normalize_text(ws.cell(row, code_col).value)

        if name and code:

            employees.append({
                "name": name,
                "code": code,
                "name_key": normalize_name(name)
            })

    if not employees:
        raise ValueError("No employees found in the master file.")

    return employees


# ============================================================
# INPUT OOO FILE
# ============================================================

def load_ooo_data(input_file):
    """
    Input file MUST contain only:

    Date | Status | Employee Name

    Example:

    9/25/2026 | Out of Office | Cristen
    9/24/2026 | Out of Office | Tracy
    """

    wb = load_workbook(input_file, data_only=True)

    ws = wb.active

    headers = {}

    for cell in ws[1]:

        if cell.value is not None:
            headers[normalize_text(cell.value).lower()] = cell.column

    date_col = None
    status_col = None
    name_col = None

    for key, col in headers.items():

        if key == "date":
            date_col = col

        elif key == "status":
            status_col = col

        elif key in [
            "employee name",
            "employee",
            "name"
        ]:
            name_col = col

    if date_col is None:
        raise ValueError("Input file must contain a 'Date' column.")

    if status_col is None:
        raise ValueError("Input file must contain a 'Status' column.")

    if name_col is None:
        raise ValueError("Input file must contain an 'Employee Name' column.")

    records = []

    for row in range(2, ws.max_row + 1):

        raw_date = ws.cell(row, date_col).value
        raw_status = ws.cell(row, status_col).value
        raw_name = ws.cell(row, name_col).value

        if not raw_date or not raw_name:
            continue

        parsed_date = parse_date(raw_date)

        if parsed_date is None:
            continue

        records.append({
            "date": parsed_date,
            "status": normalize_status(raw_status),
            "name": normalize_text(raw_name),
            "name_key": normalize_name(raw_name)
        })

    if not records:
        raise ValueError(
            "No valid records were found in the uploaded file."
        )

    return records


# ============================================================
# DETERMINE WEEK
# ============================================================

def get_week_dates(records):
    """
    Finds the week represented by the uploaded data.

    If input contains 9/25/2026,
    week becomes:

    Monday 9/21
    Tuesday 9/22
    Wednesday 9/23
    Thursday 9/24
    Friday 9/25
    """

    dates = [
        r["date"]
        for r in records
        if r.get("date") is not None
    ]

    if not dates:
        raise ValueError("No valid dates found.")

    # Use earliest date in uploaded file
    reference_date = min(dates)

    monday = reference_date - timedelta(
        days=reference_date.weekday()
    )

    friday = monday + timedelta(days=4)

    return [
        monday + timedelta(days=i)
        for i in range(5)
    ]


# ============================================================
# BUILD DAILY OOO MAP
# ============================================================

def build_ooo_map(records, employees):
    """
    Creates:

    {
        date: {
            employee_name: employee_code
        }
    }
    """

    employee_lookup = {
        employee["name_key"]: employee["code"]
        for employee in employees
    }

    ooo_map = {}

    for record in records:

        if record["status"] != "OOO":
            continue

        date_value = record["date"]
        name_key = record["name_key"]

        # Only recognize employees from master file
        if name_key not in employee_lookup:
            continue

        code = employee_lookup[name_key]

        if date_value not in ooo_map:
            ooo_map[date_value] = []

        if code not in ooo_map[date_value]:
            ooo_map[date_value].append(code)

    return ooo_map


# ============================================================
# CREATE EXCEL REPORT
# ============================================================

def create_report(records):

    employees = load_master_employees()

    week_dates = get_week_dates(records)

    ooo_map = build_ooo_map(records, employees)

    # --------------------------------------------------------
    # CREATE WORKBOOK
    # --------------------------------------------------------

    wb = Workbook()
    ws = wb.active
    ws.title = "ORB Schedule"

    # --------------------------------------------------------
    # COLORS
    # --------------------------------------------------------

    BLACK = "000000"
    WHITE = "FFFFFF"

    ONLINE_BLUE = "00A6D6"
    OOO_YELLOW = "FFC000"

    LIGHT_GREY = "D9E1F2"

    RED = "C00000"

    # --------------------------------------------------------
    # FONTS
    # --------------------------------------------------------

    header_font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color=WHITE
    )

    normal_font = Font(
        name="Calibri",
        size=11
    )

    online_font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color=WHITE
    )

    ooo_label_font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color=RED
    )

    ooo_code_font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color=RED
    )

    # --------------------------------------------------------
    # FILLS
    # --------------------------------------------------------

    black_fill = PatternFill(
        fill_type="solid",
        fgColor=BLACK
    )

    online_fill = PatternFill(
        fill_type="solid",
        fgColor=ONLINE_BLUE
    )

    ooo_fill = PatternFill(
        fill_type="solid",
        fgColor=OOO_YELLOW
    )

    grey_fill = PatternFill(
        fill_type="solid",
        fgColor=LIGHT_GREY
    )

    # --------------------------------------------------------
    # BORDER
    # --------------------------------------------------------

    thin_side = Side(
        style="thin",
        color="808080"
    )

    thin_border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side
    )

    # --------------------------------------------------------
    # COLUMN WIDTHS
    # --------------------------------------------------------

    ws.column_dimensions["A"].width = 24

    for col in range(2, 7):
        ws.column_dimensions[
            get_column_letter(col)
        ].width = 18

    # --------------------------------------------------------
    # ROW HEIGHTS
    # --------------------------------------------------------

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 42
    ws.row_dimensions[4].height = 30

    # --------------------------------------------------------
    # WEEKDAY HEADER
    # --------------------------------------------------------

    weekdays = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday"
    ]

    for index, day_name in enumerate(weekdays, start=2):

        cell = ws.cell(
            row=1,
            column=index
        )

        cell.value = day_name
        cell.fill = black_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # --------------------------------------------------------
    # DATE ROW
    # --------------------------------------------------------

    for index, day_date in enumerate(
        week_dates,
        start=2
    ):

        cell = ws.cell(
            row=2,
            column=index
        )

        # Excel date
        cell.value = day_date

        cell.number_format = "m/d"

        cell.fill = black_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

    # --------------------------------------------------------
    # LABELS
    # --------------------------------------------------------

    ws["A3"] = "Online"
    ws["A4"] = "Out of Office"

    ws["A3"].fill = online_fill
    ws["A3"].font = online_font

    ws["A4"].fill = ooo_fill
    ws["A4"].font = ooo_label_font

    ws["A3"].alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    ws["A4"].alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    # --------------------------------------------------------
    # DAILY ONLINE / OOO
    # --------------------------------------------------------

    all_codes = [
        employee["code"]
        for employee in employees
    ]

    for index, day_date in enumerate(
        week_dates,
        start=2
    ):

        # --------------------------------------------
        # OOO CODES
        # --------------------------------------------

        ooo_codes = ooo_map.get(
            day_date,
            []
        )

        # --------------------------------------------
        # ONLINE CODES
        # --------------------------------------------

        online_codes = [
            code
            for code in all_codes
            if code not in ooo_codes
        ]

        # --------------------------------------------
        # ONLINE CELL
        # --------------------------------------------

        online_cell = ws.cell(
            row=3,
            column=index
        )

        online_cell.value = "/".join(
            online_codes
        )

        online_cell.fill = grey_fill
        online_cell.font = normal_font
        online_cell.border = thin_border

        online_cell.alignment = Alignment(
            horizontal="left",
            vertical="top",
            wrap_text=True
        )

        # --------------------------------------------
        # OOO CELL
        # --------------------------------------------

        ooo_cell = ws.cell(
            row=4,
            column=index
        )

        ooo_cell.value = "/".join(
            ooo_codes
        )

        ooo_cell.fill = grey_fill
        ooo_cell.border = thin_border

        ooo_cell.font = ooo_code_font

        ooo_cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True
        )

        # Red font ONLY if OOO exists
        if ooo_codes:
            ooo_cell.font = ooo_code_font
        else:
            ooo_cell.font = normal_font

    # --------------------------------------------------------
    # BORDERS FOR HEADER
    # --------------------------------------------------------

    for row in range(1, 5):

        for col in range(2, 7):

            ws.cell(
                row=row,
                column=col
            ).border = thin_border

    # --------------------------------------------------------
    # WORK QUEUE SECTION
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
            size=11
        )

        ws.cell(
            row=row,
            column=1
        ).alignment = Alignment(
            horizontal="left"
        )

    # --------------------------------------------------------
    # FREEZE
    # --------------------------------------------------------

    ws.freeze_panes = "B3"

    # --------------------------------------------------------
    # PAGE SETUP
    # --------------------------------------------------------

    ws.sheet_view.showGridLines = True

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # --------------------------------------------------------
    # OUTPUT
    # --------------------------------------------------------

    start_date = week_dates[0]
    end_date = week_dates[-1]

    filename = (
        f"ORB_OOO_Schedule_"
        f"{start_date.strftime('%Y%m%d')}_"
        f"{end_date.strftime('%Y%m%d')}.xlsx"
    )

    output_path = os.path.join(
        OUTPUT_FOLDER,
        filename
    )

    wb.save(output_path)

    return output_path


# ============================================================
# FLASK ROUTES
# ============================================================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )


@app.route(
    "/generate",
    methods=["POST"]
)
def generate():

    try:

        if "file" not in request.files:

            return render_template(
                "index.html",
                error="Please upload the OOO Excel file."
            )

        uploaded_file = request.files["file"]

        if not uploaded_file.filename:

            return render_template(
                "index.html",
                error="Please select an Excel file."
            )

        # Save uploaded file
        input_path = os.path.join(
            UPLOAD_FOLDER,
            uploaded_file.filename
        )

        uploaded_file.save(
            input_path
        )

        # Read input
        records = load_ooo_data(
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

        print("ERROR:", str(e))

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
