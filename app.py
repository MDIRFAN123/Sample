from flask import Flask, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, date, timedelta
import os
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ============================================================
# EMPLOYEE MASTER - HARD CODED
# ============================================================

EMPLOYEE_CODES = {
    "Velina": "VR",
    "Cristen": "CR",
    "Tracy": "TR",
    "Shakima": "SL",
    "Holly": "HB"
}


# ============================================================
# HELPERS
# ============================================================

def normalize_name(value):
    if value is None:
        return ""

    value = str(value).strip()

    value = re.sub(r"\s+", " ", value)

    return value.lower()


def clean_name(value):
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value).strip()
    )


def normalize_status(value):

    if value is None:
        return ""

    value = str(value).strip().lower()

    value = value.replace("-", " ")
    value = value.replace("_", " ")

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    if value in [
        "ooo",
        "out of office",
        "outof office",
        "outoffice",
        "out of the office"
    ]:
        return "OOO"

    return value.upper()


# ============================================================
# DATE PARSER
# ============================================================

def parse_date(value):

    if value is None:
        return None

    # Python datetime
    if isinstance(value, datetime):
        return value.date()

    # Python date
    if isinstance(value, date):
        return value

    # Excel date / number
    if isinstance(value, (int, float)):

        try:

            from openpyxl.utils.datetime import from_excel

            result = from_excel(value)

            if isinstance(result, datetime):
                return result.date()

            return result

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
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%d-%m-%Y"
    ]

    for fmt in formats:

        try:

            return datetime.strptime(
                text,
                fmt
            ).date()

        except ValueError:
            continue

    # Last attempt
    try:

        parsed = datetime.fromisoformat(
            text
        )

        return parsed.date()

    except Exception:
        pass

    return None


# ============================================================
# READ OOO EXCEL
# ============================================================

def read_ooo_file(filepath):

    from openpyxl import load_workbook

    wb = load_workbook(
        filepath,
        data_only=True
    )

    ws = wb.active

    # --------------------------------------------------------
    # Find columns
    # --------------------------------------------------------

    headers = {}

    for cell in ws[1]:

        if cell.value is not None:

            headers[
                str(cell.value).strip().lower()
            ] = cell.column

    date_col = None
    status_col = None
    employee_col = None

    for header, column in headers.items():

        if header in [
            "date",
            "ooo date",
            "out of office date"
        ]:
            date_col = column

        elif header in [
            "status",
            "ooo status"
        ]:
            status_col = column

        elif header in [
            "employee name",
            "employee",
            "name"
        ]:
            employee_col = column

    # --------------------------------------------------------
    # Validate columns
    # --------------------------------------------------------

    if date_col is None:

        raise ValueError(
            "Input Excel must contain a 'Date' column."
        )

    if status_col is None:

        raise ValueError(
            "Input Excel must contain a 'Status' column."
        )

    if employee_col is None:

        raise ValueError(
            "Input Excel must contain an 'Employee Name' column."
        )

    # --------------------------------------------------------
    # Read records
    # --------------------------------------------------------

    records = []

    for row in range(
        2,
        ws.max_row + 1
    ):

        raw_date = ws.cell(
            row,
            date_col
        ).value

        raw_status = ws.cell(
            row,
            status_col
        ).value

        raw_employee = ws.cell(
            row,
            employee_col
        ).value

        if raw_date is None:
            continue

        if raw_employee is None:
            continue

        parsed_date = parse_date(
            raw_date
        )

        if parsed_date is None:
            continue

        employee_name = clean_name(
            raw_employee
        )

        status = normalize_status(
            raw_status
        )

        records.append({
            "date": parsed_date,
            "status": status,
            "employee": employee_name,
            "employee_key": normalize_name(
                employee_name
            )
        })

    if not records:

        raise ValueError(
            "No valid records were found in the uploaded Excel file."
        )

    return records


# ============================================================
# GET WEEK
# ============================================================

def get_week_dates(records):

    valid_dates = [
        record["date"]
        for record in records
        if record["date"] is not None
    ]

    if not valid_dates:

        raise ValueError(
            "No valid dates found."
        )

    # Earliest date determines the week
    reference_date = min(
        valid_dates
    )

    # Monday
    monday = (
        reference_date -
        timedelta(
            days=reference_date.weekday()
        )
    )

    # Monday-Friday
    week_dates = [
        monday + timedelta(days=i)
        for i in range(5)
    ]

    return week_dates


# ============================================================
# GET EMPLOYEE CODE
# ============================================================

def get_employee_code(employee_name):

    employee_key = normalize_name(
        employee_name
    )

    # Exact match
    for name, code in EMPLOYEE_CODES.items():

        if normalize_name(name) == employee_key:

            return code

    # Partial match as backup
    for name, code in EMPLOYEE_CODES.items():

        master_key = normalize_name(name)

        if (
            employee_key in master_key
            or
            master_key in employee_key
        ):

            return code

    return None


# ============================================================
# GET OOO EMPLOYEES FOR DATE
# ============================================================

def get_ooo_codes(
    current_date,
    records
):

    ooo_codes = []

    for record in records:

        if record["date"] != current_date:
            continue

        if record["status"] != "OOO":
            continue

        employee_code = get_employee_code(
            record["employee"]
        )

        if employee_code is None:
            continue

        if employee_code not in ooo_codes:

            ooo_codes.append(
                employee_code
            )

    return ooo_codes


# ============================================================
# GET ONLINE EMPLOYEES FOR DATE
# ============================================================

def get_online_codes(
    current_date,
    records
):

    # Employees OOO today
    ooo_codes = get_ooo_codes(
        current_date,
        records
    )

    # ALL employees from hard-coded master
    all_codes = list(
        EMPLOYEE_CODES.values()
    )

    # Remove OOO employees
    online_codes = [
        code
        for code in all_codes
        if code not in ooo_codes
    ]

    return online_codes


# ============================================================
# CREATE EXCEL REPORT
# ============================================================

def create_report(records):

    # --------------------------------------------------------
    # Determine week
    # --------------------------------------------------------

    week_dates = get_week_dates(
        records
    )

    # --------------------------------------------------------
    # Create workbook
    # --------------------------------------------------------

    wb = Workbook()

    ws = wb.active

    ws.title = "ORB Schedule"

    # --------------------------------------------------------
    # COLORS
    # --------------------------------------------------------

    BLACK = "000000"
    WHITE = "FFFFFF"

    SYNCHRONY_BLUE = "00AEEF"

    OOO_YELLOW = "FFC000"

    LIGHT_GREY = "E7E6E6"

    RED = "C00000"

    # --------------------------------------------------------
    # STYLES
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

    black_fill = PatternFill(
        fill_type="solid",
        fgColor=BLACK
    )

    online_label_fill = PatternFill(
        fill_type="solid",
        fgColor=SYNCHRONY_BLUE
    )

    ooo_label_fill = PatternFill(
        fill_type="solid",
        fgColor=OOO_YELLOW
    )

    data_fill = PatternFill(
        fill_type="solid",
        fgColor=LIGHT_GREY
    )

    # --------------------------------------------------------
    # COLUMN WIDTH
    # --------------------------------------------------------

    ws.column_dimensions["A"].width = 25

    for column in range(2, 7):

        ws.column_dimensions[
            get_column_letter(column)
        ].width = 18

    # --------------------------------------------------------
    # ROW HEIGHT
    # --------------------------------------------------------

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 40
    ws.row_dimensions[4].height = 35

    # --------------------------------------------------------
    # DAY NAMES
    # --------------------------------------------------------

    day_names = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday"
    ]

    for index, day_name in enumerate(
        day_names,
        start=2
    ):

        cell = ws.cell(
            row=1,
            column=index
        )

        cell.value = day_name

        cell.fill = black_fill

        cell.font = Font(
            name="Calibri",
            size=11,
            bold=True,
            color=WHITE
        )

        cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        cell.border = thin_border

    # --------------------------------------------------------
    # DATES
    # --------------------------------------------------------

    for index, current_date in enumerate(
        week_dates,
        start=2
    ):

        cell = ws.cell(
            row=2,
            column=index
        )

        cell.value = current_date

        cell.number_format = "m/d"

        cell.fill = black_fill

        cell.font = Font(
            name="Calibri",
            size=11,
            bold=True,
            color=WHITE
        )

        cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        cell.border = thin_border

    # --------------------------------------------------------
    # LEFT LABELS
    # --------------------------------------------------------

    ws["A3"] = "Online"

    ws["A4"] = "Out of Office"

    ws["A3"].fill = online_label_fill

    ws["A4"].fill = ooo_label_fill

    ws["A3"].font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color=WHITE
    )

    ws["A4"].font = Font(
        name="Calibri",
        size=11,
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

    ws["A3"].border = thin_border
    ws["A4"].border = thin_border

    # --------------------------------------------------------
    # DAILY DATA
    # --------------------------------------------------------

    for index, current_date in enumerate(
        week_dates,
        start=2
    ):

        # Get OOO
        ooo_codes = get_ooo_codes(
            current_date,
            records
        )

        # Get Online
        online_codes = get_online_codes(
            current_date,
            records
        )

        # -----------------------------------------------
        # ONLINE
        # -----------------------------------------------

        online_cell = ws.cell(
            row=3,
            column=index
        )

        online_cell.value = "/".join(
            online_codes
        )

        online_cell.fill = data_fill

        online_cell.font = Font(
            name="Calibri",
            size=11
        )

        online_cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True
        )

        online_cell.border = thin_border

        # -----------------------------------------------
        # OOO
        # -----------------------------------------------

        ooo_cell = ws.cell(
            row=4,
            column=index
        )

        ooo_cell.value = "/".join(
            ooo_codes
        )

        ooo_cell.fill = data_fill

        if ooo_codes:

            ooo_cell.font = Font(
                name="Calibri",
                size=11,
                bold=True,
                color=RED
            )

        else:

            ooo_cell.font = Font(
                name="Calibri",
                size=11
            )

        ooo_cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True
        )

        ooo_cell.border = thin_border

    # --------------------------------------------------------
    # OTHER ROWS FROM YOUR FORMAT
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
            horizontal="left",
            vertical="center"
        )

    # --------------------------------------------------------
    # FREEZE
    # --------------------------------------------------------

    ws.freeze_panes = "B3"

    # --------------------------------------------------------
    # PAGE SETUP
    # --------------------------------------------------------

    ws.page_setup.orientation = "landscape"

    ws.page_setup.fitToWidth = 1

    ws.page_setup.fitToHeight = 1

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # --------------------------------------------------------
    # OUTPUT NAME
    # --------------------------------------------------------

    start_date = week_dates[0]

    end_date = week_dates[-1]

    filename = (
        "ORB_OOO_Schedule_"
        + start_date.strftime("%Y%m%d")
        + "_"
        + end_date.strftime("%Y%m%d")
        + ".xlsx"
    )

    output_path = os.path.join(
        OUTPUT_FOLDER,
        filename
    )

    wb.save(
        output_path
    )

    return output_path


# ============================================================
# HOME
# ============================================================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )


# ============================================================
# GENERATE REPORT
# ============================================================

@app.route(
    "/generate",
    methods=["POST"]
)
def generate():

    try:

        uploaded_file = request.files.get(
            "file"
        )

        if uploaded_file is None:

            return render_template(
                "index.html",
                error="Please upload an Excel file."
            )

        if uploaded_file.filename == "":

            return render_template(
                "index.html",
                error="Please select an Excel file."
            )

        # ----------------------------------------------------
        # Save uploaded file
        # ----------------------------------------------------

        input_filename = (
            uploaded_file.filename
        )

        input_path = os.path.join(
            UPLOAD_FOLDER,
            input_filename
        )

        uploaded_file.save(
            input_path
        )

        # ----------------------------------------------------
        # Read OOO
        # ----------------------------------------------------

        records = read_ooo_file(
            input_path
        )

        # ----------------------------------------------------
        # Generate report
        # ----------------------------------------------------

        output_path = create_report(
            records
        )

        # ----------------------------------------------------
        # Download automatically
        # ----------------------------------------------------

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

        print(
            "ERROR:",
            repr(e)
        )

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
