import os
from datetime import datetime, date, timedelta
from copy import copy

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file
)

from openpyxl import load_workbook
from werkzeug.utils import secure_filename


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

EMPLOYEE_MASTER_FILE = os.path.join(
    BASE_DIR,
    "ORB_Employee_Master.xlsx"
)

REPORT_TEMPLATE_FILE = os.path.join(
    BASE_DIR,
    "ORB_OOO_Schedule.xlsx"
)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


app = Flask(__name__)

app.secret_key = "orb-ooo-schedule"

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


# ============================================================
# TEMPLATE STRUCTURE
# ============================================================

# Based on your actual Excel template:
#
#        C       D       E       F       G
#        Mon     Tue     Wed     Thu     Fri
#
# Row 2 = Date
# Row 3 = Online
# Row 4 = Out of Office

DAY_COLUMNS = {
    0: "C",   # Monday
    1: "D",   # Tuesday
    2: "E",   # Wednesday
    3: "F",   # Thursday
    4: "G"    # Friday
}

DATE_ROW = 2
ONLINE_ROW = 3
OOO_ROW = 4


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize(value):

    if value is None:
        return ""

    return " ".join(
        str(value).strip().lower().split()
    )


# ============================================================
# PARSE DATE
# ============================================================

def parse_date(value):

    if value is None:
        raise ValueError(
            "Date cannot be empty."
        )

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = str(value).strip()

    formats = [
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m-%d-%Y",
        "%m-%d-%y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%m/%d",
        "%m-%d"
    ]

    for fmt in formats:

        try:

            result = datetime.strptime(
                text,
                fmt
            )

            # If year was not supplied
            # use current year.

            if "%Y" not in fmt and "%y" not in fmt:

                result = result.replace(
                    year=date.today().year
                )

            return result.date()

        except ValueError:
            continue

    raise ValueError(
        f"Invalid date: {value}. "
        f"Please use MM/DD/YYYY."
    )


# ============================================================
# GET WEEK
# ============================================================

def get_monday(input_date):

    return input_date - timedelta(
        days=input_date.weekday()
    )


def get_friday(monday):

    return monday + timedelta(
        days=4
    )


# ============================================================
# READ EMPLOYEE MASTER
# ============================================================

def read_employee_master():

    if not os.path.exists(
        EMPLOYEE_MASTER_FILE
    ):

        raise FileNotFoundError(
            "ORB_Employee_Master.xlsx "
            "was not found."
        )

    wb = load_workbook(
        EMPLOYEE_MASTER_FILE,
        data_only=True
    )

    if "Employee Master" in wb.sheetnames:

        ws = wb["Employee Master"]

    else:

        ws = wb.active


    # --------------------------------------------------------
    # Find columns
    # --------------------------------------------------------

    name_col = None
    code_col = None

    for col in range(
        1,
        ws.max_column + 1
    ):

        header = normalize(
            ws.cell(1, col).value
        )

        if header in (
            "employee name",
            "employee"
        ):

            name_col = col

        elif header in (
            "employee code",
            "code"
        ):

            code_col = col


    if name_col is None:

        raise ValueError(
            "Employee Master must contain "
            "'Employee Name'."
        )

    if code_col is None:

        raise ValueError(
            "Employee Master must contain "
            "'Employee Code'."
        )


    # --------------------------------------------------------
    # Read employees
    # --------------------------------------------------------

    employee_map = {}

    for row in range(
        2,
        ws.max_row + 1
    ):

        name = ws.cell(
            row,
            name_col
        ).value

        code = ws.cell(
            row,
            code_col
        ).value


        if not name and not code:

            continue


        if not name or not code:

            continue


        employee_map[
            normalize(name)
        ] = str(code).strip()


    if not employee_map:

        raise ValueError(
            "No employees found in "
            "ORB_Employee_Master.xlsx."
        )


    return employee_map


# ============================================================
# READ INPUT FILE
# ============================================================

def read_input_file(file_path):

    wb = load_workbook(
        file_path,
        data_only=True
    )

    ws = wb.active


    # --------------------------------------------------------
    # Find input columns
    # --------------------------------------------------------

    date_col = None
    status_col = None
    employee_col = None


    for col in range(
        1,
        ws.max_column + 1
    ):

        header = normalize(
            ws.cell(1, col).value
        )

        if header == "date":

            date_col = col

        elif header == "status":

            status_col = col

        elif header in (
            "employee name",
            "employee"
        ):

            employee_col = col


    missing = []


    if date_col is None:
        missing.append("Date")


    if status_col is None:
        missing.append("Status")


    if employee_col is None:
        missing.append("Employee Name")


    if missing:

        raise ValueError(
            "Missing columns: "
            + ", ".join(missing)
        )


    records = []


    # --------------------------------------------------------
    # Read rows
    # --------------------------------------------------------

    for row in range(
        2,
        ws.max_row + 1
    ):

        date_value = ws.cell(
            row,
            date_col
        ).value

        status_value = ws.cell(
            row,
            status_col
        ).value

        employee_value = ws.cell(
            row,
            employee_col
        ).value


        # Ignore blank rows.

        if (
            date_value is None
            and status_value is None
            and employee_value is None
        ):

            continue


        if (
            date_value is None
            or status_value is None
            or employee_value is None
        ):

            raise ValueError(
                f"Row {row}: Date, Status "
                f"and Employee Name are required."
            )


        status = normalize(
            status_value
        )


        # We only expect OOO records.

        if status not in (
            "ooo",
            "out of office"
        ):

            raise ValueError(
                f"Row {row}: Status must be "
                f"'OOO' or 'Out of Office'."
            )


        records.append({

            "date": parse_date(
                date_value
            ),

            "employee": str(
                employee_value
            ).strip()

        })


    if not records:

        raise ValueError(
            "No OOO records found."
        )


    return records


# ============================================================
# GENERATE ONE WEEK
# ============================================================

def generate_week_report(
    all_records,
    monday
):

    # --------------------------------------------------------
    # Read employee master
    # --------------------------------------------------------

    employee_map = read_employee_master()


    # Example:
    #
    # Velina  -> VR
    # Cristen -> CR
    # Tracy   -> TP
    # Shakima -> HB
    #

    all_employee_codes = list(
        employee_map.values()
    )


    # --------------------------------------------------------
    # Load template
    # --------------------------------------------------------

    if not os.path.exists(
        REPORT_TEMPLATE_FILE
    ):

        raise FileNotFoundError(
            "ORB_OOO_Schedule.xlsx "
            "was not found."
        )


    wb = load_workbook(
        REPORT_TEMPLATE_FILE
    )


    # --------------------------------------------------------
    # Select worksheet
    # --------------------------------------------------------

    if "Weekly Schedule" in wb.sheetnames:

        ws = wb["Weekly Schedule"]

    else:

        ws = wb.active


    # --------------------------------------------------------
    # Update dates
    # --------------------------------------------------------

    for day_number in range(5):

        current_date = (
            monday
            + timedelta(
                days=day_number
            )
        )

        column = DAY_COLUMNS[
            day_number
        ]


        # Example:
        #
        # C2 = 9/28
        # D2 = 9/29
        # etc.

        ws[f"{column}{DATE_ROW}"] = (
            f"{current_date.month}/"
            f"{current_date.day}"
        )


    # --------------------------------------------------------
    # Get records for this week
    # --------------------------------------------------------

    friday = get_friday(
        monday
    )


    week_records = [

        record

        for record in all_records

        if (
            monday
            <= record["date"]
            <= friday
        )

    ]


    # --------------------------------------------------------
    # Build OOO mapping
    #
    # Example:
    #
    # Thursday -> ["VR", "CR"]
    #
    # --------------------------------------------------------

    ooo_by_date = {}


    for record in week_records:

        employee_name = normalize(
            record["employee"]
        )


        if employee_name not in employee_map:

            raise ValueError(
                f"Employee '{record['employee']}' "
                f"is not present in "
                f"ORB_Employee_Master.xlsx."
            )


        employee_code = employee_map[
            employee_name
        ]


        current_date = record[
            "date"
        ]


        if current_date not in ooo_by_date:

            ooo_by_date[
                current_date
            ] = []


        if employee_code not in ooo_by_date[
            current_date
        ]:

            ooo_by_date[
                current_date
            ].append(
                employee_code
            )


    # ========================================================
    # PROCESS MONDAY-FRIDAY
    # ========================================================

    for day_number in range(5):

        current_date = (
            monday
            + timedelta(
                days=day_number
            )
        )


        column = DAY_COLUMNS[
            day_number
        ]


        # ----------------------------------------------------
        # Employees OOO on this day
        # ----------------------------------------------------

        ooo_codes = ooo_by_date.get(
            current_date,
            []
        )


        # ----------------------------------------------------
        # ONLINE
        #
        # Everyone starts as Online.
        #
        # Remove OOO employees.
        # ----------------------------------------------------

        online_codes = [

            code

            for code in all_employee_codes

            if code not in ooo_codes

        ]


        # ----------------------------------------------------
        # Write ONLINE
        #
        # C3 / D3 / E3 / F3 / G3
        # ----------------------------------------------------

        online_cell = ws[
            f"{column}{ONLINE_ROW}"
        ]


        online_cell.value = "/".join(
            online_codes
        )


        # ----------------------------------------------------
        # Write OUT OF OFFICE
        #
        # C4 / D4 / E4 / F4 / G4
        # ----------------------------------------------------

        ooo_cell = ws[
            f"{column}{OOO_ROW}"
        ]


        ooo_cell.value = "/".join(
            ooo_codes
        )


        # ----------------------------------------------------
        # Make OOO red
        # ----------------------------------------------------

        if ooo_codes:

            new_font = copy(
                ooo_cell.font
            )

            new_font.color = "C00000"

            new_font.bold = True

            ooo_cell.font = new_font


    # ========================================================
    # SAVE REPORT
    # ========================================================

    filename = (
        "ORB_OOO_Schedule_"
        + monday.strftime("%Y%m%d")
        + "_"
        + friday.strftime("%Y%m%d")
        + ".xlsx"
    )


    output_path = os.path.join(
        OUTPUT_FOLDER,
        filename
    )


    wb.save(
        output_path
    )


    return {

        "filename": filename,

        "path": output_path,

        "monday": monday,

        "friday": friday,

        "records": week_records

    }


# ============================================================
# GENERATE ALL WEEKS
# ============================================================

def generate_reports(records):

    weeks = {}


    # --------------------------------------------------------
    # Group dates into Monday-Friday weeks
    # --------------------------------------------------------

    for record in records:

        monday = get_monday(
            record["date"]
        )


        if monday not in weeks:

            weeks[monday] = []


        weeks[monday].append(
            record
        )


    reports = []


    # --------------------------------------------------------
    # Generate each week
    # --------------------------------------------------------

    for monday in sorted(
        weeks.keys()
    ):

        report = generate_week_report(
            records,
            monday
        )


        reports.append(
            report
        )


    return reports


# ============================================================
# HOME
# ============================================================

@app.route(
    "/",
    methods=["GET", "POST"]
)
def index():

    # --------------------------------------------------------
    # GET
    # --------------------------------------------------------

    if request.method == "GET":

        return render_template(
            "index.html"
        )


    # --------------------------------------------------------
    # Uploaded file
    # --------------------------------------------------------

    uploaded_file = request.files.get(
        "ooo_file"
    )


    if (
        not uploaded_file
        or not uploaded_file.filename
    ):

        flash(
            "Please select an Excel file.",
            "error"
        )

        return redirect(
            url_for("index")
        )


    # --------------------------------------------------------
    # Check master file
    # --------------------------------------------------------

    if not os.path.exists(
        EMPLOYEE_MASTER_FILE
    ):

        flash(
            "ORB_Employee_Master.xlsx "
            "is missing.",
            "error"
        )

        return redirect(
            url_for("index")
        )


    # --------------------------------------------------------
    # Check template
    # --------------------------------------------------------

    if not os.path.exists(
        REPORT_TEMPLATE_FILE
    ):

        flash(
            "ORB_OOO_Schedule.xlsx "
            "is missing.",
            "error"
        )

        return redirect(
            url_for("index")
        )


    # --------------------------------------------------------
    # Check file type
    # --------------------------------------------------------

    filename = secure_filename(
        uploaded_file.filename
    )


    if not filename.lower().endswith(
        (
            ".xlsx",
            ".xlsm"
        )
    ):

        flash(
            "Please upload an Excel "
            "file (.xlsx or .xlsm).",
            "error"
        )

        return redirect(
            url_for("index")
        )


    # --------------------------------------------------------
    # Save uploaded file
    # --------------------------------------------------------

    input_path = os.path.join(
        UPLOAD_FOLDER,
        filename
    )


    uploaded_file.save(
        input_path
    )


    # --------------------------------------------------------
    # Generate
    # --------------------------------------------------------

    try:

        records = read_input_file(
            input_path
        )


        reports = generate_reports(
            records
        )


        return render_template(

            "index.html",

            success=True,

            reports=reports

        )


    except Exception as error:

        flash(
            str(error),
            "error"
        )

        return redirect(
            url_for("index")
        )


# ============================================================
# DOWNLOAD
# ============================================================

@app.route(
    "/download/<filename>"
)
def download(filename):

    safe_filename = os.path.basename(
        filename
    )


    file_path = os.path.join(
        OUTPUT_FOLDER,
        safe_filename
    )


    if not os.path.exists(
        file_path
    ):

        flash(
            "Report file not found.",
            "error"
        )

        return redirect(
            url_for("index")
        )


    return send_file(
        file_path,
        as_attachment=True
    )


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route("/health")
def health():

    return {
        "status": "OK"
    }


# ============================================================
# START APPLICATION
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
