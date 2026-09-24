import os
import re
from datetime import datetime, date, timedelta

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
# TEXT HELPERS
# ============================================================

def normalize(value):

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value).strip().lower()
    )


def normalize_header(value):

    return normalize(value).replace("_", " ")


# ============================================================
# DATE HELPERS
# ============================================================

def parse_date(value):

    if value is None:
        raise ValueError("Date is required.")

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

            parsed = datetime.strptime(
                text,
                fmt
            )

            # If user gives only month/day,
            # use current year.

            if "%Y" not in fmt and "%y" not in fmt:

                parsed = parsed.replace(
                    year=date.today().year
                )

            return parsed.date()

        except ValueError:
            continue

    raise ValueError(
        f"Invalid date '{value}'. "
        f"Use MM/DD/YYYY."
    )


def get_monday(value):

    return value - timedelta(
        days=value.weekday()
    )


def get_friday(value):

    return get_monday(value) + timedelta(
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
            "was not found next to app.py."
        )

    wb = load_workbook(
        EMPLOYEE_MASTER_FILE,
        data_only=True
    )

    if "Employee Master" in wb.sheetnames:

        ws = wb["Employee Master"]

    else:

        ws = wb.active

    headers = {}

    for col in range(
        1,
        ws.max_column + 1
    ):

        header = normalize_header(
            ws.cell(1, col).value
        )

        if header:

            headers[header] = col

    name_col = (
        headers.get("employee name")
        or headers.get("employee")
    )

    code_col = (
        headers.get("employee code")
        or headers.get("code")
    )

    if not name_col:

        raise ValueError(
            "ORB_Employee_Master.xlsx must "
            "contain 'Employee Name'."
        )

    if not code_col:

        raise ValueError(
            "ORB_Employee_Master.xlsx must "
            "contain 'Employee Code'."
        )

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

        if name is None and code is None:
            continue

        if not name or not code:

            raise ValueError(
                f"Employee Master row {row} "
                f"must contain Employee Name "
                f"and Employee Code."
            )

        employee_map[
            normalize(name)
        ] = str(code).strip()

    if not employee_map:

        raise ValueError(
            "No employees were found in "
            "ORB_Employee_Master.xlsx."
        )

    return employee_map


# ============================================================
# READ INPUT EXCEL
#
# Expected columns:
#
# Date | Status | Employee Name
#
# ============================================================

def read_input_file(file_path):

    wb = load_workbook(
        file_path,
        data_only=True
    )

    ws = wb.active

    headers = {}

    for col in range(
        1,
        ws.max_column + 1
    ):

        header = normalize_header(
            ws.cell(1, col).value
        )

        if header:

            headers[header] = col

    date_col = headers.get("date")

    status_col = headers.get("status")

    employee_col = (
        headers.get("employee name")
        or headers.get("employee")
    )

    missing = []

    if not date_col:
        missing.append("Date")

    if not status_col:
        missing.append("Status")

    if not employee_col:
        missing.append("Employee Name")

    if missing:

        raise ValueError(
            "Input Excel is missing: "
            + ", ".join(missing)
        )

    records = []

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
# FIND REPORT TEMPLATE STRUCTURE
# ============================================================

def find_schedule_structure(ws):

    online_row = None

    ooo_row = None

    # --------------------------------------------------------
    # Find Online and Out of Office rows
    # --------------------------------------------------------

    for row in range(
        1,
        ws.max_row + 1
    ):

        for col in range(
            1,
            ws.max_column + 1
        ):

            value = normalize(
                ws.cell(
                    row,
                    col
                ).value
            )

            if value == "online":

                online_row = row

            elif value == "out of office":

                ooo_row = row

    if online_row is None:

        raise ValueError(
            "Could not find 'Online' "
            "in ORB_OOO_Schedule.xlsx."
        )

    if ooo_row is None:

        raise ValueError(
            "Could not find 'Out of Office' "
            "in ORB_OOO_Schedule.xlsx."
        )

    # --------------------------------------------------------
    # Find Monday-Friday columns
    # --------------------------------------------------------

    weekday_cols = {}

    header_row = None

    for row in range(
        1,
        min(ws.max_row, 20) + 1
    ):

        found = {}

        for col in range(
            1,
            ws.max_column + 1
        ):

            value = normalize(
                ws.cell(
                    row,
                    col
                ).value
            )

            if value in (
                "monday",
                "tuesday",
                "wednesday",
                "thursday",
                "friday"
            ):

                found[value] = col

        if len(found) == 5:

            header_row = row

            weekday_cols = found

            break

    if header_row is None:

        raise ValueError(
            "Could not find Monday-Friday "
            "headers in the report template."
        )

    return {

        "online_row": online_row,

        "ooo_row": ooo_row,

        "header_row": header_row,

        "date_row": header_row + 1,

        "weekday_cols": weekday_cols

    }


# ============================================================
# UPDATE DATES
# ============================================================

def update_week_dates(
    ws,
    structure,
    monday
):

    for i in range(5):

        current_date = (
            monday + timedelta(days=i)
        )

        day_name = (
            current_date
            .strftime("%A")
            .lower()
        )

        col = structure[
            "weekday_cols"
        ][day_name]

        # Day name

        ws.cell(
            structure["header_row"],
            col
        ).value = current_date.strftime(
            "%A"
        )

        # Date

        ws.cell(
            structure["date_row"],
            col
        ).value = (
            f"{current_date.month}/"
            f"{current_date.day}"
        )


# ============================================================
# CLEAR CELL
# ============================================================

def clear_cell(cell):

    cell.value = None


# ============================================================
# ADD CODES TO CELL
# ============================================================

def set_codes(cell, codes):

    cell.value = "/".join(
        str(code)
        for code in codes
    )


# ============================================================
# MARK OOO
# ============================================================

def mark_ooo(cell):

    try:

        cell.font = cell.font.copy(
            color="C00000",
            bold=True
        )

    except Exception:

        pass


# ============================================================
# GENERATE ONE WEEK
# ============================================================

def process_week(
    records,
    monday
):

    # --------------------------------------------------------
    # Check template
    # --------------------------------------------------------

    if not os.path.exists(
        REPORT_TEMPLATE_FILE
    ):

        raise FileNotFoundError(
            "ORB_OOO_Schedule.xlsx "
            "was not found."
        )

    # --------------------------------------------------------
    # Read employee master
    # --------------------------------------------------------

    employee_map = read_employee_master()

    # Example:
    #
    # {
    #     "velina": "VR",
    #     "cristen": "CR",
    #     "tracy": "TP",
    #     "shakima": "HB"
    # }

    all_employee_codes = list(
        employee_map.values()
    )

    # --------------------------------------------------------
    # Load report template
    # --------------------------------------------------------

    wb = load_workbook(
        REPORT_TEMPLATE_FILE
    )

    if "Weekly Schedule" in wb.sheetnames:

        ws = wb["Weekly Schedule"]

    else:

        ws = wb.active

    # --------------------------------------------------------
    # Find structure
    # --------------------------------------------------------

    structure = find_schedule_structure(
        ws
    )

    # --------------------------------------------------------
    # Update dates
    # --------------------------------------------------------

    update_week_dates(
        ws,
        structure,
        monday
    )

    friday = get_friday(
        monday
    )

    # --------------------------------------------------------
    # Get this week's records
    # --------------------------------------------------------

    week_records = [

        record

        for record in records

        if (
            monday
            <= record["date"]
            <= friday
        )

    ]

    # --------------------------------------------------------
    # Validate employees
    # --------------------------------------------------------

    for record in week_records:

        employee_key = normalize(
            record["employee"]
        )

        if employee_key not in employee_map:

            raise ValueError(
                f"Employee "
                f"'{record['employee']}' "
                f"is not present in "
                f"ORB_Employee_Master.xlsx."
            )

    # --------------------------------------------------------
    # Create OOO mapping
    #
    # Date -> Employee Codes
    #
    # --------------------------------------------------------

    ooo_by_date = {}

    for record in week_records:

        employee_key = normalize(
            record["employee"]
        )

        employee_code = employee_map[
            employee_key
        ]

        record_date = record["date"]

        if record_date not in ooo_by_date:

            ooo_by_date[
                record_date
            ] = []

        if employee_code not in ooo_by_date[
            record_date
        ]:

            ooo_by_date[
                record_date
            ].append(
                employee_code
            )

    # --------------------------------------------------------
    # PROCESS MONDAY-FRIDAY
    # --------------------------------------------------------

    for day_number in range(5):

        current_date = (
            monday
            + timedelta(
                days=day_number
            )
        )

        day_name = (
            current_date
            .strftime("%A")
            .lower()
        )

        column = structure[
            "weekday_cols"
        ][day_name]

        # ====================================================
        # GET OOO EMPLOYEES FOR THIS DAY
        # ====================================================

        ooo_codes = ooo_by_date.get(
            current_date,
            []
        )

        # ====================================================
        # ONLINE
        #
        # Start with EVERY employee.
        # Remove employees who are OOO.
        # ====================================================

        online_codes = [

            code

            for code in all_employee_codes

            if code not in ooo_codes

        ]

        # ----------------------------------------------------
        # Clear existing Online cells
        #
        # This clears everything between Online and OOO.
        # ----------------------------------------------------

        for row in range(
            structure["online_row"],
            structure["ooo_row"]
        ):

            ws.cell(
                row,
                column
            ).value = None

        # ----------------------------------------------------
        # Put ALL ONLINE employees into first Online cell
        # ----------------------------------------------------

        online_cell = ws.cell(
            structure["online_row"],
            column
        )

        set_codes(
            online_cell,
            online_codes
        )

        # ====================================================
        # OUT OF OFFICE
        # ====================================================

        ooo_cell = ws.cell(
            structure["ooo_row"],
            column
        )

        # Clear old OOO value first.

        ooo_cell.value = None

        if ooo_codes:

            set_codes(
                ooo_cell,
                ooo_codes
            )

            mark_ooo(
                ooo_cell
            )

    # --------------------------------------------------------
    # SAVE FILE
    # --------------------------------------------------------

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
# GENERATE REPORTS
# ============================================================

def generate_reports(records):

    weeks = {}

    # --------------------------------------------------------
    # Group records by week
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

        report = process_week(
            records,
            monday
        )

        reports.append(
            report
        )

    return reports


# ============================================================
# HOME / UPLOAD
# ============================================================

@app.route(
    "/",
    methods=["GET", "POST"]
)
def index():

    # --------------------------------------------------------
    # Display UI
    # --------------------------------------------------------

    if request.method == "GET":

        return render_template(
            "index.html"
        )

    # --------------------------------------------------------
    # Get uploaded file
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
    # Check employee master
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
    # Check report template
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
    # Validate extension
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
    # Generate report
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
        debug=False
    )
