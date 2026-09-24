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
# HELPER FUNCTIONS
# ============================================================

def normalize(value):
    """
    Converts text to a standard format for comparison.
    """

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
# DATE PARSING
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

            # If year is not provided,
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
        f"Use MM/DD/YYYY format."
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
            "Employee master must contain "
            "'Employee Name'."
        )


    if not code_col:

        raise ValueError(
            "Employee master must contain "
            "'Employee Code'."
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
                f"Employee master row {row} "
                f"must contain both "
                f"Employee Name and Employee Code."
            )


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
# READ USER INPUT EXCEL
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


        # Ignore completely blank rows.

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
                f"Row {row}: "
                f"Date, Status and Employee Name "
                f"are required."
            )


        status = normalize(
            status_value
        )


        # Only OOO records are expected.

        if status not in (
            "out of office",
            "ooo"
        ):

            raise ValueError(
                f"Row {row}: Status must be "
                f"'Out of Office' or 'OOO'."
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
            "No OOO records found "
            "in the uploaded file."
        )


    return records


# ============================================================
# FIND REPORT STRUCTURE
# ============================================================

def find_schedule_structure(ws):

    online_row = None

    ooo_row = None


    # Find Online / Out of Office rows.

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


    # Find weekday columns.

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


        if len(found) >= 3:

            header_row = row

            weekday_cols = found

            break


    if header_row is None:

        raise ValueError(
            "Could not find Monday-Friday "
            "headers in report template."
        )


    required_days = (
        "monday",
        "tuesday",
        "wednesday",
        "thursday",
        "friday"
    )


    for day in required_days:

        if day not in weekday_cols:

            raise ValueError(
                f"Missing {day.title()} "
                f"column in report template."
            )


    return {

        "online_row": online_row,

        "ooo_row": ooo_row,

        "header_row": header_row,

        "date_row": header_row + 1,

        "weekday_cols": weekday_cols

    }


# ============================================================
# CELL HELPERS
# ============================================================

def get_codes(cell):

    if cell.value is None:

        return []


    return [

        x.strip()

        for x in str(
            cell.value
        ).split("/")

        if x.strip()

    ]


def remove_code(cell, code):

    codes = get_codes(cell)


    cell.value = "/".join(

        x

        for x in codes

        if x.lower() != code.lower()

    )


def add_code(cell, code):

    codes = get_codes(cell)


    exists = any(

        x.lower() == code.lower()

        for x in codes

    )


    if not exists:

        codes.append(code)


    cell.value = "/".join(codes)


def mark_ooo(cell):

    # Keep existing formatting and
    # simply make the OOO value red.

    font = cell.font.copy(
        color="C00000",
        bold=True
    )

    cell.font = font


# ============================================================
# UPDATE WEEK DATES
# ============================================================

def update_week_dates(
    ws,
    structure,
    monday
):

    for i in range(5):

        current_date = (
            monday
            + timedelta(days=i)
        )

        day_name = (
            current_date
            .strftime("%A")
            .lower()
        )

        col = structure[
            "weekday_cols"
        ][day_name]


        # Weekday

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
# GENERATE ONE WEEK
# ============================================================

def process_week(
    records,
    monday
):

    if not os.path.exists(
        REPORT_TEMPLATE_FILE
    ):

        raise FileNotFoundError(
            "ORB_OOO_Schedule.xlsx "
            "was not found."
        )


    employee_map = (
        read_employee_master()
    )


    wb = load_workbook(
        REPORT_TEMPLATE_FILE
    )


    # Use Weekly Schedule if available.

    if "Weekly Schedule" in wb.sheetnames:

        ws = wb["Weekly Schedule"]

    else:

        ws = wb.active


    structure = find_schedule_structure(
        ws
    )


    update_week_dates(
        ws,
        structure,
        monday
    )


    friday = get_friday(
        monday
    )


    # Records belonging to this week.

    week_records = [

        record

        for record in records

        if (
            monday
            <= record["date"]
            <= friday
        )

    ]


    # Validate employees first.

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


    # ========================================================
    # APPLY OOO
    # ========================================================

    for record in week_records:

        employee_key = normalize(
            record["employee"]
        )

        employee_code = (
            employee_map[
                employee_key
            ]
        )


        day_name = (
            record["date"]
            .strftime("%A")
            .lower()
        )


        column = (
            structure[
                "weekday_cols"
            ][day_name]
        )


        # ----------------------------------------------------
        # IMPORTANT:
        # Remove employee from ONLINE
        # if employee is OOO.
        # ----------------------------------------------------

        for row in range(

            structure["online_row"],

            structure["ooo_row"]

        ):

            cell = ws.cell(
                row,
                column
            )

            remove_code(
                cell,
                employee_code
            )


        # ----------------------------------------------------
        # Add employee to OOO.
        # ----------------------------------------------------

        ooo_cell = ws.cell(
            structure["ooo_row"],
            column
        )


        add_code(
            ooo_cell,
            employee_code
        )


        mark_ooo(
            ooo_cell
        )


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
# MAIN PAGE
# ============================================================

@app.route(
    "/",
    methods=["GET", "POST"]
)
def index():

    if request.method == "GET":

        return render_template(
            "index.html"
        )


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


    # Check master file.

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


    # Check report template.

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


    filename = secure_filename(
        uploaded_file.filename
    )


    if not filename.lower().endswith(
        (".xlsx", ".xlsm")
    ):

        flash(
            "Please upload an Excel "
            "file (.xlsx or .xlsm).",
            "error"
        )

        return redirect(
            url_for("index")
        )


    input_path = os.path.join(
        UPLOAD_FOLDER,
        filename
    )


    uploaded_file.save(
        input_path
    )


    try:

        # Read uploaded Excel.

        records = read_input_file(
            input_path
        )


        # Automatically identify
        # all weeks and generate reports.

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
# DOWNLOAD REPORT
# ============================================================

@app.route(
    "/download/<filename>"
)
def download(filename):

    # Prevent path traversal.

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
# RUN APPLICATION
# ============================================================

if __name__ == "__main__":

    app.run(

        host="0.0.0.0",

        port=5000,

        debug=True

    )
