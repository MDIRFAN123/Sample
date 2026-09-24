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
# APPLICATION CONFIGURATION
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

UPLOAD_FOLDER = os.path.join(
    BASE_DIR,
    "uploads"
)

OUTPUT_FOLDER = os.path.join(
    BASE_DIR,
    "output"
)

EMPLOYEE_MASTER_FILE = os.path.join(
    BASE_DIR,
    "ORB_Employee_Master.xlsx"
)

REPORT_TEMPLATE_FILE = os.path.join(
    BASE_DIR,
    "ORB_OOO_Schedule.xlsx"
)


os.makedirs(
    UPLOAD_FOLDER,
    exist_ok=True
)

os.makedirs(
    OUTPUT_FOLDER,
    exist_ok=True
)


app = Flask(__name__)

app.secret_key = "orb-ooo-schedule"

app.config[
    "MAX_CONTENT_LENGTH"
] = 50 * 1024 * 1024


# ============================================================
# EXCEL TEMPLATE LAYOUT
# ============================================================

# Latest template:
#
#        B          C          D          E          F
#        Monday     Tuesday    Wednesday  Thursday   Friday
#
# Row 1 = Day
# Row 2 = Online
# Row 3 = Out of Office
#
# A2 = Online
# A3 = Out of Office
#
# ============================================================

DAY_COLUMNS = {
    0: "B",       # Monday
    1: "C",       # Tuesday
    2: "D",       # Wednesday
    3: "E",       # Thursday
    4: "F"        # Friday
}

ONLINE_ROW = 2

OOO_ROW = 3


# ============================================================
# TEXT NORMALIZATION
# ============================================================

def normalize(value):

    if value is None:
        return ""

    return " ".join(
        str(value)
        .strip()
        .lower()
        .split()
    )


# ============================================================
# DATE PARSER
# ============================================================

def parse_date(value):

    if value is None:
        raise ValueError(
            "Date cannot be empty."
        )

    # Excel datetime
    if isinstance(
        value,
        datetime
    ):
        return value.date()

    # Excel date
    if isinstance(
        value,
        date
    ):
        return value

    text = str(
        value
    ).strip()

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

            # If year is not provided,
            # use current year.

            if (
                "%Y" not in fmt
                and "%y" not in fmt
            ):

                result = result.replace(
                    year=date.today().year
                )

            return result.date()

        except ValueError:

            continue

    raise ValueError(
        f"Invalid date '{value}'. "
        "Please use MM/DD/YYYY."
    )


# ============================================================
# WEEK HELPERS
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

    if (
        "Employee Master"
        in wb.sheetnames
    ):

        ws = wb[
            "Employee Master"
        ]

    else:

        ws = wb.active


    name_col = None

    code_col = None


    # --------------------------------------------------------
    # Find columns
    # --------------------------------------------------------

    for col in range(
        1,
        ws.max_column + 1
    ):

        header = normalize(
            ws.cell(
                1,
                col
            ).value
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


        if (
            name is None
            and code is None
        ):

            continue


        if (
            name is None
            or code is None
        ):

            continue


        employee_map[
            normalize(name)
        ] = str(
            code
        ).strip()


    if not employee_map:

        raise ValueError(
            "No employees found in "
            "ORB_Employee_Master.xlsx."
        )


    return employee_map


# ============================================================
# READ INPUT EXCEL
#
# Required columns:
#
# Date
# Status
# Employee Name
#
# ============================================================

def read_input_file(
    file_path
):

    wb = load_workbook(
        file_path,
        data_only=True
    )

    ws = wb.active


    date_col = None

    status_col = None

    employee_col = None


    # --------------------------------------------------------
    # Find headers
    # --------------------------------------------------------

    for col in range(
        1,
        ws.max_column + 1
    ):

        header = normalize(
            ws.cell(
                1,
                col
            ).value
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

        missing.append(
            "Date"
        )


    if status_col is None:

        missing.append(
            "Status"
        )


    if employee_col is None:

        missing.append(
            "Employee Name"
        )


    if missing:

        raise ValueError(
            "Input Excel is missing: "
            + ", ".join(missing)
        )


    records = []


    # --------------------------------------------------------
    # Read data
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


        # Ignore completely blank rows

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
                "Date, Status and "
                "Employee Name are required."
            )


        status = normalize(
            status_value
        )


        # ----------------------------------------------------
        # Only OOO records are required
        # ----------------------------------------------------

        if status not in (
            "ooo",
            "out of office"
        ):

            raise ValueError(
                f"Row {row}: Status must be "
                "'OOO' or 'Out of Office'."
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
# WRITE CELL WHILE PRESERVING FORMAT
# ============================================================

def write_cell(
    ws,
    cell_address,
    value,
    red=False
):

    cell = ws[
        cell_address
    ]


    cell.value = value


    # --------------------------------------------------------
    # Preserve existing font
    # --------------------------------------------------------

    if red:

        new_font = copy(
            cell.font
        )

        new_font.color = "C00000"

        new_font.bold = True

        cell.font = new_font


# ============================================================
# GENERATE ONE WEEK
# ============================================================

def generate_week_report(
    records,
    monday
):

    # --------------------------------------------------------
    # Read master
    # --------------------------------------------------------

    employee_map = (
        read_employee_master()
    )


    # --------------------------------------------------------
    # Complete employee list
    # --------------------------------------------------------

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
    # Use correct sheet
    # --------------------------------------------------------

    if (
        "Weekly Schedule"
        in wb.sheetnames
    ):

        ws = wb[
            "Weekly Schedule"
        ]

    else:

        ws = wb.active


    # --------------------------------------------------------
    # Determine Friday
    # --------------------------------------------------------

    friday = get_friday(
        monday
    )


    # ========================================================
    # BUILD OOO MAP
    # ========================================================

    ooo_by_date = {}


    for record in records:

        record_date = record[
            "date"
        ]


        # Only this week

        if not (
            monday
            <= record_date
            <= friday
        ):

            continue


        employee_name = normalize(
            record[
                "employee"
            ]
        )


        # ----------------------------------------------------
        # Validate employee
        # ----------------------------------------------------

        if employee_name not in employee_map:

            raise ValueError(
                f"Employee "
                f"'{record['employee']}' "
                "is not present in "
                "ORB_Employee_Master.xlsx."
            )


        employee_code = employee_map[
            employee_name
        ]


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


    # ========================================================
    # MONDAY - FRIDAY
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
        # OOO employees for this date
        # ----------------------------------------------------

        ooo_codes = ooo_by_date.get(
            current_date,
            []
        )


        # ----------------------------------------------------
        # ONLINE EMPLOYEES
        #
        # Start with everyone.
        #
        # Remove OOO employees.
        # ----------------------------------------------------

        online_codes = [

            code

            for code
            in all_employee_codes

            if code
            not in ooo_codes

        ]


        # ----------------------------------------------------
        # Convert to display text
        # ----------------------------------------------------

        online_text = "/".join(
            online_codes
        )


        ooo_text = "/".join(
            ooo_codes
        )


        # ----------------------------------------------------
        # WRITE ONLINE
        #
        # Monday    = B2
        # Tuesday   = C2
        # Wednesday = D2
        # Thursday  = E2
        # Friday    = F2
        # ----------------------------------------------------

        write_cell(
            ws,
            f"{column}{ONLINE_ROW}",
            online_text,
            red=False
        )


        # ----------------------------------------------------
        # WRITE OOO
        #
        # Monday    = B3
        # Tuesday   = C3
        # Wednesday = D3
        # Thursday  = E3
        # Friday    = F3
        # ----------------------------------------------------

        write_cell(
            ws,
            f"{column}{OOO_ROW}",
            ooo_text,
            red=bool(ooo_codes)
        )


    # ========================================================
    # SAVE REPORT
    # ========================================================

    filename = (
        "ORB_OOO_Schedule_"
        + monday.strftime(
            "%Y%m%d"
        )
        + "_"
        + friday.strftime(
            "%Y%m%d"
        )
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
        "friday": friday
    }


# ============================================================
# GENERATE REPORTS
# ============================================================

def generate_reports(
    records
):

    # --------------------------------------------------------
    # Find all weeks
    # --------------------------------------------------------

    weeks = set()


    for record in records:

        monday = get_monday(
            record["date"]
        )

        weeks.add(
            monday
        )


    reports = []


    # --------------------------------------------------------
    # Generate every required week
    # --------------------------------------------------------

    for monday in sorted(
        weeks
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
# HOME / UPLOAD
# ============================================================

@app.route(
    "/",
    methods=[
        "GET",
        "POST"
    ]
)
def index():

    # --------------------------------------------------------
    # Display page
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
    # Check master
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
    # Save input
    # --------------------------------------------------------

    input_path = os.path.join(
        UPLOAD_FOLDER,
        filename
    )


    uploaded_file.save(
        input_path
    )


    # ========================================================
    # AUTOMATIC GENERATION
    # ========================================================

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
# DOWNLOAD REPORT
# ============================================================

@app.route(
    "/download/<filename>"
)
def download(
    filename
):

    filename = os.path.basename(
        filename
    )


    file_path = os.path.join(
        OUTPUT_FOLDER,
        filename
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

@app.route(
    "/health"
)
def health():

    return {
        "status": "OK"
    }


# ============================================================
# START SERVER
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
