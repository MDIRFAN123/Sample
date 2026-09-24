import os
import re
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

MASTER_FILE = os.path.join(
    BASE_DIR,
    "ORB_OOO_Schedule.xlsx"
)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


app = Flask(__name__)

app.secret_key = "orb-ooo-schedule-secret"

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# ============================================================
# EMPLOYEE MAPPING
#
# These are fallback mappings.
#
# If the master Excel contains an
# "Employee Mapping" sheet, that mapping is used first.
# ============================================================

DEFAULT_EMPLOYEE_CODES = {
    "velina": "VR",
    "cristen": "CR",
    "tracy": "TP",
    "shakima": "HB",
}


# ============================================================
# NORMALIZE NAME
# ============================================================

def normalize_name(value):

    if value is None:
        return ""

    value = str(value).strip().lower()

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value


# ============================================================
# NORMALIZE HEADER
# ============================================================

def normalize_header(value):

    if value is None:
        return ""

    value = str(value).strip().lower()

    value = value.replace("_", " ")

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value


# ============================================================
# DATE PARSER
# ============================================================

def parse_excel_date(value):

    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = str(value).strip()

    date_formats = [
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m/%d",
        "%m-%d-%Y",
        "%m-%d-%y",
        "%m-%d",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
    ]

    for date_format in date_formats:

        try:

            parsed = datetime.strptime(
                text,
                date_format
            )

            # If year is not supplied,
            # use current year.
            if (
                "%Y" not in date_format
                and "%y" not in date_format
            ):
                parsed = parsed.replace(
                    year=date.today().year
                )

            return parsed.date()

        except ValueError:

            continue

    raise ValueError(
        f"Invalid date '{value}'. "
        "Please use MM/DD/YYYY."
    )


# ============================================================
# WEEK FUNCTIONS
# ============================================================

def get_monday(input_date):

    return (
        input_date
        - timedelta(
            days=input_date.weekday()
        )
    )


def get_week_dates(input_date):

    monday = get_monday(input_date)

    return [
        monday + timedelta(days=i)
        for i in range(5)
    ]


def display_date(value):

    return f"{value.month}/{value.day}"


# ============================================================
# EMPLOYEE MAPPING
# ============================================================

def load_employee_mapping(workbook):

    mapping = {}

    # --------------------------------------------------------
    # Employee Mapping sheet
    # --------------------------------------------------------

    if "Employee Mapping" in workbook.sheetnames:

        ws = workbook["Employee Mapping"]

        for row in range(
            2,
            ws.max_row + 1
        ):

            employee = ws.cell(
                row,
                1
            ).value

            code = ws.cell(
                row,
                2
            ).value

            if (
                employee is not None
                and code is not None
            ):

                mapping[
                    normalize_name(employee)
                ] = str(
                    code
                ).strip().upper()

    # --------------------------------------------------------
    # Fallback mapping
    # --------------------------------------------------------

    for employee, code in DEFAULT_EMPLOYEE_CODES.items():

        if employee not in mapping:

            mapping[
                employee
            ] = code.upper()

    return mapping


def get_employee_code(
    employee_name,
    mapping
):

    normalized = normalize_name(
        employee_name
    )

    if normalized in mapping:

        return mapping[
            normalized
        ]

    raise ValueError(
        f"Employee '{employee_name}' "
        "was not found in the Employee Mapping."
    )


# ============================================================
# CODE HELPERS
# ============================================================

def split_codes(value):

    if value is None:
        return []

    text = str(value).strip()

    if not text:
        return []

    return [
        item.strip()
        for item in text.split("/")
        if item.strip()
    ]


def remove_from_online(
    cell,
    employee_code
):

    existing_codes = split_codes(
        cell.value
    )

    employee_code = (
        employee_code
        .strip()
        .upper()
    )

    updated_codes = []

    for code in existing_codes:

        if code.upper() != employee_code:

            updated_codes.append(code)

    cell.value = "/".join(
        updated_codes
    )


def add_to_ooo(
    cell,
    employee_code
):

    existing_codes = split_codes(
        cell.value
    )

    employee_code = (
        employee_code
        .strip()
        .upper()
    )

    existing_upper = [
        code.upper()
        for code in existing_codes
    ]

    if employee_code not in existing_upper:

        existing_codes.append(
            employee_code
        )

    cell.value = "/".join(
        existing_codes
    )


# ============================================================
# OOO FONT
# ============================================================

def make_ooo_cell_red(cell):

    new_font = copy(
        cell.font
    )

    new_font.color = "C00000"

    new_font.bold = True

    cell.font = new_font


# ============================================================
# FIND SCHEDULE STRUCTURE
#
# Dynamically identifies:
#
# - Online row
# - Out of Office row
# - Monday-Friday columns
# - Header row
# - Date row
# ============================================================

def find_schedule_structure(
    worksheet
):

    weekdays = {
        "monday": None,
        "tuesday": None,
        "wednesday": None,
        "thursday": None,
        "friday": None
    }

    online_row = None
    online_col = None

    ooo_row = None
    ooo_col = None

    # --------------------------------------------------------
    # Find Online / OOO
    # --------------------------------------------------------

    for row in range(
        1,
        worksheet.max_row + 1
    ):

        for column in range(
            1,
            worksheet.max_column + 1
        ):

            value = worksheet.cell(
                row,
                column
            ).value

            if value is None:
                continue

            text = str(
                value
            ).strip().lower()

            if text == "online":

                online_row = row
                online_col = column

            elif text == "out of office":

                ooo_row = row
                ooo_col = column

    if online_row is None:

        raise ValueError(
            "Could not find the 'Online' section "
            "in the ORB template."
        )

    if ooo_row is None:

        raise ValueError(
            "Could not find the 'Out of Office' section "
            "in the ORB template."
        )

    if ooo_row <= online_row:

        raise ValueError(
            "The Out of Office section must be "
            "below the Online section."
        )

    # --------------------------------------------------------
    # Find weekday header row
    # --------------------------------------------------------

    weekday_names = set(
        weekdays.keys()
    )

    header_row = None

    for row in range(
        1,
        min(
            worksheet.max_row,
            15
        ) + 1
    ):

        found_days = []

        for column in range(
            1,
            worksheet.max_column + 1
        ):

            value = worksheet.cell(
                row,
                column
            ).value

            if value is None:
                continue

            normalized = (
                str(value)
                .strip()
                .lower()
            )

            if normalized in weekday_names:

                found_days.append(
                    normalized
                )

        if len(found_days) >= 3:

            header_row = row
            break

    if header_row is None:

        raise ValueError(
            "Could not find Monday-Friday "
            "headers in the ORB template."
        )

    # --------------------------------------------------------
    # Find weekday columns
    # --------------------------------------------------------

    for column in range(
        1,
        worksheet.max_column + 1
    ):

        value = worksheet.cell(
            header_row,
            column
        ).value

        if value is None:
            continue

        normalized = (
            str(value)
            .strip()
            .lower()
        )

        if normalized in weekdays:

            weekdays[
                normalized
            ] = column

    # --------------------------------------------------------
    # Validate weekdays
    # --------------------------------------------------------

    missing_days = [
        day
        for day, column in weekdays.items()
        if column is None
    ]

    if missing_days:

        raise ValueError(
            "Could not find these weekday columns: "
            + ", ".join(missing_days)
        )

    date_row = header_row + 1

    return {
        "online_row": online_row,
        "online_col": online_col,
        "ooo_row": ooo_row,
        "ooo_col": ooo_col,
        "header_row": header_row,
        "date_row": date_row,
        "weekday_columns": weekdays
    }


# ============================================================
# EMPLOYEE CHECK
# ============================================================

def employee_exists_in_master(
    worksheet,
    employee_name
):

    target = normalize_name(
        employee_name
    )

    for row in range(
        1,
        worksheet.max_row + 1
    ):

        for column in range(
            1,
            worksheet.max_column + 1
        ):

            value = worksheet.cell(
                row,
                column
            ).value

            if value is None:
                continue

            text = str(value)

            parts = [
                normalize_name(part)
                for part in text.split("/")
            ]

            if target in parts:

                return True

    return False


# ============================================================
# UPDATE WEEK HEADERS
# ============================================================

def update_week_headers(
    worksheet,
    structure,
    monday
):

    week_dates = get_week_dates(
        monday
    )

    weekday_columns = structure[
        "weekday_columns"
    ]

    header_row = structure[
        "header_row"
    ]

    date_row = structure[
        "date_row"
    ]

    for current_date in week_dates:

        weekday_name = (
            current_date
            .strftime("%A")
            .lower()
        )

        column = weekday_columns[
            weekday_name
        ]

        worksheet.cell(
            header_row,
            column
        ).value = current_date.strftime(
            "%A"
        )

        worksheet.cell(
            date_row,
            column
        ).value = display_date(
            current_date
        )


# ============================================================
# UPDATE WEEK LABEL
# ============================================================

def update_week_label(
    worksheet,
    structure,
    monday,
    friday
):

    new_label = (
        f"{display_date(monday)}-"
        f"{display_date(friday)}"
    )

    date_range_pattern = re.compile(
        r"\d{1,2}/\d{1,2}"
    )

    # --------------------------------------------------------
    # Look for an existing date-range cell
    # --------------------------------------------------------

    for row in range(
        1,
        min(
            worksheet.max_row,
            15
        ) + 1
    ):

        for column in range(
            1,
            min(
                worksheet.max_column,
                5
            ) + 1
        ):

            cell = worksheet.cell(
                row,
                column
            )

            value = cell.value

            if value is None:
                continue

            text = str(value).strip()

            if (
                date_range_pattern.search(text)
                and "-" in text
            ):

                cell.value = new_label

                return

    # --------------------------------------------------------
    # Fallback
    # --------------------------------------------------------

    worksheet.cell(
        structure["online_row"],
        1
    ).value = new_label


# ============================================================
# GENERATE ONE WEEK REPORT
# ============================================================

def generate_week_report(
    master_file,
    week_records,
    monday
):

    friday = (
        monday
        + timedelta(days=4)
    )

    # --------------------------------------------------------
    # Load master workbook
    # --------------------------------------------------------

    workbook = load_workbook(
        master_file
    )

    # --------------------------------------------------------
    # Select schedule sheet
    # --------------------------------------------------------

    if "Weekly Schedule" in workbook.sheetnames:

        worksheet = workbook[
            "Weekly Schedule"
        ]

    else:

        worksheet = workbook.active

    # --------------------------------------------------------
    # Detect structure
    # --------------------------------------------------------

    structure = find_schedule_structure(
        worksheet
    )

    # --------------------------------------------------------
    # Employee mapping
    # --------------------------------------------------------

    employee_mapping = (
        load_employee_mapping(
            workbook
        )
    )

    # --------------------------------------------------------
    # Update week
    # --------------------------------------------------------

    update_week_headers(
        worksheet,
        structure,
        monday
    )

    update_week_label(
        worksheet,
        structure,
        monday,
        friday
    )

    # --------------------------------------------------------
    # Dynamic rows
    # --------------------------------------------------------

    online_start_row = structure[
        "online_row"
    ]

    ooo_row = structure[
        "ooo_row"
    ]

    weekday_columns = structure[
        "weekday_columns"
    ]

    # --------------------------------------------------------
    # Process OOO records
    # --------------------------------------------------------

    for record in week_records:

        ooo_date = record["date"]

        employee_name = record[
            "employee"
        ]

        # Only process records in this week
        if not (
            monday
            <= ooo_date
            <= friday
        ):

            continue

        # ----------------------------------------------------
        # Employee code
        # ----------------------------------------------------

        employee_code = (
            get_employee_code(
                employee_name,
                employee_mapping
            )
        )

        # ----------------------------------------------------
        # Check employee
        # ----------------------------------------------------

        if not employee_exists_in_master(
            worksheet,
            employee_name
        ):

            raise ValueError(
                f"Employee '{employee_name}' "
                "was not found in the ORB master schedule."
            )

        # ----------------------------------------------------
        # Determine day column
        # ----------------------------------------------------

        weekday_name = (
            ooo_date
            .strftime("%A")
            .lower()
        )

        excel_column = weekday_columns[
            weekday_name
        ]

        # ----------------------------------------------------
        # Remove from Online
        # ----------------------------------------------------

        for online_row in range(
            online_start_row,
            ooo_row
        ):

            online_cell = worksheet.cell(
                online_row,
                excel_column
            )

            remove_from_online(
                online_cell,
                employee_code
            )

        # ----------------------------------------------------
        # Add to OOO
        # ----------------------------------------------------

        ooo_cell = worksheet.cell(
            ooo_row,
            excel_column
        )

        add_to_ooo(
            ooo_cell,
            employee_code
        )

        # ----------------------------------------------------
        # OOO formatting
        # ----------------------------------------------------

        make_ooo_cell_red(
            ooo_cell
        )

    # --------------------------------------------------------
    # Output filename
    # --------------------------------------------------------

    output_filename = (
        "ORB_OOO_Schedule_"
        f"{monday.strftime('%Y%m%d')}_"
        f"{friday.strftime('%Y%m%d')}.xlsx"
    )

    output_path = os.path.join(
        OUTPUT_FOLDER,
        output_filename
    )

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    workbook.save(
        output_path
    )

    return {
        "filename": output_filename,
        "path": output_path,
        "monday": monday,
        "friday": friday,
        "records": week_records
    }


# ============================================================
# GENERATE ALL REPORTS
# ============================================================

def generate_reports(
    master_file,
    records
):

    weekly_records = {}

    # --------------------------------------------------------
    # Group records by week
    # --------------------------------------------------------

    for record in records:

        input_date = record[
            "date"
        ]

        monday = get_monday(
            input_date
        )

        if monday not in weekly_records:

            weekly_records[
                monday
            ] = []

        weekly_records[
            monday
        ].append(record)

    generated_reports = []

    # --------------------------------------------------------
    # Generate each week
    # --------------------------------------------------------

    for monday in sorted(
        weekly_records.keys()
    ):

        report = generate_week_report(
            master_file,
            weekly_records[monday],
            monday
        )

        generated_reports.append(
            report
        )

    return generated_reports


# ============================================================
# READ INPUT EXCEL
#
# Required columns:
# Date
# Status
# Employee Name
# ============================================================

def read_ooo_input(
    file_path
):

    workbook = load_workbook(
        file_path,
        data_only=True
    )

    worksheet = workbook.active

    # --------------------------------------------------------
    # Read headers
    # --------------------------------------------------------

    headers = {}

    for column in range(
        1,
        worksheet.max_column + 1
    ):

        header = normalize_header(
            worksheet.cell(
                1,
                column
            ).value
        )

        if header:

            headers[
                header
            ] = column

    # --------------------------------------------------------
    # Required columns
    # --------------------------------------------------------

    date_column = headers.get(
        "date"
    )

    status_column = headers.get(
        "status"
    )

    employee_column = headers.get(
        "employee name"
    )

    if not employee_column:

        employee_column = headers.get(
            "employee"
        )

    if not date_column:

        raise ValueError(
            "The input Excel must contain "
            "a 'Date' column."
        )

    if not status_column:

        raise ValueError(
            "The input Excel must contain "
            "a 'Status' column."
        )

    if not employee_column:

        raise ValueError(
            "The input Excel must contain "
            "an 'Employee Name' column."
        )

    records = []

    # --------------------------------------------------------
    # Read rows
    # --------------------------------------------------------

    for row in range(
        2,
        worksheet.max_row + 1
    ):

        date_value = worksheet.cell(
            row,
            date_column
        ).value

        status_value = worksheet.cell(
            row,
            status_column
        ).value

        employee_value = worksheet.cell(
            row,
            employee_column
        ).value

        # Ignore blank rows
        if (
            date_value is None
            and status_value is None
            and employee_value is None
        ):

            continue

        if date_value is None:

            raise ValueError(
                f"Row {row}: Date is missing."
            )

        if status_value is None:

            raise ValueError(
                f"Row {row}: Status is missing."
            )

        if employee_value is None:

            raise ValueError(
                f"Row {row}: Employee Name is missing."
            )

        # ----------------------------------------------------
        # Status
        # ----------------------------------------------------

        status = (
            str(status_value)
            .strip()
            .lower()
        )

        if status not in [
            "out of office",
            "ooo"
        ]:

            raise ValueError(
                f"Row {row}: Status must be "
                "'Out of Office' or 'OOO'."
            )

        # ----------------------------------------------------
        # Date
        # ----------------------------------------------------

        parsed_date = parse_excel_date(
            date_value
        )

        employee_name = (
            str(employee_value)
            .strip()
        )

        records.append({

            "date": parsed_date,

            "status": "Out of Office",

            "employee": employee_name,

            "employee_normalized":
                normalize_name(
                    employee_name
                )

        })

    if not records:

        raise ValueError(
            "No OOO records were found "
            "in the input Excel."
        )

    return records


# ============================================================
# HOME / UPLOAD
#
# Selecting a file automatically submits the form.
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
            "Please upload the OOO input Excel file.",
            "error"
        )

        return redirect(
            url_for("index")
        )

    # --------------------------------------------------------
    # Master template
    # --------------------------------------------------------

    if not os.path.exists(
        MASTER_FILE
    ):

        flash(
            "ORB_OOO_Schedule.xlsx was not found. "
            "Please place it in the same folder as app.py.",
            "error"
        )

        return redirect(
            url_for("index")
        )

    # --------------------------------------------------------
    # Save uploaded input
    # --------------------------------------------------------

    filename = secure_filename(
        uploaded_file.filename
    )

    input_path = os.path.join(
        UPLOAD_FOLDER,
        filename
    )

    uploaded_file.save(
        input_path
    )

    try:

        # ----------------------------------------------------
        # Read input
        # ----------------------------------------------------

        records = read_ooo_input(
            input_path
        )

        # ----------------------------------------------------
        # Generate
        # ----------------------------------------------------

        reports = generate_reports(
            MASTER_FILE,
            records
        )

        # ----------------------------------------------------
        # Render result
        # ----------------------------------------------------

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
            "Report file was not found.",
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
# RUN
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=False
    )
