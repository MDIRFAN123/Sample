import os
import re
from datetime import datetime, timedelta

import pandas as pd
from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from copy import copy


# ============================================================
# CONFIGURATION
# ============================================================

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MASTER_FILE = os.path.join(
    BASE_DIR,
    "ORB_Employee_Master.xlsx"
)

TEMPLATE_FILE = os.path.join(
    BASE_DIR,
    "ORB_OOO_Schedule.xlsx"
)

OUTPUT_FOLDER = os.path.join(
    BASE_DIR,
    "output"
)

os.makedirs(
    OUTPUT_FOLDER,
    exist_ok=True
)


# ============================================================
# EXCEL LAYOUT
# ============================================================

# Template structure:
#
#        B          C          D          E          F
#     Monday     Tuesday    Wednesday   Thursday   Friday
#
# Row 2 = Date
# Row 3 = Online
# Row 4 = Out of Office

DAY_COLUMNS = {
    0: "B",
    1: "C",
    2: "D",
    3: "E",
    4: "F"
}

DATE_ROW = 2
ONLINE_ROW = 3
OOO_ROW = 4


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def normalize(value):
    """
    Normalize text for reliable comparison.
    """

    if value is None:
        return ""

    value = str(value).strip().lower()

    # Remove extra spaces
    value = re.sub(r"\s+", " ", value)

    return value


def clean_header(value):
    """
    Normalize Excel column headers.
    """

    if value is None:
        return ""

    value = str(value).strip().lower()

    value = value.replace("_", " ")
    value = value.replace("-", " ")

    value = re.sub(r"\s+", " ", value)

    return value


# ============================================================
# READ EMPLOYEE MASTER
# ============================================================

def read_employee_master():

    if not os.path.exists(MASTER_FILE):

        raise FileNotFoundError(
            "ORB_Employee_Master.xlsx was not found."
        )

    df = pd.read_excel(
        MASTER_FILE,
        dtype=str
    )

    # Clean headers
    df.columns = [
        clean_header(column)
        for column in df.columns
    ]

    # Expected:
    # Employee Name
    # Employee Code

    name_column = None
    code_column = None

    for column in df.columns:

        if column in [
            "employee name",
            "name",
            "employee"
        ]:
            name_column = column

        if column in [
            "employee code",
            "code",
            "employee id"
        ]:
            code_column = column

    if not name_column:
        raise ValueError(
            "Employee Name column not found "
            "in ORB_Employee_Master.xlsx."
        )

    if not code_column:
        raise ValueError(
            "Employee Code column not found "
            "in ORB_Employee_Master.xlsx."
        )

    employee_map = {}

    for _, row in df.iterrows():

        name = normalize(
            row[name_column]
        )

        code = str(
            row[code_column]
        ).strip()

        if (
            name
            and code
            and code.lower() != "nan"
        ):

            employee_map[name] = code

    if not employee_map:

        raise ValueError(
            "No employees found in "
            "ORB_Employee_Master.xlsx."
        )

    return employee_map


# ============================================================
# READ OOO INPUT
# ============================================================

def read_ooo_file(file):

    df = pd.read_excel(
        file,
        dtype=str
    )

    if df.empty:

        raise ValueError(
            "The uploaded Excel file is empty."
        )

    # Clean column names
    df.columns = [
        clean_header(column)
        for column in df.columns
    ]

    # --------------------------------------------------------
    # Find required columns
    # --------------------------------------------------------

    date_column = None
    status_column = None
    employee_column = None

    for column in df.columns:

        if column in [
            "date",
            "ooo date"
        ]:
            date_column = column

        elif column in [
            "status",
            "ooo status"
        ]:
            status_column = column

        elif column in [
            "employee name",
            "employee",
            "name"
        ]:
            employee_column = column

    if not date_column:
        raise ValueError(
            "Date column is required."
        )

    if not status_column:
        raise ValueError(
            "Status column is required."
        )

    if not employee_column:
        raise ValueError(
            "Employee Name column is required."
        )

    # --------------------------------------------------------
    # Keep only the three required fields
    # --------------------------------------------------------

    df = df[
        [
            date_column,
            status_column,
            employee_column
        ]
    ].copy()

    df.columns = [
        "date",
        "status",
        "employee"
    ]

    # --------------------------------------------------------
    # Convert dates
    # --------------------------------------------------------

    df["date"] = pd.to_datetime(
        df["date"],
        errors="coerce"
    )

    # Remove invalid dates
    df = df[
        df["date"].notna()
    ]

    # --------------------------------------------------------
    # Only process Out of Office records
    # --------------------------------------------------------

    valid_statuses = [
        "ooo",
        "out of office",
        "out-of-office",
        "out_of_office",
        "outoffice"
    ]

    df["status_normalized"] = (
        df["status"]
        .fillna("")
        .astype(str)
        .apply(normalize)
    )

    df = df[
        df["status_normalized"].isin(
            valid_statuses
        )
    ].copy()

    if df.empty:

        raise ValueError(
            "No Out of Office records were "
            "found in the uploaded file."
        )

    return df


# ============================================================
# GET MONDAY OF WEEK
# ============================================================

def get_monday(date_value):

    return date_value - timedelta(
        days=date_value.weekday()
    )


# ============================================================
# GENERATE REPORT
# ============================================================

def generate_report(df):

    # --------------------------------------------------------
    # Read Employee Master
    # --------------------------------------------------------

    employee_map = read_employee_master()

    print("\n==============================")
    print("EMPLOYEE MASTER")
    print("==============================")

    print(employee_map)

    # --------------------------------------------------------
    # Determine report week
    #
    # The earliest date in the uploaded file
    # determines the week.
    # --------------------------------------------------------

    first_date = df["date"].min().date()

    monday = get_monday(
        first_date
    )

    friday = monday + timedelta(
        days=4
    )

    print("\nREPORT WEEK:")
    print(monday, "to", friday)

    # --------------------------------------------------------
    # Create OOO map
    #
    # {
    #    date: [employee codes]
    # }
    # --------------------------------------------------------

    ooo_by_date = {}

    unknown_employees = []

    for _, row in df.iterrows():

        current_date = (
            row["date"].date()
        )

        employee_name = normalize(
            row["employee"]
        )

        # Ignore records outside
        # Monday-Friday of report week

        if not (
            monday
            <= current_date
            <= friday
        ):
            continue

        if employee_name not in employee_map:

            unknown_employees.append(
                row["employee"]
            )

            continue

        employee_code = employee_map[
            employee_name
        ]

        if current_date not in ooo_by_date:

            ooo_by_date[
                current_date
            ] = []

        if employee_code not in (
            ooo_by_date[current_date]
        ):

            ooo_by_date[
                current_date
            ].append(
                employee_code
            )

    # --------------------------------------------------------
    # Unknown employee check
    # --------------------------------------------------------

    if unknown_employees:

        unknown_list = ", ".join(
            sorted(
                set(
                    str(x)
                    for x in unknown_employees
                )
            )
        )

        raise ValueError(
            "The following employee(s) are "
            "not present in "
            "ORB_Employee_Master.xlsx: "
            + unknown_list
        )

    # --------------------------------------------------------
    # Load template
    # --------------------------------------------------------

    if not os.path.exists(
        TEMPLATE_FILE
    ):

        raise FileNotFoundError(
            "ORB_OOO_Schedule.xlsx was not found."
        )

    wb = load_workbook(
        TEMPLATE_FILE
    )

    ws = wb.active

    # --------------------------------------------------------
    # Get ALL employees
    #
    # IMPORTANT:
    # This is what drives Online.
    # --------------------------------------------------------

    all_employee_codes = list(
        employee_map.values()
    )

    print("\nALL EMPLOYEES:")
    print(all_employee_codes)

    print("\nOOO BY DATE:")
    print(ooo_by_date)

    # --------------------------------------------------------
    # Generate Monday-Friday
    # --------------------------------------------------------

    for day_number in range(5):

        current_date = (
            monday
            + timedelta(days=day_number)
        )

        column = DAY_COLUMNS[
            day_number
        ]

        # ====================================================
        # DATE
        # ====================================================

        date_cell = ws[
            f"{column}{DATE_ROW}"
        ]

        date_cell.value = (
            f"{current_date.month}/"
            f"{current_date.day}"
        )

        # ====================================================
        # OOO
        # ====================================================

        ooo_codes = ooo_by_date.get(
            current_date,
            []
        )

        # ====================================================
        # ONLINE
        #
        # Everyone from Employee Master
        # EXCEPT employees OOO on this date.
        # ====================================================

        online_codes = [
            code
            for code in all_employee_codes
            if code not in ooo_codes
        ]

        print(
            "\nDATE:",
            current_date
        )

        print(
            "OOO:",
            ooo_codes
        )

        print(
            "ONLINE:",
            online_codes
        )

        # ====================================================
        # ONLINE DISPLAY
        #
        # Split across two lines like reference.
        # ====================================================

        if len(online_codes) <= 3:

            online_text = "/".join(
                online_codes
            )

        else:

            midpoint = (
                len(online_codes) + 1
            ) // 2

            first_line = online_codes[
                :midpoint
            ]

            second_line = online_codes[
                midpoint:
            ]

            online_text = "/".join(
                first_line
            )

            if second_line:

                online_text += (
                    "\n"
                    + "/".join(
                        second_line
                    )
                )

        # ====================================================
        # WRITE ONLINE
        # ====================================================

        online_cell = ws[
            f"{column}{ONLINE_ROW}"
        ]

        online_cell.value = online_text

        # Preserve existing formatting
        # but enable wrapping.

        existing_alignment = copy(
            online_cell.alignment
        )

        online_cell.alignment = Alignment(
            horizontal=existing_alignment.horizontal
            or "left",

            vertical=existing_alignment.vertical
            or "center",

            wrap_text=True
        )

        # ====================================================
        # WRITE OOO
        # ====================================================

        ooo_cell = ws[
            f"{column}{OOO_ROW}"
        ]

        ooo_cell.value = "/".join(
            ooo_codes
        )

        # ====================================================
        # OOO RED FONT
        # ====================================================

        if ooo_codes:

            existing_font = copy(
                ooo_cell.font
            )

            ooo_cell.font = Font(
                name=existing_font.name,
                size=existing_font.size,
                bold=True,
                italic=existing_font.italic,
                color="C00000"
            )

        # ====================================================
        # If no OOO, keep cell blank
        # ====================================================

        else:

            ooo_cell.value = ""

    # ========================================================
    # SAVE OUTPUT
    # ========================================================

    filename = (
        "ORB_OOO_Schedule_"
        + monday.strftime("%m%d%Y")
        + "_"
        + friday.strftime("%m%d%Y")
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
# HOME PAGE
# ============================================================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )


# ============================================================
# UPLOAD
# ============================================================

@app.route(
    "/upload",
    methods=["POST"]
)
def upload():

    try:

        if "file" not in request.files:

            return render_template(
                "index.html",
                error="Please select an Excel file."
            )

        file = request.files["file"]

        if not file.filename:

            return render_template(
                "index.html",
                error="Please select an Excel file."
            )

        # ----------------------------------------------------
        # Read input
        # ----------------------------------------------------

        df = read_ooo_file(
            file
        )

        # ----------------------------------------------------
        # Generate report
        # ----------------------------------------------------

        output_path = generate_report(
            df
        )

        # ----------------------------------------------------
        # Automatically return generated file
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
            "\nERROR:",
            str(e)
        )

        return render_template(
            "index.html",
            error=str(e)
        )


# ============================================================
# RUN APPLICATION
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
