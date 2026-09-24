            structure["ooo_row"]
        ):
            remove_code(ws.cell(row, col), code)

        # Add employee to Out of Office.
        ooo_cell = ws.cell(structure["ooo_row"], col)
        add_code(ooo_cell, code)
        red_font(ooo_cell)

    output_name = (
        f"ORB_OOO_Schedule_"
        f"{monday.strftime('%Y%m%d')}_"
        f"{friday.strftime('%Y%m%d')}.xlsx"
    )

    output_path = os.path.join(OUTPUT_FOLDER, output_name)
    wb.save(output_path)

    return {
        "filename": output_name,
        "path": output_path,
        "monday": monday,
        "friday": friday,
        "records": week_records
    }


def generate_reports(records):
    weeks = {}

    for record in records:
        monday = monday_of(record["date"])
        weeks.setdefault(monday, []).append(record)

    reports = []

    for monday in sorted(weeks):
        reports.append(
            process_week(
                records,
                monday
            )
        )

    return reports


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html")

    uploaded = request.files.get("ooo_file")

    if not uploaded or not uploaded.filename:
        flash("Please select an Excel file.", "error")
        return redirect(url_for("index"))

    if not os.path.exists(EMPLOYEE_MASTER_FILE):
        flash(
            "ORB_Employee_Master.xlsx is missing. "
            "Place it next to app.py.",
            "error"
        )
        return redirect(url_for("index"))

    if not os.path.exists(REPORT_TEMPLATE_FILE):
        flash(
            "ORB_OOO_Schedule.xlsx is missing. "
            "Place it next to app.py.",
            "error"
        )
        return redirect(url_for("index"))

    filename = secure_filename(uploaded.filename)

    if not filename.lower().endswith((".xlsx", ".xlsm")):
        flash("Please upload an Excel file (.xlsx or .xlsm).", "error")
        return redirect(url_for("index"))

    input_path = os.path.join(UPLOAD_FOLDER, filename)
    uploaded.save(input_path)

    try:
        records = read_input_file(input_path)
        reports = generate_reports(records)

        return render_template(
            "index.html",
            success=True,
            reports=reports
        )

    except Exception as exc:
        flash(str(exc), "error")
        return redirect(url_for("index"))


@app.route("/download/<filename>")
def download(filename):
    safe_name = os.path.basename(filename)
    path = os.path.join(OUTPUT_FOLDER, safe_name)

    if not os.path.exists(path):
        flash("Report file was not found.", "error")
        return redirect(url_for("index"))

    return send_file(path, as_attachment=True)


@app.route("/health")
def health():
    return {"status": "OK"}


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
