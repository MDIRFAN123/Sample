# Workbook Automation Tool

Flask application that looks up a project task in Oracle, downloads its workbook, masks common PII from the displayed sample, and returns workbook/sheet metadata to the existing UI.

## Run locally

1. Create a virtual environment and install dependencies: `python -m venv .venv` then `.venv\\Scripts\\pip install -r requirements.txt`.
2. Set `ORA_USER`, `ORA_PASS`, and `ORA_DSN` in your environment. Update the four SQL templates in `db.py` by replacing `APP_SCHEMA` and the placeholder table/column names with your Oracle schema.
3. Run `python app.py`, then open `http://127.0.0.1:5000`.

The SQL templates use the bind names `:project_ref` and `:task_no`. Keep those bind names when adapting the query bodies.

## API

`POST /fetch_workbook`

```json
{"project_reference": "PRJ-001", "task_number": "42"}
```

The response includes `workbook_name`, `sheets`, `sheet_data`, and `pii_status`, matching the current browser client. `GET /health` is available for a basic readiness check.
