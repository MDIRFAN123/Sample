"""Flask application entry point for Workbook Automation Tool."""

from __future__ import annotations

from pathlib import Path
from uuid import uuid4

from flask import Flask, jsonify, render_template, request
from werkzeug.utils import secure_filename

from config import Settings
from db import DatabaseConfigurationError, OracleDatabase, TaskNotFoundError
from services.pii_service import PIIScanner
from services.workbook_service import WorkbookError, WorkbookProcessor
from downloader import download_to_local


def create_app(settings: Settings | None = None) -> Flask:
    settings = settings or Settings.from_environment()
    app = Flask(__name__)
    database = OracleDatabase(settings)
    processor = WorkbookProcessor()
    pii_scanner = PIIScanner()

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.get("/health")
    def health():
        return {"status": "ok"}

    @app.post("/scan_pii")
    def scan_pii():
        payload = request.get_json(silent=True) or {}
        prompt = payload.get("prompt", "")
        sections = payload.get("sections")
        worksheet = str(payload.get("worksheet", "")).strip()
        if not isinstance(prompt, str):
            return jsonify(error="prompt must be a string."), 400
        if sections is not None and (
            not isinstance(sections, dict)
            or not all(isinstance(key, str) and isinstance(value, str) for key, value in sections.items())
        ):
            return jsonify(error="sections must be an object containing text values."), 400
        total_length = len(prompt) + sum(len(value) for value in (sections or {}).values())
        if total_length > 1_000_000:
            return jsonify(error="prompt must be 1,000,000 characters or fewer."), 413
        if sections:
            return jsonify(pii_scanner.scan_sections(sections, worksheet))
        return jsonify(pii_scanner.scan(prompt))

    @app.post("/fetch_workbook")
    def fetch_workbook():
        payload = request.get_json(silent=True) or {}
        project_reference = str(payload.get("project_reference", "")).strip()
        task_number = str(payload.get("task_number", "")).strip()
        if not project_reference or not task_number:
            return jsonify(error="project_reference and task_number are required."), 400
        if len(project_reference) > 100 or len(task_number) > 100:
            return jsonify(error="Project reference and task number must be 100 characters or fewer."), 400

        try:
            task = database.get_task(project_reference, task_number)
            if not task.workbook_url or not task.filename:
                return jsonify(
                    status="manual_required",
                    message=(
                        "Workbook download link not found. Please download the "
                        "workbook manually using the task link below."
                    ),
                    task_url=task.task_url,
                    expected_filename=task.filename,
                )
            try:
                workbook_path = download_to_local(task.workbook_url, task.filename)
            except Exception as error:
                app.logger.warning("Automatic workbook download failed: %s", error)
                return jsonify(
                    status="manual_required",
                    message=(
                        "Please download the workbook manually. It will be loaded "
                        "automatically once detected (or prompt for upload if "
                        "auto-detection isn't available)."
                    ),
                    task_url=task.task_url,
                    expected_filename=task.filename,
                )
            payload = processor.build_payload(workbook_path, task.offer_description)
            payload["status"] = "ready"
            payload["upload_source"] = "auto"
            return jsonify(payload)
        except TaskNotFoundError as error:
            return jsonify(error=str(error)), 404
        except (DatabaseConfigurationError, WorkbookError) as error:
            app.logger.warning("Workbook fetch failed: %s", error)
            return jsonify(error=str(error)), 422
        except Exception:
            app.logger.exception("Unexpected workbook fetch failure")
            return jsonify(error="Unable to fetch the workbook. Please contact support if the problem persists."), 500

    @app.post("/lookup_workbook")
    def lookup_workbook():
        payload = request.get_json(silent=True) or {}
        project_reference = str(payload.get("project_reference", "")).strip()
        task_number = str(payload.get("task_number", "")).strip()
        if not project_reference or not task_number:
            return jsonify(error="project_reference and task_number are required."), 400
        try:
            task = database.get_task(project_reference, task_number)
            return jsonify(
                status=(
                    "found"
                    if task.workbook_url and task.filename
                    else "manual_required"
                ),
                task_url=task.task_url,
                filename=task.filename,
                message=(
                    None
                    if task.workbook_url and task.filename
                    else (
                        "Workbook download link not found. Please download the "
                        "workbook manually using the task link below."
                    )
                ),
            )
        except TaskNotFoundError as error:
            return jsonify(error=str(error)), 404
        except DatabaseConfigurationError as error:
            return jsonify(error=str(error)), 422

    @app.post("/upload_workbook")
    def upload_workbook():
        uploaded = request.files.get("workbook")
        project_reference = str(request.form.get("project_reference", "")).strip()
        task_number = str(request.form.get("task_number", "")).strip()
        if not uploaded or not uploaded.filename:
            return jsonify(error="Select a workbook file to upload."), 400

        filename = secure_filename(uploaded.filename)
        extension = Path(filename).suffix.lower()
        if extension not in {".xlsx", ".xlsm"}:
            return jsonify(error="Upload an .xlsx or .xlsm workbook."), 400

        settings.download_dir.mkdir(parents=True, exist_ok=True)
        workbook_path = settings.download_dir / f"manual-{uuid4().hex}-{filename}"

        try:
            uploaded.save(workbook_path)
            offer_description = (
                database.fetch_offer_description(project_reference, task_number)
                if project_reference and task_number
                else ""
            )
            payload = processor.build_payload(workbook_path, offer_description)
            payload["status"] = "ready"
            payload["upload_source"] = "manual"
            return jsonify(payload)
        except (DatabaseConfigurationError, WorkbookError) as error:
            app.logger.warning("Manual workbook upload failed: %s", error)
            return jsonify(error=str(error)), 422
        finally:
            workbook_path.unlink(missing_ok=True)

    return app


if __name__ == "__main__":
    create_app().run(host="127.0.0.1", port=5000, debug=True)
