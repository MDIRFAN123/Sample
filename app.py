"""Flask application entry point for Workbook Automation Tool."""

from __future__ import annotations

from pathlib import Path
import re
import shutil
import time
import webbrowser
from uuid import uuid4

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from config import Settings
from db import DatabaseConfigurationError, OracleDatabase, TaskNotFoundError
from services.pii_service import PIIScanner
from services.workbook_service import WorkbookError, WorkbookProcessor


def create_app(settings: Settings | None = None) -> Flask:
    settings = settings or Settings.from_environment()
    app = Flask(__name__)
    database = OracleDatabase(settings)
    processor = WorkbookProcessor()
    pii_scanner = PIIScanner()
    workbook_sessions: dict[str, dict] = {}
    monitor_sessions: dict[str, dict] = {}

    def retain_workbook(source: Path) -> tuple[str, Path]:
        session_id = uuid4().hex
        session_dir = settings.download_dir / "sessions" / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        retained = session_dir / source.name
        shutil.copy2(source, retained)
        workbook_sessions[session_id] = {
            "source": retained,
            "temporary": None,
        }
        return session_id, retained

    def build_session_payload(source: Path, offer_description: str, upload_source: str) -> dict:
        session_id, retained = retain_workbook(source)
        result = processor.build_payload(retained, offer_description)
        result.update(
            status="ready",
            upload_source=upload_source,
            workbook_session_id=session_id,
        )
        return result

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
            if not webbrowser.open(task.workbook_url):
                return jsonify(
                    status="manual_required",
                    message=(
                        "The workbook download could not be opened automatically. "
                        "Please use the Workfront task link to download it."
                    ),
                    task_url=task.task_url,
                    expected_filename=task.filename,
                )
            return jsonify(
                status="monitor_required",
                message=(
                    "The workbook download was opened in your browser. "
                    "Monitoring Downloads folder..."
                ),
                task_url=task.task_url,
                expected_filename=task.filename,
            )
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
            return jsonify(build_session_payload(workbook_path, offer_description, "manual"))
        except (DatabaseConfigurationError, WorkbookError) as error:
            app.logger.warning("Manual workbook upload failed: %s", error)
            return jsonify(error=str(error)), 422
        finally:
            workbook_path.unlink(missing_ok=True)

    @app.post("/monitor_workbook/start")
    def start_workbook_monitor():
        payload = request.get_json(silent=True) or {}
        project_reference = str(payload.get("project_reference", "")).strip()
        task_number = str(payload.get("task_number", "")).strip()
        expected_filename = str(payload.get("expected_filename", "")).strip()
        if not project_reference or not task_number:
            return jsonify(error="Project reference and task number are required."), 400
        monitor_id = uuid4().hex
        monitor_sessions[monitor_id] = {
            "project_reference": project_reference,
            "task_number": task_number,
            "expected_filename": Path(expected_filename).name,
            "started_at": time.time(),
            "cancelled": False,
            "detected_path": None,
        }
        return jsonify(
            status="monitoring",
            monitor_id=monitor_id,
            timeout_seconds=settings.download_monitor_timeout,
            directory=str(settings.monitored_download_dir),
            message="Monitoring Downloads folder...",
        )

    @app.post("/monitor_workbook/poll")
    def poll_workbook_monitor():
        payload = request.get_json(silent=True) or {}
        monitor_id = str(payload.get("monitor_id", ""))
        monitor = monitor_sessions.get(monitor_id)
        if not monitor:
            return jsonify(error="Monitoring session was not found."), 404
        if monitor["cancelled"]:
            monitor_sessions.pop(monitor_id, None)
            return jsonify(status="cancelled", message="Workbook monitoring cancelled.")
        if time.time() - monitor["started_at"] >= settings.download_monitor_timeout:
            monitor_sessions.pop(monitor_id, None)
            return jsonify(status="timeout", message="No matching workbook was detected before timeout.")

        if monitor.get("detected_path"):
            detected = Path(monitor["detected_path"])
            try:
                task = database.get_task(
                    monitor["project_reference"],
                    monitor["task_number"],
                )
                result = build_session_payload(detected, task.offer_description, "detected")
                monitor_sessions.pop(monitor_id, None)
                result["message"] = "Workbook uploaded successfully."
                return jsonify(result)
            except (DatabaseConfigurationError, WorkbookError) as error:
                return jsonify(error=str(error)), 422

        directory = settings.monitored_download_dir
        expected = monitor["expected_filename"].lower()
        expected_stem = Path(expected).stem
        project_token = re.sub(r"[^a-z0-9]", "", monitor["project_reference"].lower())
        task_token = re.sub(r"[^a-z0-9]", "", monitor["task_number"].lower())
        candidates: list[Path] = []
        if directory.is_dir():
            for candidate in directory.iterdir():
                if candidate.suffix.lower() not in {".xlsx", ".xlsm"}:
                    continue
                try:
                    if candidate.stat().st_mtime < monitor["started_at"]:
                        continue
                except OSError:
                    continue
                normalized = re.sub(r"[^a-z0-9]", "", candidate.stem.lower())
                filename_matches = bool(
                    expected_stem
                    and (
                        candidate.name.lower() == expected
                        or candidate.stem.lower().startswith(f"{expected_stem} (")
                    )
                )
                identifiers_match = (
                    project_token in normalized
                    and task_token in normalized
                )
                if filename_matches or (not expected and identifiers_match):
                    candidates.append(candidate)

        if not candidates:
            return jsonify(status="monitoring", message="Waiting for workbook download...")

        detected = max(candidates, key=lambda path: path.stat().st_mtime)
        try:
            size_before = detected.stat().st_size
            time.sleep(0.25)
            if detected.stat().st_size != size_before:
                return jsonify(status="monitoring", message="Workbook detected. Waiting for download to finish...")
            monitor["detected_path"] = str(detected)
            return jsonify(
                status="monitoring",
                message="Workbook detected. Uploading workbook...",
            )
        except OSError:
            return jsonify(status="monitoring", message="Waiting for workbook download...")

    @app.post("/monitor_workbook/cancel")
    def cancel_workbook_monitor():
        payload = request.get_json(silent=True) or {}
        monitor = monitor_sessions.get(str(payload.get("monitor_id", "")))
        if monitor:
            monitor["cancelled"] = True
        return jsonify(status="cancelled")

    @app.post("/temporary_workbook")
    def create_temporary_workbook():
        payload = request.get_json(silent=True) or {}
        session_id = str(payload.get("workbook_session_id", ""))
        selected_sheets = payload.get("selected_sheets")
        session = workbook_sessions.get(session_id)
        if not session:
            return jsonify(error="Workbook session was not found."), 404
        if not isinstance(selected_sheets, list) or not selected_sheets:
            return jsonify(error="Select at least one worksheet."), 400
        from openpyxl import load_workbook

        workbook = load_workbook(session["source"], data_only=False)
        try:
            missing = [name for name in selected_sheets if name not in workbook.sheetnames]
            if missing:
                return jsonify(error=f"Unknown worksheet: {missing[0]}"), 400
            for worksheet in list(workbook.worksheets):
                if worksheet.title not in selected_sheets:
                    workbook.remove(worksheet)
            selected_objects = {worksheet.title: worksheet for worksheet in workbook.worksheets}
            workbook._sheets = [selected_objects[name] for name in selected_sheets]
            temporary_name = f"validation-{session_id[:8]}.xlsx"
            temporary_path = session["source"].parent / temporary_name
            workbook.save(temporary_path)
        finally:
            workbook.close()
        prior = session.get("temporary")
        if prior and prior != temporary_path:
            Path(prior).unlink(missing_ok=True)
        session["temporary"] = temporary_path
        return jsonify(
            attachment_id=session_id,
            filename=temporary_name,
            sheets=selected_sheets,
            size=temporary_path.stat().st_size,
            download_url=f"/temporary_workbook/{session_id}",
        )

    @app.get("/temporary_workbook/<session_id>")
    def download_temporary_workbook(session_id: str):
        session = workbook_sessions.get(session_id)
        path = Path(session["temporary"]) if session and session.get("temporary") else None
        if not path or not path.is_file():
            return jsonify(error="Temporary workbook was not found."), 404
        return send_file(path, as_attachment=True, download_name=path.name)

    @app.post("/workbook_session/end")
    def end_workbook_session():
        payload = request.get_json(silent=True) or {}
        session_id = str(payload.get("workbook_session_id", ""))
        session = workbook_sessions.pop(session_id, None)
        if session:
            shutil.rmtree(Path(session["source"]).parent, ignore_errors=True)
        return jsonify(status="ended")

    return app


if __name__ == "__main__":
    create_app().run(host="127.0.0.1", port=5000, debug=True)
