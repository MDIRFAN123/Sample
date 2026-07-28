"""Oracle access layer for Workfront workbook metadata.

Replace ``APP_SCHEMA`` and the placeholder table/column names below with the
ones from your Oracle database. The column order is intentional: it matches
the tuples consumed by ``fetch_file_info`` and ``fetch_offer_description``.
"""

from __future__ import annotations

from dataclasses import dataclass

from config import Settings


# ---------------------------------------------------------------------------
# SQL templates - replace APP_SCHEMA/table names and join predicates as needed.
# The :project_ref and :task_no names must not be changed without also changing
# the execute calls below.
# ---------------------------------------------------------------------------
WORKBOOK_DOCUMENT_SQL = """
    SELECT
        wd.reference_number,  -- ref
        wd.task_id,
        wd.task_name,
        wd.document_name,    -- doc_name
        wd.download_url,
        wd.filename,
        wd.task_number,
        wd.created_date,
        wd.email_platform
    FROM APP_SCHEMA.WORKBOOK_DOCUMENT wd
    WHERE wd.reference_number = :project_ref
      AND wd.task_number = :task_no
    ORDER BY wd.created_date DESC
    FETCH FIRST 1 ROW ONLY
"""

WORKBOOK_TASK_SQL = """
    SELECT
        wt.reference_number, -- ref
        wt.task_id,
        wt.project_id,
        wt.name,
        wt.task_number,
        wt.email_platform
    FROM APP_SCHEMA.WORKBOOK_TASK wt
    WHERE wt.reference_number = :project_ref
      AND wt.task_number = :task_no
    FETCH FIRST 1 ROW ONLY
"""

DATA_SHARE_SQL = """
    SELECT
        ds.reference_number, -- ref
        ds.task_id,          -- task_id_ds
        ds.project_id,
        ds.name,
        ds.task_number,
        ds.email_platform
    FROM APP_SCHEMA.DATA_SHARE ds
    WHERE ds.reference_number = :project_ref
      AND ds.task_number = :task_no
    FETCH FIRST 1 ROW ONLY
"""

OFFER_DESCRIPTION_SQL = """
    SELECT
        od.reference_number, -- ref
        od.task_id,          -- task_id_ds in the existing result shape
        od.project_id,
        od.name,
        od.task_number,
        od.offer_type,
        od.offer_description
    FROM APP_SCHEMA.OFFER_DESCRIPTION od
    WHERE od.reference_number = :project_ref
      AND od.task_number = :task_no
    FETCH FIRST 1 ROW ONLY
"""


class DatabaseConfigurationError(RuntimeError):
    """Raised when Oracle connection settings are missing."""


class TaskNotFoundError(LookupError):
    """Raised when no workbook-related record is found for a task."""


@dataclass(frozen=True)
class TaskRecord:
    workbook_url: str | None
    filename: str
    offer_description: str = ""
    task_url: str | None = None


class OracleDatabase:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def get_task(self, project_reference: str, task_number: str) -> TaskRecord:
        file_info = self.fetch_file_info(project_reference, task_number)
        if not file_info:
            raise TaskNotFoundError("No workbook record was found for that project reference and task number.")

        # Prefer the direct document download. Fall back to the Workfront
        # version URL when the database only provides a version/task identifier.
        document = file_info.get("workbook_document")
        task = document or file_info.get("workbook_task") or file_info.get("data_share")
        if not task:
            raise TaskNotFoundError("No workbook link is available for this task.")

        workbook_url = task.get("download_url")

        filename = task.get("filename") or ""

        return TaskRecord(
            workbook_url=workbook_url,
            filename=filename,
            task_url=task.get("workbook_link") or task.get("data_share_link"),
            offer_description=self.fetch_offer_description(project_reference, task_number),
        )

    def fetch_file_info(self, project_ref: str, task_no: str) -> dict | None:
        """Return the existing app.py-compatible metadata shape."""
        workbook_document = None
        workbook_task = None
        data_share = None

        with self._connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(WORKBOOK_DOCUMENT_SQL, project_ref=project_ref, task_no=task_no)
                row = cursor.fetchone()
                print("WORKBOOK_DOCUMENT:", row)
                if row:
                    (
                        _ref, task_id, _task_name, doc_name, download_url, filename,
                        _task_number, _created_date, email_platform,
                    ) = row
                    task_id_str = str(task_id) if task_id is not None else None
                    workbook_document = {
                        "source": "WORKBOOK_DOCUMENT",
                        "task_id": task_id_str,
                        "download_url": download_url,
                        "filename": filename or doc_name,
                        "email_platform": email_platform,
                        "workbook_link": self._workbook_link(task_id_str),
                    }

                cursor.execute(WORKBOOK_TASK_SQL, project_ref=project_ref, task_no=task_no)
                row = cursor.fetchone()
                print("WORKBOOK_TASK:", row)
                if row:
                    _ref, task_id, _project_id, _name, _task_number, email_platform = row
                    task_id_str = str(task_id) if task_id is not None else None
                    workbook_task = {
                        "source": "WORKBOOK_TASK",
                        "task_id": task_id_str,
                        "download_url": None,
                        "filename": "",
                        "email_platform": email_platform,
                        "workbook_link": self._workbook_link(task_id_str),
                    }

                cursor.execute(DATA_SHARE_SQL, project_ref=project_ref, task_no=task_no)
                row = cursor.fetchone()
                print("DATA_SHARE:", row)
               
                if row:
                    _ref, task_id_ds, _project_id, _name, _task_number, email_platform = row
                    task_id_ds_str = str(task_id_ds) if task_id_ds is not None else None
                    data_share = {
                        "source": "DATA_SHARE",
                        "task_id_ds": task_id_ds_str,
                        "download_url": None,
                        "filename": "",
                        "email_platform": email_platform,
                        "data_share_link": self._data_share_link(task_id_ds_str),
                    }

        result: dict[str, dict] = {}
        if data_share:
            result["data_share"] = data_share
        if workbook_document:
            result["workbook_document"] = workbook_document
        if workbook_task:
            result["workbook_task"] = workbook_task
        return result or None

    def fetch_offer_description(self, project_ref: str, task_no: str) -> str:
        with self._connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(OFFER_DESCRIPTION_SQL, project_ref=project_ref, task_no=task_no)
                row = cursor.fetchone()
                return str(row[6]) if row and row[6] is not None else ""

    def _connection(self):
        if not (self.settings.ora_user and self.settings.ora_password and self.settings.ora_dsn):
            raise DatabaseConfigurationError("Missing ORA_USER / ORA_PASS / ORA_DSN environment variables.")
        try:
            import oracledb
        except ImportError as error:
            raise DatabaseConfigurationError("Install the oracledb package to use Oracle.") from error
        return oracledb.connect(user=self.settings.ora_user, password=self.settings.ora_password, dsn=self.settings.ora_dsn)

    def _workbook_link(self, version_id: str | None) -> str | None:
        return f"{self.settings.workbook_base_url}{version_id}" if version_id else None

    def _data_share_link(self, task_id: str | None) -> str | None:
        return f"{self.settings.data_share_base_url}{task_id}" if task_id else None
