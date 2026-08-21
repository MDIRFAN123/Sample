"""Oracle access layer for Workfront workbook metadata.

Replace ``APP_SCHEMA`` and the placeholder table/column names below with the
ones from your Oracle database. The column order is intentional: it matches
the tuples consumed by ``fetch_file_info`` and ``fetch_offer_description``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json

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

# DSR queries are backend-only and retain the supplied legacy query semantics.
DSR_PROJECT_SQL = "SELECT BODY FROM cmktsch.WF_PROJ WHERE REFERENCENUMBER = :refnum"
DSR_SEGMENTATION_SQL = """
    SELECT b.Client_Brand, b.seg_group, b.Segment_Name,
           b.Segment_Offer_Description, b.Target_Criteria, b.PRODUCT_TYPE,
           b.CHANNEL, b.SEG_LANGUAGE, b.control, b.percent_segment,
           b.commcode, b.Creative_Code, b.CIS_MEMO
    FROM cmktsch.CC_SEG_MATRIX_HEADER a
    RIGHT JOIN cmktsch.CC_SEG_MATRIX_DETAIL b ON a.id = b.header_id
    WHERE a.project_id IN (
        SELECT proj.id FROM cmktsch.WF_PROJ proj
        WHERE proj.last_update_date > sysdate - 365
    )
      AND project_reference_number = :refnum
"""


class DatabaseConfigurationError(RuntimeError):
    """Raised when Oracle connection settings are missing."""


class TaskNotFoundError(LookupError):
    """Raised when no workbook-related record is found for a task."""


@dataclass(frozen=True)
class TaskRecord:
    workbook_url: str | None
    filename: str
    workfront_content: dict[str, object] = field(default_factory=dict)
    task_url: str | None = None


class OracleDatabase:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def get_task(self, project_reference: str, task_number: str | None = None) -> TaskRecord:
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
            workfront_content=self.fetch_workfront_content(project_reference, task_number),
        )

    def fetch_file_info(self, project_ref: str, task_no: str | None = None) -> dict | None:
        """Return the existing app.py-compatible metadata shape."""
        workbook_document = None
        workbook_task = None
        data_share = None

        with self._connection() as connection:
            with connection.cursor() as cursor:
                document_sql = WORKBOOK_DOCUMENT_SQL
                task_sql = WORKBOOK_TASK_SQL
                data_share_sql = DATA_SHARE_SQL
                bind = {"project_ref": project_ref, "task_no": task_no}
                if not task_no:
                    document_sql = document_sql.replace("      AND wd.task_number = :task_no\n", "")
                    task_sql = task_sql.replace("      AND wt.task_number = :task_no\n", "")
                    data_share_sql = data_share_sql.replace("      AND ds.task_number = :task_no\n", "")
                    bind.pop("task_no")
                cursor.execute(document_sql, **bind)
                row = cursor.fetchone()
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

                cursor.execute(task_sql, **bind)
                row = cursor.fetchone()
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

                cursor.execute(data_share_sql, **bind)
                row = cursor.fetchone()
               
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
        content = self.fetch_workfront_content(project_ref, task_no)
        value = content.get("OFFER_DESCRIPTION")
        return str(value) if value is not None else ""

    def fetch_workfront_content(self, project_ref: str, task_no: str | None = None) -> dict[str, object]:
        """Return every query column in database-returned column order."""
        with self._connection() as connection:
            with connection.cursor() as cursor:
                sql = OFFER_DESCRIPTION_SQL
                bind = {"project_ref": project_ref, "task_no": task_no}
                if not task_no:
                    sql = sql.replace("      AND od.task_number = :task_no\n", "")
                    bind.pop("task_no")
                cursor.execute(sql, **bind)
                row = cursor.fetchone()
                if not row:
                    return {}
                columns = [str(description[0]) for description in cursor.description]
                return dict(zip(columns, row))

    def fetch_dsr_project_body(self, project_reference_number: str) -> dict[str, object]:
        with self._connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(DSR_PROJECT_SQL, refnum=project_reference_number)
                row = cursor.fetchone()
        if not row or row[0] is None:
            raise TaskNotFoundError("No DSR data was found for that project reference number.")
        parsed = self._decode_json_object(row[0])
        return parsed

    @staticmethod
    def _decode_json_object(value: object) -> dict[str, object]:
        """Normalize Oracle JSON, text, bytes, or CLOB values to a dictionary."""
        if isinstance(value, dict):
            return value
        if isinstance(value, bytes):
            value = value.decode("utf-8")
        elif not isinstance(value, str):
            reader = getattr(value, "read", None)
            if not callable(reader):
                raise ValueError("The DSR project data has an unsupported format.")
            value = reader()
            if isinstance(value, bytes):
                value = value.decode("utf-8")
        if not isinstance(value, str):
            raise ValueError("The DSR project data has an unsupported format.")
        parsed = json.loads(value)
        if not isinstance(parsed, dict):
            raise ValueError("The DSR project data has an invalid format.")
        return parsed

    def fetch_dsr_segmentation(self, project_reference_number: str) -> list[tuple]:
        with self._connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(DSR_SEGMENTATION_SQL, refnum=project_reference_number)
                return cursor.fetchall()

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
