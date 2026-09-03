"""Oracle access layer for Workfront workbook metadata.

Replace ``APP_SCHEMA`` and the placeholder table/column names below with the
ones from your Oracle database. The column order is intentional: it matches
the tuples consumed by ``fetch_file_info`` and ``fetch_offer_description``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json
import os

# Oracle Thick Mode must be initialized only once per Python process.
_ORACLE_CLIENT_INITIALIZED = False

from config import Settings


# ---------------------------------------------------------------------------
# SQL templates - replace APP_SCHEMA/table names and join predicates as needed.
# The :project_ref and :task_no names must not be changed without also changing
# the execute calls below.
# ---------------------------------------------------------------------------
WORKBOOK_DOCUMENT_SQL = """
    SELECT
        distinct
        proj.REFERENCENUMBER as "ref",
        task.id as task_id,
        task.name as "task_name",
        DOCU.name as doc_name,
        'https://internal/document/download?versionID='
            || DOCU.CURRENTVERSIONID
            || '&ID='
            || DOCU.ID "download_url",
        DOCU.NAME || ' ' || DOCV.EXT "filename",
        TASK.TASKNUMBER as task_number,
        DOCV.created_date,

        (select pv.value255
         from CMKTSCH.WF_PARAMETERVALUES255 pv
         where pv.id = proj.id
           and pv.key = 'DE:Email Platform') as "email_platform"

    from
        CMKTSCH.WF_PROJ proj
        left join CMKTSCH.WF_TASK task
            on proj.id = task.projectid
        LEFT JOIN CMKTSCH.WF_DOCU DOCU
            ON task.ID = DOCU.TASKID
        LEFT JOIN CMKTSCH.WF_DOCV DOCV
            ON DOCU.CURRENTVERSIONID = DOCV.ID

    where
        LASTVERSIONNUM in ('1')

        and proj.REFERENCENUMBER = :project_ref

        and ((:current_task = 1 and TASK.TASKNUMBER = :task_no)
             or (:current_task = 0 and upper(task.name) like '%DEVELOP WORKBOOK%'
                 and TASK.TASKNUMBER = (
                     SELECT MAX(tk.TASKNUMBER)
                     FROM CMKTSCH.WF_TASK tk
                     WHERE tk.projectid = proj.id
                       AND upper(tk.name) LIKE '%DEVELOP WORKBOOK%'
                       AND tk.TASKNUMBER < :task_no
                 )))

        and DOCV.created_date = (
            select max(DOCV.created_date)
            from
                CMKTSCH.WF_PROJ proj
                left join CMKTSCH.WF_TASK task
                    on proj.id = task.projectid
                LEFT JOIN CMKTSCH.WF_DOCU DOCU
                    ON task.ID = DOCU.TASKID
                LEFT JOIN CMKTSCH.WF_DOCV DOCV
                    ON DOCU.CURRENTVERSIONID = DOCV.ID

            where
                LASTVERSIONNUM in ('1')

                and proj.REFERENCENUMBER = :project_ref

                and ((:current_task = 1 and TASK.TASKNUMBER = :task_no)
                     or (:current_task = 0 and upper(task.name) like '%DEVELOP WORKBOOK%'
                         and TASK.TASKNUMBER = (
                             SELECT MAX(tk.TASKNUMBER)
                             FROM CMKTSCH.WF_TASK tk
                             WHERE tk.projectid = proj.id
                               AND upper(tk.name) LIKE '%DEVELOP WORKBOOK%'
                               AND tk.TASKNUMBER < :task_no
                         )))
        )
"""

WORKBOOK_TASK_SQL = """
    select distinct
        prj.REFERENCENUMBER as ref,
        tsk.id as task_id,
        prj.id as project_id,
        tsk.NAME as name,
        tsk.TASKNUMBER as task_number,
        (select pv.value255
         from cmktsch.wf_parameterValues255 pv
         where pv.id = prj.id
           and pv.key = 'DE:Email Platform') as "email_platform"

    FROM cmktsch.wf_task tsk
    inner join cmktsch.wf_proj prj
        on tsk.projectid = prj.id
    where
        (upper(tsk.NAME) like '%DEVELOP WORKBOOK%')
        AND prj.REFERENCENUMBER = :project_ref
        AND EXISTS (
            SELECT 1
            FROM cmktsch.wf_task entered_task
            WHERE entered_task.projectid = prj.id
              AND entered_task.TASKNUMBER = :task_no
        )
        AND tsk.TASKNUMBER = COALESCE(
            (
                SELECT MAX(current_task.TASKNUMBER)
                FROM cmktsch.wf_task current_task
                WHERE current_task.projectid = prj.id
                  AND current_task.TASKNUMBER = :task_no
                  AND upper(current_task.name) LIKE '%DEVELOP WORKBOOK%'
            ),
            (
                SELECT MAX(previous_task.TASKNUMBER)
                FROM cmktsch.wf_task previous_task
                WHERE previous_task.projectid = prj.id
                  AND upper(previous_task.name) LIKE '%DEVELOP WORKBOOK%'
                  AND previous_task.TASKNUMBER < :task_no
            )
        )
"""

DATA_SHARE_SQL = """
    select distinct
        prj.REFERENCENUMBER as ref,
        tsk.id as task_id_ds,
        prj.id as project_id,
        tsk.NAME as name,
        tsk.TASKNUMBER as task_number,
        (select pv.value255
         from cmktsch.wf_parameterValues255 pv
         where pv.id = prj.id
           and pv.key = 'DE:Email Platform') as "email_platform"

    FROM cmktsch.wf_task tsk
    inner join cmktsch.wf_proj prj
        on tsk.projectid = prj.id
    where
        (upper(tsk.NAME) like '%DATA SHARE APPROVAL%')

        AND prj.REFERENCENUMBER = :project_ref
        AND tsk.TASKNUMBER = :task_no
"""

OFFER_DESCRIPTION_SQL = """

    select distinct
        (select pv.value255
         from cmktsch.wf_parameterValues255 pv
         where pv.id = prj.id
           and pv.key = 'DE:Select Offer Type') as "Offer_Type",

        (select pv.value255
         from cmktsch.wf_parameterValues255 pv
         where pv.id = prj.id
           and pv.key = 'DE:Offer Description - details') as "Offer_Description"

    FROM cmktsch.wf_task tsk
    inner join cmktsch.wf_proj prj
        on tsk.projectid = prj.id
    where prj.REFERENCENUMBER = :project_ref
      AND tsk.TASKNUMBER = :task_no
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
                cursor.execute(task_sql, **bind)
                row = cursor.fetchone()
                is_current_task = bool(
                    row and str(row[4]).strip() == str(task_no).strip()
                )
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

                row = None
                if workbook_task:
                    cursor.execute(document_sql, current_task=int(is_current_task), **bind)
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
        """Fetch and decode the DSR BODY while the Oracle connection is open.

        WF_PROJ.BODY may be returned as an Oracle LOB/CLOB.  A LOB belongs to
        the connection that created it, so reading it after the connection
        closes causes DPY-1001 (not connected to database).
        """
        with self._connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(DSR_PROJECT_SQL, refnum=project_reference_number)
                row = cursor.fetchone()

                if not row or row[0] is None:
                    raise TaskNotFoundError(
                        "No DSR data was found for that project reference number."
                    )

                # IMPORTANT: decode/read the CLOB before leaving the
                # connection context.
                return self._decode_json_object(row[0])

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
        if not (
            self.settings.ora_user
            and self.settings.ora_password
            and self.settings.ora_dsn
        ):
            raise DatabaseConfigurationError(
                "Missing ORA_USER / ORA_PASS / ORA_DSN environment variables."
            )

        try:
            import oracledb

            global _ORACLE_CLIENT_INITIALIZED

            # Initialize Thick Mode only once. Do NOT call
            # oracledb.is_thick_mode() because that API is not available in
            # all python-oracledb versions.
            if not _ORACLE_CLIENT_INITIALIZED:
                client_lib_dir = os.getenv("ORACLE_CLIENT_LIB_DIR", "").strip()

                if client_lib_dir:
                    oracledb.init_oracle_client(lib_dir=client_lib_dir)
                else:
                    # Let python-oracledb discover the Oracle Client from the
                    # system environment/PATH, matching the previously
                    # working main.py deployment.
                    oracledb.init_oracle_client()

                _ORACLE_CLIENT_INITIALIZED = True

        except ImportError as error:
            raise DatabaseConfigurationError(
                "Install the oracledb package to use Oracle."
            ) from error
        except Exception as error:
            raise DatabaseConfigurationError(
                f"Unable to initialize Oracle Thick Mode: {error}"
            ) from error

        try:
            return oracledb.connect(
                user=self.settings.ora_user,
                password=self.settings.ora_password,
                dsn=self.settings.ora_dsn,
            )
        except Exception as error:
            raise DatabaseConfigurationError(
                f"Unable to connect to Oracle using ORA_DSN: {self.settings.ora_dsn}"
            ) from error

    def _workbook_link(self, version_id: str | None) -> str | None:
        return f"{self.settings.workbook_base_url}{version_id}" if version_id else None

    def _data_share_link(self, task_id: str | None) -> str | None:
        return f"{self.settings.data_share_base_url}{task_id}" if task_id else None
