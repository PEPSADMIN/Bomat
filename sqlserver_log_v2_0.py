"""
PEPS BOM Automation Tool — SQL Server Activity Logger
Version: 2.0
Date: 13 Jun 2026

Writes user activity and run/replace/approval history to the SQL Server
BOM database on 192.168.0.133 asynchronously (background thread queue).

If SQL Server is unavailable the main app is NEVER blocked or crashed —
all failures are silently swallowed and printed to the server console.

Install dependency once:
    pip install pyodbc

Usage (from app_v2_1.py):
    from sqlserver_log_v2_0 import activity_log
    activity_log.login('Hari', 'admin', '192.168.0.10')
    activity_log.bom_download('Hari', 'admin', run_id=5, product_count=3, filename='run.zip')
    activity_log.global_replace('Hari', 'admin', snap_id=2, old='RML-A', new='RML-B',
                                 files=12, rows=48, reason='Sourcing change')
"""

import queue
import threading
import traceback
from datetime import datetime
from typing import Optional

# ── Connection configuration ──────────────────────────────────────────────────
_SERVER   = '192.168.0.133'
_DATABASE = 'BOM'
_UID      = 'sa'
_PWD      = 'Admin@123'
_DRIVER   = 'ODBC Driver 17 for SQL Server'   # fallback tries 13 and 18

_CONN_STR_TEMPLATE = (
    "DRIVER={{{driver}}};"
    "SERVER={server};"
    "DATABASE={database};"
    "UID={uid};"
    "PWD={pwd};"
    "TrustServerCertificate=yes;"
    "Connect Timeout=5;"
)

# ── Internal queue + worker ───────────────────────────────────────────────────
_q: queue.Queue = queue.Queue(maxsize=500)   # drop oldest if queue is full
_worker_started = False
_conn_available = None   # None = untested, True/False after first attempt


def _get_conn_str():
    """Try ODBC drivers in order of preference."""
    for driver in [_DRIVER, 'ODBC Driver 18 for SQL Server', 'ODBC Driver 13 for SQL Server',
                   'SQL Server']:
        try:
            import pyodbc
            cs = _CONN_STR_TEMPLATE.format(
                driver=driver, server=_SERVER, database=_DATABASE,
                uid=_UID, pwd=_PWD
            )
            c = pyodbc.connect(cs, timeout=5)
            c.close()
            return cs
        except Exception:
            continue
    return None


_conn_str_cache: Optional[str] = None


def _get_connection():
    global _conn_str_cache
    import pyodbc
    if _conn_str_cache is None:
        _conn_str_cache = _get_conn_str()
    if _conn_str_cache is None:
        raise RuntimeError('No compatible ODBC driver found or SQL Server unreachable')
    return pyodbc.connect(_conn_str_cache, timeout=5)


def _worker():
    """Background daemon thread — drains the queue and executes SQL inserts."""
    global _conn_available
    while True:
        try:
            task = _q.get(timeout=30)
            if task is None:
                break
            sql, params = task
            try:
                conn = _get_connection()
                conn.execute(sql, params)
                conn.commit()
                conn.close()
                if _conn_available is not True:
                    _conn_available = True
                    print('✓ SQL Server activity logger connected to BOM database')
            except Exception as ex:
                if _conn_available is not False:
                    _conn_available = False
                    print(f'⚠ SQL Server activity logger unavailable: {ex}')
        except queue.Empty:
            continue
        except Exception:
            traceback.print_exc()


def _ensure_worker():
    global _worker_started
    if not _worker_started:
        t = threading.Thread(target=_worker, daemon=True, name='sqlserver-log')
        t.start()
        _worker_started = True


def _enqueue(sql: str, params: tuple):
    _ensure_worker()
    try:
        _q.put_nowait((sql, params))
    except queue.Full:
        pass   # queue full → silently drop; never block the main request


# ── Public API ────────────────────────────────────────────────────────────────

class ActivityLogger:
    """Fire-and-forget activity logger to SQL Server BOM database."""

    # ── User events ───────────────────────────────────────────────────────────

    def _log(self, username: str, role: str, activity: str,
             detail: str = '', ip: str = '', session_id: str = ''):
        _enqueue(
            "INSERT INTO BOM_UserActivity "
            "(username, user_role, activity, detail, ip_address, session_id, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (username, role, activity, detail[:1000], ip[:50], session_id[:100],
             datetime.utcnow())
        )

    def login(self, username: str, role: str, ip: str = '', session_id: str = ''):
        self._log(username, role, 'login', f'User logged in from {ip}', ip, session_id)

    def logout(self, username: str, role: str, session_id: str = ''):
        self._log(username, role, 'logout', 'User logged out', session_id=session_id)

    def bom_view(self, username: str, role: str, product_name: str):
        self._log(username, role, 'bom_view', f'Viewed: {product_name}')

    def bom_download(self, username: str, role: str, run_id: int = 0,
                     product_count: int = 0, filename: str = '', rows: int = 0):
        detail = (f'run_id={run_id} | products={product_count} '
                  f'| rows={rows} | file={filename}')
        self._log(username, role, 'bom_download', detail)

    def global_replace(self, username: str, role: str, snap_id: int = 0,
                       old_code: str = '', new_code: str = '',
                       files: int = 0, rows: int = 0, reason: str = ''):
        detail = (f'snap_id={snap_id} | {old_code} → {new_code} '
                  f'| files={files} rows={rows} | reason={reason}')
        self._log(username, role, 'global_replace', detail)

    def nearest_bom(self, username: str, role: str, query: str):
        self._log(username, role, 'nearest_bom', f'Query: {query}')

    def new_product(self, username: str, role: str, product_name: str):
        self._log(username, role, 'new_product', f'Created: {product_name}')

    def approval_submit(self, username: str, role: str, approval_id: int,
                        action_type: str, summary: str):
        detail = f'approval_id={approval_id} | type={action_type} | {summary}'
        self._log(username, role, 'approval_submit', detail)

    def approval_resolve(self, reviewer: str, role: str, approval_id: int,
                         status: str, reason: str = ''):
        detail = f'approval_id={approval_id} | {status} | reason={reason}'
        self._log(reviewer, role, f'approval_{status}', detail)

    # ── Run history ───────────────────────────────────────────────────────────

    def log_run(self, sqlite_run_id: int, run_label: str, status: str,
                product_count: int, rows_generated: int, run_by: str,
                mode: str = 'FULL', output_mode: str = 'ZIP',
                output_filename: str = '', approval_ref: str = '',
                warnings: int = 0, completed_at: Optional[datetime] = None):
        _enqueue(
            "INSERT INTO BOM_RunHistory "
            "(sqlite_run_id, run_label, status, product_count, rows_generated, "
            " run_by, mode, output_mode, output_filename, approval_ref, warnings, "
            " created_at, completed_at) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (sqlite_run_id, run_label[:400], status, product_count, rows_generated,
             run_by[:100], mode, output_mode, output_filename[:400], approval_ref[:200],
             warnings, datetime.utcnow(), completed_at)
        )

    # ── Global replace history ────────────────────────────────────────────────

    def log_replace(self, sqlite_snap_id: int, old_code: str, new_code: str,
                    reason: str, approval_ref: str, run_by: str,
                    files_changed: int, rows_changed: int, status: str = 'active'):
        _enqueue(
            "INSERT INTO BOM_GlobalReplace "
            "(sqlite_snap_id, old_code, new_code, reason, approval_ref, run_by, "
            " files_changed, rows_changed, status, created_at) "
            "VALUES (?,?,?,?,?,?,?,?,?,?)",
            (sqlite_snap_id, old_code[:300], new_code[:300], reason[:1000],
             approval_ref[:200], run_by[:100], files_changed, rows_changed,
             status, datetime.utcnow())
        )

    # ── Approval history ──────────────────────────────────────────────────────

    def log_approval(self, sqlite_approval_id: int, username: str, user_role: str,
                     action_type: str, action_summary: str, status: str = 'pending',
                     reviewer_name: str = '', rejection_reason: str = '',
                     reviewed_at: Optional[datetime] = None):
        _enqueue(
            "INSERT INTO BOM_ApprovalHistory "
            "(sqlite_approval_id, username, user_role, action_type, action_summary, "
            " status, reviewer_name, rejection_reason, created_at, reviewed_at) "
            "VALUES (?,?,?,?,?,?,?,?,?,?)",
            (sqlite_approval_id, username[:100], user_role[:50], action_type[:100],
             action_summary[:1000], status[:50], reviewer_name[:100],
             rejection_reason[:1000], datetime.utcnow(), reviewed_at)
        )


# Singleton instance imported by app_v2_1.py
activity_log = ActivityLogger()
