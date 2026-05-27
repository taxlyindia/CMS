"""
database.py — Dual-mode database backend
  • PostgreSQL (psycopg2) when DATABASE_URL is set  →  production (Hostinger)
  • SQLite                                           →  local development fallback
TaxlyCMS — Companies Act 2013 Compliance CRM
"""
import os, hashlib, uuid, json
from datetime import datetime, date, timedelta
from pathlib import Path

BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "data" / "uploads"

# ── Detect mode ───────────────────────────────────────────────────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", "")
# Only use PostgreSQL when the URL is actually a Postgres URL
_url_lower = DATABASE_URL.lower()
USE_POSTGRES  = bool(DATABASE_URL) and (
    _url_lower.startswith("postgres") or
    _url_lower.startswith("postgresql")
)

try:
    import psycopg2, psycopg2.extras
    _PG_AVAILABLE = True
except ImportError:
    _PG_AVAILABLE = False

if USE_POSTGRES and not _PG_AVAILABLE:
    raise RuntimeError(
        "DATABASE_URL is set but psycopg2 is not installed.\n"
        "Run: pip install psycopg2-binary"
    )

# ── SQLite path (dev only) ────────────────────────────────────────────────────
DB_PATH = BASE_DIR / "data" / "compli.db"


# ══════════════════════════════════════════════════════════════════════════════
# PostgreSQL connection
# ══════════════════════════════════════════════════════════════════════════════
import logging as _log
_logger = _log.getLogger(__name__)

# ── Build connection kwargs once ────────────────────────────────────────────
def _pg_kwargs():
    from urllib.parse import urlparse
    p = urlparse(DATABASE_URL)
    return dict(
        host     = p.hostname,
        port     = p.port or 5432,
        dbname   = p.path.lstrip("/"),
        user     = p.username,
        password = p.password,
        sslmode  = os.environ.get("DB_SSLMODE", "require"),
        # Keep-alive: detects dead connections before we try to use them
        keepalives          = 1,
        keepalives_idle     = 30,   # seconds idle before first keepalive probe
        keepalives_interval = 10,   # seconds between probes
        keepalives_count    = 5,    # drop after 5 failed probes
        connect_timeout     = 10,   # fail fast instead of hanging
    )


# ── Connection health check ─────────────────────────────────────────────────
def _conn_is_alive(conn) -> bool:
    """
    Cheaply verify that a pooled connection is still usable.
    Catches the two errors shown in the screenshot:
      • psycopg2.OperationalError: SSL error: decryption failed or bad record mac
      • psycopg2.OperationalError: SSL SYSCALL error: EOF detected
    """
    try:
        conn.poll()   # non-blocking check — raises if connection is dead
        # Also do a lightweight round-trip to catch silent TCP drops
        with conn.cursor() as cur:
            cur.execute("SELECT 1")
        return True
    except Exception:
        return False


# ── Optional connection pool (disable with DB_USE_POOL=0) ──────────────────
_USE_POOL = os.environ.get("DB_USE_POOL", "1") == "1"
_PG_POOL  = None
_POOL_LOCK = None

def _get_pg_pool():
    """
    Returns a ThreadedConnectionPool, or None if pool is disabled/failed.
    Pool is created lazily and protected by a threading.Lock so it is
    initialised exactly once even under concurrent gunicorn workers.
    """
    global _PG_POOL, _POOL_LOCK
    if not _USE_POOL:
        return None

    # Initialise the lock on first call (avoids import-time threading overhead)
    if _POOL_LOCK is None:
        import threading
        _POOL_LOCK = threading.Lock()

    if _PG_POOL is not None:
        return _PG_POOL

    with _POOL_LOCK:
        if _PG_POOL is not None:     # double-checked locking
            return _PG_POOL
        try:
            from psycopg2 import pool as _pg_pool
            _PG_POOL = _pg_pool.ThreadedConnectionPool(
                minconn = int(os.environ.get("DB_POOL_MIN", "1")),
                maxconn = int(os.environ.get("DB_POOL_MAX", "8")),
                **_pg_kwargs(),
            )
            _logger.info("PostgreSQL connection pool created (min=%s max=%s)",
                         os.environ.get("DB_POOL_MIN","1"),
                         os.environ.get("DB_POOL_MAX","8"))
        except Exception as e:
            _logger.warning("Connection pool init failed (%s) — using per-request connections", e)
            _PG_POOL = None
    return _PG_POOL


def _pg_conn():
    """
    Obtain a live PostgreSQL connection.
    If the pool is enabled, pop a connection and health-check it.
    If the health-check fails, discard the stale connection and open a fresh one.
    Falls back to per-request connections when the pool is unavailable.
    """
    pool = _get_pg_pool()
    if pool:
        try:
            conn = pool.getconn()
            if conn is None:
                raise Exception("Pool returned None")
            # ── Health check: discard stale / dead SSL connections ─────────
            if not _conn_is_alive(conn):
                _logger.debug("Discarding stale pooled connection, opening fresh one")
                try:
                    pool.putconn(conn, close=True)   # close=True destroys it
                except Exception:
                    pass
                # Open a fresh connection outside the pool for this request
                conn = psycopg2.connect(**_pg_kwargs())
            return conn
        except Exception as e:
            _logger.warning("Pool.getconn failed (%s), falling back to direct connect", e)

    # Direct connection (no pool, or pool exhausted/failed)
    return psycopg2.connect(**_pg_kwargs())


class PGConnection:
    """
    Thin wrapper that makes psycopg2 behave like sqlite3 for our row/rows helpers.
    On close():
      - If conn came from the pool and is healthy → return to pool (putconn)
      - If conn is broken → discard with close=True so the pool replaces it
      - If conn was a direct connection → just close()
    """
    def __init__(self, conn):
        self._conn  = conn
        self._pool  = _get_pg_pool()

    def cursor(self):
        return self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

    def commit(self):
        self._conn.commit()

    def rollback(self):
        try:
            self._conn.rollback()
        except Exception:
            pass

    def close(self):
        if self._pool:
            try:
                # Only return healthy connections to the pool
                if _conn_is_alive(self._conn):
                    self._pool.putconn(self._conn)
                else:
                    _logger.debug("Closing broken connection (not returned to pool)")
                    self._pool.putconn(self._conn, close=True)
                return
            except Exception as e:
                _logger.debug("putconn failed (%s), closing directly", e)
        try:
            self._conn.close()
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════════════════
# SQLite connection (dev fallback)
# ══════════════════════════════════════════════════════════════════════════════
class SQLiteConnection:
    """Thin wrapper around sqlite3 that returns dict rows — same API as PGConnection."""
    def __init__(self, conn):
        self._conn = conn
    def cursor(self):
        self._conn.row_factory = _sqlite3().Row
        c = self._conn.cursor()
        return _DictCursor(c)
    def commit(self):   self._conn.commit()
    def rollback(self): self._conn.rollback()
    def close(self):    self._conn.close()


def _sqlite3():
    import sqlite3
    return sqlite3


class _DictCursor:
    """Wraps sqlite3 cursor to return plain dicts (like RealDictCursor)."""
    def __init__(self, cur):
        self._cur = cur
        self.description = None
        self.rowcount = 0
    def execute(self, sql, params=None):
        # Convert %s → ? for SQLite
        sql = sql.replace('%s', '?')
        # Convert PostgreSQL-specific syntax → SQLite
        sql = sql.replace('NOW()', "datetime('now')")
        sql = sql.replace('CURRENT_DATE', "date('now')")
        sql = sql.replace('ILIKE', 'LIKE')
        # Remove PostgreSQL-only clauses before simple INSERT OR IGNORE
        import re as _re
        sql = _re.sub(r'ON CONFLICT[^;]*DO NOTHING', '', sql)
        sql = _re.sub(r'ON CONFLICT[^;]*DO UPDATE[^;]*', '', sql)
        sql = _re.sub(r'::\w+', '', sql)   # ::text, ::jsonb casts
        # Convert INSERT INTO → INSERT OR IGNORE INTO for upsert behaviour
        if sql.strip().upper().startswith('INSERT INTO'):
            sql = sql.replace('INSERT INTO', 'INSERT OR IGNORE INTO', 1)
        if params:
            # Flatten RealDict params if needed
            self._cur.execute(sql, params)
        else:
            self._cur.execute(sql)
        self.description = self._cur.description or []
        self.rowcount = self._cur.rowcount or 0
        return self
    def fetchone(self):
        r = self._cur.fetchone()
        return dict(r) if r else None
    def fetchall(self):
        return [dict(r) for r in (self._cur.fetchall() or [])]


def _sqlite_conn():
    import sqlite3
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return SQLiteConnection(conn)


# ══════════════════════════════════════════════════════════════════════════════
# Public API
# ══════════════════════════════════════════════════════════════════════════════
def get_db():
    """
    Returns a connection wrapper.
    For PostgreSQL: always autocommit=False so callers explicitly commit/rollback.
    """
    if USE_POSTGRES:
        conn = _pg_conn()
        try:
            conn.autocommit = False   # explicit transaction control
        except Exception:
            pass   # already set or connection is being health-checked
        return PGConnection(conn)
    return _sqlite_conn()


def _serialize_val(v):
    """Convert PostgreSQL-specific types to JSON-serializable Python types."""
    from datetime import date, datetime
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.isoformat()
    if isinstance(v, date):
        return v.isoformat()            # "2025-12-11"
    try:
        from decimal import Decimal
        if isinstance(v, Decimal):
            return float(v)
    except ImportError:
        pass
    return v

def _serialize_dict(d):
    if d is None: return None
    return {k: _serialize_val(v) for k, v in d.items()}

def row(r):
    if r is None: return None
    return _serialize_dict(dict(r) if not isinstance(r, dict) else r)

def rows(rs):
    return [_serialize_dict(dict(r) if not isinstance(r, dict) else r) for r in (rs or [])]

def hash_pw(pw):
    return hashlib.sha256(pw.encode()).hexdigest()


# ══════════════════════════════════════════════════════════════════════════════
# Schema — PostgreSQL (production)
# ══════════════════════════════════════════════════════════════════════════════
SCHEMA_PG = """
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

CREATE TABLE IF NOT EXISTS tenants (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, slug TEXT UNIQUE NOT NULL,
    email TEXT NOT NULL, phone TEXT DEFAULT '', address TEXT DEFAULT '',
    plan_id TEXT, status TEXT DEFAULT 'pending', approved_by TEXT,
    approved_at TIMESTAMPTZ, trial_ends DATE, billing_cycle TEXT DEFAULT 'monthly',
    max_users INTEGER DEFAULT 5, max_companies INTEGER DEFAULT 5,
    notes TEXT DEFAULT '', created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS plans (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, description TEXT DEFAULT '',
    price_monthly NUMERIC(12,2) DEFAULT 0, price_annual NUMERIC(12,2) DEFAULT 0,
    max_users INTEGER DEFAULT 5, max_companies INTEGER DEFAULT 5,
    max_documents INTEGER DEFAULT 100, features JSONB DEFAULT '[]',
    is_active INTEGER DEFAULT 1, created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS users (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, email TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'staff',
    is_active INTEGER DEFAULT 1, is_platform_admin INTEGER DEFAULT 0,
    tenant_id TEXT REFERENCES tenants(id) ON DELETE SET NULL,
    created_at TIMESTAMPTZ DEFAULT NOW(), last_login TIMESTAMPTZ
);
CREATE TABLE IF NOT EXISTS companies (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, cin TEXT UNIQUE,
    incorporation_date DATE, registered_office TEXT, pan TEXT, tan TEXT,
    email TEXT, phone TEXT, authorized_capital NUMERIC(20,2) DEFAULT 0,
    paid_up_capital NUMERIC(20,2) DEFAULT 0, business_activity TEXT,
    company_type TEXT DEFAULT 'Private Limited', roc TEXT, category TEXT,
    sub_category TEXT, status TEXT DEFAULT 'active', letterhead_logo TEXT,
    letterhead_address TEXT, letterhead_footer TEXT, gstin TEXT, website TEXT,
    tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(), updated_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS person_directory (
    id TEXT PRIMARY KEY, tenant_id TEXT REFERENCES tenants(id),
    name TEXT NOT NULL, pan TEXT, din TEXT, aadhaar TEXT,
    email TEXT, mobile TEXT, address TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_person_pan ON person_directory(pan,tenant_id) WHERE pan IS NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS ux_person_din ON person_directory(din,tenant_id) WHERE din IS NOT NULL;

CREATE TABLE IF NOT EXISTS directors (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    person_id TEXT REFERENCES person_directory(id) ON DELETE SET NULL,
    name TEXT NOT NULL, din TEXT, pan TEXT, aadhaar TEXT, email TEXT, mobile TEXT,
    address TEXT, designation TEXT DEFAULT 'Director', date_of_appointment DATE,
    date_of_cessation DATE, mca_user_id TEXT, mca_password TEXT, mca_notes TEXT,
    is_active INTEGER DEFAULT 1, tenant_id TEXT REFERENCES tenants(id),
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS director_kyc (
    id TEXT PRIMARY KEY, director_id TEXT NOT NULL UNIQUE REFERENCES directors(id) ON DELETE CASCADE,
    last_kyc_date DATE, next_due_date DATE, kyc_status TEXT DEFAULT 'pending',
    notes TEXT, tenant_id TEXT REFERENCES tenants(id), updated_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS auditors (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    name TEXT NOT NULL, firm_name TEXT, membership_no TEXT, frn TEXT, pan TEXT,
    address TEXT, email TEXT, phone TEXT, appointment_date DATE,
    nature_of_appointment TEXT DEFAULT 'Regular Auditor',
    appointment_type TEXT DEFAULT 'AGM Appointment',
    start_date DATE, end_date DATE, srn_adt1 TEXT, adt1_file TEXT,
    is_active INTEGER DEFAULT 1, notes TEXT,
    tenant_id TEXT REFERENCES tenants(id), created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS shareholders (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    person_id TEXT REFERENCES person_directory(id) ON DELETE SET NULL,
    name TEXT NOT NULL, folio_no TEXT, pan TEXT, email TEXT, mobile TEXT, address TEXT,
    share_class TEXT DEFAULT 'Equity', shares_held INTEGER DEFAULT 0,
    face_value NUMERIC(10,2) DEFAULT 10, date_of_entry DATE,
    mca_user_id TEXT, mca_password TEXT, is_active INTEGER DEFAULT 1,
    tenant_id TEXT REFERENCES tenants(id), created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS dsc_records (
    id TEXT PRIMARY KEY, company_id TEXT REFERENCES companies(id) ON DELETE SET NULL,
    director_id TEXT REFERENCES directors(id) ON DELETE SET NULL,
    holder_name TEXT NOT NULL, holder_type TEXT DEFAULT 'Director',
    dsc_class TEXT DEFAULT 'Class 3', issued_by TEXT,
    valid_from DATE, valid_to DATE, token_type TEXT,
    custody_status TEXT DEFAULT 'With Client', custody_date DATE, custody_notes TEXT,
    alert_30day_sent INTEGER DEFAULT 0, is_active INTEGER DEFAULT 1, notes TEXT,
    tenant_id TEXT REFERENCES tenants(id), created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS dsc_custody_log (
    id TEXT PRIMARY KEY, dsc_id TEXT NOT NULL REFERENCES dsc_records(id) ON DELETE CASCADE,
    action TEXT NOT NULL, action_date DATE, from_party TEXT, to_party TEXT,
    notes TEXT, recorded_by TEXT, tenant_id TEXT REFERENCES tenants(id),
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS meetings (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    meeting_type TEXT NOT NULL, meeting_no TEXT, meeting_date DATE NOT NULL,
    meeting_time TEXT, venue TEXT, agenda TEXT, notes TEXT,
    minutes_drafted INTEGER DEFAULT 0, status TEXT DEFAULT 'scheduled',
    tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS tasks (
    id TEXT PRIMARY KEY, company_id TEXT REFERENCES companies(id) ON DELETE SET NULL,
    title TEXT NOT NULL, description TEXT, assigned_to TEXT REFERENCES users(id) ON DELETE SET NULL,
    due_date DATE, priority TEXT DEFAULT 'medium', status TEXT DEFAULT 'pending',
    module TEXT, entity_id TEXT, created_by TEXT, completed_at TIMESTAMPTZ,
    task_leader TEXT REFERENCES users(id) ON DELETE SET NULL,
    task_manager TEXT REFERENCES users(id) ON DELETE SET NULL,
    team_members JSONB DEFAULT '[]', note TEXT DEFAULT '',
    billable INTEGER DEFAULT 0, estimated_hrs NUMERIC(6,2) DEFAULT 0,
    actual_hrs NUMERIC(6,2) DEFAULT 0, tenant_id TEXT REFERENCES tenants(id),
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS alerts (
    id TEXT PRIMARY KEY, company_id TEXT REFERENCES companies(id) ON DELETE CASCADE,
    entity_type TEXT NOT NULL, entity_id TEXT NOT NULL,
    alert_type TEXT NOT NULL, title TEXT NOT NULL, message TEXT,
    due_date DATE, severity TEXT DEFAULT 'medium', status TEXT DEFAULT 'active',
    tenant_id TEXT REFERENCES tenants(id), created_at TIMESTAMPTZ DEFAULT NOW(),
    resolved_at TIMESTAMPTZ
);
CREATE TABLE IF NOT EXISTS charges (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    charge_id TEXT, srn TEXT, date_of_creation DATE, amount NUMERIC(20,2) DEFAULT 0,
    charge_holder TEXT, assets_charged TEXT, charge_type TEXT DEFAULT 'Hypothecation',
    date_of_modification DATE, date_of_satisfaction DATE, status TEXT DEFAULT 'Open',
    remarks TEXT, tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS investments (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    investee_name TEXT NOT NULL, investee_type TEXT DEFAULT 'Company',
    investment_type TEXT DEFAULT 'Equity Shares', amount NUMERIC(20,2) DEFAULT 0,
    date_of_investment DATE, board_resolution_date DATE, srn_mgb4 TEXT, purpose TEXT,
    remarks TEXT, tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS loans_guarantees (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    party_name TEXT NOT NULL, party_type TEXT DEFAULT 'Company',
    transaction_type TEXT DEFAULT 'Loan', amount NUMERIC(20,2) DEFAULT 0,
    date_of_transaction DATE, rate_of_interest NUMERIC(6,2) DEFAULT 0,
    repayment_date DATE, security TEXT, board_resolution_date DATE,
    outstanding_amount NUMERIC(20,2) DEFAULT 0, status TEXT DEFAULT 'Active',
    remarks TEXT, tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS related_party_transactions (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    party_name TEXT NOT NULL, relationship TEXT, nature_of_transaction TEXT,
    amount NUMERIC(20,2) DEFAULT 0, date_of_transaction DATE,
    date_of_board_approval DATE, date_of_shareholders_approval DATE,
    terms TEXT, justification TEXT, remarks TEXT,
    tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS director_interests (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    director_id TEXT, director_name TEXT NOT NULL, din TEXT,
    entity_name TEXT NOT NULL, entity_type TEXT DEFAULT 'Company',
    nature_of_interest TEXT, date_of_disclosure DATE, date_of_board_resolution DATE,
    remarks TEXT, tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS esop_grants (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    employee_name TEXT NOT NULL, designation TEXT, employee_id TEXT, grant_date DATE,
    options_granted INTEGER DEFAULT 0, exercise_price NUMERIC(10,2) DEFAULT 0,
    vesting_date DATE, vesting_period TEXT, options_exercised INTEGER DEFAULT 0,
    options_lapsed INTEGER DEFAULT 0, status TEXT DEFAULT 'Active', remarks TEXT,
    tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS statutory_registers (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    register_type TEXT NOT NULL, register_name TEXT NOT NULL, section_ref TEXT,
    entry_date DATE, folio_ref TEXT, entry_data TEXT, remarks TEXT, recorded_by TEXT,
    tenant_id TEXT REFERENCES tenants(id), created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS document_templates (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, category TEXT NOT NULL, description TEXT,
    template_body TEXT NOT NULL, placeholders JSONB DEFAULT '[]',
    is_system INTEGER DEFAULT 0, is_active INTEGER DEFAULT 1,
    tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(), updated_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS documents (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    template_id TEXT REFERENCES document_templates(id) ON DELETE SET NULL,
    doc_type TEXT NOT NULL, doc_name TEXT NOT NULL, content TEXT, file_path TEXT,
    module TEXT, entity_id TEXT, date_from TEXT DEFAULT '', date_to TEXT DEFAULT '',
    tenant_id TEXT REFERENCES tenants(id), generated_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS custom_placeholders (
    id TEXT PRIMARY KEY, name TEXT UNIQUE NOT NULL, label TEXT NOT NULL,
    description TEXT DEFAULT '', ph_type TEXT DEFAULT 'text',
    tenant_id TEXT REFERENCES tenants(id), created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS user_permissions (
    id TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    module TEXT NOT NULL, action TEXT NOT NULL, granted INTEGER NOT NULL DEFAULT 1,
    granted_by TEXT, granted_at TIMESTAMPTZ DEFAULT NOW(), note TEXT DEFAULT '',
    UNIQUE(user_id, module, action)
);
CREATE TABLE IF NOT EXISTS permission_presets (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, description TEXT DEFAULT '',
    permissions JSONB DEFAULT '[]', created_by TEXT, created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE TABLE IF NOT EXISTS audit_log (
    id TEXT PRIMARY KEY, user_id TEXT, action TEXT NOT NULL, module TEXT,
    entity_id TEXT, detail TEXT, ip TEXT,
    tenant_id TEXT REFERENCES tenants(id), created_at TIMESTAMPTZ DEFAULT NOW()
);

/* ── Performance Indexes ── */
CREATE INDEX IF NOT EXISTS idx_users_tenant        ON users(tenant_id);
CREATE INDEX IF NOT EXISTS idx_users_email         ON users(email);
CREATE INDEX IF NOT EXISTS idx_companies_tenant    ON companies(tenant_id);
CREATE INDEX IF NOT EXISTS idx_directors_company   ON directors(company_id);
CREATE INDEX IF NOT EXISTS idx_directors_tenant    ON directors(tenant_id);
CREATE INDEX IF NOT EXISTS idx_tasks_leader        ON tasks(task_leader);
CREATE INDEX IF NOT EXISTS idx_tasks_manager       ON tasks(task_manager);
CREATE INDEX IF NOT EXISTS idx_tasks_assignee      ON tasks(assigned_to);
CREATE INDEX IF NOT EXISTS idx_tasks_status        ON tasks(status);
CREATE INDEX IF NOT EXISTS idx_tasks_tenant        ON tasks(tenant_id);
CREATE INDEX IF NOT EXISTS idx_alerts_status       ON alerts(status);
CREATE INDEX IF NOT EXISTS idx_alerts_company      ON alerts(company_id);
CREATE INDEX IF NOT EXISTS idx_alerts_tenant       ON alerts(tenant_id);
CREATE INDEX IF NOT EXISTS idx_meetings_company    ON meetings(company_id);
CREATE INDEX IF NOT EXISTS idx_meetings_date       ON meetings(meeting_date);
CREATE INDEX IF NOT EXISTS idx_audit_log_tenant    ON audit_log(tenant_id);
CREATE INDEX IF NOT EXISTS idx_audit_log_user      ON audit_log(user_id);
CREATE INDEX IF NOT EXISTS idx_dsc_tenant          ON dsc_records(tenant_id);
CREATE INDEX IF NOT EXISTS idx_dsc_valid_to        ON dsc_records(valid_to);

/* ── Composite indexes for common query patterns ── */
CREATE INDEX IF NOT EXISTS idx_tasks_tenant_status   ON tasks(tenant_id, status);
CREATE INDEX IF NOT EXISTS idx_tasks_assignee_status ON tasks(assigned_to, status);
CREATE INDEX IF NOT EXISTS idx_tasks_due_active      ON tasks(due_date) WHERE status NOT IN ('completed','cancelled');
CREATE INDEX IF NOT EXISTS idx_alerts_tenant_active  ON alerts(tenant_id, status);
CREATE INDEX IF NOT EXISTS idx_alerts_company_status ON alerts(company_id, status);
CREATE INDEX IF NOT EXISTS idx_meetings_tenant_date  ON meetings(tenant_id, meeting_date);
CREATE INDEX IF NOT EXISTS idx_directors_active      ON directors(company_id, is_active);
CREATE INDEX IF NOT EXISTS idx_esop_company          ON esop_grants(company_id, status);

/* ── New tables for improvements ── */

/* Custom alert rules */
CREATE TABLE IF NOT EXISTS custom_alert_rules (
    id TEXT PRIMARY KEY,
    tenant_id TEXT REFERENCES tenants(id),
    company_id TEXT REFERENCES companies(id) ON DELETE CASCADE,
    rule_name TEXT NOT NULL,
    entity_type TEXT NOT NULL,          /* director/auditor/dsc/meeting/filing/custom */
    condition_field TEXT NOT NULL,       /* e.g. valid_to, due_date, next_kyc_date */
    condition_days INTEGER DEFAULT 30,   /* alert X days before */
    severity TEXT DEFAULT 'medium',
    is_active INTEGER DEFAULT 1,
    created_by TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_custom_rules_tenant ON custom_alert_rules(tenant_id, is_active);

/* Company group / subsidiary mapping */
ALTER TABLE companies ADD COLUMN IF NOT EXISTS parent_company_id TEXT REFERENCES companies(id) ON DELETE SET NULL;
ALTER TABLE companies ADD COLUMN IF NOT EXISTS group_role TEXT DEFAULT 'standalone'; /* holding/subsidiary/associate/standalone */
CREATE INDEX IF NOT EXISTS idx_companies_parent ON companies(parent_company_id);

/* Director photo and signature */
ALTER TABLE directors ADD COLUMN IF NOT EXISTS photo_url TEXT;
ALTER TABLE directors ADD COLUMN IF NOT EXISTS signature_url TEXT;

/* ESOP vesting schedule */
CREATE TABLE IF NOT EXISTS esop_vesting_schedule (
    id TEXT PRIMARY KEY,
    esop_grant_id TEXT NOT NULL REFERENCES esop_grants(id) ON DELETE CASCADE,
    vesting_date DATE NOT NULL,
    options_vesting INTEGER DEFAULT 0,
    options_vested INTEGER DEFAULT 0,     /* actually exercised/confirmed */
    cliff INTEGER DEFAULT 0,              /* 1 = cliff vesting event */
    notes TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_esop_vesting_grant ON esop_vesting_schedule(esop_grant_id, vesting_date);

/* Compliance calendar entries */
CREATE TABLE IF NOT EXISTS compliance_calendar (
    id TEXT PRIMARY KEY,
    tenant_id TEXT REFERENCES tenants(id),
    company_id TEXT REFERENCES companies(id) ON DELETE CASCADE,
    form_code TEXT NOT NULL,             /* MGT-7, AOC-4, ADT-1 … */
    form_name TEXT NOT NULL,
    due_date DATE NOT NULL,
    financial_year TEXT,                 /* e.g. 2025-26 */
    status TEXT DEFAULT 'pending',       /* pending/filed/overdue/na */
    filed_date DATE,
    srn TEXT,
    notes TEXT,
    alert_generated INTEGER DEFAULT 0,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_cal_company_fy   ON compliance_calendar(company_id, financial_year);
CREATE INDEX IF NOT EXISTS idx_cal_tenant_due   ON compliance_calendar(tenant_id, due_date);
CREATE INDEX IF NOT EXISTS idx_cal_status        ON compliance_calendar(status, due_date);
"""

# ── SQLite schema (dev) ───────────────────────────────────────────────────────
SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS tenants (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, slug TEXT UNIQUE NOT NULL,
    email TEXT NOT NULL, phone TEXT DEFAULT '', address TEXT DEFAULT '',
    plan_id TEXT, status TEXT DEFAULT 'pending', approved_by TEXT,
    approved_at TEXT, trial_ends TEXT, billing_cycle TEXT DEFAULT 'monthly',
    max_users INTEGER DEFAULT 5, max_companies INTEGER DEFAULT 5,
    notes TEXT DEFAULT '', created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS plans (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, description TEXT DEFAULT '',
    price_monthly REAL DEFAULT 0, price_annual REAL DEFAULT 0,
    max_users INTEGER DEFAULT 5, max_companies INTEGER DEFAULT 5,
    max_documents INTEGER DEFAULT 100, features TEXT DEFAULT '[]',
    is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS users (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, email TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'staff',
    is_active INTEGER DEFAULT 1, is_platform_admin INTEGER DEFAULT 0,
    tenant_id TEXT, created_at TEXT DEFAULT (datetime('now')), last_login TEXT
);
CREATE TABLE IF NOT EXISTS companies (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, cin TEXT UNIQUE,
    incorporation_date TEXT, registered_office TEXT, pan TEXT, tan TEXT,
    email TEXT, phone TEXT, authorized_capital REAL DEFAULT 0,
    paid_up_capital REAL DEFAULT 0, business_activity TEXT,
    company_type TEXT DEFAULT 'Private Limited', roc TEXT, category TEXT,
    sub_category TEXT, status TEXT DEFAULT 'active', letterhead_logo TEXT,
    letterhead_address TEXT, letterhead_footer TEXT, gstin TEXT, website TEXT,
    tenant_id TEXT, created_by TEXT,
    created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS directors (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, name TEXT NOT NULL,
    din TEXT, pan TEXT, aadhaar TEXT, email TEXT, mobile TEXT, address TEXT,
    designation TEXT DEFAULT 'Director', date_of_appointment TEXT,
    date_of_cessation TEXT, mca_user_id TEXT, mca_password TEXT, mca_notes TEXT,
    is_active INTEGER DEFAULT 1, tenant_id TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS director_kyc (
    id TEXT PRIMARY KEY,
    director_id TEXT NOT NULL REFERENCES directors(id) ON DELETE CASCADE,
    company_id TEXT REFERENCES companies(id) ON DELETE CASCADE,
    financial_year TEXT,
    kyc_date TEXT,
    last_kyc_date TEXT,
    next_due_date TEXT,
    kyc_status TEXT DEFAULT 'pending',
    mobile TEXT,
    email TEXT,
    address TEXT,
    notes TEXT,
    tenant_id TEXT,
    updated_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS auditors (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, name TEXT NOT NULL,
    firm_name TEXT, membership_no TEXT, frn TEXT, pan TEXT, address TEXT,
    email TEXT, phone TEXT, appointment_date TEXT,
    nature_of_appointment TEXT DEFAULT 'Regular Auditor',
    appointment_type TEXT DEFAULT 'AGM Appointment',
    start_date TEXT, end_date TEXT, srn_adt1 TEXT, adt1_file TEXT,
    is_active INTEGER DEFAULT 1, notes TEXT, tenant_id TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS shareholders (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, name TEXT NOT NULL,
    folio_no TEXT, pan TEXT, email TEXT, mobile TEXT, address TEXT,
    share_class TEXT DEFAULT 'Equity', shares_held INTEGER DEFAULT 0,
    face_value REAL DEFAULT 10, date_of_entry TEXT,
    mca_user_id TEXT, mca_password TEXT, is_active INTEGER DEFAULT 1,
    tenant_id TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS dsc_records (
    id TEXT PRIMARY KEY, company_id TEXT, director_id TEXT,
    holder_name TEXT NOT NULL, holder_type TEXT DEFAULT 'Director',
    dsc_class TEXT DEFAULT 'Class 3', issued_by TEXT,
    valid_from TEXT, valid_to TEXT, token_type TEXT,
    custody_status TEXT DEFAULT 'With Client', custody_date TEXT, custody_notes TEXT,
    alert_30day_sent INTEGER DEFAULT 0, is_active INTEGER DEFAULT 1, notes TEXT,
    tenant_id TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS dsc_custody_log (
    id TEXT PRIMARY KEY, dsc_id TEXT NOT NULL, action TEXT NOT NULL,
    action_date TEXT, from_party TEXT, to_party TEXT, notes TEXT, recorded_by TEXT,
    tenant_id TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS meetings (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, meeting_type TEXT NOT NULL,
    meeting_no TEXT, meeting_date TEXT NOT NULL, meeting_time TEXT, venue TEXT,
    agenda TEXT, notes TEXT, minutes_drafted INTEGER DEFAULT 0,
    status TEXT DEFAULT 'scheduled', tenant_id TEXT, created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS tasks (
    id TEXT PRIMARY KEY, company_id TEXT, title TEXT NOT NULL, description TEXT,
    assigned_to TEXT, due_date TEXT, priority TEXT DEFAULT 'medium',
    status TEXT DEFAULT 'pending', module TEXT, entity_id TEXT, created_by TEXT,
    completed_at TEXT, task_leader TEXT, task_manager TEXT,
    team_members TEXT DEFAULT '[]', note TEXT DEFAULT '',
    billable INTEGER DEFAULT 0, estimated_hrs REAL DEFAULT 0, actual_hrs REAL DEFAULT 0,
    tenant_id TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS alerts (
    id TEXT PRIMARY KEY, company_id TEXT, entity_type TEXT NOT NULL,
    entity_id TEXT NOT NULL, alert_type TEXT NOT NULL,
    title TEXT NOT NULL, message TEXT, due_date TEXT,
    severity TEXT DEFAULT 'medium', status TEXT DEFAULT 'active',
    tenant_id TEXT, created_at TEXT DEFAULT (datetime('now')), resolved_at TEXT
);
CREATE TABLE IF NOT EXISTS charges (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, charge_id TEXT, srn TEXT,
    date_of_creation TEXT, amount REAL DEFAULT 0, charge_holder TEXT,
    assets_charged TEXT, charge_type TEXT DEFAULT 'Hypothecation',
    date_of_modification TEXT, date_of_satisfaction TEXT, status TEXT DEFAULT 'Open',
    remarks TEXT, tenant_id TEXT, created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS investments (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, investee_name TEXT NOT NULL,
    investee_type TEXT DEFAULT 'Company', investment_type TEXT DEFAULT 'Equity Shares',
    amount REAL DEFAULT 0, date_of_investment TEXT, board_resolution_date TEXT,
    srn_mgb4 TEXT, purpose TEXT, remarks TEXT, tenant_id TEXT, created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS loans_guarantees (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, party_name TEXT NOT NULL,
    party_type TEXT DEFAULT 'Company', transaction_type TEXT DEFAULT 'Loan',
    amount REAL DEFAULT 0, date_of_transaction TEXT, rate_of_interest REAL DEFAULT 0,
    repayment_date TEXT, security TEXT, board_resolution_date TEXT,
    outstanding_amount REAL DEFAULT 0, status TEXT DEFAULT 'Active', remarks TEXT,
    tenant_id TEXT, created_by TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS related_party_transactions (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, party_name TEXT NOT NULL,
    relationship TEXT, nature_of_transaction TEXT, amount REAL DEFAULT 0,
    date_of_transaction TEXT, date_of_board_approval TEXT,
    date_of_shareholders_approval TEXT, terms TEXT, justification TEXT, remarks TEXT,
    tenant_id TEXT, created_by TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS director_interests (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, director_id TEXT,
    director_name TEXT NOT NULL, din TEXT, entity_name TEXT NOT NULL,
    entity_type TEXT DEFAULT 'Company', nature_of_interest TEXT,
    date_of_disclosure TEXT, date_of_board_resolution TEXT, remarks TEXT,
    tenant_id TEXT, created_by TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS esop_grants (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, employee_name TEXT NOT NULL,
    designation TEXT, employee_id TEXT, grant_date TEXT,
    options_granted INTEGER DEFAULT 0, exercise_price REAL DEFAULT 0,
    vesting_date TEXT, vesting_period TEXT, options_exercised INTEGER DEFAULT 0,
    options_lapsed INTEGER DEFAULT 0, status TEXT DEFAULT 'Active', remarks TEXT,
    tenant_id TEXT, created_by TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS statutory_registers (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, register_type TEXT NOT NULL,
    register_name TEXT NOT NULL, section_ref TEXT, entry_date TEXT, folio_ref TEXT,
    entry_data TEXT, remarks TEXT, recorded_by TEXT, tenant_id TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS document_templates (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, category TEXT NOT NULL, description TEXT,
    template_body TEXT NOT NULL, placeholders TEXT DEFAULT '[]',
    is_system INTEGER DEFAULT 0, is_active INTEGER DEFAULT 1,
    tenant_id TEXT, created_by TEXT,
    created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS documents (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL, template_id TEXT,
    doc_type TEXT NOT NULL, doc_name TEXT NOT NULL, content TEXT, file_path TEXT,
    module TEXT, entity_id TEXT, date_from TEXT DEFAULT '', date_to TEXT DEFAULT '',
    tenant_id TEXT, generated_by TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS custom_placeholders (
    id TEXT PRIMARY KEY, name TEXT UNIQUE NOT NULL, label TEXT NOT NULL,
    description TEXT DEFAULT '', ph_type TEXT DEFAULT 'text',
    tenant_id TEXT, created_by TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS user_permissions (
    id TEXT PRIMARY KEY, user_id TEXT NOT NULL, module TEXT NOT NULL,
    action TEXT NOT NULL, granted INTEGER NOT NULL DEFAULT 1,
    granted_by TEXT, granted_at TEXT DEFAULT (datetime('now')), note TEXT DEFAULT '',
    UNIQUE(user_id, module, action)
);
CREATE TABLE IF NOT EXISTS permission_presets (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, description TEXT DEFAULT '',
    permissions TEXT DEFAULT '[]', created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS audit_log (
    id TEXT PRIMARY KEY, user_id TEXT, action TEXT NOT NULL, module TEXT,
    entity_id TEXT, detail TEXT, ip TEXT, tenant_id TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

-- ── P2 Feature tables ──────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS compliance_calendar (
    id TEXT PRIMARY KEY,
    tenant_id TEXT,
    company_id TEXT REFERENCES companies(id) ON DELETE CASCADE,
    form_code TEXT NOT NULL,
    form_name TEXT NOT NULL,
    due_date TEXT NOT NULL,
    financial_year TEXT,
    status TEXT DEFAULT 'pending',
    filed_date TEXT,
    srn TEXT,
    notes TEXT,
    alert_generated INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS custom_alert_rules (
    id TEXT PRIMARY KEY,
    tenant_id TEXT,
    company_id TEXT,
    rule_name TEXT NOT NULL,
    entity_type TEXT NOT NULL,
    condition_field TEXT NOT NULL,
    condition_days INTEGER DEFAULT 30,
    severity TEXT DEFAULT 'medium',
    is_active INTEGER DEFAULT 1,
    created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS esop_vesting_schedule (
    id TEXT PRIMARY KEY,
    esop_grant_id TEXT NOT NULL REFERENCES esop_grants(id) ON DELETE CASCADE,
    vesting_date TEXT NOT NULL,
    options_vesting INTEGER DEFAULT 0,
    options_vested INTEGER DEFAULT 0,
    cliff INTEGER DEFAULT 0,
    notes TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

-- ── ALTER TABLE equivalents for SQLite (add columns if missing) ─────────────
-- SQLite doesn't support ALTER TABLE ADD COLUMN IF NOT EXISTS
-- These are handled by ensure_columns() in app.py

"""
def init_db():
    """Create all tables and seed initial data."""
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    (BASE_DIR / "data" / "pdfs").mkdir(parents=True, exist_ok=True)
    (BASE_DIR / "data" / "legal_docs").mkdir(parents=True, exist_ok=True)

    conn = get_db(); c = conn.cursor()

    if USE_POSTGRES:
        for stmt in [s.strip() for s in SCHEMA_PG.split(';') if s.strip()]:
            try: c.execute(stmt)
            except Exception as e: print(f"  Schema: {e}")
    else:
        import sqlite3 as _sq3
        raw = _sq3.connect(str(DB_PATH))
        raw.executescript(SCHEMA_SQLITE)
        raw.commit(); raw.close()
        # Run column migrations for any new columns not in original schema
        ensure_columns()
        conn.close()
        conn = get_db(); c = conn.cursor()

    conn.commit()

    c.execute("SELECT COUNT(*) FROM users")
    r = c.fetchone()
    cnt = r['count'] if isinstance(r, dict) and 'count' in r else (list(r.values())[0] if isinstance(r, dict) else r[0])
    if int(cnt or 0) == 0:
        _seed(c, conn)

    # Only seed templates if none exist yet
    c.execute("SELECT COUNT(*) FROM document_templates")
    _tr = c.fetchone()
    _tcnt = _tr.get('count', list(_tr.values())[0]) if isinstance(_tr, dict) else _tr[0]
    if int(_tcnt or 0) == 0:
        _seed_templates(c, conn)

    # ── Advanced feature tables ──────────────────────────────
    _adv = [
        """CREATE TABLE IF NOT EXISTS esign_requests (id TEXT PRIMARY KEY, tenant_id TEXT, doc_title TEXT, doc_id TEXT, provider TEXT DEFAULT 'leegality', signatories TEXT, status TEXT DEFAULT 'pending', deadline TEXT, created_by TEXT, created_at TEXT)""",
        """CREATE TABLE IF NOT EXISTS portal_links (id TEXT PRIMARY KEY, tenant_id TEXT, company_id TEXT, token TEXT UNIQUE, access_level TEXT DEFAULT 'full_readonly', client_name TEXT, expires_at TEXT, active INTEGER DEFAULT 1, view_count INTEGER DEFAULT 0, created_by TEXT, created_at TEXT)""",
        """CREATE TABLE IF NOT EXISTS wa_templates (id TEXT PRIMARY KEY, tenant_id TEXT, name TEXT, body TEXT, created_at TEXT)""",
        """CREATE TABLE IF NOT EXISTS wa_message_log (id TEXT PRIMARY KEY, tenant_id TEXT, phone TEXT, message TEXT, type TEXT, status TEXT DEFAULT 'sent', sent_at TEXT)""",
        """CREATE TABLE IF NOT EXISTS wa_schedules (id TEXT PRIMARY KEY, tenant_id TEXT, alert_type TEXT, send_time TEXT DEFAULT '09:00', recipients TEXT DEFAULT 'directors', active INTEGER DEFAULT 1, created_at TEXT)""",
    ]
    for _s in _adv:
        try: c.execute(_s)
        except Exception: pass
    conn.commit()
    conn
    print(f"DB initialised ({'PostgreSQL' if USE_POSTGRES else 'SQLite dev mode'})")


def ensure_columns():
    """
    Add any new columns that may not exist in older SQLite databases.
    Safe to run on every startup — skips columns that already exist.
    PostgreSQL uses ALTER TABLE ADD COLUMN IF NOT EXISTS in the schema string.
    SQLite does not support IF NOT EXISTS on ALTER TABLE, so we do it here.
    """
    if USE_POSTGRES:
        return  # PG schema handles this with IF NOT EXISTS
    try:
        import sqlite3 as _sq3
        raw = _sq3.connect(str(DB_PATH))
        cur = raw.cursor()
        # companies → parent_company_id, group_role
        existing = {row[1] for row in cur.execute("PRAGMA table_info(companies)")}
        # shareholders distinctive_no columns (migration)
        sh_existing = {r[1] for r in cur.execute("PRAGMA table_info(shareholders)")}
        if "distinctive_no_from" not in sh_existing:
            cur.execute("ALTER TABLE shareholders ADD COLUMN distinctive_no_from INTEGER")
        if "distinctive_no_to" not in sh_existing:
            cur.execute("ALTER TABLE shareholders ADD COLUMN distinctive_no_to INTEGER")
        # share_transfers table (migration)
        try:
            cur.execute("CREATE TABLE IF NOT EXISTS share_transfers ("
                        "id TEXT PRIMARY KEY, company_id TEXT, transferor_id TEXT, transferee_id TEXT,"
                        "shares_transferred INTEGER, transfer_date TEXT, distinctive_no_from INTEGER,"
                        "distinctive_no_to INTEGER, remarks TEXT, tenant_id TEXT, created_at TEXT)")
        except Exception: pass

        if "parent_company_id" not in existing:
            cur.execute("ALTER TABLE companies ADD COLUMN parent_company_id TEXT")
        if "group_role" not in existing:
            cur.execute("ALTER TABLE companies ADD COLUMN group_role TEXT DEFAULT 'standalone'")
        # directors → photo_url, signature_url
        existing = {row[1] for row in cur.execute("PRAGMA table_info(directors)")}
        if "photo_url" not in existing:
            cur.execute("ALTER TABLE directors ADD COLUMN photo_url TEXT")
        if "signature_url" not in existing:
            cur.execute("ALTER TABLE directors ADD COLUMN signature_url TEXT")
        # director_kyc → extra columns
        existing = {row[1] for row in cur.execute("PRAGMA table_info(director_kyc)")}
        for col_name, col_type in [("company_id","TEXT"),("financial_year","TEXT"),
                                    ("kyc_date","TEXT"),("mobile","TEXT"),
                                    ("email","TEXT"),("address","TEXT")]:
            if col_name not in existing:
                try: cur.execute(f"ALTER TABLE director_kyc ADD COLUMN {col_name} {col_type}")
                except Exception: pass
        raw.commit(); raw.close()
    except Exception as _e:
        import logging; logging.getLogger(__name__).warning(f"ensure_columns: {_e}")


def ensure_custom_placeholders_table():
    """Create the custom_placeholders table if it doesn't exist."""
    try:
        conn = get_db(); c = conn.cursor()
        if USE_POSTGRES:
            c.execute("""CREATE TABLE IF NOT EXISTS custom_placeholders (
                id TEXT PRIMARY KEY, tenant_id TEXT, company_id TEXT,
                placeholder TEXT NOT NULL, value TEXT,
                created_at TIMESTAMPTZ DEFAULT NOW())""")
        else:
            c.execute("""CREATE TABLE IF NOT EXISTS custom_placeholders (
                id TEXT PRIMARY KEY, tenant_id TEXT, company_id TEXT,
                placeholder TEXT NOT NULL, value TEXT,
                created_at TEXT DEFAULT (datetime('now')))""")
        conn.commit(); conn.close()
    except Exception as _e:
        import logging; logging.getLogger(__name__).warning(f"ensure_custom_placeholders_table: {_e}")


def ensure_permission_tables():
    """Create the permissions table if it doesn't exist."""
    try:
        conn = get_db(); c = conn.cursor()
        if USE_POSTGRES:
            c.execute("""CREATE TABLE IF NOT EXISTS user_permissions (
                id TEXT PRIMARY KEY, user_id TEXT NOT NULL, module TEXT NOT NULL,
                can_view INTEGER DEFAULT 1, can_create INTEGER DEFAULT 0,
                can_update INTEGER DEFAULT 0, can_delete INTEGER DEFAULT 0,
                tenant_id TEXT, UNIQUE(user_id, module))""")
        else:
            c.execute("""CREATE TABLE IF NOT EXISTS user_permissions (
                id TEXT PRIMARY KEY, user_id TEXT NOT NULL, module TEXT NOT NULL,
                can_view INTEGER DEFAULT 1, can_create INTEGER DEFAULT 0,
                can_update INTEGER DEFAULT 0, can_delete INTEGER DEFAULT 0,
                tenant_id TEXT, UNIQUE(user_id, module))""")
        conn.commit(); conn.close()
    except Exception as _e:
        import logging; logging.getLogger(__name__).warning(f"ensure_permission_tables: {_e}")


def _seed(c, conn):
    today = date.today()
    tid   = "default-tenant-001"

    _ex(c, """INSERT INTO tenants (id,name,slug,email,phone,address,status,max_users,max_companies)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
        (tid,"Taxly India Private Limited","taxlyindia","info@taxlyindia.com",
         "+91 88829 35471","L-30B, LGF, Malviya Nagar, New Delhi-110017","active",10,10))

    _ex(c, """INSERT INTO plans (id,name,description,price_monthly,price_annual,max_users,max_companies,max_documents,features,is_active)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,1)""",
        (str(uuid.uuid4()),"Starter","For small firms",2999,29990,3,3,50,'["companies","tasks"]'))
    _ex(c, """INSERT INTO plans (id,name,description,price_monthly,price_annual,max_users,max_companies,max_documents,features,is_active)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,1)""",
        (str(uuid.uuid4()),"Professional","Most popular",6999,69990,10,10,500,'["all"]'))
    _ex(c, """INSERT INTO plans (id,name,description,price_monthly,price_annual,max_users,max_companies,max_documents,features,is_active)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,1)""",
        (str(uuid.uuid4()),"Enterprise","Large firms",0,0,50,50,999,'["all"]'))

    _ex(c, """INSERT INTO users (id,name,email,password,role,is_active,is_platform_admin)
               VALUES (%s,%s,%s,%s,%s,1,1)""",
        (str(uuid.uuid4()),"Platform Admin","platform@taxlycms.in",hash_pw("platform@2025"),"superadmin"))

    for nm, em, pw, rl in [
        ("Super Admin","admin@compli.in","admin123","superadmin"),
        ("Priya Sharma","manager@compli.in","manager123","manager"),
        ("Rahul Verma","staff@compli.in","staff123","staff"),
    ]:
        _ex(c, "INSERT INTO users (id,name,email,password,role,is_active,tenant_id) VALUES (%s,%s,%s,%s,%s,1,%s)",
            (str(uuid.uuid4()),nm,em,hash_pw(pw),rl,tid))

    c.execute("SELECT id FROM users WHERE email=%s", ("admin@compli.in",))
    r = c.fetchone(); uid = (r['id'] if isinstance(r,dict) else r[0]) if r else str(uuid.uuid4())

    co1 = str(uuid.uuid4())
    _ex(c, """INSERT INTO companies (id,name,cin,incorporation_date,registered_office,pan,tan,
               email,phone,authorized_capital,paid_up_capital,business_activity,
               company_type,roc,tenant_id,created_by) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
        (co1,"Innovate Solutions Pvt Ltd","U72200MH2018PTC309876","2018-03-15",
         "12 Nariman Point, Mumbai 400021","AABCI1234P","MUMA12345A","cs@innovate.in",
         "022-22001234",5000000,1000000,"Software Development","Private Limited","RoC-Mumbai",tid,uid))

    co2 = str(uuid.uuid4())
    _ex(c, """INSERT INTO companies (id,name,cin,incorporation_date,registered_office,pan,tan,
               email,phone,authorized_capital,paid_up_capital,business_activity,
               company_type,roc,tenant_id,created_by) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
        (co2,"GreenTech Manufacturing Ltd","L24100DL2015PLC285432","2015-07-22",
         "45 Connaught Place, New Delhi 110001","AABCG5678Q","DELA98765B","legal@greentech.in",
         "011-43210000",50000000,20000000,"Manufacturing","Public Limited","RoC-Delhi",tid,uid))

    d1=str(uuid.uuid4()); d2=str(uuid.uuid4())
    _ex(c, "INSERT INTO directors (id,company_id,name,din,pan,email,mobile,designation,date_of_appointment,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        (d1,co1,"Arjun Mehta","00123456","ABCPM1234D","arjun@innovate.in","9876543210","Managing Director","2018-03-15",tid))
    _ex(c, "INSERT INTO directors (id,company_id,name,din,pan,email,mobile,designation,date_of_appointment,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        (d2,co1,"Sneha Patel","00234567","ABCSP5678E","sneha@innovate.in","9123456789","Director","2019-06-01",tid))

    for did, last_kyc, st in [(d1,"2024-08-15","filed"),(d2,"2023-09-10","overdue")]:
        _ex(c, "INSERT INTO director_kyc (id,director_id,last_kyc_date,next_due_date,kyc_status,tenant_id) VALUES (%s,%s,%s,%s,%s,%s)",
            (str(uuid.uuid4()),did,last_kyc,"2025-09-30",st,tid))

    aud1=str(uuid.uuid4()); near=(today+timedelta(days=25)).isoformat()
    _ex(c, "INSERT INTO auditors (id,company_id,name,firm_name,membership_no,frn,pan,email,appointment_date,start_date,end_date,srn_adt1,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        (aud1,co1,"CA Ramesh Gupta","Gupta & Associates","123456","012345W","ABCPG3456G","ca@gupta.in","2023-09-30","2023-09-30",near,"ADT1202312345",tid))

    for nm,folio,shares in [("Arjun Mehta","F0001",50000),("Sneha Patel","F0002",30000)]:
        _ex(c,"INSERT INTO shareholders (id,company_id,name,folio_no,shares_held,date_of_entry,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s)",
            (str(uuid.uuid4()),co1,nm,folio,shares,"2018-03-15",tid))

    dsc1=str(uuid.uuid4()); dsc_exp=(today+timedelta(days=22)).isoformat()
    _ex(c,"INSERT INTO dsc_records (id,company_id,director_id,holder_name,holder_type,dsc_class,issued_by,valid_from,valid_to,token_type,custody_status,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        (dsc1,co1,d1,"Arjun Mehta","Director","Class 3","eMudhra","2022-06-01",dsc_exp,"ePass 2003","With Us",tid))

    _ex(c,"INSERT INTO meetings (id,company_id,meeting_type,meeting_no,meeting_date,meeting_time,venue,status,tenant_id,created_by) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        (str(uuid.uuid4()),co1,"Board","BM-2025-01",(today+timedelta(days=7)).isoformat(),"11:00","Registered Office","scheduled",tid,uid))

    c.execute("SELECT id FROM users WHERE email=%s",("admin@compli.in",)); r=c.fetchone(); sa=r['id'] if isinstance(r,dict) else r[0]
    c.execute("SELECT id FROM users WHERE email=%s",("manager@compli.in",)); r=c.fetchone(); mg=r['id'] if r else None
    c.execute("SELECT id FROM users WHERE email=%s",("staff@compli.in",)); r=c.fetchone(); st=r['id'] if r else None
    if isinstance(mg,dict): mg=mg['id']
    if isinstance(st,dict): st=st['id']

    for title,pri,status,dd,mod in [
        ("File ADT-1 for Innovate Solutions","high","pending",20,"auditor"),
        ("Director KYC for Sneha Patel","critical","pending",5,"director"),
        ("Renew DSC — Arjun Mehta","critical","pending",22,"dsc"),
        ("Prepare AGM Notice","medium","in_progress",30,"meeting"),
        ("ROC Annual Return Filing","high","pending",60,"filing"),
        ("Test Task with Roles","high","pending",42,"general"),
    ]:
        _ex(c,"INSERT INTO tasks (id,company_id,title,priority,status,due_date,module,task_leader,task_manager,assigned_to,tenant_id,created_by) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            (str(uuid.uuid4()),co1,title,pri,status,(today+timedelta(days=dd)).isoformat(),mod,sa,mg,st,tid,uid))

    for etype,eid,atype,ttl,msg,dd_val,sev in [
        ("auditor",aud1,"expiry","Auditor Appointment Expiring Soon",f"CA Ramesh Gupta expires {near}",near,"critical"),
        ("dsc",dsc1,"dsc_expiry",f"DSC Expiring — Arjun Mehta",f"DSC expires {dsc_exp}",dsc_exp,"critical"),
        ("director",d2,"kyc_due","Director KYC Overdue — Sneha Patel","DIR-3 KYC overdue","2025-09-30","critical"),
    ]:
        _ex(c,"INSERT INTO alerts (id,company_id,entity_type,entity_id,alert_type,title,message,due_date,severity,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            (str(uuid.uuid4()),co1,etype,eid,atype,ttl,msg,dd_val,sev,tid))

    conn.commit()


def _ex(c, sql, params):
    """Execute with IGNORE on conflict for SQLite compatibility."""
    try: c.execute(sql, params)
    except Exception: pass


def _seed_templates(c, conn):
    c.execute("SELECT COUNT(*) FROM document_templates WHERE is_system=1")
    r = c.fetchone()
    cnt = r.get('count', list(r.values())[0]) if isinstance(r, dict) else r[0]
    if int(cnt or 0) > 0: return
    templates = [
        ("Board Resolution — General","resolution","General-purpose board resolution.",
         "BOARD RESOLUTION\n\n{{company_name}}\n\nRESOLVED THAT {{resolution_text}}\n\n{{director_name}} | {{designation}} | DIN: {{din}}\nDate: {{resolution_date}}",
         ["company_name","resolution_text","director_name","designation","din","resolution_date"]),
        ("AGM Notice","notice","Notice of Annual General Meeting.",
         "NOTICE OF AGM\n\n{{company_name}} | CIN: {{cin}}\n\nNOTICE: The {{agm_number}} AGM will be held on {{meeting_date}} at {{venue}}.\n\n{{director_name}}",
         ["company_name","cin","agm_number","meeting_date","venue","director_name"]),
        ("Board Meeting Minutes","minutes","Minutes of Board meeting.",
         "MINUTES OF BOARD MEETING\n\n{{company_name}} | Date: {{meeting_date}}\n\nDIRECTORS PRESENT:\n{{directors_present}}\n\nRESOLUTIONS:\n{{resolutions_passed}}",
         ["company_name","meeting_date","directors_present","resolutions_passed"]),
        ("Authorisation Letter","letter","General authorisation letter.",
         "AUTHORISATION LETTER\n\n{{company_name}}\nDate: {{letter_date}}\n\n{{authorised_person}} is authorised to {{purpose}}.\n\nFor {{company_name}}\n{{director_name}}",
         ["company_name","letter_date","authorised_person","purpose","director_name"]),
        ("Director's Report","report","Annual Director's Report under Section 134.",
         "DIRECTORS' REPORT\n\nTo, The Members, {{company_name}}\n\nYour Directors present the {{report_year}} Annual Report.\n\n{{director_name}} | DIN: {{din}}",
         ["company_name","report_year","director_name","din"]),
    ]
    for name, cat, desc, body, ph in templates:
        _ex(c, """INSERT INTO document_templates
            (id,name,category,description,template_body,placeholders,is_system,tenant_id)
            VALUES (%s,%s,%s,%s,%s,%s,1,'default-tenant-001')""",
            (str(uuid.uuid4()),name,cat,desc,body,json.dumps(ph)))
    conn.commit()


def write_audit_log(user_id: str, action: str, module: str = None,
                    entity_id: str = None, detail: str = None,
                    ip: str = None, tenant_id: str = None):
    """Write an audit log entry. Call from route handlers for sensitive operations."""
    try:
        conn = get_db(); c = conn.cursor()
        c.execute("""INSERT INTO audit_log
                     (id, user_id, action, module, entity_id, detail, ip, tenant_id)
                     VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""",
                  (str(uuid.uuid4()), user_id, action, module,
                   entity_id, detail, ip, tenant_id))
        conn.commit(); conn.close()
    except Exception as e:
        import logging
        logging.getLogger(__name__).error(f"audit_log write failed: {e}")


