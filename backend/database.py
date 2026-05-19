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
USE_POSTGRES  = bool(DATABASE_URL)

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
def _pg_conn():
    from urllib.parse import urlparse
    p = urlparse(DATABASE_URL)
    return psycopg2.connect(
        host     = p.hostname,
        port     = p.port or 5432,
        dbname   = p.path.lstrip("/"),
        user     = p.username,
        password = p.password,
        sslmode  = os.environ.get("DB_SSLMODE", "require"),
    )


class PGConnection:
    """Thin wrapper: makes psycopg2 behave like sqlite3 for our row/rows helpers."""
    def __init__(self, conn):
        self._conn = conn
    def cursor(self):
        return self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    def commit(self):   self._conn.commit()
    def rollback(self): self._conn.rollback()
    def close(self):    self._conn.close()


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
    if USE_POSTGRES:
        conn = _pg_conn()
        conn.autocommit = False
        return PGConnection(conn)
    return _sqlite_conn()


def row(r):
    if r is None: return None
    return dict(r) if not isinstance(r, dict) else r

def rows(rs):
    return [dict(r) if not isinstance(r, dict) else r for r in (rs or [])]

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
CREATE TABLE IF NOT EXISTS directors (
    id TEXT PRIMARY KEY, company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    name TEXT NOT NULL, din TEXT, pan TEXT, aadhaar TEXT, email TEXT, mobile TEXT,
    address TEXT, designation TEXT DEFAULT 'Director', date_of_appointment DATE,
    date_of_cessation DATE, mca_user_id TEXT, mca_password TEXT,
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
    date_of_cessation TEXT, mca_user_id TEXT, mca_password TEXT,
    is_active INTEGER DEFAULT 1, tenant_id TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS director_kyc (
    id TEXT PRIMARY KEY, director_id TEXT NOT NULL UNIQUE,
    last_kyc_date TEXT, next_due_date TEXT, kyc_status TEXT DEFAULT 'pending',
    notes TEXT, tenant_id TEXT, updated_at TEXT DEFAULT (datetime('now'))
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
        conn.close()
        conn = get_db(); c = conn.cursor()

    conn.commit()

    c.execute("SELECT COUNT(*) FROM users")
    r = c.fetchone()
    cnt = r['count'] if isinstance(r, dict) and 'count' in r else (list(r.values())[0] if isinstance(r, dict) else r[0])
    if int(cnt or 0) == 0:
        _seed(c, conn)

    _seed_templates(c, conn)
    conn.commit(); conn.close()
    print(f"DB initialised ({'PostgreSQL' if USE_POSTGRES else 'SQLite dev mode'})")


def ensure_columns(): pass
def ensure_custom_placeholders_table(): pass
def ensure_permission_tables(): pass


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
