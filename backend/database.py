"""
database.py — PostgreSQL backend (Hostinger VPS / Cloud DB)
TaxlyCMS — Companies Act 2013 Compliance CRM
"""
import os, hashlib, uuid, json
from datetime import datetime, date, timedelta
from pathlib import Path
from urllib.parse import urlparse

# ── Connection config ─────────────────────────────────────────────────────────
# Set DATABASE_URL in .env or environment:
#   postgresql://user:password@host:5432/dbname
#   OR individual vars: DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD

try:
    import psycopg2
    import psycopg2.extras
    HAS_PG = True
except ImportError:
    HAS_PG = False

BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "data" / "uploads"


def _pg_conn():
    """Return a psycopg2 connection from DATABASE_URL or individual env vars."""
    url = os.environ.get("DATABASE_URL", "")
    if url:
        p = urlparse(url)
        return psycopg2.connect(
            host=p.hostname, port=p.port or 5432,
            dbname=p.path.lstrip("/"), user=p.username, password=p.password,
            sslmode=os.environ.get("DB_SSLMODE", "require"),
        )
    return psycopg2.connect(
        host    = os.environ.get("DB_HOST",     "localhost"),
        port    = int(os.environ.get("DB_PORT", "5432")),
        dbname  = os.environ.get("DB_NAME",     "taxlycms"),
        user    = os.environ.get("DB_USER",     "taxlycms"),
        password= os.environ.get("DB_PASSWORD", ""),
        sslmode = os.environ.get("DB_SSLMODE",  "require"),
    )


class PGConnection:
    """Thin wrapper that makes psycopg2 behave like sqlite3 for our helpers."""
    def __init__(self, conn):
        self._conn = conn
    def cursor(self):
        return self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    def commit(self):  self._conn.commit()
    def rollback(self): self._conn.rollback()
    def close(self):   self._conn.close()
    def execute(self, sql, params=None):
        cur = self.cursor()
        cur.execute(sql, params or ())
        return cur


def get_db():
    if not HAS_PG:
        raise RuntimeError("psycopg2 not installed. Run: pip install psycopg2-binary")
    conn = _pg_conn()
    conn.autocommit = False
    return PGConnection(conn)


def row(r):
    if r is None: return None
    return dict(r)

def rows(rs):
    return [dict(r) for r in (rs or [])]

def hash_pw(pw):
    return hashlib.sha256(pw.encode()).hexdigest()


# ── Schema (PostgreSQL-compatible) ────────────────────────────────────────────
SCHEMA_SQL = """
-- Enable UUID generation
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

CREATE TABLE IF NOT EXISTS tenants (
    id              TEXT PRIMARY KEY,
    name            TEXT NOT NULL,
    slug            TEXT UNIQUE NOT NULL,
    email           TEXT NOT NULL,
    phone           TEXT DEFAULT '',
    address         TEXT DEFAULT '',
    plan_id         TEXT,
    status          TEXT DEFAULT 'pending',
    approved_by     TEXT,
    approved_at     TIMESTAMPTZ,
    trial_ends      DATE,
    billing_cycle   TEXT DEFAULT 'monthly',
    max_users       INTEGER DEFAULT 5,
    max_companies   INTEGER DEFAULT 5,
    notes           TEXT DEFAULT '',
    created_at      TIMESTAMPTZ DEFAULT NOW(),
    updated_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS plans (
    id              TEXT PRIMARY KEY,
    name            TEXT NOT NULL,
    description     TEXT DEFAULT '',
    price_monthly   NUMERIC(12,2) DEFAULT 0,
    price_annual    NUMERIC(12,2) DEFAULT 0,
    max_users       INTEGER DEFAULT 5,
    max_companies   INTEGER DEFAULT 5,
    max_documents   INTEGER DEFAULT 100,
    features        JSONB DEFAULT '[]',
    is_active       INTEGER DEFAULT 1,
    created_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS users (
    id                  TEXT PRIMARY KEY,
    name                TEXT NOT NULL,
    email               TEXT UNIQUE NOT NULL,
    password            TEXT NOT NULL,
    role                TEXT NOT NULL DEFAULT 'staff',
    is_active           INTEGER DEFAULT 1,
    is_platform_admin   INTEGER DEFAULT 0,
    tenant_id           TEXT REFERENCES tenants(id) ON DELETE SET NULL,
    created_at          TIMESTAMPTZ DEFAULT NOW(),
    last_login          TIMESTAMPTZ
);

CREATE TABLE IF NOT EXISTS companies (
    id                  TEXT PRIMARY KEY,
    name                TEXT NOT NULL,
    cin                 TEXT UNIQUE,
    incorporation_date  DATE,
    registered_office   TEXT,
    pan                 TEXT,
    tan                 TEXT,
    email               TEXT,
    phone               TEXT,
    authorized_capital  NUMERIC(20,2) DEFAULT 0,
    paid_up_capital     NUMERIC(20,2) DEFAULT 0,
    business_activity   TEXT,
    company_type        TEXT DEFAULT 'Private Limited',
    roc                 TEXT,
    category            TEXT,
    sub_category        TEXT,
    status              TEXT DEFAULT 'active',
    letterhead_logo     TEXT,
    letterhead_address  TEXT,
    letterhead_footer   TEXT,
    gstin               TEXT,
    website             TEXT,
    tenant_id           TEXT REFERENCES tenants(id),
    created_by          TEXT,
    created_at          TIMESTAMPTZ DEFAULT NOW(),
    updated_at          TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS directors (
    id                  TEXT PRIMARY KEY,
    company_id          TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    name                TEXT NOT NULL,
    din                 TEXT,
    pan                 TEXT,
    aadhaar             TEXT,
    email               TEXT,
    mobile              TEXT,
    address             TEXT,
    designation         TEXT DEFAULT 'Director',
    date_of_appointment DATE,
    date_of_cessation   DATE,
    mca_user_id         TEXT,
    mca_password        TEXT,
    is_active           INTEGER DEFAULT 1,
    tenant_id           TEXT REFERENCES tenants(id),
    created_at          TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS director_kyc (
    id              TEXT PRIMARY KEY,
    director_id     TEXT NOT NULL UNIQUE REFERENCES directors(id) ON DELETE CASCADE,
    last_kyc_date   DATE,
    next_due_date   DATE,
    kyc_status      TEXT DEFAULT 'pending',
    notes           TEXT,
    tenant_id       TEXT REFERENCES tenants(id),
    updated_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS auditors (
    id                      TEXT PRIMARY KEY,
    company_id              TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    name                    TEXT NOT NULL,
    firm_name               TEXT,
    membership_no           TEXT,
    frn                     TEXT,
    pan                     TEXT,
    address                 TEXT,
    email                   TEXT,
    phone                   TEXT,
    appointment_date        DATE,
    nature_of_appointment   TEXT DEFAULT 'Regular Auditor',
    appointment_type        TEXT DEFAULT 'AGM Appointment',
    start_date              DATE,
    end_date                DATE,
    srn_adt1                TEXT,
    adt1_file               TEXT,
    is_active               INTEGER DEFAULT 1,
    notes                   TEXT,
    tenant_id               TEXT REFERENCES tenants(id),
    created_at              TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS shareholders (
    id              TEXT PRIMARY KEY,
    company_id      TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    name            TEXT NOT NULL,
    folio_no        TEXT,
    pan             TEXT,
    email           TEXT,
    mobile          TEXT,
    address         TEXT,
    share_class     TEXT DEFAULT 'Equity',
    shares_held     INTEGER DEFAULT 0,
    face_value      NUMERIC(10,2) DEFAULT 10,
    date_of_entry   DATE,
    mca_user_id     TEXT,
    mca_password    TEXT,
    is_active       INTEGER DEFAULT 1,
    tenant_id       TEXT REFERENCES tenants(id),
    created_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS dsc_records (
    id                  TEXT PRIMARY KEY,
    company_id          TEXT REFERENCES companies(id) ON DELETE SET NULL,
    director_id         TEXT REFERENCES directors(id) ON DELETE SET NULL,
    holder_name         TEXT NOT NULL,
    holder_type         TEXT DEFAULT 'Director',
    dsc_class           TEXT DEFAULT 'Class 3',
    issued_by           TEXT,
    valid_from          DATE,
    valid_to            DATE,
    token_type          TEXT,
    custody_status      TEXT DEFAULT 'With Client',
    custody_date        DATE,
    custody_notes       TEXT,
    alert_30day_sent    INTEGER DEFAULT 0,
    is_active           INTEGER DEFAULT 1,
    notes               TEXT,
    tenant_id           TEXT REFERENCES tenants(id),
    created_at          TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS dsc_custody_log (
    id          TEXT PRIMARY KEY,
    dsc_id      TEXT NOT NULL REFERENCES dsc_records(id) ON DELETE CASCADE,
    action      TEXT NOT NULL,
    action_date DATE,
    from_party  TEXT,
    to_party    TEXT,
    notes       TEXT,
    recorded_by TEXT,
    tenant_id   TEXT REFERENCES tenants(id),
    created_at  TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS meetings (
    id              TEXT PRIMARY KEY,
    company_id      TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    meeting_type    TEXT NOT NULL,
    meeting_no      TEXT,
    meeting_date    DATE NOT NULL,
    meeting_time    TEXT,
    venue           TEXT,
    agenda          TEXT,
    notes           TEXT,
    minutes_drafted INTEGER DEFAULT 0,
    status          TEXT DEFAULT 'scheduled',
    tenant_id       TEXT REFERENCES tenants(id),
    created_by      TEXT,
    created_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS tasks (
    id              TEXT PRIMARY KEY,
    company_id      TEXT REFERENCES companies(id) ON DELETE SET NULL,
    title           TEXT NOT NULL,
    description     TEXT,
    assigned_to     TEXT REFERENCES users(id) ON DELETE SET NULL,
    due_date        DATE,
    priority        TEXT DEFAULT 'medium',
    status          TEXT DEFAULT 'pending',
    module          TEXT,
    entity_id       TEXT,
    created_by      TEXT,
    completed_at    TIMESTAMPTZ,
    task_leader     TEXT REFERENCES users(id) ON DELETE SET NULL,
    task_manager    TEXT REFERENCES users(id) ON DELETE SET NULL,
    team_members    JSONB DEFAULT '[]',
    note            TEXT DEFAULT '',
    billable        INTEGER DEFAULT 0,
    estimated_hrs   NUMERIC(6,2) DEFAULT 0,
    actual_hrs      NUMERIC(6,2) DEFAULT 0,
    tenant_id       TEXT REFERENCES tenants(id),
    created_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS alerts (
    id          TEXT PRIMARY KEY,
    company_id  TEXT REFERENCES companies(id) ON DELETE CASCADE,
    entity_type TEXT NOT NULL,
    entity_id   TEXT NOT NULL,
    alert_type  TEXT NOT NULL,
    title       TEXT NOT NULL,
    message     TEXT,
    due_date    DATE,
    severity    TEXT DEFAULT 'medium',
    status      TEXT DEFAULT 'active',
    tenant_id   TEXT REFERENCES tenants(id),
    created_at  TIMESTAMPTZ DEFAULT NOW(),
    resolved_at TIMESTAMPTZ
);

CREATE TABLE IF NOT EXISTS charges (
    id                      TEXT PRIMARY KEY,
    company_id              TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    charge_id               TEXT,
    srn                     TEXT,
    date_of_creation        DATE,
    amount                  NUMERIC(20,2) DEFAULT 0,
    charge_holder           TEXT,
    assets_charged          TEXT,
    charge_type             TEXT DEFAULT 'Hypothecation',
    date_of_modification    DATE,
    date_of_satisfaction    DATE,
    status                  TEXT DEFAULT 'Open',
    remarks                 TEXT,
    tenant_id               TEXT REFERENCES tenants(id),
    created_by              TEXT,
    created_at              TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS investments (
    id                      TEXT PRIMARY KEY,
    company_id              TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    investee_name           TEXT NOT NULL,
    investee_type           TEXT DEFAULT 'Company',
    investment_type         TEXT DEFAULT 'Equity Shares',
    amount                  NUMERIC(20,2) DEFAULT 0,
    date_of_investment      DATE,
    board_resolution_date   DATE,
    srn_mgb4                TEXT,
    purpose                 TEXT,
    remarks                 TEXT,
    tenant_id               TEXT REFERENCES tenants(id),
    created_by              TEXT,
    created_at              TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS loans_guarantees (
    id                      TEXT PRIMARY KEY,
    company_id              TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    party_name              TEXT NOT NULL,
    party_type              TEXT DEFAULT 'Company',
    transaction_type        TEXT DEFAULT 'Loan',
    amount                  NUMERIC(20,2) DEFAULT 0,
    date_of_transaction     DATE,
    rate_of_interest        NUMERIC(6,2) DEFAULT 0,
    repayment_date          DATE,
    security                TEXT,
    board_resolution_date   DATE,
    outstanding_amount      NUMERIC(20,2) DEFAULT 0,
    status                  TEXT DEFAULT 'Active',
    remarks                 TEXT,
    tenant_id               TEXT REFERENCES tenants(id),
    created_by              TEXT,
    created_at              TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS related_party_transactions (
    id                              TEXT PRIMARY KEY,
    company_id                      TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    party_name                      TEXT NOT NULL,
    relationship                    TEXT,
    nature_of_transaction           TEXT,
    amount                          NUMERIC(20,2) DEFAULT 0,
    date_of_transaction             DATE,
    date_of_board_approval          DATE,
    date_of_shareholders_approval   DATE,
    terms                           TEXT,
    justification                   TEXT,
    remarks                         TEXT,
    tenant_id                       TEXT REFERENCES tenants(id),
    created_by                      TEXT,
    created_at                      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS director_interests (
    id                          TEXT PRIMARY KEY,
    company_id                  TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    director_id                 TEXT,
    director_name               TEXT NOT NULL,
    din                         TEXT,
    entity_name                 TEXT NOT NULL,
    entity_type                 TEXT DEFAULT 'Company',
    nature_of_interest          TEXT,
    date_of_disclosure          DATE,
    date_of_board_resolution    DATE,
    remarks                     TEXT,
    tenant_id                   TEXT REFERENCES tenants(id),
    created_by                  TEXT,
    created_at                  TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS esop_grants (
    id                  TEXT PRIMARY KEY,
    company_id          TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    employee_name       TEXT NOT NULL,
    designation         TEXT,
    employee_id         TEXT,
    grant_date          DATE,
    options_granted     INTEGER DEFAULT 0,
    exercise_price      NUMERIC(10,2) DEFAULT 0,
    vesting_date        DATE,
    vesting_period      TEXT,
    options_exercised   INTEGER DEFAULT 0,
    options_lapsed      INTEGER DEFAULT 0,
    status              TEXT DEFAULT 'Active',
    remarks             TEXT,
    tenant_id           TEXT REFERENCES tenants(id),
    created_by          TEXT,
    created_at          TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS statutory_registers (
    id              TEXT PRIMARY KEY,
    company_id      TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    register_type   TEXT NOT NULL,
    register_name   TEXT NOT NULL,
    section_ref     TEXT,
    entry_date      DATE,
    folio_ref       TEXT,
    entry_data      TEXT,
    remarks         TEXT,
    recorded_by     TEXT,
    tenant_id       TEXT REFERENCES tenants(id),
    created_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS document_templates (
    id              TEXT PRIMARY KEY,
    name            TEXT NOT NULL,
    category        TEXT NOT NULL,
    description     TEXT,
    template_body   TEXT NOT NULL,
    placeholders    JSONB DEFAULT '[]',
    is_system       INTEGER DEFAULT 0,
    is_active       INTEGER DEFAULT 1,
    tenant_id       TEXT REFERENCES tenants(id),
    created_by      TEXT,
    created_at      TIMESTAMPTZ DEFAULT NOW(),
    updated_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS documents (
    id              TEXT PRIMARY KEY,
    company_id      TEXT NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    template_id     TEXT REFERENCES document_templates(id) ON DELETE SET NULL,
    doc_type        TEXT NOT NULL,
    doc_name        TEXT NOT NULL,
    content         TEXT,
    file_path       TEXT,
    module          TEXT,
    entity_id       TEXT,
    date_from       TEXT DEFAULT '',
    date_to         TEXT DEFAULT '',
    tenant_id       TEXT REFERENCES tenants(id),
    generated_by    TEXT,
    created_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS custom_placeholders (
    id          TEXT PRIMARY KEY,
    name        TEXT UNIQUE NOT NULL,
    label       TEXT NOT NULL,
    description TEXT DEFAULT '',
    ph_type     TEXT DEFAULT 'text',
    tenant_id   TEXT REFERENCES tenants(id),
    created_by  TEXT,
    created_at  TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS user_permissions (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    module      TEXT NOT NULL,
    action      TEXT NOT NULL,
    granted     INTEGER NOT NULL DEFAULT 1,
    granted_by  TEXT,
    granted_at  TIMESTAMPTZ DEFAULT NOW(),
    note        TEXT DEFAULT '',
    UNIQUE(user_id, module, action)
);

CREATE TABLE IF NOT EXISTS permission_presets (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    description TEXT DEFAULT '',
    permissions JSONB DEFAULT '[]',
    created_by  TEXT,
    created_at  TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS audit_log (
    id          TEXT PRIMARY KEY,
    user_id     TEXT,
    action      TEXT NOT NULL,
    module      TEXT,
    entity_id   TEXT,
    detail      TEXT,
    ip          TEXT,
    tenant_id   TEXT REFERENCES tenants(id),
    created_at  TIMESTAMPTZ DEFAULT NOW()
);
"""


def init_db():
    """Create all tables and seed initial data."""
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    (BASE_DIR / "data" / "pdfs").mkdir(parents=True, exist_ok=True)
    (BASE_DIR / "data" / "legal_docs").mkdir(parents=True, exist_ok=True)

    conn = get_db(); c = conn.cursor()
    # Run schema statement by statement
    for stmt in [s.strip() for s in SCHEMA_SQL.split(';') if s.strip()]:
        try:
            c.execute(stmt)
        except Exception as e:
            print(f"Schema warning: {e}")
    conn.commit()

    # Seed if empty
    c.execute("SELECT COUNT(*) FROM users")
    r = c.fetchone()
    if (r['count'] if isinstance(r, dict) else r[0]) == 0:
        _seed(c, conn)

    _seed_templates(c, conn)
    conn.commit()
    conn.close()


def ensure_columns():
    """Add any missing columns (safe to run repeatedly)."""
    pass  # Schema handles this with IF NOT EXISTS


def ensure_custom_placeholders_table():
    pass  # Already in main schema


def ensure_permission_tables():
    pass  # Already in main schema


# ── Seed helpers ──────────────────────────────────────────────────────────────
def _seed(c, conn):
    today = date.today()

    # Default tenant
    tid = "default-tenant-001"
    c.execute("""INSERT INTO tenants (id,name,slug,email,phone,address,status,max_users,max_companies)
                 VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING""",
              (tid, "Taxly India Private Limited", "taxlyindia",
               "info@taxlyindia.com", "+91 88829 35471",
               "L-30B, LGF, Malviya Nagar, New Delhi-110017",
               "active", 10, 10))

    # Platform admin
    pa_id = str(uuid.uuid4())
    c.execute("""INSERT INTO users (id,name,email,password,role,is_active,is_platform_admin)
                 VALUES (%s,%s,%s,%s,%s,1,1) ON CONFLICT DO NOTHING""",
              (pa_id, "Platform Admin", "platform@taxlycms.in",
               hash_pw("platform@2025"), "superadmin"))

    # Tenant users
    uid = str(uuid.uuid4())
    for nm, em, pw, rl in [
        ("Super Admin",  "admin@compli.in",   "admin123",   "superadmin"),
        ("Priya Sharma", "manager@compli.in", "manager123", "manager"),
        ("Rahul Verma",  "staff@compli.in",   "staff123",   "staff"),
    ]:
        c.execute("""INSERT INTO users (id,name,email,password,role,tenant_id)
                     VALUES (%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING""",
                  (str(uuid.uuid4()), nm, em, hash_pw(pw), rl, tid))

    # Get the superadmin uid
    c.execute("SELECT id FROM users WHERE email='admin@compli.in'")
    r = c.fetchone(); uid = r['id'] if r else uid

    # Companies
    co1 = str(uuid.uuid4())
    c.execute("""INSERT INTO companies
        (id,name,cin,incorporation_date,registered_office,pan,tan,email,phone,
         authorized_capital,paid_up_capital,business_activity,company_type,roc,tenant_id,created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
        (co1,"Innovate Solutions Pvt Ltd","U72200MH2018PTC309876","2018-03-15",
         "12 Nariman Point, Mumbai 400021","AABCI1234P","MUMA12345A","cs@innovate.in",
         "022-22001234",5000000,1000000,"Software Development","Private Limited","RoC-Mumbai",tid,uid))

    co2 = str(uuid.uuid4())
    c.execute("""INSERT INTO companies
        (id,name,cin,incorporation_date,registered_office,pan,tan,email,phone,
         authorized_capital,paid_up_capital,business_activity,company_type,roc,tenant_id,created_by)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
        (co2,"GreenTech Manufacturing Ltd","L24100DL2015PLC285432","2015-07-22",
         "45 Connaught Place, New Delhi 110001","AABCG5678Q","DELA98765B","legal@greentech.in",
         "011-43210000",50000000,20000000,"Manufacturing - Renewable Energy","Public Limited","RoC-Delhi",tid,uid))

    # Directors
    d1 = str(uuid.uuid4()); d2 = str(uuid.uuid4()); d3 = str(uuid.uuid4())
    c.execute("INSERT INTO directors (id,company_id,name,din,pan,email,mobile,address,designation,date_of_appointment,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (d1,co1,"Arjun Mehta","00123456","ABCPM1234D","arjun@innovate.in","9876543210","123 Bandra West, Mumbai 400050","Managing Director","2018-03-15",tid))
    c.execute("INSERT INTO directors (id,company_id,name,din,pan,email,mobile,address,designation,date_of_appointment,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (d2,co1,"Sneha Patel","00234567","ABCSP5678E","sneha@innovate.in","9123456789","56 Andheri East, Mumbai 400069","Director","2019-06-01",tid))
    c.execute("INSERT INTO directors (id,company_id,name,din,pan,email,mobile,designation,date_of_appointment,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (d3,co2,"Vikram Singh","00345678","ABCVS9012F","vikram@greentech.in","9988776655","Chairman & Managing Director","2015-07-22",tid))

    # KYC
    for did, last_kyc, status in [(d1,"2024-08-15","filed"),(d2,"2023-09-10","overdue"),(d3,"2024-07-20","filed")]:
        c.execute("INSERT INTO director_kyc (id,director_id,last_kyc_date,next_due_date,kyc_status,tenant_id) VALUES (%s,%s,%s,%s,%s,%s)",
                  (str(uuid.uuid4()),did,last_kyc,"2025-09-30",status,tid))

    # Auditors
    aud1 = str(uuid.uuid4())
    near = (today+timedelta(days=25)).isoformat()
    c.execute("INSERT INTO auditors (id,company_id,name,firm_name,membership_no,frn,pan,email,appointment_date,nature_of_appointment,appointment_type,start_date,end_date,srn_adt1,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (aud1,co1,"CA Ramesh Gupta","Gupta & Associates","123456","012345W","ABCPG3456G","ca@gupta-assoc.in","2023-09-30","Regular Auditor","AGM Appointment","2023-09-30",near,"ADT1202312345",tid))
    aud2 = str(uuid.uuid4())
    c.execute("INSERT INTO auditors (id,company_id,name,firm_name,membership_no,frn,pan,email,appointment_date,nature_of_appointment,appointment_type,start_date,end_date,srn_adt1,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (aud2,co2,"CA Priya Nair","Nair & Co","234567","023456E","ABCPN7890H","priya@nairca.in","2024-09-28","Subsequent Auditor","AGM Appointment","2024-09-28","2025-09-27","ADT1202423456",tid))

    # Shareholders
    for nm, folio, shares in [("Arjun Mehta","F0001",50000),("Sneha Patel","F0002",30000),("Angel Investors LLP","F0003",20000)]:
        c.execute("INSERT INTO shareholders (id,company_id,name,folio_no,shares_held,date_of_entry,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                  (str(uuid.uuid4()),co1,nm,folio,shares,"2018-03-15",tid))

    # DSC
    dsc1 = str(uuid.uuid4()); dsc_exp1 = (today+timedelta(days=22)).isoformat()
    c.execute("INSERT INTO dsc_records (id,company_id,director_id,holder_name,holder_type,dsc_class,issued_by,valid_from,valid_to,token_type,custody_status,custody_date,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (dsc1,co1,d1,"Arjun Mehta","Director","Class 3","eMudhra","2022-06-01",dsc_exp1,"ePass 2003","With Us","2024-10-01",tid))
    dsc2 = str(uuid.uuid4()); dsc_exp2 = (today+timedelta(days=180)).isoformat()
    c.execute("INSERT INTO dsc_records (id,company_id,director_id,holder_name,holder_type,dsc_class,issued_by,valid_from,valid_to,token_type,custody_status,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (dsc2,co1,d2,"Sneha Patel","Director","Class 3","NSDL","2023-01-15",dsc_exp2,"USB Token","With Client",tid))

    # Meeting
    c.execute("INSERT INTO meetings (id,company_id,meeting_type,meeting_no,meeting_date,meeting_time,venue,agenda,status,tenant_id,created_by) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (str(uuid.uuid4()),co1,"Board","BM-2025-01",(today+timedelta(days=7)).isoformat(),"11:00","Registered Office, Mumbai","1. Financial Statements\n2. Director KYC Status","scheduled",tid,uid))

    # Tasks
    c.execute("SELECT id FROM users WHERE email='admin@compli.in'")
    r = c.fetchone(); sa_id = r['id'] if r else uid
    c.execute("SELECT id FROM users WHERE email='manager@compli.in'")
    r = c.fetchone(); mgr_id = r['id'] if r else None
    c.execute("SELECT id FROM users WHERE email='staff@compli.in'")
    r = c.fetchone(); staff_id = r['id'] if r else None

    for title, pri, status, due_days, mod in [
        ("File ADT-1 for Innovate Solutions","high","pending",20,"auditor"),
        ("Director KYC for Sneha Patel","critical","pending",5,"director"),
        ("Renew DSC — Arjun Mehta","critical","pending",22,"dsc"),
        ("Prepare AGM Notice","medium","in_progress",30,"meeting"),
        ("ROC Annual Return Filing","high","pending",60,"filing"),
    ]:
        c.execute("INSERT INTO tasks (id,company_id,title,priority,status,due_date,module,task_leader,task_manager,assigned_to,tenant_id,created_by) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                  (str(uuid.uuid4()),co1,title,pri,status,(today+timedelta(days=due_days)).isoformat(),
                   mod,sa_id,mgr_id,staff_id,tid,uid))

    # Alerts
    c.execute("INSERT INTO alerts (id,company_id,entity_type,entity_id,alert_type,title,message,due_date,severity,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (str(uuid.uuid4()),co1,"auditor",aud1,"expiry","Auditor Appointment Expiring Soon",
               f"CA Ramesh Gupta expires {near}. File ADT-1.",near,"critical",tid))
    c.execute("INSERT INTO alerts (id,company_id,entity_type,entity_id,alert_type,title,message,due_date,severity,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (str(uuid.uuid4()),co1,"dsc",dsc1,"dsc_expiry",f"DSC Expiring — Arjun Mehta",
               f"DSC expires {dsc_exp1}. Renew immediately.",dsc_exp1,"critical",tid))
    c.execute("INSERT INTO alerts (id,company_id,entity_type,entity_id,alert_type,title,message,due_date,severity,tenant_id) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
              (str(uuid.uuid4()),co1,"director",d2,"kyc_due","Director KYC Overdue — Sneha Patel",
               "DIR-3 KYC overdue. File immediately.","2025-09-30","critical",tid))
    conn.commit()


def _seed_templates(c, conn):
    c.execute("SELECT COUNT(*) FROM document_templates WHERE is_system=1")
    r = c.fetchone()
    cnt = r['count'] if isinstance(r, dict) else r[0]
    if cnt and int(cnt) > 0:
        return
    templates = [
        ("Board Resolution — General","resolution","General-purpose board resolution.",
         "BOARD RESOLUTION\n\n{{company_name}}\n\nRESOLVED THAT {{resolution_text}}\n\nFor {{company_name}}\n{{director_name}} | {{designation}} | DIN: {{din}}\nDate: {{resolution_date}}",
         ["company_name","resolution_text","director_name","designation","din","resolution_date"]),
        ("AGM Notice","notice","Notice of Annual General Meeting.",
         "NOTICE OF ANNUAL GENERAL MEETING\n\n{{company_name}} | CIN: {{cin}}\n\nNOTICE is hereby given that the {{agm_number}} AGM will be held on {{meeting_date}} at {{venue}}.\n\nBy Order of the Board\n{{director_name}}",
         ["company_name","cin","agm_number","meeting_date","venue","director_name"]),
        ("Director's Report","report","Annual Director's Report under Section 134.",
         "DIRECTORS' REPORT\n\nTo, The Members, {{company_name}}\n\nYour Directors present the {{report_year}} Annual Report.\n\nFor {{company_name}}\n{{director_name}} | DIN: {{din}}",
         ["company_name","report_year","director_name","din"]),
        ("Authorisation Letter","letter","General authorisation letter.",
         "AUTHORISATION LETTER\n\n{{company_name}}\nDate: {{letter_date}}\n\n{{authorised_person}} is authorised to {{purpose}}.\n\nFor {{company_name}}\n{{director_name}}",
         ["company_name","letter_date","authorised_person","purpose","director_name"]),
        ("Board Meeting Minutes","minutes","Minutes of Board of Directors meeting.",
         "MINUTES OF BOARD MEETING\n\n{{company_name}} | Date: {{meeting_date}}\n\nDIRECTORS PRESENT:\n{{directors_present}}\n\nRESOLUTIONS:\n{{resolutions_passed}}",
         ["company_name","meeting_date","directors_present","resolutions_passed"]),
    ]
    for name, cat, desc, body, ph in templates:
        c.execute("""INSERT INTO document_templates
            (id,name,category,description,template_body,placeholders,is_system,tenant_id)
            VALUES (%s,%s,%s,%s,%s,%s,1,'default-tenant-001')
            ON CONFLICT DO NOTHING""",
            (str(uuid.uuid4()), name, cat, desc, body, json.dumps(ph)))
    conn.commit()
