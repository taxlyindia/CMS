-- ============================================================
-- TaxlyCMS — PostgreSQL Schema  v2.0
-- Companies Act 2013 Compliance CRM — Multi-tenant
--
-- Usage:
--   psql $DATABASE_URL -f sql/001_schema.sql
--   (or run via Hostinger phpPgAdmin query tool)
-- ============================================================

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

-- ── Performance Indexes ──────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_companies_tenant     ON companies(tenant_id);
CREATE INDEX IF NOT EXISTS idx_directors_company    ON directors(company_id);
CREATE INDEX IF NOT EXISTS idx_directors_tenant     ON directors(tenant_id);
CREATE INDEX IF NOT EXISTS idx_auditors_company     ON auditors(company_id);
CREATE INDEX IF NOT EXISTS idx_shareholders_company ON shareholders(company_id);
CREATE INDEX IF NOT EXISTS idx_tasks_tenant         ON tasks(tenant_id);
CREATE INDEX IF NOT EXISTS idx_tasks_assigned       ON tasks(assigned_to);
CREATE INDEX IF NOT EXISTS idx_tasks_leader         ON tasks(task_leader);
CREATE INDEX IF NOT EXISTS idx_tasks_manager        ON tasks(task_manager);
CREATE INDEX IF NOT EXISTS idx_tasks_company        ON tasks(company_id);
CREATE INDEX IF NOT EXISTS idx_alerts_tenant        ON alerts(tenant_id);
CREATE INDEX IF NOT EXISTS idx_alerts_status        ON alerts(status, tenant_id);
CREATE INDEX IF NOT EXISTS idx_alerts_company       ON alerts(company_id);
CREATE INDEX IF NOT EXISTS idx_users_tenant         ON users(tenant_id);
CREATE INDEX IF NOT EXISTS idx_users_email          ON users(email);
CREATE INDEX IF NOT EXISTS idx_meetings_company     ON meetings(company_id);
CREATE INDEX IF NOT EXISTS idx_meetings_date        ON meetings(meeting_date);
CREATE INDEX IF NOT EXISTS idx_documents_company    ON documents(company_id);
CREATE INDEX IF NOT EXISTS idx_dsc_company          ON dsc_records(company_id);
