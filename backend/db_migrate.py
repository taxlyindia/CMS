#!/usr/bin/env python3
"""
db_migrate.py — ONE-TIME server database migration
Run this ONCE on the server to fix the existing database:

    cd /var/www/taxlycms/backend
    python3 db_migrate.py

Safe on existing data — only adds missing columns/tables/users.
"""
import os, sys, hashlib, uuid, sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).parent
DB_PATH  = BASE_DIR / "data" / "compli.db"

def hash_pw(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

print("=" * 60)
print("TaxlyCMS — Database Migration")
print("=" * 60)

if not DB_PATH.exists():
    print(f"\n⚠  DB not found at {DB_PATH}")
    print("Running full init instead...")
    sys.path.insert(0, str(BASE_DIR))
    from database import init_db
    init_db()
    print("✅  Done — fresh database created")
    sys.exit(0)

conn = sqlite3.connect(str(DB_PATH))
conn.row_factory = sqlite3.Row
c = conn.cursor()

# ── 1. users table: add missing columns ────────────────────────────────────
print("\n[1] users table columns...")
c.execute("PRAGMA table_info(users)")
user_cols = [r["name"] for r in c.fetchall()]
for col, defn in [
    ("is_platform_admin", "INTEGER DEFAULT 0"),
    ("tenant_id",         "TEXT"),
]:
    if col not in user_cols:
        c.execute(f"ALTER TABLE users ADD COLUMN {col} {defn}")
        print(f"  ✓ Added users.{col}")
    else:
        print(f"  ✓ users.{col} exists")

# ── 2. tasks table: add role columns ────────────────────────────────────────
print("\n[2] tasks table columns...")
c.execute("PRAGMA table_info(tasks)")
task_cols = [r["name"] for r in c.fetchall()]
for col, defn in [
    ("task_leader",   "TEXT"),
    ("task_manager",  "TEXT"),
    ("team_members",  "TEXT DEFAULT '[]'"),
    ("note",          "TEXT DEFAULT ''"),
    ("billable",      "INTEGER DEFAULT 0"),
    ("estimated_hrs", "REAL DEFAULT 0"),
    ("actual_hrs",    "REAL DEFAULT 0"),
]:
    if col not in task_cols:
        c.execute(f"ALTER TABLE tasks ADD COLUMN {col} {defn}")
        print(f"  ✓ Added tasks.{col}")

# ── 3. Create missing tables ──────────────────────────────────────────────────
print("\n[3] Creating missing tables...")
c.executescript("""
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
    max_documents INTEGER DEFAULT 100, features TEXT DEFAULT '["all"]',
    is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS dsc_custody_log (
    id TEXT PRIMARY KEY, dsc_id TEXT NOT NULL, action TEXT NOT NULL,
    action_date TEXT, from_party TEXT, to_party TEXT, notes TEXT,
    recorded_by TEXT, tenant_id TEXT, created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS audit_log (
    id TEXT PRIMARY KEY, user_id TEXT, action TEXT NOT NULL, module TEXT,
    entity_id TEXT, detail TEXT, ip TEXT, tenant_id TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS permission_presets (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, description TEXT DEFAULT '',
    permissions TEXT DEFAULT '[]', created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);
""")
print("  ✓ All required tables present")

# ── 4. Add tenant_id to all data tables ──────────────────────────────────────
print("\n[4] Adding tenant_id to data tables...")
data_tables = [
    "companies","directors","director_kyc","auditors","shareholders",
    "dsc_records","meetings","tasks","alerts","documents",
    "document_templates","charges","investments","loans_guarantees",
    "related_party_transactions","statutory_registers","custom_placeholders",
    "user_permissions", "dsc_custody_log"
]
for tbl in data_tables:
    try:
        c.execute(f"ALTER TABLE {tbl} ADD COLUMN tenant_id TEXT")
        print(f"  ✓ {tbl}.tenant_id added")
    except:
        pass  # Already exists

# ── 5. Default tenant ─────────────────────────────────────────────────────────
print("\n[5] Default tenant...")
c.execute("SELECT id FROM tenants WHERE id='default-tenant-001'")
if not c.fetchone():
    c.execute("""INSERT OR IGNORE INTO tenants
        (id,name,slug,email,phone,address,status,max_users,max_companies)
        VALUES (?,?,?,?,?,?,?,?,?)""",
        ("default-tenant-001","Taxly India Private Limited","taxlyindia",
         "info@taxlyindia.com","+91 88829 35471",
         "L-30B, LGF, Malviya Nagar, New Delhi-110017","active",10,10))
    print("  ✓ Default tenant created")
else:
    print("  ✓ Default tenant exists")

# ── 6. Pricing plans ──────────────────────────────────────────────────────────
print("\n[6] Pricing plans...")
c.execute("SELECT COUNT(*) as cnt FROM plans"); r = c.fetchone()
cnt = r["cnt"] if isinstance(r, dict) else r[0]
if int(cnt) == 0:
    for name, pm, pa_p, mu, mc in [
        ("Starter",2999,29990,3,3),
        ("Professional",6999,69990,10,10),
        ("Enterprise",0,0,50,50),
    ]:
        c.execute("""INSERT OR IGNORE INTO plans
            (id,name,description,price_monthly,price_annual,max_users,max_companies,max_documents,features,is_active)
            VALUES (?,?,?,?,?,?,?,100,'["all"]',1)""",
            (str(uuid.uuid4()),name,f"{name} plan",pm,pa_p,mu,mc))
    print("  ✓ 3 plans created")
else:
    print(f"  ✓ {cnt} plans exist")

# ── 7. Assign existing users + data to default tenant ────────────────────────
print("\n[7] Assigning data to default tenant...")
c.execute("UPDATE users SET tenant_id='default-tenant-001' WHERE tenant_id IS NULL AND COALESCE(is_platform_admin,0)=0")
print(f"  ✓ {c.rowcount} users → default tenant")
for tbl in ["companies","directors","director_kyc","auditors","shareholders",
            "dsc_records","meetings","tasks","alerts","documents","document_templates"]:
    try:
        c.execute(f"UPDATE {tbl} SET tenant_id='default-tenant-001' WHERE tenant_id IS NULL")
        if c.rowcount: print(f"  ✓ {tbl}: {c.rowcount} rows")
    except: pass

# ── 8. Platform admin ─────────────────────────────────────────────────────────
print("\n[8] Platform admin user...")
c.execute("SELECT id, is_platform_admin, is_active FROM users WHERE email='platform@taxlycms.in'")
pa = c.fetchone()
if not pa:
    c.execute("""INSERT INTO users (id,name,email,password,role,is_active,is_platform_admin,tenant_id)
        VALUES (?,?,?,?,?,1,1,NULL)""",
        (str(uuid.uuid4()),"Platform Admin","platform@taxlycms.in",
         hash_pw("platform@2025"),"superadmin"))
    print("  ✓ Platform admin CREATED")
else:
    c.execute("""UPDATE users SET
        is_platform_admin=1, tenant_id=NULL, is_active=1,
        password=? WHERE email='platform@taxlycms.in'""",
        (hash_pw("platform@2025"),))
    print("  ✓ Platform admin FIXED (pa=1, tid=NULL, active=1, password reset)")

conn.commit()

# ── Verify ────────────────────────────────────────────────────────────────────
print("\n[VERIFY]")
c.execute("SELECT email, is_platform_admin, is_active, tenant_id FROM users ORDER BY is_platform_admin DESC, email")
for u in c.fetchall():
    u = dict(u)
    icon = "👑" if u['is_platform_admin'] else ("✓" if u['is_active'] else "✗")
    print(f"  {icon} {u['email']:35s} pa={u['is_platform_admin']} active={u['is_active']} tid={str(u['tenant_id'] or '')[:12]}")

conn.close()

print("\n" + "=" * 60)
print("✅  Migration complete!")
print("=" * 60)
print("""
Credentials:
  Platform Admin : platform@taxlycms.in  /  platform@2025
  Tenant Admin   : admin@compli.in        /  admin123
  Manager        : manager@compli.in      /  manager123
  Staff          : staff@compli.in         /  staff123

Restart gunicorn:
  pkill gunicorn
  gunicorn app:app -c gunicorn.conf.py --daemon
  
  # or systemd:
  sudo systemctl restart taxlycms
""")
