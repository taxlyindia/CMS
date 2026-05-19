# TaxlyCMS — Corporate Compliance Management System

Multi-tenant SaaS for Companies Act 2013 compliance management.

## Project Structure

```
taxlycms/
├── backend/              # Flask API server (Python)
│   ├── app.py            # Main Flask application (3000+ lines)
│   ├── auth.py           # JWT auth + RBAC + multitenant
│   ├── database.py       # PostgreSQL connection + helpers
│   ├── compliance.py     # Compliance checks engine
│   ├── requirements.txt  # Python dependencies
│   ├── gunicorn.conf.py  # Production server config
│   └── data/             # Uploads, PDFs, logos (gitignored)
│
├── frontend/             # Single-page application (HTML/JS/CSS)
│   └── index.html        # Full SPA — no build step needed
│
├── sql/                  # Database scripts
│   ├── 001_schema.sql    # Full PostgreSQL schema (run first)
│   └── 002_seed.sql      # Demo data (run after schema)
│
├── .env.example          # Environment variable template
├── .gitignore
├── Procfile              # For Heroku/Railway-style deployment
└── README.md
```

## Quick Start — Hostinger

### 1. Database Setup
```sql
-- In Hostinger phpPgAdmin or via psql:
\i sql/001_schema.sql
\i sql/002_seed.sql
```

### 2. Configure Environment
```bash
cp .env.example .env
# Edit .env with your Hostinger PostgreSQL credentials
```

### 3. Install & Run
```bash
cd backend
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
gunicorn app:app -c gunicorn.conf.py
```

## Login Credentials (Demo)

| Role | Email | Password |
|------|-------|----------|
| Platform Admin | platform@taxlycms.in | platform@2025 |
| Tenant Admin | admin@compli.in | admin123 |
| Manager | manager@compli.in | manager123 |
| Staff | staff@compli.in | staff123 |

⚠️ Change all passwords before going live.

## Tech Stack

- **Backend**: Python / Flask / psycopg2
- **Frontend**: Vanilla HTML + CSS + JS (no build step)
- **Database**: PostgreSQL 14+
- **Auth**: JWT (PyJWT)
- **Docs**: ReportLab PDF + XlsxWriter Excel
- **Server**: Gunicorn + Nginx

## Developed by

Taxly India Private Limited  
L-30B, LGF, Malviya Nagar, New Delhi – 110017  
+91 88829 35471 · info@taxlyindia.com
