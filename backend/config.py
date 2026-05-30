"""
TaxlyCMS — Central Configuration
All environment variables with defaults.
Import this in app.py: from config import Config
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).parent


class Config:
    # ── Application ──────────────────────────────────────────────────────
    SECRET_KEY       = os.environ.get("SECRET_KEY",        "taxlycms-dev-secret-change-in-prod")
    FRONTEND_URL     = os.environ.get("FRONTEND_URL",      "http://localhost:5000")
    ALLOWED_ORIGINS  = os.environ.get("ALLOWED_ORIGINS",   "*")
    COOKIE_SECURE    = os.environ.get("COOKIE_SECURE",     "0") == "1"
    TOKEN_EXPIRY_HOURS = int(os.environ.get("TOKEN_EXPIRY_HOURS", "24"))
    DEBUG            = os.environ.get("DEBUG",             "0") == "1"

    # ── Database ─────────────────────────────────────────────────────────
    DATABASE_URL     = os.environ.get("DATABASE_URL",      "")   # PostgreSQL DSN
    # SQLite fallback (dev only)
    SQLITE_PATH      = os.environ.get("SQLITE_PATH",       str(BASE_DIR / "data" / "taxlycms.db"))

    # ── File Storage ─────────────────────────────────────────────────────
    UPLOAD_DIR       = BASE_DIR / "data" / "uploads"
    ATTACH_DIR       = BASE_DIR / "data" / "attachments"
    WORD_TPL_DIR     = BASE_DIR / "data" / "word_templates"
    ATTACH_MAX_MB    = int(os.environ.get("ATTACH_MAX_MB", "20"))

    # ── Email (SMTP) ─────────────────────────────────────────────────────
    MAIL_SERVER      = os.environ.get("MAIL_SERVER",       "smtp.gmail.com")
    MAIL_PORT        = int(os.environ.get("MAIL_PORT",     "587"))
    MAIL_USE_TLS     = os.environ.get("MAIL_USE_TLS",      "1") == "1"
    MAIL_USERNAME    = os.environ.get("MAIL_USERNAME",     "")
    MAIL_PASSWORD    = os.environ.get("MAIL_PASSWORD",     "")
    MAIL_FROM        = os.environ.get("MAIL_FROM",         "") or os.environ.get("MAIL_USERNAME", "noreply@taxlycms.in")
    MAIL_ENABLED     = bool(os.environ.get("MAIL_USERNAME", ""))

    # ── WhatsApp (Meta Cloud API) ─────────────────────────────────────────
    WA_TOKEN         = os.environ.get("WA_TOKEN",          "")   # Meta access token
    WA_PHONE_ID      = os.environ.get("WA_PHONE_ID",       "")   # Meta phone number ID
    WA_ENABLED       = bool(os.environ.get("WA_TOKEN", ""))

    # ── Twilio (SMS / WhatsApp fallback) ─────────────────────────────────
    TWILIO_ACCOUNT_SID   = os.environ.get("TWILIO_ACCOUNT_SID",   "")
    TWILIO_AUTH_TOKEN    = os.environ.get("TWILIO_AUTH_TOKEN",     "")
    TWILIO_FROM_NUMBER   = os.environ.get("TWILIO_FROM_NUMBER",    "")
    TWILIO_WHATSAPP_FROM = os.environ.get("TWILIO_WHATSAPP_FROM",  "whatsapp:+14155238886")

    # ── E-Signature ───────────────────────────────────────────────────────
    ESIGN_API_KEY        = os.environ.get("ESIGN_API_KEY",         "")   # Leegality / DocuSign key
    ESIGN_PROVIDER       = os.environ.get("ESIGN_PROVIDER",        "internal")
    LEEGALITY_API_KEY    = os.environ.get("LEEGALITY_API_KEY",     "")
    DOCUSIGN_CLIENT_ID   = os.environ.get("DOCUSIGN_CLIENT_ID",    "")
    DOCUSIGN_SECRET      = os.environ.get("DOCUSIGN_SECRET",       "")

    # ── AI (Anthropic) ────────────────────────────────────────────────────
    ANTHROPIC_API_KEY    = os.environ.get("ANTHROPIC_API_KEY",     "")

    # ── JWT ───────────────────────────────────────────────────────────────
    JWT_SECRET           = os.environ.get("JWT_SECRET", SECRET_KEY)

    @classmethod
    def summary(cls):
        """Print config summary (safe — no secrets)."""
        return {
            "MAIL_ENABLED":  cls.MAIL_ENABLED,
            "MAIL_SERVER":   cls.MAIL_SERVER,
            "MAIL_FROM":     cls.MAIL_FROM,
            "WA_ENABLED":    cls.WA_ENABLED,
            "WA_PHONE_ID":   cls.WA_PHONE_ID[:8]+"…" if cls.WA_PHONE_ID else "",
            "ESIGN_PROVIDER": cls.ESIGN_PROVIDER,
            "ANTHROPIC_KEY": "set" if cls.ANTHROPIC_API_KEY else "not set",
            "DATABASE_URL":  "set" if cls.DATABASE_URL else "sqlite fallback",
        }


# ── .env file loader (dev convenience) ───────────────────────────────────
def load_dotenv(env_path=None):
    """Load a .env file into os.environ (no external dependency)."""
    path = env_path or (BASE_DIR / ".env")
    if not path.exists():
        return
    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = val


# Auto-load .env if present
load_dotenv()
