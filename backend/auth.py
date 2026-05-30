"""auth.py — JWT + RBAC + Multitenant  (v2.1 — improved)"""
import os, jwt, uuid, hashlib, json
from datetime import datetime, timedelta
from functools import wraps
from flask import request, jsonify, g
from database import get_db, row

# ── Secret Key ────────────────────────────────────────────────────────────────
SECRET_KEY = os.environ.get("FLASK_SECRET_KEY") or os.environ.get("SECRET_KEY")
if not SECRET_KEY:
    import secrets as _sec
    SECRET_KEY = _sec.token_hex(32)
    import warnings
    warnings.warn(
        "FLASK_SECRET_KEY not set — using random key (tokens won't survive restarts)",
        stacklevel=2,
    )
TOKEN_EXPIRY_HOURS = int(os.environ.get("TOKEN_EXPIRY_HOURS", "12"))

# ── Role hierarchy ────────────────────────────────────────────────────────────
ROLE_HIERARCHY = {"superadmin": 3, "manager": 2, "staff": 1}

# ── Default permissions per role ──────────────────────────────────────────────
DEFAULT_PERMISSIONS = {
    "company":     {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager"], "update": ["superadmin","manager"], "delete": ["superadmin"]},
    "director":    {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager"], "update": ["superadmin","manager"], "delete": ["superadmin"]},
    "auditor":     {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager"], "update": ["superadmin","manager"], "delete": ["superadmin"]},
    "shareholder": {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager"], "update": ["superadmin","manager"], "delete": ["superadmin"]},
    "dsc":         {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager"], "update": ["superadmin","manager","staff"], "delete": ["superadmin"]},
    "meeting":     {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager","staff"], "update": ["superadmin","manager"], "delete": ["superadmin"]},
    "document":    {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager","staff"], "delete": ["superadmin","manager"]},
    "task":        {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager","staff"], "update": ["superadmin","manager","staff"], "delete": ["superadmin","manager"]},
    "user":        {"read": ["superadmin","manager"], "create": ["superadmin"], "update": ["superadmin"], "delete": ["superadmin"]},
    "alert":       {"read": ["superadmin","manager","staff"], "update": ["superadmin","manager","staff"]},
    "register":    {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager"], "update": ["superadmin","manager"], "delete": ["superadmin"]},
    "template":    {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager"], "update": ["superadmin"], "delete": ["superadmin"]},
    "export":      {"excel": ["superadmin","manager","staff"], "pdf": ["superadmin","manager","staff"], "bulk_upload": ["superadmin","manager"]},
    "charge":      {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager"], "update": ["superadmin","manager"], "delete": ["superadmin"]},
    "investment":  {"read": ["superadmin","manager","staff"], "create": ["superadmin","manager"], "update": ["superadmin","manager"], "delete": ["superadmin"]},
}

ALL_MODULES = {
    "company":     {"label": "Companies",           "actions": ["read","create","update","delete"]},
    "director":    {"label": "Directors",            "actions": ["read","create","update","delete"]},
    "auditor":     {"label": "Auditors",             "actions": ["read","create","update","delete"]},
    "shareholder": {"label": "Shareholders",         "actions": ["read","create","update","delete"]},
    "dsc":         {"label": "DSC Records",          "actions": ["read","create","update","delete"]},
    "meeting":     {"label": "Meetings",             "actions": ["read","create","update","delete"]},
    "document":    {"label": "Documents",            "actions": ["read","create","delete"]},
    "task":        {"label": "Tasks",                "actions": ["read","create","update","delete"]},
    "register":    {"label": "Statutory Registers",  "actions": ["read","create","update","delete"]},
    "template":    {"label": "Doc Templates",        "actions": ["read","create","update","delete"]},
    "alert":       {"label": "Alerts",               "actions": ["read","update"]},
    "charge":      {"label": "Charges / CHG-1",      "actions": ["read","create","update","delete"]},
    "investment":  {"label": "Investments / MBP-3",  "actions": ["read","create","update","delete"]},
    "export":      {"label": "Reports & Export",     "actions": ["excel","pdf","bulk_upload"]},
    "user":        {"label": "User Management",      "actions": ["read","create","update","delete"]},
}


def hash_pw(pw: str) -> str:
    """Hash password using bcrypt (preferred) or SHA-256 fallback."""
    try:
        import bcrypt as _bcrypt
        return _bcrypt.hashpw(pw.encode(), _bcrypt.gensalt()).decode()
    except ImportError:
        return hashlib.sha256(pw.encode()).hexdigest()


def verify_pw(plain: str, stored: str) -> bool:
    """Verify a password against its stored hash (bcrypt or SHA-256)."""
    if not plain or not stored:
        return False
    if stored.startswith("$2b$") or stored.startswith("$2a$"):
        try:
            import bcrypt as _bcrypt
            return _bcrypt.checkpw(plain.encode(), stored.encode())
        except Exception:
            return False
    # SHA-256 legacy
    return hashlib.sha256(plain.encode()).hexdigest() == stored


def make_token(user_id, role, name, tenant_id=None, is_platform_admin=False):
    payload = {
        "sub": user_id,
        "role": role,
        "name": name,
        "tid": tenant_id,
        "pa": is_platform_admin,
        "iat": datetime.utcnow(),
        "exp": datetime.utcnow() + timedelta(hours=TOKEN_EXPIRY_HOURS),
    }
    return jwt.encode(payload, SECRET_KEY, algorithm="HS256")


def decode_token(token):
    try:
        return jwt.decode(token, SECRET_KEY, algorithms=["HS256"])
    except jwt.ExpiredSignatureError:
        return None
    except jwt.InvalidTokenError:
        return None
    except Exception:
        return None


def get_token():
    auth = request.headers.get("Authorization", "")
    return auth[7:] if auth.startswith("Bearer ") else request.cookies.get("token")


def load_user_permissions(user_id):
    """Load per-user permission overrides from DB."""
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute(
            "SELECT module, action, granted FROM user_permissions WHERE user_id=%s",
            (user_id,),
        )
        overrides = {}
        for _r in c.fetchall():
            if isinstance(_r, dict):
                vals = list(_r.values())
                overrides[(vals[0], vals[1])] = bool(vals[2])
            else:
                overrides[(_r[0], _r[1])] = bool(_r[2])
        conn.close()
        return overrides
    except Exception as _e:
        import logging
        logging.getLogger(__name__).debug(f"load_user_permissions: {_e}")
        return {}


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = get_token()
        if not token:
            return jsonify({"error": "Authentication required"}), 401
        payload = decode_token(token)
        if not payload:
            return jsonify({"error": "Invalid or expired token"}), 401
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE id=%s AND is_active=1", (payload["sub"],))
        user = row(c.fetchone())
        conn.close()
        if not user:
            return jsonify({"error": "User not found or disabled"}), 401
        g.user = user
        g.user_id = user["id"]
        g.role = user["role"]
        g.tenant_id = user.get("tenant_id")
        g.is_platform_admin = bool(user.get("is_platform_admin", 0))
        g.perm_overrides = load_user_permissions(user["id"])
        return f(*args, **kwargs)
    return decorated


def platform_admin_required(f):
    """Only platform admin can access this endpoint."""
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not g.is_platform_admin:
            return jsonify({"error": "Platform admin access required"}), 403
        return f(*args, **kwargs)
    return decorated


def require_role(*roles):
    def decorator(f):
        @wraps(f)
        @login_required
        def decorated(*args, **kwargs):
            if g.role not in roles:
                return jsonify({"error": "Access denied"}), 403
            return f(*args, **kwargs)
        return decorated
    return decorator


def can(module, action):
    role = getattr(g, "role", None)
    if not role:
        return False
    if getattr(g, "is_platform_admin", False):
        return True
    if role == "superadmin":
        return True
    overrides = getattr(g, "perm_overrides", {})
    key = (module, action)
    if key in overrides:
        return overrides[key]
    return role in (DEFAULT_PERMISSIONS.get(module, {}).get(action, []))


def tenant_scope(query, alias=""):
    """Append tenant_id filter — returns (query, params_list).
    DEPRECATED: prefer passing tenant_id as %s param directly.
    """
    tid = getattr(g, "tenant_id", None)
    if tid:
        col = f"{alias}tenant_id" if alias else "tenant_id"
        return query + f" AND {col} = %s", [tid]
    return query, []


# ── Simple in-memory rate limiter (use Redis in prod) ────────────────────────
import time, threading

_rl_lock = threading.Lock()
_rl_store: dict = {}


def _rate_limit(key: str, max_calls: int = 5, window_secs: int = 60) -> bool:
    """Returns True if allowed, False if rate-limited."""
    now = time.time()
    with _rl_lock:
        calls = [t for t in _rl_store.get(key, []) if now - t < window_secs]
        if len(calls) >= max_calls:
            return False
        calls.append(now)
        _rl_store[key] = calls
    return True


def rate_limit_login(ip: str) -> bool:
    """Max 10 login attempts per IP per minute."""
    return _rate_limit(f"login:{ip}", max_calls=10, window_secs=60)
