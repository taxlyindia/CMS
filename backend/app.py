"""
app.py — Companies Act 2013 Compliance CRM  v2.0
Flask REST API + JWT + RBAC
All 7 new features: Company Master PDF, MCA Credentials, DSC Records,
  Auditor Nature, All Statutory Registers, Document Template Store
"""
import os, uuid, json, io
import xlsxwriter
import openpyxl
from pathlib import Path
from datetime import datetime, date, timedelta
from flask import Flask, request, jsonify, send_file, g, send_from_directory
from werkzeug.utils import secure_filename

from database import get_db, row, rows, hash_pw, init_db, ensure_custom_placeholders_table, ensure_permission_tables, ensure_custom_placeholders_table
from auth import login_required, require_role, platform_admin_required, make_token, hash_pw as auth_hash, can, get_token, DEFAULT_PERMISSIONS, ALL_MODULES, tenant_scope
from compliance import (run_compliance_checks, generate_document, build_context,
                        extract_placeholders, get_active_entities, AUTO_FILLED,
                        generate_company_master_pdf, generate_register_pdf,
                        REGISTER_DEFINITIONS)

BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "data" / "uploads"
ALLOWED_EXT = {".pdf",".doc",".docx",".jpg",".jpeg",".png"}

app = Flask(__name__, static_folder=str(BASE_DIR/"static"))
app.config["MAX_CONTENT_LENGTH"] = 10*1024*1024

@app.errorhandler(500)
def handle_500(e):
    import traceback as _tb
    tb = _tb.format_exc()
    app.logger.error("500 ERROR:\n" + tb)
    return jsonify({"error": "Server error — " + str(e), "detail": tb[-800:]}), 500

# ══ HELPERS ══════════════════════════════════════════════════════════════════
def _kyc_due():
    today = date.today()
    yr = today.year if today.month>=4 else today.year-1
    return f"{yr+1}-09-30"

def _kyc_status(due_str):
    if not due_str: return "pending"
    d = date.fromisoformat(due_str[:10])
    days = (d-date.today()).days
    if days < 0: return "overdue"
    if days <= 30: return "due_soon"
    return "compliant"

def _enrich_auditor(a):
    if a.get("end_date"):
        d=date.fromisoformat(a["end_date"][:10]); days=(d-date.today()).days
        a["days_to_expiry"]=days
        a["expiry_status"]="expired" if days<0 else ("expiring_soon" if days<=30 else "valid")
    return a

# ══ STATIC / INDEX ═══════════════════════════════════════════════════════════
@app.route("/")
def index():
    from flask import make_response
    resp = make_response(send_from_directory(str(BASE_DIR), "index.html"))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"]        = "no-cache"
    resp.headers["Expires"]       = "0"
    return resp

@app.route("/static/<path:fn>")
def statics(fn): return send_from_directory(str(BASE_DIR/"static"),fn)

# ══ AUTH ═════════════════════════════════════════════════════════════════════
@app.route("/api/auth/login", methods=["POST"])
def login():
    d=request.get_json(silent=True, force=True) or {}
    email=(d.get("email") or "").strip().lower(); pw=d.get("password") or ""
    if not email or not pw: return jsonify({"error":"Email and password required"}),400
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT * FROM users WHERE email=%s AND is_active=1",(email,))
    user=row(c.fetchone())
    if not user or user["password"]!=hash_pw(pw): conn.close(); return jsonify({"error":"Invalid credentials"}),401
    is_pa = bool(user.get("is_platform_admin", 0))
    if not is_pa and user.get("tenant_id"):
        c.execute("SELECT status FROM tenants WHERE id=%s", (user["tenant_id"],))
        t = c.fetchone()
        if t and t[0] not in ("active",):
            conn.close(); return jsonify({"error":f"Account {t[0]}. Contact support."}),403
    c.execute("UPDATE users SET last_login=NOW() WHERE id=%s",(user["id"],))
    conn.commit(); conn.close()
    return jsonify({
        "token": make_token(user["id"],user["role"],user["name"],user.get("tenant_id"),is_pa),
        "user":  {"id":user["id"],"name":user["name"],"email":user["email"],"role":user["role"]},
        "is_platform_admin": is_pa,
        "tenant_id": user.get("tenant_id"),
    })

@app.route("/api/auth/me")
@login_required
def me(): return jsonify({"id":g.user["id"],"name":g.user["name"],"email":g.user["email"],"role":g.user["role"],"is_platform_admin":bool(g.user.get("is_platform_admin",0)),"tenant_id":g.user.get("tenant_id")})

@app.route("/api/auth/change-password", methods=["POST"])
@login_required
def change_password():
    d=request.get_json(silent=True, force=True) or {}
    old=d.get("old_password",""); new=d.get("new_password","")
    if not old or not new or len(new)<6: return jsonify({"error":"Valid passwords required (min 6 chars)"}),400
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT password FROM users WHERE id=%s",(g.user_id,))
    r=c.fetchone()
    if not r or r[0]!=hash_pw(old): conn.close(); return jsonify({"error":"Current password incorrect"}),400
    c.execute("UPDATE users SET password=%s WHERE id=%s",(hash_pw(new),g.user_id))
    conn.commit(); conn.close(); return jsonify({"success":True})

# ══ USERS ════════════════════════════════════════════════════════════════════
@app.route("/api/users")
@require_role("superadmin", "manager")
def list_users():
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT id,name,email,role,is_active,created_at,last_login FROM users ORDER BY name")
    return jsonify(rows(c.fetchall()))

@app.route("/api/users", methods=["POST"])
@require_role("superadmin")
def create_user():
    d = request.get_json(silent=True, force=True) or {}
    if not d.get("name"):     return jsonify({"error": "Name is required"}), 400
    if not d.get("email"):    return jsonify({"error": "Email is required"}), 400
    if not d.get("password"): return jsonify({"error": "Password is required"}), 400
    if len(d["password"]) < 6: return jsonify({"error": "Password must be at least 6 characters"}), 400
    role = d.get("role", "staff")
    # Superadmin can create any role; validate role value
    if role not in ("staff", "manager", "superadmin"):
        return jsonify({"error": "Invalid role. Must be staff, manager, or superadmin"}), 400
    uid = str(uuid.uuid4()); conn = get_db(); c = conn.cursor()
    try:
        c.execute(
            "INSERT INTO users (id,name,email,password,role,is_active,tenant_id) VALUES (%s,%s,?,%s,?,1,%s)",
            (uid, d["name"].strip(), d["email"].strip().lower(), hash_pw(d["password"]), role, g.tenant_id)
        )
        conn.commit()
    except Exception as e:
        conn.close()
        if "UNIQUE" in str(e): return jsonify({"error": "Email already exists"}), 400
        return jsonify({"error": str(e)}), 400
    c.execute("SELECT id,name,email,role,is_active,created_at FROM users WHERE id=%s", (uid,))
    result = row(c.fetchone()); conn.close()
    return jsonify(result), 201

@app.route("/api/users/<uid>", methods=["PUT"])
@require_role("superadmin")
def update_user(uid):
    d = request.get_json(silent=True, force=True) or {}
    # Prevent superadmin from editing themselves into a lower role
    if uid == g.user_id and "role" in d and d["role"] != "superadmin":
        return jsonify({"error": "Cannot change your own role"}), 400
    conn = get_db(); c = conn.cursor()
    fields = {k: d[k] for k in ["name", "email", "role", "is_active"] if k in d}
    if "password" in d and d["password"]:
        if len(d["password"]) < 6:
            return jsonify({"error": "Password must be at least 6 characters"}), 400
        fields["password"] = hash_pw(d["password"])
    if not fields: return jsonify({"error": "Nothing to update"}), 400
    c.execute(
        f"UPDATE users SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",
        list(fields.values()) + [uid]
    )
    conn.commit()
    c.execute("SELECT id,name,email,role,is_active,created_at,last_login FROM users WHERE id=%s", (uid,))
    result = row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/users/<uid>/pending-tasks")
@require_role("superadmin")
def user_pending_tasks(uid):
    """Return count and list of pending tasks assigned to/led/managed by this user."""
    conn = get_db(); c = conn.cursor()
    c.execute("""
        SELECT t.id, t.title, t.status, t.priority, t.due_date,
               co.name AS company_name,
               CASE
                 WHEN t.task_leader=%s  THEN 'leader'
                 WHEN t.task_manager=%s THEN 'manager'
                 ELSE 'assignee'
               END AS user_role_in_task
        FROM tasks t
        LEFT JOIN companies co ON t.company_id=co.id
        WHERE (t.task_leader=%s OR t.task_manager=%s OR t.assigned_to=%s)
          AND t.status NOT IN ('completed','cancelled')
        ORDER BY CASE t.priority WHEN 'critical' THEN 1 WHEN 'high' THEN 2 WHEN 'medium' THEN 3 ELSE 4 END
    """, (uid, uid, uid, uid, uid))
    cols  = [d[0] for d in c.description]
    tasks = [dict(zip(cols, r)) for r in c.fetchall()]
    conn.close()
    return jsonify({"count": len(tasks), "tasks": tasks})

@app.route("/api/users/<uid>/reassign-tasks", methods=["POST"])
@require_role("superadmin")
def reassign_tasks(uid):
    """Reassign all pending tasks from outgoing user to a replacement user."""
    d   = request.get_json(silent=True, force=True) or {}
    to  = d.get("to_user_id", "").strip()
    if not to: return jsonify({"error": "to_user_id is required"}), 400
    if to == uid: return jsonify({"error": "Cannot reassign to the same user"}), 400

    conn = get_db(); c = conn.cursor()
    # Verify target user exists
    c.execute("SELECT id, role FROM users WHERE id=%s AND is_active=1", (to,))
    target = c.fetchone()
    if not target: conn.close(); return jsonify({"error": "Target user not found"}), 404

    # Reassign each role column
    c.execute("UPDATE tasks SET task_leader=%s  WHERE task_leader=%s  AND status NOT IN ('completed','cancelled')", (to, uid))
    leader_cnt = c.rowcount
    c.execute("UPDATE tasks SET task_manager=%s WHERE task_manager=%s AND status NOT IN ('completed','cancelled')", (to, uid))
    manager_cnt = c.rowcount
    c.execute("UPDATE tasks SET assigned_to=%s  WHERE assigned_to=%s  AND status NOT IN ('completed','cancelled')", (to, uid))
    assignee_cnt = c.rowcount
    conn.commit(); conn.close()

    return jsonify({
        "success":  True,
        "reassigned": leader_cnt + manager_cnt + assignee_cnt,
        "leader":   leader_cnt,
        "manager":  manager_cnt,
        "assignee": assignee_cnt,
    })

@app.route("/api/users/<uid>", methods=["DELETE"])
@require_role("superadmin")
def delete_user(uid):
    if uid == g.user_id:
        return jsonify({"error": "Cannot delete your own account"}), 400
    conn = get_db(); c = conn.cursor()
    # Check for pending tasks
    c.execute("""SELECT COUNT(*) FROM tasks
                 WHERE (task_leader=%s OR task_manager=%s OR assigned_to=%s)
                 AND status NOT IN ('completed','cancelled')""", (uid,uid,uid))
    pending = c.fetchone()[0]
    if pending > 0:
        conn.close()
        return jsonify({"error": f"User has {pending} pending task(s). Reassign them first.", "pending": pending}), 409
    c.execute("DELETE FROM users WHERE id=%s", (uid,))
    conn.commit(); conn.close()
    return jsonify({"success": True})

@app.route("/api/users/<uid>/reset-password", methods=["POST"])
@require_role("superadmin")
def reset_user_password(uid):
    d = request.get_json(silent=True, force=True) or {}
    new_pw = d.get("password", "")
    if len(new_pw) < 6:
        return jsonify({"error": "Password must be at least 6 characters"}), 400
    conn = get_db(); c = conn.cursor()
    c.execute("UPDATE users SET password=%s WHERE id=%s", (hash_pw(new_pw), uid))
    conn.commit(); conn.close()
    return jsonify({"success": True})

# ══ DASHBOARD ════════════════════════════════════════════════════════════════
@app.route("/api/dashboard")
@login_required
def dashboard():
    run_compliance_checks()
    uid  = g.user_id
    role = g.role
    conn = get_db(); c = conn.cursor()
    today = date.today()

    # ── Role-scoped task filter ───────────────────────────────────────────────
    if role == "superadmin":
        task_scope = "1=1"
        task_params = []
        alert_scope = "1=1"
        alert_params = []
    elif role == "manager":
        task_scope = "t.task_manager = %s"
        task_params = [uid]
        alert_scope = "a.company_id IN (SELECT DISTINCT company_id FROM tasks WHERE task_manager=%s AND status NOT IN ('completed','cancelled'))"
        alert_params = [uid]
    else:
        task_scope = "t.assigned_to = %s"
        task_params = [uid]
        alert_scope = "a.company_id IN (SELECT DISTINCT company_id FROM tasks WHERE assigned_to=%s AND status NOT IN ('completed','cancelled'))"
        alert_params = [uid]

    # ── Stats ────────────────────────────────────────────────────────────────
    c.execute("SELECT COUNT(*) FROM companies WHERE status='active'"); total_co  = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM directors WHERE is_active=1");     total_dir = c.fetchone()[0]
    c.execute(f"SELECT COUNT(*) FROM alerts a WHERE a.status='active' AND {alert_scope}", alert_params)
    active_alerts = c.fetchone()[0]
    c.execute(f"SELECT COUNT(*) FROM alerts a WHERE a.status='active' AND a.severity='critical' AND {alert_scope}", alert_params)
    crit = c.fetchone()[0]
    c.execute(f"SELECT COUNT(*) FROM tasks t WHERE t.status NOT IN ('completed','cancelled') AND {task_scope}", task_params)
    open_tasks = c.fetchone()[0]
    c.execute(f"SELECT COUNT(*) FROM tasks t WHERE t.status NOT IN ('completed','cancelled') AND t.due_date<=%s AND {task_scope}",
              [(today+timedelta(days=7)).isoformat()] + task_params)
    due_soon = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM dsc_records WHERE is_active=1"); total_dsc = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM dsc_records WHERE is_active=1 AND valid_to<?", (today.isoformat(),)); dsc_exp  = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM dsc_records WHERE is_active=1 AND valid_to BETWEEN ? AND ?",
              (today.isoformat(), (today+timedelta(days=30)).isoformat())); dsc_soon = c.fetchone()[0]

    # ── Alert panels ─────────────────────────────────────────────────────────
    def role_alerts(entity_type, limit=5):
        c.execute(f"""SELECT a.*, co.name AS company_name FROM alerts a
                     LEFT JOIN companies co ON a.company_id=co.id
                     WHERE a.entity_type=%s AND a.status='active' AND {alert_scope}
                     ORDER BY a.due_date LIMIT {limit}""",
                  [entity_type] + alert_params)
        return rows(c.fetchall())

    aud_alerts = role_alerts("auditor")
    din_alerts = role_alerts("director")
    dsc_alerts = role_alerts("dsc")

    c.execute("""SELECT m.*,co.name as company_name FROM meetings m
                 JOIN companies co ON m.company_id=co.id
                 WHERE m.meeting_date>=%s AND m.status='scheduled'
                 ORDER BY m.meeting_date LIMIT 5""", (today.isoformat(),))
    meetings = rows(c.fetchall())

    c.execute(f"""SELECT t.*, co.name AS company_name FROM tasks t
                 LEFT JOIN companies co ON t.company_id=co.id
                 WHERE t.status NOT IN ('completed','cancelled') AND {task_scope}
                 ORDER BY CASE t.priority WHEN 'critical' THEN 1 WHEN 'high' THEN 2
                          WHEN 'medium' THEN 3 ELSE 4 END, t.due_date LIMIT 8""",
              task_params)
    tasks = rows(c.fetchall())

    conn.close()
    return jsonify({
        "stats": {
            "total_companies": total_co, "total_directors": total_dir,
            "active_alerts": active_alerts, "critical_alerts": crit,
            "open_tasks": open_tasks, "due_soon_tasks": due_soon,
            "total_dsc": total_dsc, "dsc_expired": dsc_exp, "dsc_expiring_soon": dsc_soon,
        },
        "auditor_alerts":   aud_alerts,
        "din_alerts":       din_alerts,
        "director_alerts":  din_alerts,
        "dsc_alerts":       dsc_alerts,
        "upcoming_meetings": meetings,
        "pending_tasks":    tasks,
    })

# ══ COMPANIES ════════════════════════════════════════════════════════════════
@app.route("/api/companies")
@login_required
def list_companies():
    q=request.args.get("q","").strip(); conn=get_db(); c=conn.cursor()
    tid = g.tenant_id
    if q:
        if tid: c.execute("SELECT * FROM companies WHERE tenant_id=%s AND (name LIKE ? OR cin LIKE ? OR pan LIKE ?) ORDER BY name",[tid,f"%{q}%",f"%{q}%",f"%{q}%"])
        else:   c.execute("SELECT * FROM companies WHERE name LIKE ? OR cin LIKE ? OR pan LIKE ? ORDER BY name",[f"%{q}%"]*3)
    else:
        if tid: c.execute("SELECT * FROM companies WHERE tenant_id=%s ORDER BY name",(tid,))
        else:   c.execute("SELECT * FROM companies ORDER BY name")
    return jsonify(rows(c.fetchall()))

@app.route("/api/companies/<cid>")
@login_required
def get_company(cid):
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT * FROM companies WHERE id=%s",(cid,)); co=row(c.fetchone())
    if not co: return jsonify({"error":"Not found"}),404
    c.execute("""SELECT d.*,k.last_kyc_date,k.next_due_date,k.kyc_status FROM directors d
                 LEFT JOIN director_kyc k ON d.id=k.director_id WHERE d.company_id=%s AND d.is_active=1""",(cid,))
    co["directors"]=rows(c.fetchall())
    c.execute("SELECT * FROM auditors WHERE company_id=%s ORDER BY created_at DESC",(cid,))
    co["auditors"]=[_enrich_auditor(a) for a in rows(c.fetchall())]
    c.execute("SELECT * FROM shareholders WHERE company_id=%s AND is_active=1",(cid,))
    co["shareholders"]=rows(c.fetchall())
    c.execute("SELECT * FROM dsc_records WHERE company_id=%s AND is_active=1",(cid,))
    co["dsc_records"]=rows(c.fetchall())
    c.execute("SELECT * FROM meetings WHERE company_id=%s ORDER BY meeting_date DESC LIMIT 5",(cid,))
    co["recent_meetings"]=rows(c.fetchall())
    c.execute("SELECT * FROM alerts WHERE company_id=%s AND status='active' ORDER BY severity DESC",(cid,))
    co["alerts"]=rows(c.fetchall())
    conn.close(); return jsonify(co)

@app.route("/api/companies/<cid>/master-pdf")
@login_required
def company_master_pdf(cid):
    try:
        pdf=generate_company_master_pdf(cid)
        conn=get_db(); c=conn.cursor()
        c.execute("SELECT name FROM companies WHERE id=%s",(cid,)); co=row(c.fetchone()); conn.close()
        fname=f"CompanyMaster_{(co['name'] if co else cid)[:20].replace(' ','_')}.pdf"
        return send_file(io.BytesIO(pdf),mimetype="application/pdf",as_attachment=True,download_name=fname)
    except Exception as e:
        return jsonify({"error":str(e)}),400

@app.route("/api/companies", methods=["POST"])
@login_required
def create_company():
    if not can("company","create"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}
    if not d.get("name"): return jsonify({"error":"Company name required"}),400
    cid=str(uuid.uuid4()); tid=g.tenant_id; conn=get_db(); c=conn.cursor()
    c.execute("""INSERT INTO companies
        (id,name,cin,incorporation_date,registered_office,pan,tan,email,phone,
         authorized_capital,paid_up_capital,business_activity,company_type,roc,
         letterhead_address,letterhead_footer,created_by)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
        (cid,d["name"],d.get("cin"),d.get("incorporation_date"),d.get("registered_office"),
         (d.get("pan") or "").upper() or None,(d.get("tan") or "").upper() or None,
         d.get("email"),d.get("phone"),float(d.get("authorized_capital") or 0),
         float(d.get("paid_up_capital") or 0),d.get("business_activity"),
         d.get("company_type","Private Limited"),d.get("roc"),
         d.get("letterhead_address"),d.get("letterhead_footer"),g.user_id))
    conn.commit()
    c.execute("SELECT * FROM companies WHERE id=%s",(cid,)); result=row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/companies/<cid>", methods=["PUT"])
@login_required
def update_company(cid):
    if not can("company","update"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}
    fields={k:d[k] for k in ["name","cin","incorporation_date","registered_office","pan","tan","email","phone",
             "authorized_capital","paid_up_capital","business_activity","company_type","roc","status",
             "letterhead_address","letterhead_footer"] if k in d}
    fields["updated_at"]=datetime.utcnow().isoformat()
    conn=get_db(); c=conn.cursor()
    c.execute(f"UPDATE companies SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[cid])
    conn.commit()
    c.execute("SELECT * FROM companies WHERE id=%s",(cid,)); result=row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/companies/<cid>", methods=["DELETE"])
@require_role("superadmin")
def delete_company(cid):
    conn=get_db(); c=conn.cursor()
    c.execute("DELETE FROM companies WHERE id=%s",(cid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ══ DIRECTORS ════════════════════════════════════════════════════════════════
@app.route("/api/directors")
@login_required
def all_directors():
    conn=get_db(); c=conn.cursor()
    c.execute("""SELECT d.*,k.last_kyc_date,k.next_due_date,k.kyc_status,
                        co.name as company_name,co.cin as company_cin
                 FROM directors d LEFT JOIN director_kyc k ON d.id=k.director_id
                 LEFT JOIN companies co ON d.company_id=co.id WHERE d.is_active=1 ORDER BY d.name""")
    return jsonify(rows(c.fetchall()))

@app.route("/api/companies/<cid>/directors")
@login_required
def list_directors(cid):
    conn=get_db(); c=conn.cursor()
    c.execute("""SELECT d.*,k.last_kyc_date,k.next_due_date,k.kyc_status FROM directors d
                 LEFT JOIN director_kyc k ON d.id=k.director_id WHERE d.company_id=%s ORDER BY d.name""",(cid,))
    return jsonify(rows(c.fetchall()))

@app.route("/api/directors", methods=["POST"])
@login_required
def create_director():
    if not can("director","create"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}
    if not d.get("company_id") or not d.get("name"): return jsonify({"error":"company_id and name required"}),400
    did=str(uuid.uuid4()); conn=get_db(); c=conn.cursor()
    c.execute("""INSERT INTO directors
        (id,company_id,name,din,pan,aadhaar,email,mobile,address,designation,
         date_of_appointment,mca_user_id,mca_password)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
        (did,d["company_id"],d["name"],d.get("din"),(d.get("pan") or "").upper() or None,
         d.get("aadhaar"),d.get("email"),d.get("mobile"),d.get("address"),
         d.get("designation","Director"),d.get("date_of_appointment"),
         d.get("mca_user_id"),d.get("mca_password")))
    last_kyc=d.get("last_kyc_date"); due=_kyc_due()
    c.execute("INSERT INTO director_kyc (id,director_id,last_kyc_date,next_due_date,kyc_status) VALUES (%s,%s,?,%s,%s)",
              (str(uuid.uuid4()),did,last_kyc,due,_kyc_status(due)))
    conn.commit()
    c.execute("""SELECT d.*,k.last_kyc_date,k.next_due_date,k.kyc_status FROM directors d
                 LEFT JOIN director_kyc k ON d.id=k.director_id WHERE d.id=%s""",(did,))
    result=row(c.fetchone()); conn.close(); return jsonify(result),201

@app.route("/api/directors/<did>", methods=["PUT"])
@login_required
def update_director(did):
    if not can("director","update"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}; conn=get_db(); c=conn.cursor()
    fields={k:d[k] for k in ["name","din","pan","aadhaar","email","mobile","address","designation",
             "date_of_appointment","date_of_cessation","is_active","mca_user_id","mca_password"] if k in d}
    if fields:
        c.execute(f"UPDATE directors SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[did])
    if "last_kyc_date" in d:
        due=_kyc_due(); st=_kyc_status(due)
        c.execute("UPDATE director_kyc SET last_kyc_date=%s,next_due_date=%s,kyc_status=%s,updated_at=NOW() WHERE director_id=%s",
                  (d["last_kyc_date"],due,st,did))
    conn.commit()
    c.execute("""SELECT d.*,k.last_kyc_date,k.next_due_date,k.kyc_status FROM directors d
                 LEFT JOIN director_kyc k ON d.id=k.director_id WHERE d.id=%s""",(did,))
    result=row(c.fetchone()); conn.close(); return jsonify(result)

@app.route("/api/directors/<did>", methods=["DELETE"])
@login_required
def delete_director(did):
    if not can("director","delete"): return jsonify({"error":"Insufficient permissions"}),403
    conn=get_db(); c=conn.cursor()
    c.execute("UPDATE directors SET is_active=0 WHERE id=%s",(did,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ══ AUDITORS ═════════════════════════════════════════════════════════════════
@app.route("/api/auditors")
@login_required
def all_auditors():
    conn=get_db(); c=conn.cursor()
    c.execute("""SELECT a.*,co.name as company_name,co.cin as company_cin
                 FROM auditors a JOIN companies co ON a.company_id=co.id
                 WHERE a.is_active=1 ORDER BY a.end_date""")
    return jsonify([_enrich_auditor(a) for a in rows(c.fetchall())])

@app.route("/api/companies/<cid>/auditors")
@login_required
def list_auditors(cid):
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT * FROM auditors WHERE company_id=%s ORDER BY created_at DESC",(cid,))
    return jsonify([_enrich_auditor(a) for a in rows(c.fetchall())])

@app.route("/api/auditors", methods=["POST"])
@login_required
def create_auditor():
    if not can("auditor","create"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}
    if not d.get("company_id") or not d.get("name"): return jsonify({"error":"company_id and name required"}),400
    aid=str(uuid.uuid4())
    end=d.get("end_date")
    if not end and d.get("start_date"):
        sd=date.fromisoformat(d["start_date"][:10]); end=sd.replace(year=sd.year+1).isoformat()
    conn=get_db(); c=conn.cursor()
    c.execute("UPDATE auditors SET is_active=0 WHERE company_id=%s AND is_active=1",(d["company_id"],))
    c.execute("""INSERT INTO auditors
        (id,company_id,name,firm_name,membership_no,frn,pan,address,email,phone,
         appointment_date,nature_of_appointment,appointment_type,start_date,end_date,srn_adt1,notes)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
        (aid,d["company_id"],d["name"],d.get("firm_name"),d.get("membership_no"),d.get("frn"),
         (d.get("pan") or "").upper() or None,d.get("address"),d.get("email"),d.get("phone"),
         d.get("appointment_date"),d.get("nature_of_appointment","Regular Auditor"),
         d.get("appointment_type","AGM Appointment"),d.get("start_date"),end,d.get("srn_adt1"),d.get("notes")))
    conn.commit()
    c.execute("SELECT * FROM auditors WHERE id=%s",(aid,)); result=row(c.fetchone()); conn.close()
    return jsonify(_enrich_auditor(result)),201

@app.route("/api/auditors/<aid>", methods=["PUT"])
@login_required
def update_auditor(aid):
    if not can("auditor","update"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}
    fields={k:d[k] for k in ["name","firm_name","membership_no","frn","pan","address","email","phone",
             "appointment_date","nature_of_appointment","appointment_type","start_date","end_date",
             "srn_adt1","is_active","notes"] if k in d}
    if not fields: return jsonify({"error":"Nothing to update"}),400
    conn=get_db(); c=conn.cursor()
    c.execute(f"UPDATE auditors SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[aid])
    conn.commit()
    c.execute("SELECT * FROM auditors WHERE id=%s",(aid,)); result=row(c.fetchone()); conn.close()
    return jsonify(_enrich_auditor(result))

@app.route("/api/auditors/<aid>", methods=["DELETE"])
@login_required
def delete_auditor(aid):
    if not can("auditor","delete"): return jsonify({"error":"Insufficient permissions"}),403
    conn=get_db(); c=conn.cursor()
    c.execute("UPDATE auditors SET is_active=0 WHERE id=%s",(aid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

@app.route("/api/auditors/<aid>/upload", methods=["POST"])
@login_required
def upload_adt1(aid):
    if "file" not in request.files: return jsonify({"error":"No file"}),400
    f=request.files["file"]
    if not f.filename: return jsonify({"error":"No file selected"}),400
    ext=Path(f.filename).suffix.lower()
    if ext not in ALLOWED_EXT: return jsonify({"error":f"Type {ext} not allowed"}),400
    fname=f"adt1_{aid}{ext}"; f.save(str(UPLOAD_DIR/fname))
    conn=get_db(); c=conn.cursor()
    c.execute("UPDATE auditors SET adt1_file=%s WHERE id=%s",(fname,aid)); conn.commit(); conn.close()
    return jsonify({"success":True,"file":fname})

# ══ SHAREHOLDERS ═════════════════════════════════════════════════════════════
@app.route("/api/companies/<cid>/shareholders")
@login_required
def list_shareholders(cid):
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT * FROM shareholders WHERE company_id=%s AND is_active=1 ORDER BY name",(cid,))
    return jsonify(rows(c.fetchall()))

@app.route("/api/shareholders", methods=["POST"])
@login_required
def create_shareholder():
    if not can("shareholder","create"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}; sid=str(uuid.uuid4()); conn=get_db(); c=conn.cursor()
    c.execute("""INSERT INTO shareholders
        (id,company_id,name,folio_no,pan,email,mobile,address,share_class,shares_held,face_value,
         date_of_entry,mca_user_id,mca_password)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s)""",
        (sid,d["company_id"],d["name"],d.get("folio_no"),(d.get("pan") or "").upper() or None,
         d.get("email"),d.get("mobile"),d.get("address"),d.get("share_class","Equity"),
         int(d.get("shares_held",0)),float(d.get("face_value",10)),d.get("date_of_entry"),
         d.get("mca_user_id"),d.get("mca_password")))
    conn.commit()
    c.execute("SELECT * FROM shareholders WHERE id=%s",(sid,)); result=row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/shareholders/<sid>", methods=["PUT"])
@login_required
def update_shareholder(sid):
    if not can("shareholder","update"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}
    fields={k:d[k] for k in ["name","folio_no","pan","email","mobile","address","share_class",
             "shares_held","face_value","is_active","mca_user_id","mca_password"] if k in d}
    conn=get_db(); c=conn.cursor()
    c.execute(f"UPDATE shareholders SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[sid])
    conn.commit(); c.execute("SELECT * FROM shareholders WHERE id=%s",(sid,)); result=row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/shareholders/<sid>", methods=["DELETE"])
@login_required
def delete_shareholder(sid):
    conn=get_db(); c=conn.cursor()
    c.execute("UPDATE shareholders SET is_active=0 WHERE id=%s",(sid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ══ DSC RECORDS ══════════════════════════════════════════════════════════════
@app.route("/api/dsc")
@login_required
def all_dsc():
    cid=request.args.get("company_id",""); conn=get_db(); c=conn.cursor()
    if cid:
        c.execute("""SELECT d.*,co.name as company_name FROM dsc_records d
                     LEFT JOIN companies co ON d.company_id=co.id
                     WHERE d.company_id=%s AND d.is_active=1 ORDER BY d.valid_to""",(cid,))
    else:
        c.execute("""SELECT d.*,co.name as company_name FROM dsc_records d
                     LEFT JOIN companies co ON d.company_id=co.id
                     WHERE d.is_active=1 ORDER BY d.valid_to""")
    result=rows(c.fetchall())
    today=date.today()
    for r in result:
        if r.get("valid_to"):
            days=(date.fromisoformat(r["valid_to"][:10])-today).days
            r["days_to_expiry"]=days
            r["expiry_status"]="expired" if days<0 else ("expiring_soon" if days<=30 else "valid")
    conn.close(); return jsonify(result)

@app.route("/api/dsc", methods=["POST"])
@login_required
def create_dsc():
    if not can("dsc","create"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}; dsc_id=str(uuid.uuid4()); conn=get_db(); c=conn.cursor()
    c.execute("""INSERT INTO dsc_records
        (id,company_id,director_id,holder_name,holder_type,dsc_class,issued_by,
         valid_from,valid_to,token_type,custody_status,custody_date,custody_notes,notes)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s)""",
        (dsc_id,d.get("company_id"),d.get("director_id"),d["holder_name"],
         d.get("holder_type","Director"),d.get("dsc_class","Class 3"),d.get("issued_by"),
         d.get("valid_from"),d.get("valid_to"),d.get("token_type"),
         d.get("custody_status","With Client"),d.get("custody_date"),
         d.get("custody_notes"),d.get("notes")))
    conn.commit()
    c.execute("SELECT * FROM dsc_records WHERE id=%s",(dsc_id,)); result=row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/dsc/<dsc_id>", methods=["PUT"])
@login_required
def update_dsc(dsc_id):
    if not can("dsc","update"): return jsonify({"error":"Insufficient permissions"}),403
    d=request.get_json(silent=True, force=True) or {}
    fields={k:d[k] for k in ["holder_name","holder_type","dsc_class","issued_by","valid_from","valid_to",
             "token_type","custody_status","custody_date","custody_notes","is_active","notes"] if k in d}
    conn=get_db(); c=conn.cursor()
    if fields:
        c.execute(f"UPDATE dsc_records SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[dsc_id])
    # Log custody change
    if "custody_status" in d:
        c.execute("SELECT custody_status FROM dsc_records WHERE id=%s",(dsc_id,))
        old=c.fetchone()
        c.execute("""INSERT INTO dsc_custody_log (id,dsc_id,action,action_date,from_party,to_party,notes,recorded_by)
                     VALUES (%s,%s,?,%s,?,%s,?,%s)""",
                  (str(uuid.uuid4()),dsc_id,"custody_change",date.today().isoformat(),
                   old[0] if old else "",d["custody_status"],d.get("custody_notes",""),g.user_id))
    conn.commit()
    c.execute("SELECT * FROM dsc_records WHERE id=%s",(dsc_id,)); result=row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/dsc/<dsc_id>", methods=["DELETE"])
@login_required
def delete_dsc(dsc_id):
    if not can("dsc","delete"): return jsonify({"error":"Insufficient permissions"}),403
    conn=get_db(); c=conn.cursor()
    c.execute("UPDATE dsc_records SET is_active=0 WHERE id=%s",(dsc_id,)); conn.commit(); conn.close()
    return jsonify({"success":True})

@app.route("/api/dsc/<dsc_id>/custody-log")
@login_required
def dsc_custody_log(dsc_id):
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT l.*,u.name as recorded_by_name FROM dsc_custody_log l LEFT JOIN users u ON l.recorded_by=u.id WHERE l.dsc_id=%s ORDER BY l.created_at DESC",(dsc_id,))
    return jsonify(rows(c.fetchall()))

# ══ MEETINGS ═════════════════════════════════════════════════════════════════
@app.route("/api/meetings")
@login_required
def list_meetings():
    cid=request.args.get("company_id",""); conn=get_db(); c=conn.cursor()
    q="SELECT m.*,co.name as company_name FROM meetings m JOIN companies co ON m.company_id=co.id"
    params=[]
    if cid: q+=" WHERE m.company_id=%s"; params.append(cid)
    q+=" ORDER BY m.meeting_date DESC"
    c.execute(q,params); return jsonify(rows(c.fetchall()))

@app.route("/api/meetings", methods=["POST"])
@login_required
def create_meeting():
    d=request.get_json(silent=True, force=True) or {}; mid=str(uuid.uuid4()); conn=get_db(); c=conn.cursor()
    c.execute("""INSERT INTO meetings (id,company_id,meeting_type,meeting_no,meeting_date,meeting_time,venue,agenda,status,created_by)
                 VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s)""",
        (mid,d["company_id"],d.get("meeting_type","Board"),d.get("meeting_no"),
         d["meeting_date"],d.get("meeting_time",""),d.get("venue",""),d.get("agenda",""),
         d.get("status","scheduled"),g.user_id))
    conn.commit()
    c.execute("SELECT m.*,co.name as company_name FROM meetings m JOIN companies co ON m.company_id=co.id WHERE m.id=%s",(mid,))
    result=row(c.fetchone()); conn.close(); return jsonify(result),201

@app.route("/api/meetings/<mid>", methods=["PUT"])
@login_required
def update_meeting(mid):
    d=request.get_json(silent=True, force=True) or {}
    fields={k:d[k] for k in ["meeting_type","meeting_no","meeting_date","meeting_time","venue","agenda","notes","minutes_drafted","status"] if k in d}
    conn=get_db(); c=conn.cursor()
    c.execute(f"UPDATE meetings SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[mid])
    conn.commit()
    c.execute("SELECT m.*,co.name as company_name FROM meetings m JOIN companies co ON m.company_id=co.id WHERE m.id=%s",(mid,))
    result=row(c.fetchone()); conn.close(); return jsonify(result)

@app.route("/api/meetings/<mid>", methods=["DELETE"])
@login_required
def delete_meeting(mid):
    conn=get_db(); c=conn.cursor(); c.execute("DELETE FROM meetings WHERE id=%s",(mid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ══ STATUTORY REGISTERS ══════════════════════════════════════════════════════
@app.route("/api/registers/types")
@login_required
def register_types():
    return jsonify([{"key":k,"name":v["name"],"section":v["section"]} for k,v in REGISTER_DEFINITIONS.items()])

@app.route("/api/registers/<reg_type>/<cid>")
@login_required
def get_register(reg_type, cid):
    reg=REGISTER_DEFINITIONS.get(reg_type)
    if not reg: return jsonify({"error":f"Unknown register: {reg_type}"}),400
    conn=get_db(); c=conn.cursor()
    # Run the display query (column-based)
    c.execute(reg["query"],(cid,))
    rows_raw=c.fetchall()
    data=[list(r) for r in rows_raw]
    # Also fetch IDs for deletable registers
    id_queries = {
        "MBP-1": "SELECT id FROM director_interests WHERE company_id=%s ORDER BY date_of_disclosure",
        "CHG-1": "SELECT id FROM charges WHERE company_id=%s ORDER BY date_of_creation",
        "MBP-3": "SELECT id FROM investments WHERE company_id=%s ORDER BY date_of_investment",
        "MBP-2": "SELECT id FROM loans_guarantees WHERE company_id=%s ORDER BY date_of_transaction",
        "RPT-188": "SELECT id FROM related_party_transactions WHERE company_id=%s ORDER BY date_of_transaction",
        "SH-6": "SELECT id FROM esop_grants WHERE company_id=%s ORDER BY grant_date",
    }
    ids=[]
    if reg_type in id_queries:
        c.execute(id_queries[reg_type],(cid,))
        ids=[r[0] for r in c.fetchall()]
    c.execute("SELECT name,cin FROM companies WHERE id=%s",(cid,))
    co=row(c.fetchone()); conn.close()
    return jsonify({"register":reg["name"],"section":reg["section"],"company":co,
                    "columns":reg["columns"],"data":data,"_ids":ids})

@app.route("/api/registers/<reg_type>/<cid>/pdf")
@login_required
def download_register_pdf(reg_type, cid):
    try:
        pdf=generate_register_pdf(cid,reg_type)
        reg=REGISTER_DEFINITIONS.get(reg_type,{})
        fname=f"{reg.get('name','Register').replace(' ','_')}_{cid[:8]}.pdf"
        return send_file(io.BytesIO(pdf),mimetype="application/pdf",as_attachment=True,download_name=fname)
    except Exception as e:
        return jsonify({"error":str(e)}),400

# ══ DOCUMENT TEMPLATES ═══════════════════════════════════════════════════════
@app.route("/api/document-templates")
@login_required
def list_doc_templates():
    cat=request.args.get("category",""); conn=get_db(); c=conn.cursor()
    q="SELECT id,name,category,description,placeholders,is_system,is_active,created_at FROM document_templates WHERE is_active=1"
    params=[]
    if cat: q+=" AND category=%s"; params.append(cat)
    q+=" ORDER BY is_system DESC,name"
    c.execute(q,params); return jsonify(rows(c.fetchall()))

@app.route("/api/document-templates/<tid>")
@login_required
def get_doc_template(tid):
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT * FROM document_templates WHERE id=%s",(tid,))
    t=row(c.fetchone()); conn.close()
    if not t: return jsonify({"error":"Not found"}),404
    return jsonify(t)

@app.route("/api/document-templates", methods=["POST"])
@login_required
def create_doc_template():
    d=request.get_json(silent=True, force=True) or {}
    if not d.get("name") or not d.get("template_body"): return jsonify({"error":"name and template_body required"}),400
    import re
    placeholders=list(set(re.findall(r'\{\{([^}]+)\}\}',d["template_body"])))
    tid=str(uuid.uuid4()); conn=get_db(); c=conn.cursor()
    c.execute("""INSERT INTO document_templates (id,name,category,description,template_body,placeholders,is_system,created_by)
                 VALUES (%s,%s,?,%s,?,%s,0,%s)""",
        (tid,d["name"],d.get("category","resolution"),d.get("description",""),
         d["template_body"],json.dumps(placeholders),g.user_id))
    conn.commit()
    c.execute("SELECT * FROM document_templates WHERE id=%s",(tid,)); result=row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/document-templates/<tid>", methods=["PUT"])
@login_required
def update_doc_template(tid):
    d=request.get_json(silent=True, force=True) or {}; conn=get_db(); c=conn.cursor()
    c.execute("SELECT is_system FROM document_templates WHERE id=%s",(tid,))
    t=c.fetchone()
    if t and t[0] and g.role!="superadmin": conn.close(); return jsonify({"error":"System templates can only be edited by superadmin"}),403
    import re
    fields={k:d[k] for k in ["name","category","description","template_body","is_active"] if k in d}
    if "template_body" in d: fields["placeholders"]=json.dumps(list(set(re.findall(r'\{\{([^}]+)\}\}',d["template_body"]))))
    fields["updated_at"]=datetime.utcnow().isoformat()
    c.execute(f"UPDATE document_templates SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[tid])
    conn.commit()
    c.execute("SELECT * FROM document_templates WHERE id=%s",(tid,)); result=row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/document-templates/<tid>", methods=["DELETE"])
@require_role("superadmin","manager")
def delete_doc_template(tid):
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT is_system FROM document_templates WHERE id=%s",(tid,))
    t=c.fetchone()
    if t and t[0]: conn.close(); return jsonify({"error":"Cannot delete system templates"}),400
    c.execute("UPDATE document_templates SET is_active=0 WHERE id=%s",(tid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ── Document Generation from template ────────────────────────────────────────
@app.route("/api/companies/<cid>/active-entities")
@login_required
def active_entities(cid):
    """Return only active directors, current auditors, active shareholders for document generation."""
    try:
        data = get_active_entities(cid)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 400

@app.route("/api/templates/placeholders", methods=["POST"])
@login_required
def get_template_placeholders():
    """Parse placeholders from a template and classify auto-filled vs manual."""
    d = request.get_json(silent=True, force=True) or {}
    tid = d.get("template_id")
    if not tid: return jsonify({"error": "template_id required"}), 400
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT template_body, placeholders FROM document_templates WHERE id=%s", (tid,))
    tmpl = row(c.fetchone()); conn.close()
    if not tmpl: return jsonify({"error": "Template not found"}), 404
    all_phs = extract_placeholders(tmpl["template_body"])
    auto = [p for p in all_phs if p in AUTO_FILLED]
    manual = [p for p in all_phs if p not in AUTO_FILLED]
    return jsonify({"all": all_phs, "auto_filled": auto, "manual": manual})

@app.route("/api/documents/generate", methods=["POST"])
@login_required
def gen_document():
    d = request.get_json(silent=True, force=True) or {}
    template_id = d.get("template_id"); company_id = d.get("company_id")
    if not template_id or not company_id:
        return jsonify({"error": "template_id and company_id required"}), 400
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT * FROM document_templates WHERE id=%s", (template_id,))
    tmpl = row(c.fetchone()); conn.close()
    if not tmpl: return jsonify({"error": "Template not found"}), 404
    try:
        ctx = build_context(
            company_id,
            extra=d.get("extra_context", {}),
            director_id=d.get("director_id"),
            director2_id=d.get("director2_id"),
            auditor_id=d.get("auditor_id"),
        )
        content = generate_document(tmpl["template_body"], ctx)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    all_phs = extract_placeholders(tmpl["template_body"])
    doc_id = str(uuid.uuid4()); conn = get_db(); c = conn.cursor()
    c.execute("INSERT INTO documents (id,company_id,template_id,doc_type,doc_name,content,module,generated_by) VALUES (%s,%s,?,%s,?,%s,?,%s)",
              (doc_id, company_id, template_id, tmpl["category"], tmpl["name"], content, "document_engine", g.user_id))
    conn.commit(); conn.close()
    return jsonify({"id": doc_id, "template_name": tmpl["name"], "category": tmpl["category"],
                    "content": content, "context": ctx,
                    "placeholders": all_phs,
                    "auto_filled": [p for p in all_phs if p in AUTO_FILLED],
                    "manual": [p for p in all_phs if p not in AUTO_FILLED]})

@app.route("/api/documents/generate/pdf", methods=["POST"])
@login_required
def gen_document_pdf():
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
    from reportlab.lib.units import inch
    from reportlab.lib import colors

    d=request.get_json(silent=True, force=True) or {}
    template_id=d.get("template_id"); company_id=d.get("company_id")
    if not template_id or not company_id: return jsonify({"error":"template_id and company_id required"}),400
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT * FROM document_templates WHERE id=%s",(template_id,))
    tmpl=row(c.fetchone()); conn.close()
    if not tmpl: return jsonify({"error":"Template not found"}),404
    ctx = build_context(
        company_id,
        extra=d.get("extra_context", {}),
        director_id=d.get("director_id"),
        director2_id=d.get("director2_id"),
        auditor_id=d.get("auditor_id"),
    )
    content=generate_document(tmpl["template_body"],ctx)

    # Get company for letterhead
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT * FROM companies WHERE id=%s",(company_id,)); co=row(c.fetchone()); conn.close()

    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,topMargin=0.8*inch,bottomMargin=0.8*inch,leftMargin=1.0*inch,rightMargin=1.0*inch)
    NAVY=colors.HexColor("#0f2d5c"); BLUE=colors.HexColor("#1a56db"); GREY=colors.HexColor("#64748b")

    def sty(n,sz=9.5,bold=False,col=None,align=0):
        return ParagraphStyle(n,fontName="Helvetica-Bold" if bold else "Helvetica",
                               fontSize=sz,textColor=col or colors.HexColor("#1e293b"),
                               alignment=align,leading=sz*1.5)
    story=[]
    # Letterhead
    if co:
        story.append(Paragraph(co["name"].upper(),sty("ln",15,True,NAVY,1)))
        lh=co.get("letterhead_address") or co.get("registered_office","")
        if lh: story.append(Paragraph(lh,sty("la",8,False,GREY,1)))
        story.append(Spacer(1,4))
        story.append(HRFlowable(width="100%",thickness=2,color=BLUE,spaceAfter=2))
        story.append(HRFlowable(width="100%",thickness=0.5,color=NAVY,spaceAfter=10))

    for line in content.split("\n"):
        stripped=line.strip()
        if not stripped: story.append(Spacer(1,5))
        elif stripped.isupper() and len(stripped)<80: story.append(Paragraph(stripped,sty("h",11,True,NAVY)))
        else: story.append(Paragraph(line.replace("&","&amp;").replace("<","&lt;"),sty("b")))

    if co:
        story.append(Spacer(1,16))
        story.append(HRFlowable(width="100%",thickness=0.5,color=GREY))
        lf=co.get("letterhead_footer") or f"CIN: {co.get('cin','')} | PAN: {co.get('pan','')} | {co.get('registered_office','')}"
        story.append(Paragraph(lf,sty("ft",7,False,GREY,1)))

    doc.build(story)
    buf.seek(0)
    fname=f"{tmpl['name'].replace(' ','_')[:30]}.pdf"
    return send_file(buf,mimetype="application/pdf",as_attachment=True,download_name=fname)

@app.route("/api/documents/<company_id>")
@login_required
def list_documents(company_id):
    conn=get_db(); c=conn.cursor()
    c.execute("""SELECT d.*,t.name as template_name FROM documents d
                 LEFT JOIN document_templates t ON d.template_id=t.id
                 WHERE d.company_id=%s ORDER BY d.created_at DESC""",(company_id,))
    return jsonify(rows(c.fetchall()))


# ══ REGISTER ENTRY CRUD ═══════════════════════════════════════════════════════
# Each register that needs data entry has full GET/POST/PUT/DELETE routes
# ─────────────────────────────────────────────────────────────────────────────

# ── Charges (CHG-1, Section 85) ──────────────────────────────────────────────
@app.route("/api/charges")
@login_required
def list_charges():
    cid = request.args.get("company_id",""); conn = get_db(); c = conn.cursor()
    if cid:
        c.execute("""SELECT ch.*,co.name as company_name FROM charges ch
                     JOIN companies co ON ch.company_id=co.id WHERE ch.company_id=%s ORDER BY ch.date_of_creation DESC""",(cid,))
    else:
        c.execute("""SELECT ch.*,co.name as company_name FROM charges ch
                     JOIN companies co ON ch.company_id=co.id ORDER BY ch.date_of_creation DESC""")
    return jsonify(rows(c.fetchall()))

@app.route("/api/charges", methods=["POST"])
@login_required
def create_charge():
    if not can("company","create"): return jsonify({"error":"Insufficient permissions"}),403
    d = request.get_json(silent=True, force=True) or {}
    if not d.get("company_id") or not d.get("charge_holder"): return jsonify({"error":"company_id and charge_holder required"}),400
    rid = str(uuid.uuid4()); conn = get_db(); c = conn.cursor()
    c.execute("""INSERT INTO charges
        (id,company_id,charge_id,srn,date_of_creation,amount,charge_holder,assets_charged,
         charge_type,date_of_modification,date_of_satisfaction,status,remarks,created_by)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s)""",
        (rid,d["company_id"],d.get("charge_id"),d.get("srn"),d.get("date_of_creation"),
         float(d.get("amount") or 0),d["charge_holder"],d.get("assets_charged"),
         d.get("charge_type","Hypothecation"),d.get("date_of_modification"),
         d.get("date_of_satisfaction"),d.get("status","Open"),d.get("remarks"),g.user_id))
    conn.commit()
    c.execute("SELECT * FROM charges WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/charges/<rid>", methods=["PUT"])
@login_required
def update_charge(rid):
    if not can("company","update"): return jsonify({"error":"Insufficient permissions"}),403
    d = request.get_json(silent=True, force=True) or {}
    fields = {k:d[k] for k in ["charge_id","srn","date_of_creation","amount","charge_holder",
              "assets_charged","charge_type","date_of_modification","date_of_satisfaction",
              "status","remarks"] if k in d}
    conn = get_db(); c = conn.cursor()
    c.execute(f"UPDATE charges SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[rid])
    conn.commit()
    c.execute("SELECT * FROM charges WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/charges/<rid>", methods=["DELETE"])
@require_role("superadmin","manager")
def delete_charge(rid):
    conn = get_db(); c = conn.cursor()
    c.execute("DELETE FROM charges WHERE id=%s",(rid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ── Director Interests (MBP-1, Section 189) ──────────────────────────────────
@app.route("/api/director-interests")
@login_required
def list_director_interests():
    cid = request.args.get("company_id",""); conn = get_db(); c = conn.cursor()
    if cid:
        c.execute("SELECT * FROM director_interests WHERE company_id=%s ORDER BY date_of_disclosure DESC",(cid,))
    else:
        c.execute("""SELECT di.*,co.name as company_name FROM director_interests di
                     JOIN companies co ON di.company_id=co.id ORDER BY di.date_of_disclosure DESC""")
    return jsonify(rows(c.fetchall()))

@app.route("/api/director-interests", methods=["POST"])
@login_required
def create_director_interest():
    d = request.get_json(silent=True, force=True) or {}
    if not d.get("company_id") or not d.get("director_name"): return jsonify({"error":"company_id and director_name required"}),400
    rid = str(uuid.uuid4()); conn = get_db(); c = conn.cursor()
    c.execute("""INSERT INTO director_interests
        (id,company_id,director_id,director_name,din,entity_name,entity_type,
         nature_of_interest,date_of_disclosure,date_of_board_resolution,remarks,created_by)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s)""",
        (rid,d["company_id"],d.get("director_id"),d["director_name"],d.get("din"),
         d.get("entity_name",""),d.get("entity_type","Company"),d.get("nature_of_interest"),
         d.get("date_of_disclosure"),d.get("date_of_board_resolution"),d.get("remarks"),g.user_id))
    conn.commit()
    c.execute("SELECT * FROM director_interests WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/director-interests/<rid>", methods=["PUT"])
@login_required
def update_director_interest(rid):
    d = request.get_json(silent=True, force=True) or {}
    fields = {k:d[k] for k in ["director_name","din","entity_name","entity_type",
              "nature_of_interest","date_of_disclosure","date_of_board_resolution","remarks"] if k in d}
    conn = get_db(); c = conn.cursor()
    c.execute(f"UPDATE director_interests SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[rid])
    conn.commit()
    c.execute("SELECT * FROM director_interests WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/director-interests/<rid>", methods=["DELETE"])
@require_role("superadmin","manager")
def delete_director_interest(rid):
    conn = get_db(); c = conn.cursor()
    c.execute("DELETE FROM director_interests WHERE id=%s",(rid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ── ESOP Grants (SH-6, Section 62) ───────────────────────────────────────────
@app.route("/api/esop-grants")
@login_required
def list_esop_grants():
    cid = request.args.get("company_id",""); conn = get_db(); c = conn.cursor()
    if cid:
        c.execute("SELECT * FROM esop_grants WHERE company_id=%s ORDER BY grant_date DESC",(cid,))
    else:
        c.execute("""SELECT eg.*,co.name as company_name FROM esop_grants eg
                     JOIN companies co ON eg.company_id=co.id ORDER BY eg.grant_date DESC""")
    return jsonify(rows(c.fetchall()))

@app.route("/api/esop-grants", methods=["POST"])
@login_required
def create_esop_grant():
    d = request.get_json(silent=True, force=True) or {}
    if not d.get("company_id") or not d.get("employee_name"): return jsonify({"error":"company_id and employee_name required"}),400
    rid = str(uuid.uuid4()); conn = get_db(); c = conn.cursor()
    c.execute("""INSERT INTO esop_grants
        (id,company_id,employee_name,designation,employee_id,grant_date,options_granted,
         exercise_price,vesting_date,vesting_period,options_exercised,options_lapsed,status,remarks,created_by)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
        (rid,d["company_id"],d["employee_name"],d.get("designation"),d.get("employee_id"),
         d.get("grant_date"),int(d.get("options_granted") or 0),float(d.get("exercise_price") or 0),
         d.get("vesting_date"),d.get("vesting_period"),int(d.get("options_exercised") or 0),
         int(d.get("options_lapsed") or 0),d.get("status","Active"),d.get("remarks"),g.user_id))
    conn.commit()
    c.execute("SELECT * FROM esop_grants WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/esop-grants/<rid>", methods=["PUT"])
@login_required
def update_esop_grant(rid):
    d = request.get_json(silent=True, force=True) or {}
    fields = {k:d[k] for k in ["employee_name","designation","employee_id","grant_date",
              "options_granted","exercise_price","vesting_date","vesting_period",
              "options_exercised","options_lapsed","status","remarks"] if k in d}
    conn = get_db(); c = conn.cursor()
    c.execute(f"UPDATE esop_grants SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[rid])
    conn.commit()
    c.execute("SELECT * FROM esop_grants WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/esop-grants/<rid>", methods=["DELETE"])
@require_role("superadmin","manager")
def delete_esop_grant(rid):
    conn = get_db(); c = conn.cursor()
    c.execute("DELETE FROM esop_grants WHERE id=%s",(rid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ── Investments (MBP-3, Section 186) ─────────────────────────────────────────
@app.route("/api/investments")
@login_required
def list_investments():
    cid = request.args.get("company_id",""); conn = get_db(); c = conn.cursor()
    if cid:
        c.execute("SELECT * FROM investments WHERE company_id=%s ORDER BY date_of_investment DESC",(cid,))
    else:
        c.execute("""SELECT i.*,co.name as company_name FROM investments i
                     JOIN companies co ON i.company_id=co.id ORDER BY i.date_of_investment DESC""")
    return jsonify(rows(c.fetchall()))

@app.route("/api/investments", methods=["POST"])
@login_required
def create_investment():
    d = request.get_json(silent=True, force=True) or {}
    if not d.get("company_id") or not d.get("investee_name"): return jsonify({"error":"company_id and investee_name required"}),400
    rid = str(uuid.uuid4()); conn = get_db(); c = conn.cursor()
    c.execute("""INSERT INTO investments
        (id,company_id,investee_name,investee_type,investment_type,amount,date_of_investment,
         board_resolution_date,srn_mgb4,purpose,remarks,created_by)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s)""",
        (rid,d["company_id"],d["investee_name"],d.get("investee_type","Company"),
         d.get("investment_type","Equity Shares"),float(d.get("amount") or 0),
         d.get("date_of_investment"),d.get("board_resolution_date"),d.get("srn_mgb4"),
         d.get("purpose"),d.get("remarks"),g.user_id))
    conn.commit()
    c.execute("SELECT * FROM investments WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/investments/<rid>", methods=["PUT"])
@login_required
def update_investment(rid):
    d = request.get_json(silent=True, force=True) or {}
    fields = {k:d[k] for k in ["investee_name","investee_type","investment_type","amount",
              "date_of_investment","board_resolution_date","srn_mgb4","purpose","remarks"] if k in d}
    conn = get_db(); c = conn.cursor()
    c.execute(f"UPDATE investments SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[rid])
    conn.commit()
    c.execute("SELECT * FROM investments WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/investments/<rid>", methods=["DELETE"])
@require_role("superadmin","manager")
def delete_investment(rid):
    conn = get_db(); c = conn.cursor()
    c.execute("DELETE FROM investments WHERE id=%s",(rid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ── Loans & Guarantees (MBP-2, Section 186) ───────────────────────────────────
@app.route("/api/loans-guarantees")
@login_required
def list_loans():
    cid = request.args.get("company_id",""); conn = get_db(); c = conn.cursor()
    if cid:
        c.execute("SELECT * FROM loans_guarantees WHERE company_id=%s ORDER BY date_of_transaction DESC",(cid,))
    else:
        c.execute("""SELECT lg.*,co.name as company_name FROM loans_guarantees lg
                     JOIN companies co ON lg.company_id=co.id ORDER BY lg.date_of_transaction DESC""")
    return jsonify(rows(c.fetchall()))

@app.route("/api/loans-guarantees", methods=["POST"])
@login_required
def create_loan():
    d = request.get_json(silent=True, force=True) or {}
    if not d.get("company_id") or not d.get("party_name"): return jsonify({"error":"company_id and party_name required"}),400
    rid = str(uuid.uuid4()); conn = get_db(); c = conn.cursor()
    c.execute("""INSERT INTO loans_guarantees
        (id,company_id,party_name,party_type,transaction_type,amount,date_of_transaction,
         rate_of_interest,repayment_date,security,board_resolution_date,outstanding_amount,status,remarks,created_by)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
        (rid,d["company_id"],d["party_name"],d.get("party_type","Company"),
         d.get("transaction_type","Loan"),float(d.get("amount") or 0),d.get("date_of_transaction"),
         float(d.get("rate_of_interest") or 0),d.get("repayment_date"),d.get("security"),
         d.get("board_resolution_date"),float(d.get("outstanding_amount") or 0),
         d.get("status","Active"),d.get("remarks"),g.user_id))
    conn.commit()
    c.execute("SELECT * FROM loans_guarantees WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/loans-guarantees/<rid>", methods=["PUT"])
@login_required
def update_loan(rid):
    d = request.get_json(silent=True, force=True) or {}
    fields = {k:d[k] for k in ["party_name","party_type","transaction_type","amount",
              "date_of_transaction","rate_of_interest","repayment_date","security",
              "board_resolution_date","outstanding_amount","status","remarks"] if k in d}
    conn = get_db(); c = conn.cursor()
    c.execute(f"UPDATE loans_guarantees SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[rid])
    conn.commit()
    c.execute("SELECT * FROM loans_guarantees WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/loans-guarantees/<rid>", methods=["DELETE"])
@require_role("superadmin","manager")
def delete_loan(rid):
    conn = get_db(); c = conn.cursor()
    c.execute("DELETE FROM loans_guarantees WHERE id=%s",(rid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ── Related Party Transactions (Section 188) ──────────────────────────────────
@app.route("/api/related-party-transactions")
@login_required
def list_rpt():
    cid = request.args.get("company_id",""); conn = get_db(); c = conn.cursor()
    if cid:
        c.execute("SELECT * FROM related_party_transactions WHERE company_id=%s ORDER BY date_of_transaction DESC",(cid,))
    else:
        c.execute("""SELECT r.*,co.name as company_name FROM related_party_transactions r
                     JOIN companies co ON r.company_id=co.id ORDER BY r.date_of_transaction DESC""")
    return jsonify(rows(c.fetchall()))

@app.route("/api/related-party-transactions", methods=["POST"])
@login_required
def create_rpt():
    d = request.get_json(silent=True, force=True) or {}
    if not d.get("company_id") or not d.get("party_name"): return jsonify({"error":"company_id and party_name required"}),400
    rid = str(uuid.uuid4()); conn = get_db(); c = conn.cursor()
    c.execute("""INSERT INTO related_party_transactions
        (id,company_id,party_name,relationship,nature_of_transaction,amount,date_of_transaction,
         date_of_board_approval,date_of_shareholders_approval,terms,justification,remarks,created_by)
        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
        (rid,d["company_id"],d["party_name"],d.get("relationship"),d.get("nature_of_transaction"),
         float(d.get("amount") or 0),d.get("date_of_transaction"),d.get("date_of_board_approval"),
         d.get("date_of_shareholders_approval"),d.get("terms"),d.get("justification"),
         d.get("remarks"),g.user_id))
    conn.commit()
    c.execute("SELECT * FROM related_party_transactions WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result),201

@app.route("/api/related-party-transactions/<rid>", methods=["PUT"])
@login_required
def update_rpt(rid):
    d = request.get_json(silent=True, force=True) or {}
    fields = {k:d[k] for k in ["party_name","relationship","nature_of_transaction","amount",
              "date_of_transaction","date_of_board_approval","date_of_shareholders_approval",
              "terms","justification","remarks"] if k in d}
    conn = get_db(); c = conn.cursor()
    c.execute(f"UPDATE related_party_transactions SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[rid])
    conn.commit()
    c.execute("SELECT * FROM related_party_transactions WHERE id=%s",(rid,)); result = row(c.fetchone()); conn.close()
    return jsonify(result)

@app.route("/api/related-party-transactions/<rid>", methods=["DELETE"])
@require_role("superadmin","manager")
def delete_rpt(rid):
    conn = get_db(); c = conn.cursor()
    c.execute("DELETE FROM related_party_transactions WHERE id=%s",(rid,)); conn.commit(); conn.close()
    return jsonify({"success":True})

# ══ ALERTS ═══════════════════════════════════════════════════════════════════
@app.route("/api/alerts")
@login_required
def list_alerts():
    uid  = g.user_id
    role = g.role
    cid  = request.args.get("company_id", "")
    sev  = request.args.get("severity",   "")
    atype= request.args.get("type",       "")
    conn = get_db(); c = conn.cursor()

    base = """SELECT a.*, co.name AS company_name
              FROM alerts a
              LEFT JOIN companies co ON a.company_id = co.id
              WHERE a.status = 'active'"""
    params = []

    if role == "superadmin":
        # ── Super Admin: all alerts ──────────────────────────────────────────
        pass

    elif role == "manager":
        # ── Manager: alerts for companies/entities where they manage tasks ───
        # Include: alerts on companies where they are task_manager on ANY task
        # Plus: alerts for entity_type=director/auditor/dsc where related task has task_manager=me
        base += """ AND (
            a.company_id IN (
                SELECT DISTINCT company_id FROM tasks WHERE task_manager = %s AND status NOT IN ('completed','cancelled')
            )
            OR EXISTS (
                SELECT 1 FROM tasks t WHERE t.task_manager = %s
                  AND t.status NOT IN ('completed','cancelled')
                  AND (
                    (a.entity_type = 'director'  AND t.module = 'director')  OR
                    (a.entity_type = 'auditor'   AND t.module = 'auditor')   OR
                    (a.entity_type = 'dsc'       AND t.module = 'dsc')       OR
                    (a.entity_type = 'filing'    AND t.module = 'filing')    OR
                    (a.entity_type = 'meeting'   AND t.module = 'meeting')   OR
                    a.entity_type  = 'company'
                  )
            )
        )"""
        params.extend([uid, uid])

    else:
        # ── Staff: only alerts for entities/companies where they are assigned ─
        base += """ AND (
            a.company_id IN (
                SELECT DISTINCT company_id FROM tasks WHERE assigned_to = %s AND status NOT IN ('completed','cancelled')
            )
            AND (
                EXISTS (
                    SELECT 1 FROM tasks t WHERE t.assigned_to = %s
                      AND t.company_id = a.company_id
                      AND t.status NOT IN ('completed','cancelled')
                      AND (
                        (a.entity_type = 'director'  AND t.module = 'director')  OR
                        (a.entity_type = 'auditor'   AND t.module = 'auditor')   OR
                        (a.entity_type = 'dsc'       AND t.module = 'dsc')       OR
                        (a.entity_type = 'filing'    AND t.module = 'filing')    OR
                        (a.entity_type = 'meeting'   AND t.module = 'meeting')   OR
                        a.entity_type  = 'company'
                      )
                )
            )
        )"""
        params.extend([uid, uid])

    # Common filters
    if cid:   base += " AND a.company_id = %s"; params.append(cid)
    if sev:   base += " AND a.severity = %s";   params.append(sev)
    if atype: base += " AND a.entity_type = %s";params.append(atype)

    base += " ORDER BY CASE a.severity WHEN 'critical' THEN 1 WHEN 'high' THEN 2 WHEN 'medium' THEN 3 ELSE 4 END, a.due_date"
    c.execute(base, params)
    return jsonify(rows(c.fetchall()))

@app.route("/api/alerts/<aid>/dismiss", methods=["POST"])
@login_required
def dismiss_alert(aid):
    conn=get_db(); c=conn.cursor()
    c.execute("UPDATE alerts SET status='dismissed',resolved_at=NOW() WHERE id=%s",(aid,))
    conn.commit(); conn.close(); return jsonify({"success":True})

@app.route("/api/alerts/<aid>/resolve", methods=["POST"])
@login_required
def resolve_alert(aid):
    conn=get_db(); c=conn.cursor()
    c.execute("UPDATE alerts SET status='resolved',resolved_at=NOW() WHERE id=%s",(aid,))
    conn.commit(); conn.close(); return jsonify({"success":True})

# ══ TASKS ════════════════════════════════════════════════════════════════════
@app.route("/api/tasks")
@login_required
def list_tasks():
    cid=request.args.get("company_id",""); status=request.args.get("status","")
    conn=get_db(); c=conn.cursor()
    q="""SELECT t.*,
               co.name  AS company_name,
               ul.name  AS task_leader_name,
               um.name  AS task_manager_name,
               ua.name  AS assigned_to_name
               FROM tasks t
               LEFT JOIN companies co ON t.company_id=co.id
               LEFT JOIN users ul ON t.task_leader=ul.id
               LEFT JOIN users um ON t.task_manager=um.id
               LEFT JOIN users ua ON t.assigned_to=ua.id
               WHERE 1=1"""
    params=[]
    if cid: q+=" AND t.company_id=%s"; params.append(cid)
    if status: q+=" AND t.status=%s"; params.append(status)
    q+=" ORDER BY CASE t.priority WHEN 'critical' THEN 1 WHEN 'high' THEN 2 WHEN 'medium' THEN 3 ELSE 4 END,t.due_date"
    c.execute(q,params); return jsonify(rows(c.fetchall()))

@app.route("/api/tasks", methods=["POST"])
@login_required
def create_task():
    d=request.get_json(silent=True, force=True) or {}; tid=str(uuid.uuid4()); conn=get_db(); c=conn.cursor()
    c.execute("""INSERT INTO tasks (id,company_id,title,description,assigned_to,due_date,priority,status,module,entity_id,created_by)
                 VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
        (tid,d.get("company_id"),d["title"],d.get("description"),d.get("assigned_to"),
         d.get("due_date"),d.get("priority","medium"),d.get("status","pending"),
         d.get("module"),d.get("entity_id"),g.user_id))
    conn.commit()
    c.execute("SELECT t.*,co.name as company_name FROM tasks t LEFT JOIN companies co ON t.company_id=co.id WHERE t.id=%s",(tid,))
    result=row(c.fetchone()); conn.close(); return jsonify(result),201

@app.route("/api/tasks/<tid>", methods=["PUT"])
@login_required
def update_task(tid):
    d=request.get_json(silent=True, force=True) or {}
    fields={k:d[k] for k in ["title","description","assigned_to","due_date","priority","status","module"] if k in d}
    if d.get("status")=="completed": fields["completed_at"]=datetime.utcnow().isoformat()
    conn=get_db(); c=conn.cursor()
    c.execute(f"UPDATE tasks SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",list(fields.values())+[tid])
    conn.commit()
    c.execute("SELECT t.*,co.name as company_name FROM tasks t LEFT JOIN companies co ON t.company_id=co.id WHERE t.id=%s",(tid,))
    result=row(c.fetchone()); conn.close(); return jsonify(result)

@app.route("/api/tasks/<tid>", methods=["DELETE"])
@login_required
def delete_task(tid):
    conn=get_db(); c=conn.cursor(); c.execute("DELETE FROM tasks WHERE id=%s",(tid,)); conn.commit(); conn.close()
    return jsonify({"success":True})


# ══ EXCEL EXPORT ══════════════════════════════════════════════════════════════

@app.route("/api/export/excel", methods=["POST"])
@login_required
def export_excel():
    """
    Universal Excel export endpoint.
    Body: { module, company_id (optional), date_from, date_to, company_name }
    Returns a formatted .xlsx file.
    """
    d = request.get_json(silent=True, force=True) or {}
    module      = d.get("module", "")
    company_id  = d.get("company_id", "")
    date_from   = d.get("date_from", "")
    date_to     = d.get("date_to", "")
    co_name     = d.get("company_name", "All Companies")

    conn = get_db(); c = conn.cursor()

    # ── Date filter helper ────────────────────────────────────────────────────
    def date_filter(col, params):
        clause = ""
        if date_from:
            clause += f" AND {col} >= %s"
            params.append(date_from)
        if date_to:
            clause += f" AND {col} <= %s"
            params.append(date_to)
        return clause

    def co_filter(col, params):
        if company_id:
            params.append(company_id)
            return f" AND {col} = %s"
        return ""

    # ── Fetch data per module ─────────────────────────────────────────────────
    sheets = {}

    if module in ("companies", "all"):
        params = []
        q = "SELECT name,cin,company_type,pan,tan,gstin,email,phone,incorporation_date,registered_office,authorized_capital,paid_up_capital,business_activity,roc,status,created_at FROM companies WHERE 1=1"
        q += date_filter("created_at", params)
        c.execute(q, params)
        sheets["Companies"] = {
            "headers": ["Company Name","CIN","Type","PAN","TAN","GSTIN","Email","Phone","Incorporation Date","Registered Office","Auth Capital","Paid-up Capital","Business Activity","ROC","Status","Created"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("directors", "all"):
        params = []
        q = """SELECT d.name,d.din,d.pan,d.designation,co.name,d.email,d.mobile,
                      d.date_of_appointment,d.date_of_cessation,
                      k.kyc_status,k.last_kyc_date,k.next_due_date,
                      CASE d.is_active WHEN 1 THEN 'Active' ELSE 'Inactive' END
               FROM directors d
               LEFT JOIN director_kyc k ON d.id=k.director_id
               LEFT JOIN companies co ON d.company_id=co.id WHERE 1=1"""
        q += co_filter("d.company_id", params)
        q += date_filter("d.date_of_appointment", params)
        q += " ORDER BY co.name,d.name"
        c.execute(q, params)
        sheets["Directors"] = {
            "headers": ["Name","DIN","PAN","Designation","Company","Email","Mobile","Date Appointed","Date Ceased","KYC Status","Last KYC","KYC Due Date","Status"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("auditors", "all"):
        params = []
        q = """SELECT a.name,a.firm_name,a.membership_no,a.frn,a.pan,co.name,
                      a.nature_of_appointment,a.appointment_type,
                      a.start_date,a.end_date,a.srn_adt1,a.email,
                      CASE a.is_active WHEN 1 THEN 'Active' ELSE 'Inactive' END
               FROM auditors a LEFT JOIN companies co ON a.company_id=co.id WHERE 1=1"""
        q += co_filter("a.company_id", params)
        q += date_filter("a.start_date", params)
        q += " ORDER BY co.name,a.start_date"
        c.execute(q, params)
        sheets["Auditors"] = {
            "headers": ["Auditor Name","Firm Name","Membership No","FRN","PAN","Company","Nature","Appointment Type","Start Date","End Date","ADT-1 SRN","Email","Status"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("shareholders", "all"):
        params = []
        q = """SELECT s.name,s.folio_no,s.pan,co.name,s.share_class,
                      s.shares_held,s.face_value,s.email,s.mobile,s.date_of_entry,
                      CASE s.is_active WHEN 1 THEN 'Active' ELSE 'Inactive' END
               FROM shareholders s LEFT JOIN companies co ON s.company_id=co.id WHERE 1=1"""
        q += co_filter("s.company_id", params)
        q += date_filter("s.date_of_entry", params)
        q += " ORDER BY co.name,s.folio_no"
        c.execute(q, params)
        sheets["Shareholders"] = {
            "headers": ["Name","Folio No","PAN","Company","Share Class","Shares Held","Face Value","Email","Mobile","Date of Entry","Status"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("dsc", "all"):
        params = []
        q = """SELECT d.holder_name,d.holder_type,d.dsc_class,d.issued_by,
                      d.token_type,d.valid_from,d.valid_to,
                      d.custody_status,d.custody_date,co.name,d.notes
               FROM dsc_records d LEFT JOIN companies co ON d.company_id=co.id WHERE d.is_active=1"""
        q += co_filter("d.company_id", params)
        q += date_filter("d.valid_from", params)
        q += " ORDER BY d.valid_to"
        c.execute(q, params)
        sheets["DSC Records"] = {
            "headers": ["Holder Name","Holder Type","Class","Issued By","Token","Valid From","Valid To","Custody Status","Custody Date","Company","Notes"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("meetings", "all"):
        params = []
        q = """SELECT m.meeting_type,m.meeting_no,m.meeting_date,m.meeting_time,
                      m.venue,co.name,m.status,
                      CASE m.minutes_drafted WHEN 1 THEN 'Yes' ELSE 'No' END,
                      m.agenda
               FROM meetings m LEFT JOIN companies co ON m.company_id=co.id WHERE 1=1"""
        q += co_filter("m.company_id", params)
        q += date_filter("m.meeting_date", params)
        q += " ORDER BY m.meeting_date DESC"
        c.execute(q, params)
        sheets["Meetings"] = {
            "headers": ["Type","Meeting No","Date","Time","Venue","Company","Status","Minutes Done","Agenda"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("alerts", "all"):
        params = []
        q = """SELECT a.title,a.entity_type,a.alert_type,a.severity,a.due_date,
                      a.status,co.name,a.message,a.created_at
               FROM alerts a LEFT JOIN companies co ON a.company_id=co.id WHERE 1=1"""
        q += co_filter("a.company_id", params)
        q += date_filter("a.due_date", params)
        q += " ORDER BY CASE a.severity WHEN 'critical' THEN 1 WHEN 'high' THEN 2 ELSE 3 END,a.due_date"
        c.execute(q, params)
        sheets["Alerts"] = {
            "headers": ["Title","Entity Type","Alert Type","Severity","Due Date","Status","Company","Message","Created"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("tasks", "all"):
        params = []
        q = """SELECT t.title,t.priority,t.status,t.due_date,t.module,
                      co.name,u.name,t.description,t.created_at,t.completed_at
               FROM tasks t
               LEFT JOIN companies co ON t.company_id=co.id
               LEFT JOIN users u ON t.assigned_to=u.id WHERE 1=1"""
        q += co_filter("t.company_id", params)
        q += date_filter("t.due_date", params)
        q += " ORDER BY CASE t.priority WHEN 'critical' THEN 1 WHEN 'high' THEN 2 WHEN 'medium' THEN 3 ELSE 4 END,t.due_date"
        c.execute(q, params)
        sheets["Tasks"] = {
            "headers": ["Title","Priority","Status","Due Date","Module","Company","Assigned To","Description","Created","Completed"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("charges", "all"):
        params = []
        q = """SELECT ch.charge_id,ch.charge_type,ch.charge_holder,ch.assets_charged,
                      ch.amount,ch.date_of_creation,ch.date_of_modification,
                      ch.date_of_satisfaction,ch.status,ch.srn,co.name,ch.remarks
               FROM charges ch LEFT JOIN companies co ON ch.company_id=co.id WHERE 1=1"""
        q += co_filter("ch.company_id", params)
        q += date_filter("ch.date_of_creation", params)
        q += " ORDER BY ch.date_of_creation"
        c.execute(q, params)
        sheets["Charges"] = {
            "headers": ["Charge ID","Type","Charge Holder","Assets Charged","Amount","Date Created","Date Modified","Date Satisfied","Status","SRN","Company","Remarks"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("investments", "all"):
        params = []
        q = """SELECT i.investee_name,i.investee_type,i.investment_type,i.amount,
                      i.date_of_investment,i.board_resolution_date,i.srn_mgb4,
                      i.purpose,co.name,i.remarks
               FROM investments i LEFT JOIN companies co ON i.company_id=co.id WHERE 1=1"""
        q += co_filter("i.company_id", params)
        q += date_filter("i.date_of_investment", params)
        q += " ORDER BY i.date_of_investment"
        c.execute(q, params)
        sheets["Investments"] = {
            "headers": ["Investee Name","Type","Investment Type","Amount","Date","Board Resolution Date","MGT-14 SRN","Purpose","Company","Remarks"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("loans", "all"):
        params = []
        q = """SELECT lg.party_name,lg.party_type,lg.transaction_type,lg.amount,
                      lg.date_of_transaction,lg.rate_of_interest,lg.repayment_date,
                      lg.outstanding_amount,lg.security,lg.status,
                      lg.board_resolution_date,co.name,lg.remarks
               FROM loans_guarantees lg LEFT JOIN companies co ON lg.company_id=co.id WHERE 1=1"""
        q += co_filter("lg.company_id", params)
        q += date_filter("lg.date_of_transaction", params)
        q += " ORDER BY lg.date_of_transaction"
        c.execute(q, params)
        sheets["Loans & Guarantees"] = {
            "headers": ["Party Name","Party Type","Transaction Type","Amount","Date","Interest %","Repayment Date","Outstanding","Security","Status","Board Resolution","Company","Remarks"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("rpt", "all"):
        params = []
        q = """SELECT r.party_name,r.relationship,r.nature_of_transaction,r.amount,
                      r.date_of_transaction,r.date_of_board_approval,
                      r.date_of_shareholders_approval,r.terms,co.name,r.remarks
               FROM related_party_transactions r LEFT JOIN companies co ON r.company_id=co.id WHERE 1=1"""
        q += co_filter("r.company_id", params)
        q += date_filter("r.date_of_transaction", params)
        q += " ORDER BY r.date_of_transaction"
        c.execute(q, params)
        sheets["Related Party Txns"] = {
            "headers": ["Party Name","Relationship","Nature","Amount","Date","Board Approval","Shareholder Approval","Terms","Company","Remarks"],
            "rows": [list(r) for r in c.fetchall()]
        }

    if module in ("kyc", "all"):
        params = []
        q = """SELECT d.name,d.din,d.pan,co.name,d.designation,
                      k.last_kyc_date,k.next_due_date,k.kyc_status,d.email,d.mobile
               FROM directors d
               LEFT JOIN director_kyc k ON d.id=k.director_id
               LEFT JOIN companies co ON d.company_id=co.id
               WHERE d.is_active=1"""
        q += co_filter("d.company_id", params)
        q += " ORDER BY k.kyc_status DESC,k.next_due_date"
        c.execute(q, params)
        sheets["Director KYC"] = {
            "headers": ["Director Name","DIN","PAN","Company","Designation","Last KYC Date","KYC Due Date","KYC Status","Email","Mobile"],
            "rows": [list(r) for r in c.fetchall()]
        }

    conn.close()

    if not sheets:
        return jsonify({"error": f"Unknown module: {module}"}), 400

    # ── Build Excel workbook ──────────────────────────────────────────────────
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True, "remove_timezone": True})

    # Formats
    fmt_title  = wb.add_format({"bold":True,"font_size":14,"font_color":"#0f2044","border":0})
    fmt_meta   = wb.add_format({"italic":True,"font_size":9,"font_color":"#5a6a88","border":0})
    fmt_hdr    = wb.add_format({"bold":True,"bg_color":"#0f2044","font_color":"#ffffff",
                                 "font_size":10,"border":1,"border_color":"#0f2044",
                                 "text_wrap":True,"valign":"vcenter","align":"center"})
    fmt_row_a  = wb.add_format({"font_size":10,"border":1,"border_color":"#e2e8f0",
                                 "bg_color":"#f8fafc","valign":"vcenter"})
    fmt_row_b  = wb.add_format({"font_size":10,"border":1,"border_color":"#e2e8f0",
                                 "bg_color":"#ffffff","valign":"vcenter"})
    fmt_num    = wb.add_format({"font_size":10,"border":1,"border_color":"#e2e8f0",
                                 "num_format":"#,##0.00","align":"right"})
    fmt_date   = wb.add_format({"font_size":10,"border":1,"border_color":"#e2e8f0",
                                 "align":"center"})
    fmt_red    = wb.add_format({"font_size":10,"border":1,"border_color":"#e2e8f0",
                                 "font_color":"#ef4444","bold":True})
    fmt_green  = wb.add_format({"font_size":10,"border":1,"border_color":"#e2e8f0",
                                 "font_color":"#10b981","bold":True})
    fmt_orange = wb.add_format({"font_size":10,"border":1,"border_color":"#e2e8f0",
                                 "font_color":"#f59e0b","bold":True})

    numeric_keywords = {"amount","capital","value","shares","rate","outstanding","granted","exercised","lapsed","price"}
    date_keywords    = {"date","from","to","appointed","ceased","entry","created","modified","satisfied","due","expiry","valid"}
    status_values    = {"active","inactive","open","satisfied","pending","overdue","critical","high","medium","low","filed","compliant"}

    period_str = ""
    if date_from or date_to:
        period_str = f"Period: {date_from or 'Start'} to {date_to or 'Today'}"

    for sheet_name, data in sheets.items():
        ws = wb.add_worksheet(sheet_name[:31])
        headers = data["headers"]
        rows    = data["rows"]

        # Title rows
        ws.merge_range(0, 0, 0, max(len(headers)-1,1), f"TAXLY-CMS — {sheet_name.upper()} REPORT", fmt_title)
        meta_parts = [f"Company: {co_name}"]
        if period_str: meta_parts.append(period_str)
        meta_parts.append(f"Generated: {date.today().strftime('%d %b %Y')}")
        ws.merge_range(1, 0, 1, max(len(headers)-1,1), "   |   ".join(meta_parts), fmt_meta)
        ws.merge_range(2, 0, 2, max(len(headers)-1,1), f"Total Records: {len(rows)}", fmt_meta)

        ws.set_row(0, 22); ws.set_row(1, 14); ws.set_row(2, 14); ws.set_row(3, 20)

        # Headers
        for ci, hdr in enumerate(headers):
            ws.write(3, ci, hdr, fmt_hdr)

        # Auto column widths
        col_widths = [max(len(str(h)), 8) for h in headers]

        # Data rows
        for ri, row_data in enumerate(rows):
            row_fmt = fmt_row_a if ri % 2 == 0 else fmt_row_b
            for ci, val in enumerate(row_data):
                val_str = str(val) if val is not None else ""
                hdr_lower = headers[ci].lower() if ci < len(headers) else ""

                # Choose format
                if val is None or val == "":
                    ws.write(ri+4, ci, "—", row_fmt)
                elif any(kw in hdr_lower for kw in numeric_keywords):
                    try:
                        ws.write_number(ri+4, ci, float(val), fmt_num)
                    except: ws.write(ri+4, ci, val_str, row_fmt)
                elif any(kw in hdr_lower for kw in date_keywords) and val_str and len(val_str) >= 8:
                    ws.write(ri+4, ci, val_str[:10], fmt_date)
                elif val_str.lower() in {"critical","expired","overdue","inactive","fail"}:
                    ws.write(ri+4, ci, val_str, fmt_red)
                elif val_str.lower() in {"active","valid","filed","compliant","completed","open"}:
                    ws.write(ri+4, ci, val_str, fmt_green)
                elif val_str.lower() in {"high","expiring_soon","pending","due_soon","in_progress"}:
                    ws.write(ri+4, ci, val_str, fmt_orange)
                else:
                    ws.write(ri+4, ci, val_str, row_fmt)

                col_widths[ci] = max(col_widths[ci], min(len(val_str), 40))

        # Set column widths
        for ci, w in enumerate(col_widths):
            ws.set_column(ci, ci, w + 2)

        # Freeze header rows
        ws.freeze_panes(4, 0)
        ws.autofilter(3, 0, 3, len(headers)-1)

    # Summary sheet for "all" exports
    if module == "all":
        ws_sum = wb.add_worksheet("Summary")
        ws_sum.set_column(0, 0, 28); ws_sum.set_column(1, 1, 12)
        ws_sum.merge_range(0, 0, 0, 1, "TAXLY-CMS — FULL EXPORT SUMMARY", fmt_title)
        ws_sum.merge_range(1, 0, 1, 1, f"Company: {co_name}   |   {period_str}   |   {date.today().strftime('%d %b %Y')}", fmt_meta)
        ws_sum.set_row(0, 22); ws_sum.set_row(1, 14); ws_sum.set_row(3, 18)
        ws_sum.write(3, 0, "Module", fmt_hdr); ws_sum.write(3, 1, "Records", fmt_hdr)
        for ri, (sname, sdata) in enumerate(sheets.items()):
            sum_fmt = fmt_row_a if ri%2==0 else fmt_row_b
            ws_sum.write(ri+4, 0, sname, sum_fmt)
            ws_sum.write_number(ri+4, 1, len(sdata["rows"]), fmt_num)

    wb.close()
    buf.seek(0)

    # Filename
    safe_co = co_name.replace(" ","_").replace("/","_")[:25]
    safe_mod = module.title()
    fname = f"TaxlyCMS_{safe_mod}_{safe_co}_{date.today().strftime('%Y%m%d')}.xlsx"

    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=fname)


# ══ BULK UPLOAD ═══════════════════════════════════════════════════════════════

# Module → (sheet_name, required_cols, optional_cols, insert_fn)
BULK_TEMPLATES = {
    "companies": {
        "sheet": "Companies",
        "headers": ["Company Name*","CIN","Company Type","PAN","TAN","GSTIN","Email","Phone",
                    "Incorporation Date (YYYY-MM-DD)","Registered Office","Authorised Capital","Paid-up Capital",
                    "Business Activity","ROC"],
        "sample": [["Innovate Solutions Pvt Ltd","U72200MH2018PTC309876","Private Limited","AABCI1234P",
                    "MUMA12345A","27AABCI1234P1Z5","cs@company.in","022-12345678",
                    "2018-03-15","12 Nariman Point, Mumbai","5000000","1000000","Software Development","RoC-Mumbai"]],
    },
    "directors": {
        "sheet": "Directors",
        "headers": ["Company CIN*","Director Name*","DIN","PAN","Designation","Email","Mobile",
                    "Date of Appointment (YYYY-MM-DD)","Address","MCA User ID","MCA Password"],
        "sample": [["U72200MH2018PTC309876","Arjun Mehta","00123456","ABCPM1234D",
                    "Managing Director","arjun@company.in","9876543210","2018-03-15",
                    "123 MG Road, Mumbai","arjun_mca","password123"]],
    },
    "auditors": {
        "sheet": "Auditors",
        "headers": ["Company CIN*","Auditor Name*","Firm Name","Membership No","FRN","PAN","Email",
                    "Nature of Appointment","Appointment Type","Start Date (YYYY-MM-DD)",
                    "End Date (YYYY-MM-DD)","ADT-1 SRN"],
        "sample": [["U72200MH2018PTC309876","CA Ramesh Gupta","Gupta & Associates","123456",
                    "012345W","ABCPG3456G","ca@firm.in",
                    "Subsequent Auditor","AGM Appointment","2024-09-30","2025-09-29","ADT1202312345"]],
    },
    "shareholders": {
        "sheet": "Shareholders",
        "headers": ["Company CIN*","Holder Name*","Folio No","PAN","Email","Mobile",
                    "Share Class","Shares Held","Face Value","Date of Entry (YYYY-MM-DD)",
                    "Address","MCA User ID","MCA Password"],
        "sample": [["U72200MH2018PTC309876","Arjun Mehta","F0001","ABCPM1234D",
                    "arjun@company.in","9876543210","Equity","50000","10","2018-03-15",
                    "123 MG Road, Mumbai","arjun_mca","password123"]],
    },
    "directors_kyc": {
        "sheet": "Director KYC",
        "headers": ["Director DIN*","Last KYC Date (YYYY-MM-DD)*"],
        "sample": [["00123456","2024-08-15"],["00234567","2024-09-10"]],
    },
    "meetings": {
        "sheet": "Meetings",
        "headers": ["Company CIN*","Meeting Type*","Meeting No","Date (YYYY-MM-DD)*","Time (HH:MM)",
                    "Venue","Status","Agenda"],
        "sample": [["U72200MH2018PTC309876","Board","BM-2025-01","2025-04-15","11:00",
                    "Registered Office, Mumbai","scheduled",
                    "1. Financial review\n2. Director KYC status"]],
    },
    "tasks": {
        "sheet": "Tasks",
        "headers": ["Title*","Company CIN","Priority","Status","Due Date (YYYY-MM-DD)","Module","Description"],
        "sample": [["File ADT-1","U72200MH2018PTC309876","high","pending","2025-09-30","auditor",
                    "File ADT-1 for statutory auditor appointment"]],
    },
    "charges": {
        "sheet": "Charges",
        "headers": ["Company CIN*","Charge Holder*","Charge Type","Amount","Date of Creation (YYYY-MM-DD)",
                    "Assets Charged","Status","Charge ID","SRN"],
        "sample": [["U72200MH2018PTC309876","State Bank of India","Hypothecation","10000000",
                    "2024-01-15","Plant & Machinery","Open","CHG-001","CHG120234567"]],
    },
    "investments": {
        "sheet": "Investments",
        "headers": ["Company CIN*","Investee Name*","Investee Type","Investment Type","Amount",
                    "Date of Investment (YYYY-MM-DD)","Board Resolution Date (YYYY-MM-DD)","Purpose"],
        "sample": [["U72200MH2018PTC309876","TechVentures Ltd","Company","Equity Shares",
                    "500000","2024-01-15","2024-01-10","Strategic investment"]],
    },
}

@app.route("/api/bulk/template/<module>")
@login_required
def bulk_template(module):
    """Download pre-formatted Excel template for bulk upload."""
    tmpl = BULK_TEMPLATES.get(module)
    if not tmpl: return jsonify({"error": f"No template for module: {module}"}), 400

    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})

    # Formats
    fmt_title   = wb.add_format({"bold":True,"font_size":14,"font_color":"#0f2044"})
    fmt_info    = wb.add_format({"italic":True,"font_size":9,"font_color":"#5a6a88"})
    fmt_hdr     = wb.add_format({"bold":True,"bg_color":"#0f2044","font_color":"#ffffff",
                                  "font_size":10,"border":1,"border_color":"#0f2044",
                                  "text_wrap":True,"valign":"vcenter","align":"center"})
    fmt_hdr_req = wb.add_format({"bold":True,"bg_color":"#ef4444","font_color":"#ffffff",
                                  "font_size":10,"border":1,"border_color":"#c53030",
                                  "text_wrap":True,"valign":"vcenter","align":"center"})
    fmt_sample  = wb.add_format({"font_size":10,"bg_color":"#f0fdf4","border":1,
                                  "border_color":"#bbf7d0","font_color":"#166534"})
    fmt_input   = wb.add_format({"font_size":10,"bg_color":"#fafafa","border":1,
                                  "border_color":"#e2e8f0"})

    ws = wb.add_worksheet(tmpl["sheet"])
    headers = tmpl["headers"]

    # Title
    ws.merge_range(0,0,0,len(headers)-1, f"TAXLY-CMS — {tmpl['sheet'].upper()} BULK UPLOAD TEMPLATE", fmt_title)
    ws.merge_range(1,0,1,len(headers)-1,
        "▶ Red headers = Required  |  Grey = Optional  |  Green rows = Sample data (delete before upload)", fmt_info)
    ws.merge_range(2,0,2,len(headers)-1,
        "• Date format: YYYY-MM-DD  |  Do not change column order  |  Upload via Bulk Upload button", fmt_info)
    ws.set_row(0, 22); ws.set_row(1, 14); ws.set_row(2, 14); ws.set_row(3, 22)

    # Headers
    for ci, hdr in enumerate(headers):
        is_req = hdr.endswith("*")
        label  = hdr.rstrip("*")
        ws.write(3, ci, label, fmt_hdr_req if is_req else fmt_hdr)
        ws.set_column(ci, ci, max(len(label)+4, 14))

    # Sample rows
    for ri, row_data in enumerate(tmpl["sample"]):
        for ci, val in enumerate(row_data):
            ws.write(ri+4, ci, val, fmt_sample)

    # Blank input rows (10 rows ready for data)
    for ri in range(len(tmpl["sample"]), len(tmpl["sample"])+15):
        for ci in range(len(headers)):
            ws.write(ri+4, ci, "", fmt_input)

    ws.freeze_panes(4, 0)
    wb.close()
    buf.seek(0)

    fname = f"TaxlyCMS_Upload_Template_{module.title()}.xlsx"
    return send_file(buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name=fname)


@app.route("/api/bulk/upload/<module>", methods=["POST"])
@login_required
def bulk_upload(module):
    """Process bulk upload Excel file."""
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files["file"]
    if not f.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Only .xlsx files supported"}), 400

    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb.active
        rows_data = list(ws.iter_rows(min_row=5, values_only=True))  # Skip 4 header rows
    except Exception as e:
        return jsonify({"error": f"Cannot read file: {str(e)}"}), 400

    conn = get_db(); c = conn.cursor()
    inserted = 0; skipped = 0; errors = []

    def safe(v, default=""):
        if v is None: return default
        s = str(v).strip()
        return s if s and s.lower() not in ("none","null","nan") else default

    def get_co_id(cin):
        c.execute("SELECT id FROM companies WHERE cin=%s", (cin,))
        r = c.fetchone()
        return r[0] if r else None

    try:
        for ri, row in enumerate(rows_data, start=5):
            if not row or all(v is None or str(v).strip() == "" for v in row):
                continue  # skip empty rows

            try:
                if module == "companies":
                    name = safe(row[0])
                    if not name: skipped += 1; continue
                    c.execute("""INSERT INTO companies
                        (id,name,cin,company_type,pan,tan,gstin,email,phone,
                         incorporation_date,registered_office,authorized_capital,
                         paid_up_capital,business_activity,roc,created_by)
                        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s)""",
                        (str(uuid.uuid4()), name, safe(row[1]),
                         safe(row[2],"Private Limited"),
                         (safe(row[3])).upper() or None, (safe(row[4])).upper() or None,
                         safe(row[5]), safe(row[6]), safe(row[7]),
                         safe(row[8]), safe(row[9]),
                         float(safe(row[10]) or 0), float(safe(row[11]) or 0),
                         safe(row[12]), safe(row[13]), g.user_id))
                    if c.rowcount: inserted += 1
                    else: skipped += 1

                elif module == "directors":
                    cin  = safe(row[0]); name = safe(row[1])
                    if not cin or not name: skipped += 1; continue
                    co_id = get_co_id(cin)
                    if not co_id:
                        errors.append(f"Row {ri}: Company CIN '{cin}' not found"); skipped += 1; continue
                    did = str(uuid.uuid4())
                    c.execute("""INSERT INTO directors
                        (id,company_id,name,din,pan,designation,email,mobile,
                         date_of_appointment,address,mca_user_id,mca_password)
                        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s)""",
                        (did, co_id, name, safe(row[2]), (safe(row[3])).upper() or None,
                         safe(row[4],"Director"), safe(row[5]), safe(row[6]),
                         safe(row[7]), safe(row[8]), safe(row[9]), safe(row[10])))
                    from datetime import date as _date
                    nd = f"{_date.today().year}-09-30"
                    c.execute("INSERT INTO director_kyc (id,director_id,next_due_date,kyc_status) VALUES (%s,%s,?,%s)",
                              (str(uuid.uuid4()), did, nd, "pending"))
                    inserted += 1

                elif module == "directors_kyc":
                    din = safe(row[0]); kyc_date = safe(row[1])
                    if not din or not kyc_date: skipped += 1; continue
                    c.execute("SELECT id FROM directors WHERE din=%s", (din,))
                    dr = c.fetchone()
                    if not dr:
                        errors.append(f"Row {ri}: DIN '{din}' not found"); skipped += 1; continue
                    from datetime import date as _date
                    c.execute("""UPDATE director_kyc SET last_kyc_date=%s,kyc_status='filed',
                                 next_due_date=%s,updated_at=NOW()
                                 WHERE director_id=%s""",
                              (kyc_date, f"{_date.today().year+1}-09-30", dr[0]))
                    if c.rowcount: inserted += 1
                    else: skipped += 1

                elif module == "shareholders":
                    cin  = safe(row[0]); name = safe(row[1])
                    if not cin or not name: skipped += 1; continue
                    co_id = get_co_id(cin)
                    if not co_id:
                        errors.append(f"Row {ri}: Company CIN '{cin}' not found"); skipped += 1; continue
                    c.execute("""INSERT INTO shareholders
                        (id,company_id,name,folio_no,pan,email,mobile,share_class,
                         shares_held,face_value,date_of_entry,address,mca_user_id,mca_password)
                        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,?,%s)""",
                        (str(uuid.uuid4()), co_id, name, safe(row[2]),
                         (safe(row[3])).upper() or None, safe(row[4]), safe(row[5]),
                         safe(row[6],"Equity"), int(float(safe(row[7]) or 0)),
                         float(safe(row[8]) or 10), safe(row[9]), safe(row[10]),
                         safe(row[11]), safe(row[12])))
                    inserted += 1

                elif module == "auditors":
                    cin  = safe(row[0]); name = safe(row[1])
                    if not cin or not name: skipped += 1; continue
                    co_id = get_co_id(cin)
                    if not co_id:
                        errors.append(f"Row {ri}: Company CIN '{cin}' not found"); skipped += 1; continue
                    c.execute("""INSERT INTO auditors
                        (id,company_id,name,firm_name,membership_no,frn,pan,email,
                         nature_of_appointment,appointment_type,start_date,end_date,srn_adt1)
                        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
                        (str(uuid.uuid4()), co_id, name, safe(row[2]), safe(row[3]),
                         safe(row[4]), (safe(row[5])).upper() or None, safe(row[6]),
                         safe(row[7],"Subsequent Auditor"), safe(row[8],"AGM Appointment"),
                         safe(row[9]), safe(row[10]), safe(row[11])))
                    inserted += 1

                elif module == "meetings":
                    cin  = safe(row[0]); mtype = safe(row[1]); mdate = safe(row[3])
                    if not cin or not mtype or not mdate: skipped += 1; continue
                    co_id = get_co_id(cin)
                    if not co_id:
                        errors.append(f"Row {ri}: Company CIN '{cin}' not found"); skipped += 1; continue
                    c.execute("""INSERT INTO meetings
                        (id,company_id,meeting_type,meeting_no,meeting_date,meeting_time,
                         venue,status,agenda,created_by)
                        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s)""",
                        (str(uuid.uuid4()), co_id, mtype, safe(row[2]), mdate,
                         safe(row[4],""), safe(row[5],""), safe(row[6],"scheduled"),
                         safe(row[7]), g.user_id))
                    inserted += 1

                elif module == "tasks":
                    title = safe(row[0])
                    if not title: skipped += 1; continue
                    cin   = safe(row[1])
                    co_id = get_co_id(cin) if cin else None
                    c.execute("""INSERT INTO tasks
                        (id,company_id,title,priority,status,due_date,module,description,created_by)
                        VALUES (%s,%s,?,%s,?,%s,?,%s,%s)""",
                        (str(uuid.uuid4()), co_id, title,
                         safe(row[2],"medium"), safe(row[3],"pending"),
                         safe(row[4]), safe(row[5]), safe(row[6]), g.user_id))
                    inserted += 1

                elif module == "charges":
                    cin    = safe(row[0]); holder = safe(row[1])
                    if not cin or not holder: skipped += 1; continue
                    co_id = get_co_id(cin)
                    if not co_id:
                        errors.append(f"Row {ri}: Company CIN '{cin}' not found"); skipped += 1; continue
                    c.execute("""INSERT INTO charges
                        (id,company_id,charge_holder,charge_type,amount,date_of_creation,
                         assets_charged,status,charge_id,srn,created_by)
                        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
                        (str(uuid.uuid4()), co_id, holder,
                         safe(row[2],"Hypothecation"), float(safe(row[3]) or 0),
                         safe(row[4]), safe(row[5]), safe(row[6],"Open"),
                         safe(row[7]), safe(row[8]), g.user_id))
                    inserted += 1

                elif module == "investments":
                    cin  = safe(row[0]); name = safe(row[1])
                    if not cin or not name: skipped += 1; continue
                    co_id = get_co_id(cin)
                    if not co_id:
                        errors.append(f"Row {ri}: Company CIN '{cin}' not found"); skipped += 1; continue
                    c.execute("""INSERT INTO investments
                        (id,company_id,investee_name,investee_type,investment_type,amount,
                         date_of_investment,board_resolution_date,purpose,created_by)
                        VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s)""",
                        (str(uuid.uuid4()), co_id, name,
                         safe(row[2],"Company"), safe(row[3],"Equity Shares"),
                         float(safe(row[4]) or 0), safe(row[5]),
                         safe(row[6]), safe(row[7]), g.user_id))
                    inserted += 1

                else:
                    conn.close()
                    return jsonify({"error": f"Upload not supported for module: {module}"}), 400

            except Exception as row_err:
                errors.append(f"Row {ri}: {str(row_err)}")
                skipped += 1

        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({"error": str(e)}), 500

    conn.close()
    return jsonify({
        "success": True,
        "inserted": inserted,
        "skipped": skipped,
        "errors": errors[:20],  # cap error list
        "message": f"✓ {inserted} record(s) uploaded successfully. {skipped} skipped."
    })


# ══ SMART DOCUMENT GENERATION WITH PERIOD RANGE ══════════════════════════════

@app.route("/api/documents/generate", methods=["POST"])
@login_required
def generate_doc():
    """Generate document with period-range-aware smart placeholders."""
    d = request.get_json(silent=True, force=True) or {}
    tid   = d.get("template_id")
    cid   = d.get("company_id")
    d_id  = d.get("director_id")
    d2_id = d.get("director2_id")
    a_id  = d.get("auditor_id")
    extra = d.get("extra_context") or {}
    date_from = d.get("date_from","")
    date_to   = d.get("date_to","")
    rich_mode = d.get("rich_mode", False)  # True = return HTML, False = plain text

    if not tid or not cid:
        return jsonify({"error":"template_id and company_id required"}),400

    conn = get_db(); c = conn.cursor()
    c.execute("SELECT template_body,name,category FROM document_templates WHERE id=%s", (tid,))
    tpl = c.fetchone()
    if not tpl: conn.close(); return jsonify({"error":"Template not found"}),404

    body   = tpl[0] or ""
    tname  = tpl[1]
    tcat   = tpl[2]

    # ── Build base CRM context ─────────────────────────────────────────────
    from compliance import build_context
    try:
        ctx = build_context(cid, extra=extra,
                            director_id=d_id, director2_id=d2_id, auditor_id=a_id)
    except Exception as e:
        conn.close(); return jsonify({"error": str(e)}), 400

    # ── Smart dynamic placeholders with period range ───────────────────────
    # board_meetings_list / board_meetings_numbered / bm_dates_only
    mtg_params = [cid]
    mtg_sql = """SELECT meeting_type,meeting_no,meeting_date,meeting_time,venue,status
                 FROM meetings WHERE company_id=%s"""
    if date_from: mtg_sql += " AND meeting_date>=%s"; mtg_params.append(date_from)
    if date_to:   mtg_sql += " AND meeting_date<=%s"; mtg_params.append(date_to)
    mtg_sql += " ORDER BY meeting_date"
    c.execute(mtg_sql, mtg_params)
    meetings = [dict(zip([d[0] for d in c.description], row)) for row in c.fetchall()]

    # Build meeting smart placeholders
    bm_list = [m for m in meetings if 'board' in m['meeting_type'].lower()]
    agm_list = [m for m in meetings if 'agm' in m['meeting_type'].lower() or 'annual' in m['meeting_type'].lower()]
    all_list = meetings

    def fmt_meeting(i, m):
        d_str = m.get('meeting_date','')
        try:
            from datetime import datetime as _dt
            d_str = _dt.strptime(d_str, '%Y-%m-%d').strftime('%d %B %Y')
        except: pass
        return f"  BM-{i}: {d_str}" + (f" at {m['meeting_time']}" if m.get('meeting_time') else "") +                (f", {m['venue']}" if m.get('venue') else "")

    NL = chr(10)
    ctx["board_meetings_list"]     = NL.join(fmt_meeting(i+1, m) for i,m in enumerate(bm_list)) or "[No board meetings in this period]"
    ctx["board_meetings_numbered"] = NL.join(f"{i+1}. {(m.get('meeting_date') or '')}" for i,m in enumerate(bm_list)) or "[No board meetings]"
    ctx["board_meetings_count"]    = str(len(bm_list))
    ctx["all_meetings_list"]       = NL.join(fmt_meeting(i+1, m) for i,m in enumerate(all_list)) or "[No meetings in this period]"
    ctx["agm_meetings_list"]       = NL.join(fmt_meeting(i+1, m) for i,m in enumerate(agm_list)) or "[No AGMs in this period]"

    # Individual BM slots — unlimited, up to however many exist
    for i, m in enumerate(bm_list, 1):
        d_str = m.get('meeting_date','')
        try:
            from datetime import datetime as _dt
            d_str = _dt.strptime(d_str, '%Y-%m-%d').strftime('%d %B %Y')
        except: pass
        ctx[f"board_meeting_{i}"] = d_str
        ctx[f"bm_{i}_date"]       = d_str
        ctx[f"bm_{i}_venue"]      = m.get('venue','')
        ctx[f"bm_{i}_time"]       = m.get('meeting_time','')

    # DSC records in period
    dsc_params = [cid]
    dsc_sql = "SELECT holder_name,holder_type,valid_to FROM dsc_records WHERE company_id=%s AND is_active=1"
    if date_to: dsc_sql += " AND valid_to>=%s"; dsc_params.append(date_to)
    c.execute(dsc_sql, dsc_params)
    dsc_rows = c.fetchall()
    ctx["dsc_list"] = NL.join(f"  {r[0]} ({r[1]}) — valid to {r[2]}" for r in dsc_rows) or "[No DSC records]"

    # Charges
    chrg_params = [cid]
    chrg_sql = "SELECT charge_holder,charge_type,amount,date_of_creation FROM charges WHERE company_id=%s"
    if date_from: chrg_sql += " AND date_of_creation>=%s"; chrg_params.append(date_from)
    if date_to:   chrg_sql += " AND date_of_creation<=%s"; chrg_params.append(date_to)
    c.execute(chrg_sql, chrg_params)
    charges = c.fetchall()
    ctx["charges_list"] = NL.join(f"  {r[0]} — {r[1]} — ₹{r[2]:,.0f} (created {r[3]})" for r in charges) or "[No charges in this period]"

    # Director KYC status
    c.execute("""SELECT d.name,k.kyc_status,k.next_due_date
                 FROM directors d LEFT JOIN director_kyc k ON d.id=k.director_id
                 WHERE d.company_id=%s AND d.is_active=1""", (cid,))
    kyc_rows = c.fetchall()
    ctx["director_kyc_list"] = NL.join(f"  {r[0]}: {r[1] or 'pending'} (due {r[2] or 'N/A'})" for r in kyc_rows) or "[No directors]"

    ctx["period_from"] = date_from or ctx.get("financial_year","")
    ctx["period_to"]   = date_to   or ctx.get("year_ended_on","")

    conn.close()

    # ── Resolve all {{placeholders}} in body ──────────────────────────────
    import re as _re
    def resolve(body, ctx):
        def replace(m):
            key = m.group(1).strip()
            val = ctx.get(key)
            if val is not None:
                return str(val)
            return f"[{key} — TO BE FILLED]"
        return _re.sub(r'\{\{([^}]+)\}\}', replace, body)

    content = resolve(body, ctx)

    # Save to documents table
    c2 = get_db(); c3 = c2.cursor()
    c3.execute("""INSERT OR REPLACE INTO documents
        (id,company_id,template_id,doc_name,doc_type,content,generated_by,date_from,date_to)
        VALUES (%s,%s,?,%s,?,%s,?,%s,%s)""",
        (str(uuid.uuid4()), cid, tid, tname, tcat, content,
         g.user_id, date_from, date_to))
    c2.commit(); c2.close()

    return jsonify({"content": content, "context": {k:v for k,v in ctx.items() if isinstance(v,str) and len(v)<200}})


@app.route("/api/documents/generate/pdf", methods=["POST"])
@login_required
def generate_doc_pdf():
    """Generate PDF for a document."""
    from compliance import generate_document
    d = request.get_json(silent=True, force=True) or {}
    tid   = d.get("template_id")
    cid   = d.get("company_id")
    extra = d.get("extra_context") or {}
    date_from = d.get("date_from","")
    date_to   = d.get("date_to","")

    if not tid or not cid:
        return jsonify({"error":"template_id and company_id required"}),400

    # Re-use generate_doc to get resolved content
    from flask import current_app
    with current_app.test_request_context():
        pass

    # Call the text generator directly
    from compliance import build_context
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT template_body,name FROM document_templates WHERE id=%s", (tid,))
    tpl = c.fetchone()
    conn.close()
    if not tpl: return jsonify({"error":"Template not found"}),404

    try:
        ctx = build_context(cid, extra=extra,
                            director_id=d.get("director_id"),
                            director2_id=d.get("director2_id"),
                            auditor_id=d.get("auditor_id"))
        pdf_bytes = generate_document(tid, cid,
                                      director_id=d.get("director_id"),
                                      director2_id=d.get("director2_id"),
                                      auditor_id=d.get("auditor_id"),
                                      extra=extra)
        return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf",
                         as_attachment=True, download_name=f"{tpl[1]}.pdf")
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/custom-placeholders", methods=["GET"])
@login_required
def list_custom_placeholders():
    """List all manually defined custom placeholders."""
    conn = get_db(); c = conn.cursor()
    try:
        c.execute("SELECT * FROM custom_placeholders ORDER BY name")
        result = rows(c.fetchall())
    except:
        result = []
    conn.close()
    return jsonify(result)

@app.route("/api/custom-placeholders", methods=["POST"])
@login_required
def create_custom_placeholder():
    """Create a custom placeholder."""
    d = request.get_json(silent=True, force=True) or {}
    name  = d.get("name","").strip().lower().replace(" ","_")
    label = d.get("label","").strip()
    desc  = d.get("description","").strip()
    ph_type = d.get("ph_type","text")
    if not name or not label:
        return jsonify({"error":"name and label required"}),400
    conn = get_db(); c = conn.cursor()
    try:
        c.execute("""INSERT INTO custom_placeholders (id,name,label,description,ph_type,created_by)
                     VALUES (%s,%s,?,%s,?,%s)""",
                  (str(uuid.uuid4()), name, label, desc, ph_type, g.user_id))
        conn.commit()
    except Exception as e:
        conn.close(); return jsonify({"error":str(e)}),400
    conn.close()
    return jsonify({"success":True,"name":name})

@app.route("/api/custom-placeholders/<pid>", methods=["DELETE"])
@login_required
def delete_custom_placeholder(pid):
    conn = get_db(); c = conn.cursor()
    c.execute("DELETE FROM custom_placeholders WHERE id=%s", (pid,))
    conn.commit(); conn.close()
    return jsonify({"success":True})


# ══ USER PERMISSION MANAGEMENT ═══════════════════════════════════════════════

@app.route("/api/permissions/schema")
@login_required
def get_perm_schema():
    """Return all modules and their actions for UI rendering."""
    return jsonify({
        "modules": ALL_MODULES,
        "defaults": DEFAULT_PERMISSIONS,
    })

@app.route("/api/users/<uid>/permissions")
@login_required
def get_user_permissions(uid):
    """Get all permission overrides for a specific user."""
    if g.role != "superadmin":
        return jsonify({"error": "Only Super Admin can view permissions"}), 403
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT * FROM users WHERE id=%s", (uid,))
    user = row(c.fetchone())
    if not user: conn.close(); return jsonify({"error": "User not found"}), 404
    c.execute("""SELECT module, action, granted, granted_by, granted_at, note
                 FROM user_permissions WHERE user_id=%s ORDER BY module, action""", (uid,))
    overrides = [dict(zip([d[0] for d in c.description], r)) for r in c.fetchall()]
    conn.close()
    return jsonify({
        "user": user,
        "overrides": overrides,
        "role": user["role"],
    })

@app.route("/api/users/<uid>/permissions", methods=["POST"])
@login_required
def set_user_permissions(uid):
    """Set/update permission overrides for a user (superadmin only)."""
    if g.role != "superadmin":
        return jsonify({"error": "Only Super Admin can modify permissions"}), 403
    d = request.get_json(silent=True, force=True) or {}
    # d = { "permissions": [ {"module": "company", "action": "delete", "granted": true/false, "note": ""}, ... ] }
    permissions = d.get("permissions", [])
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT id, role FROM users WHERE id=%s", (uid,))
    user = c.fetchone()
    if not user: conn.close(); return jsonify({"error": "User not found"}), 404
    # Cannot set permissions for superadmin
    if user[1] == "superadmin":
        conn.close(); return jsonify({"error": "Cannot restrict Super Admin permissions"}), 400
    for perm in permissions:
        module  = perm.get("module","").strip()
        action  = perm.get("action","").strip()
        granted = 1 if perm.get("granted", True) else 0
        note    = perm.get("note", "")
        if not module or not action: continue
        c.execute("""INSERT INTO user_permissions (id, user_id, module, action, granted, granted_by, note)
                     VALUES (%s, ?, ?, ?, ?, ?, ?)
                     ON CONFLICT(user_id, module, action)
                     DO UPDATE SET granted=excluded.granted, granted_by=excluded.granted_by,
                                   note=excluded.note, granted_at=CURRENT_TIMESTAMP""",
                  (str(uuid.uuid4()), uid, module, action, granted, g.user_id, note))
    conn.commit(); conn.close()
    return jsonify({"success": True, "updated": len(permissions)})

@app.route("/api/users/<uid>/permissions/reset", methods=["POST"])
@login_required
def reset_user_permissions(uid):
    """Remove ALL overrides for a user (revert to role defaults)."""
    if g.role != "superadmin":
        return jsonify({"error": "Only Super Admin can reset permissions"}), 403
    d = request.get_json(silent=True, force=True) or {}
    module = d.get("module")  # optional: reset only one module
    conn = get_db(); c = conn.cursor()
    if module:
        c.execute("DELETE FROM user_permissions WHERE user_id=%s AND module=%s", (uid, module))
    else:
        c.execute("DELETE FROM user_permissions WHERE user_id=%s", (uid,))
    conn.commit(); conn.close()
    return jsonify({"success": True})

@app.route("/api/permission-presets", methods=["GET"])
@login_required
def list_presets():
    """List all saved permission presets."""
    conn = get_db(); c = conn.cursor()
    try:
        c.execute("SELECT * FROM permission_presets ORDER BY name")
        result = rows(c.fetchall())
    except: result = []
    conn.close()
    return jsonify(result)

@app.route("/api/permission-presets", methods=["POST"])
@login_required
def create_preset():
    """Save current permission set as a named preset."""
    if g.role != "superadmin":
        return jsonify({"error": "Only Super Admin can create presets"}), 403
    d = request.get_json(silent=True, force=True) or {}
    name = d.get("name","").strip()
    if not name: return jsonify({"error": "Preset name required"}), 400
    conn = get_db(); c = conn.cursor()
    c.execute("""INSERT INTO permission_presets (id, name, description, permissions, created_by)
                 VALUES (%s,%s,?,%s,%s)""",
              (str(uuid.uuid4()), name, d.get("description",""),
               __import__("json").dumps(d.get("permissions",[])), g.user_id))
    conn.commit(); conn.close()
    return jsonify({"success": True})

@app.route("/api/permission-presets/<pid>", methods=["DELETE"])
@login_required
def delete_preset(pid):
    if g.role != "superadmin": return jsonify({"error":"Forbidden"}), 403
    conn = get_db(); c = conn.cursor()
    c.execute("DELETE FROM permission_presets WHERE id=%s", (pid,))
    conn.commit(); conn.close()
    return jsonify({"success": True})

@app.route("/api/users/<uid>/permissions/apply-preset/<pid>", methods=["POST"])
@login_required
def apply_preset(uid, pid):
    """Apply a named preset to a user."""
    if g.role != "superadmin": return jsonify({"error":"Forbidden"}), 403
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT permissions FROM permission_presets WHERE id=%s", (pid,))
    preset = c.fetchone()
    if not preset: conn.close(); return jsonify({"error":"Preset not found"}),404
    import json as _json
    permissions = _json.loads(preset[0])
    for perm in permissions:
        module = perm.get("module",""); action = perm.get("action","")
        granted = 1 if perm.get("granted", True) else 0
        if not module or not action: continue
        c.execute("""INSERT INTO user_permissions (id,user_id,module,action,granted,granted_by,note)
                     VALUES (%s,%s,?,%s,?,%s,%s)
                     ON CONFLICT(user_id,module,action)
                     DO UPDATE SET granted=excluded.granted, granted_by=excluded.granted_by,
                                   granted_at=CURRENT_TIMESTAMP""",
                  (str(uuid.uuid4()),uid,module,action,granted,g.user_id,f"Applied preset: {pid}"))
    conn.commit(); conn.close()
    return jsonify({"success":True,"applied":len(permissions)})


# ══ PERSONAL DASHBOARD ════════════════════════════════════════════════════════

@app.route("/api/my-dashboard")
@login_required
def my_dashboard():
    """Role-aware personal dashboard.
    
    Role mapping:
      superadmin → Task Leader   (task_leader = me)
      manager    → Task Manager  (task_manager = me)
      staff      → Assignee      (assigned_to = me)
    
    Each panel shows OTHER roles' names for context.
    """
    uid  = g.user_id
    role = g.role
    conn = get_db(); c = conn.cursor()
    from datetime import date, timedelta
    today = date.today()
    week  = today + timedelta(days=7)

    # Column filter based on role
    if role == "superadmin":
        my_col = "task_leader"
    elif role == "manager":
        my_col = "task_manager"
    else:  # staff / any other
        my_col = "assigned_to"

    # Task summary counts (only my tasks per role)
    c.execute(f"""SELECT status, COUNT(*) FROM tasks WHERE {my_col}=%s GROUP BY status""", (uid,))
    task_counts = {r[0]: r[1] for r in c.fetchall()}

    c.execute(f"""SELECT COUNT(*) FROM tasks WHERE {my_col}=%s
                  AND due_date=%s AND status NOT IN ('completed','cancelled')""",
              (uid, str(today)))
    due_today = c.fetchone()[0]

    c.execute(f"""SELECT COUNT(*) FROM tasks WHERE {my_col}=%s
                  AND due_date BETWEEN ? AND ? AND status NOT IN ('completed','cancelled')""",
              (uid, str(today), str(week)))
    due_week = c.fetchone()[0]

    c.execute(f"""SELECT COUNT(*) FROM tasks WHERE {my_col}=%s
                  AND due_date < ? AND status NOT IN ('completed','cancelled')""",
              (uid, str(today)))
    overdue = c.fetchone()[0]

    # Full task list — only tasks where I play MY role
    # Always join all three roles so frontend can show "who else is on this task"
    task_sql = f"""
        SELECT t.*,
               co.name  AS company_name,
               ul.name  AS task_leader_name,
               um.name  AS task_manager_name,
               ua.name  AS assigned_to_name,
               ul.id    AS leader_id,
               um.id    AS manager_id,
               ua.id    AS assignee_id
        FROM tasks t
        LEFT JOIN companies co ON t.company_id = co.id
        LEFT JOIN users ul ON t.task_leader  = ul.id
        LEFT JOIN users um ON t.task_manager = um.id
        LEFT JOIN users ua ON t.assigned_to  = ua.id
        WHERE t.{my_col} = %s
        ORDER BY
            CASE t.priority WHEN 'critical' THEN 1 WHEN 'high' THEN 2
                            WHEN 'medium'   THEN 3 ELSE 4 END,
            t.due_date
    """
    c.execute(task_sql, (uid,))
    my_tasks = [dict(zip([d[0] for d in c.description], r)) for r in c.fetchall()]

    # ── Role-wise team summary (superadmin + manager see this) ──────────────
    role_summary = []
    if role in ("superadmin", "manager"):
        c.execute("""
            SELECT u.id, u.name, u.role,
                   COUNT(CASE WHEN t.status='pending'     THEN 1 END) AS pending,
                   COUNT(CASE WHEN t.status='in_progress' THEN 1 END) AS in_progress,
                   COUNT(CASE WHEN t.status='completed'   THEN 1 END) AS completed,
                   COUNT(CASE WHEN t.due_date < CURRENT_DATE AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS overdue
            FROM users u
            LEFT JOIN tasks t ON (
                (u.role='superadmin' AND t.task_leader=u.id) OR
                (u.role='manager'    AND t.task_manager=u.id) OR
                (u.role='staff'      AND t.assigned_to=u.id)
            )
            WHERE u.is_active=1
            GROUP BY u.id
            ORDER BY CASE u.role WHEN 'superadmin' THEN 1 WHEN 'manager' THEN 2 ELSE 3 END, u.name
        """)
        role_summary = [dict(zip([d[0] for d in c.description], r)) for r in c.fetchall()]

    conn.close()
    return jsonify({
        "my_role":    role,
        "my_col":     my_col,
        "summary": {
            "pending":     task_counts.get("pending", 0),
            "in_progress": task_counts.get("in_progress", 0),
            "completed":   task_counts.get("completed", 0),
            "due_today":   due_today,
            "due_week":    due_week,
            "overdue":     overdue,
        },
        "my_tasks":     my_tasks,
        "role_summary": role_summary,
    })


# ══ ROLE-WISE TASK WIDGET ═════════════════════════════════════════════════════

@app.route("/api/tasks/rolewise-summary")
@login_required
def tasks_rolewise_summary():
    """
    HIERARCHY PENDENCY TRACKER:
    Shows the complete pending workload across the logged-in user's chain.

    Super Admin (Leader) sees:
      - Their own pending (task_leader = me)  [Leader column]
      - All managers' pending                  [Manager column]
      - All assignees' pending                 [Assignee column]

    Manager sees:
      - Their own pending (task_manager = me)  [Manager column]
      - All assignees under their tasks         [Assignee column]
      - Their leader's count                    [Leader column]

    Staff sees:
      - Only their own (assigned_to = me)       [Assignee column]
    
    Per module breakdown for the full hierarchy.
    """
    uid  = g.user_id
    role = g.role
    conn = get_db(); c = conn.cursor()

    # ── Build hierarchy scope ─────────────────────────────────────────────
    # Get all tasks visible in this user's chain
    if role == "superadmin":
        # Leader sees everything — all tasks where they are leader
        # PLUS counts for each subordinate role across those tasks
        c.execute("""
            SELECT
                COALESCE(t.module,'general') AS module,
                COUNT(CASE WHEN t.task_leader=%s  AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS leader_cnt,
                COUNT(CASE WHEN t.task_manager IS NOT NULL AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS manager_cnt,
                COUNT(CASE WHEN t.assigned_to  IS NOT NULL AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS assignee_cnt,
                COUNT(CASE WHEN t.status NOT IN ('completed','cancelled') THEN 1 END) AS total
            FROM tasks t
            WHERE t.task_leader=%s
            GROUP BY COALESCE(t.module,'general')
            ORDER BY total DESC
        """, (uid, uid))
    elif role == "manager":
        # Manager sees their own tasks + subordinate counts on those tasks
        c.execute("""
            SELECT
                COALESCE(t.module,'general') AS module,
                COUNT(CASE WHEN t.task_leader IS NOT NULL AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS leader_cnt,
                COUNT(CASE WHEN t.task_manager=%s AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS manager_cnt,
                COUNT(CASE WHEN t.assigned_to  IS NOT NULL AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS assignee_cnt,
                COUNT(CASE WHEN t.status NOT IN ('completed','cancelled') THEN 1 END) AS total
            FROM tasks t
            WHERE t.task_manager=%s
            GROUP BY COALESCE(t.module,'general')
            ORDER BY total DESC
        """, (uid, uid))
    else:
        # Staff sees only their own
        c.execute("""
            SELECT
                COALESCE(t.module,'general') AS module,
                0 AS leader_cnt,
                0 AS manager_cnt,
                COUNT(CASE WHEN t.assigned_to=%s AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS assignee_cnt,
                COUNT(CASE WHEN t.assigned_to=%s AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS total
            FROM tasks t
            WHERE t.assigned_to=%s
            GROUP BY COALESCE(t.module,'general')
            ORDER BY total DESC
        """, (uid, uid, uid))

    modules = [{"module":r[0],"leader":r[1],"manager":r[2],"assignee":r[3],"total":r[4]}
               for r in c.fetchall()]

    totals = {
        "leader":   sum(m["leader"]   for m in modules),
        "manager":  sum(m["manager"]  for m in modules),
        "assignee": sum(m["assignee"] for m in modules),
        "total":    sum(m["total"]    for m in modules),
    }

    # ── Per-person breakdown for the hierarchy strip ──────────────────────
    if role == "superadmin":
        # All people on tasks where I am leader
        c.execute("""
            SELECT u.id, u.name, u.role,
                COUNT(CASE WHEN (
                    (u.role='superadmin' AND t.task_leader=u.id) OR
                    (u.role='manager'    AND t.task_manager=u.id) OR
                    (u.role='staff'      AND t.assigned_to=u.id)
                ) AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS pending,
                COUNT(CASE WHEN (
                    (u.role='superadmin' AND t.task_leader=u.id) OR
                    (u.role='manager'    AND t.task_manager=u.id) OR
                    (u.role='staff'      AND t.assigned_to=u.id)
                ) AND t.status='completed' THEN 1 END) AS completed,
                COUNT(CASE WHEN (
                    (u.role='superadmin' AND t.task_leader=u.id) OR
                    (u.role='manager'    AND t.task_manager=u.id) OR
                    (u.role='staff'      AND t.assigned_to=u.id)
                ) AND t.due_date < CURRENT_DATE AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS overdue
            FROM users u
            JOIN tasks t ON t.task_leader=%s
            WHERE u.is_active=1
            GROUP BY u.id
            HAVING pending > 0 OR completed > 0
            ORDER BY CASE u.role WHEN 'superadmin' THEN 1 WHEN 'manager' THEN 2 ELSE 3 END, pending DESC
        """, (uid,))
    elif role == "manager":
        c.execute("""
            SELECT u.id, u.name, u.role,
                COUNT(CASE WHEN t.assigned_to=u.id AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS pending,
                COUNT(CASE WHEN t.assigned_to=u.id AND t.status='completed' THEN 1 END) AS completed,
                COUNT(CASE WHEN t.assigned_to=u.id AND t.due_date < CURRENT_DATE
                           AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS overdue
            FROM users u
            JOIN tasks t ON t.task_manager=%s
            WHERE u.is_active=1
            GROUP BY u.id
            HAVING pending > 0 OR completed > 0
            ORDER BY pending DESC
        """, (uid,))
    else:
        c.execute("""
            SELECT ? as id, u.name, u.role, 
                COUNT(CASE WHEN t.status NOT IN ('completed','cancelled') THEN 1 END) AS pending,
                COUNT(CASE WHEN t.status='completed' THEN 1 END) AS completed,
                COUNT(CASE WHEN t.due_date < CURRENT_DATE AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS overdue
            FROM users u JOIN tasks t ON t.assigned_to=u.id
            WHERE u.id=%s
            GROUP BY u.id
        """, (uid, uid))

    person_cols = [d[0] for d in c.description]
    persons = [dict(zip(person_cols, r)) for r in c.fetchall()]

    conn.close()
    return jsonify({
        "uid":     uid,
        "my_role": role,
        "totals":  totals,
        "modules": modules,
        "persons": persons,   # hierarchy chain
    })

@app.route("/api/tasks/by-role-module")
@login_required
def tasks_by_role_module():
    """Return tasks filtered by module + my role for the drill-down page."""
    uid    = g.user_id
    module = request.args.get("module","")
    role   = request.args.get("role","")    # leader | manager | assignee | all
    status = request.args.get("status","")

    conn = get_db(); c = conn.cursor()
    sql = """
        SELECT t.*,
               co.name AS company_name,
               ul.name AS task_leader_name,
               um.name AS task_manager_name,
               ua.name AS assigned_to_name
        FROM tasks t
        LEFT JOIN companies co ON t.company_id=co.id
        LEFT JOIN users ul ON t.task_leader=ul.id
        LEFT JOIN users um ON t.task_manager=um.id
        LEFT JOIN users ua ON t.assigned_to=ua.id
        WHERE 1=1
    """
    params = []
    if role == "leader":   sql += " AND t.task_leader=%s";  params.append(uid)
    elif role == "manager": sql += " AND t.task_manager=%s"; params.append(uid)
    elif role == "assignee":sql += " AND t.assigned_to=%s";  params.append(uid)
    else:                  sql += " AND (t.task_leader=%s OR t.task_manager=%s OR t.assigned_to=%s)"; params.extend([uid,uid,uid])
    if module and module != 'all': sql += " AND COALESCE(t.module,'general')=%s"; params.append(module)
    if status: sql += " AND t.status=%s"; params.append(status)
    sql += " ORDER BY CASE t.priority WHEN 'critical' THEN 1 WHEN 'high' THEN 2 WHEN 'medium' THEN 3 ELSE 4 END, t.due_date"

    c.execute(sql, params)
    cols  = [d[0] for d in c.description]
    tasks = [dict(zip(cols, r)) for r in c.fetchall()]
    conn.close()
    return jsonify(tasks)


@app.route("/api/tasks/subordinate-view")
@login_required  
def tasks_subordinate_view():
    """
    Superior view: I can see ALL tasks in my chain of command.
    - If I am leader (superadmin): see all tasks I lead + their managers + assignees
    - If I am manager: see all tasks I manage + their assignees
    - Also returns subordinate pending summary grouped by person
    """
    uid  = g.user_id
    role = g.role
    conn = get_db(); c = conn.cursor()

    if role == "superadmin":
        # I am leader — see all tasks where task_leader = me
        c.execute("""
            SELECT t.*,
                   co.name AS company_name,
                   ul.name AS task_leader_name,
                   um.name AS task_manager_name, um.id AS manager_id,
                   ua.name AS assigned_to_name,  ua.id AS assignee_id
            FROM tasks t
            LEFT JOIN companies co ON t.company_id=co.id
            LEFT JOIN users ul ON t.task_leader=ul.id
            LEFT JOIN users um ON t.task_manager=um.id
            LEFT JOIN users ua ON t.assigned_to=ua.id
            WHERE t.task_leader=%s
            ORDER BY CASE t.priority WHEN 'critical' THEN 1 WHEN 'high' THEN 2
                                     WHEN 'medium' THEN 3 ELSE 4 END, t.due_date
        """, (uid,))
    elif role == "manager":
        # I am manager — see all tasks where task_manager = me
        c.execute("""
            SELECT t.*,
                   co.name AS company_name,
                   ul.name AS task_leader_name, ul.id AS leader_id,
                   um.name AS task_manager_name,
                   ua.name AS assigned_to_name,  ua.id AS assignee_id
            FROM tasks t
            LEFT JOIN companies co ON t.company_id=co.id
            LEFT JOIN users ul ON t.task_leader=ul.id
            LEFT JOIN users um ON t.task_manager=um.id
            LEFT JOIN users ua ON t.assigned_to=ua.id
            WHERE t.task_manager=%s
            ORDER BY CASE t.priority WHEN 'critical' THEN 1 WHEN 'high' THEN 2
                                     WHEN 'medium' THEN 3 ELSE 4 END, t.due_date
        """, (uid,))
    else:
        conn.close()
        return jsonify({"tasks": [], "subordinates": [], "my_role": role})

    cols  = [d[0] for d in c.description]
    tasks = [dict(zip(cols, r)) for r in c.fetchall()]

    # Subordinate summary — people below me and their pending counts
    if role == "superadmin":
        # Subordinates = all managers + assignees on my tasks
        c.execute("""
            SELECT u.id, u.name, u.role,
                   COUNT(CASE WHEN t.status NOT IN ('completed','cancelled') THEN 1 END) AS pending,
                   COUNT(CASE WHEN t.status='completed' THEN 1 END) AS completed,
                   COUNT(CASE WHEN t.due_date < CURRENT_DATE AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS overdue,
                   COUNT(*) AS total
            FROM users u
            JOIN tasks t ON (t.task_manager=u.id OR t.assigned_to=u.id)
            WHERE t.task_leader=%s AND u.id != %s AND u.is_active=1
            GROUP BY u.id
            ORDER BY pending DESC
        """, (uid, uid))
    else:  # manager
        # Subordinates = assignees on tasks I manage
        c.execute("""
            SELECT u.id, u.name, u.role,
                   COUNT(CASE WHEN t.status NOT IN ('completed','cancelled') THEN 1 END) AS pending,
                   COUNT(CASE WHEN t.status='completed' THEN 1 END) AS completed,
                   COUNT(CASE WHEN t.due_date < CURRENT_DATE AND t.status NOT IN ('completed','cancelled') THEN 1 END) AS overdue,
                   COUNT(*) AS total
            FROM users u
            JOIN tasks t ON t.assigned_to=u.id
            WHERE t.task_manager=%s AND u.id != %s AND u.is_active=1
            GROUP BY u.id
            ORDER BY pending DESC
        """, (uid, uid))

    sub_cols = [d[0] for d in c.description]
    subordinates = [dict(zip(sub_cols, r)) for r in c.fetchall()]

    conn.close()
    return jsonify({
        "my_role":     role,
        "tasks":       tasks,
        "subordinates": subordinates,
        "totals": {
            "total":     len(tasks),
            "pending":   sum(1 for t in tasks if t["status"] == "pending"),
            "in_progress": sum(1 for t in tasks if t["status"] == "in_progress"),
            "completed": sum(1 for t in tasks if t["status"] == "completed"),
            "overdue":   sum(1 for t in tasks if t.get("due_date","") < str(__import__("datetime").date.today()) and t["status"] not in ("completed","cancelled")),
        }
    })


# ══ PLATFORM ADMIN — TENANT & PLAN MANAGEMENT ════════════════════════════════

@app.route("/api/platform/plans")
@platform_admin_required
def list_plans():
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT * FROM plans ORDER BY price_monthly")
    return jsonify(rows(c.fetchall()))

@app.route("/api/platform/plans", methods=["POST"])
@platform_admin_required
def create_plan():
    d=request.get_json(silent=True,force=True) or {}
    if not d.get("name"): return jsonify({"error":"Name required"}),400
    pid=str(uuid.uuid4()); conn=get_db(); c=conn.cursor()
    c.execute("""INSERT INTO plans (id,name,description,price_monthly,price_annual,
                max_users,max_companies,max_documents,features,is_active)
                VALUES (%s,%s,?,%s,?,%s,?,%s,?,1)""",
             (pid,d["name"],d.get("description",""),
              float(d.get("price_monthly",0)),float(d.get("price_annual",0)),
              int(d.get("max_users",5)),int(d.get("max_companies",5)),
              int(d.get("max_documents",100)),
              __import__("json").dumps(d.get("features",[]))))
    conn.commit(); c.execute("SELECT * FROM plans WHERE id=%s", (pid,))
    result=row(c.fetchone()); conn.close(); return jsonify(result),201

@app.route("/api/platform/plans/<pid>", methods=["PUT"])
@platform_admin_required
def update_plan(pid):
    d=request.get_json(silent=True,force=True) or {}
    conn=get_db(); c=conn.cursor()
    fields={k:d[k] for k in ["name","description","price_monthly","price_annual",
                               "max_users","max_companies","max_documents","is_active"] if k in d}
    if "features" in d: fields["features"]=__import__("json").dumps(d["features"])
    if not fields: return jsonify({"error":"Nothing to update"}),400
    c.execute(f"UPDATE plans SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",
              list(fields.values())+[pid])
    conn.commit(); c.execute("SELECT * FROM plans WHERE id=%s", (pid,))
    result=row(c.fetchone()); conn.close(); return jsonify(result)

@app.route("/api/platform/plans/<pid>", methods=["DELETE"])
@platform_admin_required
def delete_plan(pid):
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT COUNT(*) FROM tenants WHERE plan_id=%s", (pid,))
    if c.fetchone()[0]: conn.close(); return jsonify({"error":"Plan in use by tenants"}),409
    c.execute("DELETE FROM plans WHERE id=%s", (pid,))
    conn.commit(); conn.close(); return jsonify({"success":True})

@app.route("/api/platform/tenants")
@platform_admin_required
def list_tenants():
    conn=get_db(); c=conn.cursor()
    c.execute("""SELECT t.*, p.name as plan_name, p.price_monthly,
                 (SELECT COUNT(*) FROM users u WHERE u.tenant_id=t.id AND u.is_active=1) as user_count,
                 (SELECT COUNT(*) FROM companies co WHERE co.tenant_id=t.id) as company_count
                 FROM tenants t LEFT JOIN plans p ON t.plan_id=p.id
                 ORDER BY t.created_at DESC""")
    return jsonify(rows(c.fetchall()))

@app.route("/api/platform/tenants", methods=["POST"])
@platform_admin_required
def create_tenant():
    d=request.get_json(silent=True,force=True) or {}
    if not d.get("name"):  return jsonify({"error":"Name required"}),400
    if not d.get("email"): return jsonify({"error":"Email required"}),400
    if not d.get("admin_name"): return jsonify({"error":"Admin name required"}),400
    if not d.get("admin_password"): return jsonify({"error":"Admin password required"}),400

    conn=get_db(); c=conn.cursor()
    # Get plan limits
    plan_id=d.get("plan_id")
    max_u=5; max_co=5
    if plan_id:
        c.execute("SELECT max_users,max_companies FROM plans WHERE id=%s", (plan_id,))
        pl=c.fetchone()
        if pl: max_u,max_co=pl[0],pl[1]

    # Create slug from name
    import re
    slug=re.sub(r'[^a-z0-9]+','-',d["name"].lower()).strip('-')
    base=slug
    for i in range(1,100):
        c.execute("SELECT 1 FROM tenants WHERE slug=%s", (slug,))
        if not c.fetchone(): break
        slug=f"{base}-{i}"

    tid=str(uuid.uuid4())
    try:
        c.execute("""INSERT INTO tenants (id,name,slug,email,phone,address,plan_id,
                     status,max_users,max_companies,notes)
                     VALUES (%s,%s,?,%s,?,%s,?,%s,?,%s,%s)""",
                 (tid,d["name"],slug,d["email"],d.get("phone",""),d.get("address",""),
                  plan_id,"active",max_u,max_co,d.get("notes","")))
    except Exception as e:
        conn.close(); return jsonify({"error":str(e)}),400

    # Create tenant superadmin
    aid=str(uuid.uuid4())
    admin_email=d.get("admin_email",d["email"]).strip().lower()
    c.execute("""INSERT INTO users (id,name,email,password,role,is_active,tenant_id)
                 VALUES (%s,%s,?,%s,?,1,%s)""",
             (aid,d["admin_name"],admin_email,hash_pw(d["admin_password"]),"superadmin",tid))

    # Seed default document templates for this tenant (each gets new UUID)
    c.execute("""SELECT id,name,category,description,template_body FROM document_templates
                 WHERE tenant_id='default-tenant-001' LIMIT 5""")
    for row_t in c.fetchall():
        c.execute("""INSERT INTO document_templates
                     (id,name,category,description,template_body,created_by,tenant_id)
                     VALUES (%s,%s,?,%s,?,%s,%s)""",
                 (str(uuid.uuid4()),row_t[1],row_t[2],row_t[3],row_t[4],aid,tid))

    c.execute("SELECT t.*,p.name as plan_name FROM tenants t LEFT JOIN plans p ON t.plan_id=p.id WHERE t.id=%s",(tid,))
    result=row(c.fetchone())
    conn.commit(); conn.close()
    if result: result["admin_email"]=admin_email
    return jsonify(result or {"id":tid,"name":d["name"],"admin_email":admin_email}),201

@app.route("/api/platform/tenants/<tid>", methods=["PUT"])
@platform_admin_required
def update_tenant(tid):
    d=request.get_json(silent=True,force=True) or {}
    conn=get_db(); c=conn.cursor()
    fields={k:d[k] for k in ["name","email","phone","address","plan_id","status",
                               "max_users","max_companies","notes","billing_cycle"] if k in d}
    if "approved" in d and d["approved"]:
        fields["status"]="active"
        fields["approved_by"]=g.user_id
        fields["approved_at"]=datetime.utcnow().isoformat()

    # If plan changed, update max limits from new plan
    if "plan_id" in fields:
        c.execute("SELECT max_users,max_companies FROM plans WHERE id=%s", (fields["plan_id"],))
        pl=c.fetchone()
        if pl:
            if "max_users"     not in d: fields["max_users"]     = pl[0]
            if "max_companies" not in d: fields["max_companies"] = pl[1]

    if not fields: return jsonify({"error":"Nothing to update"}),400
    fields["updated_at"]=datetime.utcnow().isoformat()
    c.execute(f"UPDATE tenants SET {','.join(k+'=%s' for k in fields)} WHERE id=%s",
              list(fields.values())+[tid])
    # Disable/enable all users in tenant on status change
    if "status" in fields:
        active=1 if fields["status"]=="active" else 0
        c.execute("UPDATE users SET is_active=%s WHERE tenant_id=%s AND is_platform_admin=0",(active,tid))
    conn.commit()
    c.execute("SELECT t.*,p.name as plan_name FROM tenants t LEFT JOIN plans p ON t.plan_id=p.id WHERE t.id=%s",(tid,))
    result=row(c.fetchone()); conn.close(); return jsonify(result)

@app.route("/api/platform/tenants/<tid>", methods=["DELETE"])
@platform_admin_required
def delete_tenant(tid):
    conn=get_db(); c=conn.cursor()
    # Soft delete — mark cancelled and disable users
    c.execute("UPDATE tenants SET status='cancelled',updated_at=NOW() WHERE id=%s",(tid,))
    c.execute("UPDATE users SET is_active=0 WHERE tenant_id=%s AND is_platform_admin=0",(tid,))
    conn.commit(); conn.close()
    return jsonify({"success":True})

@app.route("/api/platform/tenants/<tid>/users")
@platform_admin_required
def tenant_users(tid):
    conn=get_db(); c=conn.cursor()
    c.execute("""SELECT id,name,email,role,is_active,created_at,last_login
                 FROM users WHERE tenant_id=%s ORDER BY role,name""", (tid,))
    return jsonify(rows(c.fetchall()))

@app.route("/api/platform/tenants/<tid>/stats")
@platform_admin_required
def tenant_stats(tid):
    conn=get_db(); c=conn.cursor()
    stats={}
    for tbl,col in [("users","id"),("companies","id"),("tasks","id"),("documents","id"),("alerts","id")]:
        c.execute(f"SELECT COUNT(*) FROM {tbl} WHERE tenant_id=%s", (tid,))
        stats[tbl]=c.fetchone()[0]
    conn.close(); return jsonify(stats)

@app.route("/api/platform/summary")
@platform_admin_required
def platform_summary():
    conn=get_db(); c=conn.cursor()
    c.execute("SELECT COUNT(*) FROM tenants WHERE status='active'");   active=c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tenants WHERE status='pending'");  pending=c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tenants WHERE status='suspended'");suspended=c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tenants");                         total=c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM users WHERE is_platform_admin=0 AND is_active=1"); total_users=c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM companies"); total_co=c.fetchone()[0]
    c.execute("SELECT COALESCE(SUM(p.price_monthly),0) FROM tenants t JOIN plans p ON t.plan_id=p.id WHERE t.status='active' AND t.billing_cycle='monthly'")
    mrr=c.fetchone()[0]
    conn.close()
    return jsonify({"active":active,"pending":pending,"suspended":suspended,"total":total,
                    "total_users":total_users,"total_companies":total_co,"mrr":mrr})

# ══ SCHEDULER ════════════════════════════════════════════════════════════════
@app.route("/api/scheduler/run", methods=["POST"])
@require_role("superadmin","manager")
def run_scheduler():
    results=run_compliance_checks(); return jsonify({"success":True,"results":results})

# ══ BOOT ═════════════════════════════════════════════════════════════════════
def ensure_columns():
    """Add any columns that may be missing from older databases."""
    conn = get_db(); c = conn.cursor()
    migrations = [
        ("companies", "gstin",   "TEXT"),
        ("companies", "tan",     "TEXT"),
        ("companies", "website", "TEXT"),
        ("directors", "mca_notes", "TEXT"),
        ("shareholders", "mca_notes", "TEXT"),
    ]
    for table, col, coltype in migrations:
        try:
            c.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}")
            conn.commit()
        except Exception:
            pass  # Column already exists
    conn.close()

if __name__ == "__main__":
    init_db(); ensure_columns(); ensure_custom_placeholders_table(); ensure_permission_tables(); run_compliance_checks()
    print("\n" + "="*62)
    print("  TAXLY-CMS v2.0 - Corporate Compliance Management System")
    print("="*62)
    print("  URL    : http://127.0.0.1:5000")
    print("  Admin  : admin@compli.in / admin123")
    print("  Manager: manager@compli.in / manager123")
    print("  Staff  : staff@compli.in / staff123")
    print("="*62 + "\n")
    app.run(host="127.0.0.1",port=5000,debug=False)
