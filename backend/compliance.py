"""compliance.py — Alert scheduler, Company Master PDF, Document Engine  v2.0"""
import uuid, json, re, io
from datetime import date, timedelta
from pathlib import Path
from database import get_db, rows, row

BASE_DIR = Path(__file__).parent

# ══════════════════════════════════════════════════════════════════════════════
#  COMPLIANCE CHECKER (runs daily / on-demand)
# ══════════════════════════════════════════════════════════════════════════════

def _count(c):
    r = c.fetchone()
    if r is None: return 0
    if isinstance(r, dict): return int(list(r.values())[0] or 0)
    return int(r[0] or 0)


def run_compliance_checks():
    today = date.today()
    results = {"auditor_expiry":0,"din_kyc":0,"dsc_expiry":0,"meeting_pending":0}
    conn = get_db(); c = conn.cursor()

    # ── Auditor expiry ───────────────────────────────────────────────────────
    c.execute("SELECT a.*,co.name as company_name FROM auditors a JOIN companies co ON a.company_id=co.id WHERE a.is_active=1 AND a.end_date IS NOT NULL")
    for aud in rows(c.fetchall()):
        end = _pd(aud["end_date"])
        if not end: continue
        days = (end-today).days
        if days < 0: sev,msg = "critical",f"EXPIRED {abs(days)}d ago"
        elif days <= 7: sev,msg = "critical",f"Expires in {days}d"
        elif days <= 30: sev,msg = "high",f"Expires in {days}d"
        else: continue
        _upsert(c,aud["company_id"],"auditor",aud["id"],"expiry",
                f"Auditor Expiring — {aud['name']}",
                f"{aud.get('firm_name',aud['name'])}'s appointment {msg}. File ADT-1.",
                aud["end_date"],sev)
        results["auditor_expiry"]+=1

    # ── DIN KYC ─────────────────────────────────────────────────────────────
    c.execute("""SELECT d.*,k.next_due_date,k.kyc_status,co.name as company_name
                 FROM directors d JOIN director_kyc k ON d.id=k.director_id
                 JOIN companies co ON d.company_id=co.id WHERE d.is_active=1""")
    for dr in rows(c.fetchall()):
        due = _pd(dr["next_due_date"])
        if not due: continue
        days = (due-today).days
        if days < 0: sev,msg = "critical",f"OVERDUE {abs(days)}d"
        elif days <= 7: sev,msg = "critical",f"Due in {days}d"
        elif days <= 30: sev,msg = "high",f"Due in {days}d"
        else: continue
        _upsert(c,dr["company_id"],"director",dr["id"],"kyc_due",
                f"DIR-3 KYC Due — {dr['name']}",
                f"DIN KYC for {dr['name']} (DIN:{dr.get('din','?')}) {msg}. Avoid DIN deactivation.",
                dr["next_due_date"],sev)
        c.execute("UPDATE director_kyc SET kyc_status=%s WHERE director_id=%s",
                  ("overdue" if days<0 else "pending",dr["id"]))
        results["din_kyc"]+=1

    # ── DSC expiry ──────────────────────────────────────────────────────────
    c.execute("""SELECT d.*,co.name as company_name FROM dsc_records d
                 LEFT JOIN companies co ON d.company_id=co.id WHERE d.is_active=1""")
    for dsc in rows(c.fetchall()):
        end = _pd(dsc["valid_to"])
        if not end: continue
        days = (end-today).days
        if days < 0: sev,msg = "critical",f"EXPIRED {abs(days)}d ago"
        elif days <= 7: sev,msg = "critical",f"Expires in {days}d"
        elif days <= 30: sev,msg = "high",f"Expires in {days}d"
        else: continue
        _upsert(c,dsc["company_id"],"dsc",dsc["id"],"dsc_expiry",
                f"DSC Expiring — {dsc['holder_name']}",
                f"DSC of {dsc['holder_name']} ({dsc['dsc_class']}) {msg}. Renew immediately.",
                dsc["valid_to"],sev)
        results["dsc_expiry"]+=1

    # ── Meeting minutes pending ──────────────────────────────────────────────
    c.execute("""SELECT m.*,co.name as company_name FROM meetings m
                 JOIN companies co ON m.company_id=co.id
                 WHERE m.status='scheduled' AND m.meeting_date < %s""",(today.isoformat(),))
    for mtg in rows(c.fetchall()):
        _upsert(c,mtg["company_id"],"meeting",mtg["id"],"minutes_pending",
                f"Minutes Pending — {mtg['meeting_type']} {mtg['meeting_no'] or ''}",
                f"{mtg['meeting_type']} on {mtg['meeting_date']} — minutes not drafted.",
                mtg["meeting_date"],"medium")
        results["meeting_pending"]+=1

    conn.commit(); conn.close()
    return results

def _pd(s):
    """Parse date from string OR datetime.date/datetime.datetime object."""
    if s is None: return None
    if isinstance(s, date): return s  # Already a date object (PostgreSQL)
    try: return date.fromisoformat(str(s)[:10])
    except: return None

def _upsert(c,company_id,entity_type,entity_id,alert_type,title,message,due_date,severity):
    c.execute("SELECT id FROM alerts WHERE entity_type=%s AND entity_id=%s AND alert_type=%s AND status='active'",
              (entity_type,entity_id,alert_type))
    ex = c.fetchone()
    if ex:
        c.execute("UPDATE alerts SET severity=%s,message=%s,due_date=%s WHERE id=%s",(severity,message,due_date,ex['id'] if isinstance(ex,dict) else ex[0]))
    else:
        c.execute("INSERT INTO alerts (id,company_id,entity_type,entity_id,alert_type,title,message,due_date,severity) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                  (str(uuid.uuid4()),company_id,entity_type,entity_id,alert_type,title,message,due_date,severity))


# ══════════════════════════════════════════════════════════════════════════════
#  DOCUMENT ENGINE — fill template with context
# ══════════════════════════════════════════════════════════════════════════════
def generate_document(template_body: str, context: dict) -> str:
    """Render template — replace all {{key}} with context values."""
    out = template_body
    for k, v in context.items():
        out = out.replace("{{"+k+"}}", str(v or ""))
    # Remaining unfilled placeholders — leave visible so user knows what's missing
    out = re.sub(r'\{\{([^}]+)\}\}', r'[\1 — TO BE FILLED]', out)
    return out


def extract_placeholders(template_body: str) -> list:
    """Extract all unique {{placeholder}} names from a template."""
    return sorted(set(re.findall(r'\{\{([^}]+)\}\}', template_body)))


# Placeholders that are auto-filled from company master — no manual input needed
AUTO_FILLED = {
    # Company
    "company_name","cin","pan","tan","gstin","incorporation_date","registered_office",
    "company_email","phone","business_activity","authorized_capital","paid_up_capital",
    "roc","company_type","financial_year_end","financial_year","place",
    # Date auto-fills
    "resolution_date","notice_date","letter_date","signing_date","year_ended_on",
    "current_year","previous_year","previous_year_end",
    # Director auto-fills
    "directors_list","directors_present",
    "director_name","designation","din","director_address","director_pan",
    "director_mobile","director_email","date_of_appointment",
    "director2_name","din2","designation2","director2_address",
    # Auditor auto-fills
    "auditor_name","auditor_firm_name","frn","srn_adt1","auditor_details",
    "auditor_membership","auditor_start_date","auditor_end_date",
    "nature_of_appointment","signing_partner_name","auditor_eligibility",
    # Shareholder auto-fills
    "shareholder1_name","share1_no","share1_nv","share1_total","share1_pct",
    "shareholder2_name","share2_no","share2_nv","share2_total","share2_pct",
    "shareholder3_name","share3_no","share3_nv","share3_total","share3_pct",
    "total_shares","total_paid_up",
}

# Placeholders that are fulfilled by choosing a specific active person
PERSON_FIELDS = {
    "director_name":"director","designation":"director","din":"director",
    "director_mobile":"director","director_email":"director","director_pan":"director",
    "director2_name":"director2","din2":"director2","designation2":"director2",
    "auditor_name":"auditor","auditor_firm_name":"auditor","frn":"auditor",
    "srn_adt1":"auditor","auditor_pan":"auditor","auditor_email":"auditor",
    "shareholder_name":"shareholder","shareholder_folio":"shareholder",
    "shareholder_shares":"shareholder","shareholder_pan":"shareholder",
}


def get_active_entities(company_id: str) -> dict:
    """
    Return ONLY active/current directors, auditors, shareholders for a company.
    Retired directors (is_active=0 or date_of_cessation set), resigned auditors
    (is_active=0), and removed shareholders (is_active=0) are excluded.
    """
    conn = get_db(); c = conn.cursor()

    # Active directors only: is_active=1 AND no cessation date (or cessation in future)
    c.execute("""
        SELECT d.id,d.name,d.designation,d.din,d.pan,d.email,d.mobile,d.address,
               k.kyc_status,k.next_due_date
        FROM directors d
        LEFT JOIN director_kyc k ON d.id=k.director_id
        WHERE d.company_id=%s
          AND d.is_active=1
          AND (d.date_of_cessation IS NULL OR d.date_of_cessation > CURRENT_DATE)
        ORDER BY d.designation, d.name
    """, (company_id,))
    directors = rows(c.fetchall())

    # Current auditor only: is_active=1 AND appointment not expired
    c.execute("""
        SELECT id,name,firm_name,membership_no,frn,pan,email,
               nature_of_appointment,appointment_type,start_date,end_date,srn_adt1
        FROM auditors
        WHERE company_id=%s
          AND is_active=1
          AND (end_date IS NULL OR end_date >= CURRENT_DATE)
        ORDER BY start_date DESC
        LIMIT 5
    """, (company_id,))
    auditors = rows(c.fetchall())

    # Active shareholders only: is_active=1
    c.execute("""
        SELECT id,name,folio_no,pan,shares_held,share_class,email,mobile,face_value
        FROM shareholders
        WHERE company_id=%s AND is_active=1
        ORDER BY name
    """, (company_id,))
    shareholders = rows(c.fetchall())

    conn.close()
    return {"directors": directors, "auditors": auditors, "shareholders": shareholders}


def build_context(company_id: str, extra: dict = None,
                  director_id: str = None, director2_id: str = None,
                  auditor_id: str = None) -> dict:
    """
    Build document context from CRM data.
    Only active directors/auditors/shareholders are selected.
    Specific person IDs can be passed to pick exactly who signs the document.
    """
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT * FROM companies WHERE id=%s", (company_id,))
    co = row(c.fetchone())
    if not co: conn.close(); raise ValueError("Company not found")

    entities = get_active_entities(company_id)
    dirs  = entities["directors"]
    auds  = entities["auditors"]
    shs   = entities["shareholders"]

    conn.close()

    # Base company context
    ctx = {
        "company_name":       co["name"],
        "cin":                co["cin"] or "",
        "pan":                co["pan"] or "",
        "tan":                co["tan"] or "",
        "gstin":              co.get("gstin","") or "",
        "incorporation_date": co["incorporation_date"] or "",
        "registered_office":  co["registered_office"] or "",
        "company_email":      co["email"] or "",
        "phone":              co["phone"] or "",
        "business_activity":  co["business_activity"] or "",
        "authorized_capital": f"{co['authorized_capital']:,.0f}" if co["authorized_capital"] else "0",
        "paid_up_capital":    f"{co['paid_up_capital']:,.0f}" if co["paid_up_capital"] else "0",
        "roc":                co.get("roc","") or "",
        "company_type":       co.get("company_type","Private Limited"),
        "resolution_date":    date.today().strftime("%d %B %Y"),
        "notice_date":        date.today().strftime("%d %B %Y"),
        "letter_date":        date.today().strftime("%d %B %Y"),
        "place":              (co["registered_office"] or "").split(",")[-1].strip() or "Mumbai",
        "financial_year":     str(date.today().year),
        "directors_list":     "\n".join([f"   {i+1}. {d['name']}, {d['designation']} (DIN: {d.get('din','—')})" for i,d in enumerate(dirs)]),
        "directors_present":  "\n".join([f"   {d['name']} — {d['designation']}" for d in dirs]),
    }

    # Primary signatory director (use selected ID, else first active director)
    primary = None
    if director_id:
        primary = next((d for d in dirs if d["id"] == director_id), None)
    if not primary and dirs:
        primary = dirs[0]
    if primary:
        ctx.update({
            "director_name":   primary["name"],
            "designation":     primary["designation"],
            "din":             primary.get("din",""),
            "director_pan":    primary.get("pan",""),
            "director_mobile": primary.get("mobile",""),
            "director_email":  primary.get("email",""),
            "director_address":primary.get("address",""),
        })

    # Second signatory director
    secondary = None
    if director2_id:
        secondary = next((d for d in dirs if d["id"] == director2_id), None)
    if not secondary:
        others = [d for d in dirs if not primary or d["id"] != primary["id"]]
        if others: secondary = others[0]
    if secondary:
        ctx.update({
            "director2_name":    secondary["name"],
            "din2":              secondary.get("din",""),
            "designation2":      secondary.get("designation",""),
            "director2_address": secondary.get("address",""),
        })

    # Auditor (use selected ID, else current active auditor)
    aud = None
    if auditor_id:
        aud = next((a for a in auds if a["id"] == auditor_id), None)
    if not aud and auds:
        aud = auds[0]
    if aud:
        ctx.update({
            "auditor_name":       aud["name"],
            "auditor_firm_name":  aud["firm_name"] or aud["name"],
            "frn":                aud["frn"] or "",
            "srn_adt1":           aud["srn_adt1"] or "",
            "auditor_pan":        aud.get("pan",""),
            "auditor_email":      aud.get("email",""),
            "auditor_membership": aud.get("membership_no",""),
            "auditor_details":    f"{aud['firm_name'] or aud['name']} (FRN: {aud.get('frn','')}) — {aud.get('nature_of_appointment','Statutory Auditor')}",
            "nature_of_appointment": aud.get("nature_of_appointment",""),
            "appointment_type":   aud.get("appointment_type",""),
            "auditor_start_date": aud.get("start_date",""),
            "auditor_end_date":   aud.get("end_date",""),
        })

    # Financial year context
    from datetime import date as _date
    today    = _date.today()
    fy_start = today.year if today.month >= 4 else today.year - 1
    fy_end   = fy_start + 1
    ctx.update({
        "financial_year":     f"{fy_start}-{str(fy_end)[-2:]}",
        "financial_year_end": f"31 March {fy_end}",
        "year_ended_on":      f"31 March {fy_end}",
        "current_year":       f"01/04/{fy_start} to 31/03/{fy_end}",
        "previous_year":      f"01/04/{fy_start-1} to 31/03/{fy_start}",
        "previous_year_end":  f"31 March {fy_start}",
        "signing_date":       today.strftime("%d %B %Y"),
        "notice_date":        today.strftime("%d %B %Y"),
        "letter_date":        today.strftime("%d %B %Y"),
        "resolution_date":    today.strftime("%d %B %Y"),
        "date_of_appointment": primary.get("date_of_appointment","") if primary else "",
    })

    # Auditor extended fields
    if aud:
        ctx["signing_partner_name"] = extra.get("signing_partner_name","") if extra else ""
        ctx["auditor_eligibility"]  = aud.get("end_date","") or ""

    # Shareholders (auto-fill up to 3 shareholders)
    for i, sh in enumerate(shs[:3], 1):
        total = (sh.get("shares_held") or 0) * (sh.get("face_value") or 10)
        pct   = round((sh.get("shares_held",0) / max(sum(s.get("shares_held",0) for s in shs),1)) * 100, 2) if shs else 0
        ctx.update({
            f"shareholder{i}_name": sh.get("name",""),
            f"share{i}_no":         str(sh.get("shares_held",0)),
            f"share{i}_nv":         str(sh.get("face_value",10)),
            f"share{i}_total":      f"{total:,.0f}",
            f"share{i}_pct":        str(pct),
        })
    total_shs  = sum(s.get("shares_held",0) for s in shs)
    total_paid = sum((s.get("shares_held",0))*(s.get("face_value",10)) for s in shs)
    ctx["total_shares"]    = str(total_shs)
    ctx["total_paid_up"]   = f"{total_paid:,.0f}"

    # Apply extra context (manual overrides from user) — these win over auto values
    if extra:
        ctx.update({k: v for k, v in extra.items() if v not in (None, "")})

    return ctx


# ══════════════════════════════════════════════════════════════════════════════
#  COMPANY MASTER PDF
# ══════════════════════════════════════════════════════════════════════════════
def generate_company_master_pdf(company_id: str) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import inch, cm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                     TableStyle, HRFlowable, PageBreak)
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

    conn = get_db(); c = conn.cursor()
    c.execute("SELECT * FROM companies WHERE id=%s", (company_id,))
    co = row(c.fetchone())
    if not co: conn.close(); raise ValueError("Company not found")
    c.execute("""SELECT d.*,k.last_kyc_date,k.next_due_date,k.kyc_status
                 FROM directors d LEFT JOIN director_kyc k ON d.id=k.director_id
                 WHERE d.company_id=%s ORDER BY d.date_of_appointment""", (company_id,))
    dirs = rows(c.fetchall())
    c.execute("SELECT * FROM auditors WHERE company_id=%s ORDER BY created_at DESC", (company_id,))
    auds = rows(c.fetchall())
    c.execute("SELECT * FROM shareholders WHERE company_id=%s AND is_active=1 ORDER BY folio_no", (company_id,))
    shs = rows(c.fetchall())
    c.execute("SELECT * FROM dsc_records WHERE company_id=%s AND is_active=1", (company_id,))
    dscs = rows(c.fetchall())
    conn.close()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=0.7*inch, bottomMargin=0.7*inch,
                            leftMargin=0.85*inch, rightMargin=0.85*inch)

    # ── Colour palette ───────────────────────────────────────────
    NAVY = colors.HexColor("#0f2d5c")
    BLUE = colors.HexColor("#1a56db")
    LIGHT = colors.HexColor("#e8f0fe")
    GREY = colors.HexColor("#64748b")
    ROW_ALT = colors.HexColor("#f8fafc")
    WHITE = colors.white

    def sty(name,size=9,bold=False,color=None,align=TA_LEFT,leading=None):
        return ParagraphStyle(name,fontName="Helvetica-Bold" if bold else "Helvetica",
                               fontSize=size,textColor=color or colors.HexColor("#1e293b"),
                               alignment=align,leading=leading or (size*1.4))

    story = []

    # ── LETTERHEAD ───────────────────────────────────────────────
    lh_addr = co.get("letterhead_address") or co.get("registered_office") or ""
    lh_footer = co.get("letterhead_footer") or f"CIN: {co.get('cin','')} | PAN: {co.get('pan','')} | Email: {co.get('email','')}"

    story.append(Paragraph(co["name"].upper(), sty("h1",18,True,NAVY,TA_CENTER,24)))
    story.append(Paragraph(lh_addr, sty("addr",8,False,GREY,TA_CENTER)))
    story.append(Spacer(1,4))
    story.append(HRFlowable(width="100%",thickness=2,color=BLUE,spaceAfter=2))
    story.append(HRFlowable(width="100%",thickness=0.5,color=NAVY,spaceAfter=8))
    story.append(Paragraph("COMPANY MASTER DATA", sty("title",13,True,BLUE,TA_CENTER)))
    story.append(Paragraph(f"Generated on {date.today().strftime('%d %B %Y')}", sty("dt",8,False,GREY,TA_CENTER)))
    story.append(Spacer(1,14))

    def section_heading(text):
        story.append(Spacer(1,10))
        story.append(Table([[Paragraph(text, sty("sh",10,True,WHITE))]],
                           colWidths=["100%"],
                           style=TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY),
                                             ("ROWPADDING",(0,0),(-1,-1),6),
                                             ("BOX",(0,0),(-1,-1),0,WHITE)])))
        story.append(Spacer(1,6))

    def kv_table(pairs, cols=2):
        """pairs = [(label,value), ...]"""
        cell_sty = TableStyle([
            ("FONTNAME",(0,0),(-1,-1),"Helvetica"),("FONTSIZE",(0,0),(-1,-1),8.5),
            ("ROWPADDING",(0,0),(-1,-1),5),
            ("BACKGROUND",(0,0),(0,-1),LIGHT),
            ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
            ("TEXTCOLOR",(0,0),(0,-1),NAVY),
            ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#cbd5e1")),
        ])
        if cols == 2:
            # Pair up into rows
            tdata = []
            for i in range(0,len(pairs),2):
                r = list(pairs[i])
                if i+1 < len(pairs): r += list(pairs[i+1])
                else: r += ["",""]
                tdata.append([Paragraph(str(r[0]),sty("kl",8.5,True,NAVY)),
                               Paragraph(str(r[1]),sty("kv",8.5)),
                               Paragraph(str(r[2]),sty("kl",8.5,True,NAVY)),
                               Paragraph(str(r[3]),sty("kv",8.5))])
            w = doc.width
            t = Table(tdata, colWidths=[w*0.18,w*0.32,w*0.18,w*0.32])
        else:
            tdata = [[Paragraph(str(p[0]),sty("kl",8.5,True,NAVY)),
                      Paragraph(str(p[1]),sty("kv",8.5))] for p in pairs]
            t = Table(tdata, colWidths=[doc.width*0.25,doc.width*0.75])
        t.setStyle(cell_sty)
        story.append(t)

    def data_table(headers, data_rows, col_widths=None):
        tdata = [[Paragraph(h, sty("th",8,True,WHITE)) for h in headers]]
        for i,dr in enumerate(data_rows):
            row_data = [Paragraph(str(v or "—"), sty("td",8)) for v in dr]
            tdata.append(row_data)
        w = doc.width
        cw = col_widths or [w/len(headers)]*len(headers)
        t = Table(tdata, colWidths=cw)
        ts = TableStyle([
            ("BACKGROUND",(0,0),(-1,0),BLUE),("TEXTCOLOR",(0,0),(-1,0),WHITE),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),
            ("ROWPADDING",(0,0),(-1,-1),5),
            ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#cbd5e1")),
        ])
        for i in range(1,len(tdata)):
            if i%2==0: ts.add("BACKGROUND",(0,i),(-1,i),ROW_ALT)
        t.setStyle(ts)
        story.append(t)

    # ── 1. Company Details ───────────────────────────────────────
    section_heading("1. COMPANY DETAILS")
    kv_table([
        ("Company Name", co["name"]), ("CIN", co.get("cin") or "—"),
        ("Company Type", co.get("company_type") or "—"), ("Status", (co.get("status") or "active").upper()),
        ("Incorporation Date", co.get("incorporation_date") or "—"), ("ROC", co.get("roc") or "—"),
        ("PAN", co.get("pan") or "—"), ("TAN", co.get("tan") or "—"),
        ("Email", co.get("email") or "—"), ("Phone", co.get("phone") or "—"),
        ("Authorised Capital", f"Rs. {co.get('authorized_capital',0):,.0f}"), ("Paid-up Capital", f"Rs. {co.get('paid_up_capital',0):,.0f}"),
        ("Business Activity", co.get("business_activity") or "—"), ("Registered Office", co.get("registered_office") or "—"),
    ])

    # ── 2. Directors ─────────────────────────────────────────────
    section_heading("2. DIRECTORS")
    if dirs:
        data_table(
            ["Name","Designation","DIN","PAN","Email","Mobile","Appointed","KYC Status","KYC Due"],
            [(d["name"],d["designation"],d.get("din",""),d.get("pan",""),
              d.get("email",""),d.get("mobile",""),d.get("date_of_appointment",""),
              (d.get("kyc_status") or "pending").upper(),d.get("next_due_date","")) for d in dirs],
            col_widths=[doc.width*w for w in [0.14,0.12,0.09,0.10,0.16,0.10,0.10,0.10,0.09]]
        )
    else:
        story.append(Paragraph("No directors recorded.", sty("empty",9,False,GREY)))

    # ── 3. Statutory Auditors ────────────────────────────────────
    section_heading("3. STATUTORY AUDITORS")
    if auds:
        data_table(
            ["Auditor/Firm","Membership","FRN","Nature","Start","End","ADT-1 SRN","Status"],
            [(a.get("firm_name") or a["name"],a.get("membership_no",""),a.get("frn",""),
              a.get("nature_of_appointment",""),a.get("start_date",""),a.get("end_date",""),
              a.get("srn_adt1",""),"Active" if a.get("is_active") else "Inactive") for a in auds],
            col_widths=[doc.width*w for w in [0.18,0.10,0.10,0.14,0.10,0.10,0.14,0.08]]
        )
    else:
        story.append(Paragraph("No auditor records.", sty("empty",9,False,GREY)))

    # ── 4. Shareholders ──────────────────────────────────────────
    section_heading("4. SHAREHOLDERS (Register of Members)")
    if shs:
        total_shares = sum(s.get("shares_held",0) for s in shs)
        data_table(
            ["Name","Folio","PAN","Share Class","Shares","% Holding","Face Value","Email","Mobile"],
            [(s["name"],s.get("folio_no",""),s.get("pan",""),s.get("share_class","Equity"),
              f"{s.get('shares_held',0):,}",
              f"{(s.get('shares_held',0)/total_shares*100):.1f}%" if total_shares else "0%",
              f"Rs.{s.get('face_value',10)}",s.get("email",""),s.get("mobile","")) for s in shs],
            col_widths=[doc.width*w for w in [0.14,0.08,0.09,0.09,0.09,0.07,0.09,0.16,0.10]]
        )
        story.append(Spacer(1,4))
        story.append(Paragraph(f"Total Shares: {total_shares:,}", sty("total",9,True,NAVY)))
    else:
        story.append(Paragraph("No shareholders recorded.", sty("empty",9,False,GREY)))

    # ── 5. DSC Records ───────────────────────────────────────────
    section_heading("5. DSC RECORDS")
    if dscs:
        today = date.today()
        def dsc_status(valid_to):
            if not valid_to: return "Unknown"
            d = _pd(valid_to)
            if d is None: return "Unknown"
            diff = (d-today).days
            if diff < 0: return f"EXPIRED {abs(diff)}d ago"
            if diff <= 30: return f"Expiring in {diff}d"
            return f"Valid ({diff}d)"
        data_table(
            ["Holder","Type","Class","Issued By","Valid From","Valid To","Token","Custody","Status"],
            [(d["holder_name"],d.get("holder_type",""),d.get("dsc_class",""),d.get("issued_by",""),
              d.get("valid_from",""),d.get("valid_to",""),d.get("token_type",""),
              d.get("custody_status",""),dsc_status(d.get("valid_to"))) for d in dscs],
            col_widths=[doc.width*w for w in [0.14,0.09,0.08,0.10,0.10,0.10,0.10,0.10,0.13]]
        )
    else:
        story.append(Paragraph("No DSC records.", sty("empty",9,False,GREY)))

    # ── Footer ───────────────────────────────────────────────────
    story.append(Spacer(1,18))
    story.append(HRFlowable(width="100%",thickness=0.5,color=colors.HexColor("#94a3b8")))
    story.append(Spacer(1,4))
    story.append(Paragraph(lh_footer, sty("footer",7.5,False,GREY,TA_CENTER)))
    story.append(Paragraph("CONFIDENTIAL — Generated by Taxly-CMS | Taxly India Private Limited", sty("conf",7,False,GREY,TA_CENTER)))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
#  REGISTER PDF GENERATOR  (with company letterhead)
# ══════════════════════════════════════════════════════════════════════════════
REGISTER_DEFINITIONS = {
    "MGT-1": {
        "name": "Register of Members",
        "section": "Section 88(1)(a) / Rule 3",
        "add_route": "shareholders",
        "columns": ["Folio No","Name","PAN","Address","Share Class","Shares Held","Face Value (Rs.)","Date of Entry","Mobile","Email"],
        "query": "SELECT folio_no,name,pan,address,share_class,shares_held,face_value,date_of_entry,mobile,email FROM shareholders WHERE company_id=%s AND is_active=1 ORDER BY folio_no",
    },
    "MGT-2": {
        "name": "Register of Debenture Holders",
        "section": "Section 88(1)(b)",
        "add_route": "shareholders",
        "columns": ["Folio No","Name","PAN","Debenture Type","Amount (Rs.)","Face Value","Date of Issue","Mobile","Email"],
        "query": "SELECT folio_no,name,pan,share_class,shares_held,face_value,date_of_entry,mobile,email FROM shareholders WHERE company_id=%s AND share_class='Debenture' ORDER BY folio_no",
    },
    "DIR-3": {
        "name": "Register of Directors and KMP",
        "section": "Section 170 / Rule 17",
        "add_route": "directors",
        "columns": ["Name","DIN","PAN","Designation","Date Appointed","Date Ceased","Address","Email","Mobile","KYC Status","KYC Due Date"],
        "query": """SELECT d.name,d.din,d.pan,d.designation,d.date_of_appointment,d.date_of_cessation,
                           d.address,d.email,d.mobile,k.kyc_status,k.next_due_date
                    FROM directors d LEFT JOIN director_kyc k ON d.id=k.director_id
                    WHERE d.company_id=%s ORDER BY d.date_of_appointment""",
    },
    "MBP-1": {
        "name": "Register of Director Interests",
        "section": "Section 189 / Rule 16",
        "add_route": "director_interests",
        "columns": ["Director Name","DIN","Entity Name","Entity Type","Nature of Interest","Date of Disclosure","Board Resolution Date","Remarks"],
        "query": "SELECT director_name,din,entity_name,entity_type,nature_of_interest,date_of_disclosure,date_of_board_resolution,remarks FROM director_interests WHERE company_id=%s ORDER BY date_of_disclosure",
    },
    "ADT-1": {
        "name": "Register of Auditors",
        "section": "Section 139 / Rule 5",
        "add_route": "auditors",
        "columns": ["Auditor/Firm","Membership No","FRN","PAN","Email","Nature of Appointment","Appt Type","Start Date","End Date","ADT-1 SRN"],
        "query": "SELECT firm_name,membership_no,frn,pan,email,nature_of_appointment,appointment_type,start_date,end_date,srn_adt1 FROM auditors WHERE company_id=%s ORDER BY start_date",
    },
    "CHG-1": {
        "name": "Register of Charges",
        "section": "Section 85 / Rule 3",
        "add_route": "charges",
        "columns": ["Charge ID","Type","Charge Holder","Assets Charged","Amount (Rs.)","Date Created","Date Modified","Date Satisfied","Status","Remarks"],
        "query": "SELECT charge_id,charge_type,charge_holder,assets_charged,amount,date_of_creation,date_of_modification,date_of_satisfaction,status,remarks FROM charges WHERE company_id=%s ORDER BY date_of_creation",
    },
    "SH-6": {
        "name": "Register of ESOPs",
        "section": "Section 62(1)(b) / Rule 12",
        "add_route": "esop_grants",
        "columns": ["Employee Name","Designation","Employee ID","Grant Date","Options Granted","Exercise Price","Vesting Date","Vesting Period","Options Exercised","Options Lapsed","Status"],
        "query": "SELECT employee_name,designation,employee_id,grant_date,options_granted,exercise_price,vesting_date,vesting_period,options_exercised,options_lapsed,status FROM esop_grants WHERE company_id=%s ORDER BY grant_date",
    },
    "MBP-3": {
        "name": "Register of Investments",
        "section": "Section 186 / Rule 12",
        "add_route": "investments",
        "columns": ["Investee Name","Type","Investment Type","Amount (Rs.)","Date of Investment","Board Resolution Date","SRN (MGT-14)","Purpose","Remarks"],
        "query": "SELECT investee_name,investee_type,investment_type,amount,date_of_investment,board_resolution_date,srn_mgb4,purpose,remarks FROM investments WHERE company_id=%s ORDER BY date_of_investment",
    },
    "MBP-2": {
        "name": "Register of Loans & Guarantees",
        "section": "Section 186 / Rule 11",
        "add_route": "loans_guarantees",
        "columns": ["Party Name","Party Type","Transaction Type","Amount (Rs.)","Date","Rate of Interest (%)","Repayment Date","Security","Board Resolution","Outstanding (Rs.)","Status"],
        "query": "SELECT party_name,party_type,transaction_type,amount,date_of_transaction,rate_of_interest,repayment_date,security,board_resolution_date,outstanding_amount,status FROM loans_guarantees WHERE company_id=%s ORDER BY date_of_transaction",
    },
    "RPT-188": {
        "name": "Register of Related Party Transactions",
        "section": "Section 188 / Rule 15",
        "add_route": "related_party_transactions",
        "columns": ["Party Name","Relationship","Nature of Transaction","Amount (Rs.)","Date","Board Approval","Shareholder Approval","Terms","Remarks"],
        "query": "SELECT party_name,relationship,nature_of_transaction,amount,date_of_transaction,date_of_board_approval,date_of_shareholders_approval,terms,remarks FROM related_party_transactions WHERE company_id=%s ORDER BY date_of_transaction",
    },
    "DSC": {
        "name": "DSC Custody Register",
        "section": "Internal Record",
        "add_route": "dsc",
        "columns": ["Holder Name","Type","Class","Issued By","Valid From","Valid To","Token Type","Custody Status","Custody Date","Notes"],
        "query": "SELECT holder_name,holder_type,dsc_class,issued_by,valid_from,valid_to,token_type,custody_status,custody_date,notes FROM dsc_records WHERE company_id=%s AND is_active=1",
    },
}


def generate_register_pdf(company_id: str, register_type: str) -> bytes:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    from reportlab.lib.enums import TA_CENTER, TA_LEFT

    reg = REGISTER_DEFINITIONS.get(register_type)
    if not reg: raise ValueError(f"Unknown register: {register_type}")

    conn = get_db(); c = conn.cursor()
    c.execute("SELECT * FROM companies WHERE id=%s", (company_id,))
    co = row(c.fetchone())
    if not co: conn.close(); raise ValueError("Company not found")
    c.execute(reg["query"], (company_id,))
    data_rows = c.fetchall()
    conn.close()

    NAVY = colors.HexColor("#0f2d5c"); BLUE = colors.HexColor("#1a56db")
    GREY = colors.HexColor("#64748b"); WHITE = colors.white
    ROW_ALT = colors.HexColor("#f8fafc")

    use_landscape = len(reg["columns"]) > 6
    pagesize = landscape(A4) if use_landscape else A4
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=pagesize,
                            topMargin=0.6*inch, bottomMargin=0.6*inch,
                            leftMargin=0.7*inch, rightMargin=0.7*inch)
    story = []

    def sty(n,sz=9,bold=False,col=None,align=TA_LEFT):
        return ParagraphStyle(n,fontName="Helvetica-Bold" if bold else "Helvetica",
                               fontSize=sz,textColor=col or colors.HexColor("#1e293b"),alignment=align)

    lh_addr = co.get("letterhead_address") or co.get("registered_office") or ""
    story.append(Paragraph(co["name"].upper(), sty("co",16,True,NAVY,TA_CENTER)))
    story.append(Paragraph(lh_addr, sty("addr",8,False,GREY,TA_CENTER)))
    story.append(Spacer(1,3))
    story.append(HRFlowable(width="100%",thickness=2,color=BLUE,spaceAfter=2))
    story.append(HRFlowable(width="100%",thickness=0.5,color=NAVY,spaceAfter=6))
    story.append(Paragraph(f"{reg['name'].upper()}", sty("t",12,True,BLUE,TA_CENTER)))
    story.append(Paragraph(f"As required under {reg['section']} of the Companies Act, 2013", sty("s",8,False,GREY,TA_CENTER)))
    story.append(Paragraph(f"CIN: {co.get('cin','')}  |  As on {date.today().strftime('%d %B %Y')}", sty("dt",8,False,GREY,TA_CENTER)))
    story.append(Spacer(1,12))

    # Build table
    page_w = doc.width
    n_cols = len(reg["columns"])
    col_w = [page_w/n_cols]*n_cols

    tdata = [[Paragraph(h, sty("h",7.5,True,WHITE)) for h in reg["columns"]]]
    for i, dr in enumerate(data_rows):
        # Support both dict rows (PostgreSQL RealDictCursor) and tuple rows (SQLite)
        row_vals = list(dr.values()) if isinstance(dr, dict) else list(dr)
        tdata.append([Paragraph(str(v or "—") if v is not None else "—", sty("d",7.5)) for v in row_vals])

    if not data_rows:
        tdata.append([Paragraph("No entries recorded.", sty("d",8,False,GREY))]+[""]*(n_cols-1))

    ts = TableStyle([
        ("BACKGROUND",(0,0),(-1,0),BLUE),("TEXTCOLOR",(0,0),(-1,0),WHITE),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),7.5),
        ("ROWPADDING",(0,0),(-1,-1),4),
        ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#cbd5e1")),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
    ])
    for i in range(1,len(tdata)):
        if i%2==0: ts.add("BACKGROUND",(0,i),(-1,i),ROW_ALT)
    t = Table(tdata, colWidths=col_w)
    t.setStyle(ts)
    story.append(t)

    story.append(Spacer(1,16))
    story.append(HRFlowable(width="100%",thickness=0.5,color=colors.HexColor("#94a3b8")))
    story.append(Spacer(1,3))
    lh_footer = co.get("letterhead_footer") or f"CIN: {co.get('cin','')} | PAN: {co.get('pan','')} | {co.get('registered_office','')}"
    story.append(Paragraph(lh_footer, sty("ft",7,False,GREY,TA_CENTER)))
    story.append(Paragraph("CONFIDENTIAL — Generated by Taxly-CMS | Taxly India Private Limited", sty("cf",6.5,False,GREY,TA_CENTER)))

    doc.build(story)
    buf.seek(0)
    return buf.read()
