#!/usr/bin/env python3
"""
Run this on the server to find EXACT errors:
  cd /path/to/backend
  python3 diagnose.py
"""
import os, sys, traceback, re

print("="*60)
print("SERVER DIAGNOSTIC — Checking exact issues")
print("="*60)

# ── 1. Check which files are actually loaded ──────────────────
print("\n[1] FILE VERSIONS")
for fname in ['app.py','database.py','auth.py','compliance.py']:
    if os.path.exists(fname):
        with open(fname) as f: content = f.read()
        lines = content.count('\n')
        
        # Check for ? placeholders in VALUES
        q_in_values = len(re.findall(r'VALUES\s*\([^)]*\?[^)]*\)', content, re.DOTALL))
        
        # Check for fetchone()[0] 
        int_access = len(re.findall(r'fetchone\(\)\[0\]', content))
        
        # Check zip bug
        zip_bugs = len(re.findall(r'dict\(zip\(\[d\[0\] for d in c\.description\],\s*r\)\)', content))
        
        print(f"\n  {fname}: {lines} lines")
        print(f"    ? in VALUES:    {q_in_values} {'← NEEDS FIX' if q_in_values else '✓'}")
        print(f"    fetchone()[0]:  {int_access} {'← NEEDS FIX' if int_access else '✓'}")
        print(f"    zip(desc,r) bug:{zip_bugs} {'← NEEDS FIX' if zip_bugs else '✓'}")
        
        # Find specific lines with ?
        if q_in_values > 0:
            print(f"    Lines with ? in VALUES:")
            for i, line in enumerate(content.split('\n')):
                if 'VALUES' in line and '?' in line:
                    print(f"      L{i+1}: {line.strip()[:70]}")
    else:
        print(f"  {fname}: NOT FOUND in current directory!")

# ── 2. Live API test ──────────────────────────────────────────
print("\n[2] LIVE API TEST")
try:
    os.environ.pop('DATABASE_URL', None)
    sys.path.insert(0, '.')
    from database import init_db, USE_POSTGRES
    print(f"  DB mode: {'PostgreSQL' if USE_POSTGRES else 'SQLite'}")
    
    from app import app
    
    with app.test_client() as c:
        r = c.post('/api/auth/login', 
                   json={'email':'admin@compli.in','password':'admin123'},
                   content_type='application/json')
        if r.status_code != 200:
            print(f"  Login failed: {r.status_code} {r.get_json()}")
            sys.exit(1)
        
        token = r.get_json()['token']
        sa = {'Authorization': f'Bearer {token}'}
        cos = c.get('/api/companies', headers=sa).get_json()
        cid = cos[0]['id'] if cos else None
        
        tests = [
            ('GET /api/my-dashboard',           '/api/my-dashboard'),
            ('GET /api/dashboard',              '/api/dashboard'),
            ('GET /api/tasks/rolewise-summary', '/api/tasks/rolewise-summary'),
            ('GET /api/auditors',               '/api/auditors'),
            ('GET /api/dsc',                    '/api/dsc'),
        ]
        
        for label, url in tests:
            try:
                resp = c.get(url, headers=sa)
                d = resp.get_json()
                if resp.status_code == 200:
                    print(f"  ✓ {label}: 200")
                    # Check data integrity
                    items = d if isinstance(d, list) else []
                    for key, val in (d.items() if isinstance(d,dict) else []):
                        if isinstance(val, list) and val and isinstance(val[0],dict):
                            items = val; break
                    if items and items[0]:
                        first = items[0]
                        bad = [(k,v) for k,v in first.items() if v==k and isinstance(v,str)]
                        if bad:
                            print(f"    ← DATA BUG (column=value): {bad[:3]}")
                else:
                    print(f"  ✗ {label}: {resp.status_code}")
                    print(f"    {str(d)[:200]}")
            except Exception as e:
                print(f"  ✗ {label}: {e}")
                traceback.print_exc()

        # Test POST operations
        if cid:
            for label, url, body in [
                ('POST auditor', '/api/auditors', {
                    'company_id':cid,'name':'Diag CA','firm_name':'Diag Co',
                    'membership_no':'999111','appointment_type':'AGM Appointment',
                    'start_date':'2025-09-30','end_date':'2026-09-29'}),
                ('POST dsc', '/api/dsc', {
                    'company_id':cid,'holder_name':'Diag Holder',
                    'dsc_class':'Class 3','valid_from':'2025-01-01',
                    'valid_to':'2028-01-01','custody_status':'With Us'}),
            ]:
                try:
                    resp = c.post(url, json=body, headers=sa,
                                  content_type='application/json')
                    if resp.status_code in [200,201]:
                        print(f"  ✓ {label}: {resp.status_code}")
                    else:
                        print(f"  ✗ {label}: {resp.status_code}")
                        print(f"    {str(resp.get_json())[:200]}")
                except Exception as e:
                    print(f"  ✗ {label}: {e}")
                    traceback.print_exc()

except Exception as e:
    print(f"  ERROR: {e}")
    traceback.print_exc()

print("\n" + "="*60)
print("DONE — share this output to identify exact fixes needed")
print("="*60)
