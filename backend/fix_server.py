#!/usr/bin/env python3
"""
Self-fixing patcher — run on server in backend/ directory:
  cd /path/to/backend
  python3 fix_server.py
  
This script patches app.py, auth.py, database.py in-place.
"""
import re, os, sys, ast, shutil
from datetime import datetime

BACKUP_SUFFIX = f".bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

def backup(fname):
    shutil.copy(fname, fname + BACKUP_SUFFIX)
    print(f"  Backed up → {fname}{BACKUP_SUFFIX}")

def save(fname, content):
    try:
        ast.parse(content)
    except SyntaxError as e:
        print(f"  ✗ SYNTAX ERROR in {fname} L{e.lineno}: {e.msg} — NOT SAVED")
        return False
    with open(fname, 'w') as f:
        f.write(content)
    print(f"  ✓ {fname} saved")
    return True

total_fixes = 0

# ══════════════════════════════════════════════════════════════
# FIX app.py
# ══════════════════════════════════════════════════════════════
print("\n[app.py]")
if not os.path.exists('app.py'):
    print("  app.py not found — run from backend/ directory")
    sys.exit(1)

backup('app.py')
with open('app.py') as f:
    app = f.read()

# ── Fix A: ALL ? in execute() blocks → %s ──────────────────────────────────
def replace_in_execute_blocks(content):
    """Replace ? with %s inside every c.execute() call."""
    count = 0
    
    # Triple-quoted SQL strings
    def fix_triple(m):
        nonlocal count
        original = m.group(0)
        fixed = original.replace('?', '%s')
        if fixed != original:
            count += fixed.count('%s') - original.count('%s')
        return fixed
    
    content = re.sub(r'c\.execute\s*\(\s*""".*?"""', fix_triple, content, flags=re.DOTALL)
    
    # Single-quoted SQL strings with ? 
    def fix_single(m):
        nonlocal count
        original = m.group(0)
        fixed = original.replace('?', '%s')
        if fixed != original:
            count += 1
        return fixed
    
    content = re.sub(r'c\.execute\s*\(\s*"[^"]*\?[^"]*"', fix_single, content)
    
    # VALUES lines with mixed ? and %s
    lines = content.split('\n')
    new_lines = []
    for line in lines:
        if '?' in line and '%s' in line and ('VALUES' in line or 'SET ' in line):
            new_line = line.replace('?', '%s')
            if new_line != line:
                count += 1
            new_lines.append(new_line)
        else:
            new_lines.append(line)
    content = '\n'.join(new_lines)
    
    return content, count

app, fix_count = replace_in_execute_blocks(app)
total_fixes += fix_count
print(f"  Fix A: ? → %s in SQL: {fix_count} replacements")

# ── Fix B: fetchone()[0] → _count(c) ──────────────────────────────────────
count_b = app.count('= c.fetchone()[0]') + app.count('=c.fetchone()[0]')
app = re.sub(r'=\s*c\.fetchone\(\)\[0\]', '= _count(c)', app)
total_fixes += count_b
print(f"  Fix B: fetchone()[0] → _count(c): {count_b} fixes")

# ── Fix C: Add _count and _rv helpers if missing ───────────────────────────
if '_count' not in app or 'def _count' not in app:
    HELPERS = '''
def _count(c):
    """Extract COUNT(*) from dict or tuple row."""
    r = c.fetchone()
    if r is None: return 0
    if isinstance(r, dict): return int(list(r.values())[0] or 0)
    return int(r[0] or 0)

def _rv(r, idx_or_key):
    """Safe row value — works for dict rows and tuple rows."""
    if isinstance(r, dict):
        if isinstance(idx_or_key, int):
            return list(r.values())[idx_or_key]
        return r.get(idx_or_key)
    return r[idx_or_key]

'''
    # Insert before first @app.route
    idx = app.find('@app.route(')
    app = app[:idx] + HELPERS + app[idx:]
    total_fixes += 1
    print(f"  Fix C: _count/_rv helpers added")
else:
    print(f"  Fix C: helpers already present ✓")

# ── Fix D: zip(description, r) without isinstance guard ────────────────────
pattern = r'dict\(zip\(\[d\[0\] for d in c\.description\],\s*r\)\)'
count_d = len(re.findall(pattern, app))
app = re.sub(
    pattern,
    '(dict(r) if isinstance(r,dict) else dict(zip([d[0] for d in c.description],r)))',
    app
)
total_fixes += count_d
print(f"  Fix D: zip(description,r) unguarded: {count_d} fixes")

# ── Fix E: zip(cols, r) unguarded in list comprehensions ───────────────────
count_e = 0
for var in ['cols', 'sub_cols', 'person_cols']:
    pattern_e = rf'\[dict\(zip\({var},\s*r\)\) for r in c\.fetchall\(\)\]'
    count_e += len(re.findall(pattern_e, app))
    app = re.sub(
        pattern_e,
        f'[dict(r) if isinstance(r,dict) else dict(zip({var},r)) for r in c.fetchall()]',
        app
    )
total_fixes += count_e
print(f"  Fix E: zip(cols,r) unguarded: {count_e} fixes")

# ── Fix F: my_tasks zip bug ─────────────────────────────────────────────────
OLD_TASKS_ZIP = 'my_tasks = [dict(zip([d[0] for d in c.description], r)) for r in c.fetchall()]'
if OLD_TASKS_ZIP in app:
    app = app.replace(OLD_TASKS_ZIP, 'my_tasks = rows(c.fetchall())')
    total_fixes += 1
    print(f"  Fix F: my_tasks zip bug fixed")

# ── Fix G: role_summary zip bug ─────────────────────────────────────────────
OLD_RS_ZIP = '        role_summary = [dict(zip([d[0] for d in c.description], r)) for r in c.fetchall()]'
if OLD_RS_ZIP in app:
    app = app.replace(OLD_RS_ZIP, '        role_summary = rows(c.fetchall())')
    total_fixes += 1
    print(f"  Fix G: role_summary zip bug fixed")

# ── Fix H: task_counts dict comprehension ──────────────────────────────────
OLD_TC = '    task_counts = {r[0]: r[1] for r in c.fetchall()}'
NEW_TC = '''    task_counts = {}
    for _r in c.fetchall():
        if isinstance(_r, dict):
            _vals = list(_r.values()); task_counts[_vals[0]] = _vals[1]
        else:
            task_counts[_r[0]] = _r[1]'''
if OLD_TC in app:
    app = app.replace(OLD_TC, NEW_TC, 1)
    total_fixes += 1
    print(f"  Fix H: task_counts dict comprehension fixed")

# ── Fix I: DSC update old[0] ────────────────────────────────────────────────
OLD_DSC = 'old[0] if old else ""'
if OLD_DSC in app:
    app = app.replace(OLD_DSC, "(old.get('custody_status') if isinstance(old,dict) else old[0]) if old else \"\"", 1)
    total_fixes += 1
    print(f"  Fix I: DSC old[0] fixed")

# ── Fix J: tpl[0/1/2] doc template fetch ───────────────────────────────────
OLD_TPL = '    body   = tpl[0] or ""\n    tname  = tpl[1]\n    tcat   = tpl[2]'
NEW_TPL = ('    body   = (tpl[\'template_body\'] if isinstance(tpl,dict) else tpl[0]) or ""\n'
           "    tname  = tpl['name']     if isinstance(tpl,dict) else tpl[1]\n"
           "    tcat   = tpl['category'] if isinstance(tpl,dict) else tpl[2]")
if OLD_TPL in app:
    app = app.replace(OLD_TPL, NEW_TPL, 1)
    total_fixes += 1
    print(f"  Fix J: tpl[0/1/2] fixed")

# ── Fix K: meetings zip bug ─────────────────────────────────────────────────
OLD_MTG = '    meetings = [dict(zip([d[0] for d in c.description], row)) for row in c.fetchall()]'
if OLD_MTG in app:
    app = app.replace(OLD_MTG, '    meetings = rows(c.fetchall())', 1)
    total_fixes += 1
    print(f"  Fix K: meetings zip bug fixed")

# ── Fix L: password change r[0] ─────────────────────────────────────────────
OLD_PW = "    if not r or r[0]!=hash_pw(old):"
if OLD_PW in app:
    NEW_PW = "    _rpw = r['password'] if isinstance(r,dict) else r[0]\n    if not r or _rpw!=hash_pw(old):"
    app = app.replace(OLD_PW, NEW_PW, 1)
    total_fixes += 1
    print(f"  Fix L: password check r[0] fixed")

# ── Fix M: SQLite datetime functions (shouldn't be in app.py) ───────────────
count_m = 0
for old, new in [("datetime('now')", "NOW()"), ("date('now')", "CURRENT_DATE")]:
    c2 = app.count(old)
    app = app.replace(old, new)
    count_m += c2
if count_m:
    total_fixes += count_m
    print(f"  Fix M: SQLite date functions → PG: {count_m} fixes")

save('app.py', app)

# ══════════════════════════════════════════════════════════════
# FIX auth.py
# ══════════════════════════════════════════════════════════════
print("\n[auth.py]")
if os.path.exists('auth.py'):
    backup('auth.py')
    with open('auth.py') as f: auth = f.read()
    
    # Fix ? in execute calls
    auth, cnt = replace_in_execute_blocks(auth)
    total_fixes += cnt
    if cnt: print(f"  ? → %s: {cnt}")
    
    # Fix permissions dict comprehension
    OLD_AUTH = "        overrides = {(r[0], r[1]): bool(r[2]) for r in c.fetchall()}"
    if OLD_AUTH in auth:
        NEW_AUTH = """        overrides = {}
        for _r in c.fetchall():
            if isinstance(_r, dict):
                _v = list(_r.values())
                overrides[(_v[0], _v[1])] = bool(_v[2])
            else:
                overrides[(_r[0], _r[1])] = bool(_r[2])"""
        auth = auth.replace(OLD_AUTH, NEW_AUTH, 1)
        total_fixes += 1
        print(f"  Permissions dict comprehension fixed")
    
    save('auth.py', auth)

# ══════════════════════════════════════════════════════════════
# FIX database.py
# ══════════════════════════════════════════════════════════════
print("\n[database.py]")
if os.path.exists('database.py'):
    backup('database.py')
    with open('database.py') as f: db = f.read()
    
    # Ensure SQLite timeout and check_same_thread
    if 'sqlite3.connect(str(DB_PATH))' in db and 'timeout' not in db:
        db = db.replace(
            'sqlite3.connect(str(DB_PATH))',
            'sqlite3.connect(str(DB_PATH), timeout=30, check_same_thread=False)'
        )
        total_fixes += 1
        print(f"  SQLite timeout added")
    
    # Ensure description defaults to []
    OLD_DESC = "        self.description = self._cur.description"
    NEW_DESC = "        self.description = self._cur.description or []"
    if OLD_DESC in db and 'or []' not in db:
        db = db.replace(OLD_DESC, NEW_DESC, 1)
        total_fixes += 1
        print(f"  _DictCursor.description defaults to []")
    
    save('database.py', db)

# ══════════════════════════════════════════════════════════════
# FIX compliance.py
# ══════════════════════════════════════════════════════════════
print("\n[compliance.py]")
if os.path.exists('compliance.py'):
    backup('compliance.py')
    with open('compliance.py') as f: comp = f.read()
    
    comp_new, cnt = replace_in_execute_blocks(comp)
    total_fixes += cnt
    if cnt: print(f"  ? → %s: {cnt}")
    
    # Add _count helper if missing
    if '_count' not in comp_new and 'def _count' not in comp_new:
        COMP_HELPER = '\ndef _count(c):\n    r=c.fetchone()\n    if r is None: return 0\n    if isinstance(r,dict): return int(list(r.values())[0] or 0)\n    return int(r[0] or 0)\n\n'
        idx = comp_new.find('\ndef ')
        comp_new = comp_new[:idx] + COMP_HELPER + comp_new[idx:]
        total_fixes += 1
        print(f"  _count helper added")
    
    comp_new = re.sub(r'=\s*c\.fetchone\(\)\[0\]', '= _count(c)', comp_new)
    
    save('compliance.py', comp_new)

# ══════════════════════════════════════════════════════════════
# FINAL VERIFICATION
# ══════════════════════════════════════════════════════════════
print("\n[VERIFICATION]")
for fname in ['app.py', 'auth.py', 'database.py', 'compliance.py']:
    if not os.path.exists(fname): continue
    with open(fname) as f: content = f.read()
    
    # Count remaining issues
    q_issues = sum(1 for m in re.finditer(r'\.execute\s*\((.*?)\)', content, re.DOTALL)
                   if '?' in m.group(1) and '%s' not in m.group(1) and 'lambda' not in m.group(1))
    int_issues = len(re.findall(r'fetchone\(\)\[0\]', content))
    zip_issues = len(re.findall(r'dict\(zip\(\[d\[0\] for d in c\.description\],\s*r\)\)', content))
    
    all_ok = q_issues == 0 and int_issues == 0 and zip_issues == 0
    print(f"  {'✓' if all_ok else '✗'} {fname}: ? in SQL={q_issues}, fetchone()[0]={int_issues}, zip bug={zip_issues}")

print(f"\nTotal fixes applied: {total_fixes}")
print("\nRestart gunicorn to apply changes:")
print("  pkill gunicorn && gunicorn app:app -c gunicorn.conf.py --daemon")
print("  # or: sudo systemctl restart taxlycms")
