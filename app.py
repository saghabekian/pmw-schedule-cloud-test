import os, sqlite3, html, urllib.parse, subprocess, platform
from datetime import datetime
from functools import wraps
from flask import Flask, request, redirect, url_for, session, render_template_string, flash, jsonify, send_from_directory
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
try:
    import psycopg
    from psycopg.rows import dict_row
except Exception:
    psycopg = None
    dict_row = None


APP_NAME = "PMW Ticket + Fabrication"
APP_VERSION = "v30 Clear Ticket Cells"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "pmw_schedule.db")
UPLOAD_FOLDER = os.path.join(APP_DIR, "uploads")
EXPORT_FOLDER = os.path.join(APP_DIR, "exports")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)
CLOUD_TICKET_FOLDER = os.path.join(APP_DIR, 'cloud_ticket_files')
os.makedirs(CLOUD_TICKET_FOLDER, exist_ok=True)

DISPLAY_COLS = [1,2,3,4,5,6]
ROLE_LEVEL = {"viewer":1,"editor":2,"admin":3}


def auto_upgrade_sqlite_schema():
    """Safe startup upgrade for existing SQLite DBs on Render/free hosting.
    Adds newer columns that older DB files may be missing.
    """
    try:
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()

        cur.execute("""CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            password_hash TEXT,
            role TEXT,
            active INTEGER DEFAULT 1,
            created_at TEXT
        )""")

        cur.execute("""CREATE TABLE IF NOT EXISTS workbook_cells(
            sheet_name TEXT,
            row_num INTEGER,
            col_num INTEGER,
            value TEXT DEFAULT '',
            updated_by TEXT DEFAULT '',
            updated_at TEXT DEFAULT '',
            PRIMARY KEY(sheet_name,row_num,col_num)
        )""")

        for coldef in [
            "bg_color TEXT DEFAULT ''",
            "text_color TEXT DEFAULT ''",
            "link_path TEXT DEFAULT ''",
            "link_label TEXT DEFAULT ''",
            "font_size TEXT DEFAULT ''",
            "bold TEXT DEFAULT ''",
            "rich_html TEXT DEFAULT ''"
        ]:
            try:
                cur.execute("ALTER TABLE workbook_cells ADD COLUMN " + coldef)
            except sqlite3.OperationalError:
                pass

        cur.execute("""CREATE TABLE IF NOT EXISTS ticket_links(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_number TEXT,
            subject TEXT,
            sender TEXT,
            received TEXT,
            file_path TEXT UNIQUE,
            created_at TEXT
        )""")

        cur.execute("""CREATE TABLE IF NOT EXISTS audit_log(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            action TEXT,
            details TEXT,
            created_at TEXT
        )""")

        con.commit()
        con.close()
        print("SQLite schema auto-upgrade complete:", DB_PATH)
    except Exception as e:
        print("SQLite schema auto-upgrade failed:", repr(e))

auto_upgrade_sqlite_schema()


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "pmw-local-dev-secret")

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = "postgresql://" + DATABASE_URL[len("postgres://"):]
USE_POSTGRES = bool(DATABASE_URL and psycopg)

class PgCursor:
    def __init__(self, cur):
        self.cur = cur

    def _convert_query(self, query):
        q = query.strip()
        upper = q.upper()

        # SQLite upsert compatibility for workbook cells
        if upper.startswith("INSERT OR REPLACE INTO WORKBOOK_CELLS"):
            q = q.replace("INSERT OR REPLACE INTO workbook_cells", "INSERT INTO workbook_cells")
            if "ON CONFLICT" not in q:
                q += """ ON CONFLICT(sheet_name,row_num,col_num) DO UPDATE SET
                    value=EXCLUDED.value,
                    bg_color=EXCLUDED.bg_color,
                    text_color=EXCLUDED.text_color,
                    link_path=EXCLUDED.link_path,
                    link_label=EXCLUDED.link_label,
                    font_size=EXCLUDED.font_size,
                    bold=EXCLUDED.bold,
                    rich_html=EXCLUDED.rich_html,
                    updated_by=EXCLUDED.updated_by,
                    updated_at=EXCLUDED.updated_at"""
        elif upper.startswith("INSERT OR REPLACE INTO"):
            q = q.replace("INSERT OR REPLACE INTO", "INSERT INTO")

        return q.replace("?", "%s")

    def execute(self, query, params=()):
        self.cur.execute(self._convert_query(query), params or ())
        return self

    def fetchone(self):
        return self.cur.fetchone()

    def fetchall(self):
        return self.cur.fetchall()

    @property
    def lastrowid(self):
        try:
            self.cur.execute("SELECT LASTVAL() AS id")
            row = self.cur.fetchone()
            return row["id"] if row else None
        except Exception:
            return None

class PgConnection:
    def __init__(self, con):
        self.con = con

    def cursor(self):
        return PgCursor(self.con.cursor(row_factory=dict_row))

    def execute(self, query, params=()):
        cur = self.cursor()
        return cur.execute(query, params)

    def commit(self):
        self.con.commit()

    def close(self):
        self.con.close()


def db():
    if USE_POSTGRES:
        con = psycopg.connect(DATABASE_URL)
        return PgConnection(con)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    con=db(); cur=con.cursor()

    if USE_POSTGRES:
        cur.execute("""CREATE TABLE IF NOT EXISTS users(
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE,
            password_hash TEXT,
            role TEXT,
            active INTEGER DEFAULT 1,
            created_at TEXT
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS workbook_cells(
            sheet_name TEXT,
            row_num INTEGER,
            col_num INTEGER,
            value TEXT DEFAULT '',
            updated_by TEXT DEFAULT '',
            updated_at TEXT DEFAULT '',
            PRIMARY KEY(sheet_name,row_num,col_num)
        )""")
        for coldef in [
            "bg_color TEXT DEFAULT ''",
            "text_color TEXT DEFAULT ''",
            "link_path TEXT DEFAULT ''",
            "link_label TEXT DEFAULT ''",
            "font_size TEXT DEFAULT ''",
            "bold TEXT DEFAULT ''",
            "rich_html TEXT DEFAULT ''"
        ]:
            try:
                cur.execute("ALTER TABLE workbook_cells ADD COLUMN " + coldef)
            except Exception:
                try:
                    con.con.rollback()
                except Exception:
                    pass

        cur.execute("""CREATE TABLE IF NOT EXISTS ticket_links(
            id SERIAL PRIMARY KEY,
            job_number TEXT,
            subject TEXT,
            sender TEXT,
            received TEXT,
            file_path TEXT UNIQUE,
            created_at TEXT
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS audit_log(
            id SERIAL PRIMARY KEY,
            username TEXT,
            action TEXT,
            details TEXT,
            created_at TEXT
        )""")
    else:
        cur.execute("""CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            password_hash TEXT,
            role TEXT,
            active INTEGER DEFAULT 1,
            created_at TEXT
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS workbook_cells(
            sheet_name TEXT,
            row_num INTEGER,
            col_num INTEGER,
            value TEXT DEFAULT '',
            updated_by TEXT DEFAULT '',
            updated_at TEXT DEFAULT '',
            PRIMARY KEY(sheet_name,row_num,col_num)
        )""")
        for coldef in [
            "bg_color TEXT DEFAULT ''",
            "text_color TEXT DEFAULT ''",
            "link_path TEXT DEFAULT ''",
            "link_label TEXT DEFAULT ''",
            "font_size TEXT DEFAULT ''",
            "bold TEXT DEFAULT ''",
            "rich_html TEXT DEFAULT ''"
        ]:
            try:
                cur.execute("ALTER TABLE workbook_cells ADD COLUMN " + coldef)
            except sqlite3.OperationalError:
                pass

        cur.execute("""CREATE TABLE IF NOT EXISTS ticket_links(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_number TEXT,
            subject TEXT,
            sender TEXT,
            received TEXT,
            file_path TEXT UNIQUE,
            created_at TEXT
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS audit_log(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            action TEXT,
            details TEXT,
            created_at TEXT
        )""")

    for u,p,r in [("admin","admin123","admin"),("shop","shop123","editor"),("viewer","view123","viewer")]:
        if not cur.execute("SELECT id FROM users WHERE username=?",(u,)).fetchone():
            cur.execute(
                "INSERT INTO users(username,password_hash,role,active,created_at) VALUES(?,?,?,?,?)",
                (u,generate_password_hash(p),r,1,datetime.now().isoformat(timespec='seconds'))
            )

    con.commit(); con.close()

def log(action, details=""):

    con=db(); con.execute("INSERT INTO audit_log(username,action,details,created_at) VALUES(?,?,?,?)",(session.get('username','system'),action,details,datetime.now().isoformat(timespec='seconds'))); con.commit(); con.close()

def can(role): return ROLE_LEVEL.get(session.get('role',''),0) >= ROLE_LEVEL[role]

def login_required(fn):
    @wraps(fn)
    def w(*a,**k):
        if not session.get('user_id'): return redirect(url_for('login'))
        return fn(*a,**k)
    return w

def role_required(role):
    def deco(fn):
        @wraps(fn)
        def w(*a,**k):
            if not can(role):
                flash("You do not have permission for that.")
                return redirect(url_for('index'))
            return fn(*a,**k)
        return w
    return deco

def clean(v):
    if v is None: return ""
    if hasattr(v, "strftime"): return v.strftime("%m/%d/%Y")
    return str(v)

def import_workbook(path):
    wb = load_workbook(path, data_only=True, keep_vba=True)
    con=db(); cur=con.cursor(); cur.execute("DELETE FROM workbook_cells")
    count=0
    for ws in wb.worksheets:
        if ws.title.lower().startswith('chart'): continue
        sheet = 'Fabrication Schedule' if ws.title == 'Sheet1' else ws.title
        for r in range(1, max(50, min(ws.max_row or 50, 90))+1):
            for c in DISPLAY_COLS:
                v = clean(ws.cell(r,c).value)
                cur.execute("INSERT OR REPLACE INTO workbook_cells(sheet_name,row_num,col_num,value,updated_at) VALUES(?,?,?,?,?)",(sheet,r,c,v,datetime.now().isoformat(timespec='seconds')))
                if v: count += 1
    con.commit(); con.close(); log("IMPORT_WORKBOOK", os.path.basename(path)); return count

def sheet_names():
    con=db()
    if USE_POSTGRES:
        rows=con.execute("""
            SELECT sheet_name
            FROM workbook_cells
            GROUP BY sheet_name
            ORDER BY MIN(CASE WHEN sheet_name='Fabrication Schedule' THEN 0 ELSE 1 END), sheet_name
        """).fetchall()
    else:
        rows=con.execute("SELECT DISTINCT sheet_name FROM workbook_cells ORDER BY CASE WHEN sheet_name='Fabrication Schedule' THEN 0 ELSE 1 END, sheet_name").fetchall()
    con.close()
    names=[r['sheet_name'] for r in rows]
    return names or ['Fabrication Schedule']

def cells_for(sheet):
    con=db(); rows=con.execute("SELECT row_num,col_num,value FROM workbook_cells WHERE sheet_name=?",(sheet,)).fetchall(); con.close()
    d={(r['row_num'],r['col_num']):r['value'] or '' for r in rows}
    for r in range(1,51):
        for c in DISPLAY_COLS: d.setdefault((r,c),"")
    return d

def cell_meta_for(sheet):
    con=db()
    rows=con.execute("SELECT row_num,col_num,value,bg_color,text_color,link_path,link_label,font_size,bold,rich_html FROM workbook_cells WHERE sheet_name=?",(sheet,)).fetchall()
    con.close()
    d={}
    for r in rows:
        d[(r['row_num'],r['col_num'])]={
            "value": r["value"] or "",
            "bg_color": r["bg_color"] or "",
            "text_color": r["text_color"] or "",
            "link_path": r["link_path"] or "",
            "link_label": r["link_label"] or "",
            "font_size": r["font_size"] or "",
            "bold": r["bold"] or "",
            "rich_html": r["rich_html"] or ""
        }
    for rr in range(1,51):
        for cc in DISPLAY_COLS:
            d.setdefault((rr,cc), {"value":"","bg_color":"","text_color":"","link_path":"","link_label":"","font_size":"","bold":"","rich_html":""})
    return d

def ticket_rows_from_drop(path):
    """Reads PMW Ticket emails drop.xlsx.
    Columns:
    B received, C sender, D subject, E job #, G saved .msg path
    """
    wb = load_workbook(path, data_only=False)
    ws = wb.active
    rows=[]
    for r in range(2, (ws.max_row or 1)+1):
        received = clean(ws.cell(r,2).value)
        sender = clean(ws.cell(r,3).value)
        subject = clean(ws.cell(r,4).value)
        job = clean(ws.cell(r,5).value).strip()
        file_path = clean(ws.cell(r,7).value).strip()
        if not file_path:
            h = ws.cell(r,6).hyperlink
            if h:
                file_path = h.target or ""
        if job or subject or file_path:
            rows.append({
                "source_row": r,
                "received": received,
                "sender": sender,
                "subject": subject,
                "job_number": job,
                "file_path": file_path
            })
    return rows

def import_one_ticket_from_drop(path, mode="newest"):
    rows = ticket_rows_from_drop(path)
    if not rows:
        return None, "No ticket rows found."

    con=db(); cur=con.cursor()
    imported = None

    # newest means last non-duplicate row in the workbook.
    candidates = list(reversed(rows)) if mode == "newest" else rows
    for t in candidates:
        fp = (t.get("file_path") or "").strip()
        if not fp:
            continue
        exists = cur.execute("SELECT id FROM ticket_links WHERE file_path=?",(fp,)).fetchone()
        if exists:
            continue
        now=datetime.now().isoformat(timespec='seconds')
        cur.execute("""INSERT INTO ticket_links(job_number,subject,sender,received,file_path,created_at)
                       VALUES(?,?,?,?,?,?)""",
                    (t.get("job_number",""), t.get("subject",""), t.get("sender",""), t.get("received",""), fp, now))
        tid = cur.lastrowid
        con.commit()
        imported = cur.execute("SELECT * FROM ticket_links WHERE id=?",(tid,)).fetchone()
        break

    con.close()
    if imported:
        log("IMPORT_ONE_TICKET", imported["file_path"] or "")
        return imported, ""
    return None, "No new ticket found. The newest rows may already be imported."


def import_last_n_tickets_from_drop(path, count=1):
    """Import the newest N non-duplicate ticket rows from the ticket drop workbook."""
    try:
        count = int(count)
    except Exception:
        count = 1
    count = max(1, min(count, 25))

    rows = ticket_rows_from_drop(path)
    if not rows:
        return [], "No ticket rows found."

    con=db(); cur=con.cursor()
    imported=[]
    now=datetime.now().isoformat(timespec='seconds')

    for t in reversed(rows):
        if len(imported) >= count:
            break
        fp = (t.get("file_path") or "").strip()
        if not fp:
            continue
        exists = cur.execute("SELECT id FROM ticket_links WHERE file_path=?",(fp,)).fetchone()
        if exists:
            continue
        cur.execute("""INSERT INTO ticket_links(job_number,subject,sender,received,file_path,created_at)
                       VALUES(?,?,?,?,?,?)""",
                    (t.get("job_number",""), t.get("subject",""), t.get("sender",""), t.get("received",""), fp, now))
        tid = cur.lastrowid
        imported.append(cur.execute("SELECT * FROM ticket_links WHERE id=?",(tid,)).fetchone())

    con.commit(); con.close()
    if imported:
        log("IMPORT_LAST_TICKETS", f"{len(imported)} ticket(s)")
        return imported, ""
    return [], "No new tickets found. The newest rows may already be imported."

def save_selected_ticket_rows_from_drop(path, selected_rows):
    """Import only checked row numbers from the ticket drop workbook."""
    wanted=set()
    for x in selected_rows:
        try:
            wanted.add(int(x))
        except Exception:
            pass

    rows=ticket_rows_from_drop(path)
    if not wanted:
        return [], "No ticket rows were selected."

    con=db(); cur=con.cursor()
    imported=[]
    skipped=0
    now=datetime.now().isoformat(timespec='seconds')

    for t in rows:
        if int(t.get("source_row", 0)) not in wanted:
            continue
        fp=(t.get("file_path") or "").strip()
        if not fp:
            skipped += 1
            continue
        exists=cur.execute("SELECT id FROM ticket_links WHERE file_path=?",(fp,)).fetchone()
        if exists:
            skipped += 1
            continue
        cur.execute("""INSERT INTO ticket_links(job_number,subject,sender,received,file_path,created_at)
                       VALUES(?,?,?,?,?,?)""",
                    (t.get("job_number",""), t.get("subject",""), t.get("sender",""), t.get("received",""), fp, now))
        tid=cur.lastrowid
        imported.append(cur.execute("SELECT * FROM ticket_links WHERE id=?",(tid,)).fetchone())

    con.commit(); con.close()
    if imported:
        log("IMPORT_SELECTED_TICKETS", f"{len(imported)} ticket(s), skipped {skipped}")
        return imported, f"Imported {len(imported)} selected ticket(s)."
    return [], f"No new tickets imported. {skipped} selected row(s) may already be imported or missing a saved .msg path."

def first_empty_schedule_row(sheet, col_num):
    con=db()
    for r in range(3,51):
        row=con.execute("SELECT value FROM workbook_cells WHERE sheet_name=? AND row_num=? AND col_num=?",(sheet,r,col_num)).fetchone()
        if not row or not (row["value"] or "").strip():
            con.close()
            return r
    con.close()
    return 50

def add_ticket_to_schedule(ticket_id, side, sheet="Fabrication Schedule"):
    con=db(); cur=con.cursor()
    t=cur.execute("SELECT * FROM ticket_links WHERE id=?",(ticket_id,)).fetchone()
    if not t:
        con.close()
        return False, "Ticket not found."

    if side == "numbering":
        sort_col, job_col, note_col = 1,2,3
    else:
        sort_col, job_col, note_col = 4,5,6

    r = first_empty_schedule_row(sheet, job_col)
    text = (t["subject"] or t["job_number"] or "Ticket").strip()
    now=datetime.now().isoformat(timespec='seconds')
    cloud_file = ''
    try:
        cloud_file = t['cloud_file'] or ''
    except Exception:
        cloud_file = ''
    path = ("/ticket_download/" + str(ticket_id)) if cloud_file.strip() else (t["file_path"] or "")
    label = (t["subject"] or t["job_number"] or "Ticket")[:80]

    for c, val in [(job_col, text), (note_col, "TICKET")]:
        cur.execute("""INSERT OR REPLACE INTO workbook_cells(
            sheet_name,row_num,col_num,value,bg_color,text_color,link_path,link_label,font_size,bold,rich_html,updated_by,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (sheet,r,c,val,"#fff066" if c==job_col else "", "", path if c==job_col else "", label if c==job_col else "", "", "", "", session.get("username",""), now))
    con.commit(); con.close()
    log("ADD_TICKET_TO_SCHEDULE", f"{side} row {r}: {text}")
    return True, f"Ticket added to {side} row {r}."

def open_path_on_this_pc(path):
    path=(path or "").strip()
    if not path:
        return False, "No path saved."
    try:
        if platform.system().lower().startswith("win"):
            os.startfile(path)
        else:
            reveal_file(path)
        return True, "Opened."
    except Exception as e:
        return False, str(e)


def upsert_workbook_cell(sheet, r, c, val='', bg='', txt='', link='', label='', fsize='', bold='', rich='', user=''):
    now = datetime.now().isoformat(timespec='seconds')
    con = db()
    cur = con.cursor()
    if USE_POSTGRES:
        cur.execute("""
            INSERT INTO workbook_cells(
                sheet_name,row_num,col_num,value,bg_color,text_color,link_path,link_label,font_size,bold,rich_html,updated_by,updated_at
            )
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(sheet_name,row_num,col_num) DO UPDATE SET
                value=EXCLUDED.value,
                bg_color=EXCLUDED.bg_color,
                text_color=EXCLUDED.text_color,
                link_path=EXCLUDED.link_path,
                link_label=EXCLUDED.link_label,
                font_size=EXCLUDED.font_size,
                bold=EXCLUDED.bold,
                rich_html=EXCLUDED.rich_html,
                updated_by=EXCLUDED.updated_by,
                updated_at=EXCLUDED.updated_at
        """, (sheet,r,c,val,bg,txt,link,label,fsize,bold,rich,user,now))
    else:
        cur.execute("""INSERT OR REPLACE INTO workbook_cells(
            sheet_name,row_num,col_num,value,bg_color,text_color,link_path,link_label,font_size,bold,rich_html,updated_by,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",(sheet,r,c,val,bg,txt,link,label,fsize,bold,rich,user,now))
    con.commit()
    con.close()
    return now


def save_posted_cells(sheet):
    for r in range(3,51):
        for c in DISPLAY_COLS:
            key=f"cell_{r}_{c}"
            if key in request.form:
                val=(request.form.get(key) or '').replace('\r',' ').replace('\n',' ').strip()

                # V30: if the user deletes the visible cell text, clear the whole cell object,
                # including hidden ticket/email link, envelope icon, colors, font settings, and rich text.
                if val == '':
                    upsert_workbook_cell(sheet,r,c,'','','','','','','','',session.get('username',''))
                    continue

                bg=(request.form.get(f"bg_{r}_{c}") or '').strip()
                txt=(request.form.get(f"txt_{r}_{c}") or '').strip()
                link=(request.form.get(f"link_{r}_{c}") or '').strip()
                label=(request.form.get(f"label_{r}_{c}") or '').strip()
                fsize=(request.form.get(f"fsize_{r}_{c}") or '').strip()
                bold=(request.form.get(f"bold_{r}_{c}") or '').strip()
                rich=(request.form.get(f"rich_{r}_{c}") or '').strip()
                upsert_workbook_cell(sheet,r,c,val,bg,txt,link,label,fsize,bold,rich,session.get('username',''))


def sort_side(sheet, key_col, job_col, note_col):
    """Sort one side of the schedule while keeping the whole row metadata together:
    visible text, background color, text color, ticket/email link, and link label.
    """
    con=db(); cur=con.cursor(); now=datetime.now().isoformat(timespec='seconds')
    rows=[]; blanks=[]

    def sort_key(v):
        s=str(v or '').strip()
        try:
            return (0, float(s))
        except Exception:
            return (1, s.lower())

    def get_cell(r, c):
        row=cur.execute("""SELECT value,bg_color,text_color,link_path,link_label,font_size,bold,rich_html
                           FROM workbook_cells
                           WHERE sheet_name=? AND row_num=? AND col_num=?""",(sheet,r,c)).fetchone()
        if not row:
            return {"value":"","bg_color":"","text_color":"","link_path":"","link_label":"","font_size":"","bold":"","rich_html":""}
        return {
            "value": row["value"] or "",
            "bg_color": row["bg_color"] or "",
            "text_color": row["text_color"] or "",
            "link_path": row["link_path"] or "",
            "link_label": row["link_label"] or "",
            "font_size": row["font_size"] or "",
            "bold": row["bold"] or "",
            "rich_html": row["rich_html"] or ""
        }

    for r in range(3,51):
        row_obj=[get_cell(r, c) for c in (key_col, job_col, note_col)]
        if any(x["value"].strip() or x["bg_color"].strip() or x["text_color"].strip() or x["link_path"].strip() or x.get("font_size","").strip() or x.get("bold","").strip() or x.get("rich_html","").strip() for x in row_obj):
            if row_obj[0]["value"].strip():
                rows.append(row_obj)
            else:
                blanks.append(row_obj)

    ordered=sorted(rows, key=lambda x: sort_key(x[0]["value"])) + blanks

    for idx,r in enumerate(range(3,51)):
        row_obj=ordered[idx] if idx < len(ordered) else [
            {"value":"","bg_color":"","text_color":"","link_path":"","link_label":"","font_size":"","bold":"","rich_html":""},
            {"value":"","bg_color":"","text_color":"","link_path":"","link_label":"","font_size":"","bold":"","rich_html":""},
            {"value":"","bg_color":"","text_color":"","link_path":"","link_label":"","font_size":"","bold":"","rich_html":""}
        ]
        for c, cell in zip((key_col,job_col,note_col), row_obj):
            cur.execute("""INSERT OR REPLACE INTO workbook_cells(
                sheet_name,row_num,col_num,value,bg_color,text_color,link_path,link_label,font_size,bold,rich_html,updated_by,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (sheet,r,c,cell["value"],cell["bg_color"],cell["text_color"],cell["link_path"],cell["link_label"],cell.get("font_size",""),cell.get("bold",""),cell.get("rich_html",""),session.get('username',''),now))
    con.commit(); con.close()

def email_body(sheet):
    d=cells_for(sheet); lines=[]
    for r in range(3,51):
        n=d.get((r,1),'').strip(); j=d.get((r,2),'').strip(); f=d.get((r,4),'').strip(); fj=d.get((r,5),'').strip()
        if n or j: lines.append(f"Numbering {n}: {j}".strip())
        if f or fj: lines.append(f"Fabrication {f}: {fj}".strip())
    return "Please see today's schedule below.\r\n\r\n" + "\r\n".join(lines) + "\r\n\r\nOpen the PMW app for the live version."


def make_schedule_pdf(sheet):
    """Create a PDF copy of the active schedule and return the full file path.
    V16 keeps PMW cell colors and marks linked ticket/email cells with an envelope.
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    meta = cell_meta_for(sheet)
    d = cells_for(sheet)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_sheet = ''.join(ch if ch.isalnum() or ch in ('-', '_') else '_' for ch in sheet)[:40]
    filename = f"PMW_Schedule_{safe_sheet}_{stamp}.pdf"
    path = os.path.join(EXPORT_FOLDER, filename)

    doc = SimpleDocTemplate(path, pagesize=landscape(letter), rightMargin=0.35*inch, leftMargin=0.35*inch, topMargin=0.30*inch, bottomMargin=0.30*inch)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('PMWTitle', parent=styles['Heading1'], alignment=1, fontName='Helvetica-Bold', fontSize=17, leading=20)
    cell_style = ParagraphStyle('PMWCell', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=9)
    head_style = ParagraphStyle('PMWHead', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=11, alignment=1)

    title = d.get((1,2),'FABRICATION SCHEDULE') or 'FABRICATION SCHEDULE'
    datev = d.get((1,4),'') or datetime.now().strftime('%m/%d/%Y')
    story = [Paragraph(f"{html.escape(title)} &nbsp;&nbsp;&nbsp; {html.escape(datev)}", title_style), Spacer(1, 6)]

    data = [[Paragraph('NUMBER', head_style), Paragraph('NUMBERING', head_style), Paragraph('STATUS/NOTES', head_style), Paragraph('NUMBER', head_style), Paragraph('FABRICATION', head_style), Paragraph('STATUS/NOTES', head_style)]]
    row_source=[]
    for r in range(3, 51):
        vals = []
        has_content = False
        for c in DISPLAY_COLS:
            m = meta.get((r,c), {})
            txt = str(m.get("value",""))
            if m.get("link_path"):
                txt = "✉ " + txt
            if txt.strip() or m.get("bg_color") or m.get("text_color"):
                has_content = True
            vals.append(Paragraph(html.escape(txt), cell_style))
        if has_content:
            data.append(vals)
            row_source.append(r)
    if len(data) == 1:
        data.append(['','','','','',''])

    tbl = Table(data, colWidths=[0.55*inch, 3.30*inch, 0.85*inch, 0.55*inch, 3.30*inch, 0.85*inch], repeatRows=1)
    style_cmds=[
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#d9ead3')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),
        ('ALIGN', (2,0), (2,-1), 'CENTER'),
        ('ALIGN', (3,0), (3,-1), 'CENTER'),
        ('ALIGN', (5,0), (5,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ]

    for pdf_row, sheet_row in enumerate(row_source, start=1):
        for ci, c in enumerate(DISPLAY_COLS):
            m = meta.get((sheet_row,c), {})
            bg = (m.get("bg_color") or "").strip()
            txt = (m.get("text_color") or "").strip()
            fsize = (m.get("font_size") or "").strip()
            bold = (m.get("bold") or "").strip()
            if bg:
                try:
                    style_cmds.append(('BACKGROUND', (ci,pdf_row), (ci,pdf_row), colors.HexColor(bg)))
                except Exception:
                    pass
            if txt:
                try:
                    style_cmds.append(('TEXTCOLOR', (ci,pdf_row), (ci,pdf_row), colors.HexColor(txt)))
                except Exception:
                    pass
            if fsize:
                try:
                    style_cmds.append(('FONTSIZE', (ci,pdf_row), (ci,pdf_row), float(fsize)))
                except Exception:
                    pass
            if bold:
                style_cmds.append(('FONTNAME', (ci,pdf_row), (ci,pdf_row), 'Helvetica-Bold'))

    tbl.setStyle(TableStyle(style_cmds))
    story.append(tbl)
    doc.build(story)
    return path


def schedule_numbers_to_rows(sheet, side, start_num, end_num):
    """Translate visible schedule numbers into real rows.
    Example: user enters 1 to 5, app finds rows where the Number column has 1-5.
    It ignores blank number rows.
    """
    meta = cell_meta_for(sheet)
    if side == "numbering":
        key_cols = [1]
    elif side == "fabrication":
        key_cols = [4]
    else:
        key_cols = [1,4]

    try:
        a = float(str(start_num).strip())
        b = float(str(end_num).strip())
    except Exception:
        return []

    low, high = min(a,b), max(a,b)
    rows = []
    for r in range(3,51):
        for c in key_cols:
            val = str(meta.get((r,c),{}).get("value","")).strip()
            if not val:
                continue
            try:
                n = float(val)
            except Exception:
                continue
            if low <= n <= high:
                rows.append(r)
                break
    return sorted(set(rows))

def make_snip_pdf(sheet, start_row, end_row, side):
    """Create a small PDF snip.
    In v17, start/end are schedule numbers typed in the Number column, not Excel row numbers.
    The PDF always prints the top title/date and section header.
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    meta=cell_meta_for(sheet)
    d=cells_for(sheet)

    selected_rows = schedule_numbers_to_rows(sheet, side, start_row, end_row)
    if not selected_rows:
        try:
            sr=max(3, int(float(str(start_row).strip())))
            er=min(50, int(float(str(end_row).strip())))
            if er < sr:
                sr, er = er, sr
            selected_rows=list(range(sr, er+1))
        except Exception:
            selected_rows=[]

    stamp=datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_sheet=''.join(ch if ch.isalnum() or ch in ('-', '_') else '_' for ch in sheet)[:40]
    filename=f"PMW_Snip_{safe_sheet}_{side}_{start_row}-{end_row}_{stamp}.pdf"
    path=os.path.join(EXPORT_FOLDER, filename)

    if side == "numbering":
        cols=[1,2,3]; headers=['NUMBER','NUMBERING','STATUS/NOTES']; widths=[0.65*inch, 5.2*inch, 1.1*inch]
    elif side == "fabrication":
        cols=[4,5,6]; headers=['NUMBER','FABRICATION','STATUS/NOTES']; widths=[0.65*inch, 5.2*inch, 1.1*inch]
    else:
        cols=DISPLAY_COLS; headers=['NUMBER','NUMBERING','STATUS/NOTES','NUMBER','FABRICATION','STATUS/NOTES']; widths=[0.55*inch,3.25*inch,0.85*inch,0.55*inch,3.25*inch,0.85*inch]

    doc=SimpleDocTemplate(path, pagesize=landscape(letter), rightMargin=0.35*inch, leftMargin=0.35*inch, topMargin=0.30*inch, bottomMargin=0.30*inch)
    styles=getSampleStyleSheet()
    title_style=ParagraphStyle('SnipTitle', parent=styles['Heading1'], alignment=1, fontName='Helvetica-Bold', fontSize=15, leading=18)
    date_style=ParagraphStyle('SnipDate', parent=styles['Normal'], alignment=1, fontName='Helvetica-Bold', fontSize=11, leading=13)
    cell_style=ParagraphStyle('SnipCell', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=10)
    head_style=ParagraphStyle('SnipHead', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=11, alignment=1)

    title = d.get((1,2),'FABRICATION SCHEDULE') or 'FABRICATION SCHEDULE'
    datev = d.get((1,4),'') or datetime.now().strftime('%m/%d/%Y')
    story=[
        Paragraph(html.escape(title), title_style),
        Paragraph(html.escape(datev), date_style),
        Spacer(1,6)
    ]

    data=[[Paragraph(h, head_style) for h in headers]]
    row_source=[]
    for r in selected_rows:
        vals=[]; has=False
        for c in cols:
            m=meta.get((r,c), {})
            txt=str(m.get("value",""))
            if m.get("link_path"):
                txt="✉ " + txt
            if txt.strip() or m.get("bg_color") or m.get("text_color"):
                has=True
            vals.append(Paragraph(html.escape(txt), cell_style))
        if has:
            data.append(vals); row_source.append(r)

    if len(data)==1:
        data.append([Paragraph('', cell_style) for _ in headers])

    tbl=Table(data, colWidths=widths, repeatRows=1)
    style_cmds=[
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#d9ead3')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('ALIGN',(0,0),(0,-1),'CENTER'),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
    ]

    for pdf_row, sheet_row in enumerate(row_source, start=1):
        for ci,c in enumerate(cols):
            m=meta.get((sheet_row,c), {})
            bg=(m.get("bg_color") or "").strip()
            txt=(m.get("text_color") or "").strip()
            fsize=(m.get("font_size") or "").strip()
            bold=(m.get("bold") or "").strip()
            if bg:
                try: style_cmds.append(('BACKGROUND',(ci,pdf_row),(ci,pdf_row),colors.HexColor(bg)))
                except Exception: pass
            if txt:
                try: style_cmds.append(('TEXTCOLOR',(ci,pdf_row),(ci,pdf_row),colors.HexColor(txt)))
                except Exception: pass
            if fsize:
                try: style_cmds.append(('FONTSIZE',(ci,pdf_row),(ci,pdf_row),float(fsize)))
                except Exception: pass
            if bold:
                style_cmds.append(('FONTNAME',(ci,pdf_row),(ci,pdf_row),'Helvetica-Bold'))

    tbl.setStyle(TableStyle(style_cmds))
    story.append(tbl)
    doc.build(story)
    return path


def open_outlook_draft_with_attachment(pdf_path, sheet):
    """Try to open an Outlook draft with the PDF already attached. Returns (ok, message)."""
    try:
        import win32com.client
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = f"PMW Schedule {datetime.now().strftime('%m-%d-%y')}"
        mail.Body = "Please see attached PMW schedule PDF.\n\n"
        mail.Attachments.Add(os.path.abspath(pdf_path))
        mail.Display(True)
        return True, 'Outlook draft opened with the PDF attached.'
    except Exception as e:
        return False, str(e)


def reveal_file(path):
    try:
        if platform.system().lower().startswith('win'):
            subprocess.Popen(['explorer', '/select,', os.path.normpath(path)])
        elif platform.system().lower() == 'darwin':
            subprocess.Popen(['open', '-R', path])
        else:
            subprocess.Popen(['xdg-open', os.path.dirname(path)])
    except Exception:
        pass


def cloud_notice_banner():
    if os.environ.get("RENDER"):
        return "<div style='background:#fff3cd;border:1px solid #d6b656;padding:7px;margin:6px;font-weight:bold'>Render v26 Clear Ticket Cells: old database columns are upgraded automatically on startup.</div>"
    return ""


def db_mode_banner():
    if USE_POSTGRES:
        return "<div style='background:#d4edda;border:2px solid #28a745;padding:7px;margin:6px;font-weight:bold'>Database Mode: PostgreSQL Persistent — users and schedule data should survive redeploys.</div>"
    if os.environ.get("RENDER"):
        return "<div style='background:#f8d7da;border:2px solid #dc3545;padding:7px;margin:6px;font-weight:bold'>Database Mode: SQLite TEMP — data can reset. Add DATABASE_URL to Render Environment.</div>"
    return ""

BASE = """
<!doctype html><html><head><meta name='viewport' content='width=device-width, initial-scale=1, maximum-scale=5, user-scalable=yes, viewport-fit=cover'><title>{{app_name}}</title>
<style>
*{box-sizing:border-box} body{margin:0;background:#e7e6e6;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#111}.top{height:38px;background:#107c41;color:white;display:flex;align-items:center;justify-content:space-between;padding:0 14px}.brand{font-weight:bold;font-size:17px}.nav a,.nav span{color:white;margin-left:14px;text-decoration:none;font-weight:bold}.toolbar{padding:7px 9px;background:#f3f3f3;border-bottom:1px solid #aaa;display:flex;gap:8px;align-items:center}.grow{flex:1}.small{font-size:12px;color:#555}.flash{background:#fff3cd;padding:7px;border-bottom:1px solid #d6b656}.tabs{display:flex;gap:3px;padding:5px 8px 0}.tab{background:#d9ead3;border:1px solid #888;border-bottom:0;padding:6px 16px;color:#111;text-decoration:none}.active{background:white;font-weight:bold}.workspace{height:calc(100vh - 88px);overflow:auto;padding:0 8px 16px}.sheetline{display:flex;gap:12px;align-items:flex-start;min-width:1320px}.sheetwrap{background:white;border:1px solid #777;box-shadow:0 1px 5px #aaa}.sheet{border-collapse:collapse;table-layout:fixed;width:1160px}.sheet col.num{width:52px}.sheet col.job{width:450px}.sheet col.note{width:78px}.sheet td{border:1px solid black;height:34px;padding:2px 5px;vertical-align:middle;overflow:hidden}.title{text-align:center;font-size:18pt;font-weight:bold}.section{text-align:center;font-size:16pt;font-weight:bold}.date{text-align:center;font-size:14pt;font-weight:bold}.center{text-align:center}.cellinput{width:100%;height:30px;border:0;background:transparent;font:inherit;outline:none;padding:3px 4px}.cellinput:focus{outline:3px solid #107c41}.numinput,.noteinput{text-align:center}.colorbar{display:flex;gap:5px;align-items:center}.swatch{width:26px;height:24px;border:1px solid #555;border-radius:3px;cursor:pointer}.swatch.white{background:white}.swatch.red{background:#ff6666}.swatch.yellow{background:#fff066}.swatch.green{background:#93d050}.swatch.blue{background:#9dc3e6}.swatch.clear{background:linear-gradient(135deg,#fff 45%,#c00 48%,#c00 52%,#fff 55%)}.linkbtn{display:inline-block;background:#ffd966;border:1px solid #a67c00;border-radius:3px;padding:1px 4px;margin-left:3px;text-decoration:none;color:#111;font-size:12px}.cellbox{display:flex;align-items:center}.cellbox .cellinput{flex:1}.selectedCell{outline:3px solid #1d4ed8!important;box-shadow:inset 0 0 0 2px white}.richEditor{min-height:120px;border:1px solid #777;padding:8px;background:white;color:#111;overflow:auto}.richEditor:focus{outline:2px solid #107c41}.richCell{width:100%;min-height:30px;border:0;background:transparent;outline:none;padding:5px 4px;white-space:pre-wrap;overflow:hidden}.richCell:focus{outline:3px solid #107c41;background:#fffde7}.plainHidden{display:none!important}.jobinput{text-align:left}.buttons{width:175px;padding-top:32px}.buttons button,.btn,button{background:#e9e9e9;border:1px solid #777;border-radius:2px;padding:6px 12px;font:11pt Calibri,Arial;cursor:pointer}.buttons button{width:165px;height:36px;margin-bottom:7px}.green{background:#107c41!important;color:white!important;border-color:#0b5f31!important;font-weight:bold}.login{max-width:430px;margin:80px auto;background:white;padding:22px;border:1px solid #888;box-shadow:0 2px 10px #999}.login input{width:100%;margin:5px 0 12px;padding:7px}.admin{background:white;border-collapse:collapse;width:100%;margin:10px}.admin th,.admin td{border:1px solid #aaa;padding:7px;text-align:left}@media print{
.top,.toolbar,.tabs,.buttons,#snipBox{display:none!important}
.workspace{height:auto!important;overflow:visible!important;padding:0!important}
.sheetline{min-width:0!important}
.sheetwrap{border:0!important;box-shadow:none!important}
*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}
.sheet td,.cellinput{background-color:inherit!important}
.cellinput{color:inherit!important}
}

/* ===== PMW MOBILE LAYOUT ===== */
.mobileTop,.mobileFab{display:none}

@media (max-width: 800px){
  body{font-size:16px;background:#f1f1f1;overflow:hidden}
  .top{
    height:auto;min-height:44px;padding:6px 8px;display:block;position:sticky;top:0;z-index:1000;
  }
  .brand{font-size:15px;line-height:18px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .nav{
    display:flex;gap:10px;overflow-x:auto;padding-top:4px;white-space:nowrap;
  }
  .nav a,.nav span{font-size:13px;margin-left:0}
  .toolbar{
    display:none !important;
  }
  .tabs{
    position:sticky;top:54px;z-index:900;background:#f1f1f1;
    padding:4px 6px;overflow-x:auto;white-space:nowrap;
  }
  .tab{padding:7px 12px;font-size:15px}
  .workspace{
    height:calc(100vh - 105px);
    overflow:auto;
    padding:0 4px 75px 4px;
    -webkit-overflow-scrolling:touch;
  }
  .sheetline{min-width:0;display:block}
  .sheetwrap{
    width:max-content;
    min-width:100%;
    box-shadow:none;
    border:0;
    background:white;
  }
  .sheet{
    width:auto;
    min-width:980px;
    font-size:17px;
  }
  .sheet col.num{width:70px}
  .sheet col.job{width:420px}
  .sheet col.note{width:95px}
  .sheet td{
    height:48px;
    padding:4px 8px;
  }
  .title{font-size:21px;white-space:nowrap}
  .section{font-size:20px}
  .date{font-size:16px}
  .cellinput{
    height:42px;
    font-size:17px;
    padding:6px 5px;
  }
  .richCell{min-height:42px;font-size:17px;padding:8px 5px}
  .buttons{
    position:fixed;
    left:0;right:0;bottom:0;
    width:100%;
    background:#107c41;
    padding:7px;
    display:flex;
    overflow-x:auto;
    gap:6px;
    z-index:1200;
    box-shadow:0 -2px 8px rgba(0,0,0,.25);
  }
  .buttons button{
    min-width:115px;
    width:auto;
    height:44px;
    margin:0;
    font-size:14px;
    background:white;
  }
  .buttons button[name='cmd'][value='clear_schedule'],
  .buttons button[name='cmd'][value='delete_comments']{
    display:none;
  }
  .mobileFab{
    display:flex;
    position:fixed;
    right:14px;
    bottom:64px;
    z-index:1300;
    flex-direction:column;
    gap:8px;
  }
  .mobileFab button{
    border:0;
    border-radius:24px;
    padding:12px 16px;
    background:#107c41;
    color:white;
    font-weight:bold;
    box-shadow:0 2px 8px rgba(0,0,0,.35);
  }
  .mobileTop{
    display:flex;
    position:sticky;
    top:92px;
    z-index:850;
    background:#fff8dc;
    border-bottom:1px solid #d6b656;
    padding:5px 6px;
    gap:5px;
    overflow-x:auto;
    white-space:nowrap;
  }
  .mobileTop button{
    height:36px;
    min-width:58px;
    font-size:13px;
    padding:5px 8px;
    background:white;
  }
  .mobileTop .red{background:#ff6666}
  .mobileTop .yellow{background:#fff066}
  .mobileTop .green{background:#93d050;color:#111}
  .mobileTop .blue{background:#9dc3e6;color:#111}
  .mobileTop .white{background:white;color:#111}
  #snipBox,#richBox{
    left:10px!important;
    right:10px!important;
    top:95px!important;
    transform:none!important;
    width:auto!important;
    max-height:70vh;
    overflow:auto;
  }
  .login{
    margin:25px 10px;
    max-width:none;
  }
  .admin{
    font-size:14px;
    margin:0;
    width:max-content;
    min-width:100%;
  }
  .flash{
    font-size:14px;
    padding:6px 8px;
  }
}


/* ===== V24 MOBILE ZOOM + BUTTON FIX ===== */
@media (max-width: 800px){
  .sheetwrap{
    transform-origin: top left;
    transition: transform .12s ease;
  }
  .mobileZoomLabel{
    font-weight:bold;
    padding:8px 10px;
    background:white;
    border:1px solid #777;
    border-radius:4px;
    color:#111;
  }
  .mobileTop button{
    color:#111 !important;
    font-weight:bold;
    min-width:70px;
  }
  .buttons{
    gap:8px !important;
    padding:8px !important;
    min-height:64px;
  }
  .buttons button{
    color:#111 !important;
    background:#fff !important;
    border:2px solid #d9d9d9 !important;
    min-width:145px !important;
    height:48px !important;
    font-size:15px !important;
    font-weight:bold !important;
    line-height:18px !important;
    padding:6px 10px !important;
    opacity:1 !important;
    visibility:visible !important;
    text-indent:0 !important;
  }
  .buttons button[name='cmd'][value='clear_schedule'],
  .buttons button[name='cmd'][value='delete_comments'],
  .buttons button[name='cmd'][value='done']{
    display:none !important;
  }
  .mobileFab{
    bottom:76px !important;
  }
  .mobileFab button{
    color:white !important;
    min-width:145px;
    font-size:17px !important;
  }
}


.userform{background:white;border:1px solid #bbb;padding:14px;margin:10px 0;box-shadow:0 1px 3px #ccc}
.userform input,.userform select{padding:8px;margin:4px;min-width:160px}
.rolebadge{display:inline-block;padding:3px 8px;border-radius:12px;font-weight:bold}
.role-admin{background:#d9ead3}
.role-editor{background:#fff2cc}
.role-viewer{background:#d9eaf7}

</style></head><body>
{% if session.get('user_id') %}<div class='top'><div class='brand'>{{app_name}} <span style='font-size:12px'>{{version}}</span></div><div class='nav'><span>{{session.username}} / {{session.role}}</span><a href='/'>Workbook</a><a href='/tickets'>Tickets</a>{% if can_admin %}<a href='/users'>Users</a><a href='/audit'>Audit</a>{% endif %}<a href='/logout'>Logout</a></div></div>{% endif %}
{% for m in get_flashed_messages() %}<div class='flash'>{{m}}</div>{% endfor %}
{{body|safe}}</body></html>
"""

def page(body): return render_template_string(BASE, body=db_mode_banner()+cloud_notice_banner()+body, app_name=APP_NAME, version=APP_VERSION, can_admin=can('admin'))

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method=='POST':
        u=request.form.get('username','').strip(); p=request.form.get('password','')
        con=db(); row=con.execute("SELECT * FROM users WHERE username=? AND active=1",(u,)).fetchone(); con.close()
        if row and check_password_hash(row['password_hash'],p):
            session.clear(); session['user_id']=row['id']; session['username']=row['username']; session['role']=row['role']; return redirect('/')
        flash('Bad username or password.')
    return page("""<div class='login'><h1>PMW Login</h1><form method='post'><label>Username</label><input name='username' autofocus><label>Password</label><input type='password' name='password'><button class='green'>Login</button></form><p>admin/admin123 &nbsp; shop/shop123 &nbsp; viewer/view123</p></div>""")

@app.route('/logout')
def logout(): session.clear(); return redirect('/login')

@app.route('/')
@login_required
def index():
    sheets=sheet_names(); active=request.args.get('sheet') or (sheets[0] if sheets else 'Fabrication Schedule')
    d=cells_for(active); meta=cell_meta_for(active); editable=can('editor')
    tabs=''.join(f"<a class='tab {'active' if s==active else ''}' href='/?sheet={urllib.parse.quote(s)}'>{html.escape(s)}</a>" for s in sheets)
    import_html = "" if not can('admin') else """<form method='post' action='/import' enctype='multipart/form-data' style='display:inline-flex;gap:5px;align-items:center'><input type='file' name='workbook' accept='.xlsm,.xlsx'><button>Import Excel</button></form>
<form method='post' action='/preview_ticket_drop' enctype='multipart/form-data' style='display:inline-flex;gap:5px;align-items:center'>
<input type='file' name='ticketlog' accept='.xlsm,.xlsx'>
<button>Open Ticket Drop / Pick Emails</button>
</form>"""
    note = "V18: multi-cell colors, font size, bold, and selected-word editor." if editable else "View only."
    color_html = ""
    if editable:
        color_html = """<div class='colorbar'>
<b>Cell:</b>
<button type='button' class='swatch red' onclick="setCellColor('#ff6666')" title='Red'></button>
<button type='button' class='swatch yellow' onclick="setCellColor('#fff066')" title='Yellow'></button>
<button type='button' class='swatch green' onclick="setCellColor('#93d050')" title='Green'></button>
<button type='button' class='swatch blue' onclick="setCellColor('#9dc3e6')" title='Blue'></button>
<button type='button' class='swatch white' onclick="setCellColor('#ffffff')" title='White'></button>
<button type='button' class='swatch clear' onclick="setCellColor('')" title='Clear'></button>
<b>Text:</b>
<button type='button' class='swatch red' onclick="setTextColor('#c00000')" title='Red text'></button>
<button type='button' class='swatch yellow' onclick="setTextColor('#bf9000')" title='Gold text'></button>
<button type='button' class='swatch green' onclick="setTextColor('#00b050')" title='Green text'></button>
<button type='button' class='swatch blue' onclick="setTextColor('#0070c0')" title='Blue text'></button>
<button type='button' class='swatch clear' onclick="setTextColor('')" title='Clear text'></button>
<b>Size:</b>
<select onchange="setFontSize(this.value); this.selectedIndex=0"><option value=''>Size</option><option value='10'>10</option><option value='11'>11</option><option value='12'>12</option><option value='14'>14</option><option value='16'>16</option><option value='18'>18</option><option value='20'>20</option></select>
<button type='button' onclick="toggleBold()"><b>B</b></button>
<button type='button' onclick="openRichTextEditor()">Edit selected words</button><button type='button' onclick="clearSelectedCells()">Clear Cell + Link</button>
<span class='small'>Drag across cells or Ctrl+click to format many cells.</span>
</div>"""
    body=f"<div class='toolbar'><b>Excel-style workbook</b><span class='small'>{note}</span>{color_html}<span class='grow'></span>{import_html}</div><div class='tabs'>{tabs}</div><div id='saveStatus' style='position:fixed;right:8px;top:48px;z-index:2000;background:#fff3cd;border:1px solid #d6b656;padding:3px 8px;font-size:12px;display:none'>Saved</div>"
    if editable:
        body += "<div class='mobileTop'><button type='button' class='red' onclick=\"setCellColor('#ff6666')\">Red</button><button type='button' class='yellow' onclick=\"setCellColor('#fff066')\">Yellow</button><button type='button' class='green' onclick=\"setCellColor('#93d050')\">Green</button><button type='button' class='blue' onclick=\"setCellColor('#9dc3e6')\">Blue</button><button type='button' class='white' onclick=\"setCellColor('#ffffff')\">White</button><button type='button' onclick=\"setCellColor('')\">Clear</button><button type='button' onclick=\"toggleBold()\"><b>B</b></button><button type='button' onclick=\"openRichTextEditor()\">Words</button><button type='button' onclick=\"clearSelectedCells()\">Clear</button><button type='button' onclick=\"mobileZoomOut()\">Zoom -</button><button type='button' onclick=\"mobileZoomIn()\">Zoom +</button><span class='mobileZoomLabel' id='mobileZoomLabel'>100%</span></div>"
    body += "<div class='workspace'>"
    if editable: body += f"<form id='sheetForm' method='post' action='/save_command'><input type='hidden' name='sheet' value='{html.escape(active)}'>"
    body += "<div class='sheetline'><div class='sheetwrap'><table class='sheet'><col class='num'><col class='job'><col class='note'><col class='num'><col class='job'><col class='note'>"
    datev=d.get((1,4),'') or datetime.now().strftime('%m/%d/%Y')
    body += f"<tr><td></td><td class='title' colspan='2'>{html.escape(d.get((1,2),'FABRICATION SCHEDULE') or 'FABRICATION SCHEDULE')}</td><td class='date' colspan='3'>{html.escape(datev)}</td></tr>"
    body += "<tr>"
    for c in DISPLAY_COLS:
        if c in (2,5): content=f"<div class='section'>{html.escape(d.get((2,c),'') or ('NUMBERING' if c==2 else 'FABRICATION'))}</div>"
        elif c in (3,6) and editable: content="<button type='submit' name='cmd' value='done' class='donebtn'>Done</button>"
        else: content=html.escape(d.get((2,c),'') or '')
        body += f"<td class='center'>{content}</td>"
    body += "</tr>"
    for r in range(3,51):
        body += "<tr>"
        for c in DISPLAY_COLS:
            v=d.get((r,c),'')
            m=meta.get((r,c),{})
            bg=m.get('bg_color',''); txt=m.get('text_color',''); link=m.get('link_path',''); label=m.get('link_label','')
            fsize=m.get('font_size',''); bold=m.get('bold',''); rich=m.get('rich_html','')
            style=''
            if bg: style += f"background-color:{html.escape(bg)};"
            if txt: style += f"color:{html.escape(txt)};"
            if fsize: style += f"font-size:{html.escape(fsize)}pt;"
            if bold: style += "font-weight:bold;"
            link_html = f"<a class='linkbtn' href='/open_link?sheet={urllib.parse.quote(active)}&row={r}&col={c}' title='{html.escape(label or link)}'>✉</a>" if link else ""
            if editable:
                cls = 'numinput' if c in (1,4) else ('jobinput' if c in (2,5) else 'noteinput')
                if rich:
                    body += f"<td style='{style}' data-row='{r}' data-col='{c}'><div class='cellbox'><div class='cellinput richCell' contenteditable='true' data-row='{r}' data-col='{c}' style='{style}'>{rich}</div><input class='plainHidden' name='cell_{r}_{c}' data-row='{r}' data-col='{c}' value='{html.escape(v, quote=True)}' autocomplete='off'><input type='hidden' name='bg_{r}_{c}' value='{html.escape(bg, quote=True)}'><input type='hidden' name='txt_{r}_{c}' value='{html.escape(txt, quote=True)}'><input type='hidden' name='link_{r}_{c}' value='{html.escape(link, quote=True)}'><input type='hidden' name='label_{r}_{c}' value='{html.escape(label, quote=True)}'><input type='hidden' name='fsize_{r}_{c}' value='{html.escape(fsize, quote=True)}'><input type='hidden' name='bold_{r}_{c}' value='{html.escape(bold, quote=True)}'><input type='hidden' name='rich_{r}_{c}' value='{html.escape(rich, quote=True)}'>{link_html}</div></td>"
                else:
                    body += f"<td style='{style}' data-row='{r}' data-col='{c}'><div class='cellbox'><input class='cellinput {cls}' name='cell_{r}_{c}' data-row='{r}' data-col='{c}' style='{style}' value='{html.escape(v, quote=True)}' autocomplete='off'><input type='hidden' name='bg_{r}_{c}' value='{html.escape(bg, quote=True)}'><input type='hidden' name='txt_{r}_{c}' value='{html.escape(txt, quote=True)}'><input type='hidden' name='link_{r}_{c}' value='{html.escape(link, quote=True)}'><input type='hidden' name='label_{r}_{c}' value='{html.escape(label, quote=True)}'><input type='hidden' name='fsize_{r}_{c}' value='{html.escape(fsize, quote=True)}'><input type='hidden' name='bold_{r}_{c}' value='{html.escape(bold, quote=True)}'><input type='hidden' name='rich_{r}_{c}' value='{html.escape(rich, quote=True)}'>{link_html}</div></td>"
            else:
                body += f"<td style='{style}'>{rich if rich else html.escape(v)} {link_html}</td>"
        body += "</tr>"
    body += "</table></div>"
    if editable:
        body += """<div class='buttons'>
<button type='submit' name='cmd' value='clear_schedule' onclick="return confirm('Clear the active schedule?')">Clear Schedule</button>
<button type='submit' name='cmd' value='email_schedule'>Email Schedule</button>
<button type='submit' name='cmd' value='print_pdf'>Print PDF</button>
<button type='submit' name='cmd' value='delete_comments'>Delete Comments</button>
<button type='submit' name='cmd' value='sort_numbering'>Sort Numbering</button>
<button type='submit' name='cmd' value='sort_fabrication'>Sort Fabrication</button>
<button type='button' onclick='window.print()'>Browser Print</button>
<button type='button' onclick='openSnipBox()'>Snip / Print / Email</button>
</div>
<div class='mobileFab'>
<button type='button' onclick='document.querySelector("button[name=cmd][value=email_schedule]").click()'>Email PDF</button>
<button type='button' onclick='document.querySelector("button[name=cmd][value=print_pdf]").click()'>Print PDF</button>
</div>"""
    body += "</div>"
    if editable:
        body += """
<div id='snipBox' style='display:none;position:fixed;right:25px;top:90px;background:white;border:2px solid #107c41;box-shadow:0 2px 12px #555;padding:12px;z-index:50;width:310px'>
<h3 style='margin:0 0 8px'>Snip Schedule Area</h3>
<p class='small'>Type the schedule numbers you want, not Excel row numbers. Example: 1 to 5.</p>
<label>Start #</label><input name='snip_start' value='1' style='width:60px'>
<label>End #</label><input name='snip_end' value='5' style='width:60px'>
<br><br>
<label>Section</label>
<select name='snip_side'>
<option value='both'>Both sides</option>
<option value='numbering'>Numbering only</option>
<option value='fabrication'>Fabrication only</option>
</select>
<br><br>
<button name='cmd' value='snip_pdf' class='green'>Create Snip PDF</button>
<button type='button' onclick='openSnipBox(false)'>Cancel</button>
</div>
<div id='richBox' style='display:none;position:fixed;left:50%;top:90px;transform:translateX(-50%);background:white;border:2px solid #107c41;box-shadow:0 2px 12px #555;padding:12px;z-index:60;width:520px;color:#111'>
<h3 style='margin:0 0 8px'>Edit Selected Words In Cell</h3>
<p class='small'>Highlight words inside this box, then use the buttons. Click Save when done.</p>
<div style='display:flex;gap:5px;margin-bottom:8px;align-items:center'>
<button type='button' onclick="richCmd('bold')"><b>B</b></button>
<button type='button' onclick="richColor('#c00000')">Red text</button>
<button type='button' onclick="richColor('#00b050')">Green text</button>
<button type='button' onclick="richColor('#0070c0')">Blue text</button>
<button type='button' onclick="richColor('#bf9000')">Gold text</button>
<select onchange="richFontSize(this.value); this.selectedIndex=0"><option value=''>Size</option><option value='10'>10</option><option value='12'>12</option><option value='14'>14</option><option value='16'>16</option><option value='18'>18</option><option value='20'>20</option></select>
</div>
<div id='richEditor' class='richEditor' contenteditable='true'></div>
<br><button type='button' class='green' onclick='saveRichText()'>Save Rich Text</button> <button type='button' onclick='closeRichTextEditor()'>Cancel</button>
</div>
<script>
(function(){
  window.selectedCells = new Set();
  window.isSelectingCells = false;

  function cellKey(el){ return el.dataset.row + '_' + el.dataset.col; }

  window.isRichCell = function(el){
    return el && el.classList && el.classList.contains('richCell');
  }

  window.syncCell = function(el){
    if(!el) return;
    const r=el.dataset.row, c=el.dataset.col;
    const plain=document.querySelector(`input[name="cell_${r}_${c}"]`);
    const rich=document.querySelector(`input[name="rich_${r}_${c}"]`);
    if(window.isRichCell(el)){
      if(plain) plain.value = (el.textContent || '').replace(/\s+/g,' ').trim();
      if(rich) rich.value = el.innerHTML;
    }
  }

  window.updateSelectedVisuals = function(){
    document.querySelectorAll('td.selectedCell').forEach(td=>td.classList.remove('selectedCell'));
    window.selectedCells.forEach(k=>{
      const [r,c]=k.split('_');
      const el=document.querySelector(`.cellinput[data-row="${r}"][data-col="${c}"]`);
      if(el) el.closest('td').classList.add('selectedCell');
    });
  }

  window.selectOnly = function(el){
    window.selectedCells.clear();
    window.selectedCells.add(cellKey(el));
    window.activeCell=el;
    window.updateSelectedVisuals();
  }

  window.addSelect = function(el){
    window.selectedCells.add(cellKey(el));
    window.activeCell=el;
    window.updateSelectedVisuals();
  }

  document.querySelectorAll('.cellinput').forEach(el=>{
    el.addEventListener('keydown', function(e){
      if(e.key === 'Enter'){
        e.preventDefault(); e.stopPropagation();
        window.syncCell(this);
        const row=parseInt(this.dataset.row,10)+1;
        const col=this.dataset.col;
        const next=document.querySelector(`.cellinput[data-row="${row}"][data-col="${col}"]`);
        if(next){ next.focus(); window.selectOnly(next); if(next.select) next.select(); }
        return false;
      }
    });
    el.addEventListener('mousedown', function(e){
      window.isSelectingCells = true;
      if(e.ctrlKey || e.metaKey){ window.addSelect(this); }
      else{ window.selectOnly(this); }
    });
    el.addEventListener('mouseover', function(){ if(window.isSelectingCells){ window.addSelect(this); } });
    el.addEventListener('focus', function(){ if(!window.isSelectingCells) window.selectOnly(this); });
    el.addEventListener('click', function(e){ if(e.ctrlKey || e.metaKey) window.addSelect(this); else if(!window.isSelectingCells) window.selectOnly(this); });
    el.addEventListener('input', function(){ window.syncCell(this); });
    el.addEventListener('blur', function(){ window.syncCell(this); });
  });

  document.addEventListener('mouseup', function(){ window.isSelectingCells=false; });

  const form=document.getElementById('sheetForm');
  if(form){ form.addEventListener('submit', function(){ document.querySelectorAll('.cellinput').forEach(window.syncCell); }); }
})();

function selectedInputs(){
  const arr=[];
  if(!window.selectedCells || window.selectedCells.size===0){
    if(window.activeCell) return [window.activeCell];
    return [];
  }
  window.selectedCells.forEach(k=>{
    const [r,c]=k.split('_');
    const el=document.querySelector(`.cellinput[data-row="${r}"][data-col="${c}"]`);
    if(el) arr.push(el);
  });
  return arr;
}

function setCellColor(color){
  const cells=selectedInputs();
  if(cells.length===0){ alert('Click a cell first.'); return; }
  cells.forEach(el=>{
    const r=el.dataset.row, c=el.dataset.col;
    el.style.backgroundColor=color || '';
    el.closest('td').style.backgroundColor=color || '';
    const h=document.querySelector(`input[name="bg_${r}_${c}"]`);
    if(h) h.value=color || '';
  });
}

function setTextColor(color){
  const cells=selectedInputs();
  if(cells.length===0){ alert('Click a cell first.'); return; }
  cells.forEach(el=>{
    const r=el.dataset.row, c=el.dataset.col;
    el.style.color=color || '';
    el.closest('td').style.color=color || '';
    const h=document.querySelector(`input[name="txt_${r}_${c}"]`);
    if(h) h.value=color || '';
  });
}

function setFontSize(size){
  const cells=selectedInputs();
  if(cells.length===0){ alert('Click a cell first.'); return; }
  cells.forEach(el=>{
    const r=el.dataset.row, c=el.dataset.col;
    el.style.fontSize=size ? size+'pt' : '';
    el.closest('td').style.fontSize=size ? size+'pt' : '';
    const h=document.querySelector(`input[name="fsize_${r}_${c}"]`);
    if(h) h.value=size || '';
  });
}

function toggleBold(){
  const cells=selectedInputs();
  if(cells.length===0){ alert('Click a cell first.'); return; }
  cells.forEach(el=>{
    const r=el.dataset.row, c=el.dataset.col;
    const h=document.querySelector(`input[name="bold_${r}_${c}"]`);
    const on = !(h && h.value === '1');
    el.style.fontWeight = on ? 'bold' : '';
    el.closest('td').style.fontWeight = on ? 'bold' : '';
    if(h) h.value = on ? '1' : '';
  });
}

function openSnipBox(show=true){
  const box=document.getElementById('snipBox');
  if(!box) return;
  box.style.display = show ? 'block' : 'none';
}

function stripTags(html){
  const d=document.createElement('div'); d.innerHTML=html; return d.textContent || d.innerText || '';
}

function htmlEscape(s){
  return (s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function openRichTextEditor(){
  if(!window.activeCell){ alert('Click one cell first.'); return; }
  const r=activeCell.dataset.row, c=activeCell.dataset.col;
  const richHidden=document.querySelector(`input[name="rich_${r}_${c}"]`);
  const editor=document.getElementById('richEditor');
  if(richHidden && richHidden.value){
    editor.innerHTML = richHidden.value;
  }else if(window.isRichCell(activeCell)){
    editor.innerHTML = activeCell.innerHTML;
  }else{
    editor.innerHTML = htmlEscape(activeCell.value || '');
  }
  document.getElementById('richBox').style.display='block';
  setTimeout(()=>editor.focus(), 50);
}

function closeRichTextEditor(){
  document.getElementById('richBox').style.display='none';
}

function richCmd(cmd){ document.execCommand(cmd, false, null); }
function richColor(color){ document.execCommand('foreColor', false, color); }
function richFontSize(size){
  const map={'10':'2','12':'3','14':'4','16':'5','18':'6','20':'7'};
  document.execCommand('fontSize', false, map[size] || '3');
}

function attachRichHandlers(div){
  div.addEventListener('keydown', function(e){
    if(e.key === 'Enter'){
      e.preventDefault(); e.stopPropagation();
      window.syncCell(this);
      const next=document.querySelector(`.cellinput[data-row="${parseInt(this.dataset.row,10)+1}"][data-col="${this.dataset.col}"]`);
      if(next){ next.focus(); window.selectOnly(next); }
      return false;
    }
  });
  div.addEventListener('focus', function(){ window.selectOnly(this); });
  div.addEventListener('click', function(){ window.selectOnly(this); });
  div.addEventListener('input', function(){ window.syncCell(this); });
  div.addEventListener('blur', function(){ window.syncCell(this); });
}

function saveRichText(){
  if(!window.activeCell){ closeRichTextEditor(); return; }
  const r=activeCell.dataset.row, c=activeCell.dataset.col;
  const editor=document.getElementById('richEditor');
  const rich=editor.innerHTML;
  const plain=stripTags(rich).replace(/\s+/g,' ').trim();

  const richHidden=document.querySelector(`input[name="rich_${r}_${c}"]`);
  const plainHidden=document.querySelector(`input[name="cell_${r}_${c}"]`);
  if(richHidden) richHidden.value=rich;
  if(plainHidden) plainHidden.value=plain;

  if(window.isRichCell(activeCell)){
    activeCell.innerHTML=rich;
    window.syncCell(activeCell);
  }else{
    const old=activeCell;
    const div=document.createElement('div');
    div.className='cellinput richCell';
    div.contentEditable='true';
    div.dataset.row=r;
    div.dataset.col=c;
    div.style.cssText=old.style.cssText;
    div.innerHTML=rich;
    old.parentNode.insertBefore(div, old);
    old.classList.add('plainHidden');
    attachRichHandlers(div);
    window.activeCell=div;
    window.selectOnly(div);
    window.syncCell(div);
  }
  closeRichTextEditor();
}
window.mobileScheduleZoom = window.mobileScheduleZoom || 1.0;
function applyMobileZoom(){
  const wrap=document.querySelector('.sheetwrap');
  const label=document.getElementById('mobileZoomLabel');
  if(!wrap) return;
  const z=window.mobileScheduleZoom;
  wrap.style.transform='scale('+z+')';
  wrap.style.width=(100/z)+'%';
  wrap.style.marginBottom=((1-z)*900)+'px';
  if(label) label.textContent=Math.round(z*100)+'%';
  try{ localStorage.setItem('pmw_mobile_zoom', String(z)); }catch(e){}
}
function mobileZoomOut(){
  window.mobileScheduleZoom=Math.max(0.45, (window.mobileScheduleZoom||1)-0.1);
  applyMobileZoom();
}
function mobileZoomIn(){
  window.mobileScheduleZoom=Math.min(1.25, (window.mobileScheduleZoom||1)+0.1);
  applyMobileZoom();
}
document.addEventListener('DOMContentLoaded', function(){
  try{
    const saved=parseFloat(localStorage.getItem('pmw_mobile_zoom')||'1');
    if(saved) window.mobileScheduleZoom=saved;
  }catch(e){}
  if(window.innerWidth <= 800) applyMobileZoom();
});

function getHiddenVal(name){ const el=document.querySelector(`input[name="${name}"]`); return el ? el.value : ''; }
function plainTextFromCell(el){
  if(!el) return '';
  if(el.classList && el.classList.contains('richCell')) return (el.textContent || '').replace(/\s+/g,' ').trim();
  return (el.value || '').trim();
}
function richHtmlFromCell(el){
  if(!el) return '';
  if(el.classList && el.classList.contains('richCell')) return el.innerHTML || '';
  const r=el.dataset.row, c=el.dataset.col;
  return getHiddenVal(`rich_${r}_${c}`);
}
let pmwSaveTimers = {};
function showSaveStatus(text){
  const s=document.getElementById('saveStatus');
  if(!s) return;
  s.textContent=text;
  s.style.display='block';
  clearTimeout(window._pmwSaveStatusTimer);
  window._pmwSaveStatusTimer=setTimeout(()=>{ s.style.display='none'; }, 1800);
}
function autosaveCell(el){
  if(!el || !el.dataset) return;
  const r=el.dataset.row, c=el.dataset.col;
  const key=r+'_'+c;
  clearTimeout(pmwSaveTimers[key]);
  pmwSaveTimers[key]=setTimeout(()=>{
    const sheetInput=document.querySelector('input[name="sheet"]');
    const payload={
      sheet: sheetInput ? sheetInput.value : 'Fabrication Schedule',
      row: r,
      col: c,
      value: plainTextFromCell(el),
      bg_color: getHiddenVal(`bg_${r}_${c}`),
      text_color: getHiddenVal(`txt_${r}_${c}`),
      link_path: getHiddenVal(`link_${r}_${c}`),
      link_label: getHiddenVal(`label_${r}_${c}`),
      font_size: getHiddenVal(`fsize_${r}_${c}`),
      bold: getHiddenVal(`bold_${r}_${c}`),
      rich_html: richHtmlFromCell(el)
    };
    showSaveStatus('Saving...');
    fetch('/autosave_cell', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify(payload)
    }).then(r=>r.json()).then(j=>{
      showSaveStatus(j.ok ? 'Saved' : 'Save failed');
    }).catch(()=>showSaveStatus('Save failed'));
  }, 450);
}
document.addEventListener('DOMContentLoaded', function(){
  document.querySelectorAll('.cellinput').forEach(el=>{
    el.addEventListener('input', ()=>autosaveCell(el));
    el.addEventListener('change', ()=>autosaveCell(el));
    el.addEventListener('blur', ()=>autosaveCell(el));
  });
  ['setCellColor','setTextColor','setFontSize','toggleBold'].forEach(fn=>{
    const old=window[fn];
    if(typeof old === 'function'){
      window[fn]=function(...args){
        const ret=old.apply(this,args);
        setTimeout(()=>{ (selectedInputs ? selectedInputs() : []).forEach(autosaveCell); }, 50);
        return ret;
      }
    }
  });
  const oldSaveRich=window.saveRichText;
  if(typeof oldSaveRich === 'function'){
    window.saveRichText=function(...args){
      const ret=oldSaveRich.apply(this,args);
      setTimeout(()=>{ if(window.activeCell) autosaveCell(window.activeCell); }, 50);
      return ret;
    }
  }
});

/* ===== V26 DESKTOP-ONLY ARROW NAVIGATION ===== */
function pmwIsDesktopForArrows(){
  return window.matchMedia && window.matchMedia('(min-width: 801px)').matches;
}
function pmwCellAt(row, col){
  return document.querySelector(`.cellinput[data-row="${row}"][data-col="${col}"]`);
}
function pmwMoveCell(current, rowDelta, colDelta){
  if(!current || !pmwIsDesktopForArrows()) return false;
  const r=parseInt(current.dataset.row || '0',10);
  const c=parseInt(current.dataset.col || '0',10);
  const next=pmwCellAt(r + rowDelta, c + colDelta);
  if(next){
    try{ if(window.syncCell) window.syncCell(current); }catch(e){}
    next.focus();
    if(next.select) next.select();
    try{
      if(window.selectOnly) window.selectOnly(next);
    }catch(e){}
    next.scrollIntoView({block:'nearest', inline:'nearest'});
    return true;
  }
  return false;
}
document.addEventListener('keydown', function(e){
  if(!pmwIsDesktopForArrows()) return;
  const el=document.activeElement;
  if(!el || !el.classList || !el.classList.contains('cellinput')) return;

  // Let Alt+Arrow keep normal cursor/navigation behavior if needed.
  if(e.altKey) return;

  // If a rich text cell is actively being edited, keep arrows inside the text.
  // Plain spreadsheet cells get Excel-like arrow movement.
  const isRich = el.classList.contains('richCell');
  if(isRich && !e.ctrlKey) return;

  let moved=false;
  if(e.key === 'ArrowUp') moved=pmwMoveCell(el,-1,0);
  if(e.key === 'ArrowDown') moved=pmwMoveCell(el,1,0);
  if(e.key === 'ArrowLeft') moved=pmwMoveCell(el,0,-1);
  if(e.key === 'ArrowRight') moved=pmwMoveCell(el,0,1);

  if(moved){
    e.preventDefault();
    e.stopPropagation();
  }
}, true);

/* ===== V30 CLEAR CELL + TICKET LINK ===== */
function clearSelectedCells(){
  const cells = (typeof selectedInputs === 'function') ? selectedInputs() : (window.activeCell ? [window.activeCell] : []);
  if(!cells || cells.length===0){ alert('Click a cell first.'); return; }
  cells.forEach(el=>{
    const r=el.dataset.row, c=el.dataset.col;
    if(el.classList && el.classList.contains('richCell')){
      el.innerHTML='';
    }else{
      el.value='';
    }
    const plain=document.querySelector(`input[name="cell_${r}_${c}"]`);
    if(plain) plain.value='';
    ['bg','txt','link','label','fsize','bold','rich'].forEach(prefix=>{
      const h=document.querySelector(`input[name="${prefix}_${r}_${c}"]`);
      if(h) h.value='';
    });
    el.style.backgroundColor='';
    el.style.color='';
    el.style.fontSize='';
    el.style.fontWeight='';
    const td=el.closest('td');
    if(td){
      td.style.backgroundColor='';
      td.style.color='';
      td.style.fontSize='';
      td.style.fontWeight='';
      const linkBtn=td.querySelector('.linkbtn');
      if(linkBtn) linkBtn.remove();
    }
    try{ if(typeof autosaveCell === 'function') autosaveCell(el); }catch(e){}
  });
}
</script>
</form>"""
    body += "</div>"
    return page(body)

@app.route('/save_command', methods=['POST'])
@login_required
@role_required('editor')
def save_command():
    sheet=request.form.get('sheet','Fabrication Schedule')
    cmd=request.form.get('cmd','save')
    save_posted_cells(sheet)
    if cmd=='sort_numbering':
        sort_side(sheet,1,2,3); log('SORT_NUMBERING',sheet); flash('Numbering sorted.')
    elif cmd=='sort_fabrication':
        sort_side(sheet,4,5,6); log('SORT_FABRICATION',sheet); flash('Fabrication sorted.')
    elif cmd=='clear_schedule':
        con=db(); con.execute("UPDATE workbook_cells SET value='', bg_color='', text_color='', link_path='', link_label='', font_size='', bold='', rich_html='' WHERE sheet_name=? AND row_num>=3 AND col_num IN (1,2,3,4,5,6)",(sheet,)); con.commit(); con.close(); log('CLEAR_SCHEDULE',sheet); flash('Schedule cleared.')
    elif cmd=='delete_comments':
        con=db(); con.execute("UPDATE workbook_cells SET value='', bg_color='', text_color='', link_path='', link_label='', font_size='', bold='', rich_html='' WHERE sheet_name=? AND row_num>=3 AND col_num IN (3,6)",(sheet,)); con.commit(); con.close(); log('DELETE_COMMENTS',sheet); flash('Comments/status cells cleared.')
    elif cmd=='print_pdf':
        try:
            pdf_path = make_schedule_pdf(sheet)
            reveal_file(pdf_path)
            log('PRINT_PDF', os.path.basename(pdf_path))
            flash('Colored schedule PDF created. The folder was opened so you can print it.')
            return redirect('/email_ready?file=' + urllib.parse.quote(os.path.basename(pdf_path)) + '&sheet=' + urllib.parse.quote(sheet))
        except Exception as e:
            log('PRINT_PDF_FAILED', str(e))
            flash('Print PDF failed: ' + str(e))
            return redirect('/?sheet='+urllib.parse.quote(sheet))
    elif cmd=='email_schedule':
        try:
            pdf_path = make_schedule_pdf(sheet)
            ok, msg = open_outlook_draft_with_attachment(pdf_path, sheet)
            log('EMAIL_SCHEDULE_PDF', os.path.basename(pdf_path) + ' | ' + msg)
            if ok:
                flash('PDF created and an Outlook email draft opened with the PDF attached.')
                return redirect('/?sheet='+urllib.parse.quote(sheet))
            else:
                reveal_file(pdf_path)
                flash('PDF created. Outlook attachment draft did not open, so the PDF folder was opened. Attach the PDF manually if needed. Reason: ' + msg)
                return redirect('/email_ready?file=' + urllib.parse.quote(os.path.basename(pdf_path)) + '&sheet=' + urllib.parse.quote(sheet))
        except Exception as e:
            log('EMAIL_SCHEDULE_PDF_FAILED', str(e))
            flash('PDF/email failed: ' + str(e))
            return redirect('/?sheet='+urllib.parse.quote(sheet))
    elif cmd=='snip_pdf':
        try:
            start_row=request.form.get('snip_start','3')
            end_row=request.form.get('snip_end','12')
            side=request.form.get('snip_side','both')
            pdf_path=make_snip_pdf(sheet, start_row, end_row, side)
            log('SNIP_PDF', os.path.basename(pdf_path))
            reveal_file(pdf_path)
            flash('Snip PDF created using the schedule numbers you typed. The folder was opened so you can print or email it.')
            return redirect('/email_ready?file=' + urllib.parse.quote(os.path.basename(pdf_path)) + '&sheet=' + urllib.parse.quote(sheet))
        except Exception as e:
            log('SNIP_PDF_FAILED', str(e))
            flash('Snip PDF failed: ' + str(e))
            return redirect('/?sheet='+urllib.parse.quote(sheet))
    return redirect('/?sheet='+urllib.parse.quote(sheet))


@app.route('/exports/<path:filename>')
@login_required
def exports(filename):
    return send_from_directory(EXPORT_FOLDER, filename, as_attachment=True)

@app.route('/email_ready')
@login_required
def email_ready():
    filename = request.args.get('file','')
    sheet = request.args.get('sheet','Fabrication Schedule')
    safe = os.path.basename(filename)
    subject = f"PMW Schedule {datetime.now().strftime('%m-%d-%y')}"
    mailto = "mailto:?subject=" + urllib.parse.quote(subject) + "&body=" + urllib.parse.quote("Please see attached PMW schedule PDF.\n\n")
    body = f"""
    <div class='workspace'><div class='login' style='max-width:650px;margin-top:35px'>
    <h2>PDF Schedule Created</h2>
    <p>The PDF was created here:</p>
    <p><b>{html.escape(safe)}</b></p>
    <p><a class='btn green' href='/exports/{urllib.parse.quote(safe)}'>Download PDF</a>
    <a class='btn' href='{mailto}'>Open Email</a>
    <a class='btn' href='/?sheet={urllib.parse.quote(sheet)}'>Back to Schedule</a></p>
    <p class='small'>Outlook could not be controlled directly on this computer. Use Download PDF, then attach it to the email window that opens.</p>
    </div></div>
    """
    return page(body)

@app.route('/import', methods=['POST'])
@login_required
@role_required('admin')
def import_route():
    f=request.files.get('workbook')
    if not f or not f.filename:
        flash('Choose an Excel workbook first.'); return redirect('/')
    name=secure_filename(f.filename); path=os.path.join(UPLOAD_FOLDER,name); f.save(path)
    try: count=import_workbook(path); flash(f'Imported {count} filled cells.')
    except Exception as e: flash(f'Import failed: {e}')
    return redirect('/')




@app.route('/preview_ticket_drop', methods=['POST'])
@login_required
@role_required('admin')
def preview_ticket_drop():
    f=request.files.get('ticketlog')
    if not f or not f.filename:
        flash('Choose the Ticket emails drop Excel file first.')
        return redirect('/')
    name=secure_filename(f.filename)
    path=os.path.join(UPLOAD_FOLDER, "ticket_preview_" + name)
    f.save(path)
    try:
        rows=ticket_rows_from_drop(path)
    except Exception as e:
        flash(f'Could not open ticket drop workbook: {e}')
        return redirect('/')

    # show newest first
    rows=list(reversed(rows))
    body="<div class='toolbar'><b>Select Ticket Emails to Import</b><span class='small'>Check the rows you want, then click Import Selected.</span></div>"
    body+=f"<form method='post' action='/import_selected_tickets'><input type='hidden' name='preview_path' value='{html.escape(path, quote=True)}'>"
    body+="<table class='admin'><tr><th>Import</th><th>Excel Row</th><th>Received</th><th>Job #</th><th>Subject</th><th>Sender</th><th>Saved Email Path</th></tr>"
    for t in rows[:300]:
        fp=t.get('file_path','') or ''
        checked = "checked" if t == rows[0] else ""
        body+=f"<tr><td><input type='checkbox' name='rows' value='{int(t.get('source_row',0))}' {checked}></td><td>{int(t.get('source_row',0))}</td><td>{html.escape(t.get('received','') or '')}</td><td>{html.escape(t.get('job_number','') or '')}</td><td>{html.escape(t.get('subject','') or '')}</td><td>{html.escape(t.get('sender','') or '')}</td><td>{html.escape(fp)}</td></tr>"
    body+="</table><div style='padding:10px'><button class='green'>Import Selected Tickets</button> <a class='btn' href='/'>Cancel</a></div></form>"
    return page(body)

@app.route('/import_selected_tickets', methods=['POST'])
@login_required
@role_required('admin')
def import_selected_tickets():
    path=request.form.get('preview_path','')
    selected=request.form.getlist('rows')
    if not path or not os.path.exists(path):
        flash('The preview file could not be found. Please select the ticket drop workbook again.')
        return redirect('/')
    try:
        imported,msg=save_selected_ticket_rows_from_drop(path, selected)
        flash(msg)
        if imported:
            return redirect('/tickets')
    except Exception as e:
        flash(f'Selected ticket import failed: {e}')
    return redirect('/')


@app.route('/import_tickets_count', methods=['POST'])
@login_required
@role_required('admin')
def import_tickets_count_route():
    f=request.files.get('ticketlog')
    if not f or not f.filename:
        flash('Choose the Ticket emails drop Excel file first.')
        return redirect('/')
    name=secure_filename(f.filename)
    path=os.path.join(UPLOAD_FOLDER,name)
    f.save(path)
    count=request.form.get('ticket_count','1')
    try:
        tickets, msg = import_last_n_tickets_from_drop(path, count)
        if tickets:
            flash(f"Imported {len(tickets)} newest ticket drop(s).")
            return redirect('/tickets')
        flash(msg)
    except Exception as e:
        flash(f'Ticket import failed: {e}')
    return redirect('/')

@app.route('/import_one_ticket', methods=['POST'])
@login_required
@role_required('admin')
def import_one_ticket_route():
    f=request.files.get('ticketlog')
    if not f or not f.filename:
        flash('Choose the Ticket emails drop Excel file first.')
        return redirect('/')
    name=secure_filename(f.filename)
    path=os.path.join(UPLOAD_FOLDER,name)
    f.save(path)
    try:
        ticket, msg = import_one_ticket_from_drop(path, "newest")
        if ticket:
            flash(f"Imported newest ticket: {ticket['job_number'] or ''} - {ticket['subject'] or ''}")
            return redirect('/tickets')
        flash(msg)
    except Exception as e:
        flash(f'Newest ticket import failed: {e}')
    return redirect('/')


@app.route('/ticket_upload/<int:ticket_id>', methods=['POST'])
@login_required
@role_required('editor')
def ticket_upload(ticket_id):
    f = request.files.get('msgfile')
    if not f or not f.filename:
        flash('Choose a .msg file first.')
        return redirect('/tickets')

    original = secure_filename(f.filename)
    if not original.lower().endswith('.msg'):
        flash('Please upload the saved Outlook .msg file only.')
        return redirect('/tickets')

    con=db()
    t=con.execute("SELECT id,job_number,subject FROM ticket_links WHERE id=?",(ticket_id,)).fetchone()
    con.close()
    if not t:
        flash('Ticket not found.')
        return redirect('/tickets')

    safe_name = f"ticket_{ticket_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{original}"
    save_path = os.path.join(CLOUD_TICKET_FOLDER, safe_name)
    f.save(save_path)

    con=db()
    con.execute("""UPDATE ticket_links
                   SET cloud_file=?, cloud_filename=?, cloud_uploaded_at=?, cloud_uploaded_by=?
                   WHERE id=?""",
                (safe_name, original, datetime.now().isoformat(timespec='seconds'), session.get('username',''), ticket_id))
    con.commit(); con.close()

    log('UPLOAD_TICKET_MSG', f'{ticket_id} / {original}')
    flash('Cloud ticket email uploaded. Users can now download/open it from the cloud link.')
    return redirect('/tickets')

@app.route('/ticket_download/<int:ticket_id>')
@login_required
def ticket_download(ticket_id):
    from flask import send_file
    con=db()
    t=con.execute("SELECT cloud_file,cloud_filename,file_path,subject FROM ticket_links WHERE id=?",(ticket_id,)).fetchone()
    con.close()
    if not t:
        flash('Ticket not found.')
        return redirect('/tickets')
    cloud_file = (t.get('cloud_file') if hasattr(t, 'get') else t['cloud_file']) or ''
    cloud_filename = (t.get('cloud_filename') if hasattr(t, 'get') else t['cloud_filename']) or ''
    if not cloud_file:
        flash('No cloud .msg file has been uploaded for this ticket yet.')
        return redirect('/tickets')
    full_path = os.path.join(CLOUD_TICKET_FOLDER, cloud_file)
    if not os.path.exists(full_path):
        flash('Cloud file is missing from storage.')
        return redirect('/tickets')
    return send_file(full_path, as_attachment=True, download_name=cloud_filename or cloud_file)

@app.route('/ticket_link_info/<int:ticket_id>')
@login_required
def ticket_link_info(ticket_id):
    con=db()
    t=con.execute("SELECT * FROM ticket_links WHERE id=?",(ticket_id,)).fetchone()
    con.close()
    if not t:
        flash('Ticket not found.')
        return redirect('/tickets')
    subject=html.escape(t['subject'] or '')
    job=html.escape(t['job_number'] or '')
    local_path=html.escape(t['file_path'] or '')
    cloud_file=(t.get('cloud_file') if hasattr(t, 'get') else t['cloud_file']) or ''
    cloud_status = f"<p><a class='btn green' href='/ticket_download/{ticket_id}'>Download Cloud .msg</a></p>" if cloud_file else "<p><b>No cloud .msg uploaded yet.</b></p>"
    body=f"""
    <h2>Ticket Email Link</h2>
    <p><b>Job:</b> {job}</p>
    <p><b>Subject:</b> {subject}</p>
    {cloud_status}
    <h3>Original Office Path</h3>
    <p class='small'>This is the old office/server path from your ticket drop workbook.</p>
    <input value="{local_path}" style="width:90%;padding:10px" onclick="this.select()">
    <p><a class='btn' href='/tickets'>Back to Tickets</a></p>
    """
    return page(body)



@app.route('/clear_cell', methods=['POST'])
@login_required
@role_required('editor')
def clear_cell():
    sheet = request.form.get('sheet') or 'Fabrication Schedule'
    try:
        r = int(request.form.get('row'))
        c = int(request.form.get('col'))
    except Exception:
        flash('Bad cell.')
        return redirect('/?sheet='+urllib.parse.quote(sheet))
    if c not in DISPLAY_COLS:
        flash('Bad cell.')
        return redirect('/?sheet='+urllib.parse.quote(sheet))
    upsert_workbook_cell(sheet,r,c,'','','','','','','','',session.get('username',''))
    log('CLEAR_CELL', f'{sheet} R{r} C{c}')
    flash('Cell cleared, including ticket/email link.')
    return redirect('/?sheet='+urllib.parse.quote(sheet))


@app.route('/tickets')
@login_required
def tickets():
    q=request.args.get('q','').strip()
    con=db()
    if q:
        rows=con.execute("""SELECT * FROM ticket_links WHERE job_number LIKE ? OR subject LIKE ? OR sender LIKE ?
                            ORDER BY id DESC LIMIT 200""",(f'%{q}%',f'%{q}%',f'%{q}%')).fetchall()
    else:
        rows=con.execute("SELECT * FROM ticket_links ORDER BY id DESC LIMIT 200").fetchall()
    con.close()
    body="<div class='toolbar'><b>Imported Tickets</b><form style='display:inline-flex;gap:5px;margin-left:10px'><input name='q' value='"+html.escape(q, quote=True)+"' placeholder='Search job, subject, sender'><button>Search</button></form><span class='small'>Use Add to Numbering/Fabrication to place a linked ticket on the schedule.</span></div>"
    body+="<table class='admin'><tr><th>Received</th><th>Job</th><th>Subject</th><th>Sender</th><th>Email Link</th><th>Cloud Upload</th><th>Add to Schedule</th></tr>"
    for r in rows:
        cloud_file = ''
        try:
            cloud_file = r['cloud_file'] or ''
        except Exception:
            cloud_file = ''
        if cloud_file:
            openlink = "<a class='btn green' href='/ticket_download/"+str(r['id'])+"'>Download .msg</a> <a class='btn' href='/ticket_link_info/"+str(r['id'])+"'>Info</a>"
        else:
            openlink = "<a class='btn' href='/ticket_link_info/"+str(r['id'])+"'>Office Path</a>"
        uploadform = "<form method='post' action='/ticket_upload/"+str(r['id'])+"' enctype='multipart/form-data' style='display:flex;gap:4px;align-items:center'><input type='file' name='msgfile' accept='.msg' style='max-width:190px'><button>Upload .msg</button></form>"
        addlinks = "<a class='btn green' href='/add_ticket_to_schedule?id="+str(r['id'])+"&side=numbering'>Numbering</a> <a class='btn green' href='/add_ticket_to_schedule?id="+str(r['id'])+"&side=fabrication'>Fabrication</a>"
        body += f"<tr><td>{html.escape(r['received'] or '')}</td><td>{html.escape(r['job_number'] or '')}</td><td>{html.escape(r['subject'] or '')}</td><td>{html.escape(r['sender'] or '')}</td><td>{openlink}</td><td>{uploadform}</td><td>{addlinks}</td></tr>"
    body+="</table>"
    return page(body)

@app.route('/open_ticket_id')
@login_required
def open_ticket_id():
    tid=request.args.get('id','')
    con=db(); r=con.execute("SELECT * FROM ticket_links WHERE id=?",(tid,)).fetchone(); con.close()
    if not r:
        flash('Ticket not found.')
        return redirect('/tickets')
    ok,msg=open_path_on_this_pc(r['file_path'] or '')
    flash('Opened ticket/email.' if ok else 'Could not open ticket/email. Reason: '+msg)
    return redirect('/tickets')

@app.route('/add_ticket_to_schedule')
@login_required
@role_required('editor')
def add_ticket_to_schedule_route():
    tid=request.args.get('id','')
    side=request.args.get('side','fabrication')
    ok,msg=add_ticket_to_schedule(tid, side)
    flash(msg)
    return redirect('/')

@app.route('/open_link')
@login_required
def open_link():
    sheet=request.args.get('sheet','Fabrication Schedule')
    row=int(request.args.get('row','0') or 0)
    col=int(request.args.get('col','0') or 0)
    con=db()
    rec=con.execute("SELECT link_path FROM workbook_cells WHERE sheet_name=? AND row_num=? AND col_num=?",(sheet,row,col)).fetchone()
    con.close()
    if not rec or not (rec['link_path'] or '').strip():
        flash('No ticket/email link is attached to that cell.')
        return redirect('/?sheet='+urllib.parse.quote(sheet))
    link = rec['link_path'] or ''
    if link.startswith('/'):
        return redirect(link)
    ok,msg=open_path_on_this_pc(link)
    flash('Opened linked ticket/email.' if ok else 'Could not open linked ticket/email. Reason: '+msg)
    return redirect('/?sheet='+urllib.parse.quote(sheet))


@app.route('/autosave_cell', methods=['POST'])
@login_required
@role_required('editor')
def autosave_cell():
    data = request.get_json(silent=True) or {}
    sheet = data.get('sheet') or 'Fabrication Schedule'
    try:
        r = int(data.get('row'))
        c = int(data.get('col'))
    except Exception:
        return jsonify({"ok": False, "error": "Bad row/col"}), 400
    if r < 1 or r > 200 or c not in DISPLAY_COLS:
        return jsonify({"ok": False, "error": "Out of range"}), 400

    val = (data.get('value') or '').replace('\r',' ').replace('\n',' ').strip()
    bg = (data.get('bg_color') or '').strip()
    txt = (data.get('text_color') or '').strip()
    link = (data.get('link_path') or '').strip()
    label = (data.get('link_label') or '').strip()
    fsize = (data.get('font_size') or '').strip()
    bold = (data.get('bold') or '').strip()
    rich = (data.get('rich_html') or '').strip()
    now = datetime.now().isoformat(timespec='seconds')

    con = db()
    con.execute("""INSERT OR REPLACE INTO workbook_cells(
        sheet_name,row_num,col_num,value,bg_color,text_color,link_path,link_label,font_size,bold,rich_html,updated_by,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (sheet,r,c,val,bg,txt,link,label,fsize,bold,rich,session.get('username',''),now))
    con.commit()
    con.close()
    return jsonify({"ok": True, "saved_at": now})


@app.route('/users')
@login_required
@role_required('admin')
def users():
    con=db()
    rows=con.execute("SELECT id,username,role,active,created_at FROM users ORDER BY username").fetchall()
    con.close()

    role_options = "<option value='viewer'>viewer - can only view</option><option value='editor'>editor - can view and edit schedule</option><option value='admin'>admin - full access</option>"

    body = """
    <h2>Admin - Users & Permissions</h2>
    <div class='userform'>
      <h3>Create New User</h3>
      <form method='post' action='/users/create'>
        <input name='username' placeholder='username' required>
        <input name='password' placeholder='password' required>
        <select name='role'>""" + role_options + """</select>
        <button class='green'>Create User</button>
      </form>
      <p class='small'><b>viewer</b> can only view. <b>editor</b> can edit schedules/tickets. <b>admin</b> can manage users.</p>
    </div>
    <table class='admin'>
      <tr><th>Username</th><th>Role</th><th>Active</th><th>Created</th><th>Change Role</th><th>Password</th><th>Status</th></tr>
    """

    for u in rows:
        uid=u['id']
        username=html.escape(u['username'] or '')
        role=u['role'] or 'viewer'
        active='Yes' if int(u['active'] or 0)==1 else 'No'
        role_class='role-' + html.escape(role)
        selected_viewer='selected' if role=='viewer' else ''
        selected_editor='selected' if role=='editor' else ''
        selected_admin='selected' if role=='admin' else ''
        disable_self = "disabled" if u['username']==session.get('username') else ""

        body += f"""
        <tr>
          <td><b>{username}</b></td>
          <td><span class='rolebadge {role_class}'>{html.escape(role)}</span></td>
          <td>{active}</td>
          <td>{html.escape(u['created_at'] or '')}</td>
          <td>
            <form method='post' action='/users/update_role' style='display:inline-flex;gap:5px'>
              <input type='hidden' name='user_id' value='{uid}'>
              <select name='role'>
                <option value='viewer' {selected_viewer}>viewer</option>
                <option value='editor' {selected_editor}>editor</option>
                <option value='admin' {selected_admin}>admin</option>
              </select>
              <button>Update</button>
            </form>
          </td>
          <td>
            <form method='post' action='/users/reset_password' style='display:inline-flex;gap:5px'>
              <input type='hidden' name='user_id' value='{uid}'>
              <input name='password' placeholder='new password'>
              <button>Reset</button>
            </form>
          </td>
          <td>
            <form method='post' action='/users/toggle_active' style='display:inline'>
              <input type='hidden' name='user_id' value='{uid}'>
              <button {disable_self}>{'Disable' if int(u['active'] or 0)==1 else 'Enable'}</button>
            </form>
          </td>
        </tr>
        """

    body += "</table>"
    return page(body)

@app.route('/users/create', methods=['POST'])
@login_required
@role_required('admin')
def users_create():
    username=(request.form.get('username') or '').strip()
    password=(request.form.get('password') or '').strip()
    role=(request.form.get('role') or 'viewer').strip()
    if role not in ('viewer','editor','admin'):
        role='viewer'
    if not username or not password:
        flash('Username and password are required.')
        return redirect('/users')
    con=db()
    try:
        con.execute("INSERT INTO users(username,password_hash,role,active,created_at) VALUES(?,?,?,?,?)",
                    (username,generate_password_hash(password),role,1,datetime.now().isoformat(timespec='seconds')))
        con.commit()
        log('CREATE_USER', f'{username} / {role}')
        flash(f'Created user {username} as {role}.')
    except Exception as e:
        flash('Could not create user: ' + str(e))
    con.close()
    return redirect('/users')

@app.route('/users/update_role', methods=['POST'])
@login_required
@role_required('admin')
def users_update_role():
    uid=request.form.get('user_id')
    role=(request.form.get('role') or 'viewer').strip()
    if role not in ('viewer','editor','admin'):
        role='viewer'
    con=db()
    row=con.execute("SELECT username FROM users WHERE id=?",(uid,)).fetchone()
    if row and row['username']==session.get('username') and role!='admin':
        flash('You cannot remove admin from your own logged-in account.')
        con.close()
        return redirect('/users')
    con.execute("UPDATE users SET role=? WHERE id=?",(role,uid))
    con.commit()
    con.close()
    log('UPDATE_USER_ROLE', f'id {uid} -> {role}')
    flash('User role updated.')
    return redirect('/users')

@app.route('/users/reset_password', methods=['POST'])
@login_required
@role_required('admin')
def users_reset_password():
    uid=request.form.get('user_id')
    password=(request.form.get('password') or '').strip()
    if not password:
        flash('Enter a new password first.')
        return redirect('/users')
    con=db()
    con.execute("UPDATE users SET password_hash=? WHERE id=?",(generate_password_hash(password),uid))
    con.commit()
    con.close()
    log('RESET_USER_PASSWORD', f'id {uid}')
    flash('Password reset.')
    return redirect('/users')

@app.route('/users/toggle_active', methods=['POST'])
@login_required
@role_required('admin')
def users_toggle_active():
    uid=request.form.get('user_id')
    con=db()
    row=con.execute("SELECT username,active FROM users WHERE id=?",(uid,)).fetchone()
    if not row:
        flash('User not found.')
        con.close()
        return redirect('/users')
    if row['username']==session.get('username'):
        flash('You cannot disable your own logged-in account.')
        con.close()
        return redirect('/users')
    new_active=0 if int(row['active'] or 0)==1 else 1
    con.execute("UPDATE users SET active=? WHERE id=?",(new_active,uid))
    con.commit()
    con.close()
    log('TOGGLE_USER_ACTIVE', f'id {uid} -> {new_active}')
    flash('User status updated.')
    return redirect('/users')


@app.route('/audit')
@login_required
@role_required('admin')
def audit():
    con=db(); rows=con.execute('SELECT * FROM audit_log ORDER BY id DESC LIMIT 200').fetchall(); con.close()
    body='<table class="admin"><tr><th>When</th><th>User</th><th>Action</th><th>Details</th></tr>'
    for r in rows: body += f"<tr><td>{r['created_at']}</td><td>{html.escape(r['username'] or '')}</td><td>{r['action']}</td><td>{html.escape(r['details'] or '')}</td></tr>"
    return page(body+'</table>')



def upgrade_ticket_cloud_columns():
    try:
        con=db(); cur=con.cursor()
        for coldef in [
            "cloud_file TEXT DEFAULT ''",
            "cloud_filename TEXT DEFAULT ''",
            "cloud_uploaded_at TEXT DEFAULT ''",
            "cloud_uploaded_by TEXT DEFAULT ''"
        ]:
            try:
                cur.execute("ALTER TABLE ticket_links ADD COLUMN " + coldef)
            except Exception:
                try:
                    con.con.rollback()
                except Exception:
                    pass
        con.commit(); con.close()
    except Exception as e:
        print("Ticket cloud column upgrade failed:", repr(e))


def startup_init_for_cloud():
    try:
        print("PMW DB MODE:", "PostgreSQL" if USE_POSTGRES else "SQLite", "DATABASE_URL set:", bool(DATABASE_URL), "psycopg:", bool(psycopg))
        init_db()
        upgrade_ticket_cloud_columns()

        con=db()
        try:
            row = con.execute("SELECT COUNT(*) AS n FROM workbook_cells").fetchone()
            n = row["n"] if row else 0
        except Exception:
            n = 0
        con.close()

        starter = os.path.join(APP_DIR, 'Ticket +Fabrication-ACTIVE(1).xlsm')
        if n == 0 and os.path.exists(starter):
            with app.test_request_context('/'):
                session['username'] = 'system'
                try:
                    import_workbook(starter)
                    print("Starter workbook imported.")
                except Exception as e:
                    print("Starter workbook import skipped:", repr(e))
    except Exception as e:
        print("Startup database initialization failed:", repr(e))

startup_init_for_cloud()

if __name__ == '__main__':
    init_db()
    con=db(); n=con.execute('SELECT COUNT(*) FROM workbook_cells').fetchone()[0]; con.close()
    starter=os.path.join(APP_DIR,'Ticket +Fabrication-ACTIVE(1).xlsm')
    if n == 0 and os.path.exists(starter):
        with app.test_request_context('/'):
            session['username']='system'
            try: import_workbook(starter)
            except Exception as e: print('Starter import skipped:',e)
    print('====================================================')
    print('PMW Ticket + Fabrication APP v30 Clear Ticket Cells')
    print('Open http://127.0.0.1:5050')
    print('====================================================')
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5050)), debug=False)
