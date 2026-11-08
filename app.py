"""CEAT Safety workplace safety portal (Python 3.10+; Excel support bundled in vendor/)."""
import os,sys,sqlite3,html,json,csv,io,zipfile,hashlib,hmac,secrets,re,mimetypes,urllib.parse,datetime,logging,threading
from pathlib import Path
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
from http.cookies import SimpleCookie
from email.parser import BytesParser
from email.policy import default
from xml.sax.saxutils import escape as xml_escape

ROOT=Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT/'vendor'))
from department_exports import DEPARTMENTS, division_access, division_matches, create_department_workbook
DATA=Path(os.environ.get('SAFETRACK_DATA',ROOT/'data'));DATA.mkdir(parents=True,exist_ok=True)
DB=DATA/'safetrack.db';UPLOAD=DATA/'uploads';UPLOAD.mkdir(exist_ok=True)
MAX_BODY=9*1024*1024
ROLES=('staff','responsible','manager','safety','admin')
def today():return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
def month():return datetime.date.today().strftime('%Y-%m')
def con():
 c=sqlite3.connect(DB,timeout=20);c.row_factory=sqlite3.Row;c.execute('PRAGMA foreign_keys=ON');c.execute('PRAGMA busy_timeout=20000');return c
def init():
 with con() as c:
  c.execute('PRAGMA journal_mode=WAL')
  c.executescript('''CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,code TEXT UNIQUE,name TEXT NOT NULL,designation TEXT,division TEXT NOT NULL,role TEXT DEFAULT 'staff',email TEXT DEFAULT '',salt TEXT NOT NULL,pwhash TEXT NOT NULL,active INTEGER DEFAULT 1);
  CREATE TABLE IF NOT EXISTS sessions(tokenhash TEXT PRIMARY KEY,uid INTEGER NOT NULL,csrf TEXT NOT NULL,expires TEXT NOT NULL,FOREIGN KEY(uid) REFERENCES users(id));
  CREATE TABLE IF NOT EXISTS reports(id INTEGER PRIMARY KEY,reporter INTEGER NOT NULL,division TEXT NOT NULL,kind TEXT NOT NULL,location TEXT NOT NULL,process TEXT,operation TEXT,description TEXT NOT NULL,priority TEXT NOT NULL,photo TEXT,assignee INTEGER,deadline TEXT,status TEXT DEFAULT 'Open',created TEXT NOT NULL,closed TEXT,FOREIGN KEY(reporter) REFERENCES users(id),FOREIGN KEY(assignee) REFERENCES users(id));
  CREATE TABLE IF NOT EXISTS actions(id INTEGER PRIMARY KEY,rid INTEGER NOT NULL,actor INTEGER NOT NULL,kind TEXT,note TEXT,photo TEXT,created TEXT NOT NULL,FOREIGN KEY(rid) REFERENCES reports(id));
  CREATE TABLE IF NOT EXISTS notices(id INTEGER PRIMARY KEY,uid INTEGER NOT NULL,rid INTEGER NOT NULL,message TEXT,seen INTEGER DEFAULT 0,created TEXT NOT NULL);
  CREATE TABLE IF NOT EXISTS confirm(uid INTEGER NOT NULL,month TEXT NOT NULL,created TEXT NOT NULL,PRIMARY KEY(uid,month));
  CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY,uid INTEGER,event TEXT,rid INTEGER,created TEXT NOT NULL);\n  CREATE TABLE IF NOT EXISTS app_meta(key TEXT PRIMARY KEY,value TEXT);''')
  # Existing databases are migrated without changing stored reports or user IDs.
  user_cols=[r[1] for r in c.execute('PRAGMA table_info(users)')]
  if 'epf' not in user_cols:c.execute("ALTER TABLE users ADD COLUMN epf TEXT DEFAULT ''")
  if 'must_change_password' not in user_cols:c.execute("ALTER TABLE users ADD COLUMN must_change_password INTEGER DEFAULT 0")
  c.execute('UPDATE users SET must_change_password=0')
  report_cols=[r[1] for r in c.execute('PRAGMA table_info(reports)')]
  if 'responsible_division' not in report_cols:c.execute("ALTER TABLE reports ADD COLUMN responsible_division TEXT DEFAULT ''")
  # One-time EPF login migration; does not create the employee with a missing EPF.
  migrated=c.execute("SELECT value FROM app_meta WHERE key='epf_login_migration_v1'").fetchone()
  if not migrated:
   # The roster sync runs after init exits to avoid nested DB connections.
   pass


def hp(pw,salt):return hashlib.pbkdf2_hmac('sha256',pw.encode(),bytes.fromhex(salt),250000).hex()
def valid_password(pw):
 return bool(re.fullmatch(r'\d{8}',str(pw or '')))
def set_pw(c,uid,pw):
 if not valid_password(pw):raise ValueError('Password must contain exactly 8 digits.')
 s=secrets.token_hex(16);c.execute('UPDATE users SET salt=?,pwhash=? WHERE id=?',(s,hp(pw,s),uid))
def import_staff():
 fp=ROOT/'staff_import.csv'
 if not fp.exists():return 0
 count=0
 with con() as c,open(fp,encoding='utf-8-sig',newline='') as f:
  for r in csv.DictReader(f):
   epf=(r.get('employee_code') or '').strip().upper()
   digits=re.sub(r'\D','',epf)
   if not re.fullmatch(r'[A-Z]?\d{4}',epf) or len(digits)!=4:continue
   code=digits;pw=digits*2
   existing=c.execute('SELECT * FROM users WHERE code=? OR code=? OR lower(name)=lower(?)',(code,epf,r['name'])).fetchone()
   if existing:
    # Convert matching pre-existing imported staff accounts to the new EPF-based login.
    conflict=c.execute('SELECT id FROM users WHERE code=? AND id!=?',(code,existing['id'])).fetchone()
    if conflict:continue
    salt=secrets.token_hex(16)
    c.execute('UPDATE users SET code=?,epf=?,name=?,designation=?,division=?,salt=?,pwhash=?,must_change_password=0 WHERE id=?',
      (code,epf,r['name'],r.get('designation',''),r['division'],salt,hp(pw,salt),existing['id']))
   else:
    salt=secrets.token_hex(16)
    c.execute('INSERT INTO users(code,epf,name,designation,division,role,email,salt,pwhash,must_change_password) VALUES(?,?,?,?,?,?,?,?,?,0)',
      (code,epf,r['name'],r.get('designation',''),r['division'],'staff','',salt,hp(pw,salt)))
    count+=1
 return count

def esc(x):return html.escape(str(x if x is not None else ''),quote=True)
def opt(items,selected=None):return ''.join(f'<option value="{esc(v)}" {"selected" if str(v)==str(selected) else ""}>{esc(label)}</option>' for v,label in items)
def field(label,name,value='',typ='text',required=False):return f'<label>{esc(label)}<input type="{typ}" name="{name}" value="{esc(value)}" {"required" if required else ""}></label>'
def badge(status):
 color='green' if status=='Closed' else ('orange' if status in ('Open','Reopened','Pending Verification') else 'blue')
 return f'<span class="pill {color}">{esc(status)}</span>'
CSS='''*{box-sizing:border-box}body{margin:0;background:#f5f7fa;color:#14283e;font:15px/1.5 system-ui,Segoe UI,Arial,sans-serif}a{color:#176f73;text-decoration:none}a:hover{text-decoration:underline}.sidebar{position:fixed;inset:0 auto 0 0;width:242px;background:#122a39;color:#e8f4f5;padding:25px 17px;overflow:auto}.brand{font-size:23px;font-weight:750;color:#fff;margin-bottom:28px}.brand small{display:block;font-size:10px;letter-spacing:2px;color:#88b6bf}.identity{padding:12px;background:#1d394a;border-radius:10px;margin-bottom:16px}.identity small{display:block;color:#bfd4dd}.sidebar nav a{display:block;padding:11px;color:#d1e4e7;border-radius:8px}.sidebar nav a:hover{background:#284c5c}.navlogout{background:none;border:0;color:#d9e7ed;cursor:pointer;padding:12px;font-size:14px}.content{margin-left:242px;padding:35px;max-width:1650px}.login-content{max-width:470px;margin:7vh auto;padding:20px}h1{font-size:29px;line-height:1.2;margin:7px 0}h2{font-size:19px;margin:0 0 18px}h3{font-size:15px;margin-bottom:7px}.muted{color:#697d8d}.eyebrow{color:#178177;letter-spacing:1.4px;font-size:11px;font-weight:800}.top{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:25px}.panel{background:#fff;border:1px solid #e1e7ee;border-radius:14px;padding:24px;box-shadow:0 3px 15px #14263b09;margin-bottom:20px}.cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:15px;margin-bottom:24px}.stat{background:#fff;border-radius:14px;padding:20px;border:1px solid #e1e7ee}.stat b{display:block;font-size:32px}.stat small{color:#748494}.twocol{display:grid;grid-template-columns:1fr 1fr;gap:20px}.stack>*{margin-bottom:20px}label{display:block;font-weight:600;font-size:13px;margin:0 0 14px}input,select,textarea{display:block;width:100%;margin-top:6px;padding:11px;border:1px solid #cfdae1;border-radius:8px;background:#fff;font:inherit;color:#14283e}textarea{min-height:100px}.primary{display:inline-block;border:0;border-radius:9px;background:#147c73;color:#fff!important;padding:11px 19px;cursor:pointer;font-weight:700;font-size:14px}.secondary{background:#e5efef;color:#176d73;border:0;border-radius:9px;padding:10px 14px;cursor:pointer}.pill{display:inline-block;padding:4px 10px;border-radius:100px;font-size:12px;background:#e4ebf4;color:#3b607c;white-space:nowrap}.pill.green{background:#e2f8ed;color:#137345}.pill.orange{background:#fff0db;color:#9e6110}.pill.blue{background:#e3f1fe;color:#2367a0}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;font-size:13px}th{text-align:left;background:#f1f5f8;color:#546979}td,th{padding:11px;border-bottom:1px solid #e9edf1;vertical-align:top}tr:hover td{background:#f9fbfd}.flash{padding:12px 16px;background:#e9f8ed;border:1px solid #bae4c7;border-radius:8px;margin-bottom:16px}.warn{background:#fff2df;border:1px solid #ffd8a7;border-radius:8px;padding:12px;margin-bottom:14px}.evidence{max-width:100%;max-height:340px;border-radius:9px}.timeline{border-left:2px solid #c6e7dd;padding-left:18px}.event{padding:8px 0 20px}.event small{display:block;color:#697d8d}dl{display:grid;grid-template-columns:130px 1fr;gap:10px}dt{font-weight:bold;color:#5c7184}dd{margin:0}.buttons{display:flex;gap:10px;flex-wrap:wrap;margin:12px 0}.hidden{display:none}@media(max-width:950px){.sidebar{position:static;width:100%;padding:16px}.sidebar nav{display:flex;flex-wrap:wrap}.content{margin:0;padding:16px}.twocol{grid-template-columns:1fr}.cards{grid-template-columns:repeat(2,1fr)}}@media(max-width:450px){.cards{grid-template-columns:1fr 1fr}.panel{padding:14px}dl{grid-template-columns:1fr}}'''

CSS += """
:root{--primary:#d71920;--ink:#171717;--bg:#f5f5f5}
body{background:#f4f4f4;color:#202020}
.sidebar{background:linear-gradient(180deg,#171717,#262626);border-right:5px solid #d71920}
.brand{color:#fff}.brand small{color:#ffb6b8}
.sidebar nav a:hover{background:#d719202b}
.eyebrow{color:#d71920}
.primary{background:#d71920;box-shadow:0 5px 12px #d7192030}
.primary:hover{background:#b81218}
.stat{border-top-color:#d71920}.stat b{color:#b51219}
a{color:#c8171e}input:focus,select:focus,textarea:focus{outline-color:#e45b60}
.pill.red{background:#ffe3e3;color:#b51219;font-weight:800}
.notification-unread{background:#fff0f0;border:1px solid #f2b9bc;border-left:5px solid #d71920;padding:14px;border-radius:10px;margin:10px 0}
.notification-read{background:#fafafa;border:1px solid #e7e7e7;padding:14px;border-radius:10px;margin:10px 0}
.notification-dot{display:inline-block;width:9px;height:9px;border-radius:50%;background:#d71920;margin-right:7px}
.division-section{margin:24px 0 22px;padding:20px;border:1px solid #e5e7eb;border-radius:16px;background:linear-gradient(135deg,#fff 0%,#fff8f8 100%)}
.division-heading{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:16px}
.division-heading h2{margin:0 0 3px}.division-heading p{margin:0;font-size:13px}
.division-count{font-size:10px;font-weight:850;letter-spacing:1px;color:#b51219;background:#ffe7e8;border:1px solid #ffd0d2;border-radius:30px;padding:7px 11px;white-space:nowrap}
.division-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
.division-card{position:relative;display:flex;min-height:116px;flex-direction:column;align-items:flex-start;justify-content:center;gap:5px;margin:0;padding:14px;border:1px solid #e4e7eb;border-radius:13px;background:#fff;cursor:pointer;overflow:hidden;transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease,background .2s ease}
.division-card:hover{transform:translateY(-5px);border-color:#d71920;box-shadow:0 12px 24px #991b1b18;background:linear-gradient(145deg,#fff,#fff4f4)}
.division-card input{position:absolute;inset:0;width:100%;height:100%;margin:0;opacity:0;cursor:pointer;z-index:2}
.division-icon{display:grid;place-items:center;width:37px;height:37px;border-radius:11px;background:#f3f4f6;font-size:21px;transition:transform .2s ease,background .2s ease}
.division-card:hover .division-icon{transform:scale(1.12) rotate(-4deg);background:#ffe3e4}
.division-name{font-size:13px;font-weight:800;color:#242424;line-height:1.25}.division-hint{font-size:10px;color:#84909b;font-weight:500}
.division-check{position:absolute;right:11px;top:10px;display:grid;place-items:center;width:21px;height:21px;border-radius:50%;background:#d71920;color:white;font-size:12px;opacity:0;transform:scale(.7);transition:all .18s ease}
.division-card:has(input:checked){border:2px solid #d71920;background:linear-gradient(145deg,#fff,#fff0f0);box-shadow:0 6px 18px #d719201c;padding:13px}
.division-card:has(input:checked) .division-check{opacity:1;transform:scale(1)}
.division-card:has(input:focus-visible){outline:3px solid #f6a4a7;outline-offset:2px}
.report-panel{padding:28px}
@media(max-width:1100px){.division-grid{grid-template-columns:repeat(3,minmax(0,1fr))}}
@media(max-width:700px){.division-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.division-section{padding:14px}.division-card{min-height:105px;padding:11px}}
@media(prefers-reduced-motion:reduce){.division-card,.division-icon,.division-check{transition:none}}
"""
CSS += '\n:root{color-scheme:dark;--ceat-red:#e21b2d;--ceat-red2:#ff4655;--surface:#17191e;--surface2:#20232a;--surface3:#292d35;--line:#343943;--muted:#a0a7b3}\n*{box-sizing:border-box}\nbody{background:radial-gradient(ellipse at 75% -10%,#3b171d 0%,#17191e 42%,#101216 100%);color:#f3f4f6}\na{color:#ff6974}\n.sidebar{background:linear-gradient(180deg,#0b0d11 0%,#17191e 70%,#211317 100%);border-right:1px solid #3a2529;box-shadow:8px 0 30px #0005}\n.brand{display:flex;align-items:center;gap:12px;color:#fff;font-size:22px;letter-spacing:.4px}\n.brand img{width:112px;height:auto;display:block;filter:drop-shadow(0 3px 8px #0008)}\n.brand small{color:#ff7b85;letter-spacing:2.2px;font-size:9px}\n.identity{background:linear-gradient(135deg,#252830,#1b1e25);border:1px solid #3b404a;color:#fff}\n.identity small,.muted{color:#a7afbc}\n.sidebar nav a{color:#d9dde5;transition:all .22s ease;border:1px solid transparent}\n.sidebar nav a:hover{background:linear-gradient(90deg,#d7192033,#d719200a);border-color:#71343a;color:#fff;transform:translateX(4px);text-decoration:none}\n.navlogout{border-radius:9px;transition:all .2s ease}.navlogout:hover{background:#d7192030;color:#fff}\n.content{color:#f0f1f4}\nh1,h2,h3{color:#fff}\n.eyebrow{color:#ff6874}\n.panel,.stat{background:linear-gradient(145deg,#20232a,#191c22);border:1px solid #343943;box-shadow:0 12px 35px #0003;color:#eef0f4;transition:transform .22s ease,border-color .22s ease,box-shadow .22s ease}\n.panel:hover,.stat:hover{border-color:#7a333a;box-shadow:0 16px 42px #0005,0 0 0 1px #d7192014}\n.stat b{color:#ff6570}.stat small{color:#aeb5c0}\ninput,select,textarea{background:#11141a;color:#f4f5f7;border:1px solid #414752;transition:border-color .2s ease,box-shadow .2s ease,background .2s ease}\ninput::placeholder,textarea::placeholder{color:#737d8a}\ninput:focus,select:focus,textarea:focus{outline:none;border-color:#ed3948;box-shadow:0 0 0 3px #e21b2d26;background:#171a20}\nlabel{color:#dfe2e8}\n.primary,.secondary,button{transition:transform .18s ease,box-shadow .18s ease,background .18s ease,border-color .18s ease}\n.primary{background:linear-gradient(135deg,#ef2939,#b90e1c);color:#fff!important;box-shadow:0 7px 20px #d7192038;border:1px solid #ff6570}\n.primary:hover{background:linear-gradient(135deg,#ff4655,#d71920);transform:translateY(-2px);box-shadow:0 11px 27px #d7192055;text-decoration:none}\n.primary:active,.secondary:active{transform:translateY(0) scale(.98)}\n.secondary{background:#2b3039;color:#f3f4f6;border:1px solid #454b56}.secondary:hover{background:#3a414d;border-color:#e33a48;transform:translateY(-2px)}\ntable{color:#e7e9ed}th{background:#2a2e37;color:#f0f1f4}td,th{border-bottom:1px solid #343943}tr:hover td{background:#2a2025}\n.pill{background:#303743;color:#e1e6ed}.pill.green{background:#153d2c;color:#8ce1b2}.pill.orange{background:#4a3218;color:#ffd18b}.pill.blue{background:#18344e;color:#9bcfff}.pill.red{background:#571c24;color:#ff9ba3}\n.flash{background:#173a2a;border:1px solid #2b7650;color:#c5f6d8}.warn{background:#3b2a17;border:1px solid #805c28;color:#ffe0ad}\n.notification-unread{background:#401c23;border:1px solid #9e343f;border-left:5px solid #f02b3b;color:#fff}\n.notification-read{background:#20232a;border:1px solid #3a404a;color:#d8dce3}\n.division-section{background:linear-gradient(135deg,#20232a,#191b21);border:1px solid #3a404a}\n.division-card{background:linear-gradient(145deg,#282c34,#1e2128);border-color:#3d434e;color:#f4f5f7}\n.division-card:hover{background:linear-gradient(145deg,#3b2228,#24242b);border-color:#ed3948;box-shadow:0 14px 28px #0006;transform:translateY(-5px)}\n.division-name{color:#f5f6f8}.division-hint{color:#9ca5b2}.division-icon{background:#363b45}\n.division-card:hover .division-icon{background:#5b252d}\n.division-card:has(input:checked){background:linear-gradient(145deg,#4a2028,#292129);border-color:#ff4655;box-shadow:0 0 0 1px #e21b2d55,0 10px 24px #0005}\n.division-count{background:#4a2028;border-color:#87313a;color:#ff9ba3}\n.login-content{max-width:560px;margin:5vh auto;padding:20px}\n.login-content:before{content:"";display:block;height:145px;margin:0 0 20px;border-radius:18px;background:linear-gradient(90deg,#0d0f13bb,#1b1012aa),url(\'/factory-bg.jpg\') center 48%/cover;box-shadow:0 15px 45px #0008;border:1px solid #443238}\n.login-content .panel{backdrop-filter:blur(8px)}\n@media(max-width:950px){.sidebar{border-right:0;border-bottom:3px solid #d71920}.content{padding:18px}.login-content{margin:2vh auto}}\n@media(prefers-reduced-motion:reduce){*{transition:none!important;scroll-behavior:auto!important}}\n'


def page(title,body,u=None,csrf='',message=''):
 nav=''
 if u:
  with con() as db:unread=db.execute('SELECT count(*) FROM notices WHERE uid=? AND seen=0',(u['id'],)).fetchone()[0]
  links=[('/','▦ Dashboard'),('/new','＋ New report'),('/notices','🔔 Notifications'+(f' <span class="pill red">{unread}</span>' if unread else '')),('/monthly','◉ Monthly tracking')]
  if u['role'] in ('manager','admin','safety'):links.extend([('/staff','♙ Staff directory'),('/compliance','▤ Staff compliance'),('/downloads','⇩ Downloads')])
  if u['role']=='admin':links.append(('/change-password','⚙ Admin password'))
  nav=f'<aside class="sidebar"><div class="brand"><img src="/ceat-logo.png" alt="CEAT logo"><small>WORKPLACE SAFETY</small></div><div class="identity">{esc(u["name"])}<small>{esc(u["role"].title())} · {esc(u["division"])}</small></div><nav>'+''.join(f'<a href="{url}">{label}</a>' for url,label in links)+'</nav>'+f'<form method="post" action="/logout"><input type="hidden" name="csrf" value="{esc(csrf)}"><button class="navlogout">Sign out →</button></form></aside>'
 return f'<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{esc(title)} · CEAT Safety</title><link rel="stylesheet" href="/style.css"></head><body>{nav}<main class="{"content" if u else "login-content"}">{f"<div class=flash>{esc(message)}</div>" if message else ""}{body}</main></body></html>'
def hidden(csrf):return f'<input type="hidden" name="csrf" value="{esc(csrf)}">'
def log(c,uid,event,rid=None):c.execute('INSERT INTO audit(uid,event,rid,created) VALUES(?,?,?,?)',(uid,event,rid,today()))
def notify(c,uid,rid,txt):
 if uid:c.execute('INSERT INTO notices(uid,rid,message,created) VALUES(?,?,?,?)',(uid,rid,txt,today()))
def managers(c,division):return c.execute("SELECT id FROM users WHERE active=1 AND (role IN ('admin','safety') OR (role='manager' AND division=?))",(division,)).fetchall()
def manage(u,r):return u['role'] in ('admin','safety') or (u['role']=='manager' and u['division']==(r['responsible_division'] or r['division']))
def view(u,r):return manage(u,r) or r['reporter']==u['id'] or r['assignee']==u['id'] or u['division']==(r['responsible_division'] or r['division'])
def xlsx_bytes(headers,rows):
 # Real XLSX Office Open XML, generated with only Python standard library.
 allrows=[headers]+[[str(v) if v is not None else '' for v in row] for row in rows]
 cells=[]
 for ri,row in enumerate(allrows,1):
  vals=[]
  for ci,v in enumerate(row):
   n=ci+1;letter=''
   while n: n,rem=divmod(n-1,26);letter=chr(65+rem)+letter
   # prevent formula injection in exports
   s=str(v)
   if s and s[0] in '=+-@\t\r':s="'"+s
   vals.append(f'<c r="{letter}{ri}" t="inlineStr"><is><t>{xml_escape(s)}</t></is></c>')
  cells.append(f'<row r="{ri}">'+''.join(vals)+'</row>')
 worksheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>'+''.join(cells)+'</sheetData></worksheet>'
 chunks={'[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>', '_rels/.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', 'xl/workbook.xml':'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Safety Reports" sheetId="1" r:id="rId1"/></sheets></workbook>', 'xl/_rels/workbook.xml.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>', 'xl/worksheets/sheet1.xml':worksheet}
 out=io.BytesIO()
 with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
  for k,v in chunks.items():z.writestr(k,'<?xml version="1.0" encoding="UTF-8"?>'+v if k!='xl/worksheets/sheet1.xml' else v)
 return out.getvalue()

class Handler(BaseHTTPRequestHandler):
 def _send(self,body,code=200,content_type='text/html; charset=utf-8',headers=None):
  if isinstance(body,str):body=body.encode('utf-8')
  self.send_response(code);self.send_header('Content-Type',content_type);self.send_header('Content-Length',str(len(body)));self.send_header('X-Content-Type-Options','nosniff');self.send_header('X-Frame-Options','DENY');self.send_header('Referrer-Policy','same-origin');self.send_header('Cache-Control','no-store');self.send_header('Content-Security-Policy',"default-src 'self'; style-src 'self'; img-src 'self' data:; form-action 'self'; base-uri 'none'")
  for k,v in (headers or {}).items():self.send_header(k,v)
  self.end_headers();self.wfile.write(body)
 def redirect(self,url,cookie=None):self._send('',303,headers={'Location':url,**({'Set-Cookie':cookie} if cookie else {})})
 def error(self,status,msg):self._send(page(str(status),f'<section class="panel"><h1>{status}</h1><p>{esc(msg)}</p><a href="/">Back to dashboard</a></section>',self.u,self.csrf),status)
 def get_form(self):
  ln=int(self.headers.get('Content-Length','0'))
  if ln<0 or ln>MAX_BODY:raise ValueError('File or form is too large (max 9 MB).')
  raw=self.rfile.read(ln);ct=self.headers.get('Content-Type','')
  if ct.startswith('multipart/form-data'):
   boundary=re.search(r'boundary=(?:"([^"]+)"|([^;]+))',ct)
   if not boundary:raise ValueError('Invalid file upload')
   msg=BytesParser(policy=default).parsebytes(b'MIME-Version: 1.0\r\nContent-Type: '+ct.encode()+b'\r\n\r\n'+raw)
   data={};files={}
   for part in msg.iter_parts():
    name=part.get_param('name',header='content-disposition')
    if not name:continue
    fn=part.get_filename()
    if fn:
     payload=part.get_payload(decode=True)
     if len(payload)>7*1024*1024:raise ValueError('Photo must be smaller than 7 MB')
     files[name]=(fn,payload)
    else:data[name]=part.get_content()
   return data,files
  return {k:v[-1] for k,v in urllib.parse.parse_qs(raw.decode('utf-8'),keep_blank_values=True).items()},{}
 def photo(self,files,key):
  if key not in files:return None
  fn,b=files[key];ext=Path(fn).suffix.lower()
  good=(ext in ('.jpg','.jpeg') and b.startswith(b'\xff\xd8\xff')) or (ext=='.png' and b.startswith(b'\x89PNG\r\n\x1a\n')) or (ext=='.webp' and b.startswith(b'RIFF') and b[8:12]==b'WEBP')
  if not good:raise ValueError('Only valid JPG, PNG, WEBP images accepted.')
  name=secrets.token_hex(16)+ext;(UPLOAD/name).write_bytes(b);return name
 def auth(self):
  self.u=None;self.csrf='';self.tokenhash=None
  ck=SimpleCookie()
  try:ck.load(self.headers.get('Cookie',''))
  except:pass
  token=ck['safetrack'].value if 'safetrack' in ck else ''
  if not token:return
  th=hashlib.sha256(token.encode()).hexdigest()
  with con() as c:
   x=c.execute('SELECT s.csrf,u.* FROM sessions s JOIN users u ON u.id=s.uid WHERE s.tokenhash=? AND s.expires>? AND u.active=1',(th,today())).fetchone()
  if x:self.u=x;self.csrf=x['csrf'];self.tokenhash=th
 def go(self,method):
  self.auth();p=urllib.parse.urlsplit(self.path).path;q=urllib.parse.parse_qs(urllib.parse.urlsplit(self.path).query)
  if p=='/style.css' and method=='GET':return self._send(CSS,content_type='text/css; charset=utf-8')
  if p=='/factory-bg.jpg' and method=='GET':
   img=ROOT/'factory-bg.jpg'
   return self._send(img.read_bytes(),content_type='image/jpeg') if img.exists() else self._send('',404)
  if p=='/ceat-logo.png' and method=='GET':
   logo=ROOT/'ceat-logo.png'
   return self._send(logo.read_bytes(),content_type='image/png') if logo.exists() else self._send('',404)
  if p=='/ceat-logo.svg' and method=='GET':
   logo=ROOT/'ceat-logo.svg'
   return self._send(logo.read_bytes(),content_type='image/svg+xml') if logo.exists() else self._send('',404)
  if p=='/health' and method=='GET':return self._send('{"status":"ok"}',content_type='application/json')
  if p=='/login':
   if method=='GET':return self._send(page('Login','<section class="panel"><div class="eyebrow">CEAT KELANI HOLDINGS · WORKPLACE SAFETY</div><h1 style="color:#d71920">CEAT Safety</h1><p class="muted">Sign in using your EPF-based username and password.</p><form method="post">'+field('Employee username (EPF digits)','code',required=True)+field('Password','password',typ='password',required=True)+'<button class="primary">Sign in</button></form></section>'))
   data,_=self.get_form()
   with con() as c:x=c.execute("SELECT * FROM users WHERE (code=? OR (epf=? AND epf!='')) AND active=1",(data.get('code','').strip(),data.get('code','').strip())).fetchone()
   if x and hmac.compare_digest(hp(data.get('password',''),x['salt']),x['pwhash']):
    token=secrets.token_urlsafe(40);csrf=secrets.token_urlsafe(24);exp=(datetime.datetime.now()+datetime.timedelta(hours=12)).strftime('%Y-%m-%d %H:%M:%S')
    with con() as c:c.execute('INSERT INTO sessions VALUES(?,?,?,?)',(hashlib.sha256(token.encode()).hexdigest(),x['id'],csrf,exp))
    cookie=f'safetrack={token}; HttpOnly; SameSite=Lax; Path=/'+('; Secure' if os.environ.get('SAFETRACK_HTTPS')=='1' else '')
    return self.redirect('/',cookie)
   return self._send(page('Login','<section class="panel"><h1>Sign in</h1><div class="warn">Invalid employee code or password.</div><form method="post">'+field('Employee username (EPF digits)','code')+field('Password','password',typ='password')+'<button class="primary">Try again</button></form></section>'),401)
  if not self.u:return self.redirect('/login')
  u=self.u;uid=u['id'];cs=self.csrf
  if method=='POST':
   data,files=self.get_form()
   if not hmac.compare_digest(data.get('csrf',''),cs):return self.error(403,'Invalid CSRF token. Refresh page.')
  else:data={};files={}
  if p=='/logout' and method=='POST':
   with con() as c:c.execute('DELETE FROM sessions WHERE tokenhash=?',(self.tokenhash,))
   return self.redirect('/login','safetrack=; Max-Age=0; Path=/; HttpOnly; SameSite=Lax')
  if p=='/change-password':
   if u['role']!='admin':return self.error(403,'Only the Safety Administrator can change account passwords. Contact your administrator.')
   if method=='GET':
    body='<div class="eyebrow">ADMIN ACCOUNT SECURITY</div><h1>Change administrator password</h1><section class="panel"><p class="warn">This page is available only to the Safety Administrator. All passwords must contain exactly 8 digits.</p><form method="post">'+hidden(cs)+field('Current password','current_password',typ='password',required=True)+field('New password (exactly 8 digits)','new_password',typ='password',required=True)+field('Confirm new password','confirm_password',typ='password',required=True)+'<button class="primary">Update administrator password</button></form></section>'
    return self._send(page('Admin password',body,u,cs))
   current=data.get('current_password','');pw=data.get('new_password','');confirm=data.get('confirm_password','')
   if not hmac.compare_digest(hp(current,u['salt']),u['pwhash']):return self.error(403,'Current password is incorrect.')
   if not valid_password(pw) or pw!=confirm:return self.error(400,'Password must be exactly 8 digits and both entries must match.')
   if hmac.compare_digest(hp(pw,u['salt']),u['pwhash']):return self.error(400,'Choose a new password different from your current password.')
   with con() as c:
    set_pw(c,uid,pw);c.execute('DELETE FROM sessions WHERE uid=? AND tokenhash!=?',(uid,self.tokenhash))
    log(c,uid,'Changed administrator password')
   return self.redirect('/')
  if p=='/' and method=='GET':
   where='1=1';args=[]
   if u['role']=='manager':where='COALESCE(NULLIF(r.responsible_division,\'\'),r.division)=?';args=[u['division']]
   elif u['role'] not in ('admin','safety'):where='(r.reporter=? OR r.assignee=? OR COALESCE(NULLIF(r.responsible_division,\'\'),r.division)=?)';args=[uid,uid,u['division']]
   with con() as c:
    rows=c.execute(f'SELECT r.*,p.name reporter_name,a.name assigned_name FROM reports r JOIN users p ON p.id=r.reporter LEFT JOIN users a ON a.id=r.assignee WHERE {where} ORDER BY r.id DESC LIMIT 250',args).fetchall()
    totals=c.execute(f"SELECT count(*) total, sum(CASE WHEN status='Closed' THEN 1 ELSE 0 END) closed, sum(CASE WHEN status!='Closed' AND deadline<date('now') THEN 1 ELSE 0 END) overdue FROM reports r WHERE {where}",args).fetchone()
    divs=[x['division'] for x in c.execute('SELECT DISTINCT division FROM users ORDER BY division').fetchall()]
   stats=[('Total reports',totals['total'] or 0),('Open / pending',(totals['total'] or 0)-(totals['closed'] or 0)),('Closed',totals['closed'] or 0),('Overdue',totals['overdue'] or 0)]
   cards='<div class="cards">'+''.join(f'<div class="stat"><small>{esc(k)}</small><b>{n}</b></div>' for k,n in stats)+'</div>'
   exp=''
   if u['role'] in ('admin','safety','manager'):
    exp='<section class="panel"><h2>Department-wise Excel downloads</h2><p class="muted">Download reports using the original department Excel templates.</p><a class="primary" href="/downloads">Open Download Reports</a></section>'
   t='<div class="tablewrap"><table><tr><th>ID</th><th>Date</th><th>Division</th><th>Type</th><th>Observation</th><th>Priority</th><th>Assigned</th><th>Status</th></tr>'+''.join(f'<tr><td><a href="/reports/{r["id"]}">#{r["id"]:04d}</a></td><td>{esc(r["created"][:10])}</td><td>{esc(r["responsible_division"] or r["division"])}</td><td>{esc(r["kind"])}</td><td>{esc(r["location"])}</td><td>{esc(r["priority"])}</td><td>{esc(r["assigned_name"] or "Unassigned")}</td><td>{badge(r["status"])}</td></tr>' for r in rows)+'</table></div>'
   return self._send(page('Dashboard','<div class="top"><div><div class="eyebrow">SAFETY OPERATIONS</div><h1>Safety Dashboard</h1><p class="muted">Safety observations and corrective actions</p></div><a class="primary" href="/new">＋ New report</a></div>'+cards+exp+'<section class="panel"><h2>Latest reports</h2>'+t+'</section>',u,cs))
  if p=='/new':
   if method=='GET':
    divisions=[d['label'] for d in DEPARTMENTS]
    div_icons={'Division A':'🏭','General':'⚙️','TB - Bead & Sq Bias':'🛞','Bias Curing':'🔥','TBR':'🚚','TW':'🏍️','Radial':'⭕','Engineering':'🛠️','LAB':'🧪','Technical':'🔧','QA':'✓','HR':'👥','Stores':'📦','Warehouse (WH)':'🏬'}
    current=next((d['label'] for d in DEPARTMENTS if division_access(d,u['division'])), 'Division A')
    division_cards='<div class="division-grid" role="radiogroup" aria-label="Responsible division">'+''.join('<label class="division-card"><input type="radio" name="responsible_division" value="'+esc(d['label'])+'" '+('checked' if d['label']==current else '')+' required><span class="division-icon">'+div_icons.get(d['label'],'🏢')+'</span><span class="division-name">'+esc(d['label'])+'</span><span class="division-hint">Select department</span><span class="division-check">✓</span></label>' for d in DEPARTMENTS)+'</div>'
    form='<div class="top"><div><div class="eyebrow">CEAT WORKPLACE SAFETY</div><h1>Report a Safety Hazard</h1><p class="pagelead">Choose the department responsible for addressing this safety observation.</p></div></div><section class="panel report-panel"><div class="warn">For immediate danger, inform your supervisor / EHS team immediately. This portal does not replace emergency escalation.</div><form method="post" enctype="multipart/form-data">'+hidden(cs)+'<div class="twocol"><label>Report type<select name="kind">'+opt([(x,x) for x in ('UC','UA','Near Miss')])+'</select></label><label>Priority (preliminary)<select name="priority">'+opt([(x,x) for x in ('Low','Medium','High','Critical')],'Medium')+'</select></label></div>'+field('Location / machine / section','location',required=True)+'<div class="twocol">'+field('Process','process')+'</div><div class="division-section"><div class="division-heading"><div><h2>Responsible division</h2><p class="muted">Select the responsible department.</p></div><span class="division-count">15 DEPARTMENTS</span></div>'+division_cards+'</div><label>Observation description<textarea name="description" maxlength="5000" required></textarea></label><label>Photo evidence (optional JPG, PNG, WEBP; max 7 MB)<input type="file" name="photo" accept="image/jpeg,image/png,image/webp"></label><button class="primary">Submit safety report</button></form></section>'
    return self._send(page('New report',form,u,cs))
   kind=data.get('kind');priority=data.get('priority');loc=data.get('location','').strip();desc=data.get('description','').strip();responsible_div=data.get('responsible_division','').strip()
   if kind not in ('UC','UA','Near Miss') or priority not in ('Low','Medium','High','Critical') or not loc or not desc or not responsible_div or len(desc)>5000:return self.error(400,'Invalid report fields')
   with con() as c:
    valid_divisions=set(x['label'] for x in DEPARTMENTS)
    if responsible_div not in valid_divisions:return self.error(400,'Select a valid responsible division.')
   photo=self.photo(files,'photo')
   with con() as c:
    rid=c.execute('INSERT INTO reports(reporter,division,responsible_division,kind,location,process,operation,description,priority,photo,created) VALUES(?,?,?,?,?,?,?,?,?,?,?)',(uid,u['division'],responsible_div,kind,loc,data.get('process','')[:200],'',desc,priority,photo,today())).lastrowid
    log(c,uid,'Created report',rid)
    recipients=set(x['id'] for x in c.execute("SELECT id FROM users WHERE active=1 AND role IN ('admin','safety')").fetchall())
    all_users=c.execute("SELECT id,division FROM users WHERE active=1").fetchall()
    recipients.update(x['id'] for x in all_users if x['division'].casefold()==responsible_div.casefold())
    for dept in DEPARTMENTS:
     if dept['label'].casefold()==responsible_div.casefold():
      recipients.update(x['id'] for x in all_users if division_access(dept,x['division']))
    for recipient_id in recipients:notify(c,recipient_id,rid,f'New {priority} {kind} report #{rid} assigned to {responsible_div}')
   return self.redirect('/reports/'+str(rid))
  match=re.fullmatch(r'/reports/(\d+)(?:/(assign|action|verify))?',p)
  if match:
   rid=int(match.group(1));task=match.group(2)
   with con() as c:
    r=c.execute('SELECT r.*,p.name reporter_name,p.code reporter_code,p.epf reporter_epf,p.designation reporter_designation,a.name assigned_name FROM reports r JOIN users p ON p.id=r.reporter LEFT JOIN users a ON a.id=r.assignee WHERE r.id=?',(rid,)).fetchone()
    if not r:return self.error(404,'Report not found')
    if not view(u,r):return self.error(403,'No permission for this report')
    if method=='POST':
     if task=='assign':
      if not manage(u,r) or r['status']=='Closed':return self.error(403,'Cannot assign')
      try:ass=int(data.get('assignee','0'));deadline=datetime.date.fromisoformat(data.get('deadline',''))
      except:return self.error(400,'Choose a responsible person and valid deadline')
      person=c.execute('SELECT * FROM users WHERE id=? AND active=1 AND division=?',(ass,r['responsible_division'] or r['division'])).fetchone()
      if not person:return self.error(400,'Responsible person must belong to this division')
      c.execute('UPDATE reports SET assignee=?,deadline=?,status=? WHERE id=?',(ass,str(deadline),'Assigned',rid));log(c,uid,'Assigned action to '+person['code'],rid);notify(c,ass,rid,f'Report #{rid} assigned to you, deadline {deadline}')
     elif task=='action':
      if not(manage(u,r) or r['assignee']==uid) or r['status']=='Closed':return self.error(403,'Cannot update this action')
      typ=data.get('kind');note=data.get('note','').strip()
      if typ not in ('Progress','Corrective Action') or not note:return self.error(400,'Provide action type and notes')
      photo=self.photo(files,'photo')
      c.execute('INSERT INTO actions(rid,actor,kind,note,photo,created) VALUES(?,?,?,?,?,?)',(rid,uid,typ,note[:3000],photo,today()))
      state='Pending Verification' if typ=='Corrective Action' else 'In Progress'
      c.execute('UPDATE reports SET status=? WHERE id=?',(state,rid));log(c,uid,typ,rid)
      if typ=='Corrective Action':
       for x in managers(c,r['responsible_division'] or r['division']):notify(c,x['id'],rid,f'Corrective action submitted for verification on report #{rid}')
     elif task=='verify':
      if not manage(u,r) or r['status']!='Pending Verification':return self.error(403,'Verification not permitted')
      verdict=data.get('verdict');note=data.get('note','').strip()
      if verdict not in ('Closed','Reopened') or not note:return self.error(400,'Invalid verification')
      c.execute('INSERT INTO actions(rid,actor,kind,note,created) VALUES(?,?,?,?,?)',(rid,uid,'Verification: '+verdict,note[:3000],today()))
      c.execute('UPDATE reports SET status=?,closed=? WHERE id=?',(verdict,today() if verdict=='Closed' else None,rid));log(c,uid,'Verified: '+verdict,rid)
      for person in set([r['reporter'],r['assignee']]):notify(c,person,rid,f'Report #{rid} verification: {verdict}')
     else:return self.error(404,'Unknown action')
     return self.redirect('/reports/'+str(rid))
    acts=c.execute('SELECT a.*,u.name actor_name FROM actions a JOIN users u ON u.id=a.actor WHERE rid=? ORDER BY a.id',(rid,)).fetchall()
    people=c.execute('SELECT id,name,code FROM users WHERE active=1 AND division=? ORDER BY name',(r['responsible_division'] or r['division'],)).fetchall()
   dl=f'<dl><dt>Observation</dt><dd>{esc(r["description"])}</dd><dt>Priority</dt><dd>{esc(r["priority"])}</dd><dt>Reported by</dt><dd>{esc(r["reporter_name"])} ({esc(r["reporter_code"])}) · EPF: {esc(r["reporter_epf"] or "Not recorded")} · {esc(r["reporter_designation"])}</dd><dt>Process</dt><dd>{esc(r["process"])}</dd><dt>Responsible division</dt><dd>{esc(r["responsible_division"] or r["division"])}</dd><dt>Assignee</dt><dd>{esc(r["assigned_name"] or "Not assigned")}</dd><dt>Deadline</dt><dd>{esc(r["deadline"] or "Not set")}</dd><dt>Closed</dt><dd>{esc(r["closed"] or "—")}</dd></dl>'
   if r['photo']:dl+=f'<h3>Before photo</h3><img class="evidence" src="/photos/{esc(r["photo"])}">'
   blocks='<section class="panel"><h2>Observation details</h2>'+dl+'</section>'
   if manage(u,r) and r['status']!='Closed':
    blocks+='<section class="panel"><h2>Assign responsible person</h2><form method="post" action="/reports/'+str(rid)+'/assign">'+hidden(cs)+'<label>Assignee<select name="assignee" required>'+opt([(x['id'],x['name']+' · '+x['code']) for x in people],r['assignee'])+'</select></label>'+field('Deadline','deadline',r['deadline'],typ='date',required=True)+'<button class="primary">Assign and notify</button></form></section>'
   if (manage(u,r) or r['assignee']==uid) and r['status']!='Closed':
    blocks+='<section class="panel"><h2>Corrective action update</h2><form method="post" enctype="multipart/form-data" action="/reports/'+str(rid)+'/action">'+hidden(cs)+'<label>Action type<select name="kind">'+opt([(x,x) for x in ('Progress','Corrective Action')])+'</select></label><label>Work completed<textarea name="note" required maxlength="3000"></textarea></label><label>After photo<input type="file" name="photo" accept="image/jpeg,image/png,image/webp"></label><button class="primary">Save action</button></form></section>'
   if manage(u,r) and r['status']=='Pending Verification':
    blocks+='<section class="panel"><h2>Review & verify</h2><form method="post" action="/reports/'+str(rid)+'/verify">'+hidden(cs)+'<label>Decision<select name="verdict">'+opt([(x,x) for x in ('Closed','Reopened')])+'</select></label><label>Verification note<textarea name="note" required></textarea></label><button class="primary">Save verification</button></form></section>'
   hist='<section class="panel"><h2>Corrective action history</h2><div class="timeline">'+''.join(f'<div class="event"><b>{esc(a["kind"])}</b><small>{esc(a["created"])} by {esc(a["actor_name"])}</small><p>{esc(a["note"])}</p>'+(f'<img class="evidence" src="/photos/{esc(a["photo"])}">' if a['photo'] else '')+'</div>' for a in acts)+'</div></section>'
   return self._send(page('Report '+str(rid),f'<div class="top"><div><div class="eyebrow">SAFETY REPORT #{rid:04d}</div><h1>{esc(r["kind"])} · {esc(r["location"])}</h1><p class="muted">{esc(r["division"])} · {esc(r["created"])}</p></div>{badge(r["status"])}</div><div class="twocol"><div>{blocks}</div><div>{hist}</div></div>',u,cs))
  pm=re.fullmatch(r'/photos/([a-f0-9]{32}\.(?:jpg|jpeg|png|webp))',p)
  if pm and method=='GET':
   name=pm.group(1)
   with con() as c:r=c.execute('SELECT * FROM reports WHERE photo=? UNION SELECT r.* FROM reports r JOIN actions a ON a.rid=r.id WHERE a.photo=? LIMIT 1',(name,name)).fetchone()
   if not r or not view(u,r):return self.error(403,'Photo restricted')
   f=UPLOAD/name
   if not f.exists():return self.error(404,'Photo missing')
   return self._send(f.read_bytes(),content_type=mimetypes.guess_type(name)[0] or 'application/octet-stream')
  if p=='/notices' and method=='GET':
   with con() as c:
    items=c.execute('SELECT * FROM notices WHERE uid=? ORDER BY id DESC LIMIT 150',(uid,)).fetchall()
    body='<div class="eyebrow">ALERT CENTRE</div><h1>Notifications</h1><p class="pagelead">Unread alerts are highlighted in red.</p><section class="panel">'+(''.join(f'<div class="{"notification-unread" if not x["seen"] else "notification-read"}">{"<span class=notification-dot></span>" if not x["seen"] else ""}<a href="/reports/{x["rid"]}"><strong>Report #{x["rid"]}</strong></a> · {esc(x["message"])}<small class="muted" style="display:block">{esc(x["created"])} · {"Unread" if not x["seen"] else "Read"}</small></div>' for x in items) or '<p class="muted">No notifications yet.</p>')+'</section>'
    c.execute('UPDATE notices SET seen=1 WHERE uid=?',(uid,))
   return self._send(page('Notifications',body,u,cs))
  if p=='/monthly':
   m=month()
   if method=='POST':
    if data.get('confirm')!='yes':return self.error(400,'Confirmation required')
    with con() as c:c.execute('INSERT OR IGNORE INTO confirm(uid,month,created) VALUES(?,?,?)',(uid,m,today()))
    return self.redirect('/monthly')
   with con() as c:
    done=c.execute('SELECT 1 FROM confirm WHERE uid=? AND month=? UNION SELECT 1 FROM reports WHERE reporter=? AND substr(created,1,7)=? LIMIT 1',(uid,m,uid,m)).fetchone()
    sql='SELECT u.division,count(*) total,sum(CASE WHEN EXISTS(SELECT 1 FROM confirm t WHERE t.uid=u.id AND t.month=?) OR EXISTS(SELECT 1 FROM reports r WHERE r.reporter=u.id AND substr(r.created,1,7)=?) THEN 1 ELSE 0 END) completed FROM users u WHERE u.active=1';args=[m,m]
    if u['role']=='manager':sql+=' AND u.division=?';args.append(u['division'])
    sql+=' GROUP BY u.division ORDER BY u.division'
    summary=c.execute(sql,args).fetchall() if u['role'] in ('manager','safety','admin') else []
   body=f'<h1>Monthly reporting · {m}</h1><section class="panel"><h2>Your status</h2><p>{"Submitted / confirmed for this month" if done else "Not yet submitted"}</p>'
   if not done:body+='<form method="post">'+hidden(cs)+'<input type="hidden" name="confirm" value="yes"><p>Confirm no hazard observed this month, only if true. Otherwise submit a safety report.</p><button class="primary">Confirm no hazard observed</button></form>'
   body+='</section><section class="panel"><h2>Division compliance</h2><div class="tablewrap"><table><tr><th>Division</th><th>Staff</th><th>Participated</th><th>Percentage</th></tr>'+''.join(f'<tr><td>{esc(x["division"])}</td><td>{x["total"]}</td><td>{x["completed"]}</td><td>{round(100*x["completed"]/x["total"])}%</td></tr>' for x in summary)+'</table></div></section>'
   return self._send(page('Monthly',body,u,cs))
  if p=='/compliance' and method=='GET':
   if u['role'] not in ('admin','safety','manager'):return self.error(403,'Not authorized')
   m=q.get('month',[month()])[0];div=q.get('division',[''])[0]
   if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',m):return self.error(400,'Invalid month')
   if u['role']=='manager':div=u['division']
   with con() as c:
    divs=[x[0] for x in c.execute('SELECT DISTINCT division FROM users ORDER BY division')]
    sql="""SELECT u.id,u.code,u.epf,u.name,u.designation,u.division,u.role,u.active,
     (SELECT count(*) FROM reports r WHERE r.reporter=u.id AND substr(r.created,1,7)=?) AS reports_count,
     (SELECT count(*) FROM confirm cf WHERE cf.uid=u.id AND cf.month=?) AS confirmations
     FROM users u WHERE u.active=1"""
    params=[m,m]
    if div:sql+=' AND u.division=?';params.append(div)
    rows=c.execute(sql+' ORDER BY u.division,u.name',params).fetchall()
   completed=sum(bool(x['reports_count'] or x['confirmations']) for x in rows)
   form='<form class="filterbar" method="get"><label>Month<input type="month" name="month" value="'+esc(m)+'"></label><label>Division<select name="division"><option value="">All divisions</option>'+opt([(d,d) for d in divs],div)+'</select></label><button class="primary">Apply filters</button></form>'
   if u['role']=='manager':form='<form class="filterbar" method="get"><label>Month<input type="month" name="month" value="'+esc(m)+'"></label><button class="primary">Apply filters</button></form>'
   stats=f'<div class="cards"><div class="stat"><small>Active staff</small><b>{len(rows)}</b></div><div class="stat"><small>Participated</small><b>{completed}</b></div><div class="stat"><small>Not participated</small><b>{len(rows)-completed}</b></div><div class="stat"><small>Compliance</small><b>{round(100*completed/len(rows)) if rows else 0}%</b></div></div>'
   lines=''.join(f'<tr><td><a href="/staff/{x["id"]}">{esc(x["name"])}</a><div class="muted small">{esc(x["code"])}</div></td><td>{esc(x["epf"] or "—")}</td><td>{esc(x["division"])}</td><td>{x["reports_count"]}</td><td>{"Yes" if x["confirmations"] else "No"}</td><td>{badge("Closed" if x["reports_count"] or x["confirmations"] else "Open")}</td></tr>' for x in rows)
   exporturl='/export?type=compliance&month='+urllib.parse.quote(m)+'&division='+urllib.parse.quote(div)
   body=f'<div class="top"><div><div class="eyebrow">STAFF PARTICIPATION</div><h1>Monthly Compliance</h1><p class="pagelead">Check exactly who reported or confirmed no hazards.</p></div><a class="primary" href="{exporturl}">⬇ Export Excel</a></div>'+stats+'<section class="panel"><h2>Filter compliance</h2>'+form+'</section><section class="panel"><h2>Staff status · '+esc(m)+'</h2><div class="tablewrap"><table><tr><th>Name / Code</th><th>EPF</th><th>Division</th><th>Reports</th><th>No-hazard confirmation</th><th>Participated</th></tr>'+lines+'</table></div></section>'
   return self._send(page('Compliance',body,u,cs))
  if p=='/downloads' and method=='GET':
   if u['role'] not in ('admin','safety','manager'):return self.error(403,'Not authorized')
   available=[d for d in DEPARTMENTS if d.get('file')]
   if u['role']=='manager':
    available=[d for d in DEPARTMENTS if d.get('file') and division_access(d,u['division'])]
   if not available:
    body='<div class="eyebrow">DATA EXPORT CENTRE</div><h1>Download Reports</h1><section class="panel"><p class="muted">No department Excel template is mapped to your division yet. Please contact the Safety Administrator.</p></section>'
    return self._send(page('Downloads',body,u,cs))
   options=opt([(d['key'],d['label']) for d in available],available[0]['key'])
   default_month=month()
   body='<div class="eyebrow">DATA EXPORT CENTRE</div><h1>Download Reports</h1><p class="pagelead">Select a department and month. Only reports saved in this system for the selected division and month will be exported. Existing Excel rows are not imported or reused.</p><section class="panel"><h2>Monthly department safety workbook</h2><p class="muted">The original Excel files are used only as blank templates. Only matching reports submitted through this system are added.</p><form action="/export" method="get"><div class="twocol"><label>Department / Division<select name="department">'+options+'</select></label><label>Month<input type="month" name="month" value="'+esc(default_month)+'" required></label></div><button class="primary">⬇ Download Monthly Excel (.xlsx)</button></form></section>'
   return self._send(page('Downloads',body,u,cs))
  profile=re.fullmatch(r'/staff/(\d+)',p)
  if profile and method=='GET':
   if u['role'] not in ('admin','safety','manager'):return self.error(403,'Not authorized')
   target=int(profile.group(1))
   with con() as c:
    person=c.execute('SELECT * FROM users WHERE id=?',(target,)).fetchone()
    if not person or (u['role']=='manager' and person['division']!=u['division']):return self.error(403,'Not permitted')
    reports=c.execute('SELECT id,created,kind,location,status FROM reports WHERE reporter=? ORDER BY id DESC LIMIT 100',(target,)).fetchall()
   details='<dl>'+''.join(f'<dt>{esc(label)}</dt><dd>{esc(person[col] or "Not recorded")}</dd>' for label,col in [('Name','name'),('EPF number','epf'),('Employee code','code'),('Division','division'),('Designation','designation'),('Role','role'),('Email','email')])+'</dl>'
   tbl=''.join(f'<tr><td><a href="/reports/{r["id"]}">#{r["id"]:04d}</a></td><td>{esc(r["created"][:10])}</td><td>{esc(r["kind"])}</td><td>{esc(r["location"])}</td><td>{badge(r["status"])}</td></tr>' for r in reports)
   body=f'<div class="eyebrow">EMPLOYEE SAFETY PROFILE</div><h1>{esc(person["name"])}</h1><div class="twocol"><section class="panel"><h2>Employee details</h2>{details}</section><section class="panel"><h2>Reporting activity</h2><div class="stat"><small>Total reports submitted</small><b>{len(reports)}</b></div><p class="muted">Latest 100 reports shown.</p></section></div><section class="panel"><h2>Submitted reports</h2><div class="tablewrap"><table><tr><th>ID</th><th>Date</th><th>Type</th><th>Location</th><th>Status</th></tr>{tbl}</table></div></section>'
   return self._send(page('Staff profile',body,u,cs))
  if p=='/staff/create' and method=='POST':
   if u['role']!='admin':return self.error(403,'Admin only')
   epf=data.get('epf','').strip().upper();name=data.get('name','').strip();div=data.get('division','').strip();designation=data.get('designation','').strip();role=data.get('role','staff')
   digits=re.sub(r'\D','',epf);code=digits
   if not re.fullmatch(r'[A-Z]?\d{4}',epf) or not code or not name or not div or role not in ROLES:return self.error(400,'Enter an EPF number with exactly 4 digits after its optional first letter, full name, division and role.')
   pw=digits*2
   if not valid_password(pw):return self.error(400,'The EPF must contain exactly 4 digits to generate the initial 8-digit password.')
   with con() as c:
    if c.execute('SELECT 1 FROM users WHERE code=?',(code,)).fetchone():return self.error(409,'Employee code already exists')
    if epf and c.execute('SELECT 1 FROM users WHERE epf=?',(epf,)).fetchone():return self.error(409,'EPF number already exists')
    salt=secrets.token_hex(16)
    c.execute('INSERT INTO users(code,epf,name,designation,division,role,email,salt,pwhash,must_change_password) VALUES(?,?,?,?,?,?,?,?,?,0)',(code,epf,name,designation,div,role,data.get('email','')[:200],salt,hp(pw,salt)))
    log(c,uid,'Created employee '+code)
   return self.redirect('/staff')
  if p=='/staff':
   if u['role'] not in ('manager','safety','admin'):return self.error(403,'Not authorized')
   with con() as c:
    sql='SELECT id,code,epf,name,designation,division,role,email,active FROM users';args=[]
    if u['role']=='manager':sql+=' WHERE division=?';args=[u['division']]
    rows=c.execute(sql+' ORDER BY division,name',args).fetchall()
   items=[]
   for x in rows:
    update=''
    if u['role']=='admin':
     update=f'<form method="post" action="/staff/{x["id"]}/update">'+hidden(cs)+'<input name="epf" placeholder="EPF number" value="'+esc(x['epf'])+'"><select name="role">'+opt([(z,z) for z in ROLES],x['role'])+'</select><input name="email" type="email" placeholder="Email" value="'+esc(x['email'])+'"><input name="password" type="password" placeholder="Set password (8 digits)"><button class="secondary">Save</button></form>'
    items.append(f'<tr><td>{esc(x["code"])}</td><td><a href="/staff/{x["id"]}">{esc(x["name"])}</a></td><td>{esc(x["epf"] or "—")}</td><td>{esc(x["division"])}</td><td>{esc(x["designation"])}</td><td>{esc(x["role"])}</td><td>{update}</td></tr>')
   create=''
   if u['role']=='admin':
    divisions=[d['label'] for d in DEPARTMENTS]
    div_select='<label>Department / Division<select name="division" required>'+opt([(d,d) for d in divisions])+'</select></label>'
    create='<section class="panel"><h2>＋ Create employee account</h2><p class="muted">Enter the employee details below. Username and initial password are generated from the EPF number.</p><form method="post" action="/staff/create">'+hidden(cs)+'<div class="twocol">'+field('Full name','name',required=True)+field('EPF number (e.g. K1434)','epf',required=True)+field('Designation','designation',required=True)+div_select+'<label>User role<select name="role">'+opt([(z,z) for z in ROLES])+'</select></label>'+field('Email address','email',typ='email')+'</div><div class="warn">Username = EPF digits only. Initial password = those 4 digits repeated twice (8 digits). Employees cannot change passwords themselves; only the administrator can reset them.</div><button class="primary">Create employee account</button></form></section>'
   body='<div class="eyebrow">EMPLOYEE MANAGEMENT</div><h1>Staff Directory</h1><p class="pagelead">Search and open an employee profile to view their safety reports. Admin can create users and update EPF numbers.</p>'+create+'<section class="panel"><h2>Employee list</h2><input id="staffsearch" placeholder="Search name, EPF, code or division" oninput="filterStaff(this.value)"><div class="tablewrap"><table><tr><th>Employee code</th><th>Name</th><th>EPF</th><th>Division</th><th>Designation</th><th>Role</th><th>Admin edits</th></tr><tbody id="staffrows">'+''.join(items)+'</tbody></table></div></section>'
   return self._send(page('Staff',body+'<script>function filterStaff(v){document.querySelectorAll("#staffrows tr").forEach(r=>r.style.display=r.innerText.toLowerCase().includes(v.toLowerCase())?"":"none")}</script>',u,cs))
  sm=re.fullmatch(r'/staff/(\d+)/update',p)
  if sm and method=='POST':
   if u['role']!='admin':return self.error(403,'Admin only')
   tid=int(sm.group(1));role=data.get('role');password=data.get('password','')
   if role not in ROLES:return self.error(400,'Bad role')
   if password and not valid_password(password):return self.error(400,'Password must be exactly 8 digits')
   with con() as c:
    if not c.execute('SELECT id FROM users WHERE id=?',(tid,)).fetchone():return self.error(404,'Staff missing')
    epf=data.get('epf','').strip()
    if len(epf)>50:return self.error(400,'EPF number too long')
    if epf and c.execute('SELECT 1 FROM users WHERE epf=? AND id!=?',(epf,tid)).fetchone():return self.error(409,'EPF already assigned to another user')
    c.execute('UPDATE users SET role=?,email=?,epf=? WHERE id=?',(role,data.get('email','')[:200],epf,tid))
    if password:set_pw(c,tid,password);c.execute('UPDATE users SET must_change_password=0 WHERE id=?',(tid,));c.execute('DELETE FROM sessions WHERE uid=?',(tid,))
    log(c,uid,'Updated user '+str(tid))
   return self.redirect('/staff')
  if p=='/export' and method=='GET':
   if u['role'] not in ('manager','safety','admin'):return self.error(403,'Not authorized')
   # Keep the existing monthly compliance export available from Staff Compliance.
   if q.get('type',[''])[0]=='compliance':
    m=q.get('month',[''])[0].strip() or month()
    div=q.get('division',[''])[0].strip()
    if u['role']=='manager':div=u['division']
    if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',m):return self.error(400,'Invalid month')
    sql="""SELECT u.code,u.epf,u.name,u.designation,u.division,
      (SELECT count(*) FROM reports r WHERE r.reporter=u.id AND substr(r.created,1,7)=?),
      (SELECT count(*) FROM confirm t WHERE t.uid=u.id AND t.month=?)
      FROM users u WHERE u.active=1"""
    args=[m,m]
    if div:sql+=' AND u.division=?';args.append(div)
    with con() as c:rows=c.execute(sql+' ORDER BY u.division,u.name',args).fetchall()
    export_rows=[(x[0],x[1],x[2],x[3],x[4],m,x[5], 'Yes' if x[6] else 'No','Yes' if x[5] or x[6] else 'No') for x in rows]
    payload=xlsx_bytes(['Employee Code','EPF','Name','Designation','Division','Month','Reports Submitted','No Hazard Confirmation','Participated'],export_rows)
    return self._send(payload,content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':f'attachment; filename="safety_staff_compliance_{m}.xlsx"'})
   department_key=q.get('department',[''])[0].strip()
   dept=next((d for d in DEPARTMENTS if d['key']==department_key),None)
   if not dept:return self.error(400,'Please select a valid department.')
   if u['role']=='manager' and not division_access(dept,u['division']):
    return self.error(403,'You can only download your own department workbook.')
   selected_month=q.get('month',[''])[0].strip()
   if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',selected_month):
    return self.error(400,'Please select a valid month for the department export.')
   sql='''SELECT r.id,r.created,r.division,r.responsible_division,r.kind,r.location,r.process,r.description,r.priority,r.deadline,r.status,r.closed,
     p.name AS reporter_name,p.code AS reporter_code,p.epf AS reporter_epf,a.name AS assigned_name,
     (SELECT ac.note FROM actions ac WHERE ac.rid=r.id ORDER BY ac.id DESC LIMIT 1) AS latest_action
     FROM reports r JOIN users p ON p.id=r.reporter LEFT JOIN users a ON a.id=r.assignee
     WHERE substr(r.created,1,7)=?
     ORDER BY r.id'''
   with con() as c:all_rows=[dict(x) for x in c.execute(sql,(selected_month,)).fetchall()]
   rows=[r for r in all_rows if division_matches(dept,r.get('responsible_division') or r['division'],r.get('process'),r.get('location'))]
   try:
    payload=create_department_workbook(dept,rows)
   except FileNotFoundError:
    logging.exception('Department template missing')
    return self.error(500,'The selected department template is missing. Please contact the administrator.')
   except Exception:
    logging.exception('Excel export failed')
    return self.error(500,'Could not create the Excel workbook. Please check the application log and template files.')
   safe_name=re.sub(r'[^A-Za-z0-9_-]+','_',dept['label']).strip('_')
   return self._send(payload,content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
     headers={'Content-Disposition':f'attachment; filename="SafeTrack_{safe_name}_{selected_month}.xlsx"'})
  return self.error(404,'Page not found')
 def do_GET(self):
  try:self.go('GET')
  except (ValueError,sqlite3.Error) as e:logging.exception('Request failed');self.error(400,str(e))
  except Exception:logging.exception('Request failed');self._send('Internal server error',500)
 def do_POST(self):
  try:self.go('POST')
  except (ValueError,sqlite3.Error) as e:logging.exception('Request failed');self.error(400,str(e))
  except Exception:logging.exception('Request failed');self._send('Internal server error',500)

def main():
 import argparse,getpass
 default_host=os.environ.get('HOST','0.0.0.0')
 default_port=int(os.environ.get('PORT','8091'))
 parser=argparse.ArgumentParser();parser.add_argument('command',nargs='?',default='serve',choices=['serve','setup']);parser.add_argument('--host',default=default_host);parser.add_argument('--port',type=int,default=default_port);args=parser.parse_args()
 init()
 with con() as c: migrated=c.execute("SELECT value FROM app_meta WHERE key='epf_login_migration_v1'").fetchone()
 if not migrated:
  import_staff()
  with con() as c:c.execute("INSERT OR REPLACE INTO app_meta(key,value) VALUES('epf_login_migration_v1','done')")
 with con() as c:
  if not c.execute("SELECT id FROM users WHERE role='admin'").fetchone():
   s=secrets.token_hex(16)
   aid=c.execute("INSERT INTO users(code,name,designation,division,role,salt,pwhash,active,must_change_password) VALUES(?,?,?,?,?,?,?,1,0)",
                 ('SAFETYADMIN','System Administrator','Administrator','SAFETY','admin',s,hp('12345678',s))).lastrowid
   set_pw(c,aid,'12345678')
 if args.command=='setup':
  print('Staff records imported/migrated from roster.');code=input('Admin employee code (new or existing): ').strip().upper()
  if not code:raise SystemExit('Employee code required')
  pw=getpass.getpass('Admin password (exactly 8 digits): ')
  if not valid_password(pw):raise SystemExit('Password must be exactly 8 digits')
  with con() as c:
   x=c.execute('SELECT * FROM users WHERE code=?',(code,)).fetchone()
   if x:aid=x['id']
   else:
    name=input('Admin name: ').strip() or 'System Administrator';salt=secrets.token_hex(16)
    aid=c.execute('INSERT INTO users(code,name,designation,division,role,salt,pwhash) VALUES(?,?,?,?,?,?,?)',(code,name,'Administrator','SAFETY','admin',salt,hp(pw,salt))).lastrowid
   c.execute('UPDATE users SET role=? WHERE id=?',('admin',aid));set_pw(c,aid,pw)
  print('Setup completed. Start with: python app.py serve')
  return
  
 server=ThreadingHTTPServer((args.host,args.port),Handler);print(f'SafeTrack running at http://{args.host}:{args.port} (Ctrl+C to stop)',flush=True)
 try:server.serve_forever()
 except KeyboardInterrupt:pass
 finally:server.server_close()
if __name__=='__main__':main()
