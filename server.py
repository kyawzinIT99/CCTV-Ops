#!/usr/bin/env python3
"""Local-first casino security operations console (camera registry and review workflow)."""
from __future__ import annotations

import argparse
import hashlib
import hmac
import http.cookies
import json
import os
import secrets
import sqlite3
import sys
import time
import uuid
import socket
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs
import getpass
import ipaddress
import io
import urllib.error
import urllib.request
from urllib.parse import quote
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("CASINO_OPS_DB", ROOT / "data" / "casino-ops.sqlite3"))
HOST = os.environ.get("CASINO_OPS_HOST", "127.0.0.1")
PORT = int(os.environ.get("CASINO_OPS_PORT", "8080"))
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-5.6-luna").strip() or "gpt-5.6-luna"
SESSION_SECONDS = 8 * 60 * 60
PBKDF2_ROUNDS = 600_000
RATE_LIMIT: dict[str, list[float]] = {}
PERMISSIONS = ("view_dashboard","manage_projects","view_devices","manage_devices","run_discovery","view_people","manage_people","view_movement","view_alerts","review_alerts","view_automation","view_audit","manage_users","manage_backups")
BACKUP_DIR = Path(os.environ.get("CASINO_OPS_BACKUP_DIR", str(DB_PATH.parent / "backups")))

SCHEMA = """
PRAGMA foreign_keys=ON;
CREATE TABLE IF NOT EXISTS projects (
  id INTEGER PRIMARY KEY,
  name TEXT NOT NULL UNIQUE,
  location TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY,
  username TEXT NOT NULL UNIQUE,
  password_salt BLOB NOT NULL,
  password_hash BLOB NOT NULL,
  permissions TEXT NOT NULL DEFAULT '[]',
  active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS sessions (
  token_hash TEXT PRIMARY KEY,
  user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  csrf_token TEXT NOT NULL,
  expires_at INTEGER NOT NULL,
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_sessions_expiry ON sessions(expires_at);
CREATE TABLE IF NOT EXISTS devices (
  id INTEGER PRIMARY KEY,
  project_id INTEGER REFERENCES projects(id),
  name TEXT NOT NULL,
  device_kind TEXT NOT NULL DEFAULT 'Unknown' CHECK(device_kind IN ('DVR','NVR','Camera','Unknown')),
  vendor_model TEXT NOT NULL DEFAULT '',
  host TEXT NOT NULL,
  port INTEGER NOT NULL CHECK(port BETWEEN 1 AND 65535),
  protocol TEXT NOT NULL CHECK(protocol IN ('RTSP','ONVIF','RTSP + ONVIF','Vendor connector')),
  stream_label TEXT NOT NULL DEFAULT '',
  coverage_role TEXT NOT NULL DEFAULT 'General' CHECK(coverage_role IN ('General','Entrance','Exit')),
  source_mode TEXT NOT NULL DEFAULT 'Recorder' CHECK(source_mode IN ('Recorder','Standalone camera')),
  monitoring_purposes TEXT NOT NULL DEFAULT '[]',
  status TEXT NOT NULL DEFAULT 'Unverified' CHECK(status IN ('Unverified','Needs verification','Configured')),
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS people (
  id INTEGER PRIMARY KEY,
  project_id INTEGER REFERENCES projects(id),
  display_name TEXT NOT NULL,
  category TEXT NOT NULL CHECK(category IN ('Staff','Manager','Watchlist')),
  position TEXT NOT NULL DEFAULT '',
  record_reference TEXT NOT NULL DEFAULT '',
  record_owner TEXT NOT NULL,
  purpose TEXT NOT NULL,
  review_date TEXT NOT NULL,
  status TEXT NOT NULL CHECK(status IN ('Pending review','Active','Inactive')),
  photo_blob BLOB,
  photo_mime TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_people_category_status ON people(category,status);
CREATE TABLE IF NOT EXISTS presence_events (
  id INTEGER PRIMARY KEY,
  project_id INTEGER NOT NULL REFERENCES projects(id),
  person_id INTEGER NOT NULL REFERENCES people(id),
  status TEXT NOT NULL CHECK(status IN ('Checked out','Returned')),
  actor TEXT NOT NULL,
  note TEXT NOT NULL DEFAULT '',
  occurred_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_presence_project_person_time ON presence_events(project_id,person_id,occurred_at DESC);
CREATE TABLE IF NOT EXISTS alerts (
  id INTEGER PRIMARY KEY,
  project_id INTEGER REFERENCES projects(id),
  source TEXT NOT NULL,
  occurred_at TEXT NOT NULL,
  summary TEXT NOT NULL,
  status TEXT NOT NULL CHECK(status IN ('Needs review','Reviewed','Dismissed','Escalated')),
  reviewer TEXT NOT NULL DEFAULT '',
  review_note TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_alerts_status_time ON alerts(status, occurred_at DESC);
CREATE TABLE IF NOT EXISTS audit_log (
  id INTEGER PRIMARY KEY,
  project_id INTEGER REFERENCES projects(id),
  actor TEXT NOT NULL,
  action TEXT NOT NULL,
  object_type TEXT NOT NULL,
  object_id INTEGER,
  detail TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_log(created_at DESC);
CREATE TABLE IF NOT EXISTS movement_events (
  id INTEGER PRIMARY KEY,
  project_id INTEGER NOT NULL REFERENCES projects(id),
  event_reference TEXT NOT NULL,
  direction TEXT NOT NULL DEFAULT 'Unknown' CHECK(direction IN ('Entry','Exit','Unknown')),
  first_camera TEXT NOT NULL,
  last_camera TEXT NOT NULL,
  first_seen TEXT NOT NULL,
  last_seen TEXT NOT NULL,
  review_status TEXT NOT NULL CHECK(review_status IN ('Unreviewed','Reviewed','Dismissed')),
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_movements_project_time ON movement_events(project_id,last_seen DESC);
CREATE TABLE IF NOT EXISTS automation_outbox (
  id TEXT PRIMARY KEY,
  project_id INTEGER NOT NULL REFERENCES projects(id),
  event_type TEXT NOT NULL,
  payload TEXT NOT NULL,
  delivery_status TEXT NOT NULL DEFAULT 'Queued' CHECK(delivery_status IN ('Queued','Delivered','Failed')),
  created_at TEXT NOT NULL,
  delivered_at TEXT,
  attempts INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_automation_project_queue ON automation_outbox(project_id,delivery_status,created_at DESC);
CREATE TABLE IF NOT EXISTS project_discovery (
  project_id INTEGER PRIMARY KEY REFERENCES projects(id) ON DELETE CASCADE,
  enabled INTEGER NOT NULL DEFAULT 0,
  cidrs TEXT NOT NULL DEFAULT '[]',
  last_scan REAL,
  last_result TEXT NOT NULL DEFAULT ''
);
"""

def utcnow() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()

def db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA busy_timeout=10000")
    con.executescript(SCHEMA)
    # Migrate databases created by the first single-site prototype.
    for table in ("devices", "people", "alerts", "audit_log"):
        cols={row[1] for row in con.execute(f"PRAGMA table_info({table})")}
        if "project_id" not in cols:
            con.execute(f"ALTER TABLE {table} ADD COLUMN project_id INTEGER REFERENCES projects(id)")
    device_cols={row[1] for row in con.execute("PRAGMA table_info(devices)")}
    if "device_kind" not in device_cols:
        con.execute("ALTER TABLE devices ADD COLUMN device_kind TEXT NOT NULL DEFAULT 'Unknown'")
    if "coverage_role" not in device_cols:con.execute("ALTER TABLE devices ADD COLUMN coverage_role TEXT NOT NULL DEFAULT 'General'")
    if "source_mode" not in device_cols:con.execute("ALTER TABLE devices ADD COLUMN source_mode TEXT NOT NULL DEFAULT 'Recorder'")
    if "monitoring_purposes" not in device_cols:con.execute("ALTER TABLE devices ADD COLUMN monitoring_purposes TEXT NOT NULL DEFAULT '[]'")
    people_cols={row[1] for row in con.execute("PRAGMA table_info(people)")}
    if "position" not in people_cols:con.execute("ALTER TABLE people ADD COLUMN position TEXT NOT NULL DEFAULT ''")
    if "photo_blob" not in people_cols:con.execute("ALTER TABLE people ADD COLUMN photo_blob BLOB")
    if "photo_mime" not in people_cols:con.execute("ALTER TABLE people ADD COLUMN photo_mime TEXT NOT NULL DEFAULT ''")
    movement_cols={row[1] for row in con.execute("PRAGMA table_info(movement_events)")}
    if "direction" not in movement_cols:con.execute("ALTER TABLE movement_events ADD COLUMN direction TEXT NOT NULL DEFAULT 'Unknown'")
    user_cols={row[1] for row in con.execute("PRAGMA table_info(users)")}
    if "permissions" not in user_cols: con.execute("ALTER TABLE users ADD COLUMN permissions TEXT NOT NULL DEFAULT '[]'")
    if "active" not in user_cols: con.execute("ALTER TABLE users ADD COLUMN active INTEGER NOT NULL DEFAULT 1")
    full_access=json.dumps(list(PERMISSIONS),separators=(",",":"))
    con.execute("UPDATE users SET permissions=? WHERE id=(SELECT MIN(id) FROM users) AND permissions='[]'",(full_access,))
    con.execute("CREATE INDEX IF NOT EXISTS idx_devices_project ON devices(project_id,created_at DESC)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_people_project ON people(project_id,category,status)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_alerts_project ON alerts(project_id,status,occurred_at DESC)")
    if not con.execute("SELECT 1 FROM projects LIMIT 1").fetchone():
        con.execute("INSERT INTO projects(name,location,created_at) VALUES('Casino Site 1','',?)",(utcnow(),))
    default_project=con.execute("SELECT id FROM projects ORDER BY id LIMIT 1").fetchone()[0]
    for table in ("devices", "people", "alerts", "audit_log"):
        con.execute(f"UPDATE {table} SET project_id=? WHERE project_id IS NULL",(default_project,))
    con.commit()
    try:
        os.chmod(DB_PATH.parent, 0o700)
        if DB_PATH.exists(): os.chmod(DB_PATH, 0o600)
    except OSError:
        pass
    return con

def init_admin() -> None:
    con = db()
    try:
        if con.execute("SELECT 1 FROM users LIMIT 1").fetchone():
            print("An administrator already exists.")
            return
        username = input("Administrator username: ").strip()
        if not username or len(username) > 80:
            raise SystemExit("Username must contain 1–80 characters.")
        password = getpass.getpass("Password (minimum 14 characters): ")
        confirm = getpass.getpass("Confirm password: ")
        if len(password) < 14 or password != confirm:
            raise SystemExit("Passwords must match and be at least 14 characters.")
        salt = secrets.token_bytes(16)
        digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, PBKDF2_ROUNDS)
        con.execute("INSERT INTO users(username,password_salt,password_hash,created_at) VALUES(?,?,?,?)", (username,salt,digest,utcnow()))
        con.commit()
        print(f"Administrator created in {DB_PATH}")
    finally:
        con.close()

def record_audit(con: sqlite3.Connection, actor: str, action: str, object_type: str, object_id: int | None, detail: str = "", project_id: int | None = None) -> None:
    con.execute("INSERT INTO audit_log(project_id,actor,action,object_type,object_id,detail,created_at) VALUES(?,?,?,?,?,?,?)", (project_id,actor,action,object_type,object_id,detail[:500],utcnow()))

def queue_automation_event(con: sqlite3.Connection, project_id: int, event_type: str, object_type: str, object_id: int | None, status: str = "") -> None:
    # Outbox payload intentionally excludes names, IDs, photos, biometric data, and footage.
    event_id=str(uuid.uuid4()); created=utcnow()
    payload={"event_id":event_id,"event_type":event_type,"event_version":1,"occurred_at":created,"project_id":project_id,"object_type":object_type,"object_id":object_id,"status":status}
    con.execute("INSERT INTO automation_outbox(id,project_id,event_type,payload,delivery_status,created_at) VALUES(?,?,?,?,'Queued',?)",(event_id,project_id,event_type,json.dumps(payload,separators=(",",":")),created))

RFC1918=(ipaddress.ip_network("10.0.0.0/8"),ipaddress.ip_network("172.16.0.0/12"),ipaddress.ip_network("192.168.0.0/16"))
def validate_discovery_networks(value: object) -> list[str]:
    if not isinstance(value,list) or len(value)>16: raise ValueError("Enter up to 16 private IPv4 subnets")
    result=[]
    for item in value:
        if not isinstance(item,str): raise ValueError("Each subnet must be CIDR text")
        try: network=ipaddress.ip_network(item.strip(),strict=False)
        except ValueError: raise ValueError(f"Invalid subnet: {item}")
        if network.version!=4 or not any(network.subnet_of(private) for private in RFC1918): raise ValueError("Discovery is limited to RFC1918 private IPv4 subnets")
        if network.num_addresses>1024: raise ValueError("Each discovery subnet must contain 1,024 addresses or fewer")
        if str(network) not in result: result.append(str(network))
    return result

def onvif_probe(cidrs: list[str], timeout: float=2.5) -> list[dict[str,str]]:
    """Send scoped WS-Discovery multicast probes; never opens a video stream."""
    found:dict[str,dict[str,str]]={}
    envelope="""<?xml version="1.0" encoding="UTF-8"?>
<e:Envelope xmlns:e="http://www.w3.org/2003/05/soap-envelope" xmlns:w="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:d="http://schemas.xmlsoap.org/ws/2005/04/discovery" xmlns:dn="http://www.onvif.org/ver10/network/wsdl"><e:Header><w:MessageID>urn:uuid:%s</w:MessageID><w:To>urn:schemas-xmlsoap-org:ws:2005:04:discovery</w:To><w:Action>http://schemas.xmlsoap.org/ws/2005/04/discovery/Probe</w:Action></e:Header><e:Body><d:Probe><d:Types>dn:NetworkVideoTransmitter</d:Types></d:Probe></e:Body></e:Envelope>"""
    for cidr in cidrs:
        network=ipaddress.ip_network(cidr)
        # Select a local interface address inside the requested subnet. This prevents
        # discovery packets from being sent out through an unrelated network adapter.
        route=socket.socket(socket.AF_INET,socket.SOCK_DGRAM)
        try:
            route.connect((str(network.network_address+1),9)); local_ip=route.getsockname()[0]
        except OSError:
            continue
        finally: route.close()
        if ipaddress.ip_address(local_ip) not in network: continue
        sock=socket.socket(socket.AF_INET,socket.SOCK_DGRAM,socket.IPPROTO_UDP)
        try:
            sock.setsockopt(socket.SOL_SOCKET,socket.SO_REUSEADDR,1)
            sock.setsockopt(socket.IPPROTO_IP,socket.IP_MULTICAST_IF,socket.inet_aton(local_ip))
            sock.setsockopt(socket.IPPROTO_IP,socket.IP_MULTICAST_TTL,1)
            sock.bind(("",0)); sock.settimeout(.25)
            body=(envelope%uuid.uuid4()).encode()
            sock.sendto(body,("239.255.255.250",3702))
            deadline=time.monotonic()+timeout
            while time.monotonic()<deadline:
                try: packet,addr=sock.recvfrom(65535)
                except socket.timeout: continue
                ip=addr[0]
                try:
                    if ipaddress.ip_address(ip) not in network: continue
                    root=ET.fromstring(packet)
                    values={node.tag.rsplit("}",1)[-1]: (node.text or "").strip() for node in root.iter()}
                    xaddr=values.get("XAddrs","").split()[0] if values.get("XAddrs") else ""
                    types=values.get("Types","")
                    endpoint=values.get("Address","")
                    found[ip]={"ip":ip,"xaddr":xaddr,"types":types[:240],"endpoint":endpoint[:120]}
                except (ET.ParseError,ValueError,IndexError): continue
        except OSError:
            continue
        finally: sock.close()
    return list(found.values())[:256]

def selected_project(handler: BaseHTTPRequestHandler, con: sqlite3.Connection) -> int | None:
    raw=handler.headers.get("X-Casino-Project", "")
    if not raw.isdigit(): return None
    project_id=int(raw)
    return project_id if con.execute("SELECT 1 FROM projects WHERE id=?",(project_id,)).fetchone() else None

def valid_permissions(value: object) -> list[str]:
    if not isinstance(value,list) or not value or any(not isinstance(p,str) or p not in PERMISSIONS for p in value):
        raise ValueError("Choose one or more valid permissions")
    return sorted(set(value))

PEOPLE_COLUMNS=("display_name","category","position","record_reference","record_owner","purpose","review_date")
LEGACY_PEOPLE_COLUMNS=("display_name","category","record_reference","record_owner","purpose","review_date")
MAX_IMPORT_BYTES=20*1024*1024
MAX_IMPORT_ROWS=10_000
def people_workbook(rows: list[dict] | None=None, include_data: bool=False) -> bytes:
    book=Workbook();sheet=book.active;sheet.title="People"
    headers=list(PEOPLE_COLUMNS)
    if include_data:headers += ["status","created_at","updated_at"]
    sheet.append(headers)
    for row in (rows or []):
        sheet.append([row.get(key,"") for key in headers])
        for cell in sheet[sheet.max_row]:
            if isinstance(cell.value,str):cell.data_type="s"
    sheet.freeze_panes="A2";sheet.auto_filter.ref=sheet.dimensions
    for cell in sheet[1]:cell.font=Font(bold=True,color="FFFFFF");cell.fill=PatternFill("solid",fgColor="176A68")
    if not include_data:
        categories=DataValidation(type="list",formula1='"Staff,Manager,Watchlist"',allow_blank=False);sheet.add_data_validation(categories);categories.add("B2:B10001")
    for row in sheet.iter_rows(min_row=2,max_col=len(headers)):
        for cell in row:cell.number_format="@"
    if "review_date" in headers:
        col=headers.index("review_date")+1
        for row in range(2,sheet.max_row+1):sheet.cell(row,col).number_format="yyyy-mm-dd"
    for col,width in zip("ABCDEFG",(28,18,24,24,24,54,16)):sheet.column_dimensions[col].width=width
    guide=book.create_sheet("Instructions");guide.append(["People import template"]);guide.append(["Add one person per row on the People sheet. Keep the exact column names."])
    guide.append(["Allowed category values: Staff, Manager, Watchlist"]);guide.append(["New imported records always start as Pending review."])
    guide.append(["Use ISO date format YYYY-MM-DD for review_date."]);guide.append(["Do not add government IDs, face images, biometrics, or formulas."])
    output=io.BytesIO();book.save(output);return output.getvalue()

def xlsx_response(handler: BaseHTTPRequestHandler, filename: str, body: bytes) -> None:
    handler.send_response(200);handler.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");handler.send_header("Content-Length",str(len(body)));handler.send_header("Content-Disposition",f'attachment; filename="{filename}"');handler.security_headers();handler.end_headers();handler.wfile.write(body)

def openai_http(method: str, path: str, payload: dict | None = None):
    key=os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:raise RuntimeError("OPENAI_API_KEY is not configured in the server environment")
    headers={"Authorization":"Bearer "+key,"Accept":"application/json"}
    body=None
    if payload is not None:
        headers["Content-Type"]="application/json";body=json.dumps(payload).encode()
    request=urllib.request.Request("https://api.openai.com/v1/"+path, data=body, headers=headers, method=method)
    with urllib.request.urlopen(request, timeout=20) as response:
        raw=response.read(1024*1024+1)
    if len(raw)>1024*1024:raise ValueError("OpenAI response exceeded the 1 MB response limit")
    return json.loads(raw.decode("utf-8"))

def openai_error(status: int) -> str:
    return {401:"OpenAI rejected the key. Check OPENAI_API_KEY.",403:"This key or project cannot access the selected model.",404:"The configured model was not found. Check OPENAI_MODEL.",429:"OpenAI rate or billing limit reached. Check API account limits and billing.",500:"OpenAI returned a server error. Try again later.",502:"OpenAI returned a server error. Try again later.",503:"OpenAI is temporarily unavailable. Try again later."}.get(status,"OpenAI request failed (HTTP %s)." % status)

def openai_output_text(response: dict) -> str:
    direct=response.get("output_text")
    if isinstance(direct,str) and direct.strip():return direct.strip()
    parts=[]
    for item in response.get("output",[]):
        if item.get("type")=="message":
            for content in item.get("content",[]):
                if content.get("type")=="output_text" and isinstance(content.get("text"),str):parts.append(content["text"])
    return "\n".join(parts).strip()

class Handler(BaseHTTPRequestHandler):
    server_version = "CasinoOpsLocal/0.1"
    def log_message(self, fmt: str, *args: object) -> None:
        sys.stderr.write("%s - %s\n" % (self.log_date_time_string(), fmt % args))

    def send_json(self, status: int, payload: object, headers: dict[str,str] | None = None) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.security_headers()
        for k,v in (headers or {}).items(): self.send_header(k,v)
        self.end_headers(); self.wfile.write(body)

    def security_headers(self) -> None:
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Security-Policy", "default-src 'self'; style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'")

    def read_json(self) -> dict:
        try:
            n = int(self.headers.get("Content-Length", "0"))
            if n < 1 or n > 32768: raise ValueError()
            value = json.loads(self.rfile.read(n))
            if not isinstance(value, dict): raise ValueError()
            return value
        except (ValueError, json.JSONDecodeError):
            raise ValueError("Invalid or oversized JSON request")

    def session(self, con: sqlite3.Connection):
        cookie = http.cookies.SimpleCookie(self.headers.get("Cookie", ""))
        morsel = cookie.get("ops_session")
        if not morsel: return None
        token_hash = hashlib.sha256(morsel.value.encode()).hexdigest()
        return con.execute("SELECT s.*,u.username,u.permissions FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token_hash=? AND s.expires_at>? AND u.active=1", (token_hash,int(time.time()))).fetchone()

    def has_permission(self, sess: sqlite3.Row, permission: str) -> bool:
        aliases={"view_devices":"manage_devices","view_people":"manage_people","view_alerts":"review_alerts"}
        try:
            permissions=json.loads(sess["permissions"])
            return permission in permissions or aliases.get(permission) in permissions
        except (ValueError,TypeError,KeyError): return False

    def require_permission(self, sess: sqlite3.Row, permission: str) -> bool:
        if self.has_permission(sess,permission): return True
        self.send_json(403,{"error":"Your account does not have permission for this action"}); return False

    def require_csrf(self, sess: sqlite3.Row) -> bool:
        # Same-origin and per-session CSRF token are required on every write.
        origin = self.headers.get("Origin")
        expected = f"http://{self.headers.get('Host','')}"
        if origin and origin.rstrip("/") != expected.rstrip("/"): return False
        return hmac.compare_digest(self.headers.get("X-CSRF-Token", ""), sess["csrf_token"])

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/":
            try: body = (ROOT / "index.html").read_bytes()
            except OSError: return self.send_error(500)
            self.send_response(200); self.send_header("Content-Type","text/html; charset=utf-8"); self.send_header("Content-Length",str(len(body))); self.security_headers(); self.end_headers(); self.wfile.write(body); return
        if path == "/app.js":
            try: body = (ROOT / "app.js").read_bytes()
            except OSError: return self.send_error(500)
            self.send_response(200); self.send_header("Content-Type","text/javascript; charset=utf-8"); self.send_header("Content-Length",str(len(body))); self.security_headers(); self.end_headers(); self.wfile.write(body); return
        if path.startswith("/api/"): return self.api_get(path)
        return self.send_error(404)

    def api_get(self, path: str) -> None:
        con=db()
        try:
            if path=="/api/bootstrap":
                ready=bool(con.execute("SELECT 1 FROM users LIMIT 1").fetchone())
                return self.send_json(200,{"admin_ready":ready,"local_only":True,"integration":"not connected"})
            sess=self.session(con)
            if not sess:
                return self.send_json(401,{"error":"Sign in required"})
            if path == "/api/me": return self.send_json(200,{"username":sess["username"],"csrf_token":sess["csrf_token"],"permissions":json.loads(sess["permissions"])})
            if path=="/api/ai/status":
                if not self.require_permission(sess,"review_alerts"):return
                return self.send_json(200,{"provider":"OpenAI","configured":bool(os.environ.get("OPENAI_API_KEY","").strip()),"model":OPENAI_MODEL})
            if path == "/api/projects":
                if not any(self.has_permission(sess,p) for p in ("view_dashboard","view_devices","view_people","view_movement","view_alerts","view_automation","view_audit","manage_devices","manage_people","review_alerts","manage_projects","manage_users","manage_backups","run_discovery")):
                    return self.send_json(403,{"error":"Your account does not have permission to view a CCTV project"})
                return self.send_json(200,[dict(r) for r in con.execute("SELECT * FROM projects ORDER BY name COLLATE NOCASE").fetchall()])
            if path == "/api/users":
                if not self.require_permission(sess,"manage_users"): return
                rows=con.execute("SELECT id,username,permissions,active,created_at FROM users ORDER BY username COLLATE NOCASE").fetchall()
                return self.send_json(200,[{**dict(r),"permissions":json.loads(r["permissions"]),"is_current_user":r["username"]==sess["username"]} for r in rows])
            if path in ("/api/backups","/api/backups/download"):
                if not self.require_permission(sess,"manage_backups"): return
                BACKUP_DIR.mkdir(parents=True,exist_ok=True)
                try:os.chmod(BACKUP_DIR,0o700)
                except OSError:pass
                if path=="/api/backups":
                    entries=[]
                    for file in sorted(BACKUP_DIR.glob("casino-ops-*.sqlite3"),key=lambda p:p.stat().st_mtime,reverse=True)[:30]:
                        if file.is_file():entries.append({"name":file.name,"size":file.stat().st_size,"modified":datetime.fromtimestamp(file.stat().st_mtime,timezone.utc).replace(microsecond=0).isoformat()})
                    return self.send_json(200,entries)
                name=parse_qs(urlparse(self.path).query).get("name",[""])[0]
                if not name.startswith("casino-ops-") or not name.endswith(".sqlite3") or Path(name).name!=name:return self.send_json(400,{"error":"Invalid backup name"})
                file=BACKUP_DIR/name
                if not file.is_file():return self.send_json(404,{"error":"Backup not found"})
                project_for_event=selected_project(self,con) or con.execute("SELECT id FROM projects ORDER BY id LIMIT 1").fetchone()[0]
                record_audit(con,sess["username"],"downloaded","backup",None,name,project_for_event);queue_automation_event(con,project_for_event,"backup.downloaded","backup",None,"downloaded");con.commit()
                body=file.read_bytes();self.send_response(200);self.send_header("Content-Type","application/vnd.sqlite3");self.send_header("Content-Length",str(len(body)));self.send_header("Content-Disposition",f'attachment; filename="{name}"');self.security_headers();self.end_headers();self.wfile.write(body);return
            project_id=selected_project(self,con)
            if not project_id: return self.send_json(400,{"error":"Select a valid CCTV project"})
            photo_parts=path.strip("/").split("/")
            if len(photo_parts)==4 and photo_parts[:2]==["api","people"] and photo_parts[2].isdigit() and photo_parts[3]=="photo":
                if not self.require_permission(sess,"view_people"):return
                row=con.execute("SELECT photo_blob,photo_mime FROM people WHERE id=? AND project_id=?",(int(photo_parts[2]),project_id)).fetchone()
                if not row or not row["photo_blob"]:return self.send_error(404)
                body=row["photo_blob"];self.send_response(200);self.send_header("Content-Type",row["photo_mime"]);self.send_header("Content-Length",str(len(body)));self.security_headers();self.send_header("X-Content-Type-Options","nosniff");self.end_headers();self.wfile.write(body);return
            if path=="/api/presence":
                if not self.require_permission(sess,"view_people"):return
                rows=con.execute("SELECT e.id,e.person_id,p.display_name,e.status,e.actor,e.note,e.occurred_at FROM presence_events e JOIN people p ON p.id=e.person_id WHERE e.project_id=? ORDER BY e.occurred_at DESC LIMIT 500",(project_id,)).fetchall()
                return self.send_json(200,[dict(r) for r in rows])
            if path=="/api/people/template.xlsx":
                if not self.require_permission(sess,"manage_people"): return
                return xlsx_response(self,"people-import-template.xlsx",people_workbook())
            if path=="/api/people/export.xlsx":
                if not self.require_permission(sess,"view_people"): return
                rows=[dict(r) for r in con.execute("SELECT display_name,category,position,record_reference,record_owner,purpose,review_date,status,created_at,updated_at FROM people WHERE project_id=? ORDER BY created_at DESC",(project_id,)).fetchall()]
                record_audit(con,sess["username"],"exported","people",None,f"rows={len(rows)}",project_id);queue_automation_event(con,project_id,"people.exported","people",None,f"rows={len(rows)}");con.commit()
                return xlsx_response(self,"people-export.xlsx",people_workbook(rows,include_data=True))
            if path == "/api/automation/outbox":
                if not self.require_permission(sess,"view_automation"): return
                rows=con.execute("SELECT id,event_type,payload,delivery_status,created_at,delivered_at,attempts FROM automation_outbox WHERE project_id=? ORDER BY created_at DESC LIMIT 200",(project_id,)).fetchall()
                return self.send_json(200,[dict(r) for r in rows])
            if path == "/api/discovery/settings":
                if not self.require_permission(sess,"run_discovery"): return
                row=con.execute("SELECT enabled,cidrs,last_scan,last_result FROM project_discovery WHERE project_id=?",(project_id,)).fetchone()
                return self.send_json(200,dict(row) if row else {"enabled":0,"cidrs":"[]","last_scan":None,"last_result":""})
            table={"/api/devices":("devices","created_at DESC","view_devices"),"/api/people":("people","created_at DESC","view_people"),"/api/alerts":("alerts","occurred_at DESC","view_alerts"),"/api/audit":("audit_log","created_at DESC","view_audit"),"/api/movements":("movement_events","last_seen DESC","view_movement")}.get(path)
            if table:
                if not self.require_permission(sess,table[2]): return
                if table[0]=="audit_log":rows=con.execute("SELECT * FROM audit_log WHERE project_id=? OR project_id IS NULL ORDER BY created_at DESC LIMIT 500",(project_id,)).fetchall()
                elif table[0]=="people":rows=con.execute("SELECT p.id,p.project_id,p.display_name,p.category,p.position,p.record_reference,p.record_owner,p.purpose,p.review_date,p.status,p.photo_mime,p.created_at,p.updated_at,(SELECT e.status FROM presence_events e WHERE e.person_id=p.id AND e.project_id=p.project_id ORDER BY e.occurred_at DESC LIMIT 1) AS presence_status,(SELECT e.occurred_at FROM presence_events e WHERE e.person_id=p.id AND e.project_id=p.project_id ORDER BY e.occurred_at DESC LIMIT 1) AS presence_at FROM people p WHERE p.project_id=? ORDER BY p.created_at DESC LIMIT 500",(project_id,)).fetchall()
                else:rows=con.execute(f"SELECT * FROM {table[0]} WHERE project_id=? ORDER BY {table[1]} LIMIT 500",(project_id,)).fetchall()
                if table[0]=="devices":
                    result=[]
                    for row in rows:
                        item=dict(row)
                        try:item["monitoring_purposes"]=json.loads(item.get("monitoring_purposes") or "[]")
                        except (TypeError,json.JSONDecodeError):item["monitoring_purposes"]=[]
                        result.append(item)
                    return self.send_json(200,result)
                return self.send_json(200,[dict(r) for r in rows])
            return self.send_json(404,{"error":"Not found"})
        finally: con.close()

    def do_POST(self) -> None:
        path=urlparse(self.path).path
        con=db()
        try:
            if path == "/api/setup": return self.first_run_setup(con)
            if path == "/api/login": return self.login(con)
            sess=self.session(con)
            if not sess: return self.send_json(401,{"error":"Sign in required"})
            if not self.require_csrf(sess): return self.send_json(403,{"error":"Request verification failed"})
            photo_parts=path.strip("/").split("/")
            if len(photo_parts)==4 and photo_parts[:2]==["api","people"] and photo_parts[2].isdigit() and photo_parts[3]=="photo":
                if not self.require_permission(sess,"manage_people"):return
                project_id=selected_project(self,con)
                if not project_id:return self.send_json(400,{"error":"Select a valid CCTV project"})
                mime=self.headers.get("Content-Type","").split(";",1)[0].strip().lower()
                try:size=int(self.headers.get("Content-Length","0"))
                except ValueError:size=0
                if size<1 or size>5*1024*1024:return self.send_json(413,{"error":"Photo must be between 1 byte and 5 MB"})
                if mime not in ("image/jpeg","image/png"):return self.send_json(415,{"error":"Use a JPEG or PNG staff photo"})
                body=self.rfile.read(size)
                if (mime=="image/jpeg" and not body.startswith(b"\xff\xd8\xff")) or (mime=="image/png" and not body.startswith(b"\x89PNG\r\n\x1a\n")):
                    return self.send_json(400,{"error":"The selected file is not a valid JPEG or PNG image"})
                person_id=int(photo_parts[2]);row=con.execute("SELECT id FROM people WHERE id=? AND project_id=?",(person_id,project_id)).fetchone()
                if not row:return self.send_json(404,{"error":"People record not found"})
                con.execute("UPDATE people SET photo_blob=?,photo_mime=?,updated_at=? WHERE id=? AND project_id=?",(sqlite3.Binary(body),mime,utcnow(),person_id,project_id))
                record_audit(con,sess["username"],"photo_added","person",person_id,"Profile photo uploaded",project_id);queue_automation_event(con,project_id,"person.photo_added","person",person_id,"photo_added");con.commit()
                return self.send_json(200,{"ok":True})
            if path=="/api/people/import.xlsx": return self.import_people(con,sess)
            try: data=self.read_json()
            except ValueError as e: return self.send_json(400,{"error":str(e)})
            actor=sess["username"]
            if path=="/api/ai/check":
                if not self.require_permission(sess,"review_alerts"):return
                project_id=selected_project(self,con)
                if not project_id:return self.send_json(400,{"error":"Select a valid CCTV project"})
                if not os.environ.get("OPENAI_API_KEY","").strip():return self.send_json(503,{"error":"OpenAI key is not configured. Set OPENAI_API_KEY in the server environment and restart the app."})
                try:
                    model=openai_http("GET","models/"+quote(OPENAI_MODEL,safe=""))
                    record_audit(con,actor,"checked","openai_connection",None,"model_accessible; no prompt sent",project_id);con.commit()
                    return self.send_json(200,{"ok":True,"model":model.get("id",OPENAI_MODEL),"message":"OpenAI key accepted and configured model is accessible. No prompt or image was sent."})
                except urllib.error.HTTPError as error:
                    record_audit(con,actor,"failed","openai_connection",None,"http_status=%s" % error.code,project_id);con.commit()
                    return self.send_json(502,{"error":openai_error(error.code)})
                except (urllib.error.URLError,TimeoutError,ConnectionError):
                    record_audit(con,actor,"failed","openai_connection",None,"network_connection_failed",project_id);con.commit()
                    return self.send_json(502,{"error":"Could not reach api.openai.com from this server. Check network access and try again."})
                except (ValueError,json.JSONDecodeError):
                    return self.send_json(502,{"error":"OpenAI returned an unreadable response."})
            if path=="/api/ai/draft":
                if not self.require_permission(sess,"review_alerts"):return
                project_id=selected_project(self,con)
                if not project_id:return self.send_json(400,{"error":"Select a valid CCTV project"})
                incident=text(data,"incident_text",2000)
                if not incident:return self.send_json(400,{"error":"Enter a short human-reviewed incident summary"})
                if not os.environ.get("OPENAI_API_KEY","").strip():return self.send_json(503,{"error":"OpenAI key is not configured. Set OPENAI_API_KEY in the server environment and restart the app."})
                payload={"model":OPENAI_MODEL,"store":False,"max_output_tokens":250,"instructions":"Draft a concise, neutral message for authorized casino staff from the user-provided human-reviewed incident summary. Do not identify people, match faces, infer identity, infer intent or guilt, or invent facts. Clearly say when information is unconfirmed. Include only supplied details. Return a short subject and message. Do not send or route the message.","input":incident}
                try:
                    response=openai_http("POST","responses",payload);draft=openai_output_text(response)
                    if not draft:raise ValueError("empty response")
                    record_audit(con,actor,"generated","staff_alert_draft",None,"human_reviewed_text_only; not_sent",project_id);queue_automation_event(con,project_id,"ai.staff_alert_draft_created","staff_alert_draft",None,"drafted_not_sent");con.commit()
                    return self.send_json(200,{"draft":draft,"model":response.get("model",OPENAI_MODEL),"sent":False})
                except urllib.error.HTTPError as error:
                    record_audit(con,actor,"failed","staff_alert_draft",None,"http_status=%s" % error.code,project_id);con.commit()
                    return self.send_json(502,{"error":openai_error(error.code)})
                except (urllib.error.URLError,TimeoutError,ConnectionError):
                    record_audit(con,actor,"failed","staff_alert_draft",None,"network_connection_failed",project_id);con.commit()
                    return self.send_json(502,{"error":"Could not reach api.openai.com from this server. Check network access and try again."})
                except (ValueError,json.JSONDecodeError):
                    record_audit(con,actor,"failed","staff_alert_draft",None,"invalid_provider_response",project_id);con.commit()
                    return self.send_json(502,{"error":"OpenAI returned an unreadable draft. Nothing was sent."})
            if path=="/api/presence":
                if not self.require_permission(sess,"manage_people"):return
                project_id=selected_project(self,con)
                if not project_id:return self.send_json(400,{"error":"Select a valid CCTV project"})
                person_id=integer(data,"person_id",1,2147483647);status=choice(data,"status",("Checked out","Returned"));note=text(data,"note",240,required=False)
                person=con.execute("SELECT category,status FROM people WHERE id=? AND project_id=?",(person_id,project_id)).fetchone()
                if not person:return self.send_json(404,{"error":"People record not found"})
                if person["category"] not in ("Staff","Manager") or person["status"]!="Active":return self.send_json(409,{"error":"Only active staff and manager records can be checked in or out"})
                previous=con.execute("SELECT status FROM presence_events WHERE project_id=? AND person_id=? ORDER BY occurred_at DESC,id DESC LIMIT 1",(project_id,person_id)).fetchone()
                if (previous and previous["status"]==status) or (not previous and status=="Returned"):
                    return self.send_json(409,{"error":"That person is already in this presence state"})
                occurred=utcnow();cur=con.execute("INSERT INTO presence_events(project_id,person_id,status,actor,note,occurred_at) VALUES(?,?,?,?,?,?)",(project_id,person_id,status,actor,note,occurred))
                record_audit(con,actor,"presence_changed","person",person_id,f"status={status}; event_id={cur.lastrowid}",project_id);queue_automation_event(con,project_id,"person.presence_changed","person",person_id,status.lower().replace(" ","_"));con.commit()
                return self.send_json(201,{"id":cur.lastrowid,"status":status,"occurred_at":occurred})
            if path == "/api/projects":
                name=text(data,"name",120); location=text(data,"location",160,required=False)
                if not self.require_permission(sess,"manage_projects"): return
                if not name:return self.send_json(400,{"error":"Project name is required"})
                cur=con.execute("INSERT INTO projects(name,location,created_at) VALUES(?,?,?)",(name,location,utcnow())); record_audit(con,actor,"created","project",cur.lastrowid,"CCTV project created",cur.lastrowid); queue_automation_event(con,cur.lastrowid,"project.created","project",cur.lastrowid,"created"); con.commit()
                return self.send_json(201,{"id":cur.lastrowid,"name":name,"location":location})
            if path == "/api/users":
                if not self.require_permission(sess,"manage_users"): return
                username=text(data,"username",80); password=text(data,"password",200); permissions=valid_permissions(data.get("permissions"))
                if len(password)<14:return self.send_json(400,{"error":"Use a unique password with at least 14 characters"})
                salt=secrets.token_bytes(16); digest=hashlib.pbkdf2_hmac("sha256",password.encode(),salt,PBKDF2_ROUNDS)
                cur=con.execute("INSERT INTO users(username,password_salt,password_hash,permissions,active,created_at) VALUES(?,?,?,?,1,?)",(username,salt,digest,json.dumps(permissions,separators=(",",":")),utcnow()))
                record_audit(con,actor,"created","user",cur.lastrowid,"Account created with individually selected permissions")
                project_for_event=selected_project(self,con) or con.execute("SELECT id FROM projects ORDER BY id LIMIT 1").fetchone()[0]
                queue_automation_event(con,project_for_event,"user.created","user",cur.lastrowid,"created")
                con.commit();return self.send_json(201,{"id":cur.lastrowid,"username":username,"permissions":permissions,"active":True})
            if path == "/api/backups":
                if not self.require_permission(sess,"manage_backups"): return
                project_id=selected_project(self,con) or con.execute("SELECT id FROM projects ORDER BY id LIMIT 1").fetchone()[0]
                stamp=datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S");name=f"casino-ops-{stamp}-{secrets.token_hex(3)}.sqlite3"
                BACKUP_DIR.mkdir(parents=True,exist_ok=True)
                try:os.chmod(BACKUP_DIR,0o700)
                except OSError:pass
                destination=BACKUP_DIR/name
                record_audit(con,actor,"started","backup",None,name);queue_automation_event(con,project_id,"backup.started","backup",None,"started");con.commit()
                try:
                    source=sqlite3.connect(DB_PATH,timeout=10);snapshot=sqlite3.connect(destination,timeout=10)
                    try:
                        source.backup(snapshot);snapshot.execute("DELETE FROM sessions");snapshot.commit();snapshot.execute("PRAGMA journal_mode=DELETE")
                    finally:source.close();snapshot.close()
                    os.chmod(destination,0o600)
                except Exception as error:
                    try:destination.unlink(missing_ok=True)
                    except OSError:pass
                    record_audit(con,actor,"failed","backup",None,type(error).__name__);queue_automation_event(con,project_id,"backup.failed","backup",None,"failed");con.commit()
                    return self.send_json(500,{"error":"Backup creation failed; check local disk space and folder permissions"})
                record_audit(con,actor,"completed","backup",None,name);queue_automation_event(con,project_id,"backup.completed","backup",None,"complete");con.commit()
                return self.send_json(201,{"name":name,"size":destination.stat().st_size,"message":"Database snapshot created. Store a copy on approved encrypted backup storage."})
            project_id=selected_project(self,con)
            if not project_id:return self.send_json(400,{"error":"Select a valid CCTV project"})
            if path == "/api/discovery/settings":
                if not self.require_permission(sess,"run_discovery"): return
                enabled=data.get("enabled")
                if not isinstance(enabled,bool):return self.send_json(400,{"error":"Discovery enabled must be true or false"})
                networks=validate_discovery_networks(data.get("cidrs",[]))
                con.execute("INSERT INTO project_discovery(project_id,enabled,cidrs) VALUES(?,?,?) ON CONFLICT(project_id) DO UPDATE SET enabled=excluded.enabled,cidrs=excluded.cidrs",(project_id,1 if enabled else 0,json.dumps(networks)))
                record_audit(con,actor,"updated","discovery_settings",project_id,f"enabled={enabled}; subnet_count={len(networks)}",project_id);queue_automation_event(con,project_id,"discovery.settings_changed","discovery",project_id,"enabled" if enabled else "disabled");con.commit()
                return self.send_json(200,{"enabled":enabled,"cidrs":networks})
            if path == "/api/discovery/run":
                if not self.require_permission(sess,"run_discovery"): return
                cfg=con.execute("SELECT * FROM project_discovery WHERE project_id=?",(project_id,)).fetchone()
                if not cfg or not cfg["enabled"]:return self.send_json(409,{"error":"Enable device discovery in this project's admin settings first"})
                networks=validate_discovery_networks(json.loads(cfg["cidrs"]))
                if not networks:return self.send_json(400,{"error":"Add at least one approved CCTV subnet before discovery"})
                if cfg["last_scan"] and time.time()-cfg["last_scan"]<60:return self.send_json(429,{"error":"Wait one minute between discovery runs"})
                con.execute("UPDATE project_discovery SET last_scan=?,last_result='Discovery running' WHERE project_id=?",(time.time(),project_id));con.commit()
                record_audit(con,actor,"started","device_discovery",project_id,f"subnet_count={len(networks)}",project_id);queue_automation_event(con,project_id,"discovery.started","discovery",project_id,"running");con.commit()
                candidates=onvif_probe(networks)
                added=[]
                for candidate in candidates:
                    exists=con.execute("SELECT id FROM devices WHERE project_id=? AND host=? LIMIT 1",(project_id,candidate["ip"])).fetchone()
                    if exists: continue
                    parsed=urlparse(candidate["xaddr"]) if candidate["xaddr"] else None
                    port=parsed.port if parsed and parsed.port else (443 if parsed and parsed.scheme=="https" else 80)
                    now=utcnow(); name=f"ONVIF candidate {candidate['ip']}"
                    cur=con.execute("INSERT INTO devices(project_id,name,device_kind,vendor_model,host,port,protocol,stream_label,status,created_at,updated_at) VALUES(?,?,'Unknown','Unknown / identify model',?,?,'ONVIF','Discovery candidate · verify device type before use','Needs verification',?,?)",(project_id,name,candidate["ip"],port,now,now))
                    record_audit(con,actor,"discovered","device",cur.lastrowid,"ONVIF discovery candidate; not connected",project_id);queue_automation_event(con,project_id,"device.discovery_candidate","device",cur.lastrowid,"needs_verification")
                    added.append({"id":cur.lastrowid,**candidate,"device_kind":"Unknown","status":"Needs verification"})
                result=f"{len(candidates)} ONVIF response(s), {len(added)} new candidate(s); none connected"
                con.execute("UPDATE project_discovery SET last_result=? WHERE project_id=?",(result,project_id));record_audit(con,actor,"completed","device_discovery",project_id,result,project_id);queue_automation_event(con,project_id,"discovery.completed","discovery",project_id,"completed");con.commit()
                return self.send_json(200,{"scanned_subnets":networks,"responses":len(candidates),"added":added,"result":result})
            if path == "/api/logout":
                cookie=http.cookies.SimpleCookie(self.headers.get("Cookie","")); m=cookie.get("ops_session")
                project_for_event=selected_project(self,con) or con.execute("SELECT id FROM projects ORDER BY id LIMIT 1").fetchone()[0]
                record_audit(con,actor,"logout","session",None,"User signed out",project_for_event);queue_automation_event(con,project_for_event,"auth.logout","session",None,"ok")
                if m: con.execute("DELETE FROM sessions WHERE token_hash=?",(hashlib.sha256(m.value.encode()).hexdigest(),))
                con.commit()
                return self.send_json(200,{"ok":True},{"Set-Cookie":"ops_session=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0"})
            if path == "/api/devices":
                if not self.require_permission(sess,"manage_devices"): return
                name=text(data,"name",120); host=text(data,"host",253); port=integer(data,"port",1,65535); protocol=choice(data,"protocol",("RTSP","ONVIF","RTSP + ONVIF","Vendor connector")); device_kind=choice(data,"device_kind",("DVR","NVR","Camera","Unknown")); vendor=text(data,"vendor_model",160,required=False); stream=text(data,"stream_label",160,required=False); coverage=choice(data,"coverage_role",("General","Entrance","Exit")); source_mode=choice(data,"source_mode",("Recorder","Standalone camera")); purposes=data.get("monitoring_purposes",[])
                allowed_purposes=("Entrance monitoring","Exit monitoring","Staff access","Gaming floor safety","Cash handling area","Incident review","Equipment area")
                if not isinstance(purposes,list) or any(p not in allowed_purposes for p in purposes):return self.send_json(400,{"error":"Choose valid monitoring purposes"})
                purposes_json=json.dumps(sorted(set(purposes)),separators=(",",":"))
                if not name or not host: return self.send_json(400,{"error":"Name and IP address or hostname are required"})
                now=utcnow(); cur=con.execute("INSERT INTO devices(project_id,name,device_kind,vendor_model,host,port,protocol,stream_label,coverage_role,source_mode,monitoring_purposes,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,'Unverified',?,?)",(project_id,name,device_kind,vendor,host,port,protocol,stream,coverage,source_mode,purposes_json,now,now)); record_audit(con,actor,"created","device",cur.lastrowid,"Source registered; connection not verified",project_id); queue_automation_event(con,project_id,"device.created","device",cur.lastrowid,"unverified"); con.commit()
                return self.send_json(201,{"id":cur.lastrowid,"status":"Unverified"})
            if path == "/api/people":
                if not self.require_permission(sess,"manage_people"): return
                name=text(data,"display_name",120); category=choice(data,"category",("Staff","Manager","Watchlist")); position=text(data,"position",120,required=False); ref=text(data,"record_reference",120,required=False); owner=text(data,"record_owner",120); purpose=text(data,"purpose",500); review=text(data,"review_date",10)
                if not all((name,owner,purpose,review)): return self.send_json(400,{"error":"Name, record owner, purpose, and review date are required"})
                if category == "Watchlist": status="Pending review"
                else: status="Pending review"
                now=utcnow(); cur=con.execute("INSERT INTO people(project_id,display_name,category,position,record_reference,record_owner,purpose,review_date,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(project_id,name,category,position,ref,owner,purpose,review,status,now,now)); record_audit(con,actor,"created","person",cur.lastrowid,f"{category} record pending review",project_id); queue_automation_event(con,project_id,"person.created","person",cur.lastrowid,"pending_review"); con.commit()
                return self.send_json(201,{"id":cur.lastrowid,"status":status})
            if path == "/api/alerts":
                if not self.require_permission(sess,"review_alerts"): return
                source=text(data,"source",160); summary=text(data,"summary",500); occurred=text(data,"occurred_at",40,required=False) or utcnow()
                if not source or not summary: return self.send_json(400,{"error":"Source and summary are required"})
                now=utcnow(); cur=con.execute("INSERT INTO alerts(project_id,source,occurred_at,summary,status,created_at,updated_at) VALUES(?,?,?,?,'Needs review',?,?)",(project_id,source,occurred,summary,now,now)); record_audit(con,actor,"created","alert",cur.lastrowid,"Manual review item; not an AI detection",project_id); queue_automation_event(con,project_id,"alert.created","alert",cur.lastrowid,"needs_review"); con.commit()
                return self.send_json(201,{"id":cur.lastrowid,"status":"Needs review"})
            return self.send_json(404,{"error":"Not found"})
        except (ValueError,TypeError) as e:
            return self.send_json(400,{"error":str(e)})
        except sqlite3.IntegrityError:
            return self.send_json(400,{"error":"The record values are invalid"})
        finally: con.close()

    def import_rejected(self, con: sqlite3.Connection, sess: sqlite3.Row, project_id: int, code: str, status: int, message: str, row_errors: list[dict] | None=None, error_count: int=0):
        record_audit(con,sess["username"],"rejected","people_import",None,f"reason={code}; error_count={error_count}",project_id)
        queue_automation_event(con,project_id,"people.import_rejected","people_import",None,f"reason={code}; error_count={error_count}");con.commit()
        result={"error":message}
        if row_errors is not None:result.update({"row_errors":row_errors,"error_count":error_count})
        return self.send_json(status,result)

    def import_people(self, con: sqlite3.Connection, sess: sqlite3.Row) -> None:
        if not self.require_permission(sess,"manage_people"): return
        project_id=selected_project(self,con)
        if not project_id:return self.send_json(400,{"error":"Select a valid CCTV project"})
        if "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" not in self.headers.get("Content-Type",""):
            return self.import_rejected(con,sess,project_id,"invalid_file_type",415,"Upload an .xlsx workbook")
        try:size=int(self.headers.get("Content-Length","0"))
        except ValueError:size=0
        if size<1 or size>MAX_IMPORT_BYTES:return self.import_rejected(con,sess,project_id,"file_size_limit",413,"Workbook must be between 1 byte and 20 MB")
        try:
            book=load_workbook(io.BytesIO(self.rfile.read(size)),read_only=True,data_only=False)
            if "People" not in book.sheetnames:return self.import_rejected(con,sess,project_id,"missing_people_sheet",400,"Workbook needs a sheet named People. Download the template for the required columns.")
            sheet=book["People"];iterator=sheet.iter_rows(values_only=False)
            header_cells=next(iterator,None)
            headers=[str(cell.value).strip() if cell.value is not None else "" for cell in (header_cells or [])]
            if tuple(headers) not in (PEOPLE_COLUMNS,LEGACY_PEOPLE_COLUMNS):return self.import_rejected(con,sess,project_id,"invalid_columns",400,"People sheet columns must match the current or legacy template exactly and in order: "+", ".join(PEOPLE_COLUMNS))
            import_columns=tuple(headers)
            prepared=[];errors=[];error_count=0;refs=set()
            for row_number,cells in enumerate(iterator,start=2):
                if row_number>MAX_IMPORT_ROWS+1:return self.import_rejected(con,sess,project_id,"row_limit",413,f"Import limit is {MAX_IMPORT_ROWS:,} rows")
                values=[cell.value for cell in cells[:len(import_columns)]]
                if all(value is None or str(value).strip()=="" for value in values):continue
                if any(cell.data_type=="f" for cell in cells[:len(import_columns)]):
                    error_count+=1
                    if len(errors)<200:errors.append({"row":row_number,"error":"Formulas are not allowed; paste values only"})
                    continue
                item={};row_errors=[]
                for index,key in enumerate(import_columns):
                    value=values[index] if index<len(values) else None
                    if key=="review_date" and hasattr(value,"strftime"):
                        value=value.strftime("%Y-%m-%d")
                    if value is None:value=""
                    if not isinstance(value,str):row_errors.append(key+" must be text (review_date must be YYYY-MM-DD)");continue
                    value=value.strip();item[key]=value
                for key,limit in (("display_name",120),("position",120),("record_reference",120),("record_owner",120),("purpose",500)):
                    if key in item and len(item[key])>limit:row_errors.append(key+f" exceeds {limit} characters")
                if item.get("category") not in ("Staff","Manager","Watchlist"):row_errors.append("category must be Staff, Manager, or Watchlist")
                if not item.get("display_name"):row_errors.append("display_name is required")
                if not item.get("record_owner"):row_errors.append("record_owner is required")
                if not item.get("purpose"):row_errors.append("purpose is required")
                try:datetime.strptime(item.get("review_date",""),"%Y-%m-%d")
                except ValueError:row_errors.append("review_date must use YYYY-MM-DD")
                ref=item.get("record_reference","")
                if ref and ref in refs:row_errors.append("record_reference duplicates another row in this workbook")
                elif ref:refs.add(ref)
                if row_errors:
                    error_count+=1
                    if len(errors)<200:errors.append({"row":row_number,"error":"; ".join(row_errors)})
                else:prepared.append({**item,"_row":row_number})
                if len(prepared)+len(errors)>MAX_IMPORT_ROWS:return self.send_json(413,{"error":f"Import limit is {MAX_IMPORT_ROWS:,} rows"})
            book.close()
            if error_count:return self.import_rejected(con,sess,project_id,"row_validation",422,"No rows were imported. Correct the listed rows and upload again.",errors,error_count)
            if not prepared:return self.import_rejected(con,sess,project_id,"empty_sheet",400,"The People sheet contains no data rows")
            con.execute("BEGIN IMMEDIATE")
            existing={r[0] for r in con.execute("SELECT record_reference FROM people WHERE project_id=? AND record_reference!=''",(project_id,)).fetchall()}
            dupes=[item["_row"] for item in prepared if item["record_reference"] and item["record_reference"] in existing]
            if dupes:
                con.rollback();return self.import_rejected(con,sess,project_id,"duplicate_reference",422,"No rows were imported. record_reference already exists in this project.",[{"row":n,"error":"record_reference already exists in this project"} for n in dupes[:200]],len(dupes))
        except Exception as error:
            if isinstance(error,(ValueError,TypeError)):return self.import_rejected(con,sess,project_id,"invalid_workbook",400,"Could not read workbook. Save it as a standard .xlsx file and try again.")
            if error.__class__.__module__.startswith("openpyxl"):return self.import_rejected(con,sess,project_id,"invalid_workbook",400,"Could not read workbook. Save it as a standard .xlsx file and try again.")
            raise
        now=utcnow()
        try:
            for item in prepared:
                con.execute("INSERT INTO people(project_id,display_name,category,position,record_reference,record_owner,purpose,review_date,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?, ?,'Pending review',?,?)",(project_id,item["display_name"],item["category"],item.get("position",""),item["record_reference"],item["record_owner"],item["purpose"],item["review_date"],now,now))
            record_audit(con,sess["username"],"imported","people",None,f"rows={len(prepared)}; status=Pending review",project_id)
            queue_automation_event(con,project_id,"people.imported","people",None,f"rows={len(prepared)}; status=pending_review")
            con.commit()
        except sqlite3.IntegrityError:
            con.rollback();return self.send_json(409,{"error":"Import could not be committed because a record conflicts with current project data. No rows were imported."})
        return self.send_json(201,{"imported":len(prepared),"status":"Pending review","message":f"Imported {len(prepared):,} records. Every record is Pending review."})

    def first_run_setup(self, con: sqlite3.Connection) -> None:
        # First account can only be created while the server is loopback-bound,
        # from a loopback client, and with a same-origin browser request.
        try:
            if not ipaddress.ip_address(self.client_address[0]).is_loopback:
                return self.send_json(403,{"error":"First-run setup is available only on this computer"})
        except ValueError:
            return self.send_json(403,{"error":"First-run setup is available only on this computer"})
        origin=self.headers.get("Origin","").rstrip("/")
        expected=f"http://{self.headers.get('Host','') }".rstrip("/")
        if not origin or not hmac.compare_digest(origin,expected):
            return self.send_json(403,{"error":"Open the local app directly to finish setup"})
        try: data=self.read_json()
        except ValueError as e:return self.send_json(400,{"error":str(e)})
        username=text(data,"username",80); password=text(data,"password",200)
        if len(password)<14:return self.send_json(400,{"error":"Use a password with at least 14 characters"})
        try:
            con.execute("BEGIN IMMEDIATE")
            if con.execute("SELECT 1 FROM users LIMIT 1").fetchone():
                con.rollback(); return self.send_json(409,{"error":"Administrator setup has already been completed"})
            salt=secrets.token_bytes(16); digest=hashlib.pbkdf2_hmac("sha256",password.encode(),salt,PBKDF2_ROUNDS)
            cur=con.execute("INSERT INTO users(username,password_salt,password_hash,permissions,created_at) VALUES(?,?,?,?,?)",(username,salt,digest,json.dumps(list(PERMISSIONS),separators=(",",":")),utcnow()))
            project_id=con.execute("SELECT id FROM projects ORDER BY id LIMIT 1").fetchone()[0]
            record_audit(con,username,"created","user",cur.lastrowid,"Initial local administrator created")
            queue_automation_event(con,project_id,"auth.admin_created","user",cur.lastrowid,"created");con.commit()
            return self.send_json(201,{"ok":True})
        except sqlite3.IntegrityError:
            con.rollback(); return self.send_json(400,{"error":"That administrator username is already in use"})

    def do_PATCH(self) -> None:
        path=urlparse(self.path).path
        con=db()
        try:
            sess=self.session(con)
            if not sess: return self.send_json(401,{"error":"Sign in required"})
            if not self.require_csrf(sess): return self.send_json(403,{"error":"Request verification failed"})
            try: data=self.read_json()
            except ValueError as e: return self.send_json(400,{"error":str(e)})
            parts=path.strip("/").split("/")
            if len(parts)!=3 or parts[0]!="api" or not parts[2].isdigit(): return self.send_json(404,{"error":"Not found"})
            kind,ident=parts[1],int(parts[2]); actor=sess["username"]; now=utcnow()
            if kind=="users":
                if not self.require_permission(sess,"manage_users"): return
                permissions=valid_permissions(data.get("permissions"));active=data.get("active")
                if not isinstance(active,bool):return self.send_json(400,{"error":"Account active must be true or false"})
                target=con.execute("SELECT id,username,active,permissions FROM users WHERE id=?",(ident,)).fetchone()
                if not target:return self.send_json(404,{"error":"Account not found"})
                if target["username"]==actor and not active:return self.send_json(400,{"error":"You cannot disable your own account"})
                retains_admin=active and "manage_users" in permissions
                other_admin=con.execute("SELECT 1 FROM users WHERE active=1 AND id!=? AND instr(permissions,'\"manage_users\"')>0 LIMIT 1",(ident,)).fetchone()
                if not (retains_admin or other_admin):return self.send_json(400,{"error":"At least one active account must retain user-management permission"})
                con.execute("UPDATE users SET permissions=?,active=? WHERE id=?",(json.dumps(permissions,separators=(",",":")),1 if active else 0,ident))
                if not active:con.execute("DELETE FROM sessions WHERE user_id=?",(ident,))
                record_audit(con,actor,"updated","user",ident,f"active={active}; permissions={','.join(permissions)}")
                project_for_event=selected_project(self,con) or con.execute("SELECT id FROM projects ORDER BY id LIMIT 1").fetchone()[0]
                queue_automation_event(con,project_for_event,"user.updated","user",ident,"active" if active else "disabled");con.commit()
                return self.send_json(200,{"ok":True})
            project_id=selected_project(self,con)
            if not project_id:return self.send_json(400,{"error":"Select a valid CCTV project"})
            if kind=="people":
                if not self.require_permission(sess,"manage_people"): return
                status=choice(data,"status",("Active","Inactive")); row=con.execute("SELECT category FROM people WHERE id=? AND project_id=?",(ident,project_id)).fetchone()
                if not row:return self.send_json(404,{"error":"Person record not found"})
                if status=="Active" and row["category"]=="Watchlist" and not text(data,"approval_note",500):return self.send_json(400,{"error":"A documented approval note is required to activate a watchlist record"})
                con.execute("UPDATE people SET status=?,updated_at=? WHERE id=? AND project_id=?",(status,now,ident,project_id)); record_audit(con,actor,"status_changed","person",ident,f"status={status}",project_id); queue_automation_event(con,project_id,"person.status_changed","person",ident,status.lower());
            elif kind=="devices":
                if not self.require_permission(sess,"manage_devices"): return
                device_kind=choice(data,"device_kind",("DVR","NVR","Camera","Unknown"));row=con.execute("SELECT id FROM devices WHERE id=? AND project_id=?",(ident,project_id)).fetchone()
                if not row:return self.send_json(404,{"error":"Device not found"})
                con.execute("UPDATE devices SET device_kind=?,updated_at=? WHERE id=? AND project_id=?",(device_kind,now,ident,project_id));record_audit(con,actor,"classified","device",ident,f"device_kind={device_kind}; classification is administrative, not automatic",project_id);queue_automation_event(con,project_id,"device.classified","device",ident,device_kind.lower())
            elif kind=="alerts":
                if not self.require_permission(sess,"review_alerts"): return
                status=choice(data,"status",("Needs review","Reviewed","Dismissed","Escalated")); note=text(data,"review_note",500,required=False); row=con.execute("SELECT id,status FROM alerts WHERE id=? AND project_id=?",(ident,project_id)).fetchone()
                if not row:return self.send_json(404,{"error":"Alert not found"})
                con.execute("UPDATE alerts SET status=?,reviewer=?,review_note=?,updated_at=? WHERE id=? AND project_id=?",(status,actor,note,now,ident,project_id)); changed=status!=row["status"]; record_audit(con,actor,"reviewed" if changed else "note_updated","alert",ident,f"status={status}; note_updated",project_id); queue_automation_event(con,project_id,"alert.reviewed" if changed else "alert.note_updated","alert",ident,status.lower())
            else:return self.send_json(404,{"error":"Not found"})
            con.commit(); return self.send_json(200,{"ok":True})
        except (ValueError,TypeError) as e:return self.send_json(400,{"error":str(e)})
        finally:con.close()

    def login(self, con: sqlite3.Connection) -> None:
        try: data=self.read_json()
        except ValueError as e:return self.send_json(400,{"error":str(e)})
        username=text(data,"username",80); password=text(data,"password",200)
        now=time.time(); failures=[t for t in RATE_LIMIT.get(username,[]) if now-t<900]; RATE_LIMIT[username]=failures
        if len(failures)>=5:return self.send_json(429,{"error":"Too many attempts. Wait 15 minutes and try again."})
        row=con.execute("SELECT * FROM users WHERE username=?",(username,)).fetchone()
        valid=False
        if row:
            digest=hashlib.pbkdf2_hmac("sha256",password.encode(),row["password_salt"],PBKDF2_ROUNDS); valid=hmac.compare_digest(digest,row["password_hash"])
        if not valid:
            RATE_LIMIT.setdefault(username,[]).append(now)
            con.execute("INSERT INTO audit_log(actor,action,object_type,detail,created_at) VALUES(?,?,?,?,?)",(username or "unknown","login_failed","session","Invalid credentials",utcnow()));con.commit()
            return self.send_json(401,{"error":"Username or password is incorrect"})
        RATE_LIMIT.pop(username,None); token=secrets.token_urlsafe(32); csrf=secrets.token_urlsafe(24); expires=int(now)+SESSION_SECONDS
        con.execute("DELETE FROM sessions WHERE expires_at<=?",(int(now),)); con.execute("INSERT INTO sessions(token_hash,user_id,csrf_token,expires_at,created_at) VALUES(?,?,?,?,?)",(hashlib.sha256(token.encode()).hexdigest(),row["id"],csrf,expires,utcnow())); record_audit(con,username,"login","session",None)
        project_id=con.execute("SELECT id FROM projects ORDER BY id LIMIT 1").fetchone()[0];queue_automation_event(con,project_id,"auth.login_succeeded","session",None,"ok");con.commit()
        secure="; Secure" if os.environ.get("CASINO_OPS_HTTPS")=="1" else ""
        return self.send_json(200,{"username":username,"csrf_token":csrf},{"Set-Cookie":f"ops_session={token}; HttpOnly; SameSite=Strict; Path=/; Max-Age={SESSION_SECONDS}{secure}"})

def text(obj: dict,key: str,limit: int,required: bool=True) -> str:
    value=obj.get(key,"")
    if not isinstance(value,str): raise ValueError(f"{key} must be text")
    value=value.strip()
    if required and not value: raise ValueError(f"{key} is required")
    if len(value)>limit: raise ValueError(f"{key} is too long")
    return value

def integer(obj: dict,key: str,low: int,high: int) -> int:
    try: value=int(obj.get(key))
    except (ValueError,TypeError): raise ValueError(f"{key} must be a number")
    if not low<=value<=high: raise ValueError(f"{key} must be between {low} and {high}")
    return value

def choice(obj: dict,key: str,allowed: tuple[str,...]) -> str:
    value=text(obj,key,80)
    if value not in allowed: raise ValueError(f"Invalid {key}")
    return value

def main() -> None:
    parser=argparse.ArgumentParser(); parser.add_argument("command",choices=("serve","init-admin"),nargs="?",default="serve"); args=parser.parse_args()
    db().close()
    if args.command=="init-admin": return init_admin()
    if HOST not in ("127.0.0.1","localhost","::1"):
        raise SystemExit("This starter is loopback-only. Do not expose it to a network; production deployment requires TLS and an approved access-control layer.")
    print(f"Casino Ops local console listening at http://{HOST}:{PORT} (this device only)")
    ThreadingHTTPServer((HOST,PORT),Handler).serve_forever()

if __name__=="__main__": main()
