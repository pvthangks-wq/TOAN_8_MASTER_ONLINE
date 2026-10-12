"""TOAN 8 MASTER -- FastAPI/SQLite; production data, no mocked scores.
Run: python app.py init; python app.py create-admin; uvicorn app:app --host 0.0.0.0 --port 8000
"""
import argparse
import csv
import hashlib
import hmac
import io
import json
import os
import re
import secrets
import sqlite3
import tempfile
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape as xml_escape

from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError, VerificationError
from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer

from catalog import CHAPTERS, VERIFIED_NOTES, catalog_json

HERE = Path(__file__).resolve().parent
DEFAULT_DB = os.environ.get("DB_PATH", str(HERE / "instance" / "toan8.db"))
COOKIE_SECURE = os.environ.get("COOKIE_SECURE", "0") == "1"
COOKIE_NAME = "toan8_session"
PH = PasswordHasher(time_cost=2, memory_cost=19456, parallelism=2)
ALLOWED_KIND = {"single", "true_false", "short", "essay"}
SQL = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL COLLATE NOCASE, full_name TEXT NOT NULL, role TEXT NOT NULL CHECK(role IN ('admin','teacher','student')), password_hash TEXT NOT NULL, disabled INTEGER NOT NULL DEFAULT 0, must_change_password INTEGER DEFAULT 0, created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS sessions(token_hash TEXT PRIMARY KEY, csrf TEXT NOT NULL, user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE, expires_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS login_failures(key TEXT PRIMARY KEY, count INTEGER NOT NULL, last_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS classes(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, school_year TEXT NOT NULL, teacher_id INTEGER NOT NULL REFERENCES users(id), created_at TEXT NOT NULL, UNIQUE(name,school_year,teacher_id));
CREATE TABLE IF NOT EXISTS enrollments(class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE, student_id INTEGER NOT NULL REFERENCES users(id), joined_at TEXT NOT NULL, PRIMARY KEY(class_id,student_id));
CREATE TABLE IF NOT EXISTS questions(id INTEGER PRIMARY KEY AUTOINCREMENT, teacher_id INTEGER NOT NULL REFERENCES users(id), lesson_id INTEGER NOT NULL CHECK(lesson_id BETWEEN 1 AND 39), kind TEXT NOT NULL, prompt TEXT NOT NULL, options_json TEXT NOT NULL DEFAULT '[]', answer TEXT NOT NULL DEFAULT '', explanation TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS assignments(id INTEGER PRIMARY KEY AUTOINCREMENT, class_id INTEGER NOT NULL REFERENCES classes(id), teacher_id INTEGER NOT NULL REFERENCES users(id), lesson_id INTEGER NOT NULL CHECK(lesson_id BETWEEN 1 AND 39), title TEXT NOT NULL, opens_at TEXT, due_at TEXT, duration_minutes INTEGER NOT NULL DEFAULT 0, show_answers INTEGER NOT NULL DEFAULT 1, show_score INTEGER NOT NULL DEFAULT 1, allow_hints INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS assignment_questions(assignment_id INTEGER NOT NULL REFERENCES assignments(id) ON DELETE CASCADE, question_id INTEGER NOT NULL REFERENCES questions(id), position INTEGER NOT NULL, points REAL NOT NULL, PRIMARY KEY(assignment_id,question_id));
CREATE TABLE IF NOT EXISTS assignment_recipients(assignment_id INTEGER NOT NULL REFERENCES assignments(id) ON DELETE CASCADE, student_id INTEGER NOT NULL REFERENCES users(id), PRIMARY KEY(assignment_id,student_id));
CREATE TABLE IF NOT EXISTS attempts(id INTEGER PRIMARY KEY AUTOINCREMENT, assignment_id INTEGER NOT NULL REFERENCES assignments(id), student_id INTEGER NOT NULL REFERENCES users(id), attempt_no INTEGER NOT NULL DEFAULT 1, status TEXT NOT NULL CHECK(status IN ('in_progress','needs_review','graded')), started_at TEXT NOT NULL, submitted_at TEXT, score REAL, max_score REAL, updated_at TEXT NOT NULL, UNIQUE(assignment_id,student_id,attempt_no));
CREATE TABLE IF NOT EXISTS attempt_answers(attempt_id INTEGER NOT NULL REFERENCES attempts(id) ON DELETE CASCADE, question_id INTEGER NOT NULL REFERENCES questions(id), answer TEXT NOT NULL DEFAULT '', points REAL, feedback TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL, PRIMARY KEY(attempt_id,question_id));
CREATE TABLE IF NOT EXISTS lesson_progress(student_id INTEGER NOT NULL REFERENCES users(id), lesson_id INTEGER NOT NULL CHECK(lesson_id BETWEEN 1 AND 39), status TEXT NOT NULL CHECK(status IN ('studying','learned','review')), updated_at TEXT NOT NULL, PRIMARY KEY(student_id,lesson_id));
CREATE TABLE IF NOT EXISTS activity(user_id INTEGER PRIMARY KEY REFERENCES users(id), last_at TEXT NOT NULL, last_heartbeat TEXT, active_seconds INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY AUTOINCREMENT, actor_id INTEGER REFERENCES users(id), action TEXT NOT NULL, target TEXT NOT NULL, created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS games(id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER NOT NULL REFERENCES users(id), game TEXT NOT NULL, correct INTEGER NOT NULL, total INTEGER NOT NULL, created_at TEXT NOT NULL);
CREATE INDEX IF NOT EXISTS idx_attempt_student ON attempts(student_id,assignment_id);
CREATE INDEX IF NOT EXISTS idx_assignment_class ON assignments(class_id);
CREATE INDEX IF NOT EXISTS idx_recipient_student ON assignment_recipients(student_id);
CREATE INDEX IF NOT EXISTS idx_enroll_student ON enrollments(student_id);
"""

def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")

def iso(raw):
    if not raw: return None
    try:
        d = datetime.fromisoformat(str(raw).replace('Z','+00:00'))
        if d.tzinfo is None: raise ValueError()
        return d.astimezone(timezone.utc).isoformat(timespec='seconds')
    except (ValueError,TypeError): raise HTTPException(400,"Ngày giờ phải có múi giờ hợp lệ.")

def textval(s, maxlen=200, required=True):
    if not isinstance(s,str): raise HTTPException(400,"Dữ liệu văn bản không hợp lệ.")
    s=s.strip()
    if required and not s: raise HTTPException(400,"Vui lòng điền thông tin bắt buộc.")
    if len(s)>maxlen: raise HTTPException(400,"Nội dung quá dài.")
    return s

def integer(v,name="Giá trị",minimum=1,maximum=1000000):
    if isinstance(v,bool): raise HTTPException(400,f"{name} không hợp lệ.")
    try:
        value=int(v)
        if str(value)!=str(v): raise ValueError()
    except (TypeError,ValueError): raise HTTPException(400,f"{name} không hợp lệ.")
    if not minimum<=value<=maximum: raise HTTPException(400,f"{name} ngoài giới hạn.")
    return value

def j(x): return json.dumps(x,ensure_ascii=False,separators=(',',':'))
def asdict(row): return dict(row) if row else None

def fail(status,msg): raise HTTPException(status, msg)


def create_app(db_path: Optional[str]=None):
    filename = str(Path(db_path or DEFAULT_DB).resolve())
    Path(filename).parent.mkdir(parents=True,exist_ok=True)

    @contextmanager
    def db():
        con=sqlite3.connect(filename,timeout=15)
        con.row_factory=sqlite3.Row
        con.execute('PRAGMA foreign_keys=ON')
        con.execute('PRAGMA busy_timeout=15000')
        try:
            yield con
            con.commit()
        except:
            con.rollback()
            raise
        finally: con.close()

    with db() as con:
        con.executescript(SQL)
        con.execute('PRAGMA journal_mode=WAL')

    app=FastAPI(title="Toán 8 Master", docs_url=None,redoc_url=None,openapi_url=None)

    @app.middleware("http")
    async def common_headers(request,call_next):
        response=await call_next(request)
        response.headers['X-Content-Type-Options']='nosniff'
        response.headers['Referrer-Policy']='no-referrer'
        response.headers['X-Frame-Options']='DENY'
        response.headers['Permissions-Policy']='camera=(), microphone=(), geolocation=()'
        response.headers['Content-Security-Policy']=("default-src 'self'; script-src 'self' https://cdn.jsdelivr.net; style-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net https://fonts.googleapis.com; font-src 'self' https://cdn.jsdelivr.net https://fonts.gstatic.com; connect-src 'self'; img-src 'self' data:; frame-ancestors 'none'")
        if COOKIE_SECURE: response.headers['Strict-Transport-Security']='max-age=31536000; includeSubDomains'
        if request.url.path.startswith('/api/'):response.headers['Cache-Control']='no-store'
        return response

    def log(con,actor,action,target):
        con.execute('INSERT INTO audit(actor_id,action,target,created_at) VALUES (?,?,?,?)',(actor,action,str(target),now()))

    def auth(request,roles=None,write=False):
        token=request.cookies.get(COOKIE_NAME,'')
        if not token: fail(401,"Chưa đăng nhập.")
        th=hashlib.sha256(token.encode()).hexdigest()
        with db() as con:
            row=con.execute('''SELECT u.id,u.username,u.full_name,u.role,u.disabled,u.must_change_password,s.csrf,s.expires_at
            FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token_hash=?''',(th,)).fetchone()
        if not row or row['disabled'] or row['expires_at']<=now(): fail(401,"Phiên đăng nhập đã hết hạn.")
        if roles and row['role'] not in roles: fail(403,"Không có quyền truy cập.")
        if write and not hmac.compare_digest(request.headers.get('x-csrf-token',''),row['csrf']):fail(403,"Yêu cầu thiếu mã bảo vệ phiên đăng nhập.")
        # Không chặn đổi mật khẩu hoặc đăng xuất khi yêu cầu đổi mật khẩu lần đầu.
        if row['must_change_password'] and request.url.path not in ('/api/me','/api/password','/api/logout'):
            fail(403,"Cần đổi mật khẩu trước khi tiếp tục.")
        return row

    def verify_teacher_class(con,user,cid):
        obj=con.execute('SELECT * FROM classes WHERE id=? AND teacher_id=?',(cid,user['id'])).fetchone()
        if not obj:fail(403,"Lớp không thuộc quyền quản lý của giáo viên.")
        return obj

    def verify_teacher_student(con,user,sid):
        row=con.execute('''SELECT 1 FROM enrollments e JOIN classes c ON e.class_id=c.id WHERE c.teacher_id=? AND e.student_id=?''',(user['id'],sid)).fetchone()
        if not row:fail(403,"Học sinh không thuộc lớp được phân quyền.")

    def question_public(q):
        d=dict(q)
        return {'id':d['id'],'lesson_id':d['lesson_id'],'kind':d['kind'],'prompt':d['prompt'],'options':json.loads(d['options_json'])}

    def question_teacher(q):
        d=question_public(q)
        d.update(answer=q['answer'],explanation=q['explanation'],created_at=q['created_at'])
        return d

    def access_attempt(con,user,attempt_id):
        a=con.execute('''SELECT at.*,ass.title,ass.due_at,ass.duration_minutes,ass.show_answers,ass.show_score,ass.allow_hints,ass.lesson_id
          FROM attempts at JOIN assignments ass ON at.assignment_id=ass.id WHERE at.id=?''',(attempt_id,)).fetchone()
        if not a or a['student_id']!=user['id']:fail(403,"Không có quyền xem lượt làm bài này.")
        return a

    def attempt_payload(con,at):
        qs=con.execute('''SELECT q.*,aq.points,ans.answer AS student_answer,ans.points AS awarded,ans.feedback
        FROM assignment_questions aq JOIN questions q ON aq.question_id=q.id
        LEFT JOIN attempt_answers ans ON ans.question_id=q.id AND ans.attempt_id=?
        WHERE aq.assignment_id=? ORDER BY aq.position''',(at['id'],at['assignment_id'])).fetchall()
        submitted=at['status']!='in_progress'
        result=[]
        for q in qs:
            public=question_public(q)
            public.update(points=q['points'],student_answer=q['student_answer'] or '')
            if submitted and at['show_answers']:
                public.update(correct_answer=q['answer'],explanation=q['explanation'],awarded=q['awarded'] if at['show_score'] else None,feedback=q['feedback'])
            elif submitted:
                public.update(awarded=q['awarded'] if at['show_score'] else None)
            result.append(public)
        return {'attempt':{'id':at['id'],'assignment_id':at['assignment_id'],'title':at['title'],
          'lesson_id':at['lesson_id'],'status':at['status'],'started_at':at['started_at'],
          'submitted_at':at['submitted_at'],'score':at['score'] if (submitted and at['show_score']) else None,
          'max_score':at['max_score'],'due_at':at['due_at'],'duration_minutes':at['duration_minutes'],
          'show_answers':bool(at['show_answers']),'show_score':bool(at['show_score'])},'questions':result}

    @app.get('/api/health')
    def health():
        # Fail health checks if persistent storage is unavailable.
        with db() as con:
            con.execute('SELECT id FROM users LIMIT 1').fetchone()
        return {'status':'ok','app':'TOÁN 8 MASTER'}

    def setup_ready(con):
        return con.execute("SELECT NOT EXISTS (SELECT 1 FROM users WHERE role='admin')").fetchone()[0] == 1

    @app.get('/api/setup/status')
    def setup_status():
        # There is no public registration: this endpoint is only for the first admin.
        token=os.environ.get('BOOTSTRAP_TOKEN','')
        with db() as con:
            first_admin_missing=setup_ready(con)
        return {'setup_available':first_admin_missing and len(token)>=32,
                'admin_exists':not first_admin_missing,
                'token_configured':len(token)>=32}

    @app.post('/api/setup/admin')
    async def setup_admin(request:Request):
        token=os.environ.get('BOOTSTRAP_TOKEN','')
        if len(token)<32:fail(403,'Thiếu mã khởi tạo an toàn ở cấu hình máy chủ.')
        # No default credentials, no anonymous admin creation without this server secret.
        raw=await request.json()
        provided=raw.get('setup_token','')
        if not isinstance(provided,str) or len(provided)>1024:fail(400,'Mã khởi tạo không hợp lệ.')
        username=textval(raw.get('username',''),60).lower()
        name=textval(raw.get('full_name',''),120)
        password=raw.get('password','')
        if not re.fullmatch(r'[a-z0-9_.-]{3,60}',username):fail(400,'Tên đăng nhập chỉ gồm chữ thường, số, dấu chấm, gạch nối, gạch dưới.')
        if not isinstance(password,str) or not 12<=len(password)<=128:fail(400,'Mật khẩu phải từ 12 đến 128 ký tự.')
        ip=request.client.host if request.client else 'unknown'
        key='setup|'+hashlib.sha256(ip.encode()).hexdigest()
        incorrect_token=False
        with db() as con:
            fail_row=con.execute('SELECT count,last_at FROM login_failures WHERE key=?',(key,)).fetchone()
            recent=bool(fail_row and datetime.fromisoformat(fail_row['last_at'])>datetime.now(timezone.utc)-timedelta(minutes=15))
            if recent and fail_row['count']>=8:fail(429,'Quá nhiều lần nhập sai. Thử lại sau 15 phút.')
            if not hmac.compare_digest(provided,token):
                count=(fail_row['count']+1) if recent else 1
                con.execute('INSERT INTO login_failures(key,count,last_at) VALUES(?,?,?) ON CONFLICT(key) DO UPDATE SET count=excluded.count,last_at=excluded.last_at',(key,count,now()))
                incorrect_token=True
        if incorrect_token:fail(403,'Mã khởi tạo không đúng.')
        with db() as con:
            # A write lock guarantees only one first administrator can be created.
            con.execute('BEGIN IMMEDIATE')
            if not setup_ready(con):fail(409,'Hệ thống đã có quản trị viên. Trang khởi tạo đã đóng.')
            try:
                cur=con.execute("INSERT INTO users(username,full_name,role,password_hash,created_at) VALUES (?,?,'admin',?,?)",(username,name,PH.hash(password),now()))
            except sqlite3.IntegrityError:fail(409,'Tên đăng nhập đã tồn tại.')
            log(con,cur.lastrowid,'initial_admin_setup','admin')
        return {'ok':True,'message':'Quản trị viên đầu tiên đã được tạo. Có thể đăng nhập.'}

    @app.get('/api/catalog')
    def catalog(request:Request):
        auth(request)
        return catalog_json()

    @app.get('/api/me')
    def me(request:Request):
        u=auth(request)
        return {'id':u['id'],'username':u['username'],'full_name':u['full_name'],'role':u['role'],
                'csrf':u['csrf'],'must_change_password':bool(u['must_change_password']),
                'ai_enabled':bool(os.environ.get('OPENAI_API_KEY'))}

    @app.post('/api/login')
    async def login(request:Request,response:Response):
        data=await request.json()
        username=textval(data.get('username',''),80).lower()
        pw=data.get('password','')
        if not isinstance(pw,str) or len(pw)>1024:fail(400,"Mật khẩu không hợp lệ.")
        key=hashlib.sha256((username+'|'+(request.client.host if request.client else '')).encode()).hexdigest()
        invalid=False
        with db() as con:
            f=con.execute('SELECT * FROM login_failures WHERE key=?',(key,)).fetchone()
            if f and f['count']>=7 and datetime.fromisoformat(f['last_at'])>datetime.now(timezone.utc)-timedelta(minutes=15):
                fail(429,"Đăng nhập sai nhiều lần. Thử lại sau 15 phút.")
            u=con.execute('SELECT * FROM users WHERE username=?',(username,)).fetchone()
            success=False
            if u and not u['disabled']:
                try: success=PH.verify(u['password_hash'],pw)
                except (VerifyMismatchError,VerificationError): pass
            if not success:
                n=(f['count']+1) if f and datetime.fromisoformat(f['last_at'])>datetime.now(timezone.utc)-timedelta(minutes=15) else 1
                con.execute('INSERT INTO login_failures(key,count,last_at) VALUES (?,?,?) ON CONFLICT(key) DO UPDATE SET count=excluded.count,last_at=excluded.last_at',(key,n,now()))
                invalid=True
            else:
                con.execute('DELETE FROM login_failures WHERE key=?',(key,))
                token=secrets.token_urlsafe(48);csrf=secrets.token_urlsafe(32)
                exp=(datetime.now(timezone.utc)+timedelta(days=7)).isoformat(timespec='seconds')
                con.execute('INSERT INTO sessions(token_hash,csrf,user_id,expires_at) VALUES (?,?,?,?)',(hashlib.sha256(token.encode()).hexdigest(),csrf,u['id'],exp))
                con.execute('INSERT INTO activity(user_id,last_at) VALUES (?,?) ON CONFLICT(user_id) DO UPDATE SET last_at=excluded.last_at',(u['id'],now()))
                log(con,u['id'],'login','session')
        if invalid:fail(401,"Tên đăng nhập hoặc mật khẩu không đúng.")
        response.set_cookie(COOKIE_NAME,token,httponly=True,secure=COOKIE_SECURE,samesite='lax',max_age=604800,path='/')
        return {'id':u['id'],'username':u['username'],'full_name':u['full_name'],'role':u['role'],'csrf':csrf,'must_change_password':bool(u['must_change_password']),'ai_enabled':bool(os.environ.get('OPENAI_API_KEY'))}

    @app.post('/api/logout')
    def logout(request:Request,response:Response):
        u=auth(request,write=True)
        tok=request.cookies.get(COOKIE_NAME,'')
        with db() as con:
            con.execute('DELETE FROM sessions WHERE token_hash=?',(hashlib.sha256(tok.encode()).hexdigest(),))
            log(con,u['id'],'logout','session')
        response.delete_cookie(COOKIE_NAME,path='/')
        return {'ok':True}

    @app.post('/api/password')
    async def password(request:Request):
        u=auth(request,write=True);d=await request.json()
        old=d.get('old_password','');new=d.get('new_password','')
        if not isinstance(new,str) or len(new)<12 or len(new)>128:fail(400,"Mật khẩu mới cần có ít nhất 12 ký tự.")
        with db() as con:
            rec=con.execute('SELECT password_hash FROM users WHERE id=?',(u['id'],)).fetchone()
            try: PH.verify(rec['password_hash'],old)
            except (VerifyMismatchError,VerificationError): fail(400,"Mật khẩu cũ không đúng.")
            con.execute('UPDATE users SET password_hash=?,must_change_password=0 WHERE id=?',(PH.hash(new),u['id']))
            th=hashlib.sha256(request.cookies[COOKIE_NAME].encode()).hexdigest()
            con.execute('DELETE FROM sessions WHERE user_id=? AND token_hash<>?',(u['id'],th))
            log(con,u['id'],'change_password',u['id'])
        return {'ok':True}

    @app.get('/api/admin/users')
    def admin_users(request:Request):
        u=auth(request,['admin'])
        with db() as con:
            teachers=[dict(r) for r in con.execute("SELECT id,username,full_name,disabled,created_at FROM users WHERE role='teacher' ORDER BY id DESC")]
            cls=[dict(r) for r in con.execute('SELECT c.id,c.name,c.school_year,u.full_name AS teacher FROM classes c JOIN users u ON u.id=c.teacher_id ORDER BY c.id DESC')]
        return {'teachers':teachers,'classes':cls}

    @app.post('/api/admin/teachers')
    async def admin_create_teacher(request:Request):
        u=auth(request,['admin'],True);d=await request.json()
        username=textval(d.get('username',''),60).lower();name=textval(d.get('full_name',''),120)
        if not re.fullmatch(r'[a-z0-9_.-]{3,60}',username):fail(400,"Tên đăng nhập dùng chữ thường, số, dấu chấm, gạch dưới hoặc gạch nối.")
        pw=secrets.token_urlsafe(14)
        with db() as con:
            try:cur=con.execute("INSERT INTO users(username,full_name,role,password_hash,must_change_password,created_at) VALUES (?,?,'teacher',?,1,?)",(username,name,PH.hash(pw),now()))
            except sqlite3.IntegrityError:fail(409,"Tên đăng nhập đã tồn tại.")
            log(con,u['id'],'create_teacher',cur.lastrowid)
        return {'id':cur.lastrowid,'username':username,'full_name':name,'temporary_password':pw,'notice':'Mật khẩu chỉ trả về một lần; giáo viên phải đổi khi đăng nhập.'}

    @app.post('/api/admin/teachers/{tid}/toggle')
    def admin_toggle(request:Request,tid:int):
        u=auth(request,['admin'],True)
        with db() as con:
            r=con.execute("SELECT disabled FROM users WHERE id=? AND role='teacher'",(tid,)).fetchone()
            if not r:fail(404,"Không tìm thấy giáo viên.")
            disabled=0 if r['disabled'] else 1
            con.execute('UPDATE users SET disabled=? WHERE id=?',(disabled,tid))
            if disabled:con.execute('DELETE FROM sessions WHERE user_id=?',(tid,))
            log(con,u['id'],'toggle_teacher',tid)
        return {'disabled':bool(disabled)}

    @app.post('/api/admin/teachers/{tid}/reset')
    def admin_reset_teacher(request:Request,tid:int):
        u=auth(request,['admin'],True)
        with db() as con:
            target=con.execute("SELECT id FROM users WHERE role='teacher' AND id=?",(tid,)).fetchone()
            if not target:fail(404,"Không tìm thấy tài khoản.")
            pw=secrets.token_urlsafe(14)
            con.execute('UPDATE users SET password_hash=?,must_change_password=1 WHERE id=?',(PH.hash(pw),tid))
            con.execute('DELETE FROM sessions WHERE user_id=?',(tid,))
            log(con,u['id'],'reset_teacher',tid)
        return {'temporary_password':pw}

    @app.get('/api/admin/backup')
    def admin_backup(request:Request):
        u=auth(request,['admin'])
        with db() as con:
            file= tempfile.NamedTemporaryFile(suffix='.db',delete=False)
            file.close()
            dest=sqlite3.connect(file.name)
            try:con.backup(dest)
            finally:dest.close()
            log(con,u['id'],'backup','database')
        from starlette.background import BackgroundTask
        return FileResponse(file.name,filename='toan8-backup-'+datetime.now(timezone.utc).strftime('%Y%m%d')+'.db',media_type='application/octet-stream',background=BackgroundTask(lambda:Path(file.name).unlink(missing_ok=True)))

    @app.get('/api/teacher/classes')
    def classes(request:Request):
        u=auth(request,['teacher'])
        with db() as con:
            return {'classes':[dict(r) for r in con.execute('''SELECT c.*,count(e.student_id) AS students FROM classes c LEFT JOIN enrollments e ON e.class_id=c.id WHERE c.teacher_id=? GROUP BY c.id ORDER BY c.id DESC''',(u['id'],))]}

    @app.post('/api/teacher/classes')
    async def create_class(request:Request):
        u=auth(request,['teacher'],True);d=await request.json()
        name=textval(d.get('name',''),50);year=textval(d.get('school_year',''),30)
        with db() as con:
            try:cur=con.execute('INSERT INTO classes(name,school_year,teacher_id,created_at) VALUES (?,?,?,?)',(name,year,u['id'],now()))
            except sqlite3.IntegrityError:fail(409,"Lớp này đã được tạo trong năm học.")
            log(con,u['id'],'create_class',cur.lastrowid)
        return {'id':cur.lastrowid,'name':name,'school_year':year}

    def create_students(con,actor,cid,rows):
        if not isinstance(rows,list) or not 1<=len(rows)<=120:fail(400,"Một lần nhập từ 1 đến 120 học sinh.")
        out=[]
        usernames=set()
        for d in rows:
            if not isinstance(d,dict):fail(400,"Danh sách học sinh không hợp lệ.")
            name=textval(d.get('full_name',''),120)
            username=textval(d.get('username',''),60).lower()
            if not re.fullmatch(r'[a-z0-9_.-]{3,60}',username):fail(400,"Tên đăng nhập không hợp lệ: "+username)
            if username in usernames:fail(409,"Tên đăng nhập bị lặp trong danh sách: "+username)
            usernames.add(username)
            pw=secrets.token_urlsafe(14)
            try: cur=con.execute("INSERT INTO users(username,full_name,role,password_hash,must_change_password,created_at) VALUES (?,?,'student',?,1,?)",(username,name,PH.hash(pw),now()))
            except sqlite3.IntegrityError:fail(409,"Tên đăng nhập đã tồn tại: "+username)
            sid=cur.lastrowid
            con.execute('INSERT INTO enrollments(class_id,student_id,joined_at) VALUES (?,?,?)',(cid,sid,now()))
            out.append({'id':sid,'full_name':name,'username':username,'temporary_password':pw})
            log(con,actor,'create_student',sid)
        return out

    @app.get('/api/teacher/classes/{cid}/students')
    def class_students(request:Request,cid:int):
        u=auth(request,['teacher'])
        with db() as con:
            verify_teacher_class(con,u,cid)
            return {'students':[dict(r) for r in con.execute('''SELECT us.id,us.username,us.full_name,us.disabled,us.created_at,a.last_at,COALESCE(a.active_seconds,0) AS active_seconds
                FROM enrollments e JOIN users us ON us.id=e.student_id LEFT JOIN activity a ON a.user_id=us.id
                WHERE e.class_id=? ORDER BY us.full_name''',(cid,))]}

    @app.post('/api/teacher/classes/{cid}/students')
    async def new_student(request:Request,cid:int):
        u=auth(request,['teacher'],True);d=await request.json()
        with db() as con:
            verify_teacher_class(con,u,cid)
            rows=create_students(con,u['id'],cid,[d])
        return {'credentials':rows}

    @app.post('/api/teacher/classes/{cid}/students/import')
    async def import_students(request:Request,cid:int):
        u=auth(request,['teacher'],True);d=await request.json()
        with db() as con:
            verify_teacher_class(con,u,cid)
            rows=create_students(con,u['id'],cid,d.get('students'))
        return {'credentials':rows}

    @app.post('/api/teacher/students/{sid}/reset')
    def reset_student(request:Request,sid:int):
        u=auth(request,['teacher'],True)
        with db() as con:
            verify_teacher_student(con,u,sid)
            pw=secrets.token_urlsafe(14)
            con.execute("UPDATE users SET password_hash=?,must_change_password=1 WHERE id=? AND role='student'",(PH.hash(pw),sid))
            con.execute('DELETE FROM sessions WHERE user_id=?',(sid,))
            log(con,u['id'],'reset_student_password',sid)
        return {'temporary_password':pw}

    @app.post('/api/teacher/students/{sid}/toggle')
    def toggle_student(request:Request,sid:int):
        u=auth(request,['teacher'],True)
        with db() as con:
            verify_teacher_student(con,u,sid)
            r=con.execute("SELECT disabled FROM users WHERE id=? AND role='student'",(sid,)).fetchone()
            if not r:fail(404,"Không tìm thấy học sinh.")
            state=0 if r['disabled'] else 1
            con.execute('UPDATE users SET disabled=? WHERE id=?',(state,sid))
            if state: con.execute('DELETE FROM sessions WHERE user_id=?',(sid,))
            log(con,u['id'],'toggle_student',sid)
        return {'disabled':bool(state)}

    @app.get('/api/teacher/questions')
    def questions(request:Request,lesson_id:Optional[int]=None):
        u=auth(request,['teacher'])
        with db() as con:
            if lesson_id and 1<=lesson_id<=39:
                rows=con.execute('SELECT * FROM questions WHERE teacher_id=? AND lesson_id=? ORDER BY id DESC',(u['id'],lesson_id))
            else:rows=con.execute('SELECT * FROM questions WHERE teacher_id=? ORDER BY id DESC',(u['id'],))
            return {'questions':[question_teacher(q) for q in rows]}

    @app.post('/api/teacher/questions')
    async def create_question(request:Request):
        u=auth(request,['teacher'],True);d=await request.json()
        lesson=integer(d.get('lesson_id'),'Bài học',1,39)
        kind=d.get('kind')
        if kind not in ALLOWED_KIND:fail(400,"Loại câu hỏi không hỗ trợ.")
        prompt=textval(d.get('prompt',''),3000)
        options=[];answer=''
        if kind=='single':
            options=d.get('options')
            if not isinstance(options,list) or len(options)<2 or len(options)>6:fail(400,"Trắc nghiệm cần từ 2 đến 6 đáp án.")
            options=[textval(o,400) for o in options]
            answer=str(integer(d.get('correct'),'Đáp án',0,len(options)-1))
        elif kind=='true_false':
            if d.get('correct') not in (True,False):fail(400,"Đúng/sai cần đáp án chuẩn.")
            answer='true' if d['correct'] else 'false'
        explanation=textval(d.get('explanation',''),3000,required=False)
        with db() as con:
            cur=con.execute('''INSERT INTO questions(teacher_id,lesson_id,kind,prompt,options_json,answer,explanation,created_at)
                   VALUES (?,?,?,?,?,?,?,?)''',(u['id'],lesson,kind,prompt,j(options),answer,explanation,now()))
            log(con,u['id'],'create_question',cur.lastrowid)
        return {'id':cur.lastrowid,'kind':kind}

    @app.delete('/api/teacher/questions/{qid}')
    def delete_question(request:Request,qid:int):
        u=auth(request,['teacher'],True)
        with db() as con:
            row=con.execute('SELECT 1 FROM questions WHERE id=? AND teacher_id=?',(qid,u['id'])).fetchone()
            if not row:fail(404,"Không tìm thấy câu hỏi.")
            used=con.execute('SELECT 1 FROM assignment_questions WHERE question_id=? LIMIT 1',(qid,)).fetchone()
            if used:fail(409,"Câu hỏi đã giao nên không được xóa để bảo toàn kết quả.")
            con.execute('DELETE FROM questions WHERE id=?',(qid,))
        return {'ok':True}

    @app.post('/api/teacher/assignments')
    async def make_assignment(request:Request):
        u=auth(request,['teacher'],True);d=await request.json()
        cid=integer(d.get('class_id'),'Lớp')
        lesson=integer(d.get('lesson_id'),'Bài học',1,39)
        title=textval(d.get('title',''),160)
        ids=d.get('question_ids',[])
        if not isinstance(ids,list) or len(ids)<1 or len(ids)>60:fail(400,"Chọn từ 1 đến 60 câu hỏi.")
        qids=[integer(x,'Câu hỏi') for x in ids]
        if len(set(qids))!=len(qids):fail(400,"Câu hỏi không được lặp.")
        students=d.get('student_ids')
        if students is not None and (not isinstance(students,list) or len(students)<1):fail(400,"Danh sách người nhận phải có học sinh.")
        opens=iso(d.get('opens_at'));due=iso(d.get('due_at'))
        if opens and due and opens>=due:fail(400,"Thời gian mở phải trước hạn nộp.")
        dur=integer(d.get('duration_minutes',0),'Thời gian làm bài',0,240)
        with db() as con:
            verify_teacher_class(con,u,cid)
            placeholders=','.join('?' for _ in qids)
            got=list(con.execute('SELECT id,lesson_id FROM questions WHERE teacher_id=? AND id IN ('+placeholders+')',(u['id'],*qids)))
            if len(got)!=len(qids) or any(x['lesson_id']!=lesson for x in got):fail(400,"Câu hỏi phải thuộc bài đang chọn và do giáo viên tạo.")
            enrolled={row['student_id'] for row in con.execute('SELECT student_id FROM enrollments WHERE class_id=?',(cid,))}
            if not enrolled:fail(400,"Lớp chưa có học sinh.")
            recipients=[integer(v,'Học sinh') for v in students] if students is not None else sorted(enrolled)
            if not recipients or len(recipients)!=len(set(recipients)) or not set(recipients)<=enrolled:fail(403,"Danh sách người nhận không thuộc lớp.")
            cur=con.execute('''INSERT INTO assignments(class_id,teacher_id,lesson_id,title,opens_at,due_at,duration_minutes,show_answers,show_score,allow_hints,created_at)
             VALUES (?,?,?,?,?,?,?,?,?,?,?)''',(cid,u['id'],lesson,title,opens,due,dur,int(bool(d.get('show_answers',True))),int(bool(d.get('show_score',True))),int(bool(d.get('allow_hints',False))),now()))
            aid=cur.lastrowid
            for pos,qid in enumerate(qids):
                con.execute('INSERT INTO assignment_questions VALUES (?,?,?,?)',(aid,qid,pos,1.0))
            for sid in recipients:con.execute('INSERT INTO assignment_recipients VALUES (?,?)',(aid,sid))
            log(con,u['id'],'assign',aid)
        return {'id':aid,'recipients':len(recipients),'total_questions':len(qids)}

    @app.get('/api/teacher/assignments')
    def teacher_assignments(request:Request,class_id:Optional[int]=None):
        u=auth(request,['teacher'])
        with db() as con:
            if class_id:verify_teacher_class(con,u,class_id)
            sql='''SELECT a.*,c.name AS class_name,
            (SELECT count(*) FROM assignment_recipients r WHERE r.assignment_id=a.id) AS assigned,
            (SELECT count(*) FROM attempts t WHERE t.assignment_id=a.id AND t.status!='in_progress') AS submitted,
            (SELECT round(avg(t.score),2) FROM attempts t WHERE t.assignment_id=a.id AND t.status='graded') AS mean_score
            FROM assignments a JOIN classes c ON c.id=a.class_id WHERE a.teacher_id=?'''
            args=[u['id']]
            if class_id:sql+=' AND a.class_id=?';args.append(class_id)
            sql+=' ORDER BY a.id DESC'
            return {'assignments':[dict(r) for r in con.execute(sql,args)]}

    @app.get('/api/teacher/assignments/{aid}/status')
    def assignment_status(request:Request,aid:int):
        u=auth(request,['teacher'])
        with db() as con:
            assignment=con.execute('SELECT a.*,c.name AS class_name FROM assignments a JOIN classes c ON c.id=a.class_id WHERE a.id=? AND a.teacher_id=?',(aid,u['id'])).fetchone()
            if not assignment:fail(403,"Không có quyền xem bài giao.")
            rows=[dict(r) for r in con.execute('''SELECT ur.id,ur.full_name,ur.username,t.id AS attempt_id,t.status AS attempt_status,t.started_at,t.submitted_at,t.score,t.max_score
            FROM assignment_recipients ar JOIN users ur ON ur.id=ar.student_id LEFT JOIN attempts t ON t.student_id=ur.id AND t.assignment_id=ar.assignment_id
            WHERE ar.assignment_id=? ORDER BY ur.full_name''',(aid,))]
            for r in rows:
                if r['attempt_status']=='graded' or r['attempt_status']=='needs_review':r['display_status']='submitted'
                elif assignment['due_at'] and assignment['due_at']<now():r['display_status']='late'
                elif r['attempt_status']=='in_progress':r['display_status']='in_progress'
                else:r['display_status']='not_started'
            return {'assignment':dict(assignment),'recipients':rows}

    @app.get('/api/teacher/attempts/{atid}')
    def teacher_review(request:Request,atid:int):
        u=auth(request,['teacher'])
        with db() as con:
            at=con.execute('''SELECT t.*,a.title,a.teacher_id,s.full_name AS student_name
            FROM attempts t JOIN assignments a ON a.id=t.assignment_id JOIN users s ON s.id=t.student_id WHERE t.id=?''',(atid,)).fetchone()
            if not at or at['teacher_id']!=u['id']:fail(403,"Không có quyền xem bài làm.")
            qs=[dict(q) for q in con.execute('''SELECT q.id,q.kind,q.prompt,q.answer,q.explanation,aq.points AS max_points,ans.answer AS student_answer,ans.points AS awarded,ans.feedback
            FROM assignment_questions aq JOIN questions q ON q.id=aq.question_id LEFT JOIN attempt_answers ans ON ans.question_id=q.id AND ans.attempt_id=?
            WHERE aq.assignment_id=? ORDER BY aq.position''',(atid,at['assignment_id']))]
            return {'attempt':dict(at),'questions':qs}

    @app.post('/api/teacher/attempts/{atid}/grade')
    async def grade_attempt(request:Request,atid:int):
        u=auth(request,['teacher'],True);d=await request.json()
        scores=d.get('scores',{})
        if not isinstance(scores,dict):fail(400,"Dữ liệu chấm không hợp lệ.")
        with db() as con:
            at=con.execute('''SELECT at.*,a.teacher_id FROM attempts at JOIN assignments a ON a.id=at.assignment_id WHERE at.id=?''',(atid,)).fetchone()
            if not at or at['teacher_id']!=u['id']:fail(403,"Không có quyền chấm bài.")
            if at['status']=='in_progress':fail(409,"Học sinh chưa nộp bài.")
            qrows=list(con.execute('''SELECT q.id,q.kind,aq.points,ans.answer,ans.points AS awarded FROM assignment_questions aq JOIN questions q ON q.id=aq.question_id
            JOIN attempt_answers ans ON ans.question_id=q.id AND ans.attempt_id=? WHERE aq.assignment_id=?''',(atid,at['assignment_id'])))
            valid={str(q['id']):q for q in qrows}
            if set(scores)-set(valid):fail(400,"Có câu hỏi không thuộc bài này.")
            for key,val in scores.items():
                q=valid[key]
                if not isinstance(val,dict):fail(400,"Dữ liệu điểm không hợp lệ.")
                try:score=float(val['points'])
                except (TypeError,ValueError,KeyError):fail(400,"Điểm chấm không hợp lệ.")
                if not 0<=score<=q['points']:fail(400,"Điểm ngoài thang của câu hỏi.")
                feedback=textval(val.get('feedback',''),1200,False)
                con.execute('UPDATE attempt_answers SET points=?,feedback=?,updated_at=? WHERE attempt_id=? AND question_id=?',(score,feedback,now(),atid,int(key)))
            tally=con.execute('SELECT sum(points) AS score,count(*) AS n,count(points) AS scored FROM attempt_answers WHERE attempt_id=?',(atid,)).fetchone()
            state='graded' if tally['n']==tally['scored'] else 'needs_review'
            con.execute('UPDATE attempts SET score=?,status=?,updated_at=? WHERE id=?',(tally['score'] or 0,state,now(),atid))
            log(con,u['id'],'grade_attempt',atid)
        return {'status':state,'score':tally['score'] or 0}

    @app.get('/api/student/assignments')
    def student_assignments(request:Request):
        u=auth(request,['student'])
        with db() as con:
            rows=con.execute('''SELECT a.id,a.title,a.lesson_id,a.opens_at,a.due_at,a.duration_minutes,a.show_answers,a.show_score,c.name AS class_name,
            (SELECT count(*) FROM assignment_questions aq WHERE aq.assignment_id=a.id) AS count,
            at.id AS attempt_id,at.status AS attempt_status,at.score,at.max_score
            FROM assignment_recipients ar JOIN assignments a ON ar.assignment_id=a.id JOIN classes c ON c.id=a.class_id
            LEFT JOIN attempts at ON at.assignment_id=a.id AND at.student_id=ar.student_id
            WHERE ar.student_id=? ORDER BY a.id DESC''',(u['id'],))
            result=[]
            for r in rows:
                item=dict(r)
                if item['attempt_status']=='in_progress': item['status']='in_progress'
                elif item['attempt_status']:item['status']='submitted'
                elif item['due_at'] and item['due_at']<now():item['status']='late'
                else:item['status']='not_started'
                if not item['show_score']:item['score']=None
                result.append(item)
            return {'assignments':result}

    @app.post('/api/student/assignments/{aid}/start')
    def student_start(request:Request,aid:int):
        u=auth(request,['student'],True)
        with db() as con:
            con.execute('BEGIN IMMEDIATE')
            a=con.execute('''SELECT a.* FROM assignments a JOIN assignment_recipients r ON r.assignment_id=a.id WHERE a.id=? AND r.student_id=?''',(aid,u['id'])).fetchone()
            if not a:fail(403,"Bài tập chưa được giao cho học sinh này.")
            old=con.execute('SELECT id FROM attempts WHERE assignment_id=? AND student_id=?',(aid,u['id'])).fetchone()
            if not old:
                t=now()
                if a['opens_at'] and a['opens_at']>t:fail(403,"Chưa đến thời gian mở bài.")
                if a['due_at'] and a['due_at']<t:fail(403,"Bài tập đã hết hạn.")
                score=con.execute('SELECT sum(points) AS s FROM assignment_questions WHERE assignment_id=?',(aid,)).fetchone()['s']
                cur=con.execute("INSERT INTO attempts(assignment_id,student_id,status,started_at,updated_at,max_score) VALUES (?,?,'in_progress',?,?,?)",(aid,u['id'],t,t,score))
                atid=cur.lastrowid
                log(con,u['id'],'start_attempt',atid)
            else:atid=old['id']
            at=access_attempt(con,u,atid)
            return attempt_payload(con,at)

    @app.get('/api/student/attempts/{atid}')
    def student_attempt(request:Request,atid:int):
        u=auth(request,['student'])
        with db() as con:
            at=access_attempt(con,u,atid)
            return attempt_payload(con,at)

    def validate_answers(con,atid,assignment_id,answer_map):
        if not isinstance(answer_map,dict):fail(400,"Câu trả lời không hợp lệ.")
        qs=list(con.execute('''SELECT q.id,q.kind,q.options_json FROM assignment_questions aq JOIN questions q ON q.id=aq.question_id WHERE aq.assignment_id=?''',(assignment_id,)))
        kinds={str(q['id']):q for q in qs}
        if set(answer_map)-set(kinds):fail(400,"Câu hỏi không có trong bài tập.")
        for k,v in answer_map.items():
            if not isinstance(v,str) or len(v)>10000:fail(400,"Nội dung câu trả lời không hợp lệ.")
            kind=kinds[k]['kind']
            if kind=='single' and v and (not v.isdigit() or int(v)>=len(json.loads(kinds[k]['options_json']))):fail(400,"Đáp án trắc nghiệm không hợp lệ.")
            if kind=='true_false' and v not in ('true','false',''):fail(400,"Đáp án đúng/sai không hợp lệ.")
            con.execute('''INSERT INTO attempt_answers(attempt_id,question_id,answer,updated_at) VALUES (?,?,?,?)
            ON CONFLICT(attempt_id,question_id) DO UPDATE SET answer=excluded.answer,updated_at=excluded.updated_at''',(atid,int(k),v,now()))
        return qs

    def may_work(at, for_submit=False):
        if at['status']!='in_progress':fail(409,"Bài đã nộp, không được sửa.")
        grace=timedelta(seconds=20) if for_submit else timedelta(0)
        current=datetime.now(timezone.utc)
        if at['due_at'] and current>datetime.fromisoformat(at['due_at'])+grace:
            fail(409,"Đã hết hạn nộp bài; hãy liên hệ giáo viên.")
        if at['duration_minutes']:
            deadline=datetime.fromisoformat(at['started_at'])+timedelta(minutes=at['duration_minutes'])
            if current>deadline+grace:fail(409,"Đã hết thời gian làm bài; hãy liên hệ giáo viên.")

    @app.put('/api/student/attempts/{atid}/answers')
    async def save_answers(request:Request,atid:int):
        u=auth(request,['student'],True);d=await request.json()
        with db() as con:
            con.execute('BEGIN IMMEDIATE')
            at=access_attempt(con,u,atid);may_work(at)
            validate_answers(con,atid,at['assignment_id'],d.get('answers',{}))
            con.execute('UPDATE attempts SET updated_at=? WHERE id=?',(now(),atid))
        return {'saved':True,'saved_at':now()}

    @app.post('/api/student/attempts/{atid}/submit')
    async def submit(request:Request,atid:int):
        u=auth(request,['student'],True);d=await request.json()
        with db() as con:
            con.execute('BEGIN IMMEDIATE')
            at=access_attempt(con,u,atid)
            if at['status']!='in_progress':return {'already_submitted':True,'result':attempt_payload(con,at)}
            may_work(at,for_submit=True)
            validate_answers(con,atid,at['assignment_id'],d.get('answers',{}))
            qs=list(con.execute('''SELECT q.id,q.kind,q.answer,aq.points FROM assignment_questions aq JOIN questions q ON q.id=aq.question_id WHERE aq.assignment_id=? ORDER BY aq.position''',(at['assignment_id'],)))
            needs_review=False;total=0
            for q in qs:
                ansrow=con.execute('SELECT answer FROM attempt_answers WHERE attempt_id=? AND question_id=?',(atid,q['id'])).fetchone()
                ans=ansrow['answer'] if ansrow else ''
                if q['kind'] in ('single','true_false'):
                    points=q['points'] if ans and ans==q['answer'] else 0
                elif ans.strip():
                    points=None;needs_review=True
                else:points=0
                if points is not None:total+=points
                con.execute('''INSERT INTO attempt_answers(attempt_id,question_id,answer,points,updated_at) VALUES (?,?,?,?,?)
                ON CONFLICT(attempt_id,question_id) DO UPDATE SET points=excluded.points,updated_at=excluded.updated_at''',(atid,q['id'],ans,points,now()))
            status='needs_review' if needs_review else 'graded'
            con.execute('UPDATE attempts SET status=?,score=?,submitted_at=?,updated_at=? WHERE id=?',(status,total,now(),now(),atid))
            log(con,u['id'],'submit_attempt',atid)
            fresh=access_attempt(con,u,atid)
            return {'already_submitted':False,'result':attempt_payload(con,fresh)}

    @app.post('/api/student/lessons/{lesson_id}/progress')
    async def mark_lesson(request:Request,lesson_id:int):
        u=auth(request,['student'],True)
        integer(lesson_id,'Bài học',1,39)
        d=await request.json();state=d.get('status')
        if state not in ('studying','learned','review'):fail(400,"Trạng thái không hợp lệ.")
        with db() as con:
            con.execute('''INSERT INTO lesson_progress(student_id,lesson_id,status,updated_at) VALUES (?,?,?,?)
            ON CONFLICT(student_id,lesson_id) DO UPDATE SET status=excluded.status,updated_at=excluded.updated_at''',(u['id'],lesson_id,state,now()))
        return {'ok':True}

    @app.get('/api/student/progress')
    def student_progress(request:Request):
        u=auth(request,['student'])
        with db() as con:
            items=[dict(r) for r in con.execute('SELECT lesson_id,status,updated_at FROM lesson_progress WHERE student_id=?',(u['id'],))]
            games=[dict(r) for r in con.execute('SELECT game,correct,total,created_at FROM games WHERE student_id=? ORDER BY id DESC LIMIT 20',(u['id'],))]
        return {'progress':items,'games':games}

    @app.get('/api/student/dashboard')
    def student_dashboard(request:Request):
        u=auth(request,['student'])
        with db() as con:
            cls=[dict(x) for x in con.execute('SELECT c.id,c.name,c.school_year FROM enrollments e JOIN classes c ON c.id=e.class_id WHERE e.student_id=?',(u['id'],))]
            count=con.execute('SELECT count(*) AS n FROM lesson_progress WHERE student_id=? AND status="learned"',(u['id'],)).fetchone()['n']
            scores=[dict(x) for x in con.execute('''SELECT at.id,at.score,at.max_score,at.submitted_at,a.title,a.show_score,at.status FROM attempts at
             JOIN assignments a ON at.assignment_id=a.id WHERE at.student_id=? AND at.status!='in_progress' ORDER BY at.submitted_at DESC LIMIT 12''',(u['id'],))]
            for s in scores:
                if not s['show_score']:s['score']=None
            a=con.execute('SELECT COALESCE(active_seconds,0) AS seconds,last_at FROM activity WHERE user_id=?',(u['id'],)).fetchone()
        return {'classes':cls,'learned':count,'scores':scores,'activity':dict(a) if a else {'seconds':0,'last_at':None}}

    @app.post('/api/activity/heartbeat')
    def heartbeat(request:Request):
        u=auth(request,write=True)
        t=datetime.now(timezone.utc)
        with db() as con:
            a=con.execute('SELECT last_heartbeat,active_seconds FROM activity WHERE user_id=?',(u['id'],)).fetchone()
            delta=0
            if a and a['last_heartbeat']:
                elapsed=(t-datetime.fromisoformat(a['last_heartbeat'])).total_seconds()
                if 12<=elapsed<=75:delta=min(int(elapsed),35)
                if elapsed<12:return {'recorded_seconds':0}
            con.execute('''INSERT INTO activity(user_id,last_at,last_heartbeat,active_seconds) VALUES (?,?,?,?)
              ON CONFLICT(user_id) DO UPDATE SET last_at=excluded.last_at,last_heartbeat=excluded.last_heartbeat,active_seconds=activity.active_seconds+?''',
              (u['id'],now(),now(),0,delta))
        return {'recorded_seconds':delta}

    @app.get('/api/teacher/dashboard')
    def teacher_dashboard(request:Request,class_id:Optional[int]=None):
        u=auth(request,['teacher'])
        with db() as con:
            if class_id:verify_teacher_class(con,u,class_id)
            ids=[r['id'] for r in con.execute('SELECT id FROM classes WHERE teacher_id=?'+(' AND id=?' if class_id else ''),(u['id'],class_id) if class_id else (u['id'],))]
            if not ids:return {'summary':{'students':0,'assignments':0,'submitted':0,'unsubmitted':0,'avg_score':None,'completion_pct':0},'students':[]}
            clauses=','.join('?' for _ in ids)
            students=[dict(r) for r in con.execute('''SELECT DISTINCT us.id,us.full_name,us.username,ac.last_at,COALESCE(ac.active_seconds,0) AS active_seconds
             FROM enrollments e JOIN users us ON us.id=e.student_id LEFT JOIN activity ac ON ac.user_id=us.id
             WHERE e.class_id IN ('''+clauses+') ORDER BY us.full_name',ids)]
            arows=[dict(r) for r in con.execute('SELECT id FROM assignments WHERE class_id IN ('+clauses+')',ids)]
            assids=[a['id'] for a in arows]
            vals=[]
            for s in students:
                stats=con.execute('''SELECT count(*) AS assigned,
                sum(CASE WHEN at.status IS NOT NULL AND at.status!='in_progress' THEN 1 ELSE 0 END) AS completed,
                avg(CASE WHEN at.status='graded' AND at.max_score>0 THEN (at.score*10.0/at.max_score) END) AS mean_score
                FROM assignment_recipients ar JOIN assignments a ON a.id=ar.assignment_id
                LEFT JOIN attempts at ON at.assignment_id=ar.assignment_id AND at.student_id=ar.student_id
                WHERE ar.student_id=? AND a.class_id IN ('''+clauses+')',(s['id'],*ids)).fetchone()
                lp=con.execute('SELECT count(*) AS n FROM lesson_progress WHERE student_id=? AND status="learned"',(s['id'],)).fetchone()['n']
                n=stats['assigned'] or 0;k=stats['completed'] or 0
                s.update(assigned=n,completed=k,missing=n-k,mean_score=round(stats['mean_score'],2) if stats['mean_score'] is not None else None,
                         learned=lp,progress=round(100*k/n) if n else 0)
                vals.append(s)
            count=sum(s['assigned'] for s in students);done=sum(s['completed'] for s in students)
            grades=[s['mean_score'] for s in students if s['mean_score'] is not None]
            return {'summary':{'students':len(students),'assignments':len(assids),'submitted':done,'unsubmitted':count-done,
              'avg_score':round(sum(grades)/len(grades),2) if grades else None,'completion_pct':round(100*done/count) if count else 0},'students':vals}

    @app.get('/api/teacher/students/{sid}/profile')
    def student_profile(request:Request,sid:int):
        u=auth(request,['teacher'])
        with db() as con:
            verify_teacher_student(con,u,sid)
            user=con.execute('SELECT id,full_name,username,created_at FROM users WHERE id=?',(sid,)).fetchone()
            progresses=[dict(r) for r in con.execute('SELECT lesson_id,status,updated_at FROM lesson_progress WHERE student_id=?',(sid,))]
            history=[dict(r) for r in con.execute('''SELECT at.id,at.status,at.score,at.max_score,at.started_at,at.submitted_at,a.title,a.lesson_id
            FROM attempts at JOIN assignments a ON a.id=at.assignment_id WHERE at.student_id=? AND a.teacher_id=? ORDER BY at.id DESC''',(sid,u['id']))]
            a=con.execute('SELECT last_at,active_seconds FROM activity WHERE user_id=?',(sid,)).fetchone()
            wrong=[dict(r) for r in con.execute('''SELECT q.lesson_id,q.id AS question_id,q.prompt,count(*) AS wrong_count
                FROM attempt_answers ans JOIN attempts at ON ans.attempt_id=at.id
                JOIN assignments a ON at.assignment_id=a.id JOIN questions q ON ans.question_id=q.id
                WHERE at.student_id=? AND a.teacher_id=? AND at.status!='in_progress' AND ans.points=0 AND TRIM(ans.answer)!=''
                GROUP BY q.id ORDER BY wrong_count DESC LIMIT 15''',(sid,u['id']))]
        return {'student':dict(user),'activity':dict(a) if a else None,'progress':progresses,'history':history,'weaknesses':wrong}

    def report_rows(con,u,cid):
        c=verify_teacher_class(con,u,cid)
        students=[dict(r) for r in con.execute('SELECT us.id,us.full_name,us.username FROM enrollments e JOIN users us ON us.id=e.student_id WHERE e.class_id=? ORDER BY us.full_name',(cid,))]
        asgs=[dict(r) for r in con.execute('SELECT id,title,due_at FROM assignments WHERE class_id=? ORDER BY id',(cid,))]
        rows=[]
        for s in students:
            scores=[];done=0
            for a in asgs:
                recipient=con.execute('SELECT 1 FROM assignment_recipients WHERE assignment_id=? AND student_id=?',(a['id'],s['id'])).fetchone()
                if not recipient:scores.append('—');continue
                att=con.execute('SELECT status,score,max_score FROM attempts WHERE assignment_id=? AND student_id=?',(a['id'],s['id'])).fetchone()
                if not att or att['status']=='in_progress':scores.append('Chưa nộp')
                else:
                    done+=1
                    if att['status']=='needs_review':scores.append('Chờ chấm')
                    else:scores.append(f"{(att['score'] or 0)*10/att['max_score']:.1f}" if att['max_score'] else '0')
            denom=sum(1 for v in scores if v!='—')
            rows.append({'name':s['full_name'],'username':s['username'],'done':done,'assigned':denom,'rate':round(100*done/denom) if denom else 0,'scores':scores})
        return c,asgs,rows

    @app.get('/api/teacher/classes/{cid}/report.xlsx')
    def report_xlsx(request:Request,cid:int):
        u=auth(request,['teacher'])
        with db() as con:c,asgs,rows=report_rows(con,u,cid)
        wb=Workbook();ws=wb.active;ws.title='Báo cáo lớp'
        ws.append(['TOÁN 8 MASTER — '+c['name']+' — '+c['school_year']]);ws.append(['Cập nhật UTC: '+now()]);ws.append([])
        headers=['STT','Họ tên','Tên đăng nhập','Đã nộp','Được giao','Hoàn thành (%)']+[a['title'] for a in asgs]
        ws.append(headers)
        for i,r in enumerate(rows,1):ws.append([i,r['name'],r['username'],r['done'],r['assigned'],r['rate'],*r['scores']])
        for cell in ws[4]:cell.fill=PatternFill('solid',fgColor='143B70');cell.font=Font(color='FFFFFF',bold=True);cell.alignment=Alignment(wrap_text=True)
        ws.freeze_panes='D5';ws.auto_filter.ref=f'A4:{get_column_letter(len(headers))}{max(4,ws.max_row)}'
        for idx,w in enumerate([7,28,24,12,12,17]+[27]*len(asgs),1):ws.column_dimensions[get_column_letter(idx)].width=w
        for row in ws.iter_rows(min_row=1):
            for cell in row:
                if isinstance(cell.value,str) and cell.value[:1] in ('=','+','-','@'):cell.value="'"+cell.value
        bio=io.BytesIO();wb.save(bio);bio.seek(0)
        headers_http={'Content-Disposition':f'attachment; filename="toan8_lop_{cid}.xlsx"'}
        return StreamingResponse(bio,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers=headers_http)

    @app.get('/api/teacher/classes/{cid}/report.pdf')
    def report_pdf(request:Request,cid:int):
        u=auth(request,['teacher'])
        with db() as con:c,asgs,rows=report_rows(con,u,cid)
        fontfile='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
        boldfile='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
        if not Path(fontfile).exists():fail(503,"Máy chủ cần cài fonts-dejavu-core để xuất PDF tiếng Việt.")
        if 'VNDejaVu' not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont('VNDejaVu',fontfile))
            pdfmetrics.registerFont(TTFont('VNDejaVu-Bold',boldfile))
        buf=io.BytesIO()
        doc=SimpleDocTemplate(buf,pagesize=A4,rightMargin=17*mm,leftMargin=17*mm,topMargin=17*mm,bottomMargin=17*mm)
        styles=getSampleStyleSheet()
        heading=ParagraphStyle('vnHeading',parent=styles['Heading1'],fontName='VNDejaVu-Bold',fontSize=15,leading=21,textColor=colors.HexColor('#143B70'),spaceAfter=10)
        normal=ParagraphStyle('vnBody',fontName='VNDejaVu',fontSize=8.1,leading=12)
        cell=ParagraphStyle('vnCell',fontName='VNDejaVu',fontSize=7,leading=10)
        cells=[[Paragraph(v,cell) for v in ['STT','Học sinh','Đã nộp','Được giao','Tỉ lệ']]]
        for i,r in enumerate(rows,1):cells.append([Paragraph(str(v),cell) for v in [i,xml_escape(r['name']),r['done'],r['assigned'],str(r['rate'])+'%']])
        grid=Table(cells,colWidths=[14*mm,76*mm,24*mm,27*mm,29*mm],repeatRows=1,hAlign='LEFT')
        grid.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#DDEDFB')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('BOTTOMPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),8),('LINEBELOW',(0,0),(-1,-1),0.3,colors.HexColor('#E2E9F2'))]))
        story=[Paragraph('TOÁN 8 MASTER — BÁO CÁO LỚP',heading),Paragraph(xml_escape(f"Lớp: {c['name']}   |   Năm học: {c['school_year']}"),normal),Paragraph('Ngày xuất (UTC): '+now(),normal),Spacer(1,8*mm),grid,Spacer(1,8*mm),Paragraph('Bảng tóm tắt dựa trên lượt nộp thực tế. Điểm chi tiết từng nhiệm vụ có trong tệp Excel.',normal)]
        doc.build(story);buf.seek(0)
        return StreamingResponse(buf,media_type='application/pdf',headers={'Content-Disposition':f'attachment; filename="toan8_lop_{cid}.pdf"'})

    # Trò chơi có ngân hàng đáp án cố định ở phía máy chủ; không tính vào điểm kiểm tra.
    GAME_ITEMS=[
      {'prompt':'A² − B² bằng biểu thức nào?','options':['(A−B)(A+B)','(A+B)²','A²+B²'],'correct':0},
      {'prompt':'(A + B)² bằng biểu thức nào?','options':['A² + B²','A²+2AB+B²','A²−2AB+B²'],'correct':1},
      {'prompt':'(A − B)² bằng biểu thức nào?','options':['A²−2AB+B²','A²−B²','A²+2AB+B²'],'correct':0},
      {'prompt':'A³ + B³ bằng biểu thức nào?','options':['(A−B)(A²+AB+B²)','(A+B)(A²−AB+B²)','(A+B)³'],'correct':1},
      {'prompt':'A³ − B³ bằng biểu thức nào?','options':['(A−B)(A²+AB+B²)','(A+B)(A²−AB+B²)','(A−B)³'],'correct':0},
    ]
    @app.get('/api/games/formulas')
    def game(request:Request):
        auth(request,['student'])
        return {'game':'Ghép công thức – Chương II','items':[{'prompt':x['prompt'],'options':x['options']} for x in GAME_ITEMS]}

    @app.post('/api/games/formulas')
    async def game_submit(request:Request):
        u=auth(request,['student'],True);d=await request.json();ans=d.get('choices')
        if not isinstance(ans,list) or len(ans)!=len(GAME_ITEMS):fail(400,"Cần trả lời đủ số câu của trò chơi.")
        if any(not isinstance(v,int) or isinstance(v,bool) or v<0 or v>2 for v in ans):fail(400,"Đáp án trò chơi không hợp lệ.")
        score=sum(int(x==item['correct']) for x,item in zip(ans,GAME_ITEMS))
        with db() as con:
            con.execute('INSERT INTO games(student_id,game,correct,total,created_at) VALUES (?,?,?,?,?)',(u['id'],'formulas',score,len(GAME_ITEMS),now()))
        return {'correct':score,'total':len(GAME_ITEMS),'note':'Điểm trò chơi không tính vào điểm bài tập/kiểm tra.'}

    @app.post('/api/ai/help')
    async def ai_help(request:Request):
        u=auth(request,['student'],True)
        if not os.environ.get('OPENAI_API_KEY'):fail(503,"Chưa kết nối API gia sư AI. Giáo viên có thể tiếp tục dạy học bình thường.")
        d=await request.json();lesson=integer(d.get('lesson_id'),'Bài học',1,39)
        message=textval(d.get('message',''),1200)
        aid=d.get('assignment_id')
        with db() as con:
            no_ai_active=con.execute('''SELECT 1 FROM attempts t JOIN assignments a ON a.id=t.assignment_id
               WHERE t.student_id=? AND t.status='in_progress' AND a.allow_hints=0 LIMIT 1''',(u['id'],)).fetchone()
            if no_ai_active:fail(403,'Gia sư AI tạm khóa vì em đang làm bài có quy định không trợ giúp.')
            if aid is not None:
                aid=integer(aid,'Nhiệm vụ')
                a=con.execute('''SELECT a.allow_hints,t.status FROM assignments a JOIN assignment_recipients r ON r.assignment_id=a.id
                LEFT JOIN attempts t ON t.assignment_id=a.id AND t.student_id=r.student_id WHERE a.id=? AND r.student_id=?''',(aid,u['id'])).fetchone()
                if not a:fail(403,"Không có quyền nhận hỗ trợ cho nhiệm vụ này.")
                if not a['allow_hints'] and (not a['status'] or a['status']=='in_progress'):fail(403,"Giáo viên đã tắt trợ giúp AI cho bài đang làm.")
        note=VERIFIED_NOTES.get(lesson)
        if note is None:fail(409,"Bài này chưa có tóm tắt được kiểm tra trong hệ thống; AI không tự tạo kiến thức thay SGK.")
        import httpx
        system=('Bạn là MATH AI cho học sinh lớp 8. Chỉ giải thích trong phạm vi kiến thức: '+note['focus']+'; công thức: '+', '.join(note['formulas'])+'. '
          'Ưu tiên gợi mở trước đáp án. Không suy đoán dữ liệu đề thiếu, không yêu cầu thông tin cá nhân. '
          'Nếu thiếu bối cảnh hãy yêu cầu học sinh giải thích thêm; không bịa công thức. '
          'Không làm thay bài kiểm tra; nhắc tư duy từng bước.')
        try:
            async with httpx.AsyncClient(timeout=25.0) as client:
                res=await client.post('https://api.openai.com/v1/responses',headers={'Authorization':'Bearer '+os.environ['OPENAI_API_KEY']},json={
                    'model':os.environ.get('OPENAI_MODEL','gpt-4.1-mini'),
                    'instructions':system,'input':message,'max_output_tokens':400})
                res.raise_for_status()
                body=res.json()
        except Exception:fail(502,"Không thể kết nối dịch vụ AI lúc này. Vui lòng thử lại.")
        answer='\n'.join(x.get('text','') for out in body.get('output',[]) for x in out.get('content',[]) if x.get('type')=='output_text')
        if not answer:fail(502,"AI chưa trả về nội dung hướng dẫn.")
        return {'reply':answer}

    # Frontend + SPA. Route không thuộc /api chuyển về màn hình ứng dụng.
    app.mount('/assets',StaticFiles(directory=HERE/'web'),name='assets')
    @app.get('/')
    def index():return FileResponse(HERE/'web'/'index.html')
    @app.get('/setup')
    def setup_page():return FileResponse(HERE/'web'/'setup.html')
    return app

app=create_app()


def cli():
    parser=argparse.ArgumentParser(description='Quản trị cục bộ Toán 8 Master')
    sub=parser.add_subparsers(dest='command',required=True)
    sub.add_parser('init')
    admin=sub.add_parser('create-admin');admin.add_argument('--username',required=True);admin.add_argument('--name',default='Quản trị hệ thống')
    args=parser.parse_args()
    if args.command=='init':print('CSDL đã khởi tạo:',DEFAULT_DB);return
    import getpass
    pw=getpass.getpass('Nhập mật khẩu quản trị (tối thiểu 12 ký tự): ')
    check=getpass.getpass('Nhập lại mật khẩu: ')
    if len(pw)<12 or pw!=check:raise SystemExit('Mật khẩu không khớp hoặc quá ngắn.')
    if not re.fullmatch('[a-z0-9_.-]{3,60}',args.username):raise SystemExit('Tên đăng nhập không hợp lệ.')
    dbfile=Path(DEFAULT_DB);con=sqlite3.connect(dbfile)
    try:
        con.execute("INSERT INTO users(username,full_name,role,password_hash,created_at) VALUES (?,?,'admin',?,?)",(args.username,args.name,PH.hash(pw),now()))
        con.commit();print('Đã tạo tài khoản quản trị:',args.username)
    except sqlite3.IntegrityError:raise SystemExit('Tên đăng nhập đã tồn tại.')
    finally:con.close()

if __name__=='__main__':cli()
