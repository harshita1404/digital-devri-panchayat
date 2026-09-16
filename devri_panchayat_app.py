from flask import Flask, request, redirect, url_for, session, render_template_string, send_file, flash
import sqlite3
from pathlib import Path
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO
import secrets
import os
import json
import urllib.request
import urllib.error

# ============================================================
# DIGITAL DEVRI PANCHAYAT
# Single-file Flask + SQLite + Excel System
# ============================================================

app = Flask(__name__)
app.secret_key = "DEVRI_PANCHAYAT_SECRET_KEY_2026_CHANGE_THIS"

BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "devri_panchayat.db"
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "admin123"


# ------------------------------------------------------------
# 5 VILLAGE RECORDS
# ------------------------------------------------------------

VILLAGES = [
    {"id": 1, "name_hi": "देवरी", "name_en": "Devri", "description": "", "population": "", "facilities": ""},
    {"id": 2, "name_hi": "पारलिया", "name_en": "Parliya", "description": "", "population": "", "facilities": ""},
    {"id": 3, "name_hi": "हापावास", "name_en": "Hapawas", "description": "", "population": "", "facilities": ""},
    {"id": 4, "name_hi": "घाघसा", "name_en": "Ghagsa", "description": "", "population": "", "facilities": ""},
    {"id": 5, "name_hi": "ठुकरावा", "name_en": "Thukrawa", "description": "", "population": "", "facilities": ""},
]

TOTAL_VILLAGES = len(VILLAGES)


# ------------------------------------------------------------
# DATABASE
# ------------------------------------------------------------

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def current_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db():

    conn = get_db()
    cur = conn.cursor()

    # Panchayat
    cur.execute("""
        CREATE TABLE IF NOT EXISTS panchayat (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name_hi TEXT,
            name_en TEXT,
            description_hi TEXT,
            contact TEXT,
            email TEXT
        )
    """)

    # Villages
    cur.execute("""
        CREATE TABLE IF NOT EXISTS villages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name_hi TEXT UNIQUE,
            name_en TEXT,
            description TEXT,
            population TEXT,
            facilities TEXT
        )
    """)

    # USERS
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            mobile TEXT NOT NULL UNIQUE,
            village TEXT NOT NULL,
            registered_at TEXT NOT NULL,
            last_seen TEXT,
            consent INTEGER DEFAULT 1
        )
    """)

    # Existing databases: safely add citizen verification fields.
    existing_columns = {row["name"] for row in cur.execute("PRAGMA table_info(users)").fetchall()}
    new_user_columns = {
        "aadhaar": "TEXT",
        "mobile_verified": "INTEGER DEFAULT 0",
        "verified_at": "TEXT",
        "last_login": "TEXT",
        "login_count": "INTEGER DEFAULT 0",
    }
    for column_name, column_type in new_user_columns.items():
        if column_name not in existing_columns:
            cur.execute(f"ALTER TABLE users ADD COLUMN {column_name} {column_type}")

    # Remove old village records from the previous version and keep exactly five.
    current_village_names = [v["name_hi"] for v in VILLAGES]
    placeholders = ",".join("?" for _ in current_village_names)
    cur.execute(
        f"DELETE FROM villages WHERE name_hi NOT IN ({placeholders})",
        current_village_names
    )

    # Do not expose old test/project/medical records that belong to the previous village list.
    cur.execute(
        f"DELETE FROM projects WHERE village NOT IN ({placeholders})",
        current_village_names
    )
    cur.execute(
        f"DELETE FROM medical WHERE village NOT IN ({placeholders})",
        current_village_names
    )

    # Complaints
    cur.execute("""
        CREATE TABLE IF NOT EXISTS complaints (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            complaint_id TEXT UNIQUE NOT NULL,
            name TEXT,
            mobile TEXT,
            village TEXT NOT NULL,
            ward TEXT,
            category TEXT NOT NULL,
            description TEXT NOT NULL,
            photo TEXT,
            location TEXT,
            status TEXT DEFAULT 'दर्ज',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)

    # Complaint updates
    cur.execute("""
        CREATE TABLE IF NOT EXISTS complaint_updates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            complaint_id TEXT NOT NULL,
            status TEXT NOT NULL,
            message TEXT,
            created_at TEXT NOT NULL
        )
    """)

    # Projects
    cur.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village TEXT NOT NULL,
            title TEXT NOT NULL,
            description TEXT,
            department TEXT,
            budget TEXT,
            progress INTEGER DEFAULT 0,
            status TEXT DEFAULT 'प्रस्तावित',
            start_date TEXT,
            expected_date TEXT,
            created_at TEXT NOT NULL
        )
    """)

    # Budget
    cur.execute("""
        CREATE TABLE IF NOT EXISTS budget (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            financial_year TEXT NOT NULL,
            category TEXT NOT NULL,
            amount TEXT NOT NULL,
            description TEXT,
            created_at TEXT NOT NULL
        )
    """)

    # Medical
    cur.execute("""
        CREATE TABLE IF NOT EXISTS medical (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village TEXT NOT NULL,
            facility_name TEXT NOT NULL,
            type TEXT,
            contact TEXT,
            timing TEXT,
            description TEXT,
            created_at TEXT NOT NULL
        )
    """)

    # Notices
    cur.execute("""
        CREATE TABLE IF NOT EXISTS notices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            content TEXT NOT NULL,
            notice_date TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)

    # Admin
    cur.execute("""
        CREATE TABLE IF NOT EXISTS admins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL
        )
    """)

    # Default Panchayat
    cur.execute("SELECT COUNT(*) AS c FROM panchayat")
    if cur.fetchone()["c"] == 0:
        cur.execute("""
            INSERT INTO panchayat
            (name_hi, name_en, description_hi, contact, email)
            VALUES (?, ?, ?, ?, ?)
        """, (
            "देवरी ग्राम पंचायत",
            "Devri Gram Panchayat",
            "देवरी ग्राम पंचायत एक डिजिटल पंचायत पहल है जिसका उद्देश्य ग्रामीण नागरिकों को पंचायत की जानकारी, शिकायत दर्ज करने, विकास कार्यों और सूचनाओं तक आसान पहुंच देना है।",
            "जानकारी अपडेट की जानी है",
            "जानकारी अपडेट की जानी है"
        ))

    # Keep the public directory synchronized with the five current villages.
    for v in VILLAGES:
        cur.execute("""
            UPDATE villages
            SET name_en=?, description=?, population=?, facilities=?
            WHERE name_hi=?
        """, (
            v["name_en"], v["description"], v["population"],
            v["facilities"], v["name_hi"]
        ))
        if cur.rowcount == 0:
            cur.execute("""
                INSERT INTO villages
                (name_hi, name_en, description, population, facilities)
                VALUES (?, ?, ?, ?, ?)
            """, (
                v["name_hi"], v["name_en"], v["description"],
                v["population"], v["facilities"]
            ))

    # Admin
    cur.execute("SELECT COUNT(*) AS c FROM admins")
    if cur.fetchone()["c"] == 0:
        cur.execute("""
            INSERT INTO admins(username, password_hash)
            VALUES (?, ?)
        """, (
            ADMIN_USERNAME,
            generate_password_hash(ADMIN_PASSWORD)
        ))

    # Demo projects only if empty
    cur.execute("SELECT COUNT(*) AS c FROM projects")
    if cur.fetchone()["c"] == 0:
        cur.execute("""
            INSERT INTO projects
            (village, title, description, department, budget,
             progress, status, start_date, expected_date, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            "देवरी",
            "ग्राम सड़क विकास कार्य",
            "सड़क सुधार एवं मरम्मत कार्य।",
            "ग्राम पंचायत",
            "जानकारी अपडेट की जानी है",
            25,
            "कार्य जारी",
            "",
            "",
            current_time()
        ))

    # Demo medical facility
    cur.execute("SELECT COUNT(*) AS c FROM medical")
    if cur.fetchone()["c"] == 0:
        cur.execute("""
            INSERT INTO medical
            (village, facility_name, type, contact, timing, description, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            "देवरी",
            "स्थानीय स्वास्थ्य सुविधा",
            "स्वास्थ्य केंद्र",
            "जानकारी अपडेट की जानी है",
            "जानकारी अपडेट की जानी है",
            "स्वास्थ्य सुविधा की विस्तृत जानकारी जल्द अपडेट की जाएगी।",
            current_time()
        ))

    # Demo notice
    cur.execute("SELECT COUNT(*) AS c FROM notices")
    if cur.fetchone()["c"] == 0:
        cur.execute("""
            INSERT INTO notices(title, content, notice_date, created_at)
            VALUES (?, ?, ?, ?)
        """, (
            "देवरी ग्राम पंचायत डिजिटल पोर्टल",
            "पंचायत से संबंधित जानकारी और शिकायतों के लिए इस पोर्टल का उपयोग करें।",
            datetime.now().strftime("%Y-%m-%d"),
            current_time()
        ))

    conn.commit()
    conn.close()


init_db()

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------

def is_admin():
    return session.get("admin_logged_in") is True


def generate_user_id():
    while True:
        uid = "DVP-U-" + secrets.token_hex(4).upper()
        conn = get_db()
        row = conn.execute(
            "SELECT id FROM users WHERE user_id = ?", (uid,)
        ).fetchone()
        conn.close()

        if not row:
            return uid


def generate_complaint_id():
    while True:
        cid = "DVP-" + str(datetime.now().year) + "-" + secrets.token_hex(3).upper()
        conn = get_db()
        row = conn.execute(
            "SELECT id FROM complaints WHERE complaint_id = ?", (cid,)
        ).fetchone()
        conn.close()

        if not row:
            return cid


def normalize_aadhaar(value):
    return "".join(ch for ch in (value or "") if ch.isdigit())


def mask_aadhaar(value):
    digits = normalize_aadhaar(value)
    return "XXXX-XXXX-" + digits[-4:] if len(digits) == 12 else "—"


def render_page(title, content, active="home", **kwargs):

    village_rows = []

    for v in VILLAGES:
        village_rows.append({
            "name_hi": v["name_hi"],
            "name_en": v["name_en"]
        })

    return render_template_string(
        BASE_TEMPLATE,
        title=title,
        content=content,
        active=active,
        villages=village_rows,
        admin=is_admin(),
        **kwargs
    )


# ------------------------------------------------------------
# BASE HTML
# ------------------------------------------------------------

BASE_TEMPLATE = """
<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>{{ title }} | Digital Devri Panchayat</title>

<style>

*{
    box-sizing:border-box;
}

body{
    margin:0;
    font-family:Arial, "Noto Sans Devanagari", sans-serif;
    background:#f4f7f3;
    color:#1f2937;
}

a{
    text-decoration:none;
    color:inherit;
}

.navbar{
    background:#14532d;
    color:white;
    padding:14px 5%;
    display:flex;
    justify-content:space-between;
    align-items:center;
    position:sticky;
    top:0;
    z-index:1000;
}

.logo{
    font-size:22px;
    font-weight:bold;
}

.navlinks{
    display:flex;
    gap:12px;
    flex-wrap:wrap;
}

.navlinks a{
    padding:8px 12px;
    border-radius:8px;
}

.navlinks a:hover{
    background:#166534;
}

.container{
    width:90%;
    max-width:1200px;
    margin:30px auto;
}

.hero{
    background:linear-gradient(135deg,#166534,#22c55e);
    color:white;
    border-radius:22px;
    padding:45px 30px;
    margin-bottom:30px;
}

.hero h1{
    font-size:38px;
    margin:0 0 12px;
}

.hero p{
    font-size:18px;
    line-height:1.7;
}

.btn{
    display:inline-block;
    background:#166534;
    color:white;
    padding:11px 17px;
    border:none;
    border-radius:9px;
    cursor:pointer;
    font-weight:bold;
}

.btn:hover{
    background:#14532d;
}

.btn-light{
    background:white;
    color:#166534;
}

.btn-danger{
    background:#b91c1c;
}

.btn-warning{
    background:#ca8a04;
}

.card-grid{
    display:grid;
    grid-template-columns:repeat(auto-fit,minmax(240px,1fr));
    gap:18px;
}

.card{
    background:white;
    border-radius:16px;
    padding:22px;
    box-shadow:0 3px 15px rgba(0,0,0,.08);
    margin-bottom:20px;
}

.card h3{
    margin-top:0;
    color:#14532d;
}

.icon{
    font-size:35px;
}

.form-card{
    background:white;
    padding:25px;
    border-radius:18px;
    box-shadow:0 3px 15px rgba(0,0,0,.08);
}

.form-group{
    margin-bottom:16px;
}

label{
    display:block;
    margin-bottom:7px;
    font-weight:bold;
}

input, select, textarea{
    width:100%;
    padding:12px;
    border:1px solid #d1d5db;
    border-radius:8px;
    font-size:15px;
}

textarea{
    min-height:120px;
    resize:vertical;
}

.stats{
    display:grid;
    grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
    gap:15px;
    margin-bottom:25px;
}

.stat{
    background:white;
    padding:20px;
    border-radius:15px;
    text-align:center;
    box-shadow:0 3px 12px rgba(0,0,0,.07);
}

.stat-number{
    font-size:32px;
    font-weight:bold;
    color:#166534;
}

table{
    width:100%;
    border-collapse:collapse;
    background:white;
}

th,td{
    padding:12px;
    border-bottom:1px solid #e5e7eb;
    text-align:left;
    vertical-align:top;
}

th{
    background:#dcfce7;
    color:#14532d;
}

.table-wrap{
    overflow-x:auto;
    border-radius:12px;
}

.badge{
    display:inline-block;
    padding:6px 10px;
    border-radius:20px;
    background:#dcfce7;
    color:#166534;
    font-size:13px;
    font-weight:bold;
}

.alert{
    padding:14px;
    margin:15px 0;
    border-radius:10px;
    background:#dcfce7;
    color:#14532d;
}

.alert-error{
    background:#fee2e2;
    color:#991b1b;
}

.timeline{
    border-left:3px solid #22c55e;
    padding-left:20px;
}

.timeline-item{
    margin-bottom:20px;
    position:relative;
}

.timeline-item::before{
    content:"";
    width:11px;
    height:11px;
    background:#166534;
    border-radius:50%;
    position:absolute;
    left:-27px;
    top:4px;
}

.footer{
    margin-top:50px;
    background:#14532d;
    color:white;
    padding:30px 5%;
    text-align:center;
}

.mobile-nav{
    display:none;
}

@media(max-width:800px){

    .navbar{
        flex-direction:column;
        gap:12px;
    }

    .navlinks{
        justify-content:center;
    }

    .hero h1{
        font-size:28px;
    }

    .mobile-nav{
        display:flex;
        position:fixed;
        bottom:0;
        left:0;
        right:0;
        background:white;
        border-top:1px solid #ddd;
        z-index:999;
        justify-content:space-around;
        padding:8px 3px;
    }

    .mobile-nav a{
        font-size:12px;
        text-align:center;
    }

    body{
        padding-bottom:65px;
    }
}

</style>
</head>

<body>

<div class="navbar">

    <div class="logo">
        🌿 Digital Devri Panchayat
    </div>

    <div class="navlinks">

        <a href="{{ url_for('home') }}">🏠 होम</a>

        <a href="{{ url_for('villages') }}">🏘️ गांव</a>

        <a href="{{ url_for('complaint') }}">📝 शिकायत</a>

        <a href="{{ url_for('track_complaint') }}">🔎 शिकायत ट्रैक</a>

        <a href="{{ url_for('projects') }}">🏗️ विकास</a>

        <a href="{{ url_for('budget') }}">💰 बजट</a>

        <a href="{{ url_for('medical') }}">🏥 स्वास्थ्य</a>

        <a href="{{ url_for('notices') }}">📢 सूचनाएं</a>
        <a href="{{ url_for('login') }}">🔐 नागरिक लॉगिन</a>

        {% if admin %}
            <a href="{{ url_for('admin_dashboard') }}">⚙️ Admin</a>
            <a href="{{ url_for('admin_logout') }}">Logout</a>
        {% else %}
            <a href="{{ url_for('admin_login') }}">🔐 Admin Login</a>
        {% endif %}

    </div>

</div>


<div class="container">

    {% with messages = get_flashed_messages(with_categories=true) %}

        {% for category, message in messages %}

            <div class="alert {% if category == 'error' %}alert-error{% endif %}">
                {{ message }}
            </div>

        {% endfor %}

    {% endwith %}

    {{ content|safe }}

</div>


<div class="mobile-nav">

    <a href="{{ url_for('home') }}">🏠<br>होम</a>

    <a href="{{ url_for('villages') }}">🏘️<br>गांव</a>

    <a href="{{ url_for('complaint') }}">📝<br>शिकायत</a>

    <a href="{{ url_for('track_complaint') }}">🔎<br>ट्रैक</a>

    <a href="{{ url_for('login') }}">🔐<br>लॉगिन</a>

</div>


<div class="footer">

    <h3>🌿 Digital Devri Panchayat</h3>

    <p>
        ग्रामीण नागरिकों के लिए सरल और डिजिटल पंचायत सेवा पोर्टल
    </p>

    <p>
        © 2026 Devri Gram Panchayat
    </p>

</div>

</body>
</html>
"""


# ============================================================
# HOME
# ============================================================

@app.route("/")
def home():
    conn = get_db()
    complaint_count = conn.execute("SELECT COUNT(*) AS c FROM complaints").fetchone()["c"]
    project_count = conn.execute("SELECT COUNT(*) AS c FROM projects").fetchone()["c"]
    notice_count = conn.execute("SELECT COUNT(*) AS c FROM notices").fetchone()["c"]
    notice_rows = conn.execute("SELECT * FROM notices ORDER BY id DESC LIMIT 3").fetchall()
    conn.close()

    schemes = [
        ("🌾", "प्रधानमंत्री किसान सम्मान निधि"),
        ("🛠️", "मनरेगा"),
        ("🏥", "मुख्यमंत्री आयुष्मान आरोग्य योजना"),
        ("💊", "मुख्यमंत्री निःशुल्क दवा एवं जांच योजना"),
        ("👵", "सामाजिक सुरक्षा पेंशन"),
        ("🍚", "NFSA / राशन संबंधी योजनाएं"),
    ]
    scheme_cards = "".join(
        f'''<div class="card"><div class="icon">{icon}</div><h3>{name}</h3><p>पात्रता और वर्तमान जानकारी आधिकारिक पोर्टल पर जांचें।</p></div>'''
        for icon, name in schemes
    )

    notice_html = "".join(
        f'''<div class="card"><h3>{n["title"]}</h3><p>{n["content"]}</p><small>{n["notice_date"]}</small></div>'''
        for n in notice_rows
    ) or '<div class="card">अभी कोई सूचना उपलब्ध नहीं है।</div>'

    content = f'''
    <div class="hero">
        <h1>🌿 डिजिटल देवरी ग्राम पंचायत</h1>
        <p>पंचायत की जानकारी, 5 गांवों की जानकारी, शिकायत, विकास कार्य और पंचायत सूचनाएं — सब एक ही स्थान पर।</p>
        <a class="btn btn-light" href="/register">👤 नागरिक पंजीकरण</a>
        <a class="btn btn-light" href="/login">🔐 नागरिक लॉगिन</a>
        <a class="btn btn-light" href="/complaint">📝 शिकायत दर्ज करें</a>
    </div>

    <div class="stats">
        <div class="stat"><div class="stat-number">5</div><div>कुल गांव</div></div>
        <div class="stat"><div class="stat-number">{complaint_count}</div><div>शिकायतें</div></div>
        <div class="stat"><div class="stat-number">{project_count}</div><div>विकास कार्य</div></div>
        <div class="stat"><div class="stat-number">{notice_count}</div><div>सूचनाएं</div></div>
    </div>

    <h2>मुख्य सेवाएं</h2>
    <div class="card-grid">
        <div class="card"><div class="icon">🏘️</div><h3>हमारे गांव</h3><p>देवरी पंचायत के 5 गांव देखें।</p><a class="btn" href="/villages">देखें</a></div>
        <div class="card"><div class="icon">📝</div><h3>शिकायत दर्ज करें</h3><p>पानी, सड़क, बिजली आदि से जुड़ी समस्या दर्ज करें।</p><a class="btn" href="/complaint">शिकायत करें</a></div>
        <div class="card"><div class="icon">🔎</div><h3>शिकायत ट्रैक करें</h3><p>अपने Complaint ID से स्थिति देखें।</p><a class="btn" href="/track">ट्रैक करें</a></div>
        <div class="card"><div class="icon">🏗️</div><h3>विकास कार्य</h3><p>प्रकाशित विकास कार्यों की जानकारी देखें।</p><a class="btn" href="/projects">देखें</a></div>
        <div class="card"><div class="icon">💰</div><h3>बजट जानकारी</h3><p>पंचायत द्वारा प्रकाशित बजट एवं व्यय जानकारी।</p><a class="btn" href="/budget">देखें</a></div>
        <div class="card"><div class="icon">🏥</div><h3>स्वास्थ्य सुविधाएं</h3><p>पंचायत द्वारा प्रकाशित स्वास्थ्य जानकारी देखें।</p><a class="btn" href="/medical">देखें</a></div>
    </div>

    <br><h2>🏛️ सरकारी योजनाएं</h2>
    <p>प्रमुख योजनाओं के नाम नीचे दिए गए हैं। पात्रता और वर्तमान लाभ की स्थिति आधिकारिक पोर्टल पर सत्यापित करें।</p>
    <div class="card-grid">{scheme_cards}</div>
    <div class="card">
        <h3>🔗 आधिकारिक योजना जानकारी</h3>
        <p>राजस्थान की योजनाओं की पंचायत/वार्ड स्तर की जानकारी Jan Soochna Portal पर देखें।</p>
        <a class="btn" href="https://jansoochna.rajasthan.gov.in/Scheme" target="_blank" rel="noopener">Jan Soochna Portal खोलें</a>
    </div>

    <h2>📢 नवीनतम सूचनाएं</h2>
    {notice_html}
    '''
    return render_page("होम", content)


# ============================================================
# USER REGISTRATION

# ============================================================

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        mobile = request.form.get("mobile", "").strip()
        village = request.form.get("village", "").strip()
        aadhaar = normalize_aadhaar(request.form.get("aadhaar", ""))
        consent = request.form.get("consent")

        if not name or not mobile or not village or not aadhaar or not consent:
            flash("कृपया सभी जरूरी जानकारी भरें।", "error")
            return redirect(url_for("register"))
        if not mobile.isdigit() or len(mobile) != 10:
            flash("कृपया 10 अंकों का मोबाइल नंबर डालें।", "error")
            return redirect(url_for("register"))
        if len(aadhaar) != 12:
            flash("कृपया 12 अंकों का Aadhaar नंबर डालें।", "error")
            return redirect(url_for("register"))
        if village not in [v["name_hi"] for v in VILLAGES]:
            flash("कृपया सही गांव चुनें।", "error")
            return redirect(url_for("register"))

        now = current_time()
        conn = get_db()
        existing = conn.execute("SELECT * FROM users WHERE mobile=?", (mobile,)).fetchone()
        if existing:
            user_id = existing["user_id"]
            conn.execute(
                """UPDATE users SET name=?, village=?, aadhaar=?, mobile_verified=1,
                   verified_at=?, last_seen=?, consent=1 WHERE mobile=?""",
                (name, village, aadhaar, now, now, mobile)
            )
            flash(f"पंजीकरण अपडेट और लॉगिन सफल! आपका User ID: {user_id}", "success")
        else:
            user_id = generate_user_id()
            conn.execute(
                """INSERT INTO users
                   (user_id,name,mobile,village,aadhaar,registered_at,last_seen,consent,
                    mobile_verified,verified_at,last_login,login_count)
                   VALUES (?,?,?,?,?,?,?,1,1,?,?,1)""",
                (user_id, name, mobile, village, aadhaar, now, now, now, now)
            )
            flash(f"पंजीकरण सफल! आपका User ID: {user_id}", "success")
        conn.commit()
        conn.close()
        session["citizen_user_id"] = user_id
        return redirect(url_for("home"))

    village_options = "".join(f'<option value="{v["name_hi"]}">{v["name_hi"]} ({v["name_en"]})</option>' for v in VILLAGES)
    content = f"""
    <div class="form-card">
        <h1>👤 नागरिक पंजीकरण</h1>
        <p>अपना नाम, मोबाइल, Aadhaar और गांव भरें। जानकारी पंचायत के सुरक्षित रिकॉर्ड में सेव होगी।</p>
        <form method="POST">
            <div class="form-group"><label>नाम *</label><input name="name" required></div>
            <div class="form-group"><label>मोबाइल नंबर *</label><input name="mobile" inputmode="numeric" maxlength="10" pattern="[0-9]{{10}}" required></div>
            <div class="form-group"><label>Aadhaar नंबर *</label><input name="aadhaar" inputmode="numeric" maxlength="12" pattern="[0-9]{{12}}" required></div>
            <div class="form-group"><label>गांव *</label><select name="village" required><option value="">गांव चुनें</option>{village_options}</select></div>
            <div class="form-group"><label><input type="checkbox" name="consent" value="1" required> मैं अपनी जानकारी पंचायत के सुरक्षित रिकॉर्ड में सेव करने की सहमति देता/देती हूं।</label></div>
            <button class="btn">✅ Register</button>
        </form>
    </div>"""
    return render_page("नागरिक पंजीकरण", content)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        mobile = request.form.get("mobile", "").strip()
        village = request.form.get("village", "").strip()
        if not mobile.isdigit() or len(mobile) != 10:
            flash("कृपया 10 अंकों का मोबाइल नंबर डालें।", "error")
            return redirect(url_for("login"))
        if village not in [v["name_hi"] for v in VILLAGES]:
            flash("कृपया सही गांव चुनें।", "error")
            return redirect(url_for("login"))
        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE mobile=? AND village=?", (mobile, village)).fetchone()
        if not user:
            conn.close()
            flash("Mobile number और गांव का रिकॉर्ड नहीं मिला। पहले सही जानकारी से registration करें।", "error")
            return redirect(url_for("login"))
        now = current_time()
        conn.execute(
            "UPDATE users SET last_seen=?, last_login=?, login_count=COALESCE(login_count,0)+1, mobile_verified=1 WHERE mobile=? AND village=?",
            (now, now, mobile, village)
        )
        conn.commit()
        conn.close()
        session["citizen_user_id"] = user["user_id"]
        flash("नागरिक लॉगिन सफल।", "success")
        return redirect(url_for("home"))
    village_options = "".join(f'<option value="{v["name_hi"]}">{v["name_hi"]} ({v["name_en"]})</option>' for v in VILLAGES)
    content = f"""
    <div class="form-card"><h1>🔐 नागरिक लॉगिन</h1>
    <p>Login के लिए अपना registered mobile number और गांव चुनें।</p>
    <form method="POST">
        <div class="form-group"><label>Registered Mobile Number *</label><input type="tel" name="mobile" inputmode="numeric" maxlength="10" minlength="10" pattern="[0-9]{{10}}" oninput="this.value=this.value.replace(/[^0-9]/g,'').slice(0,10)" required></div>
        <div class="form-group"><label>गांव *</label><select name="village" required><option value="">गांव चुनें</option>{village_options}</select></div>
        <button class="btn">🔐 Login</button>
    </form>
    <br><a class="btn btn-light" href="/register">नया नागरिक पंजीकरण</a></div>"""
    return render_page("नागरिक लॉगिन", content)


# ============================================================
# VILLAGES

# ============================================================

@app.route("/villages")
def villages():
    conn = get_db(); rows = conn.execute("SELECT * FROM villages ORDER BY id").fetchall(); conn.close()
    cards = "".join(f'''<div class="card"><div class="icon">🏘️</div><h3>{v['name_hi']}</h3><p><b>English:</b> {v['name_en']}</p><a class="btn" href="/village/{v['id']}">गांव देखें</a></div>''' for v in rows)
    content = f'''<h1>🏘️ हमारे गांव</h1><p>देवरी ग्राम पंचायत के अंतर्गत आने वाले 5 गांव।</p><div class="card-grid">{cards}</div>'''
    return render_page("हमारे गांव", content, active="villages")


@app.route("/village/<int:village_id>")
def village_detail(village_id):
    conn = get_db()
    village = conn.execute("SELECT * FROM villages WHERE id=?", (village_id,)).fetchone()
    if not village:
        conn.close(); return "Village not found", 404
    projects = conn.execute("SELECT * FROM projects WHERE village=? ORDER BY id DESC", (village["name_hi"],)).fetchall()
    medical = conn.execute("SELECT * FROM medical WHERE village=? ORDER BY id DESC", (village["name_hi"],)).fetchall()
    conn.close()
    project_html = "".join(f'''<div class="card"><h3>🏗️ {p['title']}</h3><p>{p['description'] or ""}</p><p><b>विभाग:</b> {p['department'] or "—"}</p><p><b>बजट:</b> {p['budget'] or "—"}</p><p><b>प्रगति:</b> {p['progress']}%</p><span class="badge">{p['status']}</span></div>''' for p in projects) or '<div class="card">अभी कोई विकास कार्य दर्ज नहीं है।</div>'
    medical_html = "".join(f'''<div class="card"><h3>🏥 {m['facility_name']}</h3><p><b>प्रकार:</b> {m['type'] or "—"}</p><p><b>संपर्क:</b> {m['contact'] or "—"}</p><p><b>समय:</b> {m['timing'] or "—"}</p><p>{m['description'] or ""}</p></div>''' for m in medical) or '<div class="card">अभी कोई स्वास्थ्य सुविधा दर्ज नहीं है।</div>'
    content = f'''<div class="hero"><h1>🏘️ {village["name_hi"]}</h1><p>इस गांव की आधिकारिक जानकारी पंचायत द्वारा अपडेट की जाएगी।</p></div><div class="card"><h2>गांव की जानकारी</h2><p><b>नाम:</b> {village["name_hi"]}</p><p><b>English Name:</b> {village["name_en"]}</p></div><h2>🏗️ विकास कार्य</h2>{project_html}<h2>🏥 स्वास्थ्य सुविधाएं</h2>{medical_html}'''
    return render_page(village["name_hi"], content)


# ============================================================
# COMPLAINT SUBMIT

# ============================================================

@app.route("/complaint", methods=["GET", "POST"])
def complaint():

    if request.method == "POST":

        name = request.form.get("name", "").strip()
        mobile = request.form.get("mobile", "").strip()
        village = request.form.get("village", "").strip()
        ward = request.form.get("ward", "").strip()
        category = request.form.get("category", "").strip()
        description = request.form.get("description", "").strip()
        location = request.form.get("location", "").strip()

        if not village or not category or not description:
            flash(
                "गांव, समस्या की श्रेणी और समस्या का विवरण जरूरी है।",
                "error"
            )
            return redirect(url_for("complaint"))

        if mobile and (not mobile.isdigit() or len(mobile) != 10):
            flash("मोबाइल नंबर 10 अंकों का होना चाहिए।", "error")
            return redirect(url_for("complaint"))

        complaint_id = generate_complaint_id()
        now = current_time()

        photo_filename = ""

        uploaded_file = request.files.get("photo")

        if uploaded_file and uploaded_file.filename:

            safe_name = (
                complaint_id
                + "_"
                + os.path.basename(uploaded_file.filename)
            )

            photo_filename = safe_name

            uploaded_file.save(
                UPLOAD_DIR / safe_name
            )

        conn = get_db()

        conn.execute(
            """
            INSERT INTO complaints
            (
                complaint_id,
                name,
                mobile,
                village,
                ward,
                category,
                description,
                photo,
                location,
                status,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                complaint_id,
                name,
                mobile,
                village,
                ward,
                category,
                description,
                photo_filename,
                location,
                "दर्ज",
                now,
                now
            )
        )

        conn.execute(
            """
            INSERT INTO complaint_updates
            (complaint_id, status, message, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                complaint_id,
                "दर्ज",
                "शिकायत सफलतापूर्वक दर्ज की गई।",
                now
            )
        )

        conn.commit()
        conn.close()

        flash(
            f"शिकायत दर्ज हो गई। आपका Complaint ID है: {complaint_id}",
            "success"
        )

        return redirect(
            url_for(
                "track_complaint",
                complaint_id=complaint_id
            )
        )

    village_options = "".join(
        f'<option value="{v["name_hi"]}">{v["name_hi"]}</option>'
        for v in VILLAGES
    )

    content = f"""

    <div class="form-card">

        <h1>📝 शिकायत दर्ज करें</h1>

        <form method="POST" enctype="multipart/form-data">

            <div class="form-group">
                <label>नाम</label>
                <input type="text" name="name">
            </div>

            <div class="form-group">
                <label>मोबाइल नंबर</label>
                <input type="tel" name="mobile" maxlength="10">
            </div>

            <div class="form-group">

                <label>गांव *</label>

                <select name="village" required>

                    <option value="">
                        -- गांव चुनें --
                    </option>

                    {village_options}

                </select>

            </div>

            <div class="form-group">

                <label>वार्ड</label>

                <input
                    type="text"
                    name="ward"
                    placeholder="उदाहरण: वार्ड 1"
                >

            </div>

            <div class="form-group">

                <label>समस्या की श्रेणी *</label>

                <select name="category" required>

                    <option value="">-- चुनें --</option>

                    <option>पानी</option>
                    <option>सड़क</option>
                    <option>बिजली</option>
                    <option>नाली / सफाई</option>
                    <option>स्ट्रीट लाइट</option>
                    <option>सरकारी भवन</option>
                    <option>अन्य</option>

                </select>

            </div>

            <div class="form-group">

                <label>समस्या का विवरण *</label>

                <textarea
                    name="description"
                    placeholder="अपनी समस्या विस्तार से लिखें..."
                    required
                ></textarea>

            </div>

            <div class="form-group">

                <label>फोटो</label>

                <input
                    type="file"
                    name="photo"
                    accept="image/*"
                >

            </div>

            <div class="form-group">

                <label>स्थान / पहचान</label>

                <input
                    type="text"
                    name="location"
                    placeholder="उदाहरण: स्कूल के पास"
                >

            </div>

            <button class="btn" type="submit">
                शिकायत दर्ज करें
            </button>

        </form>

    </div>

    """

    return render_page(
        "शिकायत दर्ज करें",
        content
    )


# ============================================================
# TRACK COMPLAINT
# ============================================================

@app.route("/track", methods=["GET", "POST"])
def track_complaint():

    complaint_id = request.args.get("complaint_id", "")

    if request.method == "POST":

        complaint_id = request.form.get(
            "complaint_id",
            ""
        ).strip().upper()

    complaint_row = None
    updates = []

    if complaint_id:

        conn = get_db()

        complaint_row = conn.execute(
            """
            SELECT * FROM complaints
            WHERE complaint_id=?
            """,
            (complaint_id,)
        ).fetchone()

        if complaint_row:

            updates = conn.execute(
                """
                SELECT *
                FROM complaint_updates
                WHERE complaint_id=?
                ORDER BY id ASC
                """,
                (complaint_id,)
            ).fetchall()

        conn.close()

    result_html = ""

    if complaint_id and not complaint_row:

        result_html = """

        <div class="alert alert-error">
            यह Complaint ID नहीं मिली।
        </div>

        """

    elif complaint_row:

        timeline = ""

        for u in updates:

            timeline += f"""

            <div class="timeline-item">

                <h3>{u["status"]}</h3>

                <p>{u["message"] or ""}</p>

                <small>{u["created_at"]}</small>

            </div>

            """

        result_html = f"""

        <div class="card">

            <h2>शिकायत: {complaint_row["complaint_id"]}</h2>

            <p>
                <b>गांव:</b>
                {complaint_row["village"]}
            </p>

            <p>
                <b>श्रेणी:</b>
                {complaint_row["category"]}
            </p>

            <p>
                <b>समस्या:</b>
                {complaint_row["description"]}
            </p>

            <p>
                <b>वर्तमान स्थिति:</b>
                <span class="badge">
                    {complaint_row["status"]}
                </span>
            </p>

        </div>

        <div class="card">

            <h2>📍 शिकायत स्थिति</h2>

            <div class="timeline">

                {timeline}

            </div>

        </div>

        """

    content = f"""

    <div class="form-card">

        <h1>🔎 शिकायत ट्रैक करें</h1>

        <p>
            अपना Complaint ID डालकर शिकायत की स्थिति देखें।
        </p>

        <form method="POST">

            <div class="form-group">

                <label>Complaint ID</label>

                <input
                    type="text"
                    name="complaint_id"
                    placeholder="उदाहरण: DVP-2026-ABC123"
                    value="{complaint_id}"
                    required
                >

            </div>

            <button class="btn">
                स्थिति देखें
            </button>

        </form>

    </div>

    {result_html}

    """

    return render_page(
        "शिकायत ट्रैक करें",
        content
    )


# ============================================================
# PROJECTS
# ============================================================

@app.route("/projects")
def projects():

    conn = get_db()

    rows = conn.execute(
        "SELECT * FROM projects ORDER BY id DESC"
    ).fetchall()

    conn.close()

    html = ""

    for p in rows:

        html += f"""

        <div class="card">

            <h2>🏗️ {p["title"]}</h2>

            <p>
                <b>गांव:</b> {p["village"]}
            </p>

            <p>
                {p["description"] or ""}
            </p>

            <p>
                <b>विभाग:</b> {p["department"] or "—"}
            </p>

            <p>
                <b>बजट:</b> {p["budget"] or "—"}
            </p>

            <p>
                <b>प्रगति:</b> {p["progress"]}%
            </p>

            <p>
                <span class="badge">
                    {p["status"]}
                </span>
            </p>

        </div>

        """

    content = f"""

    <h1>🏗️ विकास परियोजनाएं</h1>

    <p>
        पंचायत द्वारा प्रकाशित विकास कार्य एवं उनकी प्रगति।
    </p>

    {html or '<div class="card">अभी कोई विकास कार्य उपलब्ध नहीं है।</div>'}

    """

    return render_page(
        "विकास कार्य",
        content
    )


# ============================================================
# BUDGET
# ============================================================

@app.route("/budget")
def budget():

    conn = get_db()

    rows = conn.execute(
        "SELECT * FROM budget ORDER BY id DESC"
    ).fetchall()

    conn.close()

    html = ""

    for b in rows:

        html += f"""

        <div class="card">

            <h3>💰 {b["category"]}</h3>

            <p>
                <b>वित्तीय वर्ष:</b>
                {b["financial_year"]}
            </p>

            <p>
                <b>राशि:</b>
                ₹ {b["amount"]}
            </p>

            <p>
                {b["description"] or ""}
            </p>

        </div>

        """

    if not html:

        html = """

        <div class="card">

            <h3>💰 बजट जानकारी</h3>

            <p>
                अभी कोई सत्यापित बजट जानकारी प्रकाशित नहीं की गई है।
            </p>

        </div>

        """

    content = f"""

    <h1>💰 बजट एवं व्यय जानकारी</h1>

    <p>
        यहां केवल पंचायत द्वारा प्रकाशित/सत्यापित जानकारी दिखाई जाएगी।
    </p>

    {html}

    """

    return render_page(
        "बजट",
        content
    )


# ============================================================
# MEDICAL
# ============================================================

@app.route("/medical")
def medical():

    conn = get_db()

    rows = conn.execute(
        "SELECT * FROM medical ORDER BY id DESC"
    ).fetchall()

    conn.close()

    html = ""

    for m in rows:

        html += f"""

        <div class="card">

            <h2>🏥 {m["facility_name"]}</h2>

            <p>
                <b>गांव:</b> {m["village"]}
            </p>

            <p>
                <b>प्रकार:</b> {m["type"] or "—"}
            </p>

            <p>
                <b>संपर्क:</b> {m["contact"] or "—"}
            </p>

            <p>
                <b>समय:</b> {m["timing"] or "—"}
            </p>

            <p>
                {m["description"] or ""}
            </p>

        </div>

        """

    content = f"""

    <h1>🏥 चिकित्सा एवं स्वास्थ्य सुविधाएं</h1>

    {html}

    """

    return render_page(
        "स्वास्थ्य सुविधाएं",
        content
    )


# ============================================================
# NOTICES
# ============================================================

@app.route("/notices")
def notices():

    conn = get_db()

    rows = conn.execute(
        "SELECT * FROM notices ORDER BY id DESC"
    ).fetchall()

    conn.close()

    html = ""

    for n in rows:

        html += f"""

        <div class="card">

            <h2>📢 {n["title"]}</h2>

            <small>
                दिनांक: {n["notice_date"]}
            </small>

            <p>
                {n["content"]}
            </p>

        </div>

        """

    content = f"""

    <h1>📢 पंचायत सूचनाएं</h1>

    {html or '<div class="card">अभी कोई सूचना उपलब्ध नहीं है।</div>'}

    """

    return render_page(
        "सूचनाएं",
        content
    )


# ============================================================
# ADMIN LOGIN
# ============================================================

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )

        conn = get_db()

        admin = conn.execute(
            "SELECT * FROM admins WHERE username=?",
            (username,)
        ).fetchone()

        conn.close()

        if admin and check_password_hash(
            admin["password_hash"],
            password
        ):

            session["admin_logged_in"] = True
            session["admin_username"] = username

            flash(
                "Admin login सफल।",
                "success"
            )

            return redirect(
                url_for("admin_dashboard")
            )

        flash(
            "Username या password गलत है।",
            "error"
        )

    content = """

    <div class="form-card" style="max-width:500px;margin:auto;">

        <h1>🔐 Admin / Sarpanch Login</h1>

        <form method="POST">

            <div class="form-group">

                <label>Username</label>

                <input
                    type="text"
                    name="username"
                    required
                >

            </div>

            <div class="form-group">

                <label>Password</label>

                <input
                    type="password"
                    name="password"
                    required
                >

            </div>

            <button class="btn">
                Login
            </button>

        </form>

        <br>

        <small>
            Local demo login:
            admin / admin123
        </small>

    </div>

    """

    return render_page(
        "Admin Login",
        content
    )


@app.route("/admin/logout")
def admin_logout():

    session.clear()

    flash(
        "Admin logout हो गया।",
        "success"
    )

    return redirect(
        url_for("home")
    )


# ============================================================
# ADMIN DASHBOARD
# ============================================================

@app.route("/admin")
def admin_dashboard():

    if not is_admin():
        return redirect(url_for("admin_login"))

    conn = get_db()

    user_count = conn.execute(
        "SELECT COUNT(*) AS c FROM users"
    ).fetchone()["c"]

    complaint_count = conn.execute(
        "SELECT COUNT(*) AS c FROM complaints"
    ).fetchone()["c"]

    project_count = conn.execute(
        "SELECT COUNT(*) AS c FROM projects"
    ).fetchone()["c"]

    notice_count = conn.execute(
        "SELECT COUNT(*) AS c FROM notices"
    ).fetchone()["c"]

    village_counts = conn.execute(
        """
        SELECT village, COUNT(*) AS total
        FROM users
        GROUP BY village
        ORDER BY village
        """
    ).fetchall()

    recent_users = conn.execute(
        """
        SELECT *
        FROM users
        ORDER BY id DESC
        LIMIT 30
        """
    ).fetchall()

    recent_complaints = conn.execute(
        """
        SELECT *
        FROM complaints
        ORDER BY id DESC
        LIMIT 20
        """
    ).fetchall()

    conn.close()

    village_html = ""

    counts_dict = {
        r["village"]: r["total"]
        for r in village_counts
    }

    for v in VILLAGES:

        village_html += f"""

        <div class="stat">

            <div class="stat-number">
                {counts_dict.get(v["name_hi"], 0)}
            </div>

            <div>
                {v["name_hi"]}
            </div>

        </div>

        """

    users_html = ""

    for u in recent_users:

        users_html += f"""

        <tr>

            <td>{u["user_id"]}</td>

            <td>{u["name"]}</td>

            <td>{u["mobile"]}</td>

            <td>{u["village"]}</td>

            <td>{u["registered_at"]}</td>

        </tr>

        """

    complaint_html = ""

    for c in recent_complaints:

        complaint_html += f"""

        <tr>

            <td>{c["complaint_id"]}</td>

            <td>{c["village"]}</td>

            <td>{c["category"]}</td>

            <td>{c["description"]}</td>

            <td>

                <span class="badge">
                    {c["status"]}
                </span>

            </td>

            <td>
                <a
                    class="btn"
                    href="/admin/complaint/{c['id']}"
                >
                    Manage
                </a>
            </td>

        </tr>

        """

    content = f"""

    <div class="hero">

        <h1>⚙️ Admin Dashboard</h1>

        <p>
            Welcome, {session.get("admin_username", "Admin")}
        </p>

    </div>


    <div class="stats">

        <div class="stat">

            <div class="stat-number">
                {user_count}
            </div>

            <div>
                👥 कुल Registered Users
            </div>

        </div>

        <div class="stat">

            <div class="stat-number">
                {complaint_count}
            </div>

            <div>
                📝 कुल Complaints
            </div>

        </div>

        <div class="stat">

            <div class="stat-number">
                {project_count}
            </div>

            <div>
                🏗️ Projects
            </div>

        </div>

        <div class="stat">

            <div class="stat-number">
                {notice_count}
            </div>

            <div>
                📢 Notices
            </div>

        </div>

    </div>


    <h2>👥 Village-wise Registered Users</h2>

    <div class="stats">

        {village_html}

    </div>


    <div class="card">

        <h2>📊 User Data / Excel</h2>

        <p>
            सभी registered users का permanent database record
            Excel में export करें।
        </p>

        <a
            class="btn"
            href="/admin/users"
        >
            👥 Users List
        </a>

        <a
            class="btn"
            href="/admin/download-users-excel"
        >
            📥 Excel Download
        </a>

    </div>


    <div class="card">

        <h2>📝 Complaint Management</h2>

        <div class="table-wrap">

        <table>

            <thead>

                <tr>

                    <th>Complaint ID</th>
                    <th>गांव</th>
                    <th>Category</th>
                    <th>Problem</th>
                    <th>Status</th>
                    <th>Action</th>

                </tr>

            </thead>

            <tbody>

                {complaint_html}

            </tbody>

        </table>

        </div>

    </div>


    <div class="card">

        <h2>➕ Admin Management</h2>

        <a class="btn" href="/admin/project/add">
            Development Project Add
        </a>

        <a class="btn" href="/admin/budget/add">
            Budget Add
        </a>

        <a class="btn" href="/admin/medical/add">
            Medical Facility Add
        </a>

        <a class="btn" href="/admin/notice/add">
            Notice Add
        </a>

    </div>

    """

    return render_page(
        "Admin Dashboard",
        content
    )


# ============================================================
# ADMIN USERS
# ============================================================

@app.route("/admin/users")
def admin_users():
    if not is_admin(): return redirect(url_for("admin_login"))
    selected_village = request.args.get("village", "")
    conn = get_db()
    users = conn.execute("SELECT * FROM users WHERE village=? ORDER BY id DESC", (selected_village,)).fetchall() if selected_village else conn.execute("SELECT * FROM users ORDER BY id DESC").fetchall()
    conn.close()
    options='<option value="">सभी गांव</option>' + ''.join(f'<option value="{v['name_hi']}" {"selected" if selected_village==v["name_hi"] else ""}>{v['name_hi']}</option>' for v in VILLAGES)
    rows=''.join(f'''<tr><td>{u['user_id']}</td><td>{u['name']}</td><td>{u['mobile']}</td><td>{mask_aadhaar(u["aadhaar"])}</td><td>{u['village']}</td><td>{"Verified" if u["mobile_verified"] else "Not Verified"}</td><td>{u['registered_at']}</td><td>{u['last_login'] or "—"}</td><td>{u['login_count'] or 0}</td></tr>''' for u in users)
    content=f'''<h1>👥 Registered Citizens — Sarpanch/Admin Only</h1><div class="card"><form method="GET"><label>Village-wise Filter</label><select name="village" onchange="this.form.submit()">{options}</select></form><br><a class="btn" href="/admin/download-users-excel">📥 Excel Download</a></div><div class="card"><p>कुल registered citizens: <b>{len(users)}</b></p><p>यह जानकारी public website पर नहीं दिखाई जाती। केवल authorized Admin/Sarpanch dashboard से देखी जा सकती है।</p><div class="table-wrap"><table><thead><tr><th>User ID</th><th>Name</th><th>Mobile</th><th>Aadhaar</th><th>Village</th><th>Mobile Verified</th><th>Registered At</th><th>Last Login</th><th>Login Count</th></tr></thead><tbody>{rows}</tbody></table></div></div>'''
    return render_page("Registered Users", content)


# ============================================================
# EXCEL DOWNLOAD

# ============================================================

@app.route("/admin/download-users-excel")
def download_users_excel():
    if not is_admin(): return redirect(url_for("admin_login"))
    conn=get_db()
    users=conn.execute("SELECT user_id,name,mobile,aadhaar,village,mobile_verified,registered_at,last_seen,last_login,login_count,consent FROM users ORDER BY village,id").fetchall()
    conn.close()
    wb=Workbook(); ws=wb.active; ws.title="All Users"
    headers=["User ID","Name","Mobile","Aadhaar","Village","Mobile Verified","Registered At","Last Seen","Last Login","Login Count","Consent"]
    def setup_sheet(sheet,data_rows):
        sheet.append(headers)
        for cell in sheet[1]: cell.font=Font(bold=True); cell.alignment=Alignment(horizontal="center")
        for row in data_rows:
            sheet.append([row["user_id"],row["name"],row["mobile"],row["aadhaar"] or "",row["village"],"Yes" if row["mobile_verified"] else "No",row["registered_at"],row["last_seen"] or "",row["last_login"] or "",row["login_count"] or 0,"Yes" if row["consent"] else "No"])
        for col in range(1,len(headers)+1):
            max_length=max((len(str(c.value)) for c in sheet[get_column_letter(col)] if c.value),default=0)
            sheet.column_dimensions[get_column_letter(col)].width=min(max_length+3,35)
        sheet.freeze_panes="A2"; sheet.auto_filter.ref=sheet.dimensions
    setup_sheet(ws,users)
    for v in VILLAGES:
        safe=v["name_en"].replace("/","-").replace("\\","-")[:31]
        ws_v=wb.create_sheet(safe)
        setup_sheet(ws_v,[u for u in users if u["village"]==v["name_hi"]])
    output=BytesIO(); wb.save(output); output.seek(0)
    filename="Devri_Panchayat_Registered_Users_"+datetime.now().strftime("%Y%m%d_%H%M%S")+".xlsx"
    return send_file(output,as_attachment=True,download_name=filename,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============================================================
# ADMIN COMPLAINT MANAGEMENT

# ============================================================

@app.route("/admin/complaint/<int:complaint_db_id>", methods=["GET", "POST"])
def admin_complaint(complaint_db_id):

    if not is_admin():
        return redirect(url_for("admin_login"))

    conn = get_db()

    complaint_row = conn.execute(
        "SELECT * FROM complaints WHERE id=?",
        (complaint_db_id,)
    ).fetchone()

    if not complaint_row:

        conn.close()

        return "Complaint not found", 404

    if request.method == "POST":

        status = request.form.get(
            "status",
            "दर्ज"
        ).strip()

        message = request.form.get(
            "message",
            ""
        ).strip()

        now = current_time()

        conn.execute(
            """
            UPDATE complaints
            SET status=?, updated_at=?
            WHERE id=?
            """,
            (
                status,
                now,
                complaint_db_id
            )
        )

        conn.execute(
            """
            INSERT INTO complaint_updates
            (complaint_id, status, message, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                complaint_row["complaint_id"],
                status,
                message,
                now
            )
        )

        conn.commit()

        flash(
            "Complaint status updated.",
            "success"
        )

        return redirect(
            url_for(
                "admin_complaint",
                complaint_db_id=complaint_db_id
            )
        )

    updates = conn.execute(
        """
        SELECT *
        FROM complaint_updates
        WHERE complaint_id=?
        ORDER BY id ASC
        """,
        (complaint_row["complaint_id"],)
    ).fetchall()

    conn.close()

    timeline = ""

    for u in updates:

        timeline += f"""

        <div class="timeline-item">

            <h3>{u["status"]}</h3>

            <p>{u["message"] or ""}</p>

            <small>{u["created_at"]}</small>

        </div>

        """

    content = f"""

    <h1>📝 Complaint Management</h1>

    <div class="card">

        <h2>
            {complaint_row["complaint_id"]}
        </h2>

        <p>
            <b>नाम:</b>
            {complaint_row["name"] or "—"}
        </p>

        <p>
            <b>Mobile:</b>
            {complaint_row["mobile"] or "—"}
        </p>

        <p>
            <b>गांव:</b>
            {complaint_row["village"]}
        </p>

        <p>
            <b>Category:</b>
            {complaint_row["category"]}
        </p>

        <p>
            <b>Problem:</b>
            {complaint_row["description"]}
        </p>

        <p>
            <b>Current Status:</b>
            <span class="badge">
                {complaint_row["status"]}
            </span>
        </p>

    </div>


    <div class="form-card">

        <h2>🔄 Update Status</h2>

        <form method="POST">

            <div class="form-group">

                <label>Status</label>

                <select name="status">

                    <option>दर्ज</option>
                    <option>सत्यापित</option>
                    <option>कार्य जारी</option>
                    <option>समाधान</option>

                </select>

            </div>

            <div class="form-group">

                <label>Admin Reply / Update</label>

                <textarea
                    name="message"
                    placeholder="नागरिक को बताने के लिए अपडेट..."
                ></textarea>

            </div>

            <button class="btn">
                Update Complaint
            </button>

        </form>

    </div>


    <div class="card">

        <h2>📍 Status Timeline</h2>

        <div class="timeline">

            {timeline}

        </div>

    </div>

    """

    return render_page(
        "Complaint Management",
        content
    )


# ============================================================
# ADMIN ADD PROJECT
# ============================================================

@app.route("/admin/project/add", methods=["GET", "POST"])
def admin_add_project():

    if not is_admin():
        return redirect(url_for("admin_login"))

    if request.method == "POST":

        village = request.form.get("village")
        title = request.form.get("title")
        description = request.form.get("description")
        department = request.form.get("department")
        budget_value = request.form.get("budget")
        progress = request.form.get("progress", "0")
        status = request.form.get("status")
        start_date = request.form.get("start_date")
        expected_date = request.form.get("expected_date")

        conn = get_db()

        conn.execute(
            """
            INSERT INTO projects
            (
                village,
                title,
                description,
                department,
                budget,
                progress,
                status,
                start_date,
                expected_date,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                village,
                title,
                description,
                department,
                budget_value,
                int(progress or 0),
                status,
                start_date,
                expected_date,
                current_time()
            )
        )

        conn.commit()
        conn.close()

        flash(
            "Development project added.",
            "success"
        )

        return redirect(
            url_for("admin_dashboard")
        )

    village_options = "".join(
        f'<option>{v["name_hi"]}</option>'
        for v in VILLAGES
    )

    content = f"""

    <div class="form-card">

        <h1>🏗️ Development Project Add</h1>

        <form method="POST">

            <div class="form-group">

                <label>गांव</label>

                <select name="village">
                    {village_options}
                </select>

            </div>

            <div class="form-group">

                <label>Project Title</label>

                <input
                    name="title"
                    required
                >

            </div>

            <div class="form-group">

                <label>Description</label>

                <textarea name="description"></textarea>

            </div>

            <div class="form-group">

                <label>Department</label>

                <input name="department">

            </div>

            <div class="form-group">

                <label>Budget</label>

                <input name="budget">

            </div>

            <div class="form-group">

                <label>Progress %</label>

                <input
                    type="number"
                    name="progress"
                    min="0"
                    max="100"
                    value="0"
                >

            </div>

            <div class="form-group">

                <label>Status</label>

                <select name="status">

                    <option>प्रस्तावित</option>
                    <option>स्वीकृत</option>
                    <option>कार्य जारी</option>
                    <option>पूर्ण</option>

                </select>

            </div>

            <div class="form-group">

                <label>Start Date</label>

                <input
                    type="date"
                    name="start_date"
                >

            </div>

            <div class="form-group">

                <label>Expected Completion</label>

                <input
                    type="date"
                    name="expected_date"
                >

            </div>

            <button class="btn">
                Save Project
            </button>

        </form>

    </div>

    """

    return render_page(
        "Add Project",
        content
    )


# ============================================================
# ADMIN ADD BUDGET
# ============================================================

@app.route("/admin/budget/add", methods=["GET", "POST"])
def admin_add_budget():

    if not is_admin():
        return redirect(url_for("admin_login"))

    if request.method == "POST":

        financial_year = request.form.get(
            "financial_year"
        )

        category = request.form.get(
            "category"
        )

        amount = request.form.get(
            "amount"
        )

        description = request.form.get(
            "description"
        )

        conn = get_db()

        conn.execute(
            """
            INSERT INTO budget
            (
                financial_year,
                category,
                amount,
                description,
                created_at
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                financial_year,
                category,
                amount,
                description,
                current_time()
            )
        )

        conn.commit()
        conn.close()

        flash(
            "Budget information added.",
            "success"
        )

        return redirect(
            url_for("admin_dashboard")
        )

    content = """

    <div class="form-card">

        <h1>💰 Budget Information Add</h1>

        <form method="POST">

            <div class="form-group">

                <label>Financial Year</label>

                <input
                    name="financial_year"
                    placeholder="2026-27"
                    required
                >

            </div>

            <div class="form-group">

                <label>Category</label>

                <input
                    name="category"
                    placeholder="सड़क विकास"
                    required
                >

            </div>

            <div class="form-group">

                <label>Amount</label>

                <input
                    name="amount"
                    placeholder="500000"
                    required
                >

            </div>

            <div class="form-group">

                <label>Description</label>

                <textarea name="description"></textarea>

            </div>

            <button class="btn">
                Save Budget
            </button>

        </form>

    </div>

    """

    return render_page(
        "Add Budget",
        content
    )


# ============================================================
# ADMIN ADD MEDICAL
# ============================================================

@app.route("/admin/medical/add", methods=["GET", "POST"])
def admin_add_medical():

    if not is_admin():
        return redirect(url_for("admin_login"))

    if request.method == "POST":

        village = request.form.get("village")
        facility_name = request.form.get("facility_name")
        facility_type = request.form.get("type")
        contact = request.form.get("contact")
        timing = request.form.get("timing")
        description = request.form.get("description")

        conn = get_db()

        conn.execute(
            """
            INSERT INTO medical
            (
                village,
                facility_name,
                type,
                contact,
                timing,
                description,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                village,
                facility_name,
                facility_type,
                contact,
                timing,
                description,
                current_time()
            )
        )

        conn.commit()
        conn.close()

        flash(
            "Medical facility added.",
            "success"
        )

        return redirect(
            url_for("admin_dashboard")
        )

    village_options = "".join(
        f'<option>{v["name_hi"]}</option>'
        for v in VILLAGES
    )

    content = f"""

    <div class="form-card">

        <h1>🏥 Medical Facility Add</h1>

        <form method="POST">

            <div class="form-group">

                <label>गांव</label>

                <select name="village">

                    {village_options}

                </select>

            </div>

            <div class="form-group">

                <label>Facility Name</label>

                <input
                    name="facility_name"
                    required
                >

            </div>

            <div class="form-group">

                <label>Type</label>

                <input name="type">

            </div>

            <div class="form-group">

                <label>Contact</label>

                <input name="contact">

            </div>

            <div class="form-group">

                <label>Timing</label>

                <input name="timing">

            </div>

            <div class="form-group">

                <label>Description</label>

                <textarea name="description"></textarea>

            </div>

            <button class="btn">
                Save Facility
            </button>

        </form>

    </div>

    """

    return render_page(
        "Add Medical Facility",
        content
    )


# ============================================================
# ADMIN ADD NOTICE
# ============================================================

@app.route("/admin/notice/add", methods=["GET", "POST"])
def admin_add_notice():

    if not is_admin():
        return redirect(url_for("admin_login"))

    if request.method == "POST":

        title = request.form.get("title")
        content_text = request.form.get("content")
        notice_date = request.form.get("notice_date")

        conn = get_db()

        conn.execute(
            """
            INSERT INTO notices
            (
                title,
                content,
                notice_date,
                created_at
            )
            VALUES (?, ?, ?, ?)
            """,
            (
                title,
                content_text,
                notice_date,
                current_time()
            )
        )

        conn.commit()
        conn.close()

        flash(
            "Notice added.",
            "success"
        )

        return redirect(
            url_for("admin_dashboard")
        )

    today = datetime.now().strftime("%Y-%m-%d")

    content = f"""

    <div class="form-card">

        <h1>📢 Notice Add</h1>

        <form method="POST">

            <div class="form-group">

                <label>Notice Title</label>

                <input
                    name="title"
                    required
                >

            </div>

            <div class="form-group">

                <label>Notice Content</label>

                <textarea
                    name="content"
                    required
                ></textarea>

            </div>

            <div class="form-group">

                <label>Date</label>

                <input
                    type="date"
                    name="notice_date"
                    value="{today}"
                    required
                >

            </div>

            <button class="btn">
                Publish Notice
            </button>

        </form>

    </div>

    """

    return render_page(
        "Add Notice",
        content
    )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    print("=" * 60)
    print("DIGITAL DEVRI PANCHAYAT")
    print("=" * 60)

    print()
    print("Website: http://127.0.0.1:5000")
    print()
    print("Admin Username:", ADMIN_USERNAME)
    print("Admin Password:", ADMIN_PASSWORD)
    print()
    print("Database:", DB_FILE)
    print("Villages: Devri, Parliya, Hapawas, Ghagsa, Thukrawa")
    print("Excel: Admin Dashboard -> Registered Users -> Excel Download")
    print()
    print("Running...")
    print("=" * 60)

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )