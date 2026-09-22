from flask import Flask, request, redirect, url_for, session, render_template_string, flash, send_file
import sqlite3
from pathlib import Path
from datetime import datetime
import os
import secrets
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO

# ============================================================
# DIGITAL DEVRI PANCHAYAT
# Updated version:
# - Budget module completely removed from public/admin UI
# - Citizen complaint system retained
# - Admin complaint progress management added
# - Exactly 5 villages
# - Medical, notices, schemes and development information retained
# - Single-file Flask + SQLite application
# ============================================================

app = Flask(__name__)

# Use environment variable on production; keep fallback for the existing project.
app.secret_key = os.environ.get(
    "SECRET_KEY",
    "DEVRI_PANCHAYAT_SECRET_KEY_2026_CHANGE_THIS"
)

BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "devri_panchayat.db"
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")

ALLOWED_IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "webp"}

# ------------------------------------------------------------
# EXACTLY 5 VILLAGES
# ------------------------------------------------------------

VILLAGES = [
    {"id": 1, "name_hi": "देवरी", "name_en": "Devri"},
    {"id": 2, "name_hi": "पारलिया", "name_en": "Parliya"},
    {"id": 3, "name_hi": "हापावास", "name_en": "Hapawas"},
    {"id": 4, "name_hi": "घाघसा", "name_en": "Ghagsa"},
    {"id": 5, "name_hi": "ठुकरावा", "name_en": "Thukrawa"},
]

VILLAGE_NAMES = [v["name_hi"] for v in VILLAGES]

COMPLAINT_STATUSES = [
    "दर्ज",
    "जांच में",
    "कार्यवाही शुरू",
    "कार्य जारी",
    "समाधान",
    "बंद",
]

COMPLAINT_CATEGORIES = [
    "पानी",
    "सड़क",
    "बिजली",
    "नाली / सफाई",
    "स्ट्रीट लाइट",
    "सरकारी भवन",
    "स्वास्थ्य",
    "अन्य",
]


# ============================================================
# DATABASE HELPERS
# ============================================================

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def current_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def column_names(conn, table_name):
    return {
        row["name"]
        for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    }


def add_column_if_missing(conn, table_name, column_name, column_type):
    if column_name not in column_names(conn, table_name):
        conn.execute(
            f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}"
        )


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
            description TEXT DEFAULT '',
            population TEXT DEFAULT '',
            facilities TEXT DEFAULT ''
        )
    """)

    # Citizens
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

    # Preserve the existing user database and add missing fields.
    for name, typ in {
        "aadhaar": "TEXT",
        "mobile_verified": "INTEGER DEFAULT 0",
        "verified_at": "TEXT",
        "last_login": "TEXT",
        "login_count": "INTEGER DEFAULT 0",
    }.items():
        add_column_if_missing(conn, "users", name, typ)

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
            action_taken TEXT DEFAULT '',
            work_start_date TEXT DEFAULT '',
            completion_percent INTEGER DEFAULT 0,
            expected_completion_date TEXT DEFAULT '',
            remarks TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)

    # Preserve complaints created by the older version and migrate them.
    for name, typ in {
        "action_taken": "TEXT DEFAULT ''",
        "work_start_date": "TEXT DEFAULT ''",
        "completion_percent": "INTEGER DEFAULT 0",
        "expected_completion_date": "TEXT DEFAULT ''",
        "remarks": "TEXT DEFAULT ''",
    }.items():
        add_column_if_missing(conn, "complaints", name, typ)

    # Complaint timeline
    cur.execute("""
        CREATE TABLE IF NOT EXISTS complaint_updates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            complaint_id TEXT NOT NULL,
            status TEXT NOT NULL,
            message TEXT,
            created_at TEXT NOT NULL
        )
    """)

    # Development information.
    # IMPORTANT: budget is not displayed or editable anywhere in this version.
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

    # Medical facilities
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
    if cur.execute("SELECT COUNT(*) AS c FROM panchayat").fetchone()["c"] == 0:
        cur.execute("""
            INSERT INTO panchayat
            (name_hi, name_en, description_hi, contact, email)
            VALUES (?, ?, ?, ?, ?)
        """, (
            "देवरी ग्राम पंचायत",
            "Devri Gram Panchayat",
            "देवरी ग्राम पंचायत का डिजिटल नागरिक सेवा पोर्टल। यहां नागरिक शिकायत दर्ज कर सकते हैं और उसकी स्थिति देख सकते हैं।",
            "जानकारी अपडेट की जानी है",
            "जानकारी अपडेट की जानी है",
        ))

    # Keep exactly the 5 current villages.
    placeholders = ",".join("?" for _ in VILLAGE_NAMES)
    cur.execute(
        f"DELETE FROM villages WHERE name_hi NOT IN ({placeholders})",
        VILLAGE_NAMES
    )

    for village in VILLAGES:
        existing = cur.execute(
            "SELECT id FROM villages WHERE name_hi=?",
            (village["name_hi"],)
        ).fetchone()

        if existing:
            cur.execute("""
                UPDATE villages
                SET name_en=?, description='', population='', facilities=''
                WHERE name_hi=?
            """, (
                village["name_en"],
                village["name_hi"],
            ))
        else:
            cur.execute("""
                INSERT INTO villages
                (name_hi, name_en, description, population, facilities)
                VALUES (?, ?, '', '', '')
            """, (
                village["name_hi"],
                village["name_en"],
            ))

    # Remove old project/medical records belonging to obsolete villages.
    cur.execute(
        f"DELETE FROM projects WHERE village NOT IN ({placeholders})",
        VILLAGE_NAMES
    )
    cur.execute(
        f"DELETE FROM medical WHERE village NOT IN ({placeholders})",
        VILLAGE_NAMES
    )

    # Remove the old budget table/data completely if it exists.
    # The website no longer has any budget module.
    cur.execute("DROP TABLE IF EXISTS budget")

    # Admin
    if cur.execute("SELECT COUNT(*) AS c FROM admins").fetchone()["c"] == 0:
        cur.execute("""
            INSERT INTO admins(username, password_hash)
            VALUES (?, ?)
        """, (
            ADMIN_USERNAME,
            generate_password_hash(ADMIN_PASSWORD),
        ))

    # Keep one simple development record only if none exists.
    # It contains no public budget information.
    if cur.execute("SELECT COUNT(*) AS c FROM projects").fetchone()["c"] == 0:
        cur.execute("""
            INSERT INTO projects
            (village, title, description, department, budget,
             progress, status, start_date, expected_date, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            "देवरी",
            "ग्राम सड़क विकास कार्य",
            "सड़क सुधार एवं मरम्मत कार्य की जानकारी।",
            "ग्राम पंचायत",
            "",
            0,
            "जानकारी अपडेट की जानी है",
            "",
            "",
            current_time(),
        ))

    # Keep one simple medical placeholder if database is empty.
    if cur.execute("SELECT COUNT(*) AS c FROM medical").fetchone()["c"] == 0:
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
            current_time(),
        ))

    # Notice
    if cur.execute("SELECT COUNT(*) AS c FROM notices").fetchone()["c"] == 0:
        cur.execute("""
            INSERT INTO notices(title, content, notice_date, created_at)
            VALUES (?, ?, ?, ?)
        """, (
            "देवरी ग्राम पंचायत डिजिटल पोर्टल",
            "पंचायत से संबंधित जानकारी और शिकायतों के लिए इस पोर्टल का उपयोग करें।",
            datetime.now().strftime("%Y-%m-%d"),
            current_time(),
        ))

    conn.commit()
    conn.close()


init_db()


# ============================================================
# HELPERS
# ============================================================

def is_admin():
    return session.get("admin_logged_in") is True


def generate_user_id():
    while True:
        value = "DVP-U-" + secrets.token_hex(4).upper()
        conn = get_db()
        found = conn.execute(
            "SELECT id FROM users WHERE user_id=?",
            (value,)
        ).fetchone()
        conn.close()
        if not found:
            return value


def generate_complaint_id():
    while True:
        value = (
            "DVP-"
            + str(datetime.now().year)
            + "-"
            + secrets.token_hex(3).upper()
        )
        conn = get_db()
        found = conn.execute(
            "SELECT id FROM complaints WHERE complaint_id=?",
            (value,)
        ).fetchone()
        conn.close()
        if not found:
            return value


def normalize_digits(value):
    return "".join(ch for ch in (value or "") if ch.isdigit())


def mask_aadhaar(value):
    digits = normalize_digits(value)
    if len(digits) == 12:
        return "XXXX-XXXX-" + digits[-4:]
    return "—"


def allowed_image(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_IMAGE_EXTENSIONS
    )


def render_page(title, content):
    return render_template_string(
        BASE_TEMPLATE,
        title=title,
        content=content,
        admin=is_admin(),
    )


# ============================================================
# BASE TEMPLATE
# ============================================================

BASE_TEMPLATE = """
<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{{ title }} | Digital Devri Panchayat</title>

<style>
*{box-sizing:border-box}
body{
    margin:0;
    font-family:Arial,"Noto Sans Devanagari",sans-serif;
    background:#f4f7f3;
    color:#1f2937;
}
a{text-decoration:none;color:inherit}
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
.logo{font-size:21px;font-weight:bold}
.navlinks{display:flex;gap:7px;flex-wrap:wrap;justify-content:flex-end}
.navlinks a{padding:8px 10px;border-radius:8px;font-size:14px}
.navlinks a:hover{background:#166534}
.container{width:90%;max-width:1200px;margin:28px auto}
.hero{
    background:linear-gradient(135deg,#166534,#22c55e);
    color:white;
    border-radius:22px;
    padding:40px 28px;
    margin-bottom:25px;
}
.hero h1{font-size:36px;margin:0 0 12px}
.hero p{font-size:18px;line-height:1.7}
.btn{
    display:inline-block;
    background:#166534;
    color:white;
    padding:11px 17px;
    border:0;
    border-radius:9px;
    cursor:pointer;
    font-weight:bold;
    margin:3px;
}
.btn:hover{background:#14532d}
.btn-light{background:white;color:#166534}
.card-grid{
    display:grid;
    grid-template-columns:repeat(auto-fit,minmax(230px,1fr));
    gap:18px;
}
.card{
    background:white;
    border-radius:16px;
    padding:21px;
    box-shadow:0 3px 15px rgba(0,0,0,.08);
    margin-bottom:18px;
}
.card h2,.card h3{color:#14532d}
.icon{font-size:34px}
.form-card{
    background:white;
    padding:25px;
    border-radius:18px;
    box-shadow:0 3px 15px rgba(0,0,0,.08);
    max-width:850px;
    margin:auto;
}
.form-group{margin-bottom:16px}
label{display:block;margin-bottom:7px;font-weight:bold}
input,select,textarea{
    width:100%;
    padding:12px;
    border:1px solid #d1d5db;
    border-radius:8px;
    font-size:15px;
}
textarea{min-height:120px;resize:vertical}
.stats{
    display:grid;
    grid-template-columns:repeat(auto-fit,minmax(170px,1fr));
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
.stat-number{font-size:30px;font-weight:bold;color:#166534}
table{width:100%;border-collapse:collapse;background:white}
th,td{
    padding:11px;
    border-bottom:1px solid #e5e7eb;
    text-align:left;
    vertical-align:top;
}
th{background:#dcfce7;color:#14532d}
.table-wrap{overflow-x:auto;border-radius:12px}
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
.alert-error{background:#fee2e2;color:#991b1b}
.timeline{border-left:3px solid #22c55e;padding-left:20px}
.timeline-item{margin-bottom:20px;position:relative}
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
.progress-wrap{
    background:#e5e7eb;
    height:24px;
    border-radius:20px;
    overflow:hidden;
    margin:10px 0;
}
.progress-bar{
    background:#22c55e;
    height:100%;
    color:white;
    text-align:center;
    line-height:24px;
    font-size:13px;
    font-weight:bold;
}
.progress-small{
    color:#6b7280;
    font-size:14px;
}
.footer{
    margin-top:45px;
    background:#14532d;
    color:white;
    padding:28px 5%;
    text-align:center;
}
.mobile-nav{display:none}

@media(max-width:800px){
    .navbar{flex-direction:column;gap:12px}
    .navlinks{justify-content:center}
    .hero h1{font-size:28px}
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
    .mobile-nav a{font-size:12px;text-align:center}
    body{padding-bottom:65px}
}
</style>
</head>

<body>

<div class="navbar">
    <div class="logo">🌿 Digital Devri Panchayat</div>

    <div class="navlinks">
        <a href="{{ url_for('home') }}">🏠 होम</a>
        <a href="{{ url_for('villages') }}">🏘️ गांव</a>
        <a href="{{ url_for('complaint') }}">📝 शिकायत</a>
        <a href="{{ url_for('track_complaint') }}">🔎 शिकायत ट्रैक</a>
        <a href="{{ url_for('projects') }}">🏗️ विकास</a>
        <a href="{{ url_for('medical') }}">🏥 स्वास्थ्य</a>
        <a href="{{ url_for('notices') }}">📢 सूचनाएं</a>
        <a href="{{ url_for('login') }}">👤 नागरिक लॉगिन</a>

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
    <a href="{{ url_for('login') }}">👤<br>लॉगिन</a>
</div>

<div class="footer">
    <h3>🌿 Digital Devri Panchayat</h3>
    <p>ग्रामीण नागरिकों के लिए सरल डिजिटल पंचायत सेवा पोर्टल</p>
    <p>© 2026 Devri Gram Panchayat</p>
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
    complaint_count = conn.execute(
        "SELECT COUNT(*) AS c FROM complaints"
    ).fetchone()["c"]
    project_count = conn.execute(
        "SELECT COUNT(*) AS c FROM projects"
    ).fetchone()["c"]
    notice_count = conn.execute(
        "SELECT COUNT(*) AS c FROM notices"
    ).fetchone()["c"]
    notices_rows = conn.execute(
        "SELECT * FROM notices ORDER BY id DESC LIMIT 3"
    ).fetchall()
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
        f"""
        <div class="card">
            <div class="icon">{icon}</div>
            <h3>{name}</h3>
            <p>पात्रता और वर्तमान जानकारी आधिकारिक पोर्टल पर जांचें।</p>
        </div>
        """
        for icon, name in schemes
    )

    notice_cards = "".join(
        f"""
        <div class="card">
            <h3>{n["title"]}</h3>
            <p>{n["content"]}</p>
            <small>{n["notice_date"]}</small>
        </div>
        """
        for n in notices_rows
    )

    if not notice_cards:
        notice_cards = '<div class="card">अभी कोई सूचना उपलब्ध नहीं है।</div>'

    content = f"""
    <div class="hero">
        <h1>🌿 डिजिटल देवरी ग्राम पंचायत</h1>
        <p>
            पंचायत की जानकारी, 5 गांवों की जानकारी, शिकायत,
            विकास कार्य और पंचायत सूचनाएं — सब एक ही स्थान पर।
        </p>
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
        <div class="card">
            <div class="icon">🏘️</div>
            <h3>हमारे गांव</h3>
            <p>देवरी पंचायत के 5 गांव देखें।</p>
            <a class="btn" href="/villages">देखें</a>
        </div>

        <div class="card">
            <div class="icon">📝</div>
            <h3>शिकायत दर्ज करें</h3>
            <p>पानी, सड़क, बिजली आदि से जुड़ी समस्या दर्ज करें।</p>
            <a class="btn" href="/complaint">शिकायत करें</a>
        </div>

        <div class="card">
            <div class="icon">🔎</div>
            <h3>शिकायत ट्रैक करें</h3>
            <p>अपने Complaint ID से स्थिति और काम की प्रगति देखें।</p>
            <a class="btn" href="/track">ट्रैक करें</a>
        </div>

        <div class="card">
            <div class="icon">🏗️</div>
            <h3>विकास कार्य</h3>
            <p>उपलब्ध विकास कार्य और उनकी स्थिति देखें।</p>
            <a class="btn" href="/projects">देखें</a>
        </div>

        <div class="card">
            <div class="icon">🏥</div>
            <h3>स्वास्थ्य सुविधाएं</h3>
            <p>उपलब्ध स्वास्थ्य सुविधा की जानकारी देखें।</p>
            <a class="btn" href="/medical">देखें</a>
        </div>

        <div class="card">
            <div class="icon">📢</div>
            <h3>पंचायत सूचनाएं</h3>
            <p>जरूरी पंचायत सूचनाएं देखें।</p>
            <a class="btn" href="/notices">देखें</a>
        </div>
    </div>

    <br>
    <h2>🏛️ सरकारी योजनाएं</h2>
    <p>प्रमुख योजनाओं के नाम और आधिकारिक जानकारी के लिए नीचे देखें।</p>
    <div class="card-grid">{scheme_cards}</div>

    <div class="card">
        <h3>🔗 आधिकारिक योजना जानकारी</h3>
        <p>राजस्थान की योजनाओं की जानकारी आधिकारिक Jan Soochna Portal पर देखें।</p>
        <a class="btn" href="https://jansoochna.rajasthan.gov.in/Scheme"
           target="_blank" rel="noopener">
           Jan Soochna Portal खोलें
        </a>
    </div>

    <h2>📢 नवीनतम सूचनाएं</h2>
    {notice_cards}
    """

    return render_page("होम", content)


# ============================================================
# CITIZEN REGISTRATION
# ============================================================

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        mobile = normalize_digits(request.form.get("mobile", ""))
        village = request.form.get("village", "").strip()
        aadhaar = normalize_digits(request.form.get("aadhaar", ""))
        consent = request.form.get("consent")

        if not name or not mobile or not village or not aadhaar or not consent:
            flash("कृपया सभी जरूरी जानकारी भरें।", "error")
            return redirect(url_for("register"))

        if len(mobile) != 10:
            flash("कृपया 10 अंकों का मोबाइल नंबर डालें।", "error")
            return redirect(url_for("register"))

        if len(aadhaar) != 12:
            flash("कृपया 12 अंकों का Aadhaar नंबर डालें।", "error")
            return redirect(url_for("register"))

        if village not in VILLAGE_NAMES:
            flash("कृपया सही गांव चुनें।", "error")
            return redirect(url_for("register"))

        now = current_time()
        conn = get_db()

        existing = conn.execute(
            "SELECT * FROM users WHERE mobile=?",
            (mobile,)
        ).fetchone()

        if existing:
            user_id = existing["user_id"]
            conn.execute("""
                UPDATE users
                SET name=?, village=?, aadhaar=?, mobile_verified=1,
                    verified_at=?, last_seen=?, consent=1
                WHERE mobile=?
            """, (
                name, village, aadhaar, now, now, mobile
            ))
            message = f"पंजीकरण अपडेट और लॉगिन सफल। आपका User ID: {user_id}"
        else:
            user_id = generate_user_id()
            conn.execute("""
                INSERT INTO users
                (user_id,name,mobile,village,aadhaar,registered_at,last_seen,
                 consent,mobile_verified,verified_at,last_login,login_count)
                VALUES (?,?,?,?,?,?,?,1,1,?,?,1)
            """, (
                user_id, name, mobile, village, aadhaar,
                now, now, now, now
            ))
            message = f"पंजीकरण सफल। आपका User ID: {user_id}"

        conn.commit()
        conn.close()

        session["citizen_user_id"] = user_id
        flash(message, "success")
        return redirect(url_for("home"))

    village_options = "".join(
        f'<option value="{v["name_hi"]}">{v["name_hi"]} ({v["name_en"]})</option>'
        for v in VILLAGES
    )

    content = f"""
    <div class="form-card">
        <h1>👤 नागरिक पंजीकरण</h1>
        <p>अपना नाम, मोबाइल, Aadhaar और गांव भरें।</p>

        <form method="POST">
            <div class="form-group">
                <label>नाम *</label>
                <input name="name" required>
            </div>

            <div class="form-group">
                <label>मोबाइल नंबर *</label>
                <input name="mobile" inputmode="numeric"
                       maxlength="10" pattern="[0-9]{{10}}" required>
            </div>

            <div class="form-group">
                <label>Aadhaar नंबर *</label>
                <input name="aadhaar" inputmode="numeric"
                       maxlength="12" pattern="[0-9]{{12}}" required>
            </div>

            <div class="form-group">
                <label>गांव *</label>
                <select name="village" required>
                    <option value="">गांव चुनें</option>
                    {village_options}
                </select>
            </div>

            <div class="form-group">
                <label>
                    <input type="checkbox" name="consent" value="1" required>
                    मैं अपनी जानकारी सुरक्षित रिकॉर्ड में सेव करने की सहमति देता/देती हूं।
                </label>
            </div>

            <button class="btn">✅ Register</button>
        </form>
    </div>
    """

    return render_page("नागरिक पंजीकरण", content)


# ============================================================
# CITIZEN LOGIN
# ============================================================

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        mobile = normalize_digits(request.form.get("mobile", ""))
        village = request.form.get("village", "").strip()

        if len(mobile) != 10:
            flash("कृपया 10 अंकों का मोबाइल नंबर डालें।", "error")
            return redirect(url_for("login"))

        if village not in VILLAGE_NAMES:
            flash("कृपया सही गांव चुनें।", "error")
            return redirect(url_for("login"))

        conn = get_db()
        user = conn.execute(
            "SELECT * FROM users WHERE mobile=? AND village=?",
            (mobile, village)
        ).fetchone()

        if not user:
            conn.close()
            flash(
                "Mobile number और गांव का रिकॉर्ड नहीं मिला। पहले registration करें।",
                "error"
            )
            return redirect(url_for("login"))

        now = current_time()

        conn.execute("""
            UPDATE users
            SET last_seen=?, last_login=?,
                login_count=COALESCE(login_count,0)+1,
                mobile_verified=1
            WHERE mobile=? AND village=?
        """, (now, now, mobile, village))

        conn.commit()
        conn.close()

        session["citizen_user_id"] = user["user_id"]
        flash("नागरिक लॉगिन सफल।", "success")
        return redirect(url_for("home"))

    village_options = "".join(
        f'<option value="{v["name_hi"]}">{v["name_hi"]} ({v["name_en"]})</option>'
        for v in VILLAGES
    )

    content = f"""
    <div class="form-card">
        <h1>🔐 नागरिक लॉगिन</h1>
        <p>Registered mobile number और गांव चुनें।</p>

        <form method="POST">
            <div class="form-group">
                <label>Registered Mobile Number *</label>
                <input type="tel" name="mobile" inputmode="numeric"
                       maxlength="10" pattern="[0-9]{{10}}" required>
            </div>

            <div class="form-group">
                <label>गांव *</label>
                <select name="village" required>
                    <option value="">गांव चुनें</option>
                    {village_options}
                </select>
            </div>

            <button class="btn">🔐 Login</button>
        </form>

        <br>
        <a class="btn btn-light" href="/register">नया नागरिक पंजीकरण</a>
    </div>
    """

    return render_page("नागरिक लॉगिन", content)


# ============================================================
# VILLAGES
# ============================================================

@app.route("/villages")
def villages():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM villages ORDER BY id"
    ).fetchall()
    conn.close()

    cards = "".join(
        f"""
        <div class="card">
            <div class="icon">🏘️</div>
            <h3>{v["name_hi"]}</h3>
            <p><b>English:</b> {v["name_en"]}</p>
            <a class="btn" href="/village/{v["id"]}">गांव देखें</a>
        </div>
        """
        for v in rows
    )

    content = f"""
    <h1>🏘️ हमारे गांव</h1>
    <p>देवरी ग्राम पंचायत के अंतर्गत आने वाले 5 गांव।</p>
    <div class="card-grid">{cards}</div>
    """

    return render_page("हमारे गांव", content)


@app.route("/village/<int:village_id>")
def village_detail(village_id):
    conn = get_db()

    village = conn.execute(
        "SELECT * FROM villages WHERE id=?",
        (village_id,)
    ).fetchone()

    if not village:
        conn.close()
        return "Village not found", 404

    projects_rows = conn.execute(
        "SELECT * FROM projects WHERE village=? ORDER BY id DESC",
        (village["name_hi"],)
    ).fetchall()

    medical_rows = conn.execute(
        "SELECT * FROM medical WHERE village=? ORDER BY id DESC",
        (village["name_hi"],)
    ).fetchall()

    conn.close()

    project_html = "".join(
        f"""
        <div class="card">
            <h3>🏗️ {p["title"]}</h3>
            <p>{p["description"] or ""}</p>
            <p><b>विभाग:</b> {p["department"] or "—"}</p>
            <p><b>प्रगति:</b> {int(p["progress"] or 0)}%</p>
            <div class="progress-wrap">
                <div class="progress-bar" style="width:{max(0,min(100,int(p["progress"] or 0)))}%">
                    {max(0,min(100,int(p["progress"] or 0)))}%
                </div>
            </div>
            <p><span class="badge">{p["status"]}</span></p>
        </div>
        """
        for p in projects_rows
    )

    if not project_html:
        project_html = '<div class="card">अभी कोई विकास कार्य उपलब्ध नहीं है।</div>'

    medical_html = "".join(
        f"""
        <div class="card">
            <h3>🏥 {m["facility_name"]}</h3>
            <p><b>प्रकार:</b> {m["type"] or "—"}</p>
            <p><b>संपर्क:</b> {m["contact"] or "—"}</p>
            <p><b>समय:</b> {m["timing"] or "—"}</p>
            <p>{m["description"] or ""}</p>
        </div>
        """
        for m in medical_rows
    )

    if not medical_html:
        medical_html = '<div class="card">अभी कोई स्वास्थ्य सुविधा दर्ज नहीं है।</div>'

    content = f"""
    <div class="hero">
        <h1>🏘️ {village["name_hi"]}</h1>
        <p>{village["name_en"]}</p>
    </div>

    <div class="card">
        <h2>गांव की जानकारी</h2>
        <p><b>नाम:</b> {village["name_hi"]}</p>
        <p><b>English Name:</b> {village["name_en"]}</p>
    </div>

    <h2>🏗️ विकास कार्य</h2>
    {project_html}

    <h2>🏥 स्वास्थ्य सुविधाएं</h2>
    {medical_html}
    """

    return render_page(village["name_hi"], content)


# ============================================================
# COMPLAINT SUBMISSION
# ============================================================

@app.route("/complaint", methods=["GET", "POST"])
def complaint():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        mobile = normalize_digits(request.form.get("mobile", ""))
        village = request.form.get("village", "").strip()
        ward = request.form.get("ward", "").strip()
        category = request.form.get("category", "").strip()
        description = request.form.get("description", "").strip()
        location = request.form.get("location", "").strip()

        if not village or not category or not description:
            flash("गांव, समस्या की श्रेणी और समस्या का विवरण जरूरी है।", "error")
            return redirect(url_for("complaint"))

        if mobile and len(mobile) != 10:
            flash("मोबाइल नंबर 10 अंकों का होना चाहिए।", "error")
            return redirect(url_for("complaint"))

        if village not in VILLAGE_NAMES:
            flash("कृपया सही गांव चुनें।", "error")
            return redirect(url_for("complaint"))

        complaint_id = generate_complaint_id()
        now = current_time()
        photo_filename = ""

        uploaded_file = request.files.get("photo")

        if uploaded_file and uploaded_file.filename:
            if not allowed_image(uploaded_file.filename):
                flash("कृपया JPG, JPEG, PNG या WEBP फोटो अपलोड करें।", "error")
                return redirect(url_for("complaint"))

            original_name = secure_filename(uploaded_file.filename)
            photo_filename = complaint_id + "_" + original_name
            uploaded_file.save(UPLOAD_DIR / photo_filename)

        conn = get_db()

        conn.execute("""
            INSERT INTO complaints
            (complaint_id,name,mobile,village,ward,category,description,
             photo,location,status,action_taken,work_start_date,
             completion_percent,expected_completion_date,remarks,
             created_at,updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
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
            "",
            "",
            0,
            "",
            "",
            now,
            now,
        ))

        conn.execute("""
            INSERT INTO complaint_updates
            (complaint_id,status,message,created_at)
            VALUES (?,?,?,?)
        """, (
            complaint_id,
            "दर्ज",
            "शिकायत सफलतापूर्वक दर्ज की गई।",
            now,
        ))

        conn.commit()
        conn.close()

        flash(
            f"शिकायत दर्ज हो गई। आपका Complaint ID है: {complaint_id}",
            "success"
        )

        return redirect(
            url_for("track_complaint", complaint_id=complaint_id)
        )

    village_options = "".join(
        f'<option value="{v["name_hi"]}">{v["name_hi"]}</option>'
        for v in VILLAGES
    )

    category_options = "".join(
        f"<option>{c}</option>"
        for c in COMPLAINT_CATEGORIES
    )

    content = f"""
    <div class="form-card">
        <h1>📝 शिकायत दर्ज करें</h1>
        <p>अपने गांव की समस्या दर्ज करें।</p>

        <form method="POST" enctype="multipart/form-data">

            <div class="form-group">
                <label>नाम</label>
                <input type="text" name="name">
            </div>

            <div class="form-group">
                <label>मोबाइल नंबर</label>
                <input type="tel" name="mobile" maxlength="10" inputmode="numeric">
            </div>

            <div class="form-group">
                <label>गांव *</label>
                <select name="village" required>
                    <option value="">-- गांव चुनें --</option>
                    {village_options}
                </select>
            </div>

            <div class="form-group">
                <label>वार्ड</label>
                <input type="text" name="ward" placeholder="उदाहरण: वार्ड 1">
            </div>

            <div class="form-group">
                <label>समस्या की श्रेणी *</label>
                <select name="category" required>
                    <option value="">-- चुनें --</option>
                    {category_options}
                </select>
            </div>

            <div class="form-group">
                <label>समस्या का विवरण *</label>
                <textarea name="description"
                          placeholder="अपनी समस्या विस्तार से लिखें..."
                          required></textarea>
            </div>

            <div class="form-group">
                <label>फोटो</label>
                <input type="file" name="photo" accept="image/*">
            </div>

            <div class="form-group">
                <label>स्थान / पहचान</label>
                <input type="text" name="location"
                       placeholder="उदाहरण: स्कूल के पास">
            </div>

            <button class="btn" type="submit">📝 शिकायत दर्ज करें</button>
        </form>
    </div>
    """

    return render_page("शिकायत दर्ज करें", content)


# ============================================================
# TRACK COMPLAINT
# ============================================================

@app.route("/track", methods=["GET", "POST"])
def track_complaint():
    complaint_id = request.args.get("complaint_id", "").strip().upper()

    if request.method == "POST":
        complaint_id = request.form.get(
            "complaint_id", ""
        ).strip().upper()

    complaint_row = None
    updates = []

    if complaint_id:
        conn = get_db()

        complaint_row = conn.execute(
            "SELECT * FROM complaints WHERE complaint_id=?",
            (complaint_id,)
        ).fetchone()

        if complaint_row:
            updates = conn.execute("""
                SELECT * FROM complaint_updates
                WHERE complaint_id=?
                ORDER BY id ASC
            """, (complaint_id,)).fetchall()

        conn.close()

    result_html = ""

    if complaint_id and not complaint_row:
        result_html = """
        <div class="alert alert-error">
            यह Complaint ID नहीं मिली।
        </div>
        """

    elif complaint_row:
        completion = max(
            0,
            min(100, int(complaint_row["completion_percent"] or 0))
        )
        remaining = 100 - completion

        timeline = "".join(
            f"""
            <div class="timeline-item">
                <h3>{u["status"]}</h3>
                <p>{u["message"] or ""}</p>
                <small>{u["created_at"]}</small>
            </div>
            """
            for u in updates
        )

        action = complaint_row["action_taken"] or "अभी कोई कार्यवाही अपडेट नहीं है।"
        start_date = complaint_row["work_start_date"] or "—"
        expected_date = complaint_row["expected_completion_date"] or "—"
        remarks = complaint_row["remarks"] or "—"

        result_html = f"""
        <div class="card">
            <h2>शिकायत: {complaint_row["complaint_id"]}</h2>

            <p><b>गांव:</b> {complaint_row["village"]}</p>
            <p><b>वार्ड:</b> {complaint_row["ward"] or "—"}</p>
            <p><b>श्रेणी:</b> {complaint_row["category"]}</p>
            <p><b>समस्या:</b> {complaint_row["description"]}</p>

            <p>
                <b>वर्तमान स्थिति:</b>
                <span class="badge">{complaint_row["status"]}</span>
            </p>
        </div>

        <div class="card">
            <h2>🛠️ काम की प्रगति</h2>

            <p><b>काम पूरा:</b> {completion}%</p>

            <div class="progress-wrap">
                <div class="progress-bar" style="width:{completion}%">
                    {completion}%
                </div>
            </div>

            <p><b>काम बाकी:</b> {remaining}%</p>

            <p><b>काम शुरू होने की तारीख:</b> {start_date}</p>
            <p><b>संभावित पूरा होने की तारीख:</b> {expected_date}</p>

            <p><b>की गई कार्यवाही:</b><br>{action}</p>
            <p><b>टिप्पणी:</b><br>{remarks}</p>
        </div>

        <div class="card">
            <h2>📍 शिकायत स्थिति</h2>
            <div class="timeline">{timeline}</div>
        </div>
        """

    content = f"""
    <div class="form-card">
        <h1>🔎 शिकायत ट्रैक करें</h1>
        <p>अपना Complaint ID डालकर शिकायत की स्थिति और काम की प्रगति देखें।</p>

        <form method="POST">
            <div class="form-group">
                <label>Complaint ID</label>
                <input type="text" name="complaint_id"
                       placeholder="उदाहरण: DVP-2026-ABC123"
                       value="{complaint_id}" required>
            </div>
            <button class="btn">स्थिति देखें</button>
        </form>
    </div>

    {result_html}
    """

    return render_page("शिकायत ट्रैक करें", content)


# ============================================================
# DEVELOPMENT WORKS
# ============================================================

@app.route("/projects")
def projects():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM projects ORDER BY id DESC"
    ).fetchall()
    conn.close()

    cards = ""

    for p in rows:
        progress = max(0, min(100, int(p["progress"] or 0)))

        cards += f"""
        <div class="card">
            <h2>🏗️ {p["title"]}</h2>

            <p><b>गांव:</b> {p["village"]}</p>
            <p>{p["description"] or ""}</p>
            <p><b>विभाग:</b> {p["department"] or "—"}</p>

            <p><b>प्रगति:</b> {progress}%</p>

            <div class="progress-wrap">
                <div class="progress-bar" style="width:{progress}%">
                    {progress}%
                </div>
            </div>

            <p><span class="badge">{p["status"]}</span></p>
        </div>
        """

    if not cards:
        cards = '<div class="card">अभी कोई विकास कार्य उपलब्ध नहीं है।</div>'

    content = f"""
    <h1>🏗️ विकास कार्य</h1>
    <p>उपलब्ध विकास कार्य और उनकी स्थिति।</p>
    {cards}
    """

    return render_page("विकास कार्य", content)


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

    cards = "".join(
        f"""
        <div class="card">
            <h2>🏥 {m["facility_name"]}</h2>
            <p><b>गांव:</b> {m["village"]}</p>
            <p><b>प्रकार:</b> {m["type"] or "—"}</p>
            <p><b>संपर्क:</b> {m["contact"] or "—"}</p>
            <p><b>समय:</b> {m["timing"] or "—"}</p>
            <p>{m["description"] or ""}</p>
        </div>
        """
        for m in rows
    )

    if not cards:
        cards = '<div class="card">अभी कोई स्वास्थ्य सुविधा दर्ज नहीं है।</div>'

    content = f"""
    <h1>🏥 चिकित्सा एवं स्वास्थ्य सुविधाएं</h1>
    {cards}
    """

    return render_page("स्वास्थ्य सुविधाएं", content)


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

    cards = "".join(
        f"""
        <div class="card">
            <h2>📢 {n["title"]}</h2>
            <small>दिनांक: {n["notice_date"]}</small>
            <p>{n["content"]}</p>
        </div>
        """
        for n in rows
    )

    if not cards:
        cards = '<div class="card">अभी कोई सूचना उपलब्ध नहीं है।</div>'

    content = f"""
    <h1>📢 पंचायत सूचनाएं</h1>
    {cards}
    """

    return render_page("सूचनाएं", content)


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

        # Fallback login keeps the original project credentials working
        # even when an older database contains a different password hash.
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session["admin_logged_in"] = True
            session["admin_username"] = username
            flash("Admin login सफल।", "success")
            return redirect(url_for("admin_dashboard"))

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

            <td>{int(c["completion_percent"] or 0)}%</td>

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
                    <th>Progress</th>
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
        status = request.form.get("status", "दर्ज").strip()
        action_taken = request.form.get("action_taken", "").strip()
        work_start_date = request.form.get("work_start_date", "").strip()
        expected_completion_date = request.form.get(
            "expected_completion_date", ""
        ).strip()
        remarks = request.form.get("remarks", "").strip()

        try:
            completion_percent = int(
                request.form.get("completion_percent", "0")
            )
        except ValueError:
            completion_percent = 0

        completion_percent = max(0, min(100, completion_percent))

        if status not in COMPLAINT_STATUSES:
            status = "दर्ज"

        now = current_time()

        # Store complaint progress.
        conn.execute("""
            UPDATE complaints
            SET status=?,
                action_taken=?,
                work_start_date=?,
                completion_percent=?,
                expected_completion_date=?,
                remarks=?,
                updated_at=?
            WHERE id=?
        """, (
            status,
            action_taken,
            work_start_date,
            completion_percent,
            expected_completion_date,
            remarks,
            now,
            complaint_db_id,
        ))

        # Add a timeline entry.
        timeline_message = (
            f"कार्यवाही: {action_taken or '—'} | "
            f"काम पूरा: {completion_percent}% | "
            f"काम बाकी: {100 - completion_percent}% | "
            f"काम शुरू: {work_start_date or '—'} | "
            f"संभावित पूरा: {expected_completion_date or '—'} | "
            f"टिप्पणी: {remarks or '—'}"
        )

        conn.execute("""
            INSERT INTO complaint_updates
            (complaint_id,status,message,created_at)
            VALUES (?,?,?,?)
        """, (
            complaint_row["complaint_id"],
            status,
            timeline_message,
            now,
        ))

        conn.commit()

        flash("Complaint और work progress update हो गया।", "success")

        return redirect(
            url_for(
                "admin_complaint",
                complaint_db_id=complaint_db_id
            )
        )

    updates = conn.execute("""
        SELECT *
        FROM complaint_updates
        WHERE complaint_id=?
        ORDER BY id ASC
    """, (
        complaint_row["complaint_id"],
    )).fetchall()

    conn.close()

    completion = max(
        0,
        min(100, int(complaint_row["completion_percent"] or 0))
    )

    remaining = 100 - completion

    timeline = "".join(
        f"""
        <div class="timeline-item">
            <h3>{u["status"]}</h3>
            <p>{u["message"] or ""}</p>
            <small>{u["created_at"]}</small>
        </div>
        """
        for u in updates
    )

    content = f"""
    <h1>📝 Complaint Management</h1>

    <div class="card">
        <h2>{complaint_row["complaint_id"]}</h2>

        <p><b>नाम:</b> {complaint_row["name"] or "—"}</p>
        <p><b>Mobile:</b> {complaint_row["mobile"] or "—"}</p>
        <p><b>गांव:</b> {complaint_row["village"]}</p>
        <p><b>वार्ड:</b> {complaint_row["ward"] or "—"}</p>
        <p><b>Category:</b> {complaint_row["category"]}</p>
        <p><b>Problem:</b> {complaint_row["description"]}</p>

        <p>
            <b>Current Status:</b>
            <span class="badge">{complaint_row["status"]}</span>
        </p>
    </div>

    <div class="form-card">
        <h2>🔄 काम की प्रगति अपडेट करें</h2>

        <form method="POST">

            <div class="form-group">
                <label>Status</label>
                <select name="status">
                    {"".join(
                        f'<option {"selected" if complaint_row["status"] == s else ""}>{s}</option>'
                        for s in COMPLAINT_STATUSES
                    )}
                </select>
            </div>

            <div class="form-group">
                <label>की गई कार्यवाही</label>
                <textarea name="action_taken"
                    placeholder="उदाहरण: संबंधित विभाग को सूचना दी गई...">{complaint_row["action_taken"] or ""}</textarea>
            </div>

            <div class="form-group">
                <label>काम शुरू होने की तारीख</label>
                <input type="date"
                       name="work_start_date"
                       value="{complaint_row["work_start_date"] or ""}">
            </div>

            <div class="form-group">
                <label>काम कितना पूरा हुआ? (%)</label>
                <input type="number"
                       name="completion_percent"
                       min="0"
                       max="100"
                       value="{completion}">
                <p class="progress-small">
                    बाकी काम वेबसाइट अपने-आप {remaining}% calculate करेगी।
                </p>
            </div>

            <div class="form-group">
                <label>संभावित पूरा होने की तारीख</label>
                <input type="date"
                       name="expected_completion_date"
                       value="{complaint_row["expected_completion_date"] or ""}">
            </div>

            <div class="form-group">
                <label>अतिरिक्त टिप्पणी</label>
                <textarea name="remarks"
                    placeholder="अन्य जरूरी जानकारी...">{complaint_row["remarks"] or ""}</textarea>
            </div>

            <button class="btn">💾 Update Complaint</button>
        </form>
    </div>

    <div class="card">
        <h2>📊 वर्तमान काम की स्थिति</h2>

        <p><b>काम पूरा:</b> {completion}%</p>

        <div class="progress-wrap">
            <div class="progress-bar" style="width:{completion}%">
                {completion}%
            </div>
        </div>

        <p><b>काम बाकी:</b> {remaining}%</p>

        <p>
            <b>काम शुरू:</b>
            {complaint_row["work_start_date"] or "—"}
        </p>

        <p>
            <b>संभावित पूरा:</b>
            {complaint_row["expected_completion_date"] or "—"}
        </p>
    </div>

    <div class="card">
        <h2>📍 Status Timeline</h2>
        <div class="timeline">{timeline}</div>
    </div>
    """

    return render_page("Complaint Management", content)


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
# ============================================================
# 404
# ============================================================

@app.errorhandler(404)
def page_not_found(error):
    content = """
    <div class="card">
        <h1>404</h1>
        <p>यह पेज उपलब्ध नहीं है।</p>
        <a class="btn" href="/">होम पर जाएं</a>
    </div>
    """
    return render_page("Page Not Found", content), 404


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))

    print("=" * 60)
    print("DIGITAL DEVRI PANCHAYAT")
    print("=" * 60)
    print("Website: http://127.0.0.1:5000")
    print("Admin Username:", ADMIN_USERNAME)
    print("Admin Password:", ADMIN_PASSWORD)
    print("Database:", DB_FILE)
    print("Villages: Devri, Parliya, Hapawas, Ghagsa, Thukrawa")
    print("Budget module: REMOVED")
    print("Complaint work-progress management: ENABLED")
    print("=" * 60)

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
