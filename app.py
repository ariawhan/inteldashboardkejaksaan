import os
import re
import stat
import uuid
import base64
import hashlib
import json
import shutil
import threading
import time
import zipfile
from datetime import date
from functools import wraps
from io import BytesIO
from pathlib import Path
from urllib.parse import quote, urlparse

import pymysql
import requests
from bs4 import BeautifulSoup, NavigableString, Tag
from cryptography.fernet import Fernet, InvalidToken
from dotenv import load_dotenv
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from flask import Flask, abort, flash, g, jsonify, redirect, render_template, request, send_file, send_from_directory, session, url_for
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as reportlab_canvas
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from PIL import Image as PILImage
from playwright.sync_api import sync_playwright
from xml.sax.saxutils import escape
from werkzeug.utils import secure_filename
from werkzeug.security import check_password_hash, generate_password_hash


load_dotenv()

app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.getenv("SECRET_KEY", "dev-secret-ganti-saat-production"),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    PERMANENT_SESSION_LIFETIME=60 * 60 * 12,
    MAX_CONTENT_LENGTH=20 * 1024 * 1024,
)

DB_CONFIG = {
    "host": os.getenv("DB_HOST", "127.0.0.1"),
    "port": int(os.getenv("DB_PORT", "3306")),
    "user": os.getenv("DB_USER", "root"),
    "password": os.getenv("DB_PASSWORD", ""),
    "charset": "utf8mb4",
    "cursorclass": pymysql.cursors.DictCursor,
}
DB_NAME = os.getenv("DB_NAME", "intel_dashboard")
MIGRATIONS_DIR = Path(__file__).resolve().parent / "migrations"
STATIC_IMG_DIR = Path(__file__).resolve().parent / "static" / "img"
UPLOAD_DIR = Path(__file__).resolve().parent / "uploads" / "reports"
ALLOWED_IMAGE_TYPES = {"image/jpeg", "image/png", "image/webp"}
ISSUE_CODES = {"Ds.1", "Ds.2", "Ds.3", "Dip.1", "Dip.2", "Dip.3", "Dip.4",
               "Dsb.1", "Dsb.2", "Dsb.3", "Dsb.4", "Dek.1", "Dek.2", "Dek.3", "Dek.4",
               "Dpp.1", "Dpp.2", "Dpp.3", "Dpp.4", "Dti.1", "Dti.2", "Dti.3", "Dti.4"}
CHROME_EXECUTABLE = Path(os.getenv("CHROME_EXECUTABLE", r"C:\Program Files\Google\Chrome\Application\chrome.exe"))
INTELIZ_LOGIN_URL = "https://inteliz.kejaksaan.go.id/login"
INTELIZ_2FA_PATH = "/2fa/challenge"
INTELIZ_LAPINHAR_CREATE_URL = "https://inteliz.kejaksaan.go.id/lapinhar/create"
INTELIZ_AUTH_SESSIONS = {}
INTELIZ_AUTH_LOCK = threading.Lock()
SIPEDE_LOGIN_URL = "https://sipede.kejaksaan.go.id/login"
SIPEDE_BASE_URL = "https://sipede.kejaksaan.go.id"
SIPEDE_SURATKELUAR_CREATE_URL = "https://sipede.kejaksaan.go.id/suratkeluar/create?idSurat=125"
SIPEDE_AUTH_SESSIONS = {}
SIPEDE_AUTH_LOCK = threading.Lock()


class IntelizAuthenticationRequired(RuntimeError):
    pass

PDF_FONT_NAME = "Times-Roman"
PDF_FONT_BOLD = "Times-Bold"
PDF_FONT_ITALIC = "Times-Italic"
PDF_FONT_BOLD_ITALIC = "Times-BoldItalic"
WINDOWS_FONT_DIR = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
try:
    pdfmetrics.registerFont(TTFont("TimesNewRoman", str(WINDOWS_FONT_DIR / "times.ttf")))
    pdfmetrics.registerFont(TTFont("TimesNewRoman-Bold", str(WINDOWS_FONT_DIR / "timesbd.ttf")))
    pdfmetrics.registerFont(TTFont("TimesNewRoman-Italic", str(WINDOWS_FONT_DIR / "timesi.ttf")))
    pdfmetrics.registerFont(TTFont("TimesNewRoman-BoldItalic", str(WINDOWS_FONT_DIR / "timesbi.ttf")))
    pdfmetrics.registerFontFamily("TimesNewRoman", normal="TimesNewRoman", bold="TimesNewRoman-Bold",
                                  italic="TimesNewRoman-Italic", boldItalic="TimesNewRoman-BoldItalic")
    PDF_FONT_NAME, PDF_FONT_BOLD = "TimesNewRoman", "TimesNewRoman-Bold"
    PDF_FONT_ITALIC, PDF_FONT_BOLD_ITALIC = "TimesNewRoman-Italic", "TimesNewRoman-BoldItalic"
except (OSError, ValueError):
    pass


def initialize_database():
    connection = pymysql.connect(**DB_CONFIG)
    try:
        with connection.cursor() as cursor:
            cursor.execute(
                f"CREATE DATABASE IF NOT EXISTS `{DB_NAME}` "
                "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
            cursor.execute(f"USE `{DB_NAME}`")
            cursor.execute(
                """CREATE TABLE IF NOT EXISTS schema_migrations (
                    version VARCHAR(100) PRIMARY KEY,
                    applied_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                ) ENGINE=InnoDB"""
            )
            cursor.execute("SELECT version FROM schema_migrations")
            applied = {row["version"] for row in cursor.fetchall()}
            for migration_file in sorted(MIGRATIONS_DIR.glob("*.sql")):
                if migration_file.name in applied:
                    continue
                statements = [
                    statement.strip()
                    for statement in migration_file.read_text(encoding="utf-8").split(";")
                    if statement.strip()
                ]
                for statement in statements:
                    cursor.execute(statement)
                cursor.execute(
                    "INSERT INTO schema_migrations (version) VALUES (%s)",
                    (migration_file.name,),
                )
            admin_username = os.getenv("ADMIN_USERNAME", "admin")
            cursor.execute("SELECT id FROM users WHERE username = %s", (admin_username,))
            if cursor.fetchone() is None:
                cursor.execute(
                    "INSERT INTO users (username, full_name, password_hash, role) VALUES (%s, %s, %s, 'admin')",
                    (
                        admin_username,
                        os.getenv("ADMIN_FULL_NAME", "Administrator"),
                        generate_password_hash(os.getenv("ADMIN_PASSWORD", "admin123")),
                    ),
                )
        connection.commit()
    finally:
        connection.close()


@app.cli.command("migrate")
def migrate_command():
    """Jalankan seluruh migration database yang belum diterapkan."""
    initialize_database()
    print(f"Migration database '{DB_NAME}' selesai.")


def get_db():
    if "db" not in g:
        g.db = pymysql.connect(database=DB_NAME, autocommit=True, **DB_CONFIG)
    return g.db


@app.teardown_appcontext
def close_db(_error=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def fetch_one(query, params=()):
    with get_db().cursor() as cursor:
        cursor.execute(query, params)
        return cursor.fetchone()


def fetch_all(query, params=()):
    with get_db().cursor() as cursor:
        cursor.execute(query, params)
        return cursor.fetchall()


def login_required(view):
    @wraps(view)
    def wrapped_view(**kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return view(**kwargs)
    return wrapped_view


def admin_required(view):
    @wraps(view)
    @login_required
    def wrapped_view(**kwargs):
        if session.get("role") != "admin":
            flash("Menu tersebut hanya dapat diakses oleh admin.", "error")
            return redirect(url_for("dashboard"))
        return view(**kwargs)
    return wrapped_view


@app.route("/", methods=["GET", "POST"])
def login():
    if session.get("user_id"):
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        user = fetch_one("SELECT * FROM users WHERE username = %s", (username,))
        if user is None or not check_password_hash(user["password_hash"], request.form.get("password", "")):
            flash("Username atau kata sandi tidak sesuai.", "error")
        elif not user.get("is_active", 1):
            flash("Akun Anda sedang dinonaktifkan. Hubungi administrator.", "error")
        else:
            session.clear()
            session.permanent = request.form.get("remember") == "on"
            session.update(
                user_id=user["id"], full_name=user["full_name"],
                username=user["username"], role=user["role"]
            )
            return redirect(url_for("dashboard"))
    return render_template("login.html")


@app.route("/dashboard")
@login_required
def dashboard():
    report_filter = "" if session["role"] == "admin" else " WHERE created_by = %s"
    params = () if session["role"] == "admin" else (session["user_id"],)
    rows = fetch_all(
        "SELECT report_type, COUNT(*) total, SUM(status = 'draft') draft "
        f"FROM reports{report_filter} GROUP BY report_type", params
    )
    counts = {"lapinhar": {"total": 0, "draft": 0}, "lapinsus": {"total": 0, "draft": 0}}
    for row in rows:
        counts[row["report_type"]] = {"total": row["total"], "draft": int(row["draft"] or 0)}
    user_count = fetch_one("SELECT COUNT(*) total FROM users")["total"] if session["role"] == "admin" else None
    return render_template("dashboard.html", counts=counts, user_count=user_count, active="dashboard")


@app.route("/lapinhar")
@login_required
def lapinhar():
    conditions = ["reports.report_type = 'lapinhar'"]
    params = []
    base_where = " AND ".join(conditions)
    summary = fetch_one(
        f"""SELECT COUNT(*) AS total,
            SUM(reports.lapinsus_status='belum') AS lapinsus_belum,
            SUM(reports.inteliz_status='belum') AS inteliz_belum
            FROM reports WHERE {base_where}""", tuple(params)
    )
    report_summary = {
        "total": int(summary["total"] or 0),
        "lapinsus_belum": int(summary["lapinsus_belum"] or 0),
        "inteliz_belum": int(summary["inteliz_belum"] or 0),
    }

    search = request.args.get("q", "").strip()[:150]
    if search:
        search_columns = ["reports.report_number LIKE %s", "reports.title LIKE %s"]
        term = f"%{search}%"
        params.extend([term, term])
        search_columns.append("users.full_name LIKE %s")
        params.append(term)
        conditions.append(f"({' OR '.join(search_columns)})")
    where_clause = " AND ".join(conditions)
    try:
        page = max(1, int(request.args.get("page", "1")))
    except ValueError:
        page = 1
    per_page = 10
    filtered_total = fetch_one(
        f"""SELECT COUNT(*) AS total FROM reports
            JOIN users ON users.id=reports.created_by WHERE {where_clause}""", tuple(params)
    )["total"]
    total_pages = max(1, (int(filtered_total) + per_page - 1) // per_page)
    page = min(page, total_pages)
    offset = (page - 1) * per_page
    reports = fetch_all(
        "SELECT reports.id, reports.report_number, reports.title, reports.created_by, reports.inteliz_status, "
        "reports.lapinsus_status, reports.created_at, users.full_name AS creator_full_name, "
        "(SELECT COUNT(*) FROM report_attachments WHERE report_attachments.report_id=reports.id) AS attachment_count "
        "FROM reports JOIN users ON users.id = reports.created_by "
        f"WHERE {where_clause} ORDER BY reports.created_at DESC LIMIT %s OFFSET %s",
        tuple(params + [per_page, offset]),
    )
    return render_template("reports.html", report_type="LAPINHAR", reports=reports,
                           report_summary=report_summary, active="lapinhar",
                           search=search, page=page, total_pages=total_pages,
                           filtered_total=int(filtered_total), per_page=per_page,
                           auto_connect=request.args.get("connect_inteliz") == "1",
                           next_url=url_for("lapinhar"))


def normalize_whatsapp_number(value):
    """Convert common Indonesian phone formats into the international wa.me format."""
    number = re.sub(r"\D", "", value or "")
    if number.startswith("00"):
        number = number[2:]
    elif number.startswith("0"):
        number = "62" + number[1:]
    elif number.startswith("8"):
        number = "62" + number
    return number


def whatsapp_plain_text(content):
    """Turn editor HTML into readable WhatsApp text while preserving list numbering."""
    if not content:
        return "-"
    if "<" not in content:
        return content.strip() or "-"
    soup = BeautifulSoup(content, "html.parser")
    lines = []

    def clean_text(value):
        return re.sub(r"\s+", " ", value or "").strip()

    def render_list(list_node, level=0):
        for index, item in enumerate(list_node.find_all("li", recursive=False), 1):
            nested = item.find(["ol", "ul"], recursive=False)
            text_parts = [str(child) for child in item.contents if child is not nested]
            item_text = clean_text(BeautifulSoup("".join(text_parts), "html.parser").get_text(" ", strip=True))
            marker = f"{index}." if list_node.name == "ol" else "•"
            if item_text:
                lines.append(f"{'   ' * level}{marker} {item_text}")
            if nested:
                render_list(nested, level + 1)

    for node in soup.find_all(["p", "ol", "ul"], recursive=False):
        if node.name in {"ol", "ul"}:
            render_list(node)
        else:
            text = clean_text(node.get_text(" ", strip=True))
            if text:
                lines.append(text)
    if not lines:
        fallback = [clean_text(part) for part in soup.get_text("\n").splitlines()]
        lines.extend(part for part in fallback if part)
    return "\n".join(lines).strip() or "-"


@app.get("/lapinhar/<int:report_id>/whatsapp")
@login_required
def send_lapinhar_whatsapp(report_id):
    report = accessible_lapinhar(report_id)
    if report is None:
        flash("Laporan tidak ditemukan atau tidak dapat Anda akses.", "error")
        return redirect(url_for("lapinhar"))
    setting = fetch_one(
        "SELECT contact_name, phone_number FROM whatsapp_user_settings WHERE user_id=%s",
        (session["user_id"],),
    )
    if not setting:
        flash("Simpan nama dan nomor tujuan WhatsApp terlebih dahulu.", "error")
        return redirect(url_for("integration_settings", _anchor="whatsapp"))
    phone = normalize_whatsapp_number(setting["phone_number"])
    if not re.fullmatch(r"\d{9,15}", phone):
        flash("Nomor WhatsApp pada konfigurasi tidak valid.", "error")
        return redirect(url_for("integration_settings", _anchor="whatsapp"))
    organization_name = (report.get("organization") or "KEJAKSAAN NEGERI BULELENG").title()
    sender_name = f"Kepala Seksi Intelijen {organization_name}"
    recipient_name = setting["contact_name"].rstrip(". ") + "."
    message = "\n".join([
        "*Laporan Harian*", "", "Kepada Yth.", recipient_name, "",
        f"Dari : {sender_name}", "", f"Perihal : {report['title']}", "",
        "*I. INFORMASI YANG DIPEROLEH*",
        whatsapp_plain_text(report["facts"]), "", "*II. SUMBER INFORMASI*",
        whatsapp_plain_text(report["source_name"]), "", "*III. TREN PERKEMBANGAN / PERKIRAAN*",
        whatsapp_plain_text(report["analysis"]), "", "*IV. PENDAPAT / SARAN*",
        whatsapp_plain_text(report["recommendation"]), "", "Dum🙏",
    ])
    attachment_rows = fetch_all(
        """SELECT attachment_group, sort_order, image_path
           FROM report_attachments WHERE report_id=%s
           ORDER BY attachment_group, sort_order, id""",
        (report_id,),
    ) if request.args.get("attachments") == "1" else []
    downloads = ([url_for("download_all_lapinhar_attachments", report_id=report_id)]
                 if attachment_rows else [])
    whatsapp_url = f"https://web.whatsapp.com/send?phone={phone}&text={quote(message)}"
    return render_template("whatsapp_launch.html", whatsapp_url=whatsapp_url,
                           downloads=downloads, attachment_count=len(attachment_rows),
                           uses_archive=len(attachment_rows) > 1)


@app.route("/lapinsus")
@login_required
def lapinsus():
    conditions = ["reports.report_type = 'lapinsus'"]
    params = []
    summary_where = " AND ".join(conditions)
    summary = fetch_one(
        f"""SELECT COUNT(*) total,
            SUM(inteliz_status='belum') inteliz_belum,
            SUM(sipede_status='belum') sipede_belum
            FROM lapinsus_reports reports WHERE {summary_where}""", tuple(params),
    )
    report_summary = {
        "total": int(summary["total"] or 0),
        "inteliz_belum": int(summary["inteliz_belum"] or 0),
        "sipede_belum": int(summary["sipede_belum"] or 0),
    }
    search = request.args.get("q", "").strip()[:150]
    if search:
        columns = ["reports.report_number LIKE %s", "reports.title LIKE %s"]
        term = f"%{search}%"
        params.extend([term, term])
        columns.append("users.full_name LIKE %s")
        params.append(term)
        conditions.append(f"({' OR '.join(columns)})")
    where_clause = " AND ".join(conditions)
    try:
        page = max(1, int(request.args.get("page", "1")))
    except ValueError:
        page = 1
    per_page = 10
    filtered_total = int(fetch_one(
        f"SELECT COUNT(*) total FROM reports JOIN users ON users.id=reports.created_by WHERE {where_clause}",
        tuple(params),
    )["total"] or 0)
    total_pages = max(1, (filtered_total + per_page - 1) // per_page)
    page = min(page, total_pages)
    reports = fetch_all(
        "SELECT reports.id, reports.report_number, reports.title, reports.created_by, reports.status, reports.inteliz_status, "
        "reports.sipede_status, reports.created_at, "
        "users.full_name creator_full_name FROM reports JOIN users ON users.id=reports.created_by "
        f"WHERE {where_clause} ORDER BY reports.created_at DESC LIMIT %s OFFSET %s",
        tuple(params + [per_page, (page - 1) * per_page]),
    )
    return render_template("reports.html", report_type="LAPINSUS", reports=reports, active="lapinsus",
                           report_summary=report_summary,
                           search=search, page=page, total_pages=total_pages,
                           filtered_total=filtered_total, per_page=per_page)


def accessible_lapinsus(report_id):
    report = fetch_one("SELECT * FROM reports WHERE id=%s AND report_type='lapinsus'", (report_id,))
    if report and (session.get("role") == "admin" or report["created_by"] == session.get("user_id")):
        return report
    return None


def lapinsus_sequence_number(report_number):
    match = re.match(r"^R\.LIK-(\d+)/", report_number or "")
    return int(match.group(1)) if match else 0


def compose_lapinsus_number(sequence_number, institution_code, issue_code, report_date):
    try:
        selected_date = date.fromisoformat(report_date)
    except (TypeError, ValueError):
        return ""
    if not sequence_number or issue_code not in ISSUE_CODES:
        return ""
    return f"R.LIK-{sequence_number}/{institution_code}/{issue_code}/{selected_date.month:02d}/{selected_date.year}"


def reserve_lapinsus_number(document_year=None):
    document_year = int(document_year or date.today().year)
    connection = get_db()
    token = uuid.uuid4().hex
    try:
        connection.begin()
        with connection.cursor() as cursor:
            cursor.execute("SELECT next_number FROM document_counters WHERE document_type='lapinsus' AND document_year=%s FOR UPDATE", (document_year,))
            counter = cursor.fetchone()
            if counter is None:
                sequence_number = 1
                cursor.execute("INSERT INTO document_counters (document_type,document_year,next_number) VALUES ('lapinsus',%s,2)", (document_year,))
            else:
                sequence_number = counter["next_number"]
                cursor.execute("UPDATE document_counters SET next_number=next_number+1 WHERE document_type='lapinsus' AND document_year=%s", (document_year,))
            cursor.execute(
                "INSERT INTO document_number_reservations "
                "(reservation_token,document_type,document_year,sequence_number,created_by) VALUES (%s,'lapinsus',%s,%s,%s)",
                (token, document_year, sequence_number, session["user_id"]),
            )
        connection.commit()
        return {"reservation_token": token, "sequence_number": sequence_number, "document_year": document_year}
    except Exception:
        connection.rollback()
        raise


def lapinsus_number_available(report_number, reservation_token="", report_id=None):
    duplicate = fetch_one("SELECT id FROM reports WHERE report_number=%s", (report_number,))
    if duplicate and duplicate["id"] != report_id:
        return False, "Nomor surat sudah digunakan laporan lain. Silakan reload nomor surat."
    if report_id:
        return True, "Nomor surat tersedia."
    reservation = fetch_one(
        "SELECT sequence_number,document_year FROM document_number_reservations WHERE reservation_token=%s "
        "AND document_type='lapinsus' AND created_by=%s AND status='reserved'",
        (reservation_token, session["user_id"]),
    )
    if (not reservation or reservation["sequence_number"] != lapinsus_sequence_number(report_number)
            or reservation["document_year"] != report_number_year(report_number)):
        return False, "Reservasi nomor surat tidak berlaku. Silakan reload nomor surat."
    return True, "Nomor surat tersedia."


@app.post("/lapinsus/check-number")
@login_required
def check_lapinsus_number():
    payload = request.get_json(silent=True) or {}
    try:
        report_id = int(payload["report_id"]) if payload.get("report_id") else None
    except (TypeError, ValueError):
        report_id = None
    if report_id and accessible_lapinsus(report_id) is None:
        abort(404)
    available, message = lapinsus_number_available(
        str(payload.get("report_number", "")).strip(),
        str(payload.get("reservation_token", "")).strip(), report_id,
    )
    return jsonify(available=available, message=message), 200 if available else 409


@app.post("/lapinsus/reload-number")
@login_required
def reload_lapinsus_number():
    payload = request.get_json(silent=True) or {}
    reservation = reserve_lapinsus_number(payload.get("document_year"))
    return jsonify(**reservation)


def lapinhar_form_data():
    fields = ("report_number", "subject", "report_date", "recipient", "sender", "classification", "category_id",
              "attachment", "organization", "information", "sources", "trends", "suggestions",
              "city", "creator_position", "creator_name", "creator_rank_nip",
              "auth_position", "auth_name", "auth_rank_nip", "information_spacing",
              "sources_spacing", "trends_spacing", "suggestions_spacing")
    data = {field: request.form.get(field, "").strip() for field in fields}
    organization = fetch_one("SELECT * FROM organization_settings WHERE id = 1") or {
        "organization_name": "KEJAKSAAN NEGERI BULELENG", "address": "", "phone": "", "website": ""
    }
    data["organization"] = organization["organization_name"]
    data["institution_code"] = organization.get("institution_code") or "N.1.11"
    data["recipient"] = f"YTH. KEPALA {organization['organization_name']}"
    data["sender"] = "KASI INTELIJEN"
    data["classification"] = "RAHASIA"
    issue_code = request.form.get("issue_code", "").strip()
    data["issue_code"] = issue_code if issue_code in ISSUE_CODES else ""
    try:
        category_id = int(data["category_id"])
        data["category_id"] = category_id if 1 <= category_id <= 74 else None
    except (TypeError, ValueError):
        data["category_id"] = None
    for field in ("information_spacing", "sources_spacing", "trends_spacing", "suggestions_spacing"):
        data[field] = "1.5"
    data["city"] = os.getenv("SIGNING_CITY", "Singaraja")
    kasi = fetch_one("SELECT * FROM signatories WHERE position_code = 'kasi_intel'")
    subsection_code = request.form.get("subsection_signer", "kasubsi_1")
    if subsection_code not in {"kasi_intel", "kasubsi_1", "kasubsi_2"}:
        subsection_code = "kasubsi_1"
    subsection = fetch_one(
        "SELECT * FROM signatories WHERE position_code = %s", (subsection_code,)
    )
    data["show_authentication"] = subsection_code != "kasi_intel"
    if kasi:
        data.update(auth_position=kasi["position_name"], auth_name=kasi["full_name"], auth_rank_nip=kasi["rank_nip"])
    if subsection:
        data.update(creator_position=subsection["position_name"], creator_name=subsection["full_name"],
                    creator_rank_nip=subsection["rank_nip"])
    else:
        data.update(creator_position="", creator_name="", creator_rank_nip="")
    return data


def reserve_lapinhar_number(document_year=None):
    document_year = int(document_year or date.today().year)
    connection = get_db()
    token = uuid.uuid4().hex
    try:
        connection.begin()
        with connection.cursor() as cursor:
            cursor.execute("SELECT next_number FROM document_counters WHERE document_type=%s AND document_year=%s FOR UPDATE", ("lapinhar", document_year))
            counter = cursor.fetchone()
            if counter is None:
                cursor.execute("INSERT INTO document_counters (document_type,document_year,next_number) VALUES (%s,%s,2)",
                               ("lapinhar", document_year))
                sequence_number = 1
            else:
                sequence_number = counter["next_number"]
                cursor.execute("UPDATE document_counters SET next_number=next_number+1 WHERE document_type=%s AND document_year=%s",
                               ("lapinhar", document_year))
            cursor.execute(
                """INSERT INTO document_number_reservations
                   (reservation_token, document_type, document_year, sequence_number, created_by)
                   VALUES (%s, 'lapinhar', %s, %s, %s)""",
                (token, document_year, sequence_number, session["user_id"]),
            )
        connection.commit()
        return {"reservation_token": token, "sequence_number": sequence_number, "document_year": document_year}
    except Exception:
        connection.rollback()
        raise


def accessible_lapinhar(report_id):
    report = fetch_one("SELECT * FROM reports WHERE id = %s AND report_type = 'lapinhar'", (report_id,))
    if report and (session.get("role") == "admin" or report["created_by"] == session.get("user_id")):
        return report
    return None


def report_sequence_number(report_number):
    match = re.match(r"^R\.LIH-(\d+)/", report_number or "")
    return int(match.group(1)) if match else 0


def report_number_year(report_number):
    match = re.search(r"/(\d{4})$", report_number or "")
    return int(match.group(1)) if match else 0


def compose_lapinhar_number(sequence_number, institution_code, issue_code, report_date):
    try:
        selected_date = date.fromisoformat(report_date)
    except (TypeError, ValueError):
        return ""
    if not sequence_number or issue_code not in ISSUE_CODES:
        return ""
    return f"R.LIH-{sequence_number}/{institution_code}/{issue_code}/{selected_date.month:02d}/{selected_date.year}"


def lapinhar_number_available(report_number, reservation_token="", report_id=None):
    duplicate = fetch_one("SELECT id FROM reports WHERE report_number=%s", (report_number,))
    if duplicate and duplicate["id"] != report_id:
        return False, "Nomor surat sudah digunakan laporan lain. Silakan reload nomor surat."
    if report_id:
        return True, "Nomor surat tersedia."
    reservation = fetch_one(
        """SELECT sequence_number,document_year FROM document_number_reservations
           WHERE reservation_token=%s AND document_type='lapinhar'
           AND created_by=%s AND status='reserved'""",
        (reservation_token, session["user_id"]),
    )
    if (not reservation or reservation["sequence_number"] != report_sequence_number(report_number)
            or reservation["document_year"] != report_number_year(report_number)):
        return False, "Reservasi nomor surat tidak berlaku. Silakan reload nomor surat."
    return True, "Nomor surat tersedia."


@app.post("/lapinhar/check-number")
@login_required
def check_lapinhar_number():
    payload = request.get_json(silent=True) or {}
    try:
        report_id = int(payload["report_id"]) if payload.get("report_id") else None
    except (TypeError, ValueError):
        report_id = None
    if report_id and accessible_lapinhar(report_id) is None:
        abort(404)
    available, message = lapinhar_number_available(
        str(payload.get("report_number", "")).strip(),
        str(payload.get("reservation_token", "")).strip(), report_id,
    )
    return jsonify(available=available, message=message), 200 if available else 409


@app.post("/lapinhar/reload-number")
@login_required
def reload_lapinhar_number():
    payload = request.get_json(silent=True) or {}
    reservation = reserve_lapinhar_number(payload.get("document_year"))
    return jsonify(**reservation)


def collect_attachments():
    attachments = []
    keys = sorted((key for key in request.files if key.startswith("attachment_images_")),
                  key=lambda key: int(key.rsplit("_", 1)[-1]))
    for key in keys:
        group_number = int(key.rsplit("_", 1)[-1])
        images = []
        for uploaded in request.files.getlist(key)[:2]:
            if not uploaded.filename or uploaded.mimetype not in ALLOWED_IMAGE_TYPES:
                continue
            content = uploaded.read()
            try:
                image = PILImage.open(BytesIO(content))
                image_width, image_height = image.size
                image.verify()
            except Exception:
                continue
            uploaded.seek(0)
            image_index = len(images) + 1
            scale = max(40, min(100, request.form.get(f"attachment_scale_{group_number}_{image_index}", 100, type=int)))
            images.append({"filename": secure_filename(uploaded.filename), "content": content,
                           "width": image_width, "height": image_height, "scale": scale})
        if images:
            attachments.append({
                "group": group_number,
                "title": "",
                "layout": request.form.get(f"attachment_layout_{group_number}", "auto"),
                "images": images,
            })
    return attachments


def attachment_label(count):
    """Build the memo label from attachment pages that actually contain photos."""
    if not count:
        return "-"
    number_words = {1: "satu", 2: "dua", 3: "tiga", 4: "empat", 5: "lima",
                    6: "enam", 7: "tujuh", 8: "delapan", 9: "sembilan", 10: "sepuluh"}
    return f"{count} ({number_words.get(count, str(count))}) Lembar"


def save_report_attachments(report_id, attachments):
    report_dir = UPLOAD_DIR / str(report_id)
    report_dir.mkdir(parents=True, exist_ok=True)
    with get_db().cursor() as cursor:
        for attachment in attachments:
            for image_index, image in enumerate(attachment["images"], 1):
                filename = f"{attachment['group']}_{image_index}_{uuid.uuid4().hex[:8]}_{image['filename']}"
                path = report_dir / filename
                path.write_bytes(image["content"])
                cursor.execute(
                    """INSERT INTO report_attachments
                    (report_id, attachment_group, image_path, caption, sort_order)
                    VALUES (%s, %s, %s, %s, %s)""",
                    (report_id, attachment["group"], str(path.relative_to(Path(__file__).parent)),
                     attachment["title"], image_index),
                )


def clone_report_attachments(source_report_id, target_report_id):
    """Copy every attachment file and row without modifying the LAPINHAR originals."""
    if fetch_one("SELECT id FROM report_attachments WHERE report_id=%s LIMIT 1", (target_report_id,)):
        return 0
    rows = fetch_all(
        "SELECT attachment_group,image_path,caption,sort_order FROM report_attachments "
        "WHERE report_id=%s ORDER BY attachment_group,sort_order,id", (source_report_id,),
    )
    source_dir = (UPLOAD_DIR / str(source_report_id)).resolve()
    target_dir = (UPLOAD_DIR / str(target_report_id)).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    copied = []
    try:
        for row in rows:
            source_path = (Path(__file__).resolve().parent / row["image_path"]).resolve()
            if not source_path.is_relative_to(source_dir) or not source_path.is_file():
                continue
            target_name = f"{row['attachment_group']}_{row['sort_order']}_{uuid.uuid4().hex[:8]}_{source_path.name}"
            target_path = target_dir / target_name
            shutil.copy2(source_path, target_path)
            copied.append((row, target_path))
        if copied:
            with get_db().cursor() as cursor:
                cursor.executemany(
                    "INSERT INTO report_attachments (report_id,attachment_group,image_path,caption,sort_order) "
                    "VALUES (%s,%s,%s,%s,%s)",
                    [(target_report_id, row["attachment_group"],
                      str(path.relative_to(Path(__file__).resolve().parent)), row["caption"], row["sort_order"])
                     for row, path in copied],
                )
        return len(copied)
    except Exception:
        for _row, path in copied:
            try:
                path.unlink()
            except OSError:
                pass
        raise


def remove_attachment_files(rows, report_id):
    """Remove known attachment files without allowing OneDrive locks to break report updates."""
    upload_root = UPLOAD_DIR.resolve()
    all_removed = True
    for row in rows:
        path = (Path(__file__).parent / row["image_path"]).resolve()
        if not path.is_relative_to(upload_root) or not path.exists():
            continue
        try:
            path.chmod(stat.S_IWRITE)
            path.unlink()
        except OSError:
            all_removed = False
            app.logger.warning("Berkas lampiran sedang terkunci dan belum dapat dihapus: %s", path)
    report_dir = (UPLOAD_DIR / str(report_id)).resolve()
    if report_dir.is_relative_to(upload_root) and report_dir.exists():
        try:
            report_dir.rmdir()
        except OSError:
            pass
    return all_removed


def existing_report_files(report_id):
    """Snapshot regular files in one report folder before replacement/deletion."""
    upload_root = UPLOAD_DIR.resolve()
    report_dir = (UPLOAD_DIR / str(report_id)).resolve()
    if not report_dir.is_relative_to(upload_root) or not report_dir.is_dir():
        return []
    return [
        {"image_path": str(path.relative_to(Path(__file__).parent))}
        for path in report_dir.iterdir()
        if path.is_file()
    ]


def attachment_layout(attachment):
    requested = attachment.get("layout")
    if requested in {"side", "stack"}:
        return requested
    images = attachment["images"]
    return "side" if len(images) == 2 and all(image["height"] > image["width"] for image in images) else "stack"


@app.route("/lapinhar/create", methods=["GET", "POST"])
@login_required
def create_lapinhar():
    if request.method == "POST":
        data = lapinhar_form_data()
        attachments = collect_attachments()
        data["attachment"] = attachment_label(len(attachments))
        reservation_token = request.form.get("number_reservation_token", "")
        reservation = fetch_one(
            """SELECT * FROM document_number_reservations
               WHERE reservation_token = %s AND document_type = 'lapinhar'
               AND created_by = %s AND status = 'reserved'""",
            (reservation_token, session["user_id"]),
        )
        data["report_number"] = compose_lapinhar_number(
            reservation["sequence_number"] if reservation else 0,
            data["institution_code"], data["issue_code"], data["report_date"],
        )
        number_available, number_message = lapinhar_number_available(
            data["report_number"], reservation_token
        )
        if (not data["report_number"] or not data["subject"] or not data["information"]
                or not data["category_id"] or not data["issue_code"] or not reservation):
            flash("Nomor reservasi, nomor permasalahan, kategori, perihal, dan informasi wajib diisi.", "error")
        elif not number_available:
            flash(number_message, "error")
        else:
            with get_db().cursor() as cursor:
                cursor.execute(
                    """INSERT INTO lapinhar_reports
                    (report_type, report_number, title, report_date, facts, source_name, analysis, recommendation,
                     recipient, sender_name, classification, category_id, issue_code, attachment, organization, creator_position,
                     creator_name, creator_rank_nip, auth_position, auth_name, auth_rank_nip,
                     information_spacing, sources_spacing, trends_spacing, suggestions_spacing,
                     status, created_by)
                    VALUES ('lapinhar', %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 'draft', %s)""",
                    (data["report_number"], data["subject"], data["report_date"] or None,
                     data["information"], data["sources"], data["trends"], data["suggestions"],
                     data["recipient"], data["sender"], data["classification"], data["category_id"],
                     data["issue_code"], data["attachment"],
                     data["organization"], data["creator_position"], data["creator_name"],
                     data["creator_rank_nip"], data["auth_position"], data["auth_name"], data["auth_rank_nip"],
                     data["information_spacing"] or "1.15", data["sources_spacing"] or "1.15",
                     data["trends_spacing"] or "1.15", data["suggestions_spacing"] or "1.15",
                    session["user_id"]),
                )
                report_id = cursor.lastrowid
                cursor.execute(
                    """UPDATE document_number_reservations SET status = 'used', report_id = %s,
                       used_at = CURRENT_TIMESTAMP WHERE id = %s AND status = 'reserved'""",
                    (report_id, reservation["id"]),
                )
            save_report_attachments(report_id, attachments)
            flash("LAPINHAR berhasil disimpan.", "success")
            after_save = request.form.get("after_save", "")
            if after_save in {"docx", "print", "pdf", "sipede"}:
                return redirect(url_for("edit_lapinhar", report_id=report_id, after_save=after_save))
            return redirect(url_for("lapinhar"))
    signatories = {row["position_code"]: row for row in fetch_all("SELECT * FROM signatories")}
    organization = fetch_one("SELECT * FROM organization_settings WHERE id = 1") or {
        "organization_name": "KEJAKSAAN NEGERI BULELENG", "institution_code": "N.1.11",
        "address": "", "phone": "", "website": ""
    }
    number_reservation = None
    if request.method == "POST":
        number_reservation = fetch_one(
            """SELECT reservation_token, sequence_number FROM document_number_reservations
               WHERE reservation_token = %s AND created_by = %s AND status = 'reserved'""",
            (request.form.get("number_reservation_token", ""), session["user_id"]),
        )
    if number_reservation is None:
        number_reservation = reserve_lapinhar_number()
    return render_template("lapinhar_editor.html", active="lapinhar", today=date.today().isoformat(),
                           signatories=signatories, organization=organization,
                           number_reservation=number_reservation,
                           fixed_recipient=f"YTH. KEPALA {organization['organization_name']}",
                           fixed_sender="KASI INTELIJEN", report=None, existing_sources=[], existing_attachments=[])


@app.route("/lapinhar/<int:report_id>/edit", methods=["GET", "POST"])
@login_required
def edit_lapinhar(report_id):
    report = accessible_lapinhar(report_id)
    if report is None:
        flash("Laporan tidak ditemukan atau tidak dapat Anda akses.", "error")
        return redirect(url_for("lapinhar"))
    sequence_number = report_sequence_number(report["report_number"])
    if request.method == "POST":
        data = lapinhar_form_data()
        submitted_sequence = report_sequence_number(request.form.get("report_number", ""))
        data["report_number"] = compose_lapinhar_number(
            submitted_sequence or sequence_number, data["institution_code"], data["issue_code"], data["report_date"],
        )
        attachments = collect_attachments()
        old_attachment_files = existing_report_files(report_id) if attachments else []
        old_attachment_rows = fetch_all(
            "SELECT id FROM report_attachments WHERE report_id=%s", (report_id,)
        ) if attachments else []
        number_available, number_message = lapinhar_number_available(
            data["report_number"], report_id=report_id
        )
        if (not data["report_number"] or not data["subject"] or not data["information"]
                or not data["category_id"] or not data["issue_code"]):
            flash("Nomor permasalahan, kategori, perihal, dan informasi wajib diisi.", "error")
        elif not number_available:
            flash(number_message, "error")
        else:
            with get_db().cursor() as cursor:
                cursor.execute(
                    """UPDATE lapinhar_reports SET report_number=%s, title=%s, report_date=%s, facts=%s,
                       source_name=%s, analysis=%s, recommendation=%s, recipient=%s, sender_name=%s,
                       classification='RAHASIA', category_id=%s, issue_code=%s, attachment=%s,
                       organization=%s, creator_position=%s, creator_name=%s, creator_rank_nip=%s,
                       auth_position=%s, auth_name=%s, auth_rank_nip=%s, information_spacing='1.5',
                       sources_spacing='1.5', trends_spacing='1.5', suggestions_spacing='1.5'
                       WHERE id=%s""",
                    (data["report_number"], data["subject"], data["report_date"] or None,
                     data["information"], data["sources"], data["trends"], data["suggestions"],
                     data["recipient"], data["sender"], data["category_id"], data["issue_code"],
                     attachment_label(len(attachments)) if attachments else report["attachment"],
                     data["organization"], data["creator_position"], data["creator_name"],
                     data["creator_rank_nip"], data["auth_position"], data["auth_name"],
                     data["auth_rank_nip"], report_id),
                )
                reservation_token = request.form.get("number_reservation_token", "")
                if reservation_token:
                    cursor.execute(
                        "UPDATE document_number_reservations SET status='used',report_id=%s,used_at=CURRENT_TIMESTAMP "
                        "WHERE reservation_token=%s AND document_type='lapinhar' AND created_by=%s AND status='reserved'",
                        (report_id, reservation_token, session["user_id"]),
                    )
            if attachments:
                save_report_attachments(report_id, attachments)
                if old_attachment_rows:
                    with get_db().cursor() as cursor:
                        cursor.executemany(
                            "DELETE FROM report_attachments WHERE id=%s",
                            [(row["id"],) for row in old_attachment_rows],
                        )
                remove_attachment_files(old_attachment_files, report_id)
            flash("LAPINHAR berhasil diperbarui.", "success")
            after_save = request.form.get("after_save", "")
            if after_save in {"docx", "print", "pdf", "sipede"}:
                return redirect(url_for("edit_lapinhar", report_id=report_id, after_save=after_save))
            return redirect(url_for("lapinhar"))

    signatories = {row["position_code"]: row for row in fetch_all("SELECT * FROM signatories")}
    organization = fetch_one("SELECT * FROM organization_settings WHERE id=1")
    existing_sources = []
    for block in rich_text_blocks(report["source_name"] or ""):
        if block["text"].strip().lower() != "intelijen kejaksaan negeri buleleng":
            existing_sources.append(block["text"].strip())
    existing_attachments = fetch_all(
        "SELECT attachment_group, image_path, sort_order FROM report_attachments WHERE report_id=%s ORDER BY attachment_group, sort_order",
        (report_id,),
    )
    existing_attachment_groups = []
    for row in existing_attachments:
        group = next((item for item in existing_attachment_groups
                      if item["group"] == row["attachment_group"]), None)
        if group is None:
            group = {"group": row["attachment_group"], "images": []}
            existing_attachment_groups.append(group)
        group["images"].append({
            "url": url_for("lapinhar_attachment_file", report_id=report_id,
                           filename=Path(row["image_path"]).name),
            "filename": Path(row["image_path"]).name,
        })
    selected_signer = "kasi_intel"
    for code in ("kasubsi_1", "kasubsi_2"):
        signer = signatories.get(code)
        if not signer:
            continue
        report_nip = re.sub(r"\D", "", report.get("creator_rank_nip") or "")
        signer_nip = re.sub(r"\D", "", signer.get("rank_nip") or "")
        same_position = signer["position_name"] == report["creator_position"]
        same_name = bool(report.get("creator_name") and signer.get("full_name")
                         and signer["full_name"].casefold() == report["creator_name"].casefold())
        same_nip = bool(report_nip and signer_nip and report_nip == signer_nip)
        if same_position or same_name or same_nip:
            selected_signer = code
            break
    number_reservation = {"reservation_token": "", "sequence_number": sequence_number}
    return render_template(
        "lapinhar_editor.html", active="lapinhar", today=report["report_date"].isoformat() if report["report_date"] else date.today().isoformat(),
        signatories=signatories, organization=organization, number_reservation=number_reservation,
        fixed_recipient=report["recipient"], fixed_sender=report["sender_name"], report=report,
        existing_sources=existing_sources, existing_attachments=existing_attachments,
        existing_attachment_groups=existing_attachment_groups,
        selected_signer=selected_signer,
    )


@app.get("/lapinhar/<int:report_id>/attachments/<path:filename>")
@login_required
def lapinhar_attachment_file(report_id, filename):
    if accessible_lapinhar(report_id) is None or Path(filename).name != filename:
        abort(404)
    row = next((item for item in fetch_all(
        "SELECT image_path FROM report_attachments WHERE report_id=%s", (report_id,)
    ) if Path(item["image_path"]).name == filename), None)
    if row is None or Path(row["image_path"]).name != filename:
        abort(404)
    return send_from_directory(UPLOAD_DIR / str(report_id), filename)


@app.get("/lapinhar/<int:report_id>/attachments/<path:filename>/download")
@login_required
def download_lapinhar_attachment(report_id, filename):
    if accessible_lapinhar(report_id) is None or Path(filename).name != filename:
        abort(404)
    row = next((item for item in fetch_all(
        """SELECT attachment_group, sort_order, image_path FROM report_attachments
           WHERE report_id=%s""", (report_id,)
    ) if Path(item["image_path"]).name == filename), None)
    if row is None or Path(row["image_path"]).name != filename:
        abort(404)
    extension = Path(filename).suffix.lower() or ".jpg"
    download_name = f"Lampiran-{row['attachment_group']}-Foto-{row['sort_order']}{extension}"
    return send_from_directory(UPLOAD_DIR / str(report_id), filename,
                               as_attachment=True, download_name=download_name)


@app.get("/lapinhar/<int:report_id>/attachments/download-all")
@login_required
def download_all_lapinhar_attachments(report_id):
    report = accessible_lapinhar(report_id)
    if report is None:
        abort(404)
    rows = fetch_all(
        """SELECT attachment_group, sort_order, image_path FROM report_attachments
           WHERE report_id=%s ORDER BY attachment_group, sort_order, id""",
        (report_id,),
    )
    if not rows:
        abort(404)
    archive = BytesIO()
    report_dir = (UPLOAD_DIR / str(report_id)).resolve()
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as bundle:
        for row in rows:
            source = (Path(__file__).resolve().parent / row["image_path"]).resolve()
            if not source.is_relative_to(report_dir) or not source.is_file():
                continue
            extension = source.suffix.lower() or ".jpg"
            archive_name = f"Lampiran-{row['attachment_group']}-Foto-{row['sort_order']}{extension}"
            bundle.write(source, archive_name)
    if not archive.getvalue():
        abort(404)
    archive.seek(0)
    sequence = report_sequence_number(report["report_number"]) or report_id
    return send_file(archive, as_attachment=True,
                     download_name=f"Lampiran-RLIH-{sequence}.zip",
                     mimetype="application/zip")


@app.post("/lapinhar/<int:report_id>/delete")
@login_required
def delete_lapinhar(report_id):
    report = accessible_lapinhar(report_id)
    if report is None:
        flash("Laporan tidak ditemukan atau tidak dapat Anda akses.", "error")
        return redirect(url_for("lapinhar"))
    attachment_files = existing_report_files(report_id)
    with get_db().cursor() as cursor:
        cursor.execute("DELETE FROM report_attachments WHERE report_id=%s", (report_id,))
        cursor.execute("DELETE FROM lapinhar_reports WHERE id=%s", (report_id,))
    files_removed = remove_attachment_files(attachment_files, report_id)
    flash(f"LAPINHAR {report['report_number']} berhasil dihapus permanen.", "success")
    if not files_removed:
        flash("Sebagian berkas foto sedang dikunci OneDrive dan belum dapat dibersihkan, tetapi data laporan sudah terhapus.", "warning")
    return redirect(url_for("lapinhar"))


def lapinsus_editor_context(report=None, number_reservation=None):
    signatories = {row["position_code"]: row for row in fetch_all("SELECT * FROM signatories")}
    organization = fetch_one("SELECT * FROM organization_settings WHERE id=1") or {
        "organization_name": "KEJAKSAAN NEGERI BULELENG", "institution_code": "N.1.11",
        "address": "", "phone": "", "website": "",
    }
    existing_sources, existing_attachments, groups = [], [], []
    selected_signer = "kasubsi_1"
    if report:
        for block in rich_text_blocks(report["source_name"] or ""):
            if block["text"].strip().lower() != "intelijen kejaksaan negeri buleleng":
                existing_sources.append(block["text"].strip())
        existing_attachments = fetch_all(
            "SELECT attachment_group,image_path,sort_order FROM report_attachments "
            "WHERE report_id=%s ORDER BY attachment_group,sort_order", (report["id"],),
        )
        for row in existing_attachments:
            group = next((item for item in groups if item["group"] == row["attachment_group"]), None)
            if group is None:
                group = {"group": row["attachment_group"], "images": []}
                groups.append(group)
            group["images"].append({
                "url": url_for("lapinsus_attachment_file", report_id=report["id"],
                               filename=Path(row["image_path"]).name),
                "filename": Path(row["image_path"]).name,
            })
        for code in ("kasi_intel", "kasubsi_1", "kasubsi_2"):
            signer = signatories.get(code)
            if signer and (signer.get("position_name") == report.get("creator_position")
                           or (signer.get("full_name") or "").casefold() == (report.get("creator_name") or "").casefold()):
                selected_signer = code
                break
    return dict(
        active="lapinsus", document_kind="lapinsus", document_label="LAPINSUS",
        report_heading="LAPORAN INFORMASI KHUSUS", report_code="L. IN.2",
        list_url=url_for("lapinsus"), check_number_url=url_for("check_lapinsus_number"),
        reload_number_url=url_for("reload_lapinsus_number"),
        today=(report["report_date"].isoformat() if report and report["report_date"] else date.today().isoformat()),
        signatories=signatories, organization=organization, number_reservation=number_reservation,
        fixed_recipient=(report["recipient"] if report else f"YTH. KEPALA {organization['organization_name']}"),
        fixed_sender=(report["sender_name"] if report else "KASI INTELIJEN"), report=report,
        existing_sources=existing_sources, existing_attachments=existing_attachments,
        existing_attachment_groups=groups, selected_signer=selected_signer,
    )


def insert_lapinsus(data, reservation, attachments):
    with get_db().cursor() as cursor:
        cursor.execute(
            """INSERT INTO lapinsus_reports
            (report_type,report_number,title,report_date,facts,source_name,analysis,recommendation,
             recipient,sender_name,classification,category_id,issue_code,attachment,organization,
             creator_position,creator_name,creator_rank_nip,auth_position,auth_name,auth_rank_nip,
             information_spacing,sources_spacing,trends_spacing,suggestions_spacing,status,created_by)
            VALUES ('lapinsus',%s,%s,%s,%s,%s,%s,%s,%s,%s,'RAHASIA',%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,
                    '1.5','1.5','1.5','1.5','draft',%s)""",
            (data["report_number"], data["subject"], data["report_date"] or None,
             data["information"], data["sources"], data["trends"], data["suggestions"],
             data["recipient"], data["sender"], data["category_id"], data["issue_code"],
             attachment_label(len(attachments)), data["organization"], data["creator_position"],
             data["creator_name"], data["creator_rank_nip"], data["auth_position"], data["auth_name"],
             data["auth_rank_nip"], session["user_id"]),
        )
        report_id = cursor.lastrowid
        cursor.execute("UPDATE document_number_reservations SET status='used',report_id=%s,used_at=CURRENT_TIMESTAMP "
                       "WHERE id=%s AND status='reserved'", (report_id, reservation["id"]))
    save_report_attachments(report_id, attachments)
    return report_id


@app.post("/lapinhar/<int:report_id>/lapinsuskan")
@login_required
def convert_lapinhar_to_lapinsus(report_id):
    source = accessible_lapinhar(report_id)
    if source is None:
        flash("LAPINHAR tidak ditemukan atau tidak dapat Anda akses.", "error")
        return redirect(url_for("lapinhar"))
    existing = fetch_one("SELECT id FROM lapinsus_reports WHERE source_lapinhar_id=%s", (report_id,))
    if existing:
        flash("LAPINHAR tersebut sudah pernah dijadikan LAPINSUS.", "warning")
        return redirect(url_for("edit_lapinsus", report_id=existing["id"]))
    organization = fetch_one("SELECT institution_code FROM organization_settings WHERE id=1") or {"institution_code": "N.1.11"}
    reservation = reserve_lapinsus_number(source["report_date"].year if source.get("report_date") else date.today().year)
    report_number = compose_lapinsus_number(
        reservation["sequence_number"], organization["institution_code"],
        source.get("issue_code") or "", source["report_date"].isoformat() if source.get("report_date") else "",
    )
    if not report_number:
        flash("Tanggal atau nomor permasalahan LAPINHAR belum lengkap.", "error")
        return redirect(url_for("edit_lapinhar", report_id=report_id))
    with get_db().cursor() as cursor:
        cursor.execute(
            """INSERT INTO lapinsus_reports
            (report_type,report_number,title,report_date,source_name,facts,analysis,recommendation,
             recipient,sender_name,classification,category_id,issue_code,attachment,organization,
             creator_position,creator_name,creator_rank_nip,auth_position,auth_name,auth_rank_nip,
             information_spacing,sources_spacing,trends_spacing,suggestions_spacing,status,created_by,source_lapinhar_id)
            VALUES ('lapinsus',%s,%s,%s,%s,%s,%s,%s,%s,%s,'RAHASIA',%s,%s,'-',%s,%s,%s,%s,%s,%s,%s,
                    '1.5','1.5','1.5','1.5','draft',%s,%s)""",
            (report_number, source["title"], source["report_date"], source["source_name"], source["facts"],
             source["analysis"], source["recommendation"], source["recipient"], source["sender_name"],
             source["category_id"], source["issue_code"], source["organization"], source["creator_position"],
             source["creator_name"], source["creator_rank_nip"], source["auth_position"], source["auth_name"],
             source["auth_rank_nip"], session["user_id"], report_id),
        )
        lapinsus_id = cursor.lastrowid
        cursor.execute("UPDATE lapinhar_reports SET lapinsus_status='sudah' WHERE id=%s", (report_id,))
        cursor.execute("UPDATE document_number_reservations SET status='used',report_id=%s,used_at=CURRENT_TIMESTAMP "
                       "WHERE reservation_token=%s", (lapinsus_id, reservation["reservation_token"]))
    attachment_count = clone_report_attachments(report_id, lapinsus_id)
    if attachment_count:
        with get_db().cursor() as cursor:
            cursor.execute("UPDATE lapinsus_reports SET attachment=%s WHERE id=%s",
                           (attachment_label(len({row['attachment_group'] for row in fetch_all('SELECT attachment_group FROM report_attachments WHERE report_id=%s', (lapinsus_id,))})), lapinsus_id))
    flash(f"Draft LAPINSUS berhasil dibuat dari LAPINHAR. {attachment_count} foto lampiran ikut disalin.", "success")
    return redirect(url_for("edit_lapinsus", report_id=lapinsus_id))


@app.route("/lapinsus/create", methods=["GET", "POST"])
@login_required
def create_lapinsus():
    reservation = None
    if request.method == "POST":
        data = lapinhar_form_data()
        attachments = collect_attachments()
        token = request.form.get("number_reservation_token", "")
        reservation = fetch_one(
            "SELECT * FROM document_number_reservations WHERE reservation_token=%s "
            "AND document_type='lapinsus' AND created_by=%s AND status='reserved'",
            (token, session["user_id"]),
        )
        data["report_number"] = compose_lapinsus_number(
            reservation["sequence_number"] if reservation else 0, data["institution_code"],
            data["issue_code"], data["report_date"],
        )
        available, message = lapinsus_number_available(data["report_number"], token)
        if not all((data["report_number"], data["subject"], data["information"],
                    data["category_id"], data["issue_code"], reservation)):
            flash("Nomor reservasi, nomor permasalahan, kategori, perihal, dan informasi wajib diisi.", "error")
        elif not available:
            flash(message, "error")
        else:
            report_id = insert_lapinsus(data, reservation, attachments)
            flash("LAPINSUS berhasil disimpan.", "success")
            after_save = request.form.get("after_save", "")
            if after_save in {"docx", "print", "pdf", "sipede"}:
                return redirect(url_for("edit_lapinsus", report_id=report_id, after_save=after_save))
            return redirect(url_for("lapinsus"))
    if reservation is None:
        reservation = reserve_lapinsus_number()
    return render_template("lapinhar_editor.html", **lapinsus_editor_context(number_reservation=reservation))


@app.route("/lapinsus/<int:report_id>/edit", methods=["GET", "POST"])
@login_required
def edit_lapinsus(report_id):
    report = accessible_lapinsus(report_id)
    if report is None:
        flash("LAPINSUS tidak ditemukan atau tidak dapat Anda akses.", "error")
        return redirect(url_for("lapinsus"))
    sequence = lapinsus_sequence_number(report["report_number"])
    if request.method == "POST":
        data = lapinhar_form_data()
        submitted_sequence = lapinsus_sequence_number(request.form.get("report_number", ""))
        data["report_number"] = compose_lapinsus_number(
            submitted_sequence or sequence, data["institution_code"], data["issue_code"], data["report_date"])
        attachments = collect_attachments()
        old_files = existing_report_files(report_id) if attachments else []
        old_rows = fetch_all("SELECT id FROM report_attachments WHERE report_id=%s", (report_id,)) if attachments else []
        available, message = lapinsus_number_available(data["report_number"], report_id=report_id)
        if not all((data["report_number"], data["subject"], data["information"], data["category_id"], data["issue_code"])):
            flash("Nomor permasalahan, kategori, perihal, dan informasi wajib diisi.", "error")
        elif not available:
            flash(message, "error")
        else:
            with get_db().cursor() as cursor:
                cursor.execute(
                    """UPDATE lapinsus_reports SET report_number=%s,title=%s,report_date=%s,facts=%s,source_name=%s,
                    analysis=%s,recommendation=%s,recipient=%s,sender_name=%s,classification='RAHASIA',
                    category_id=%s,issue_code=%s,attachment=%s,organization=%s,creator_position=%s,
                    creator_name=%s,creator_rank_nip=%s,auth_position=%s,auth_name=%s,auth_rank_nip=%s,
                    information_spacing='1.5',sources_spacing='1.5',trends_spacing='1.5',suggestions_spacing='1.5'
                    WHERE id=%s""",
                    (data["report_number"], data["subject"], data["report_date"] or None, data["information"],
                     data["sources"], data["trends"], data["suggestions"], data["recipient"], data["sender"],
                     data["category_id"], data["issue_code"], attachment_label(len(attachments)) if attachments else report["attachment"],
                     data["organization"], data["creator_position"], data["creator_name"], data["creator_rank_nip"],
                     data["auth_position"], data["auth_name"], data["auth_rank_nip"], report_id),
                )
                reservation_token = request.form.get("number_reservation_token", "")
                if reservation_token:
                    cursor.execute(
                        "UPDATE document_number_reservations SET status='used',report_id=%s,used_at=CURRENT_TIMESTAMP "
                        "WHERE reservation_token=%s AND document_type='lapinsus' AND created_by=%s AND status='reserved'",
                        (report_id, reservation_token, session["user_id"]),
                    )
            if attachments:
                save_report_attachments(report_id, attachments)
                if old_rows:
                    with get_db().cursor() as cursor:
                        cursor.executemany("DELETE FROM report_attachments WHERE id=%s", [(row["id"],) for row in old_rows])
                remove_attachment_files(old_files, report_id)
            flash("LAPINSUS berhasil diperbarui.", "success")
            after_save = request.form.get("after_save", "")
            if after_save in {"docx", "print", "pdf", "sipede"}:
                return redirect(url_for("edit_lapinsus", report_id=report_id, after_save=after_save))
            return redirect(url_for("lapinsus"))
        report = accessible_lapinsus(report_id)
    reservation = {"reservation_token": "", "sequence_number": sequence}
    return render_template("lapinhar_editor.html", **lapinsus_editor_context(report, reservation))


@app.get("/lapinsus/<int:report_id>/attachments/<path:filename>")
@login_required
def lapinsus_attachment_file(report_id, filename):
    if accessible_lapinsus(report_id) is None or Path(filename).name != filename:
        abort(404)
    row = next((item for item in fetch_all("SELECT image_path FROM report_attachments WHERE report_id=%s", (report_id,))
                if Path(item["image_path"]).name == filename), None)
    if row is None:
        abort(404)
    return send_from_directory(UPLOAD_DIR / str(report_id), filename)


@app.post("/lapinsus/<int:report_id>/delete")
@login_required
def delete_lapinsus(report_id):
    report = accessible_lapinsus(report_id)
    if report is None:
        flash("LAPINSUS tidak ditemukan atau tidak dapat Anda akses.", "error")
        return redirect(url_for("lapinsus"))
    files = existing_report_files(report_id)
    with get_db().cursor() as cursor:
        cursor.execute("DELETE FROM report_attachments WHERE report_id=%s", (report_id,))
        cursor.execute("DELETE FROM lapinsus_reports WHERE id=%s", (report_id,))
    remove_attachment_files(files, report_id)
    flash(f"LAPINSUS {report['report_number']} berhasil dihapus.", "success")
    return redirect(url_for("lapinsus"))


@app.post("/lapinsus/<int:report_id>/upload-sipede")
@login_required
def upload_lapinsus_sipede(report_id):
    report = accessible_lapinsus(report_id)
    if report is None:
        return jsonify(message="LAPINSUS tidak ditemukan atau bukan milik Anda."), 403
    setting = fetch_one("""SELECT sipede_username,sipede_password_encrypted,
                        session_data_encrypted,connected_at
                        FROM sipede_user_settings WHERE user_id=%s""",
                        (session["user_id"],))
    if not setting:
        return jsonify(message="Konfigurasi Sipede belum tersedia. Isi username dan password pada Konfigurasi Integrasi."), 409
    if report.get("sipede_status") == "sudah":
        return jsonify(message="LAPINSUS sudah diupload ke Sipede."), 200
    if not setting.get("session_data_encrypted") or not setting.get("connected_at"):
        return jsonify(message="Login SIPede diperlukan.", requires_sipede_login=True), 401
    session_data = get_sipede_session_data(session["user_id"])
    if not session_data:
        return jsonify(message="Login SIPede diperlukan.", requires_sipede_login=True), 401
    try:
        http_session = requests.Session()
        for cookie in session_data.get("cookies", []):
            http_session.cookies.set(
                cookie.get("name", ""), cookie.get("value", ""),
                domain=cookie.get("domain") or "sipede.kejaksaan.go.id",
                path=cookie.get("path") or "/",
            )
        response = http_session.get(
            SIPEDE_SURATKELUAR_CREATE_URL,
            headers={"Accept": "text/html,application/xhtml+xml", "Referer": f"{SIPEDE_BASE_URL}/"},
            timeout=60,
        )
        create_context = sipede_create_context_from_html(response.text, response.url)
        if "/login" in urlparse(response.url).path.lower() or not create_context:
            with get_db().cursor() as cursor:
                cursor.execute(
                    "UPDATE sipede_user_settings SET connected_at=NULL WHERE user_id=%s",
                    (session["user_id"],),
                )
            return jsonify(message="Sesi SIPede kedaluwarsa.", requires_sipede_login=True), 401
        kajari = fetch_one("SELECT full_name,position_name FROM signatories WHERE position_code='kajari'")
        kajari_name = kajari.get("full_name", "") if kajari else ""
        signatory_id = sipede_signatory_id(create_context, kajari_name)
        if not signatory_id:
            return jsonify(
                message=(f"Penandatangan {kajari_name or 'Kepala Kejaksaan Negeri'} tidak ditemukan "
                         "pada daftar SIPede. Periksa Konfigurasi Tanda Tangan."),
            ), 422
        issue_code = sipede_issue_code_value(create_context, report.get("issue_code"))
        if not issue_code:
            return jsonify(
                message=(f"Kode masalah {report.get('issue_code') or 'LAPINSUS'} tidak ditemukan "
                         "pada daftar Kode Masalah SIPede."),
            ), 422
        create_context["selected_kode_masalah"] = issue_code
        create_context["selected_penandatangan"] = str(signatory_id)

        ajax_headers = {
            "Accept": "*/*",
            "Referer": SIPEDE_SURATKELUAR_CREATE_URL,
            "X-Requested-With": "XMLHttpRequest",
        }
        csrf_token = create_context.get("csrf_token", "")
        if not csrf_token:
            return jsonify(message="Token form Surat Keluar SIPede tidak ditemukan."), 422
        temporary_number_key = f"sipede_upload_number_{report_id}"
        submitted_sipede_number = request.form.get("sipede_number", "").strip() if request.files.get("document") else ""
        temporary_sipede_number = str(session.get(temporary_number_key) or "").strip()
        if submitted_sipede_number:
            if (not temporary_sipede_number or submitted_sipede_number != temporary_sipede_number
                    or len(submitted_sipede_number) > 150
                    or re.search(r"[\r\n]", submitted_sipede_number)):
                return jsonify(message="Nomor SIPede sementara tidak valid. Ulangi proses upload."), 409
            sipede_number = temporary_sipede_number
        else:
            sipede_number = str(report.get("sipede_number") or "").strip()
        if not sipede_number or sipede_number == "-":
            auto_number_response = http_session.get(
                f"{SIPEDE_BASE_URL}/suratkeluar/check-auto-number",
                params={"surat": "23"}, headers=ajax_headers, timeout=30,
            )
            auto_number_response.raise_for_status()
            number_response = http_session.post(
                f"{SIPEDE_BASE_URL}/getnosurat/cekno",
                headers={**ajax_headers, "X-CSRF-TOKEN": csrf_token},
                data={
                    "_token": csrf_token, "jenis_surat": "23",
                    "tanggal": str(report.get("report_date") or date.today().isoformat()),
                    "penandatangan": str(signatory_id), "sifat_surat": "R",
                },
                timeout=30,
            )
            number_response.raise_for_status()
            sipede_number = number_response.text.strip().strip('"')
            if not sipede_number or "<html" in sipede_number.lower():
                return jsonify(message="Nomor sequence surat tidak diterima dari SIPede."), 502
            session[temporary_number_key] = sipede_number
            session.modified = True

        destinations_response = http_session.get(
            f"{SIPEDE_BASE_URL}/suratkeluar/get-master-tujuan-disposisi",
            headers={**ajax_headers, "X-CSRF-TOKEN": csrf_token},
            timeout=30,
        )
        destinations_response.raise_for_status()
        try:
            destinations_payload = destinations_response.json()
        except ValueError:
            return jsonify(message="Daftar tujuan SIPede tidak dapat dibaca."), 502
        destination_rows = destinations_payload.get("data", []) if isinstance(destinations_payload, dict) else []
        destinations = []
        for item in destination_rows:
            users = item.get("user") or []
            user_name = users[0].get("nama", "NO USER") if users and isinstance(users[0], dict) else "NO USER"
            destinations.append({
                "id": str(item.get("id_tujuan_penerusan", "")),
                "position": str(item.get("nama_jabatan", "")).strip(),
                "user": str(user_name).strip() or "NO USER",
            })
        destinations = [item for item in destinations if item["id"] and item["position"]]
        if not destinations:
            return jsonify(message="Daftar tujuan disposisi SIPede masih kosong."), 422

        selected_destinations = []
        if request.files.get("document"):
            try:
                selected_destinations = json.loads(request.form.get("destinations", "[]"))
            except (TypeError, ValueError):
                selected_destinations = []
            allowed_destination_ids = {item["id"] for item in destinations}
            selected_destinations = [str(item) for item in selected_destinations
                                     if str(item) in allowed_destination_ids]
            if not selected_destinations:
                return jsonify(message="Pilih minimal satu tujuan SIPede."), 422
            document = request.files["document"]
            document_bytes = document.read()
            if not document_bytes or len(document_bytes) > 18_000_000:
                return jsonify(message="Dokumen PDF LAPINSUS tidak valid atau terlalu besar."), 422
            organization = fetch_one("SELECT * FROM organization_settings WHERE id=1") or {
                "organization_name": "Kejaksaan Negeri Buleleng"
            }
            submit_data = {
                "_token": csrf_token,
                "suratmasuk": "",
                "nomor": sipede_number,
                "tanggal": str(report.get("report_date") or date.today().isoformat()),
                "jenis": "23",
                "sifat": "R",
                "kode_masalah": issue_code,
                "tujuan": "Yth.\nKepala Kejaksaan Tinggi Bali\nDi - Denpasar",
                "dari": str((kajari or {}).get("position_name") or
                            f"Kepala {organization.get('organization_name', 'Kejaksaan Negeri Buleleng')}"),
                "hal": str(report.get("title") or report.get("subject") or "LAPINSUS"),
                "penandatangan": str(signatory_id),
                "idSurat": "125",
                "idSuratMasuk": "",
                "tujuan_surat": ",".join(selected_destinations),
                "submitModal": "ok",
            }
            submit_response = http_session.post(
                create_context.get("form_action") or f"{SIPEDE_BASE_URL}/suratkeluar",
                headers={
                    "Accept": ("text/html,application/xhtml+xml,application/xml;q=0.9,"
                               "image/avif,image/webp,image/apng,*/*;q=0.8"),
                    "Accept-Language": "id,en-US;q=0.9,en;q=0.8",
                    "Cache-Control": "max-age=0",
                    "Origin": SIPEDE_BASE_URL,
                    "Referer": SIPEDE_SURATKELUAR_CREATE_URL,
                    "Upgrade-Insecure-Requests": "1",
                    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                                   "Chrome/150.0.0.0 Safari/537.36"),
                },
                data=submit_data,
                files={"upload": (document.filename or "LAPINSUS.pdf", document_bytes, "application/pdf")},
                timeout=90,
                allow_redirects=True,
            )
            submit_response.raise_for_status()
            submit_soup = BeautifulSoup(submit_response.text, "html.parser")
            returned_form = submit_soup.select_one('#formCreateEdit, form[action$="/suratkeluar"]')
            error_node = submit_soup.select_one('.alert-danger, .invalid-feedback, .text-danger')
            if returned_form and urlparse(submit_response.url).path.rstrip("/").endswith("/create"):
                error_message = error_node.get_text(" ", strip=True) if error_node else "SIPede belum menerima surat."
                return jsonify(message=error_message), 422
            with get_db().cursor() as cursor:
                cursor.execute(
                    "UPDATE lapinsus_reports SET sipede_number=%s,sipede_status='sudah' WHERE id=%s",
                    (sipede_number, report_id),
                )
            session.pop(temporary_number_key, None)
            session.modified = True
            save_refreshed_sipede_session(session["user_id"], session_data, http_session, create_context)
            flash("LAPINSUS berhasil dikirim ke SIPede.", "success")
            return jsonify(
                message="LAPINSUS berhasil dikirim ke SIPede.",
                uploaded=True,
                sipede_number=sipede_number,
                redirect_url=url_for("lapinsus"),
            )

        save_refreshed_sipede_session(session["user_id"], session_data, http_session, create_context)
        return jsonify(
            message="Nomor sequence surat SIPede berhasil diperoleh.",
            sipede_connected=True,
            sipede_number=sipede_number,
            signatory_id=str(signatory_id),
            signatory_name=kajari_name,
            kode_masalah=issue_code,
            destinations=destinations,
        ), 202
    except requests.RequestException as exc:
        app.logger.warning("Form Surat Keluar SIPede gagal dibuka: %s", exc)
        return jsonify(message="SIPede tidak dapat dihubungi. Silakan coba lagi."), 502


def rich_text_blocks(content):
    if "<" not in content:
        return [{"text": line.strip(), "prefix": "", "align": "justify", "indent": 0,
                 "runs": [{"text": line.strip(), "bold": False, "italic": False, "underline": False}]}
                for line in content.splitlines() if line.strip()]
    soup = BeautifulSoup(content, "html.parser")
    blocks = []
    ordered_counters = {}

    def collect_runs(node, bold=False, italic=False, underline=False):
        runs = []
        for child in node.children:
            if isinstance(child, NavigableString):
                if str(child):
                    runs.append({"text": str(child), "bold": bold, "italic": italic, "underline": underline})
            elif isinstance(child, Tag) and child.name not in {"ol", "ul"}:
                runs.extend(collect_runs(child, bold or child.name in {"b", "strong"},
                                         italic or child.name in {"i", "em"},
                                         underline or child.name == "u"))
        return runs

    for element in soup.find_all(["p", "li"]):
        runs = collect_runs(element)
        text = "".join(run["text"] for run in runs).strip()
        if not text:
            continue
        classes = element.get("class", [])
        align = next((value.replace("ql-align-", "") for value in classes if value.startswith("ql-align-")), "justify")
        inline_style = element.get("style", "").replace(" ", "").lower()
        if "text-align:" in inline_style:
            align = inline_style.split("text-align:", 1)[1].split(";", 1)[0]
        indent_class = next((value for value in classes if value.startswith("ql-indent-")), "")
        indent = int(indent_class.rsplit("-", 1)[-1]) if indent_class else 0
        if "margin-left:" in inline_style:
            margin = inline_style.split("margin-left:", 1)[1].split(";", 1)[0]
            try:
                indent = max(indent, round(float(margin.replace("px", "").replace("em", "")) / (40 if "px" in margin else 1)))
            except ValueError:
                pass
        prefix = ""
        if element.name == "li":
            parent = element.parent
            list_type = element.get("data-list")
            if list_type == "ordered" or (not list_type and parent and parent.name == "ol"):
                key = id(parent)
                ordered_counters[key] = ordered_counters.get(key, 0) + 1
                prefix = f"{ordered_counters[key]}. "
            else:
                prefix = "• "
        blocks.append({"text": text, "prefix": prefix, "align": align, "indent": indent, "runs": runs})
    return blocks


def add_docx_section(document, title, content, spacing="1.15"):
    heading = document.add_paragraph()
    heading.paragraph_format.space_before = Pt(0)
    heading.paragraph_format.space_after = Pt(0)
    heading.add_run(title).bold = True
    alignments = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT,
                  "justify": WD_ALIGN_PARAGRAPH.JUSTIFY, "left": WD_ALIGN_PARAGRAPH.LEFT}
    for block in rich_text_blocks(content):
        paragraph = document.add_paragraph(style=None)
        paragraph.paragraph_format.left_indent = Cm(.7 + block["indent"] * .7)
        paragraph.paragraph_format.first_line_indent = Cm(-.7 if block["prefix"] else 0)
        paragraph.paragraph_format.alignment = alignments.get(block["align"], WD_ALIGN_PARAGRAPH.JUSTIFY)
        paragraph.paragraph_format.line_spacing = float(spacing or 1.15)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        if block["prefix"]: paragraph.add_run(block["prefix"]).bold = True
        for item in block["runs"]:
            run = paragraph.add_run(item["text"])
            run.bold, run.italic, run.underline = item["bold"], item["italic"], item["underline"]
            run.font.name, run.font.size = "Times New Roman", Pt(12)


def add_docx_label(document, label, value):
    table = document.add_table(rows=1, cols=3)
    table.autofit = False
    table.columns[0].width, table.columns[1].width, table.columns[2].width = Cm(3.2), Cm(.5), Cm(13)
    cells = table.rows[0].cells
    cells[0].text, cells[1].text, cells[2].text = label, ":", value
    for cell in cells:
        cell.vertical_alignment = 1
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def lapinhar_export_basename(report_number, report_date, subject):
    """Return a safe RLIH/RLIK export filename."""
    number_text = str(report_number or "")
    is_lapinsus = number_text.startswith("R.LIK-")
    sequence = lapinsus_sequence_number(number_text) if is_lapinsus else report_sequence_number(number_text)
    prefix = "RLIK" if is_lapinsus else "RLIH"
    try:
        selected_date = date.fromisoformat(str(report_date or ""))
        months = ("Januari", "Februari", "Maret", "April", "Mei", "Juni",
                  "Juli", "Agustus", "September", "Oktober", "November", "Desember")
        date_label = f"{selected_date.day} {months[selected_date.month - 1]} {selected_date.year}"
    except ValueError:
        date_label = ""
    clean_subject = BeautifulSoup(str(subject or ""), "html.parser").get_text(" ", strip=True)
    basename = f"{prefix}-{sequence or 'Draft'} {date_label} {clean_subject}".strip()
    basename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", basename)
    basename = re.sub(r"\s+", " ", basename).strip(" .")
    return basename[:220].rstrip(" .") or f"{prefix}-Draft"


def inteliz_credential_cipher():
    secret = str(app.config["SECRET_KEY"]).encode("utf-8")
    key = base64.urlsafe_b64encode(hashlib.sha256(secret + b":inteliz-credentials:v1").digest())
    return Fernet(key)


def get_inteliz_credentials(user_id):
    """Used by the upcoming automation without exposing the password to templates."""
    row = fetch_one(
        "SELECT inteliz_username, inteliz_password_encrypted FROM inteliz_user_settings WHERE user_id=%s",
        (user_id,),
    )
    if not row:
        return None
    try:
        password = inteliz_credential_cipher().decrypt(
            row["inteliz_password_encrypted"].encode("ascii")
        ).decode("utf-8")
    except (InvalidToken, UnicodeError, ValueError):
        app.logger.error("Kredensial Inteliz pengguna %s tidak dapat didekripsi.", user_id)
        return None
    return {"username": row["inteliz_username"], "password": password}


def get_inteliz_session_data(user_id):
    row = fetch_one(
        "SELECT session_data_encrypted FROM inteliz_user_settings WHERE user_id=%s", (user_id,)
    )
    if not row or not row["session_data_encrypted"]:
        return None
    try:
        decrypted = inteliz_credential_cipher().decrypt(
            row["session_data_encrypted"].encode("ascii")
        ).decode("utf-8")
        return json.loads(decrypted)
    except (InvalidToken, UnicodeError, ValueError, json.JSONDecodeError):
        app.logger.error("Sesi Inteliz pengguna %s tidak dapat didekripsi.", user_id)
        return None


def html_form_value(soup, field_name):
    field = soup.select_one(f'[name="{field_name}"]')
    if field is None:
        return ""
    if field.name == "select":
        selected = field.select_one("option[selected]") or field.select_one("option[value]")
        return selected.get("value", "") if selected else ""
    return field.get("value", "")


def inteliz_signatory_value(report, soup, saved_options=None):
    options = soup.select('[name="penandatangan"] option[value]')
    valid_options = [(option.get("value", "").strip(), option.get_text(" ", strip=True))
                     for option in options if option.get("value", "").strip()]
    if not valid_options:
        valid_options = [(str(item.get("value", "")).strip(), str(item.get("label", "")).strip())
                         for item in (saved_options or []) if str(item.get("value", "")).strip()]

    # Inteliz memakai pejabat autentikasi/Kasi Intelijen sebagai penandatangan,
    # bukan Kasubsi yang menyusun laporan. NIP pada database dapat mengandung spasi.
    for source in (report.get("auth_rank_nip"), report.get("creator_rank_nip")):
        digits = re.sub(r"\D", "", source or "")
        nip = digits[-18:] if len(digits) >= 18 else ""
        if nip and (not valid_options or any(value == nip for value, _ in valid_options)):
            return nip

    signer_name = (report.get("auth_name") or report.get("creator_name") or "").casefold()
    for value, label in valid_options:
        if signer_name and signer_name in label.casefold():
            return value
    return ""


def save_refreshed_inteliz_session(user_id, session_data, http_session, csrf_token):
    session_data["cookies"] = [
        {"name": cookie.name, "value": cookie.value, "domain": cookie.domain,
         "path": cookie.path or "/", "expires": cookie.expires or -1,
         "secure": bool(cookie.secure), "httpOnly": bool(cookie.has_nonstandard_attr("HttpOnly"))}
        for cookie in http_session.cookies
    ]
    session_data.setdefault("lapinhar_create", {})["csrf_token"] = csrf_token
    session_data["captured_at"] = int(time.time())
    encrypted = inteliz_credential_cipher().encrypt(
        json.dumps(session_data, ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    with get_db().cursor() as cursor:
        cursor.execute(
            "UPDATE inteliz_user_settings SET session_data_encrypted=%s WHERE user_id=%s",
            (encrypted, user_id),
        )


@app.post("/lapinhar/<int:report_id>/sync-inteliz")
@login_required
def sync_lapinhar_inteliz(report_id):
    report = accessible_lapinhar(report_id)
    if report is None:
        flash("LAPINHAR tidak ditemukan atau tidak dapat Anda akses.", "error")
        return redirect(url_for("lapinhar"))
    if report["inteliz_status"] == "sudah":
        flash("LAPINHAR tersebut sudah tersinkron ke Inteliz.", "success")
        return redirect(url_for("lapinhar"))

    session_data = get_inteliz_session_data(report["created_by"])
    if not session_data or not session_data.get("cookies"):
        flash("Sesi Inteliz pembuat laporan belum tersedia. Hubungkan Inteliz terlebih dahulu.", "error")
        if report["created_by"] == session["user_id"]:
            return redirect(url_for("lapinhar", connect_inteliz=1))
        return redirect(url_for("lapinhar"))

    http_session = requests.Session()
    for cookie in session_data["cookies"]:
        if cookie.get("expires", -1) not in (-1, None) and cookie["expires"] < time.time():
            continue
        http_session.cookies.set(cookie["name"], cookie["value"],
                                 domain=cookie.get("domain") or "inteliz.kejaksaan.go.id",
                                 path=cookie.get("path") or "/")
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "id,en-US;q=0.9,en;q=0.8",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36",
    }
    try:
        create_response = http_session.get(INTELIZ_LAPINHAR_CREATE_URL, headers=headers, timeout=45)
        if "/login" in create_response.url.lower():
            raise IntelizAuthenticationRequired("Sesi Inteliz sudah kedaluwarsa.")
        create_response.raise_for_status()
        soup = BeautifulSoup(create_response.text, "html.parser")
        csrf_token = html_form_value(soup, "_token") or (
            soup.select_one('meta[name="csrf-token"]').get("content", "")
            if soup.select_one('meta[name="csrf-token"]') else ""
        )
        saved_create = session_data.get("lapinhar_create", {})
        id_satker = html_form_value(soup, "id_satker") or saved_create.get("id_satker", "")
        signatory = inteliz_signatory_value(report, soup, saved_create.get("signatory_options"))
        if not csrf_token or not id_satker or not signatory:
            missing = []
            if not csrf_token:
                missing.append("token CSRF")
            if not id_satker:
                missing.append("ID satker")
            if not signatory:
                missing.append("penandatangan")
            app.logger.warning("Metadata create Inteliz laporan %s belum lengkap: %s; url=%s; fields=%s",
                               report_id, ", ".join(missing), create_response.url,
                               sorted({field.get("name") for field in soup.select("[name]") if field.get("name")}))
            raise RuntimeError(
                f"Data form Inteliz belum lengkap: {', '.join(missing)}. Cookie login tetap disimpan."
            )

        fields = {
            "_token": csrf_token,
            "id_satker": str(id_satker),
            "kategori_laporan": str(report["category_id"] or ""),
            "nomor": report["report_number"] or "",
            "tgl": report["report_date"].isoformat() if report["report_date"] else "",
            "info_yg_diperoleh": report["facts"] or "",
            "sumber_info": report["source_name"] or "",
            "trend_perkembangan": report["analysis"] or "",
            "pendapat_saran": report["recommendation"] or "",
            "penandatangan": signatory,
        }
        if not all(fields.values()):
            raise RuntimeError("Data LAPINHAR belum lengkap untuk dikirim ke Inteliz.")
        upload_response = http_session.post(
            "https://inteliz.kejaksaan.go.id/lapinhar",
            files={name: (None, value) for name, value in fields.items()},
            headers={**headers, "Referer": INTELIZ_LAPINHAR_CREATE_URL}, timeout=60,
        )
        if upload_response.status_code == 419 or "/login" in upload_response.url.lower():
            raise IntelizAuthenticationRequired("Sesi atau token Inteliz kedaluwarsa.")
        upload_response.raise_for_status()
        if urlparse(upload_response.url).path.rstrip("/") != "/lapinhar":
            error_soup = BeautifulSoup(upload_response.text, "html.parser")
            validation = " ".join(node.get_text(" ", strip=True) for node in
                                  error_soup.select(".invalid-feedback, .alert-danger, .text-danger") if node.get_text(strip=True))
            raise RuntimeError(validation or "Inteliz belum menerima data LAPINHAR.")
        save_refreshed_inteliz_session(report["created_by"], session_data, http_session, csrf_token)
        with get_db().cursor() as cursor:
            cursor.execute("UPDATE lapinhar_reports SET inteliz_status='sudah' WHERE id=%s", (report_id,))
        flash("LAPINHAR berhasil disinkronkan ke Inteliz.", "success")
    except IntelizAuthenticationRequired as exc:
        with get_db().cursor() as cursor:
            cursor.execute(
                "UPDATE inteliz_user_settings SET connected_at=NULL WHERE user_id=%s",
                (report["created_by"],),
            )
        flash(f"{exc} Silakan selesaikan login Inteliz pada popup.", "error")
        if report["created_by"] == session["user_id"]:
            return redirect(url_for("lapinhar", connect_inteliz=1))
        flash("Pembuat LAPINHAR harus menghubungkan ulang akun Inteliz miliknya.", "error")
        return redirect(url_for("lapinhar"))
    except (requests.RequestException, RuntimeError) as exc:
        app.logger.warning("Sinkron Inteliz laporan %s gagal: %s", report_id, exc)
        flash(f"Sinkron Inteliz gagal: {exc}", "error")
    return redirect(url_for("lapinhar"))


def read_inteliz_create_context(page):
    return page.evaluate("""() => {
        const form = document.querySelector('form[action*="lapinhar"]') || document.querySelector('form');
        const valueOf = (name) => {
            const field = document.querySelector(`[name="${name}"]`);
            return field ? field.value : '';
        };
        const optionsOf = (name) => Array.from(document.querySelectorAll(`[name="${name}"] option`))
            .map(option => ({ value: option.value, label: option.textContent.trim() }))
            .filter(option => option.value);
        return {
            csrf_token: valueOf('_token') || document.querySelector('meta[name="csrf-token"]')?.content || '',
            form_action: form?.action || 'https://inteliz.kejaksaan.go.id/lapinhar',
            form_method: (form?.method || 'post').toUpperCase(),
            id_satker: valueOf('id_satker'),
            category_options: optionsOf('kategori_laporan'),
            signatory_options: optionsOf('penandatangan'),
            field_names: form ? Array.from(form.elements).map(field => field.name).filter(Boolean) : []
        };
    }""")


def update_inteliz_auth(auth_id, **values):
    with INTELIZ_AUTH_LOCK:
        state = INTELIZ_AUTH_SESSIONS.get(auth_id)
        if state:
            state.update(values, updated_at=time.time())


def wait_inteliz_auth_input(auth_id, field, timeout=600):
    deadline = time.time() + timeout
    while time.time() < deadline:
        with INTELIZ_AUTH_LOCK:
            state = INTELIZ_AUTH_SESSIONS.get(auth_id)
            if not state or state.get("cancelled"):
                return None
            value = state.pop(field, None)
        if value:
            return value
        time.sleep(.25)
    return None


def visible_inteliz_otp_input(page):
    selectors = (
        'input[name="code"]', 'input[autocomplete="one-time-code"]', 'input[name="otp"]',
        'input[name="token"]', 'input[name="authenticator_code"]',
        'input[name="one_time_password"]', 'input[id*="otp" i]', 'input[id*="auth" i]',
    )
    for selector in selectors:
        locator = page.locator(selector).first
        if locator.count() and locator.is_visible():
            return locator
    return None


def persist_inteliz_browser_session(user_id, context, page, create_context=None):
    cookies = context.cookies(["https://inteliz.kejaksaan.go.id"])
    local_storage = page.evaluate(
        "Object.fromEntries(Array.from({length: localStorage.length}, (_, i) => "
        "[localStorage.key(i), localStorage.getItem(localStorage.key(i))]))"
    )
    payload = {
        "cookies": cookies,
        "local_storage": local_storage,
        "lapinhar_create": create_context or {},
        "captured_url": page.url,
        "captured_at": int(time.time()),
    }
    encrypted_session = inteliz_credential_cipher().encrypt(
        json.dumps(payload, ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    with app.app_context():
        with get_db().cursor() as cursor:
            cursor.execute(
                """UPDATE inteliz_user_settings SET session_data_encrypted=%s,
                   connected_at=CURRENT_TIMESTAMP WHERE user_id=%s""",
                (encrypted_session, user_id),
            )
    return len(cookies)


def run_inteliz_auth(auth_id, user_id, credentials):
    browser = None
    try:
        if not CHROME_EXECUTABLE.exists():
            raise RuntimeError("Google Chrome tidak ditemukan pada komputer server.")
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                executable_path=str(CHROME_EXECUTABLE),
                headless=os.getenv("INTELIZ_BROWSER_HEADLESS", "1") == "1",
                slow_mo=int(os.getenv("INTELIZ_BROWSER_SLOW_MO", "0")),
                args=["--disable-blink-features=AutomationControlled"],
            )
            context = browser.new_context(viewport={"width": 1280, "height": 900}, locale="id-ID")
            page = context.new_page()
            update_inteliz_auth(auth_id, status="loading", message="Membuka halaman Inteliz…")
            page.goto(INTELIZ_LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
            page.locator("#username").fill(credentials["username"])
            page.locator("#password-field").fill(credentials["password"])

            authenticated = False
            for attempt in range(3):
                if page.locator("#username").count():
                    page.locator("#username").fill(credentials["username"])
                if page.locator("#password-field").count():
                    page.locator("#password-field").fill(credentials["password"])
                captcha = page.locator("#captcha-container img").first
                captcha.wait_for(state="visible", timeout=20000)
                captcha_data = base64.b64encode(captcha.screenshot(type="png")).decode("ascii")
                update_inteliz_auth(
                    auth_id, status="captcha", captcha=f"data:image/png;base64,{captcha_data}",
                    message="Masukkan kode CAPTCHA yang tampil.",
                )
                captcha_value = wait_inteliz_auth_input(auth_id, "captcha_input")
                if captcha_value is None:
                    raise RuntimeError("Sesi login dibatalkan atau kedaluwarsa.")
                page.locator('input[name="captcha"]').fill(captcha_value)
                page.locator("#kt_sign_in_submit").click()
                page.wait_for_timeout(2500)
                otp_input = visible_inteliz_otp_input(page)
                current_path = urlparse(page.url).path.rstrip("/")
                if otp_input is None and current_path == INTELIZ_2FA_PATH:
                    try:
                        page.locator('input[name="code"]').wait_for(state="visible", timeout=10000)
                    except Exception:
                        pass
                    otp_input = visible_inteliz_otp_input(page)
                if otp_input is not None:
                    break
                if "/login" not in page.url.lower() and current_path != INTELIZ_2FA_PATH:
                    authenticated = True
                    break
                update_inteliz_auth(auth_id, status="loading", captcha=None,
                                    message="CAPTCHA belum diterima, memuat CAPTCHA baru…")

            if not authenticated:
                otp_input = visible_inteliz_otp_input(page)
                if otp_input is None:
                    raise RuntimeError("Login belum berhasil. Periksa CAPTCHA atau kredensial Inteliz.")
                for _attempt in range(3):
                    update_inteliz_auth(auth_id, status="otp", captcha=None,
                                        message="Masukkan kode autentikator Inteliz.")
                    otp_value = wait_inteliz_auth_input(auth_id, "otp_input")
                    if otp_value is None:
                        raise RuntimeError("Sesi autentikator dibatalkan atau kedaluwarsa.")
                    otp_input.fill(otp_value)
                    otp_form = otp_input.locator("xpath=ancestor::form[1]")
                    submit = otp_form.locator('button[type="submit"], input[type="submit"]').first
                    if not submit.count():
                        submit = page.locator('button[type="submit"], input[type="submit"]').first
                    if not submit.count():
                        raise RuntimeError("Tombol verifikasi autentikator tidak ditemukan.")
                    submit.click()
                    try:
                        page.wait_for_function(
                            "() => !['/login', '/2fa/challenge'].includes(location.pathname.replace(/\\/$/, ''))",
                            timeout=30000,
                        )
                    except Exception:
                        page.wait_for_timeout(1000)
                    otp_input = visible_inteliz_otp_input(page)
                    current_path = urlparse(page.url).path.rstrip("/")
                    if (otp_input is None and "/login" not in page.url.lower()
                            and current_path != INTELIZ_2FA_PATH):
                        authenticated = True
                        break
                if not authenticated:
                    raise RuntimeError("Kode autentikator belum diterima oleh Inteliz.")

            cookie_count = persist_inteliz_browser_session(user_id, context, page)
            if not cookie_count:
                raise RuntimeError("Dashboard terbuka, tetapi cookie Inteliz tidak ditemukan.")
            update_inteliz_auth(auth_id, status="loading", captcha=None,
                                message="Cookie tersimpan. Membaca formulir LAPINHAR Inteliz…")
            create_context = None
            try:
                page.goto(INTELIZ_LAPINHAR_CREATE_URL, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(1200)
                if urlparse(page.url).path.rstrip("/") == "/lapinhar/create":
                    create_context = read_inteliz_create_context(page)
                    persist_inteliz_browser_session(user_id, context, page, create_context)
            except Exception as form_exc:
                app.logger.warning("Cookie Inteliz tersimpan, metadata form belum terbaca: %s", form_exc)
            update_inteliz_auth(auth_id, status="success", captcha=None,
                                message="Inteliz berhasil terhubung. Cookie telah disimpan terenkripsi.")
    except Exception as exc:
        app.logger.warning("Login Inteliz pengguna %s gagal: %s", user_id, exc)
        update_inteliz_auth(auth_id, status="error", captcha=None, message=str(exc))
    finally:
        if browser:
            try:
                browser.close()
            except Exception:
                pass


def get_sipede_credentials(user_id):
    row = fetch_one(
        "SELECT sipede_username, sipede_password_encrypted FROM sipede_user_settings WHERE user_id=%s",
        (user_id,),
    )
    if not row or not row.get("sipede_password_encrypted"):
        return None
    try:
        password = inteliz_credential_cipher().decrypt(
            row["sipede_password_encrypted"].encode("ascii")
        ).decode("utf-8")
    except Exception:
        app.logger.error("Kredensial SIPede pengguna %s tidak dapat didekripsi.", user_id)
        return None
    return {"username": row["sipede_username"], "password": password}


def get_sipede_session_data(user_id):
    row = fetch_one(
        "SELECT session_data_encrypted FROM sipede_user_settings WHERE user_id=%s",
        (user_id,),
    )
    if not row or not row.get("session_data_encrypted"):
        return None
    try:
        decrypted = inteliz_credential_cipher().decrypt(
            row["session_data_encrypted"].encode("ascii")
        )
        return json.loads(decrypted.decode("utf-8"))
    except Exception:
        app.logger.error("Sesi SIPede pengguna %s tidak dapat didekripsi.", user_id)
        return None


def sipede_create_context_from_html(html, response_url):
    soup = BeautifulSoup(html, "html.parser")
    form = soup.select_one('#formCreateEdit, form[action$="/suratkeluar"]')
    if not form:
        return None
    token = form.select_one('input[name="_token"]') or soup.select_one('meta[name="csrf-token"]')
    def options_of(name):
        return [
            {"value": option.get("value", ""), "label": option.get_text(" ", strip=True)}
            for option in form.select(f'[name="{name}"] option[value]')
            if option.get("value")
        ]

    return {
        "csrf_token": (token.get("value") or token.get("content") or "") if token else "",
        "form_action": form.get("action") or f"{SIPEDE_BASE_URL}/suratkeluar",
        "form_method": (form.get("method") or "post").upper(),
        "form_enctype": form.get("enctype") or "multipart/form-data",
        "id_surat": (form.select_one('input[name="idSurat"]') or {}).get("value", "125"),
        "jenis_options": options_of("jenis"),
        "sifat_options": options_of("sifat"),
        "kode_masalah_options": options_of("kode_masalah"),
        "penandatangan_options": options_of("penandatangan"),
        "field_names": [field.get("name") for field in form.select("input[name],select[name],textarea[name]")],
        "captured_url": response_url,
    }


def sipede_signatory_id(create_context, configured_name):
    normalized_target = re.sub(r"[^A-Z0-9]", "", (configured_name or "").upper())
    if not normalized_target:
        return None
    for option in create_context.get("penandatangan_options", []):
        normalized_label = re.sub(r"[^A-Z0-9]", "", option.get("label", "").upper())
        if normalized_target in normalized_label:
            return option.get("value")
    return None


def sipede_issue_code_value(create_context, report_issue_code):
    target = (report_issue_code or "").strip()
    if not target:
        return None
    for option in create_context.get("kode_masalah_options", []):
        if option.get("value", "").strip().casefold() == target.casefold():
            return option.get("value", "").strip()
    normalized_target = re.sub(r"[^A-Z0-9]", "", target.upper())
    for option in create_context.get("kode_masalah_options", []):
        value = option.get("value", "").strip()
        label = option.get("label", "")
        normalized_value = re.sub(r"[^A-Z0-9]", "", value.upper())
        normalized_label = re.sub(r"[^A-Z0-9]", "", label.upper())
        if normalized_target == normalized_value or normalized_target in normalized_label:
            return value
    return None


def save_refreshed_sipede_session(user_id, session_data, http_session, create_context):
    session_data["cookies"] = [
        {
            "name": cookie.name,
            "value": cookie.value,
            "domain": cookie.domain,
            "path": cookie.path,
            "secure": cookie.secure,
        }
        for cookie in http_session.cookies
    ]
    session_data["suratkeluar_create"] = create_context or {}
    session_data["captured_url"] = (create_context or {}).get("captured_url", SIPEDE_SURATKELUAR_CREATE_URL)
    session_data["captured_at"] = int(time.time())
    encrypted = inteliz_credential_cipher().encrypt(
        json.dumps(session_data, ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    with get_db().cursor() as cursor:
        cursor.execute(
            """UPDATE sipede_user_settings SET session_data_encrypted=%s,
               connected_at=CURRENT_TIMESTAMP WHERE user_id=%s""",
            (encrypted, user_id),
        )


def update_sipede_auth(auth_id, **values):
    with SIPEDE_AUTH_LOCK:
        state = SIPEDE_AUTH_SESSIONS.get(auth_id)
        if state:
            state.update(values, updated_at=time.time())


def wait_sipede_auth_input(auth_id, field, timeout=600):
    deadline = time.time() + timeout
    while time.time() < deadline:
        with SIPEDE_AUTH_LOCK:
            state = SIPEDE_AUTH_SESSIONS.get(auth_id)
            if not state or state.get("cancelled"):
                return None
            value = state.pop(field, None)
        if value:
            return value
        time.sleep(.25)
    return None


def read_sipede_create_context(page):
    return page.evaluate("""() => {
        const form = document.querySelector('#formCreateEdit') ||
            document.querySelector('form[action$="/suratkeluar"]');
        const optionsOf = (name) => Array.from(document.querySelectorAll(`[name="${name}"] option`))
            .map(option => ({ value: option.value, label: option.textContent.trim() }))
            .filter(option => option.value);
        const valueOf = (name) => document.querySelector(`[name="${name}"]`)?.value || '';
        return {
            csrf_token: valueOf('_token') || document.querySelector('meta[name="csrf-token"]')?.content || '',
            form_action: form?.action || 'https://sipede.kejaksaan.go.id/suratkeluar',
            form_method: (form?.method || 'post').toUpperCase(),
            form_enctype: form?.enctype || 'multipart/form-data',
            id_surat: valueOf('idSurat') || '125',
            jenis_options: optionsOf('jenis'),
            sifat_options: optionsOf('sifat'),
            kode_masalah_options: optionsOf('kode_masalah'),
            penandatangan_options: optionsOf('penandatangan'),
            field_names: form ? Array.from(form.elements).map(field => field.name).filter(Boolean) : []
        };
    }""")


def persist_sipede_browser_session(user_id, context, page, create_context=None):
    cookies = context.cookies([SIPEDE_BASE_URL])
    local_storage = page.evaluate(
        "Object.fromEntries(Array.from({length: localStorage.length}, (_, i) => "
        "[localStorage.key(i), localStorage.getItem(localStorage.key(i))]))"
    )
    payload = {
        "cookies": cookies,
        "local_storage": local_storage,
        "suratkeluar_create": create_context or {},
        "captured_url": page.url,
        "captured_at": int(time.time()),
    }
    encrypted = inteliz_credential_cipher().encrypt(
        json.dumps(payload, ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    with app.app_context():
        with get_db().cursor() as cursor:
            cursor.execute(
                """UPDATE sipede_user_settings
                   SET session_data_encrypted=%s, connected_at=CURRENT_TIMESTAMP
                   WHERE user_id=%s""",
                (encrypted, user_id),
            )
    return len(cookies)


def run_sipede_auth(auth_id, user_id, credentials):
    browser = None
    try:
        if not CHROME_EXECUTABLE.exists():
            raise RuntimeError("Google Chrome tidak ditemukan pada komputer server.")
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                executable_path=str(CHROME_EXECUTABLE),
                headless=os.getenv("SIPEDE_BROWSER_HEADLESS", "0") == "1",
                slow_mo=int(os.getenv("SIPEDE_BROWSER_SLOW_MO", "250")),
                args=["--disable-blink-features=AutomationControlled"],
            )
            context = browser.new_context(viewport={"width": 1280, "height": 900}, locale="id-ID")
            page = context.new_page()
            update_sipede_auth(auth_id, status="loading", message="Membuka halaman SIPede…")
            page.goto(SIPEDE_LOGIN_URL, wait_until="domcontentloaded", timeout=60000)

            authenticated = False
            for _attempt in range(3):
                page.locator('#username, input[name="username"]').first.fill(credentials["username"])
                page.locator('input[name="password"]').first.fill(credentials["password"])
                captcha_image = page.locator('.captcha img, img[src*="captcha" i]').first
                captcha_image.wait_for(state="visible", timeout=30000)
                captcha_data = base64.b64encode(captcha_image.screenshot(type="png")).decode("ascii")
                update_sipede_auth(
                    auth_id,
                    status="captcha",
                    captcha=f"data:image/png;base64,{captcha_data}",
                    message="Masukkan kode CAPTCHA SIPede yang tampil.",
                )
                captcha_value = wait_sipede_auth_input(auth_id, "captcha_input")
                if captcha_value is None:
                    raise RuntimeError("Sesi login SIPede dibatalkan atau kedaluwarsa.")
                page.locator('#captcha, input[name="captcha"]').first.fill(captcha_value)
                page.locator('form[action*="login-post"] button[type="submit"], button[type="submit"]').first.click()
                try:
                    page.wait_for_function(
                        "() => !location.pathname.replace(/\\/$/, '').endsWith('/login')",
                        timeout=30000,
                    )
                except Exception:
                    page.wait_for_timeout(1200)
                if "/login" not in urlparse(page.url).path.lower():
                    authenticated = True
                    break
                update_sipede_auth(
                    auth_id,
                    status="loading",
                    captcha=None,
                    message="Login belum diterima. Memuat CAPTCHA baru…",
                )

            if not authenticated:
                raise RuntimeError("Login SIPede belum berhasil. Periksa CAPTCHA, username, atau password.")
            cookie_count = persist_sipede_browser_session(user_id, context, page)
            if not cookie_count:
                raise RuntimeError("Halaman SIPede terbuka, tetapi cookie sesi tidak ditemukan.")
            update_sipede_auth(
                auth_id,
                status="loading",
                captcha=None,
                message="Cookie tersimpan. Membuka formulir Surat Keluar SIPede…",
            )
            create_context = None
            try:
                page.goto(SIPEDE_SURATKELUAR_CREATE_URL, wait_until="domcontentloaded", timeout=60000)
                page.locator('#formCreateEdit, form[action$="/suratkeluar"]').first.wait_for(
                    state="attached", timeout=30000
                )
                create_context = read_sipede_create_context(page)
                persist_sipede_browser_session(user_id, context, page, create_context)
            except Exception as form_exc:
                app.logger.warning("Cookie SIPede tersimpan, metadata form Surat Keluar belum terbaca: %s", form_exc)
            update_sipede_auth(
                auth_id,
                status="success",
                captcha=None,
                message=("SIPede berhasil terhubung. Sesi, cookie, token, dan form Surat Keluar "
                         "telah disimpan terenkripsi."),
            )
    except Exception as exc:
        app.logger.warning("Login SIPede pengguna %s gagal: %s", user_id, exc)
        update_sipede_auth(auth_id, status="error", captcha=None, message=str(exc))
    finally:
        if browser:
            try:
                browser.close()
            except Exception:
                pass


@app.post("/lapinhar/export/<file_type>")
@login_required
def export_lapinhar(file_type):
    data = lapinhar_form_data()
    document_kind = request.form.get("document_kind", "lapinhar")
    is_lapinsus = document_kind == "lapinsus"
    document_label = "LAPINSUS" if is_lapinsus else "LAPINHAR"
    document_heading = "LAPORAN INFORMASI KHUSUS" if is_lapinsus else "LAPORAN INFORMASI HARIAN"
    attachments = collect_attachments()
    data["attachment"] = attachment_label(len(attachments))
    export_name = lapinhar_export_basename(data["report_number"], data["report_date"], data["subject"])
    if file_type == "docx":
        output = BytesIO()
        document = Document()
        section = document.sections[0]
        section.page_width, section.page_height = Cm(21.59), Cm(33.02)
        section.top_margin, section.bottom_margin = Cm(2), Cm(2)
        section.left_margin, section.right_margin = Cm(2), Cm(1.5)
        normal = document.styles["Normal"]
        normal.font.name, normal.font.size = "Times New Roman", Pt(12)
        normal.paragraph_format.space_before = Pt(0)
        normal.paragraph_format.space_after = Pt(0)
        normal.paragraph_format.line_spacing = 1.5
        organization = data["organization"] or "KEJAKSAAN NEGERI BULELENG"
        cover_org = document.add_paragraph()
        cover_org.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cover_org.add_run(organization).bold = True
        logo = document.add_picture(str(STATIC_IMG_DIR / "logo-kejaksaan.jpeg"), width=Cm(4.2))
        document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cover_title = document.add_paragraph()
        cover_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cover_title.paragraph_format.space_before = Pt(16)
        cover_title.add_run(document_label).bold = True
        cover_subject = document.add_paragraph()
        cover_subject.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cover_subject.add_run(data["subject"]).bold = True
        secret = document.add_picture(str(STATIC_IMG_DIR / "rahasia.jpeg"), width=Cm(17.5), height=Cm(14.5))
        document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        memo_section = document.add_section(WD_SECTION.NEW_PAGE)
        memo_section.page_width, memo_section.page_height = Cm(21.59), Cm(33.02)
        memo_section.top_margin, memo_section.bottom_margin = Cm(2), Cm(2)
        memo_section.left_margin, memo_section.right_margin = Cm(2), Cm(1.5)
        office = document.add_paragraph(organization)
        office.alignment = WD_ALIGN_PARAGRAPH.CENTER
        office.runs[0].bold, office.runs[0].underline = True, True
        document.add_paragraph("\n\n")
        title = document.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("NOTA  -  DINAS")
        run.bold, run.underline = True, True
        document.add_paragraph()
        for label, value in (("K E P A D A", data["recipient"]), ("D A R I", data["sender"]),
                             ("NOMOR", data["report_number"]), ("TANGGAL", data["report_date"]),
                             ("SIFAT", data["classification"]), ("LAMPIRAN", data["attachment"]),
                             ("PERIHAL", data["subject"])):
            add_docx_label(document, label, value)
        document.add_paragraph("\n")
        signer = document.add_paragraph()
        signer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        signer.paragraph_format.left_indent = Cm(9)
        signer.add_run(f"{data['auth_position']}\n{organization}\n\n\n\n")
        signer.add_run(data["auth_name"]).bold = True
        signer.add_run(f"\n{data['auth_rank_nip']}")

        second = document.add_section(WD_SECTION.NEW_PAGE)
        second.page_width, second.page_height = Cm(21.59), Cm(33.02)
        second.top_margin, second.bottom_margin = Cm(2), Cm(2)
        second.left_margin, second.right_margin = Cm(2), Cm(1.5)
        second.header.is_linked_to_previous = False
        second.footer.is_linked_to_previous = False
        report_header = second.header.paragraphs[0]
        report_header.text = "RAHASIA"
        report_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        report_header.runs[0].font.name, report_header.runs[0].font.size = "Times New Roman", Pt(10)
        report_footer = second.footer.paragraphs[0]
        report_footer.text = "RAHASIA"
        report_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        report_footer.runs[0].font.name, report_footer.runs[0].font.size = "Times New Roman", Pt(10)
        document.add_paragraph(f"{organization}\t{'L. IN.2' if is_lapinsus else 'L. IN.1'}")
        document.add_paragraph("Copy ke 1\nDari 1 copies")
        report_title = document.add_paragraph()
        report_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        report_title.paragraph_format.line_spacing = 1.5
        rr = report_title.add_run(document_heading)
        rr.bold, rr.underline = True, True
        number = document.add_paragraph(f"NOMOR : {data['report_number']}")
        number.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if is_lapinsus:
            subject_line = document.add_paragraph(f"PERIHAL : {data['subject']}")
            subject_line.paragraph_format.space_after = Pt(12)
        add_docx_section(document, "I.    INFORMASI YANG DIPEROLEH", data["information"], data["information_spacing"])
        add_docx_section(document, "II.   SUMBER INFORMASI", data["sources"], data["sources_spacing"])
        add_docx_section(document, "III.  TREN PERKEMBANGAN / PERKIRAAN", data["trends"], data["trends_spacing"])
        add_docx_section(document, "IV.   PENDAPAT / SARAN", data["suggestions"], data["suggestions_spacing"])
        signatures = document.add_table(rows=1, cols=2)
        signature_cells = []
        if data["show_authentication"]:
            left = signatures.rows[0].cells[0]
            left.text = f"AUTENTIKASI:\n{data['auth_position']}\n{organization}\n\n\n\n\n{data['auth_name']}\n{data['auth_rank_nip']}"
            signature_cells.append(left)
            right = signatures.rows[0].cells[1]
        else:
            right = signatures.rows[0].cells[1]
        right.text = f"{data['city']}, {data['report_date']}\nYang Membuat Laporan\n{data['creator_position']}\n\n\n\n\n{data['creator_name']}\n{data['creator_rank_nip']}"
        signature_cells.append(right)
        for cell in signature_cells:
            for paragraph in cell.paragraphs: paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for attachment in attachments:
            document.add_page_break()
            attachment_heading = document.add_paragraph(f"LAMPIRAN {document_heading}")
            attachment_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            attachment_heading.runs[0].bold = True
            document.add_paragraph(f"NOMOR : {data['report_number']}").alignment = WD_ALIGN_PARAGRAPH.CENTER
            document.add_paragraph(f"TANGGAL : {data['report_date']}").alignment = WD_ALIGN_PARAGRAPH.CENTER
            if attachment_layout(attachment) == "side" and len(attachment["images"]) == 2:
                photo_table = document.add_table(rows=1, cols=2)
                for cell, image_data in zip(photo_table.rows[0].cells, attachment["images"]):
                    ratio = image_data["width"] / image_data["height"]
                    scale = image_data["scale"] / 100
                    width = min(7.2, 18 * ratio) * scale
                    height = width / ratio
                    run = cell.paragraphs[0].add_run()
                    run.add_picture(BytesIO(image_data["content"]), width=Cm(width), height=Cm(height))
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                for image_data in attachment["images"]:
                    ratio = image_data["width"] / image_data["height"]
                    scale = image_data["scale"] / 100
                    max_width, max_height = 15.5, (9 if len(attachment["images"]) == 2 else 17)
                    width = min(max_width, max_height * ratio) * scale
                    height = width / ratio
                    document.add_picture(BytesIO(image_data["content"]), width=Cm(width), height=Cm(height))
                    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        document.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=f"{export_name}.docx",
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if file_type == "pdf":
        output = BytesIO()
        styles = getSampleStyleSheet()
        body = ParagraphStyle("ReportBody", parent=styles["BodyText"], fontName=PDF_FONT_NAME,
                              fontSize=12, leading=18, alignment=TA_JUSTIFY, spaceBefore=0, spaceAfter=0)
        heading = ParagraphStyle("ReportHeading", parent=body, fontName=PDF_FONT_BOLD,
                                 spaceBefore=0, spaceAfter=0, keepWithNext=True)
        centered = ParagraphStyle("ReportTitle", parent=heading, fontSize=12, leading=18, alignment=TA_CENTER)
        folio = (21.59*cm, 33.02*cm)
        pdf = SimpleDocTemplate(output, pagesize=folio, rightMargin=1.5*cm, leftMargin=2*cm,
                                topMargin=2*cm, bottomMargin=2*cm, title=f"LAPINHAR {data['report_number']}")
        months = ("", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
                  "Agustus", "September", "Oktober", "November", "Desember")
        try:
            report_day = date.fromisoformat(data["report_date"])
            display_date = f"{report_day.day} {months[report_day.month]} {report_day.year}"
        except (TypeError, ValueError):
            display_date = data["report_date"] or "—"

        def draw_cover(canvas, document_template):
            page_width, page_height = folio
            canvas.saveState()
            canvas.setFont(PDF_FONT_BOLD, 12)
            canvas.drawCentredString(page_width / 2, page_height - 2.2 * cm, data["organization"])
            canvas.drawImage(str(STATIC_IMG_DIR / "logo-kejaksaan.jpeg"),
                             (page_width - 4.2 * cm) / 2, page_height - 7.85 * cm,
                             width=4.2 * cm, height=4.1 * cm, preserveAspectRatio=True, mask="auto")
            canvas.drawCentredString(page_width / 2, page_height - 8.9 * cm, "LAPINHAR")
            canvas.setFont(PDF_FONT_BOLD, 12)
            words, lines, current = data["subject"].split(), [], ""
            for word in words:
                candidate = f"{current} {word}".strip()
                if canvas.stringWidth(candidate, PDF_FONT_BOLD, 12) <= 17.5 * cm:
                    current = candidate
                else:
                    lines.append(current)
                    current = word
            if current:
                lines.append(current)
            for index, line in enumerate(lines):
                canvas.drawCentredString(page_width / 2, page_height - (9.55 * cm + index * .55 * cm), line)
            canvas.saveState()
            canvas.translate(page_width * .43, 11 * cm)
            canvas.rotate(56)
            canvas.drawImage(str(STATIC_IMG_DIR / "rahasia.jpeg"),
                             -8.65 * cm, -1.375 * cm, width=17.3 * cm, height=2.75 * cm,
                             preserveAspectRatio=True, mask="auto")
            canvas.restoreState()
            canvas.drawCentredString(page_width / 2, 1.35 * cm, f"{data['city']}, {display_date}")
            canvas.restoreState()

        story = [PageBreak(), Paragraph(escape(data["organization"]), centered), Spacer(1, 3)]
        office_rule = Table([[""]], colWidths=[17.5 * cm], rowHeights=[3])
        office_rule.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, -1), 1.2, "black"),
                                         ("LINEBELOW", (0, 0), (-1, -1), .6, "black")]))
        story.extend([office_rule, Spacer(1, 42), Paragraph("<u>NOTA  -  DINAS</u>", centered), Spacer(1, 25)])
        for label, value in (("K E P A D A", data["recipient"]), ("D A R I", data["sender"]),
                             ("NOMOR", data["report_number"]), ("TANGGAL", display_date),
                             ("SIFAT", data["classification"]), ("LAMPIRAN", data["attachment"]),
                             ("PERIHAL", data["subject"])):
            story.append(Table([[label, ":", Paragraph(escape(value), body)]], colWidths=[3.3*cm,.5*cm,14*cm]))
        memo_signer = Paragraph(
            f"{escape(data['auth_position'])}<br/>{escape(data['organization'])}<br/><br/><br/><br/><br/>"
            f"<u>{escape(data['auth_name'])}</u><br/>{escape(data['auth_rank_nip'])}", centered)
        story.extend([Spacer(1, 70), Table([["", memo_signer]], colWidths=[8.5*cm, 8.5*cm]), PageBreak(),
                      Table([[Paragraph(f"<b>{escape(data['organization'])}</b>", body),
                              Paragraph("<b>L. IN.1</b><br/>Copy ke 1<br/>Dari 1 copies", body)]],
                            colWidths=[13.2*cm, 3.8*cm]), Spacer(1, 18),
                      Paragraph("<u>LAPORAN INFORMASI HARIAN</u>", centered),
                      Paragraph(f"NOMOR : {escape(data['report_number'])}", centered), Spacer(1, 12)])
        for title, content, spacing in (("I.    INFORMASI YANG DIPEROLEH", data["information"], data["information_spacing"]),
                               ("II.   SUMBER INFORMASI", data["sources"], data["sources_spacing"]),
                               ("III.  TREN PERKEMBANGAN / PERKIRAAN", data["trends"], data["trends_spacing"]),
                               ("IV.   PENDAPAT / SARAN", data["suggestions"], data["suggestions_spacing"])):
            story.append(Paragraph(title, heading))
            for block in rich_text_blocks(content):
                prefix = escape(block["prefix"])
                formatted = "".join(
                    ("<b>" if run["bold"] else "") + ("<i>" if run["italic"] else "") +
                    ("<u>" if run["underline"] else "") + escape(run["text"]) +
                    ("</u>" if run["underline"] else "") + ("</i>" if run["italic"] else "") +
                    ("</b>" if run["bold"] else "") for run in block["runs"]
                )
                pdf_alignments = {"center": TA_CENTER, "justify": TA_JUSTIFY, "left": TA_LEFT, "right": TA_RIGHT}
                block_style = ParagraphStyle(
                    f"Section{len(story)}", parent=body, leading=18,
                    alignment=pdf_alignments.get(block["align"], TA_JUSTIFY),
                    leftIndent=(block["indent"] * 18) + (18 if prefix else 0),
                    firstLineIndent=-18 if prefix else 0,
                )
                story.append(Paragraph(f"{prefix}{formatted}", block_style))
        auth_signature = Paragraph(f"AUTENTIKASI:<br/>{escape(data['auth_position'])}<br/>{escape(data['organization'])}<br/><br/><br/><br/><br/><u>{escape(data['auth_name'])}</u><br/>{escape(data['auth_rank_nip'])}", centered)
        creator_signature = Paragraph(f"{escape(data['city'])}, {escape(display_date)}<br/>Yang Membuat Laporan<br/>{escape(data['creator_position'])}<br/><br/><br/><br/><br/><u>{escape(data['creator_name'])}</u><br/>{escape(data['creator_rank_nip'])}", centered)
        if data["show_authentication"]:
            signature_table = Table([[auth_signature, creator_signature]], colWidths=[8.5*cm, 8.5*cm])
        else:
            signature_table = Table([["", creator_signature]], colWidths=[8.5*cm, 8.5*cm])
        story.extend([Spacer(1, 20), signature_table])
        for attachment in attachments:
            story.extend([PageBreak(), Paragraph("<b>LAMPIRAN LAPORAN INFORMASI HARIAN</b>", centered),
                          Paragraph(f"NOMOR : {escape(data['report_number'])}", centered),
                          Paragraph(f"TANGGAL : {escape(display_date)}", centered), Spacer(1, 14)])
            if attachment_layout(attachment) == "side" and len(attachment["images"]) == 2:
                photo_cells = []
                for image_data in attachment["images"]:
                    ratio = image_data["width"] / image_data["height"]
                    scale = image_data["scale"] / 100
                    width = min(7.5 * cm, 19 * cm * ratio) * scale
                    photo_cells.append(Image(BytesIO(image_data["content"]), width=width, height=width / ratio))
                story.append(Table([photo_cells], colWidths=[8.2 * cm, 8.2 * cm]))
            else:
                for image_data in attachment["images"]:
                    ratio = image_data["width"] / image_data["height"]
                    scale = image_data["scale"] / 100
                    max_width, max_height = 16 * cm, (9 * cm if len(attachment["images"]) == 2 else 18 * cm)
                    width = min(max_width, max_height * ratio) * scale
                    picture_flowable = Image(BytesIO(image_data["content"]), width=width, height=width / ratio)
                    picture_flowable.hAlign = "CENTER"
                    story.extend([picture_flowable, Spacer(1, 8)])
        class ClassifiedCanvas(reportlab_canvas.Canvas):
            def showPage(self):
                if self.getPageNumber() >= 3:
                    self.saveState()
                    self.setFont(PDF_FONT_NAME, 10)
                    self.drawCentredString(folio[0] / 2, folio[1] - 1.1 * cm, "RAHASIA")
                    self.drawCentredString(folio[0] / 2, .75 * cm, "RAHASIA")
                    self.restoreState()
                super().showPage()

        pdf.build(story, onFirstPage=draw_cover, canvasmaker=ClassifiedCanvas)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=f"{export_name}.pdf", mimetype="application/pdf")
    return "Format tidak didukung", 400


@app.post("/lapinhar/export-preview-pdf")
@login_required
def export_lapinhar_preview_pdf():
    payload = request.get_json(silent=True) or {}
    preview_html = payload.get("html", "")
    report_number = str(payload.get("report_number", "lapinhar"))
    report_date = str(payload.get("report_date", ""))
    subject = str(payload.get("subject", ""))
    document_kind = str(payload.get("document_kind", "lapinhar"))
    try:
        report_id = int(payload["report_id"]) if payload.get("report_id") else None
    except (TypeError, ValueError):
        report_id = None
    accessible_report = accessible_lapinsus(report_id) if document_kind == "lapinsus" else accessible_lapinhar(report_id)
    if report_id and accessible_report is None:
        abort(404)
    availability_check = lapinsus_number_available if document_kind == "lapinsus" else lapinhar_number_available
    number_available, number_message = availability_check(report_number, str(payload.get("reservation_token", "")), report_id)
    if not number_available:
        return number_message, 409
    if not preview_html or len(preview_html) > 18_000_000:
        return "Pratinjau PDF tidak valid atau terlalu besar.", 400

    soup = BeautifulSoup(preview_html, "html.parser")
    for unsafe in soup.find_all(["script", "iframe", "object", "embed"]):
        unsafe.decompose()
    for element in soup.find_all(True):
        for attribute in list(element.attrs):
            if attribute.lower().startswith("on"):
                del element.attrs[attribute]

    stylesheet = (Path(__file__).parent / "static" / "css" / "style.css").read_text(encoding="utf-8")
    printable_html = f"""<!doctype html><html><head><meta charset="utf-8">
    <style>{stylesheet}</style></head><body class="printing-preview">
    <div class="app-shell"><div class="app-main"><main class="page-content">
    <form class="report-editor"><aside class="preview-column"><div class="preview-pages">
    {str(soup)}
    </div></aside></form></main></div></div></body></html>"""

    if not CHROME_EXECUTABLE.exists():
        return "Google Chrome tidak ditemukan. Atur CHROME_EXECUTABLE pada file .env.", 500
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(executable_path=str(CHROME_EXECUTABLE), headless=True)
            page = browser.new_page(viewport={"width": 816, "height": 1248})
            page.set_content(printable_html, wait_until="networkidle")
            page.emulate_media(media="print")
            pdf_bytes = page.pdf(
                width="8.5in", height="13in", print_background=True,
                prefer_css_page_size=True,
                margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
            )
            browser.close()
    except Exception as exc:
        app.logger.exception("Gagal merender PDF pratinjau")
        return f"Gagal membuat PDF: {exc}", 500

    export_name = lapinhar_export_basename(report_number, report_date, subject)
    return send_file(BytesIO(pdf_bytes), as_attachment=True,
                     download_name=f"{export_name}.pdf", mimetype="application/pdf")


@app.route("/users", methods=["GET", "POST"])
@admin_required
def users():
    if request.method == "POST":
        username = request.form.get("username", "").strip().lower()
        full_name = request.form.get("full_name", "").strip()
        password = request.form.get("password", "")
        role = request.form.get("role", "user")
        if not username or not full_name or len(password) < 6 or role not in {"admin", "user"}:
            flash("Lengkapi data. Kata sandi minimal 6 karakter.", "error")
        elif fetch_one("SELECT id FROM users WHERE username = %s", (username,)):
            flash("Username sudah digunakan.", "error")
        else:
            with get_db().cursor() as cursor:
                cursor.execute(
                    "INSERT INTO users (username, full_name, password_hash, role) VALUES (%s, %s, %s, %s)",
                    (username, full_name, generate_password_hash(password), role),
                )
            flash("Pengguna baru berhasil dibuat.", "success")
            return redirect(url_for("users"))
    return render_template("users.html", users=fetch_all("SELECT id, username, full_name, role, is_active, created_at FROM users ORDER BY created_at DESC"), active="users")


@app.post("/users/<int:user_id>/toggle-active")
@admin_required
def toggle_user_active(user_id):
    user = fetch_one("SELECT id,full_name,is_active FROM users WHERE id=%s", (user_id,))
    if user is None:
        abort(404)
    if user_id == session["user_id"]:
        flash("Anda tidak dapat menonaktifkan akun sendiri.", "error")
    else:
        new_status = 0 if user["is_active"] else 1
        with get_db().cursor() as cursor:
            cursor.execute("UPDATE users SET is_active=%s WHERE id=%s", (new_status, user_id))
        flash(f"Akun {user['full_name']} berhasil {'diaktifkan' if new_status else 'dinonaktifkan'}.", "success")
    return redirect(url_for("users"))


@app.post("/users/<int:user_id>/change-password")
@admin_required
def change_user_password(user_id):
    user = fetch_one("SELECT id,full_name FROM users WHERE id=%s", (user_id,))
    if user is None:
        abort(404)
    password = request.form.get("new_password", "")
    confirmation = request.form.get("confirm_password", "")
    if len(password) < 6:
        flash("Password baru minimal 6 karakter.", "error")
    elif password != confirmation:
        flash("Konfirmasi password tidak sama.", "error")
    else:
        with get_db().cursor() as cursor:
            cursor.execute("UPDATE users SET password_hash=%s WHERE id=%s",
                           (generate_password_hash(password), user_id))
        flash(f"Password {user['full_name']} berhasil diganti.", "success")
    return redirect(url_for("users"))


@app.route("/settings/integrations", methods=["GET", "POST"])
@login_required
def integration_settings():
    user_id = session["user_id"]
    inteliz = fetch_one("SELECT inteliz_username,updated_at,connected_at FROM inteliz_user_settings WHERE user_id=%s", (user_id,))
    whatsapp = fetch_one("SELECT contact_name,phone_number,updated_at FROM whatsapp_user_settings WHERE user_id=%s", (user_id,))
    sipede = fetch_one("SELECT sipede_username,updated_at,connected_at FROM sipede_user_settings WHERE user_id=%s", (user_id,))
    if request.method == "POST":
        integration = request.form.get("integration", "")
        if integration == "inteliz":
            username = request.form.get("inteliz_username", "").strip()
            password = request.form.get("inteliz_password", "")
            existing = fetch_one("SELECT inteliz_password_encrypted FROM inteliz_user_settings WHERE user_id=%s", (user_id,))
            if not username or (not password and not existing):
                flash("Username dan password Inteliz wajib diisi pada konfigurasi pertama.", "error")
            else:
                encrypted = inteliz_credential_cipher().encrypt(password.encode()).decode() if password else existing["inteliz_password_encrypted"]
                with get_db().cursor() as cursor:
                    cursor.execute("""INSERT INTO inteliz_user_settings (user_id,inteliz_username,inteliz_password_encrypted)
                    VALUES (%s,%s,%s) ON DUPLICATE KEY UPDATE inteliz_username=VALUES(inteliz_username),
                    inteliz_password_encrypted=VALUES(inteliz_password_encrypted),updated_at=CURRENT_TIMESTAMP""",
                                   (user_id, username, encrypted))
                flash("Konfigurasi Inteliz berhasil disimpan.", "success")
                return redirect(url_for("integration_settings", _anchor="inteliz"))
        elif integration == "whatsapp":
            name = request.form.get("contact_name", "").strip()
            phone = normalize_whatsapp_number(request.form.get("phone_number", ""))
            if not name or not re.fullmatch(r"\d{9,15}", phone):
                flash("Nama tujuan atau nomor WhatsApp tidak valid.", "error")
            else:
                with get_db().cursor() as cursor:
                    cursor.execute("""INSERT INTO whatsapp_user_settings (user_id,contact_name,phone_number)
                    VALUES (%s,%s,%s) ON DUPLICATE KEY UPDATE contact_name=VALUES(contact_name),
                    phone_number=VALUES(phone_number),updated_at=CURRENT_TIMESTAMP""", (user_id, name, phone))
                flash("Konfigurasi WhatsApp berhasil disimpan.", "success")
                return redirect(url_for("integration_settings", _anchor="whatsapp"))
        elif integration == "sipede":
            username = request.form.get("sipede_username", "").strip()
            password = request.form.get("sipede_password", "")
            existing = fetch_one("SELECT sipede_password_encrypted FROM sipede_user_settings WHERE user_id=%s", (user_id,))
            if not username or (not password and not existing):
                flash("Username dan password Sipede wajib diisi pada konfigurasi pertama.", "error")
            else:
                encrypted = inteliz_credential_cipher().encrypt(password.encode()).decode() if password else existing["sipede_password_encrypted"]
                with get_db().cursor() as cursor:
                    cursor.execute("""INSERT INTO sipede_user_settings (user_id,sipede_username,sipede_password_encrypted)
                    VALUES (%s,%s,%s) ON DUPLICATE KEY UPDATE sipede_username=VALUES(sipede_username),
                    sipede_password_encrypted=VALUES(sipede_password_encrypted),
                    session_data_encrypted=NULL,connected_at=NULL,updated_at=CURRENT_TIMESTAMP""",
                                    (user_id, username, encrypted))
                flash("Konfigurasi Sipede berhasil disimpan terenkripsi.", "success")
                return redirect(url_for("integration_settings", _anchor="sipede"))
        else:
            flash("Jenis integrasi tidak valid.", "error")
    return render_template("integration_settings.html", active="integration_settings",
                           inteliz=inteliz, whatsapp=whatsapp, sipede=sipede)


@app.route("/settings/inteliz", methods=["GET", "POST"])
@login_required
def inteliz_settings():
    setting = fetch_one(
        "SELECT inteliz_username, updated_at, connected_at FROM inteliz_user_settings WHERE user_id=%s",
        (session["user_id"],),
    )
    if request.method == "POST":
        inteliz_username = request.form.get("inteliz_username", "").strip()
        inteliz_password = request.form.get("inteliz_password", "")
        existing = fetch_one(
            "SELECT inteliz_password_encrypted FROM inteliz_user_settings WHERE user_id=%s",
            (session["user_id"],),
        )
        if not inteliz_username:
            flash("Username Inteliz wajib diisi.", "error")
        elif not inteliz_password and not existing:
            flash("Password Inteliz wajib diisi saat konfigurasi pertama.", "error")
        else:
            encrypted_password = (
                inteliz_credential_cipher().encrypt(inteliz_password.encode("utf-8")).decode("ascii")
                if inteliz_password else existing["inteliz_password_encrypted"]
            )
            with get_db().cursor() as cursor:
                cursor.execute(
                    """INSERT INTO inteliz_user_settings
                       (user_id, inteliz_username, inteliz_password_encrypted)
                       VALUES (%s, %s, %s)
                       ON DUPLICATE KEY UPDATE inteliz_username=VALUES(inteliz_username),
                       inteliz_password_encrypted=VALUES(inteliz_password_encrypted),
                       updated_at=CURRENT_TIMESTAMP""",
                    (session["user_id"], inteliz_username, encrypted_password),
                )
            flash("Konfigurasi Inteliz berhasil disimpan dengan aman.", "success")
            return redirect(url_for("inteliz_settings"))
    next_url = request.args.get("next", "")
    if not next_url.startswith("/") or next_url.startswith("//"):
        next_url = ""
    return render_template("inteliz_settings.html", setting=setting, active="inteliz_settings",
                           auto_connect=request.args.get("connect") == "1", next_url=next_url)


@app.route("/settings/whatsapp", methods=["GET", "POST"])
@login_required
def whatsapp_settings():
    setting = fetch_one(
        "SELECT contact_name, phone_number, updated_at FROM whatsapp_user_settings WHERE user_id=%s",
        (session["user_id"],),
    )
    next_url = request.values.get("next", "")
    if not next_url.startswith("/") or next_url.startswith("//"):
        next_url = ""
    if request.method == "POST":
        contact_name = request.form.get("contact_name", "").strip()
        phone_number = request.form.get("phone_number", "").strip()
        normalized = normalize_whatsapp_number(phone_number)
        if not contact_name:
            flash("Nama tujuan WhatsApp wajib diisi.", "error")
        elif not re.fullmatch(r"\d{9,15}", normalized):
            flash("Nomor WhatsApp tidak valid. Contoh: 081234567890.", "error")
        else:
            with get_db().cursor() as cursor:
                cursor.execute(
                    """INSERT INTO whatsapp_user_settings (user_id, contact_name, phone_number)
                       VALUES (%s, %s, %s)
                       ON DUPLICATE KEY UPDATE contact_name=VALUES(contact_name),
                       phone_number=VALUES(phone_number), updated_at=CURRENT_TIMESTAMP""",
                    (session["user_id"], contact_name, normalized),
                )
            flash("Konfigurasi WhatsApp berhasil disimpan.", "success")
            return redirect(next_url or url_for("whatsapp_settings"))
    return render_template("whatsapp_settings.html", setting=setting,
                           active="whatsapp_settings", next_url=next_url)


@app.post("/settings/inteliz/connect/start")
@login_required
def start_inteliz_connection():
    credentials = get_inteliz_credentials(session["user_id"])
    if not credentials:
        return jsonify(message="Simpan username dan password Inteliz terlebih dahulu."), 400
    auth_id = uuid.uuid4().hex
    with INTELIZ_AUTH_LOCK:
        for state in INTELIZ_AUTH_SESSIONS.values():
            if state.get("user_id") == session["user_id"] and state.get("status") not in {"success", "error"}:
                state["cancelled"] = True
        INTELIZ_AUTH_SESSIONS[auth_id] = {
            "user_id": session["user_id"], "status": "starting", "message": "Menyiapkan browser server…",
            "captcha": None, "cancelled": False, "updated_at": time.time(),
        }
    threading.Thread(target=run_inteliz_auth, args=(auth_id, session["user_id"], credentials),
                     daemon=True, name=f"inteliz-auth-{session['user_id']}").start()
    return jsonify(auth_id=auth_id, status="starting"), 202


def owned_inteliz_auth(auth_id):
    with INTELIZ_AUTH_LOCK:
        state = INTELIZ_AUTH_SESSIONS.get(auth_id)
        if not state or state.get("user_id") != session.get("user_id"):
            return None
        return state


@app.get("/settings/inteliz/connect/<auth_id>/status")
@login_required
def inteliz_connection_status(auth_id):
    state = owned_inteliz_auth(auth_id)
    if state is None:
        abort(404)
    return jsonify(status=state["status"], message=state.get("message", ""),
                   captcha=state.get("captcha"))


@app.post("/settings/inteliz/connect/<auth_id>/captcha")
@login_required
def submit_inteliz_captcha(auth_id):
    value = str((request.get_json(silent=True) or {}).get("captcha", "")).strip()
    with INTELIZ_AUTH_LOCK:
        state = INTELIZ_AUTH_SESSIONS.get(auth_id)
        if not state or state.get("user_id") != session["user_id"]:
            abort(404)
        if state.get("status") != "captcha" or not value or len(value) > 20:
            return jsonify(message="Kode CAPTCHA tidak valid."), 400
        state["captcha_input"] = value
        state["status"] = "loading"
        state["message"] = "Memverifikasi CAPTCHA…"
    return jsonify(message="CAPTCHA dikirim.")


@app.post("/settings/inteliz/connect/<auth_id>/otp")
@login_required
def submit_inteliz_otp(auth_id):
    value = str((request.get_json(silent=True) or {}).get("otp", "")).strip()
    with INTELIZ_AUTH_LOCK:
        state = INTELIZ_AUTH_SESSIONS.get(auth_id)
        if not state or state.get("user_id") != session["user_id"]:
            abort(404)
        if state.get("status") != "otp" or not value or len(value) > 20:
            return jsonify(message="Kode autentikator tidak valid."), 400
        state["otp_input"] = value
        state["status"] = "loading"
        state["message"] = "Memverifikasi kode autentikator…"
    return jsonify(message="Kode autentikator dikirim.")


@app.post("/settings/inteliz/connect/<auth_id>/cancel")
@login_required
def cancel_inteliz_connection(auth_id):
    with INTELIZ_AUTH_LOCK:
        state = INTELIZ_AUTH_SESSIONS.get(auth_id)
        if not state or state.get("user_id") != session["user_id"]:
            abort(404)
        state["cancelled"] = True
        state["status"] = "error"
        state["message"] = "Proses login dibatalkan."
    return jsonify(message="Proses login dibatalkan.")


@app.post("/settings/sipede/connect/start")
@login_required
def start_sipede_connection():
    credentials = get_sipede_credentials(session["user_id"])
    if not credentials:
        return jsonify(message="Simpan username dan password SIPede terlebih dahulu."), 400
    auth_id = uuid.uuid4().hex
    with SIPEDE_AUTH_LOCK:
        for state in SIPEDE_AUTH_SESSIONS.values():
            if state.get("user_id") == session["user_id"] and state.get("status") not in {"success", "error"}:
                state["cancelled"] = True
        SIPEDE_AUTH_SESSIONS[auth_id] = {
            "user_id": session["user_id"],
            "status": "starting",
            "message": "Menyiapkan browser server…",
            "captcha": None,
            "cancelled": False,
            "updated_at": time.time(),
        }
    threading.Thread(
        target=run_sipede_auth,
        args=(auth_id, session["user_id"], credentials),
        daemon=True,
        name=f"sipede-auth-{session['user_id']}",
    ).start()
    return jsonify(auth_id=auth_id, status="starting"), 202


def owned_sipede_auth(auth_id):
    with SIPEDE_AUTH_LOCK:
        state = SIPEDE_AUTH_SESSIONS.get(auth_id)
        if not state or state.get("user_id") != session.get("user_id"):
            return None
        return state


@app.get("/settings/sipede/connect/<auth_id>/status")
@login_required
def sipede_connection_status(auth_id):
    state = owned_sipede_auth(auth_id)
    if state is None:
        abort(404)
    return jsonify(status=state["status"], message=state.get("message", ""),
                   captcha=state.get("captcha"))


@app.post("/settings/sipede/connect/<auth_id>/captcha")
@login_required
def submit_sipede_captcha(auth_id):
    value = str((request.get_json(silent=True) or {}).get("captcha", "")).strip()
    with SIPEDE_AUTH_LOCK:
        state = SIPEDE_AUTH_SESSIONS.get(auth_id)
        if not state or state.get("user_id") != session["user_id"]:
            abort(404)
        if state.get("status") != "captcha" or not value or len(value) > 20:
            return jsonify(message="Kode CAPTCHA tidak valid."), 400
        state["captcha_input"] = value
        state["status"] = "loading"
        state["message"] = "Memverifikasi CAPTCHA SIPede…"
    return jsonify(message="CAPTCHA dikirim.")


@app.post("/settings/sipede/connect/<auth_id>/cancel")
@login_required
def cancel_sipede_connection(auth_id):
    with SIPEDE_AUTH_LOCK:
        state = SIPEDE_AUTH_SESSIONS.get(auth_id)
        if not state or state.get("user_id") != session["user_id"]:
            abort(404)
        state["cancelled"] = True
        state["status"] = "error"
        state["message"] = "Proses login SIPede dibatalkan."
    return jsonify(message="Proses login SIPede dibatalkan.")


@app.route("/settings/signatories", methods=["GET", "POST"])
@admin_required
def signatory_settings():
    if request.method == "POST":
        code = request.form.get("position_code", "")
        if code not in {"kajari", "kasi_intel", "kasubsi_1", "kasubsi_2"}:
            flash("Jenis penandatangan tidak valid.", "error")
        else:
            position_name = request.form.get("position_name", "").strip()
            full_name = request.form.get("full_name", "").strip()
            rank_nip = request.form.get("rank_nip", "").strip()
            use_tte = 1 if code == "kajari" and request.form.get("use_tte") == "1" else 0
            if not position_name or not full_name or not rank_nip:
                flash("Seluruh data penandatangan wajib diisi.", "error")
            else:
                with get_db().cursor() as cursor:
                    cursor.execute(
                        """UPDATE signatories SET position_name = %s, full_name = %s, rank_nip = %s, use_tte = %s,
                           updated_at = CURRENT_TIMESTAMP WHERE position_code = %s""",
                        (position_name, full_name, rank_nip, use_tte, code),
                    )
                flash("Data penandatangan berhasil diperbarui.", "success")
                return redirect(url_for("signatory_settings"))
    rows = fetch_all("SELECT * FROM signatories ORDER BY FIELD(position_code, 'kajari', 'kasi_intel', 'kasubsi_1', 'kasubsi_2')")
    return render_template("signatory_settings.html", signatories=rows, active="signatories")


@app.route("/settings/organization", methods=["GET", "POST"])
@admin_required
def organization_settings():
    if request.method == "POST":
        organization_name = request.form.get("organization_name", "").strip()
        institution_code = request.form.get("institution_code", "").strip()
        address = request.form.get("address", "").strip()
        phone = request.form.get("phone", "").strip()
        website = request.form.get("website", "").strip()
        if not all((organization_name, institution_code, address, phone, website)):
            flash("Seluruh data instansi wajib diisi.", "error")
        else:
            with get_db().cursor() as cursor:
                cursor.execute(
                    """INSERT INTO organization_settings (id, organization_name, institution_code, address, phone, website)
                       VALUES (1, %s, %s, %s, %s, %s)
                       ON DUPLICATE KEY UPDATE organization_name = VALUES(organization_name),
                       institution_code = VALUES(institution_code), address = VALUES(address),
                       phone = VALUES(phone), website = VALUES(website),
                       updated_at = CURRENT_TIMESTAMP""",
                    (organization_name, institution_code, address, phone, website),
                )
            flash("Data instansi berhasil diperbarui.", "success")
            return redirect(url_for("organization_settings"))
    organization = fetch_one("SELECT * FROM organization_settings WHERE id = 1")
    return render_template("organization_settings.html", organization=organization, active="organization")


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


initialize_database()

if __name__ == "__main__":
    app.run(
        host=os.getenv("FLASK_HOST", "0.0.0.0"),
        port=int(os.getenv("FLASK_PORT", "5000")),
        debug=os.getenv("FLASK_DEBUG", "0") == "1",
    )
