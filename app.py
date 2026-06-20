# app.py  — FINAL merged & fixed for TalentTrack
# - Keeps your route layout and templates.
# - Robust email sending: prefers Mailtrap API (if MAILTRAP_API_TOKEN present),
#   falls back to Mailtrap SMTP sandbox. Helpful debug logs included.
# - Safer DB init: creates tables and adds missing columns if required.
# - Scheduler wrapped in try/except so app still runs when scheduler fails.
# - Improved parsing and resilient Adzuna fetching.

import os
import re
import csv
import sqlite3
import random
import json
import traceback
from datetime import datetime, timedelta
from flask import (Flask, render_template, request, redirect, url_for, flash,
                   session, send_from_directory)
from werkzeug.utils import secure_filename
import requests
import fitz      # PyMuPDF
import docx
import pandas as pd

# scheduler + tz
from apscheduler.schedulers.background import BackgroundScheduler
from pytz import timezone

# email helpers
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ---------------- App setup ----------------
app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = os.environ.get("TT_SECRET", "talenttrack-dev-secret")

# ---------------- Paths & files ----------------
BASE_DIR = os.getcwd()
BASE_UPLOAD = os.path.join(BASE_DIR, "uploads")
RESUME_UPLOAD = os.path.join(BASE_UPLOAD, "resumes")
PROOF_TECH = os.path.join(BASE_UPLOAD, "proofs", "tech")
PROOF_NONTECH = os.path.join(BASE_UPLOAD, "proofs", "nontech")
USERS_CSV = os.path.join(BASE_DIR, "users.csv")
APPLIED_JOBS_CSV = os.path.join(BASE_DIR, "applied_jobs.csv")
SKILLS_CSV = os.path.join(BASE_DIR, "skills.csv")
DB_PATH = os.path.join(BASE_DIR, "interviews.db")

for d in [BASE_UPLOAD, RESUME_UPLOAD, PROOF_TECH, PROOF_NONTECH]:
    os.makedirs(d, exist_ok=True)

# Ensure CSVs exist
if not os.path.exists(USERS_CSV):
    with open(USERS_CSV, "w", newline="") as f:
        csv.writer(f).writerow(["name", "email", "interview_date"])

if not os.path.exists(APPLIED_JOBS_CSV):
    with open(APPLIED_JOBS_CSV, "w", newline="") as f:
        csv.writer(f).writerow(["name", "email", "job_title", "company", "link", "interview_date"])

# Load skill bank (expected columns: skill,category)
if os.path.exists(SKILLS_CSV):
    try:
        skill_bank = pd.read_csv(SKILLS_CSV)
    except Exception:
        skill_bank = pd.DataFrame([{"skill":"Python","category":"Technical"},
                                   {"skill":"Communication","category":"Soft"},
                                   {"skill":"Dance","category":"Non-Technical"}])
else:
    skill_bank = pd.DataFrame([{"skill":"Python","category":"Technical"},
                               {"skill":"Communication","category":"Soft"},
                               {"skill":"Dance","category":"Non-Technical"}])

# ---------------- DB (interviews) ----------------
def db_conn():
    conn = sqlite3.connect(DB_PATH, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = db_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS interviews (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_name TEXT,
            user_email TEXT,
            job_role TEXT,
            company TEXT,
            interview_datetime TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_name TEXT,
            user_email TEXT,
            job_role TEXT,
            company TEXT,
            interview_datetime TEXT,
            moved_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()

    # Add flags if missing
    def column_exists(table, col):
        c.execute(f"PRAGMA table_info({table})")
        cols = [r[1] for r in c.fetchall()]
        return col in cols

    for col in ("reminder_24_sent", "reminder_1h_sent", "feedback_sent"):
        if not column_exists("interviews", col):
            try:
                c.execute(f"ALTER TABLE interviews ADD COLUMN {col} INTEGER DEFAULT 0")
            except Exception:
                app.logger.debug("Could not add column %s (may already exist)", col)

    conn.commit()
    conn.close()

init_db()

# ---------------- Text extraction & skill extraction ----------------
ALLOWED_EXT = {"pdf", "docx"}
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT

def extract_text_from_file(filepath):
    text = ""
    try:
        if filepath.lower().endswith(".pdf"):
            doc = fitz.open(filepath)
            for p in doc:
                text += p.get_text("text") + "\n"
        elif filepath.lower().endswith(".docx"):
            docx_obj = docx.Document(filepath)
            for para in docx_obj.paragraphs:
                text += para.text + "\n"
    except Exception as e:
        app.logger.warning("extract_text error: %s", e)
    return text

def normalize_text(t):
    t = t or ""
    t = t.replace("\n", " ").lower()
    t = re.sub(r"\s+", " ", t).strip()
    t = "".join(ch for ch in t if ch.isalnum() or ch.isspace() or ch in {"+", ".", "-", "_"})
    return t

STOPWORDS = {
    "and","or","the","a","an","to","in","on","of","with","by","for","as","at","from","is","are","be","this","that"
}
def tokens(text):
    text = normalize_text(text)
    toks = [t for t in re.split(r"\s+", text) if t and len(t) > 1 and t not in STOPWORDS]
    return toks

# Hobby/non-technical extraction (targeted)
HOBBY_KEYWORDS = [
    "dance","dancing","singing","music","guitar","piano","badminton","cricket","football","sports",
    "yoga","gym","anchoring","hosting","public speaking","nss","volunteering","community service",
    "teaching","art","painting","drawing","photography","blogging","content creation","event management",
    "writing","reading","travelling"
]

def extract_non_technical_keywords(text):
    if not text:
        return []
    text_low = text.lower()
    sections = ""
    # try to find sections titled Hobbies/Interests/Extra Curricular/Activities
    patterns = [r"hobbies?:\s*(.+?)(?:\n\n|\Z)", r"interests?:\s*(.+?)(?:\n\n|\Z)", r"extra[- ]?curricular[:\s]*(.+?)(?:\n\n|\Z)"]
    for p in patterns:
        m = re.search(p, text_low, flags=re.DOTALL)
        if m:
            sections += " " + m.group(1)
    if not sections:
        # fallback lines containing these verbs/words
        for line in text_low.splitlines():
            if any(k in line for k in ["hobby", "interest", "enjoy", "passion", "like"]):
                sections += " " + line

    found = set()
    for kw in HOBBY_KEYWORDS:
        if re.search(rf"\b{re.escape(kw)}\b", sections):
            found.add(kw.title())

    return sorted(found)

def clean_list(items):
    cleaned = []
    for x in items:
        if not x:
            continue
        x = str(x).strip()
        if len(x) < 2:
            continue
        if any(ch.isdigit() for ch in x):
            continue
        x = re.sub(r"[^a-zA-Z +\-\._]", "", x).strip()
        if len(x) < 2:
            continue
        cleaned.append(x)
    return sorted(set(cleaned))

def extract_skills_from_text(text):
    if not text:
        return [], [], []
    text_low = text.lower()

    technical = set()
    soft = set()
    nontechnical = set()

    # from skills.csv (exact word matching)
    for _, row in skill_bank.iterrows():
        skill = str(row.get("skill","")).strip()
        cat = str(row.get("category","")).lower().strip()
        if not skill:
            continue
        # only match whole-word occurrences to reduce false positives
        try:
            if re.search(rf"\b{re.escape(skill.lower())}\b", text_low):
                if "technical" in cat:
                    technical.add(skill)
                elif "soft" in cat:
                    soft.add(skill)
                else:
                    nontechnical.add(skill)
        except re.error:
            # in case the skill contains regex-invalid chars
            if skill.lower() in text_low:
                if "technical" in cat:
                    technical.add(skill)
                elif "soft" in cat:
                    soft.add(skill)
                else:
                    nontechnical.add(skill)

    # soft keywords
    SOFT_KEYWORDS = ["communication","teamwork","leadership","creativity","adaptability","problem solving","listening","presentation"]
    for s in SOFT_KEYWORDS:
        if re.search(rf"\b{s}\b", text_low):
            soft.add(s.title())

    # hobby detection
    hobby_found = extract_non_technical_keywords(text)
    for h in hobby_found:
        if not any(h.lower() == t.lower() for t in technical):
            nontechnical.add(h)

    return clean_list(technical), clean_list(soft), clean_list(nontechnical)

# ---------------- Save files helper ----------------
def save_multiple_files(files, target_dir):
    saved = []
    for f in files:
        if not f or not getattr(f, "filename", "").strip():
            continue
        name = secure_filename(f.filename)
        path = os.path.join(target_dir, name)
        base, ext = os.path.splitext(name)
        i = 1
        while os.path.exists(path):
            name = f"{base}_{i}{ext}"
            path = os.path.join(target_dir, name)
            i += 1
        f.save(path)
        saved.append(path)
    return saved

# ---------------- Scoreboard ----------------
WEIGHTS = {
    "tech_skill": 2.0,
    "nontech_skill": 2.0,
    "tech_proof": 3.0,
    "nontech_proof": 3.0,
    "soft_bonus_for_tech": 0.5,
    "soft_bonus_for_nontech": 0.5,
}

def compute_scoreboard(tech_skills, nontech_skills, soft_skills, tech_proofs, nontech_proofs):
    tech_score = WEIGHTS["tech_skill"]*len(tech_skills) + WEIGHTS["tech_proof"]*len(tech_proofs) + WEIGHTS["soft_bonus_for_tech"]*len(soft_skills)
    nontech_score = WEIGHTS["nontech_skill"]*len(nontech_skills) + WEIGHTS["nontech_proof"]*len(nontech_proofs) + WEIGHTS["soft_bonus_for_nontech"]*len(soft_skills)
    total = tech_score + nontech_score
    tech_pct = round((tech_score/total)*100,1) if total>0 else 0
    nontech_pct = round((nontech_score/total)*100,1) if total>0 else 0
    dominant = "Tech" if tech_score > nontech_score else ("Non-Tech" if nontech_score > tech_score else "Tie")
    return {"tech_score": tech_score, "nontech_score": nontech_score, "tech_pct": tech_pct, "nontech_pct": nontech_pct, "dominant": dominant}

# ---------------- Adzuna job fetcher ----------------
ADZUNA_APP_ID = os.environ.get("ADZUNA_APP_ID", "46131009")
ADZUNA_APP_KEY = os.environ.get("ADZUNA_APP_KEY", "b83d049c04d82af9cb69b82c4d86e89a")
ADZUNA_COUNTRY = os.environ.get("ADZUNA_COUNTRY", "in")

def fetch_adzuna_jobs(query="developer", location="India", limit=30):
    try:
        url = f"https://api.adzuna.com/v1/api/jobs/{ADZUNA_COUNTRY}/search/1"
        params = {
            "app_id": ADZUNA_APP_ID,
            "app_key": ADZUNA_APP_KEY,
            "what": query,
            "where": location,
            "results_per_page": limit
        }
        resp = requests.get(url, params=params, timeout=8)
        data = resp.json()
        jobs = []
        for j in data.get("results", []):
            city = j.get("location", {}).get("display_name", "") or ""
            city = city.replace("India", "").strip(" ,")
            jobs.append({
                "title": j.get("title","Unknown"),
                "company": j.get("company", {}).get("display_name","Unknown"),
                "city": city,
                "salary": j.get("salary_min") or 0,
                "work_type": j.get("contract_time") or "",
                "description": j.get("description",""),
                "link": j.get("redirect_url", "#"),
                "created": j.get("created", "")
            })
        return jobs
    except Exception as e:
        app.logger.warning("Adzuna fetch error: %s", e)
        return []

# ---------------- Job ranking ----------------
def score_job_against_user(job, resume_text, user_skills):
    jtext = " ".join([job.get("title",""), job.get("description","")]).lower()
    jtokens = tokens(jtext)
    resume_tokens = tokens(resume_text.lower())

    skill_match = 0
    for s in user_skills:
        if not s:
            continue
        if s.lower() in jtext:
            skill_match += 1
    skill_match_score = min(1.0, skill_match / max(1, len(user_skills))) if user_skills else 0

    set_j = set(jtokens)
    set_r = set(resume_tokens)
    if set_j or set_r:
        inter = set_j.intersection(set_r)
        union = set_j.union(set_r)
        resume_overlap_score = len(inter) / max(1, len(union))
    else:
        resume_overlap_score = 0

    title = job.get("title","").lower()
    title_tokens = set(tokens(title))
    title_bonus = 1.0 if len(title_tokens.intersection(set(resume_tokens)))>0 else 0.0

    recency_bonus = 0.0
    created = job.get("created","")
    try:
        if created:
            dt = datetime.fromisoformat(created.replace("Z",""))
            days_old = (datetime.now() - dt).days
            if days_old <= 7:
                recency_bonus = 0.05
            elif days_old <= 30:
                recency_bonus = 0.02
    except Exception:
        recency_bonus = 0.0

    score = (0.55 * skill_match_score) + (0.35 * resume_overlap_score) + (0.08 * title_bonus) + recency_bonus
    score = max(0.0, min(1.0, score))
    return int(round(score * 100))

# ---------------- Mailtrap API / SMTP config ----------------
# Preferred: MAILTRAP_API_TOKEN (Mailtrap API). Fallback to SMTP sandbox credentials below.
MAILTRAP_API_TOKEN = os.environ.get("MAILTRAP_API_TOKEN", "").strip()

# SMTP sandbox fallback (you used this earlier)
MAILTRAP_SMTP_HOST = "sandbox.smtp.mailtrap.io"
MAILTRAP_SMTP_PORT = 587
MAILTRAP_SMTP_USERNAME = "6a997ab0bff2eb"
MAILTRAP_SMTP_PASSWORD = "45d202c92e7493"
MAIL_FROM = "TalentTrack <noreply@talenttrack.demo>"

def send_email_smtp(to_email, subject, html_body):
    try:
        msg = MIMEMultipart()
        msg["From"] = MAIL_FROM
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(html_body, "html"))

        server = smtplib.SMTP(MAILTRAP_SMTP_HOST, MAILTRAP_SMTP_PORT, timeout=10)
        server.starttls()
        server.login(MAILTRAP_SMTP_USERNAME, MAILTRAP_SMTP_PASSWORD)
        server.sendmail(MAIL_FROM, [to_email], msg.as_string())
        server.quit()
        app.logger.info("SMTP email sent to %s", to_email)
        return True
    except Exception as e:
        app.logger.warning("SMTP send failed: %s", e)
        return False

def send_email_mailtrap_api(to_email, subject, html_body):
    """
    Send via Mailtrap API (preferred). Requires MAILTRAP_API_TOKEN env var.
    API docs: https://mailtrap.docs.apiary.io/ (sandbox & send endpoints)
    Note: Mailtrap requires your account to be provisioned / approved.
    """
    if not MAILTRAP_API_TOKEN:
        return False, "no_api_token"

    try:
        # Using Mailtrap "Send" API endpoint. Adjust if your account uses different endpoint.
        url = "https://send.api.mailtrap.io/api/send"
        payload = {
            "from": {"email": MAIL_FROM.split("<")[-1].strip(">").strip() if "<" in MAIL_FROM else MAIL_FROM, "name": MAIL_FROM.split("<")[0].strip()},
            "to": [{"email": to_email}],
            "subject": subject,
            "html": html_body
        }
        headers = {
            "Api-Token": MAILTRAP_API_TOKEN,
            "Content-Type": "application/json"
        }
        resp = requests.post(url, headers=headers, json=payload, timeout=10)
        if resp.status_code in (200, 201):
            app.logger.info("Mailtrap API send OK to %s", to_email)
            return True, resp.text
        else:
            app.logger.warning("Mailtrap API send failed: %s %s", resp.status_code, resp.text)
            return False, f"{resp.status_code} {resp.text}"
    except Exception as e:
        app.logger.warning("Mailtrap API exception: %s", e)
        return False, str(e)

def send_email(to_email, subject, html_body):
    """
    Attempt Mailtrap API first (if token present), otherwise fallback to SMTP sandbox.
    Returns True/False.
    """
    if MAILTRAP_API_TOKEN:
        ok, info = send_email_mailtrap_api(to_email, subject, html_body)
        if ok:
            return True
        else:
            app.logger.info("Mailtrap API failed (%s). Falling back to SMTP: %s", info, to_email)
            # fall through to SMTP
    # fallback SMTP
    return send_email_smtp(to_email, subject, html_body)

# ---------------- Scheduler job (reminders & feedback) ----------------
def parse_datetime_flexible(s):
    if not s:
        return None
    if isinstance(s, datetime):
        return s
    s2 = s.replace("T", " ").strip()
    fmts = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M:%S.%f", "%d-%m-%Y %H:%M", "%d/%m/%Y %H:%M"]
    for f in fmts:
        try:
            return datetime.strptime(s2, f)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(s2)
    except Exception:
        return None

def scheduler_job_check():
    try:
        conn = db_conn()
        c = conn.cursor()
        now = datetime.now()
        c.execute("SELECT * FROM interviews")
        rows = c.fetchall()
        for r in rows:
            try:
                dt = parse_datetime_flexible(r["interview_datetime"])
                if not dt:
                    continue

                # 24 hour reminder window
                delta_24 = dt - now - timedelta(hours=24)
                if not r.get("reminder_24_sent", 0):
                    if abs(delta_24.total_seconds()) <= 70:
                        subject = f"Reminder: Interview in 24 hours — {r['job_role']}"
                        body = f"<p>Hi {r['user_name'] or 'Candidate'},</p><p>Your interview for <strong>{r['job_role']}</strong> at <strong>{r['company']}</strong> is scheduled on <strong>{dt.strftime('%Y-%m-%d %H:%M')}</strong>.</p><p>This is a 24-hour reminder. Good luck!</p>"
                        if send_email(r['user_email'], subject, body):
                            c.execute("UPDATE interviews SET reminder_24_sent=1 WHERE id=?", (r["id"],))
                            conn.commit()

                # 1 hour reminder
                delta_1h = dt - now - timedelta(hours=1)
                if not r.get("reminder_1h_sent", 0):
                    if abs(delta_1h.total_seconds()) <= 70:
                        subject = f"Reminder: Interview in 1 hour — {r['job_role']}"
                        body = f"<p>Hi {r['user_name'] or 'Candidate'},</p><p>Your interview for <strong>{r['job_role']}</strong> at <strong>{r['company']}</strong> is scheduled on <strong>{dt.strftime('%Y-%m-%d %H:%M')}</strong>.</p><p>This is a 1-hour reminder. Best of luck!</p>"
                        if send_email(r['user_email'], subject, body):
                            c.execute("UPDATE interviews SET reminder_1h_sent=1 WHERE id=?", (r["id"],))
                            conn.commit()

                # feedback send & move to history
                if dt < now and not r.get("feedback_sent", 0):
                    subject = f"Interview Feedback: {r['job_role']}"
                    body = f"<p>Hi {r['user_name'] or 'Candidate'},</p><p>Thanks for interviewing for <strong>{r['job_role']}</strong> at <strong>{r['company']}</strong> on {dt.strftime('%Y-%m-%d %H:%M')}.</p><p>Please reply with any feedback or questions.</p>"
                    ok = send_email(r['user_email'], subject, body)
                    try:
                        # mark feedback_sent and move to history, delete original
                        c.execute("UPDATE interviews SET feedback_sent=1 WHERE id=?", (r["id"],))
                        c.execute("""INSERT INTO history (user_name,user_email,job_role,company,interview_datetime)
                                     VALUES (?,?,?,?,?)""", (r["user_name"], r["user_email"], r["job_role"], r["company"], r["interview_datetime"]))
                        c.execute("DELETE FROM interviews WHERE id=?", (r["id"],))
                        conn.commit()
                    except Exception:
                        conn.rollback()
            except Exception:
                app.logger.debug("Error in scheduler loop for row: %s\n%s", r["id"], traceback.format_exc())
        conn.close()
    except Exception:
        app.logger.warning("Scheduler job failed: %s", traceback.format_exc())

# Start scheduler (safe)
scheduler = None
try:
    scheduler = BackgroundScheduler(timezone=timezone("Asia/Kolkata"))
    scheduler.add_job(func=scheduler_job_check, trigger="interval", minutes=1, id="talenttrack_scheduler", replace_existing=True)
    scheduler.start()
    app.logger.info("APScheduler started.")
except Exception as e:
    app.logger.warning("Scheduler failed to start: %s", e)
    scheduler = None

# ---------------- Routes ----------------
@app.route("/")
def home():
    return render_template("home.html")

@app.route("/upload_flow", methods=["GET","POST"])
def upload_flow():
    if request.method == "GET":
        return render_template("upload_flow.html", section="upload")
    step = request.form.get("step")
    # STEP: upload resume
    if step == "upload_resume":
        file = request.files.get("resume")
        name = request.form.get("name")
        email = request.form.get("email")
        if not file or not file.filename:
            flash("Please upload resume (.pdf/.docx)", "warning")
            return redirect(url_for("upload_flow"))
        if not allowed_file(file.filename):
            flash("Only .pdf and .docx allowed", "danger")
            return redirect(url_for("upload_flow"))

        filename = secure_filename(file.filename)
        path = os.path.join(RESUME_UPLOAD, filename)
        base, ext = os.path.splitext(filename)
        i = 1
        while os.path.exists(path):
            filename = f"{base}_{i}{ext}"
            path = os.path.join(RESUME_UPLOAD, filename)
            i += 1
        file.save(path)

        text = extract_text_from_file(path)
        tech, soft, nontech = extract_skills_from_text(text)

        with open(USERS_CSV, "a", newline="") as f:
            csv.writer(f).writerow([name, email, ""])

        session["name"] = name
        session["email"] = email
        session["resume_path"] = path
        session["resume_text"] = text
        session["tech"] = tech
        session["soft"] = soft
        session["nontech"] = nontech

        return render_template("upload_flow.html", section="skills", tech=tech, soft=soft, nontech=nontech)

    # STEP: upload proofs
    if step == "upload_proofs":
        tech_files = request.files.getlist("proof_tech")
        nontech_files = request.files.getlist("proof_nontech")
        saved_tech = save_multiple_files(tech_files, PROOF_TECH)
        saved_nontech = save_multiple_files(nontech_files, PROOF_NONTECH)
        session["proof_tech_files"] = saved_tech
        session["proof_nontech_files"] = saved_nontech

        board = compute_scoreboard(session.get("tech", []), session.get("nontech", []), session.get("soft", []), saved_tech, saved_nontech)
        session["board"] = board
        return render_template("upload_flow.html", section="scoreboard", board=board)

    # STEP: job_type
    if step == "job_type":
        track_choice = request.form.get("track_choice", "Both")
        session["track_choice"] = track_choice
        return render_template("upload_flow.html", section="mode", tech=session.get("tech",[]), soft=session.get("soft",[]), nontech=session.get("nontech",[]))

    # STEP: choose_mode
    if step == "choose_mode":
        mode = request.form.get("mode")
        session["mode_choice"] = mode
        if mode == "Core":
            session["mastered_skill"] = request.form.get("mastered_skill")
        else:
            session["selected_skills"] = request.form.getlist("selected_skills")
        return redirect(url_for("job_recommendations"))

    return redirect(url_for("upload_flow"))

@app.route("/dashboard", methods=["GET","POST"])
def dashboard():
    domain_skills_map = {
        "Data Science": ["Python", "Pandas", "NumPy", "Machine Learning", "Statistics", "SQL"],
        "Web Development": ["HTML", "CSS", "JavaScript", "React", "Node.js"],
        "Mobile Development": ["Flutter", "Java", "Kotlin", "Swift"],
        "Cloud & DevOps": ["AWS", "Docker", "Linux", "Kubernetes"],
        "Cyber Security": ["Networking", "Ethical Hacking", "Linux", "Forensics"],
        "AI & Deep Learning": ["TensorFlow", "PyTorch", "Python", "ML"],
        "Business Analyst": ["Excel", "Power BI", "SQL", "Analytics"],
        "UI/UX Design": ["Figma", "Wireframing", "Prototyping"]
    }

    user_skills = set(session.get("tech", []) + session.get("soft", []) + session.get("nontech", []))
    domain_scores = {d: sum(1 for s in skills if s.lower() in (x.lower() for x in user_skills)) for d, skills in domain_skills_map.items()}
    suggested_domain = max(domain_scores, key=domain_scores.get) if domain_scores else "Data Science"

    if request.method == "POST":
        session["domain"] = request.form.get("domain")

    domain_query = session.get("domain", suggested_domain)
    jobs = fetch_adzuna_jobs(query=domain_query, location="India", limit=50)
    required = domain_skills_map.get(domain_query, [])
    missing_skills = [s for s in required if s.lower() not in (x.lower() for x in user_skills)]
    skill_gap = [len(required) - len(missing_skills), len(missing_skills)] 
    # city counts
    city_counts = {}
    invalid_lower = {"india", "in", "remote", "work from home", "hybrid"}
    for j in jobs:
        raw = (j.get("city") or "").strip()
        if not raw:
            continue
        low = raw.lower()
        if low in invalid_lower:
            continue
        city = raw.split(",")[0].strip() if "," in raw else raw
        if len(city) < 2:
            continue
        city_counts[city] = city_counts.get(city, 0) + 1

    months = ["Jan","Feb","Mar","Apr","May","Jun"]
    demand_trend = [random.randint(40,80) for _ in months]
    salary_trend = [round(random.uniform(3,8),1) for _ in months]
    radar_labels = list(domain_scores.keys())
    radar_values = list(domain_scores.values())
    city_labels = list(city_counts.keys())[:8]
    city_values = list(city_counts.values())[:8]

    courses = {"Python":"https://www.coursera.org/learn/python", "Pandas":"https://www.coursera.org/learn/data-analysis-with-python"}

    return render_template("dashboard.html",
                       domain_query=domain_query,
                       domain_list=list(domain_skills_map.keys()),
                       suggested_domain=suggested_domain,
                       missing_skills=missing_skills,
                       skill_gap=skill_gap,
                       courses=courses,
                       jobs_preview=jobs[:10],
                       months=months,
                       demand_trend=demand_trend,
                       salary_trend=salary_trend,
                       radar_labels=radar_labels,
                       radar_values=radar_values,
                       city_labels=city_labels,
                       city_values=city_values)

@app.route("/job_recommendations", methods=["GET"])
def job_recommendations():
    mode_choice = session.get("mode_choice", "Core")
    track_choice = session.get("track_choice", "Both")
    if mode_choice == "Core":
        search_skill = session.get("mastered_skill", "developer")
        user_skills = [session.get("mastered_skill")] if session.get("mastered_skill") else session.get("tech",[])
    else:
        selected = session.get("selected_skills", []) or session.get("tech",[])
        user_skills = selected if selected else session.get("tech",[])
        search_skill = selected[0] if selected else "developer"

    jobs = fetch_adzuna_jobs(query=search_skill, location="India", limit=60)
    resume_text = session.get("resume_text", "")

    scored = []
    for j in jobs:
        j["score"] = score_job_against_user(j, resume_text, user_skills)
        scored.append(j)
    scored = sorted(scored, key=lambda x: x.get("score",0), reverse=True)

    cities = sorted({j.get("city") for j in scored if j.get("city")})
    work_types = sorted({j.get("work_type") for j in scored if j.get("work_type")})

    city_filter = request.args.get("city") or ""
    work_filter = request.args.get("work_type") or ""
    try:
        min_salary = int(request.args.get("min_salary") or 0)
    except:
        min_salary = 0
    try:
        max_salary = int(request.args.get("max_salary") or 10000000)
    except:
        max_salary = 10000000

    filtered = []
    for j in scored:
        if city_filter and j.get("city") != city_filter:
            continue
        if work_filter and j.get("work_type") != work_filter:
            continue
        sal = j.get("salary") or 0
        if isinstance(sal,(int,float)) and not (min_salary <= sal <= max_salary):
            continue
        if track_choice == "Tech":
            tech_tokens = set(t.lower() for t in session.get("tech",[]))
            title_desc = (j.get("title","")+" "+ j.get("description","")).lower()
            if not any(t in title_desc for t in tech_tokens):
                continue
        if track_choice == "Non-Tech":
            non_tokens = set(t.lower() for t in session.get("nontech",[]))
            title_desc = (j.get("title","")+" "+ j.get("description","")).lower()
            if not any(t in title_desc for t in non_tokens):
                continue
        filtered.append(j)

    return render_template("job_recommendations_merged.html",
                           jobs=filtered,
                           cities=cities,
                           work_types=work_types,
                           min_salary=min_salary,
                           max_salary=max_salary)

@app.route("/mark_applied", methods=["POST"])
def mark_applied():
    name = session.get("name", "Unknown")
    email = session.get("email", "unknown@example.com")
    job_title = request.form.get("job_title")
    company = request.form.get("company")
    link = request.form.get("link")
    interview_date = request.form.get("interview_date", "")

    # 1️⃣ Save to applied_jobs.csv (existing behavior)
    with open(APPLIED_JOBS_CSV, "a", newline="") as f:
        csv.writer(f).writerow([name, email, job_title, company, link, interview_date])

    # 2️⃣ ALSO save interview to SQLite DB
    if interview_date.strip():
        conn = db_conn()
        c = conn.cursor()
        dt = interview_date.replace("T", " ")
        c.execute("""
            INSERT INTO interviews (user_name, user_email, job_role, company, interview_datetime)
            VALUES (?,?,?,?,?)
        """, (name, email, job_title, company, dt))
        conn.commit()
        conn.close()

    flash("Marked as applied!", "success")
    return redirect(request.referrer or url_for("job_recommendations"))

@app.route("/applied_jobs")
def applied_jobs():
    rows = []
    with open(APPLIED_JOBS_CSV, "r", newline="") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append(r)
    return render_template("applied_jobs.html", jobs=rows)

@app.route("/delete_applied_job", methods=["POST"])
def delete_applied_job():
    idx = int(request.form.get("index", -1))
    with open(APPLIED_JOBS_CSV, "r", newline="") as f:
        rows = list(csv.reader(f))
    if idx >= 0 and idx < len(rows)-1:
        new_rows = [rows[0]] + [r for i,r in enumerate(rows[1:]) if i != idx]
        with open(APPLIED_JOBS_CSV, "w", newline="") as f:
            csv.writer(f).writerows(new_rows)
    return redirect(url_for("applied_jobs"))

@app.route("/interviews", methods=["GET", "POST"])
def interviews():
    conn = db_conn()
    c = conn.cursor()

    # -------- ADD INTERVIEW --------
    if request.method == "POST":
        job_role = request.form.get("job_role")
        company = request.form.get("company")
        dt = request.form.get("interview_datetime").replace("T", " ")

        user_name = session.get("name") or "User"
        user_email = session.get("email") or "user@example.com"

        c.execute("""
            INSERT INTO interviews (user_name, user_email, job_role, company, interview_datetime)
            VALUES (?,?,?,?,?)
        """, (user_name, user_email, job_role, company, dt))

        conn.commit()
        return redirect("/interviews")

    # -------- ALWAYS SHOW EVERYTHING AS UPCOMING --------
    email = session.get("email") or "user@example.com"

    c.execute("SELECT * FROM interviews WHERE user_email=? ORDER BY id DESC", (email,))
    upcoming = c.fetchall()

    c.execute("SELECT * FROM history WHERE user_email=? ORDER BY id DESC", (email,))
    history = c.fetchall()

    conn.close()
    return render_template("interviews.html", upcoming=upcoming, history=history)

@app.route("/send_reminder_now/<int:interview_id>", methods=["POST"])
def send_reminder_now(interview_id):
    conn = db_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM interviews WHERE id=?", (interview_id,))
    row = c.fetchone()
    if not row:
        flash("Interview not found", "danger")
        return redirect(url_for("interviews"))
    dt = parse_datetime_flexible(row["interview_datetime"])
    subject = f"Interview Reminder — {row['job_role']}"
    body = f"<p>Hi {row['user_name'] or 'Candidate'},</p><p>This is a reminder for your interview: <strong>{row['job_role']}</strong> at <strong>{row['company']}</strong> scheduled for <strong>{dt.strftime('%Y-%m-%d %H:%M') if dt else row['interview_datetime']}</strong>.</p>"
    ok = send_email(row['user_email'], subject, body)
    if ok:
        try:
            c.execute("UPDATE interviews SET reminder_1h_sent=1, reminder_24_sent=1 WHERE id=?", (interview_id,))
            conn.commit()
        except Exception:
            conn.rollback()
        flash("Reminder email sent", "success")
    else:
        flash("Failed to send reminder", "danger")
    conn.close()
    return redirect(url_for("interviews"))

@app.route("/send_feedback_now/<int:interview_id>", methods=["POST"])
def send_feedback_now(interview_id):
    conn = db_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM interviews WHERE id=?", (interview_id,))
    row = c.fetchone()
    if not row:
        flash("Interview not found", "danger")
        return redirect(url_for("interviews"))
    subject = f"Interview Feedback — {row['job_role']}"
    body = f"<p>Hi {row['user_name'] or 'Candidate'},</p><p>Thank you for interviewing for <strong>{row['job_role']}</strong> at <strong>{row['company']}</strong>.</p><p>Please reply with any feedback or questions.</p>"
    ok = send_email(row['user_email'], subject, body)
    if ok:
        try:
            c.execute("UPDATE interviews SET feedback_sent=1 WHERE id=?", (interview_id,))
            c.execute("""INSERT INTO history (user_name,user_email,job_role,company,interview_datetime)
                         VALUES (?,?,?,?,?)""",
                      (row["user_name"], row["user_email"], row["job_role"], row["company"], row["interview_datetime"]))
            c.execute("DELETE FROM interviews WHERE id=?", (interview_id,))
            conn.commit()
        except Exception:
            conn.rollback()
        flash("Feedback email sent & interview moved to history", "success")
    else:
        flash("Failed to send feedback", "danger")
    conn.close()
    return redirect(url_for("interviews"))

@app.route("/move_expired_interviews")
def move_expired_interviews():
    conn = db_conn()
    c = conn.cursor()
    now = datetime.now()
    c.execute("SELECT * FROM interviews")
    rows = c.fetchall()
    for r in rows:
        try:
            dt = parse_datetime_flexible(r["interview_datetime"])
            if dt and dt < now:
                c.execute("""INSERT INTO history (user_name,user_email,job_role,company,interview_datetime)
                             VALUES (?,?,?,?,?)""",
                          (r["user_name"], r["user_email"], r["job_role"], r["company"], r["interview_datetime"]))
                c.execute("DELETE FROM interviews WHERE id=?", (r["id"],))
        except Exception:
            continue
    conn.commit()
    conn.close()
    return "moved"

@app.route("/force_create_db")
def force_create_db():
    try:
        conn = db_conn()
        c = conn.cursor()
        c.execute("DROP TABLE IF EXISTS interviews")
        c.execute("DROP TABLE IF EXISTS history")

        c.execute("""
            CREATE TABLE interviews (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_name TEXT,
                user_email TEXT,
                job_role TEXT,
                company TEXT,
                interview_datetime TEXT,
                reminder_24_sent INTEGER DEFAULT 0,
                reminder_1h_sent INTEGER DEFAULT 0,
                feedback_sent INTEGER DEFAULT 0
            )
        """)
        c.execute("""
            CREATE TABLE history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_name TEXT,
                user_email TEXT,
                job_role TEXT,
                company TEXT,
                interview_datetime TEXT,
                moved_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        conn.close()
        return "Database & tables successfully (re)created!"
    except Exception as e:
        return f"ERROR: {e}"

@app.route("/dbinfo")
def dbinfo():
    conn = db_conn()
    c = conn.cursor()
    c.execute("PRAGMA table_info(interviews)")
    cols = c.fetchall()
    conn.close()
    return str(cols)

@app.route('/uploads/<path:fname>')
def uploaded_file(fname):
    return send_from_directory(BASE_UPLOAD, fname, as_attachment=False)

@app.route("/health")
def health():
    return {"status":"ok"}

@app.route("/send_test")
def send_test():
    test_email = session.get("email") or "your-real-email@example.com"
    subject = "TalentTrack Test Email"
    body = "<p>This is a test from TalentTrack.</p>"
    ok = send_email(test_email, subject, body)
    return "Test email sent" if ok else "Failed to send test email (check MAILTRAP settings)."
@app.route("/delete_interview/<int:id>", methods=["POST"])
def delete_interview(id):
    conn = db_conn()
    c = conn.cursor()
    c.execute("DELETE FROM interviews WHERE id=?", (id,))
    conn.commit()
    conn.close()
    flash("Interview deleted", "success")
    return redirect("/interviews")
@app.route("/delete_history/<int:id>", methods=["POST"])
def delete_history(id):
    conn = db_conn()
    c = conn.cursor()
    c.execute("DELETE FROM history WHERE id=?", (id,))
    conn.commit()
    conn.close()
    flash("History deleted", "success")
    return redirect("/interviews")

# ---------------- Run ----------------
if __name__ == "__main__":
    try:
        port = int(os.environ.get("PORT", 5000))
        app.run(debug=True, port=port)
    finally:
        try:
            if scheduler:
                scheduler.shutdown(wait=False)
        except Exception:
            pass
