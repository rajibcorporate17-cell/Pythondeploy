

import streamlit as st
import pandas as pd
import datetime
import os
import re
import html
import json
import uuid
import textwrap
import msal
import requests
import time
from pathlib import Path

# ==========================================
# CONFIGURATION
# ==========================================
CLIENT_ID              = "cb467c35-c5f5-4987-afff-6d77d45e41a6"
CLIENT_SECRET          = "5va8Q~EERNmh46TGkduGGd_6MM2Y7IOT5MlT4cls"
TENANT_ID              = "231edc94-ddeb-473e-a5c1-c86d43e0db76"
HOSTNAME               = "datainsightio.sharepoint.com"
SITE_PATH              = "OESRegistration"
DATA_LIST_NAME         = "Student Information"
REGISTRATION_LIST_NAME = "Student Registration"
DOCUMENT_LIBRARY_NAME  = "StudentDocuments"

REQUIRED_COLUMNS = [
    ("SubmittedAt",          "Submitted At",           "text"),
    ("ConfirmationNumber",   "Confirmation Number",    "text"),
    ("Status",               "Status",                 "text"),
    ("StudentID",            "Student ID",             "text"),
    ("FullName",             "Full Name",              "text"),
    ("CNumber",              "C Number",               "text"),
    ("DHSEmail",             "DHS Email",              "text"),
    ("PersonalEmail",        "Personal Email",         "text"),
    ("Phone",                "Phone",                  "text"),
    ("Address",              "Address",                "text"),
    ("EC1Name",              "EC1 Name",               "text"),
    ("EC1Relationship",      "EC1 Relationship",       "text"),
    ("EC1Cell",              "EC1 Cell",               "text"),
    ("EC1Work",              "EC1 Work",               "text"),
    ("Semester",             "Semester",               "text"),
    ("Courses",              "Courses",                "note"),
    ("Tuition",              "Tuition",                "text"),
    ("FinancialAid",         "Financial Aid",          "text"),
    ("FAFSA",                "FAFSA",                  "text"),
    ("PaymentElection",      "Payment Election",       "text"),
    ("RequiredDocsUploaded", "Required Docs Uploaded", "text"),
    ("OptionalDocsUploaded", "Optional Docs Uploaded", "text"),
    ("DocumentFilesJSON",    "Document Files JSON",    "note"),
]

# ==========================================
# PAGE SETUP
# ==========================================
st.set_page_config(page_title="OES Student Portal", page_icon="🎓", layout="wide")

_original_streamlit_markdown = st.markdown

def normalize_html(markup: str) -> str:
    markup = textwrap.dedent(str(markup)).strip()
    return "\n".join(line.lstrip() for line in markup.splitlines())

def _safe_markdown(body, *args, **kwargs):
    if kwargs.get("unsafe_allow_html") and isinstance(body, str):
        body = normalize_html(body)
    return _original_streamlit_markdown(body, *args, **kwargs)

st.markdown = _safe_markdown

def load_css(filepath: str):
    css_path = os.path.join(os.path.dirname(__file__), filepath)
    try:
        with open(css_path, "r", encoding="utf-8") as f:
            css = f.read()
        st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning("style1.css not found.")

load_css("style1.css")

# ==========================================
# SHAREPOINT AUTH
# ==========================================
@st.cache_data(ttl=3000)
def get_access_token() -> str:
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise ConnectionError(f"Auth failed: {result.get('error_description', result)}")
    return result["access_token"]

@st.cache_data(ttl=86400)
def get_site_id() -> str:
    token = get_access_token()
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{HOSTNAME}:/sites/{SITE_PATH}",
        headers={"Authorization": f"Bearer {token}"}
    )
    if resp.status_code != 200:
        raise ConnectionError(f"Site not found: {resp.text}")
    return resp.json()["id"]

@st.cache_data(ttl=300)
def load_data() -> pd.DataFrame:
    for attempt in range(3):
        try:
            token   = get_access_token()
            site_id = get_site_id()
            headers = {"Authorization": f"Bearer {token}"}

            col_resp = requests.get(
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{DATA_LIST_NAME}/columns",
                headers=headers
            )
            if col_resp.status_code == 401:
                get_access_token.clear()
                continue

            column_mapping = {}
            if col_resp.status_code == 200:
                column_mapping = {
                    c["name"]: c["displayName"]
                    for c in col_resp.json().get("value", [])
                    if "name" in c and "displayName" in c
                }

            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}"
                f"/lists/{DATA_LIST_NAME}/items?expand=fields&$top=5000"
            )
            records = []
            while url:
                resp = requests.get(url, headers=headers)
                if resp.status_code == 401:
                    get_access_token.clear()
                    raise ValueError("Token expired mid-pagination")
                if resp.status_code != 200:
                    raise ValueError(f"Failed to fetch list: {resp.text}")
                data = resp.json()
                records.extend(item["fields"] for item in data.get("value", []))
                url = data.get("@odata.nextLink")

            df = pd.DataFrame(records)
            if column_mapping:
                df.rename(columns=column_mapping, inplace=True)
            if "Form Access #" in df.columns and "C#" not in df.columns:
                df.rename(columns={"Form Access #": "C#"}, inplace=True)
            df.columns = df.columns.astype(str).str.strip()

            if "StudentID" not in df.columns or "C#" not in df.columns:
                raise ValueError(f"Missing required columns. Found: {list(df.columns)}")
            df["StudentID"] = df["StudentID"].astype(str).str.strip().str.lower()
            df["C#"]        = df["C#"].astype(str).str.strip().str.lower()
            return df

        except Exception as e:
            if attempt == 2:
                raise e
            get_access_token.clear()

# ==========================================
# COLUMN AUTO-CREATION
# ==========================================
def ensure_registration_list_columns() -> tuple:
    token   = get_access_token()
    site_id = get_site_id()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    col_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{REGISTRATION_LIST_NAME}/columns"

    col_resp = requests.get(col_url, headers={"Authorization": f"Bearer {token}"})
    if col_resp.status_code != 200:
        raise RuntimeError(f"Cannot read '{REGISTRATION_LIST_NAME}' columns [{col_resp.status_code}]: {col_resp.text[:300]}")

    existing_cols = col_resp.json().get("value", [])
    existing_by_internal = {c.get("name","").lower(): c.get("name","") for c in existing_cols}
    existing_by_display  = {c.get("displayName","").lower().strip(): c.get("name","") for c in existing_cols}

    name_map = {}
    created  = []

    for internal_name, display_name, field_type in REQUIRED_COLUMNS:
        if internal_name.lower() in ("title","id","created","modified"):
            continue
        actual = existing_by_internal.get(internal_name.lower()) or existing_by_display.get(display_name.lower())
        if actual:
            name_map[internal_name] = actual
            continue
        col_def = {
            "name": internal_name, "displayName": display_name,
            "text": {"allowMultipleLines": True, "linesForEditing": 6} if field_type == "note" else {}
        }
        r = requests.post(col_url, headers=headers, json=col_def)
        name_map[internal_name] = r.json().get("name", internal_name) if r.status_code in (200,201) else internal_name
        if r.status_code in (200,201):
            created.append(internal_name)

    return name_map, created

# ==========================================
# SHAREPOINT WRITE — SAFE
# ==========================================
def submit_to_registration_list(submission_data: dict) -> dict:
    name_map, created_cols = ensure_registration_list_columns()
    if created_cols:
        time.sleep(4)

    token   = get_access_token()
    site_id = get_site_id()
    auth_headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    write_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{REGISTRATION_LIST_NAME}/items"

    col_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{REGISTRATION_LIST_NAME}/columns",
        headers=auth_headers
    )
    if col_resp.status_code != 200:
        raise RuntimeError(f"Cannot read list columns [{col_resp.status_code}]: {col_resp.text[:300]}")

    existing_cols = col_resp.json().get("value", [])
    valid_name_map    = {c["name"].lower(): c["name"] for c in existing_cols if c.get("name")}
    valid_display_map = {c.get("displayName","").lower().replace(" ",""): c["name"] for c in existing_cols if c.get("name") and c.get("displayName")}

    fields = {}
    for our_key, value in submission_data.items():
        mapped   = name_map.get(our_key, our_key)
        real_name = (
            valid_name_map.get(mapped.lower())
            or valid_name_map.get(our_key.lower())
            or valid_display_map.get(our_key.lower().replace(" ",""))
        )
        if not real_name:
            continue
        raw = str(value) if value is not None else ""
        fields[real_name] = raw[:3800] if field_is_note(our_key) else raw[:255]

    if not fields:
        raise RuntimeError(f"No valid fields to write. name_map={name_map}, skipped all keys.")

    resp = requests.post(write_url, headers=auth_headers, json={"fields": fields})
    if resp.status_code in (200, 201):
        return resp.json()

    # Fallback: write field by field, skip bad ones
    good_fields = {}
    bad_fields  = {}
    for fname, fval in fields.items():
        probe = requests.post(write_url, headers=auth_headers, json={"fields": {fname: fval}})
        if probe.status_code in (200, 201):
            item_id = probe.json().get("id")
            if item_id:
                requests.delete(
                    f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{REGISTRATION_LIST_NAME}/items/{item_id}",
                    headers=auth_headers
                )
            good_fields[fname] = fval
        else:
            bad_fields[fname] = probe.text[:120]

    if not good_fields:
        raise RuntimeError(f"All fields rejected. First error: {resp.text}")

    final_resp = requests.post(write_url, headers=auth_headers, json={"fields": good_fields})
    if final_resp.status_code not in (200, 201):
        raise RuntimeError(f"Student Registration list write failed [{final_resp.status_code}]: {final_resp.text}")
    return final_resp.json()

def field_is_note(key: str) -> bool:
    for internal_name, _, field_type in REQUIRED_COLUMNS:
        if internal_name == key:
            return field_type == "note"
    return False

# ==========================================
# DOCUMENT LIBRARY — FILE UPLOAD
# ==========================================
@st.cache_data(ttl=86400)
def get_document_library_drive_id() -> str:
    token   = get_access_token()
    site_id = get_site_id()
    headers = {"Authorization": f"Bearer {token}"}
    drives_resp = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives", headers=headers)
    if drives_resp.status_code == 200:
        drives = drives_resp.json().get("value", [])
        for d in drives:
            if d.get("name","").lower() == DOCUMENT_LIBRARY_NAME.lower():
                return d["id"]
        if drives:
            return drives[0]["id"]
    drive_resp = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
    if drive_resp.status_code != 200:
        raise RuntimeError(f"Cannot get drive: {drive_resp.text[:200]}")
    return drive_resp.json()["id"]

# ── Document key → SharePoint folder/file label mapping ──────────────────────
_DOC_LABEL_MAP = {
    "bls":             "BLSCertificate",
    "health":          "HealthClearance",
    "id_front":        "IDBadgeFront",
    "id_back":         "IDBadgeBack",
    "access_front":    "AccessCardFront",
    "access_back":     "AccessCardBack",
    "tuition_receipt": "TuitionPaymentReceipt",
    "pregnancy_1":     "DignityInPregnancyPart1",
    "pregnancy_2":     "DignityInPregnancyPart2",
    "pregnancy_3":     "DignityInPregnancyPart3",
}

def _ensure_sp_folder(drive_id: str, token: str, folder_path: str) -> None:
    """
    Walks each path segment and creates any missing folders via Graph API.
    Uses @microsoft.graph.conflictBehavior=ignore so re-runs are safe.
    """
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    parts   = folder_path.strip("/").split("/")
    current = ""
    for part in parts:
        current = f"{current}/{part}" if current else part
        check_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{current}"
        r = requests.get(check_url, headers=headers)
        if r.status_code == 404:
            parent_path = "/".join(current.split("/")[:-1])
            if parent_path:
                parent_url = (
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
                    f"/root:/{parent_path}:/children"
                )
            else:
                parent_url = (
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
                )
            requests.post(
                parent_url, headers=headers,
                json={"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "ignore"}
            )
        # 200 → already exists; other errors surface naturally at upload time

def upload_file_to_sharepoint(student_id: str, doc_key: str, file_bytes: bytes, filename: str,
                               last_name: str = "", first_name: str = "") -> str:
    """
    Uploads to the nested structure:
        StudentID / LastName_FI_DocLabel / LastNameFI_DocLabel.ext

    Example:
        A000123456 / Smith_J_BLSCertificate / SmithJ_BLSCertificate.pdf

    Falls back to original flat path if last_name/first_name are not supplied.
    """
    token    = get_access_token()
    drive_id = get_document_library_drive_id()

    doc_label = _DOC_LABEL_MAP.get(doc_key, doc_key)

    if last_name and first_name:
        clean          = lambda s: re.sub(r"[^A-Za-z0-9]", "", s)
        ln             = clean(last_name.strip().title())
        fi             = clean(first_name.strip()[0].upper()) if first_name.strip() else "X"
        ext            = os.path.splitext(filename)[1].lower() or ".bin"
        subfolder_name = f"{ln}_{fi}_{doc_label}"          # Smith_J_BLSCertificate
        final_filename = f"{ln}{fi}_{doc_label}{ext}"      # SmithJ_BLSCertificate.pdf
        folder_path    = f"{student_id}/{subfolder_name}"  # A000123456/Smith_J_BLSCertificate
    else:
        # Backward-compatible fallback (original behaviour)
        safe_name      = re.sub(r"[^A-Za-z0-9_.\-]", "_", filename)
        stamp          = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        final_filename = f"{doc_key}_{stamp}_{safe_name}"
        folder_path    = student_id

    _ensure_sp_folder(drive_id, token, folder_path)

    upload_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
        f"/root:/{folder_path}/{final_filename}:/content"
    )
    up_resp = requests.put(
        upload_url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"},
        data=file_bytes,
    )
    if up_resp.status_code not in (200, 201):
        raise RuntimeError(f"File upload failed for '{filename}' [{up_resp.status_code}]: {up_resp.text[:300]}")
    return up_resp.json().get("webUrl", "")

# ==========================================
# FILENAME GENERATOR: LastName_FirstInitial_DocumentName
# ==========================================
def make_sp_filename(last_name: str, first_name: str, doc_label: str, original_ext: str) -> str:
    """Generates: SmithJ_BLSCertificate.pdf"""
    clean = lambda s: re.sub(r"[^A-Za-z0-9]", "", s)
    ln = clean(last_name.strip().title())
    fi = clean(first_name.strip()[0].upper()) if first_name.strip() else "X"
    dl = re.sub(r"[^A-Za-z0-9]", "", doc_label.replace(" ","").replace("—","").replace("-",""))
    ext = original_ext.lower() if original_ext.startswith(".") else f".{original_ext.lower()}"
    return f"{ln}{fi}_{dl}{ext}"

# ==========================================
# HELPERS
# ==========================================
def clean_value(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if text.lower() in ["nan","none","null"]:
        return ""
    return text

def safe_get(data, col):
    if not data:
        return ""
    return clean_value(data.get(col, ""))

def norm_key(key):
    return re.sub(r"[^a-z0-9]", "", str(key).lower())

def first_available_value(data, candidates, default=""):
    if not data:
        return default
    for col in candidates:
        val = clean_value(data.get(col, ""))
        if val:
            return val
    normalized_map = {norm_key(k): k for k in data.keys()}
    for col in candidates:
        actual = normalized_map.get(norm_key(col))
        if actual:
            val = clean_value(data.get(actual, ""))
            if val:
                return val
    return default

def split_course_text(text):
    text = clean_value(text)
    if not text:
        return []
    parts = re.split(r"[,;\n|/]+", text)
    return [p.strip() for p in parts if p.strip()]

def get_student_semester(student_row):
    raw = first_available_value(student_row,
        ["Semester","Term","Registration Term","Cohort","Semester Term"], default="Semester 2")
    raw = clean_value(raw)
    if re.fullmatch(r"\d+", raw):
        return f"Semester {raw}"
    return raw

def get_student_courses(student_row):
    semester = get_student_semester(student_row)
    candidates = []
    m = re.search(r"(\d+)", semester or "")
    if m:
        n = m.group(1)
        candidates.extend([f"Semester {n} Courses",f"Semester {n} Course",f"Sem {n} Courses"])
    candidates.extend(["Semester 2 Courses","Course Name","Course Names","Course Code","Courses","Registered Courses"])
    course_text = first_available_value(student_row, candidates)
    courses = split_course_text(course_text)
    if courses:
        return courses
    found = []
    for col, val in student_row.items():
        col_name = str(col).strip()
        cell_val = clean_value(val).lower()
        if re.fullmatch(r"[a-zA-Z]{1,5}\d{2,4}[a-zA-Z]?", col_name):
            if cell_val and cell_val not in ["no","false","0","n/a","na"]:
                found.append(col_name)
    return found

def course_badges_html(courses):
    if not courses:
        return '<span class="course-pill empty">No registered courses found</span>'
    return "".join(f'<span class="course-pill">{html.escape(c)}</span>' for c in courses)

def format_money(value, default="$2,400"):
    value = clean_value(value)
    if not value:
        return default
    if value.startswith("$"):
        return value
    numeric = re.sub(r"[^0-9.]", "", value)
    if numeric:
        try:
            amount = float(numeric)
            return f"${int(amount):,}" if amount.is_integer() else f"${amount:,.2f}"
        except Exception:
            pass
    return value

def get_tuition(student_row, semester):
    candidates = ["Tuition","Tuition Charge","Amount","Semester Tuition","Tuition Amount"]
    m = re.search(r"(\d+)", semester or "")
    if m:
        n = m.group(1)
        candidates = [f"Semester {n} Tuition",f"Semester {n} Tuition Charge",f"Tuition Semester {n}"] + candidates
    return format_money(first_available_value(student_row, candidates), default="$2,400")

def get_full_name(s):
    return " ".join([safe_get(s,"FirstName"),safe_get(s,"MiddleName"),safe_get(s,"LastName")]).replace("  "," ").strip()

def get_student_id(s):
    return safe_get(s, "StudentID").upper()

def get_access_card(s):
    return first_available_value(s, ["Access Card #","AccessCard","AccessCard#","Access Card Number","AccessCardNumber"], default="")

# ==========================================
# SESSION STATE
# ==========================================
def init_state():
    defaults = {
        "step": 1, "student_data": None,
        "identity_data": {}, "contact_data": {}, "academic_data": {},
        "financial_data": {}, "documents_data": {},
        "confirmation_number": "", "submitted": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_state()

try:
    df = load_data()
except Exception:
    st.info("🔄 Connecting to server... please wait.")
    time.sleep(3)
    get_access_token.clear()
    load_data.clear()
    st.rerun()

# ==========================================
# UI HELPERS
# ==========================================
def render_navbar(student_name="", student_id=""):
    badge_html = (
        f"<div class='oes-navbar-right'>{html.escape(student_name)} · {html.escape(student_id.upper())}</div>"
        if student_name else ""
    )
    st.markdown(f"""
    <div class="oes-navbar">
        <div class="oes-navbar-left">
            <div class="oes-logo-badge">OES</div>
            <div class="oes-navbar-title">
                <div class="main-title">Office of Educational Services — LA County DHS</div>
                <div class="sub-title">Student Registration Portal · 2027-II</div>
            </div>
        </div>
        {badge_html}
    </div>
    """, unsafe_allow_html=True)

def render_progress(current_step=0, pct=0):
    steps = [("🔒","Verify"),("👤","Identity"),("📞","Contact"),
             ("🎓","Academics"),("💳","Financial"),("📁","Documents"),("✅","Review")]
    items_html = ""
    for i,(icon,label) in enumerate(steps):
        if i < current_step:      cls, disp = "done", "✓"
        elif i == current_step:   cls, disp = "active", icon
        else:                     cls, disp = "", icon
        items_html += (f'<div class="step-item"><div class="step-icon {cls}">{disp}</div>'
                       f'<div class="step-label {cls}">{label}</div></div>')
    st.markdown(
        f'<div class="progress-card"><div class="progress-header"><span>REGISTRATION PROGRESS</span>'
        f'<span class="progress-pct">{pct}% complete</span></div>'
        f'<div class="progress-steps">{items_html}</div></div>', unsafe_allow_html=True)

def render_page_start(s=None, progress_step=0, pct=0):
    if s: render_navbar(student_name=get_full_name(s), student_id=get_student_id(s))
    else: render_navbar()
    st.markdown('<div class="oes-page">', unsafe_allow_html=True)
    render_progress(current_step=progress_step, pct=pct)

def render_page_end():
    st.markdown('</div>', unsafe_allow_html=True)

def readonly_field(label, value):
    """Render a read-only field that looks like a text input but is not editable."""
    st.markdown(f"""
    <div style="margin-bottom:8px;">
        <div style="font-size:13px;font-weight:600;color:#3e5a7d;margin-bottom:4px;">{html.escape(label)}</div>
        <div style="background:#f8fafc;border:1.5px solid #d9e2ef;border-radius:8px;
                    padding:10px 14px;font-size:14px;color:#243b5a;font-weight:500;">
            {html.escape(clean_value(value) or "—")}
        </div>
    </div>
    """, unsafe_allow_html=True)

def sp_info_banner(msg="This information is from your official DHS record and cannot be edited."):
    st.markdown(f"""
    <div style="background:#eff6ff;border-left:4px solid #2563eb;border-radius:0 8px 8px 0;
                padding:10px 14px;font-size:12px;color:#1d4ed8;margin:8px 0 16px;">
        🔒 {msg}
    </div>
    """, unsafe_allow_html=True)

def summary_card(title, rows):
    body = ""
    for label, value in rows:
        body += (f'<div class="review-field"><div class="review-label">{html.escape(label)}</div>'
                 f'<div class="review-value">{html.escape(clean_value(value) or "N/A")}</div></div>')
    st.markdown(f"""
    <div class="review-card">
        <div class="review-card-header"><span>{html.escape(title)}</span></div>
        <div class="review-grid">{body}</div>
    </div>
    """, unsafe_allow_html=True)


# ==========================================
# STEP 1 — IDENTITY VERIFICATION
# Login: Student ID + Access Card# (NOT C#/Computer Number)
# ==========================================
if st.session_state.step == 1:
    render_page_start(progress_step=0, pct=0)
    st.markdown("""
    <div class="form-card">
        <div class="step-badge">🔒 STEP 0 OF 6</div>
        <div class="form-title">Identity Verification</div>
        <div class="form-subtitle">Enter your pre-issued Student ID and C# to access your registration.</div>
        <div class="info-banner">ℹ️ Your <b>Student ID</b> and <b>C#</b> were provided by your program coordinator.</div>
        <div class="credentials-label">Enter Your Credentials</div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login"):
        id_input    = st.text_input("Student ID *", placeholder="e.g. A0000000000")
        card_input  = st.text_input("C# *", type="password", placeholder="Enter your C#")
        if st.form_submit_button("🔒 Verify and Access My Registration"):
            # Match on StudentID + Access Card# (not C#)
            card_col = None
            for candidate in ["Access Card #","AccessCard","AccessCard#","Access Card Number","AccessCardNumber"]:
                if candidate in df.columns:
                    card_col = candidate
                    break
            if card_col:
                match = df[
                    (df["StudentID"] == id_input.strip().lower()) &
                    (df[card_col].astype(str).str.strip().str.lower() == card_input.strip().lower())
                ]
            else:
                # fallback to C# if no access card column found
                match = df[
                    (df["StudentID"] == id_input.strip().lower()) &
                    (df["C#"].astype(str).str.strip().str.lower() == card_input.strip().lower())
                ]
            if not match.empty:
                st.session_state.student_data = match.iloc[0].to_dict()
                st.session_state.step = 2
                st.rerun()
            else:
                st.error("❌ Invalid Credentials. Please check your Student ID and C# and try again.")
    render_page_end()


# ==========================================
# STEP 2 — IDENTITY CONFIRMATION
# All fields from SharePoint — READ ONLY, no editing.
# No birth date field — shows Access Card # instead.
# ==========================================
elif st.session_state.step == 2:
    s = st.session_state.student_data
    render_page_start(s, progress_step=1, pct=14)

    fname = first_available_value(s, ["FirstName","First Name"])
    mname = first_available_value(s, ["MiddleName","Middle Name"])
    lname = first_available_value(s, ["LastName","Last Name"])
    access_card = get_access_card(s)
    student_id  = get_student_id(s)

    st.markdown("""
    <div class="form-card">
        <div class="step-badge">👤 STEP 1 OF 6</div>
        <div class="form-title">Identity Confirmation</div>
        <div class="form-subtitle">Please confirm your legal name and Access Card # on file are correct.</div>
    </div>
    """, unsafe_allow_html=True)

    sp_info_banner("Your identity information comes directly from your DHS record. If anything is incorrect, contact your coordinator.")

    col1, col2 = st.columns(2)
    with col1:
        readonly_field("First Name", fname)
        readonly_field("Middle Name", mname)
    with col2:
        readonly_field("Last Name", lname)
        readonly_field("Access Card #", access_card)

    readonly_field("Student ID", student_id)

    st.markdown("<br>", unsafe_allow_html=True)
    col_back, col_next = st.columns([1,3])
    with col_back:
        if st.button("← Logout"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            init_state()
            st.rerun()
    with col_next:
        if st.button("This is correct — Continue ➡️", type="primary"):
            st.session_state.identity_data = {
                "first_name":  fname,
                "middle_name": mname,
                "last_name":   lname,
                "access_card": access_card,
            }
            st.session_state.step = 3
            st.rerun()
    render_page_end()


# ==========================================
# STEP 3 — CONTACT & ADDRESS
# SharePoint fields: read-only display
# Only EMPTY fields are editable
# Emergency contact fields: always editable (user fills in)
# ==========================================
elif st.session_state.step == 3:
    s = st.session_state.student_data
    render_page_start(s, progress_step=2, pct=28)
    st.markdown("""
    <div class="form-card">
        <div class="step-badge">📞 STEP 2 OF 6</div>
        <div class="form-title">Contact & Address</div>
        <div class="form-subtitle">Fields pre-filled from your DHS record are locked. Only empty fields can be edited.</div>
    </div>
    """, unsafe_allow_html=True)

    sp_info_banner("Pre-filled fields are from your official DHS record and cannot be changed here.")

    # Pull values from SharePoint
    sp_phone   = first_available_value(s, ["Phone1","Phone","Cell","Mobile"])
    sp_email   = first_available_value(s, ["Personal Email","PersonalEmail","Email"])
    sp_address = first_available_value(s, ["Address1","Address","Home Address"])
    sp_city    = first_available_value(s, ["City"])
    sp_state   = first_available_value(s, ["State"])
    sp_zip     = first_available_value(s, ["ZipCode","Zip Code","Zip"])

    with st.form("contact"):
        st.markdown("**Contact Information**")
        col1, col2 = st.columns(2)
        with col1:
            if sp_phone:
                readonly_field("Phone", sp_phone)
                phone = sp_phone
            else:
                phone = st.text_input("Phone *", placeholder="Enter your phone number")
        with col2:
            if sp_email:
                readonly_field("Personal Email", sp_email)
                personal_email = sp_email
            else:
                personal_email = st.text_input("Personal Email *", placeholder="Enter your personal email")

        if sp_address:
            readonly_field("Home Address", sp_address)
            address = sp_address
        else:
            address = st.text_input("Address *", placeholder="Enter your street address")

        c3, c4, c5 = st.columns([2,1,1])
        with c3:
            if sp_city:
                readonly_field("City", sp_city)
                city = sp_city
            else:
                city = st.text_input("City *", placeholder="City")
        with c4:
            if sp_state:
                readonly_field("State", sp_state)
                state = sp_state
            else:
                state = st.text_input("State *", placeholder="CA")
        with c5:
            if sp_zip:
                readonly_field("Zip Code", sp_zip)
                zipc = sp_zip
            else:
                zipc = st.text_input("Zip *", placeholder="90001")

        # st.markdown("---")
        # st.markdown("**Emergency Contact** — Please enter your emergency contact information below.")

        # ec1_name = st.text_input(
        #     "Emergency Contact Name *",
        #     value=first_available_value(s, ["EC1 Name","Emergency Contact Name"]),
        #     placeholder="Full name (optional)"
        # )
        # ec_col1, ec_col2, ec_col3 = st.columns(3)
        # with ec_col1:
        #     ec1_relationship = st.text_input(
        #         "Relationship *",
        #         value=first_available_value(s, ["EC1 Relationship"]),
        #         placeholder="e.g. Spouse, Parent (optional)"
        #     )
        # with ec_col2:
        #     ec1_cell = st.text_input(
        #         "Cell Phone *",
        #         value=first_available_value(s, ["EC1 Cell","EC1 Phone"]),
        #         placeholder="Cell number (optional)"
        #     )
        # with ec_col3:
        #     ec1_work = st.text_input(
        #         "Work Phone",
        #         value=first_available_value(s, ["EC1 Work"]),
        #         placeholder="Work number (optional)"
        #     )
        st.markdown("---")
        st.markdown(
            "**Emergency Contact** — Fields from DHS record are locked. "
            "Only empty fields can be edited (optional)."
        )

        # SharePoint emergency contact values
        sp_ec1_name = first_available_value(
            s, ["EC1 Name", "Emergency Contact Name"]
        )
        sp_ec1_relationship = first_available_value(
            s, ["EC1 Relationship"]
        )
        sp_ec1_cell = first_available_value(
            s, ["EC1 Cell", "EC1 Phone"]
        )
        sp_ec1_work = first_available_value(
            s, ["EC1 Work"]
        )

        # Emergency Contact Name
        if sp_ec1_name:
            readonly_field("Emergency Contact Name", sp_ec1_name)
            ec1_name = sp_ec1_name
        else:
            ec1_name = st.text_input(
                "Emergency Contact Name",
                placeholder="Full name (optional)"
            )

        ec_col1, ec_col2, ec_col3 = st.columns(3)

        with ec_col1:
            if sp_ec1_relationship:
                readonly_field("Relationship", sp_ec1_relationship)
                ec1_relationship = sp_ec1_relationship
            else:
                ec1_relationship = st.text_input(
                    "Relationship",
                    placeholder="e.g. Parent, Spouse (optional)"
                )

        with ec_col2:
            if sp_ec1_cell:
                readonly_field("Cell Phone", sp_ec1_cell)
                ec1_cell = sp_ec1_cell
            else:
                ec1_cell = st.text_input(
                    "Cell Phone",
                    placeholder="Cell number (optional)"
                )

        with ec_col3:
            if sp_ec1_work:
                readonly_field("Work Phone", sp_ec1_work)
                ec1_work = sp_ec1_work
            else:
                ec1_work = st.text_input(
                    "Work Phone",
                    placeholder="Work number (optional)"
                )

        col_back, col_next = st.columns([1,3])
        with col_back:
            back = st.form_submit_button("← Back")
        with col_next:
            cont = st.form_submit_button("Confirm & Continue ➡️", type="primary" if hasattr(st.form_submit_button, '__self__') else None)

        if back:
            st.session_state.step = 2
            st.rerun()
        if cont:
            st.session_state.contact_data = {
                "phone": phone,
                "personal_email": personal_email,
                "address": address,
                "city": city,
                "state": state,
                "zip": zipc,
                "ec1_name": ec1_name,
                "ec1_relationship": ec1_relationship,
                "ec1_cell": ec1_cell,
                "ec1_work": ec1_work
            }
            st.session_state.step = 4
            st.rerun()
            
    render_page_end()


# ==========================================
# STEP 4 — ACADEMIC VERIFICATION (read-only from SP)
# ==========================================
elif st.session_state.step == 4:
    s = st.session_state.student_data
    render_page_start(s, progress_step=3, pct=42)

    semester     = get_student_semester(s)
    courses      = get_student_courses(s)
    courses_text = ", ".join(courses)

    st.markdown(f"""
    <div class="form-card academic-main-card">
        <div class="step-badge">🎓 STEP 3 OF 6</div>
        <div class="form-title">Academic Registration Confirmation</div>
        <div class="form-subtitle">Your registered courses are pulled directly from your OES record. Course changes must go through the Registrar.</div>
        <div class="academic-dark-card">
            <div class="academic-label">SEMESTER</div>
            <div class="academic-semester">{html.escape(semester)}</div>
            <div class="academic-divider"></div>
            <div class="academic-label">REGISTERED COURSES</div>
            <div class="course-pill-wrap">{course_badges_html(courses)}</div>
        </div>
        <div class="info-banner academic-info-banner">
            🔒 Academic information is read-only. For changes, contact the Registrar or Office of Educational Services.
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("academic_verification"):
        st.markdown("<div class='course-confirm-title'>IS YOUR SEMESTER COURSE LIST CORRECT?</div>", unsafe_allow_html=True)
        col_ok, col_bad = st.columns(2)
        with col_ok:
            submit_ok  = st.form_submit_button("✅ Yes, Everything is Correct")
        with col_bad:
            submit_bad = st.form_submit_button("⚠️ No, This is Not Correct")
        if submit_ok:
            st.session_state.academic_data = {"semester": semester, "courses": courses_text, "academic_confirmed": "Yes"}
            st.session_state.step = 5
            st.rerun()
        if submit_bad:
            st.warning("⚠️ Please contact the Office of Educational Services or Registrar to correct your academic record before continuing.")

    col_back, col_save = st.columns([1,2])
    with col_back:
        if st.button("← Back"):
            st.session_state.step = 3
            st.rerun()
    render_page_end()


# ==========================================
# STEP 5 — FINANCIAL (read-only from SP)
# ==========================================
elif st.session_state.step == 5:
    s        = st.session_state.student_data
    render_page_start(s, progress_step=4, pct=67)

    semester      = st.session_state.academic_data.get("semester") or get_student_semester(s)
    tuition       = get_tuition(s, semester)
    financial_aid = first_available_value(s, ["Financial Aid","FinancialAid","Aid","Aid Status"], default="No")
    fafsa         = first_available_value(s, ["FAFSA","FAFSA Status","Fafsa"], default="N/A")
    aid_badge     = "No Financial Aid on File" if financial_aid.lower() in ["no","n","none","n/a",""] else f"Financial Aid: {financial_aid}"

    st.markdown(f"""
    <div class="form-card financial-card">
        <div class="step-badge">💳 STEP 4 OF 6</div>
        <div class="form-title">Financial Information</div>
        <div class="form-subtitle">Your tuition and financial aid status are from your official DHS record.</div>
        <div class="tuition-card">
            <div>
                <div class="tuition-term">2027-II {html.escape(semester)} — Tuition Charge</div>
                <div class="tuition-amount">{html.escape(tuition)}</div>
            </div>
            <div class="aid-badge">{html.escape(aid_badge)}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("financial_form"):
        st.markdown('<div class="section-line-title"><span>SELECT YOUR PAYMENT INTENT</span></div>', unsafe_allow_html=True)
        payment_choice = st.radio(
            "Payment Intent",
            ["Self-Pay — Direct payment"],
            format_func=lambda x: "💵  I will pay the tuition charge directly to the institution",
            label_visibility="collapsed"
        )
        col_back, col_next = st.columns([1,3])
        with col_back:
            back = st.form_submit_button("← Back")
        with col_next:
            cont = st.form_submit_button("Save & Continue →")
        if back:
            st.session_state.step = 4
            st.rerun()
        if cont:
            st.session_state.financial_data = {
                "tuition": tuition, "financial_aid": financial_aid,
                "fafsa": fafsa, "payment_election": payment_choice
            }
            st.session_state.step = 6
            st.rerun()
    render_page_end()


# ==========================================
# STEP 6 — DOCUMENT UPLOAD
#
# SECTION 1: MANDATORY DOCUMENTS (all required)
#   - BLS Certificate, Health Clearance, ID Badge Front/Back, Access Card Front/Back
#
# SECTION 2: CONDITIONAL DOCUMENTS (required)
#   - Tuition Payment Receipt
#
# SECTION 3: OPTIONAL DOCUMENTS (Dignity in Pregnancy — Part 1, 2, 3)
#   - For incoming 2nd Semester Students — mandatory for them but listed as optional here
#   - All 3 parts shown
#
# File saved as: LastName_FirstInitial_DocumentName.ext
# ==========================================
elif st.session_state.step == 6:
    s = st.session_state.student_data
    render_page_start(s, progress_step=5, pct=83)

    # Derive student name for filename generation
    identity2  = st.session_state.identity_data or {}
    first_name = identity2.get("first_name") or first_available_value(s, ["FirstName","First Name"])
    last_name  = identity2.get("last_name")  or first_available_value(s, ["LastName","Last Name"])

    st.markdown("""
    <div class="form-card docs-card">
        <div class="step-badge">📁 STEP 5 OF 6</div>
        <div class="form-title">Document Upload</div>
        <div class="form-subtitle">
            All documents are required. Files will be saved as
            <b>LastName_FirstInitial_DocumentName</b> in SharePoint.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 3 sections as per requirements
    mandatory_docs = [
        ("bls",          "🩺", "BLS Certificate"),
        ("health",       "🏥", "Health Clearance"),
        ("id_front",     "🪪", "ID Badge Front"),
        ("id_back",      "🪪", "ID Badge Back"),
        ("access_front", "💳", "Access Card Front"),
        ("access_back",  "💳", "Access Card Back"),
    ]
    conditional_docs = [
        ("tuition_receipt", "🧾", "Tuition Payment Receipt"),
    ]
    # Section 3: Dignity in Pregnancy — 3 parts (mandatory for 2nd semester students)
    optional_docs = [
        ("pregnancy_1", "📜", "Dignity in Pregnancy Part 1"),
        ("pregnancy_2", "📜", "Dignity in Pregnancy Part 2"),
        ("pregnancy_3", "📜", "Dignity in Pregnancy Part 3"),
    ]

    # ALL docs are required (every section)
    all_required_docs = mandatory_docs + conditional_docs + optional_docs
    docs_data = st.session_state.documents_data

    # ── Capture newly uploaded files from widget state (multi-file per slot) ──
    for key, _icon, _label in all_required_docs:
        uploaded_list = st.session_state.get(f"upload_{key}") or []
        if not isinstance(uploaded_list, list):
            uploaded_list = [uploaded_list]
        if uploaded_list:
            existing = docs_data.get(key) or []
            existing_names = {f["original_name"] for f in existing}
            for uf in uploaded_list:
                if uf.name not in existing_names:
                    fb       = bytes(uf.getbuffer())
                    ext      = Path(uf.name).suffix or ".pdf"
                    # Append index suffix when multiple files for same slot
                    idx      = len(existing) + 1
                    base_label = _label if idx == 1 else f"{_label} ({idx})"
                    sp_fname = make_sp_filename(last_name, first_name, base_label, ext)
                    existing.append({
                        "filename":      sp_fname,
                        "original_name": uf.name,
                        "file_bytes":    fb,
                        "sharepoint_url": "",
                        "size_kb":       round(len(fb) / 1024, 1),
                        "uploaded_at":   datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                    existing_names.add(uf.name)
            docs_data[key] = existing
            st.session_state.documents_data = docs_data

    total_uploaded = sum(1 for key,_,_ in all_required_docs if docs_data.get(key))
    total_required = len(all_required_docs)
    dot_color      = "green" if total_uploaded == total_required else "orange"

    # ── Overall progress bar ──────────────────────────────────────────────────
    pct_done = int((total_uploaded / total_required) * 100) if total_required else 0
    bar_color = "#16a34a" if total_uploaded == total_required else "#2563eb"
    st.markdown(f"""
    <div style="background:white;border:1px solid #d8e1ec;border-radius:12px;
                padding:16px 22px;margin-bottom:20px;
                box-shadow:0 2px 8px rgba(12,35,70,.06);">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
            <div style="font-size:13px;font-weight:800;color:#001b44;display:flex;align-items:center;gap:8px;">
                <span class="status-dot {dot_color}" style="display:inline-block;"></span>
                {total_uploaded} of {total_required} document slots filled
            </div>
            <div style="font-size:12px;font-weight:700;color:{bar_color};">{pct_done}%</div>
        </div>
        <div style="background:#e8eef6;border-radius:6px;height:7px;overflow:hidden;">
            <div style="width:{pct_done}%;height:100%;background:{bar_color};
                        border-radius:6px;transition:width .3s;"></div>
        </div>
        <div style="font-size:11px;color:#7a8eaa;margin-top:8px;">
            📎 Multiple files allowed per slot &nbsp;·&nbsp; PDF, JPG, PNG &nbsp;·&nbsp; Max 25 MB per file
        </div>
    </div>
    """, unsafe_allow_html=True)

    def render_doc_section(section_title, section_note, docs):
        st.markdown(
            f'<div class="section-line-title doc-title"><span>{html.escape(section_title)}</span></div>',
            unsafe_allow_html=True
        )
        if section_note:
            st.markdown(
                f'<div style="font-size:12px;color:#5d7290;margin:-6px 0 16px;">'
                f'{html.escape(section_note)}</div>',
                unsafe_allow_html=True
            )

        for key, icon, label in docs:
            file_list = docs_data.get(key) or []
            has_files = bool(file_list)
            n_files   = len(file_list)

            status_color  = "#15803d" if has_files else "#9b1c1c"
            status_border = "#86efac" if has_files else "#fecaca"
            status_bg     = "#f0fdf4" if has_files else "#ffffff"
            status_text   = (f"✓ {n_files} file{'s' if n_files != 1 else ''} uploaded"
                             if has_files else "Required — not yet uploaded")
            status_icon   = "✅" if has_files else "📤"

            # ── Card shell ────────────────────────────────────────────────
            st.markdown(f"""
            <div style="border:1.5px solid {status_border};border-radius:14px;
                        background:{status_bg};padding:16px 18px 10px;
                        margin:0 0 12px;box-shadow:0 2px 8px rgba(12,35,70,.05);">
                <div style="display:flex;align-items:center;gap:12px;">
                    <div style="width:42px;height:42px;border-radius:10px;flex-shrink:0;
                                background:{'#bbf7d0' if has_files else '#dbeafe'};
                                display:flex;align-items:center;justify-content:center;
                                font-size:19px;">{icon}</div>
                    <div style="flex:1;min-width:0;">
                        <div style="font-size:14px;font-weight:800;color:#001b44;
                                    white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
                            {html.escape(label)}</div>
                        <div style="font-size:11px;font-weight:700;color:{status_color};margin-top:3px;">
                            {status_icon} {html.escape(status_text)}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            # ── Per-file chips with individual remove buttons ─────────────
            if has_files:
                st.markdown('<div style="margin-top:12px;display:flex;flex-direction:column;gap:6px;">',
                            unsafe_allow_html=True)
                for idx, finfo in enumerate(file_list):
                    ext_icon = "📄" if str(finfo["filename"]).lower().endswith(".pdf") else "🖼"
                    st.markdown(f"""
                    <div style="display:flex;align-items:center;gap:10px;
                                background:white;border:1px solid #bbf7d0;border-radius:9px;
                                padding:8px 12px;">
                        <span style="font-size:16px;">{ext_icon}</span>
                        <div style="flex:1;min-width:0;">
                            <div style="font-size:12px;font-weight:700;color:#15803d;
                                        white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
                                {html.escape(finfo["filename"])}</div>
                            <div style="font-size:10px;color:#6b7280;margin-top:2px;">
                                {html.escape(finfo["original_name"])} &nbsp;·&nbsp; {finfo["size_kb"]} KB
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    # Individual remove button sits right after each chip
                    if st.button(f"✕ Remove", key=f"remove_{key}_{idx}",
                                 help=f"Remove {finfo['original_name']}"):
                        file_list.pop(idx)
                        if file_list:
                            docs_data[key] = file_list
                        else:
                            docs_data.pop(key, None)
                        if f"upload_{key}" in st.session_state:
                            del st.session_state[f"upload_{key}"]
                        st.session_state.documents_data = docs_data
                        st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("</div>", unsafe_allow_html=True)  # close card shell

            # ── File uploader (always visible, add more any time) ─────────
            upload_label = f"➕ Add {'more ' if has_files else ''}file(s) — {label}"
            st.file_uploader(
                upload_label,
                type=["pdf", "jpg", "jpeg", "png"],
                key=f"upload_{key}",
                accept_multiple_files=True,
                label_visibility="visible" if not has_files else "collapsed",
            )
            # "Add more" hint when files already exist
            if has_files:
                st.markdown(
                    f'<div style="font-size:11px;color:#2563eb;font-weight:600;'
                    f'margin:-8px 0 10px;padding-left:4px;">'
                    f'⬆ Use the uploader above to add more files to this slot</div>',
                    unsafe_allow_html=True
                )

    render_doc_section(
        "SECTION 1 — MANDATORY DOCUMENTS",
        "Required for all students.",
        mandatory_docs
    )
    render_doc_section(
        "SECTION 2 — CONDITIONAL DOCUMENTS",
        "Required if applicable.",
        conditional_docs
    )
    render_doc_section(
        "SECTION 3 — DIGNITY IN PREGNANCY CERTIFICATES",
        "Mandatory for incoming 2nd Semester Students. All 3 parts must be uploaded.",
        optional_docs
    )

    # Missing docs warning
    missing = [(key, label) for key, _, label in all_required_docs if not docs_data.get(key)]
    if missing:
        missing_lines = "".join(
            f'<div style="color:#dc2626;font-size:13px;margin:4px 0;">❌ {html.escape(label)}</div>'
            for _, label in missing
        )
        st.markdown(f"""
        <div style="background:#fff5f5;border:1.5px solid #fca5a5;border-radius:10px;
                    padding:14px 18px;margin:12px 0;max-width:780px;">
            <div style="font-weight:800;color:#b91c1c;font-size:13px;margin-bottom:8px;">
                ⚠️ Please upload ALL documents before continuing:
            </div>
            {missing_lines}
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="background:#f0fdf4;border:1.5px solid #86efac;border-radius:10px;
                    padding:12px 18px;margin:12px 0;max-width:780px;
                    color:#15803d;font-weight:700;font-size:13px;">
            ✅ All documents uploaded — ready to submit!
        </div>
        """, unsafe_allow_html=True)

    # Connection test
    with st.expander("🔧 Test SharePoint Connection", expanded=False):
        if st.button("Run Connection Test"):
            test_log = []
            try:
                test_log.append("1️⃣ Getting access token...")
                tok = get_access_token()
                test_log.append(f"✅ Token OK (length={len(tok)})")
                test_log.append("2️⃣ Getting site ID...")
                sid = get_site_id()
                test_log.append(f"✅ Site ID: {sid}")
                test_log.append(f"3️⃣ Checking '{REGISTRATION_LIST_NAME}' list columns...")
                h = {"Authorization": f"Bearer {tok}"}
                lr = requests.get(
                    f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{REGISTRATION_LIST_NAME}/columns",
                    headers=h)
                if lr.status_code == 200:
                    cols = [(c["name"], c["displayName"]) for c in lr.json().get("value", [])]
                    test_log.append(f"✅ {len(cols)} columns found:")
                    for cn, dn in cols:
                        test_log.append(f"   • internal='{cn}'  display='{dn}'")
                else:
                    test_log.append(f"❌ List error [{lr.status_code}]: {lr.text[:300]}")
                test_log.append(f"4️⃣ Finding drive...")
                drive_id = get_document_library_drive_id()
                test_log.append(f"✅ Drive ID: {drive_id[:30]}...")
                test_log.append("5️⃣ Running ensure_registration_list_columns...")
                name_map, created = ensure_registration_list_columns()
                if created:
                    test_log.append(f"✅ Created new columns: {created}")
                else:
                    test_log.append("✅ All required columns exist.")
                test_log.append(f"   name_map: {name_map}")
            except Exception as te:
                test_log.append(f"❌ Exception: {type(te).__name__}: {te}")
            for line in test_log:
                st.write(line)

    col_back, col_next = st.columns([1, 3])
    with col_back:
        if st.button("← Back"):
            st.session_state.step = 5
            st.rerun()
    with col_next:
        if st.button("Submit Registration →", type="primary"):
            if missing:
                st.warning(f"⚠️ Please upload all {total_required} required documents before submitting. Missing: {len(missing)}")
                st.stop()

            s2         = st.session_state.student_data
            contact2   = st.session_state.contact_data   or {}
            academic2  = st.session_state.academic_data  or {}
            financial2 = st.session_state.financial_data or {}

            full_name2 = " ".join(filter(None, [
                identity2.get("first_name") or first_available_value(s2, ["FirstName","First Name"]),
                identity2.get("middle_name") or first_available_value(s2, ["MiddleName","Middle Name"]),
                identity2.get("last_name")  or first_available_value(s2, ["LastName","Last Name"]),
            ])).strip() or get_full_name(s2)

            student_id2   = get_student_id(s2)
            semester2     = academic2.get("semester") or get_student_semester(s2)
            tuition2      = financial2.get("tuition") or get_tuition(s2, semester2)
            full_address2 = ", ".join(v for v in [
                contact2.get("address"), contact2.get("city"),
                contact2.get("state"),   contact2.get("zip")
            ] if v)

            prog  = st.progress(0, text="Starting upload…")
            erbox = st.empty()

            try:
                # Phase 1: Upload all files to SharePoint (multi-file per slot)
                uploaded_sp = {}
                # Flatten: list of (slot_key, file_dict) pairs
                file_items = []
                for k, file_list in docs_data.items():
                    if isinstance(file_list, list):
                        for finfo in file_list:
                            if finfo.get("file_bytes"):
                                file_items.append((k, finfo))
                    elif isinstance(file_list, dict) and file_list.get("file_bytes"):
                        file_items.append((k, file_list))  # backward compat

                _upload_last  = identity2.get("last_name")  or first_available_value(s2, ["LastName", "Last Name"])
                _upload_first = identity2.get("first_name") or first_available_value(s2, ["FirstName", "First Name"])

                for i, (dk, dinfo) in enumerate(file_items):
                    pct = int(10 + (i / max(len(file_items), 1)) * 55)
                    prog.progress(pct, text=f"Uploading {dinfo.get('filename', dk)}…")
                    sp_url = upload_file_to_sharepoint(
                        student_id2, dk, dinfo["file_bytes"], dinfo["filename"],
                        last_name=_upload_last, first_name=_upload_first,
                    )
                    # Store list of uploads per slot key
                    if dk not in uploaded_sp:
                        uploaded_sp[dk] = []
                    uploaded_sp[dk].append({
                        "filename": dinfo["filename"],
                        "url":      sp_url,
                        "size_kb":  dinfo.get("size_kb", 0),
                    })

                # Phase 2: Build submission payload
                prog.progress(70, text="Preparing registration record…")
                confirmation2  = f"REG-2027-{str(uuid.uuid4().int)[-5:]}"
                courses_val    = academic2.get("courses") or ", ".join(get_student_courses(s2))

                # Human-readable category labels for each doc slot
                _DOC_CATEGORY_LABELS = {
                    "bls":             "Section 1 — BLS Certificate",
                    "health":          "Section 1 — Health Clearance",
                    "id_front":        "Section 1 — ID Badge Front",
                    "id_back":         "Section 1 — ID Badge Back",
                    "access_front":    "Section 1 — Access Card Front",
                    "access_back":     "Section 1 — Access Card Back",
                    "tuition_receipt": "Section 2 — Tuition Payment Receipt",
                    "pregnancy_1":     "Section 3 — Dignity in Pregnancy Part 1",
                    "pregnancy_2":     "Section 3 — Dignity in Pregnancy Part 2",
                    "pregnancy_3":     "Section 3 — Dignity in Pregnancy Part 3",
                }

                # Build structured document record: location + category per file
                doc_files_structured = {}
                for slot_key, file_entries in uploaded_sp.items():
                    category_label = _DOC_CATEGORY_LABELS.get(slot_key, slot_key)
                    enriched = []
                    for fe in file_entries:
                        enriched.append({
                            "filename":  fe["filename"],
                            "location":  fe["url"],          # full SharePoint URL
                            "category":  category_label,     # which section/type
                            "size_kb":   fe["size_kb"],
                        })
                    doc_files_structured[slot_key] = enriched

                doc_files_json     = json.dumps(doc_files_structured, indent=2)
                total_files_uploaded = sum(len(v) for v in uploaded_sp.values())

                submission_fields = {
                    "Title":                f"{full_name2}({student_id2})",
                    "SubmittedAt":           datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "ConfirmationNumber":    confirmation2,
                    "Status":               "Submitted - Doc Review",
                    "StudentID":            student_id2,
                    "FullName":             full_name2,
                    "CNumber":              safe_get(s2, "C#"),
                    "DHSEmail":             first_available_value(s2, ["DHS Email","DHSEmail","School Email"]),
                    "PersonalEmail":        contact2.get("personal_email") or first_available_value(s2, ["Personal Email","Email"]),
                    "Phone":                contact2.get("phone") or first_available_value(s2, ["Phone1","Phone"]),
                    "Address":              full_address2,
                    "EC1Name":              contact2.get("ec1_name",""),
                    "EC1Relationship":      contact2.get("ec1_relationship",""),
                    "EC1Cell":              contact2.get("ec1_cell",""),
                    "EC1Work":              contact2.get("ec1_work",""),
                    "Semester":             semester2,
                    "Courses":              courses_val,
                    "Tuition":              tuition2,
                    "FinancialAid":         financial2.get("financial_aid") or "No",
                    "FAFSA":               financial2.get("fafsa") or "N/A",
                    "PaymentElection":      financial2.get("payment_election") or "Self-Pay",
                    "RequiredDocsUploaded": f"{total_uploaded}/{total_required} slots ({total_files_uploaded} files)",
                    "OptionalDocsUploaded": "0",
                    "DocumentFilesJSON":    doc_files_json,
                }

                prog.progress(82, text="Saving registration record to SharePoint…")
                submit_to_registration_list(submission_fields)

                prog.progress(100, text="✅ Submitted successfully!")
                st.session_state.confirmation_number = confirmation2
                st.session_state.submitted           = True
                st.session_state.uploaded_docs_sp    = uploaded_sp
                st.session_state.step = 7
                st.rerun()

            except Exception as err:
                prog.empty()
                erbox.error(f"❌ Submission failed: {type(err).__name__}: {err}")

    render_page_end()


# ==========================================
# STEP 7 — SUCCESS
# ==========================================
elif st.session_state.step == 7:
    s = st.session_state.student_data
    render_navbar(student_name=get_full_name(s), student_id=get_student_id(s))
    st.markdown('<div class="oes-page success-page">', unsafe_allow_html=True)

    confirmation   = st.session_state.confirmation_number or "REG-2027-00000"
    semester       = st.session_state.academic_data.get("semester") or get_student_semester(s)
    personal_email = st.session_state.contact_data.get("personal_email") or first_available_value(s, ["Personal Email","Email"])
    dhs_email      = first_available_value(s, ["DHS Email","DHSEmail","School Email"])
    identity       = st.session_state.identity_data  or {}
    contact        = st.session_state.contact_data   or {}
    academic       = st.session_state.academic_data  or {}
    financial      = st.session_state.financial_data or {}
    uploaded_sp    = st.session_state.get("uploaded_docs_sp", {})

    full_name = " ".join(filter(None, [
        identity.get("first_name")  or first_available_value(s, ["FirstName","First Name"]),
        identity.get("middle_name") or first_available_value(s, ["MiddleName","Middle Name"]),
        identity.get("last_name")   or first_available_value(s, ["LastName","Last Name"]),
    ])).strip() or get_full_name(s)

    st.markdown(f"""
    <div class="success-card">
        <div class="success-icon">✅</div>
        <div class="success-title">Registration Submitted Successfully!</div>
        <div class="confirmation-pill">{html.escape(confirmation)}</div>
        <div class="success-note">Save this confirmation number for your records</div>
        <p>Your registration for <b>2027-II {html.escape(semester)}</b> has been submitted and is pending OES review.
        A confirmation will be sent to <b>{html.escape(dhs_email)}</b> and <b>{html.escape(personal_email)}</b>.</p>
        <div class="next-box">
            <div class="next-title">WHAT HAPPENS NEXT</div>
            <div class="next-item"><span>1</span> Registrar staff will review your submission and documents within <b>3–5 business days.</b></div>
            <div class="next-item"><span>2</span> Finance staff will review your payment election.</div>
            <div class="next-item"><span>3</span> If documents are missing or expired, you will receive an email with a resubmission deadline.</div>
            <div class="next-item"><span>4</span> Once approved, you will receive a <b>Registered</b> status email with your confirmed course schedule.</div>
        </div>
        <div class="status-title">Registration Review Status</div>
        <div class="review-status-bar">
            <div class="status-step done"><span></span>Submitted</div>
            <div class="status-line green"></div>
            <div class="status-step active"><span></span>Doc Review</div>
            <div class="status-line"></div>
            <div class="status-step"><span></span>Fin Review</div>
            <div class="status-line"></div>
            <div class="status-step"><span></span>Approved</div>
            <div class="status-line"></div>
            <div class="status-step"><span></span>Registered</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-line-title" style="max-width:780px;margin:28px auto 10px"><span>YOUR SUBMISSION SUMMARY</span></div>', unsafe_allow_html=True)

    summary_card("IDENTITY", [
        ("FULL NAME",    full_name),
        ("STUDENT ID",   get_student_id(s)),
        ("ACCESS CARD #", identity.get("access_card") or get_access_card(s)),
        ("DHS EMAIL",    dhs_email),
        ("CONFIRMATION #", confirmation),
    ])

    full_address = ", ".join(v for v in [
        contact.get("address"), contact.get("city"),
        contact.get("state"),   contact.get("zip")
    ] if v)
    summary_card("CONTACT & EMERGENCY", [
        ("HOME ADDRESS",     full_address),
        ("PHONE",            contact.get("phone")),
        ("PERSONAL EMAIL",   contact.get("personal_email")),
        ("EC1 NAME",         contact.get("ec1_name")),
        ("EC1 RELATIONSHIP", contact.get("ec1_relationship")),
        ("EC1 CELL",         contact.get("ec1_cell")),
        ("EC1 WORK",         contact.get("ec1_work")),
    ])

    summary_card("ACADEMICS", [
        ("SEMESTER",  academic.get("semester") or semester),
        ("COURSES",   academic.get("courses") or ", ".join(get_student_courses(s))),
        ("CONFIRMED", "✓ Yes"),
    ])

    summary_card("FINANCIAL", [
        ("TUITION",          financial.get("tuition")),
        ("FINANCIAL AID",    financial.get("financial_aid") or "No"),
        ("FAFSA",            financial.get("fafsa") or "N/A"),
        ("PAYMENT ELECTION", financial.get("payment_election") or "Self-Pay"),
    ])

    if uploaded_sp:
        doc_label_map = {
            "bls":             "BLS Certificate",
            "health":          "Health Clearance",
            "id_front":        "ID Badge Front",
            "id_back":         "ID Badge Back",
            "access_front":    "Access Card Front",
            "access_back":     "Access Card Back",
            "tuition_receipt": "Tuition Payment Receipt",
            "pregnancy_1":     "Dignity in Pregnancy Part 1",
            "pregnancy_2":     "Dignity in Pregnancy Part 2",
            "pregnancy_3":     "Dignity in Pregnancy Part 3",
        }
        doc_rows = ""
        for dk, file_entries in uploaded_sp.items():
            label = doc_label_map.get(dk, dk)
            # file_entries is a list of dicts
            if isinstance(file_entries, dict):
                file_entries = [file_entries]  # backward compat
            links_html = ""
            for finfo in file_entries:
                url   = finfo.get("url","")
                fname = finfo.get("filename","")
                size  = finfo.get("size_kb","")
                link  = (f'<a href="{html.escape(url)}" target="_blank" style="color:#1d4ed8;">📎 {html.escape(fname)} ({size} KB)</a>'
                         if url else f"📎 {html.escape(fname)} ({size} KB)")
                links_html += f'<div style="margin:2px 0;">{link}</div>'
            doc_rows += (f'<div class="review-field" style="grid-column:span 2;">'
                         f'<div class="review-label">{html.escape(label)}</div>'
                         f'<div class="review-value">{links_html}</div></div>')
        st.markdown(f"""
        <div class="review-card" style="max-width:780px;margin:0 auto 12px;">
            <div class="review-card-header"><span>DOCUMENTS UPLOADED TO SHAREPOINT</span></div>
            <div class="review-grid">{doc_rows}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col_c, col_btn, col_d = st.columns([1,2,1])
    with col_btn:
        if st.button("🔄 Start New Registration", type="primary"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            init_state()
            st.rerun()
    render_page_end()


else:
    st.error("Unknown step. Please restart.")
    if st.button("Restart"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        init_state()
        st.rerun()


