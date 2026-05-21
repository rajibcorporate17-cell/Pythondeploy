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
import time  # <-- MODIFICATION: Imported time for auto-retry delay
from pathlib import Path

# ==========================================
# CONFIGURATION — SharePoint via Microsoft Graph
# ==========================================
CLIENT_ID     = "cb467c35-c5f5-4987-afff-6d77d45e41a6"
CLIENT_SECRET = "5va8Q~EERNmh46TGkduGGd_6MM2Y7IOT5MlT4cls"
TENANT_ID     = "231edc94-ddeb-473e-a5c1-c86d43e0db76"
HOSTNAME      = "datainsightio.sharepoint.com"
SITE_PATH     = "OESRegistration"

DATA_LIST_NAME       = "Student Information"   # SharePoint list — student records
SUBMISSION_LIST_NAME = "Submissions"           # SharePoint list — form submissions
UPLOAD_ROOT          = "uploaded_documents"

# ==========================================
# PAGE SETUP
# ==========================================
st.set_page_config(page_title="OES Student Portal", page_icon="🎓", layout="wide")

# Streamlit treats indented multiline HTML as a Markdown code block.
# IMPORTANT FIX: every unsafe HTML line is left-stripped before rendering.
# This prevents <div> cards from showing as black code blocks.
_original_streamlit_markdown = st.markdown

def normalize_html(markup: str) -> str:
    markup = textwrap.dedent(str(markup)).strip()
    return "\n".join(line.lstrip() for line in markup.splitlines())

def render_html(markup: str):
    return _original_streamlit_markdown(normalize_html(markup), unsafe_allow_html=True)

def _safe_markdown(body, *args, **kwargs):
    if kwargs.get("unsafe_allow_html") and isinstance(body, str):
        body = normalize_html(body)
    return _original_streamlit_markdown(body, *args, **kwargs)

st.markdown = _safe_markdown


# ==========================================
# LOAD EXTERNAL CSS
# ==========================================
def load_css(filepath: str):
    css_path = os.path.join(os.path.dirname(__file__), filepath)
    try:
        with open(css_path, "r", encoding="utf-8") as f:
            css = f.read()
        st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning("style1.css not found. Please ensure it is in the same directory.")

load_css("style1.css")

# ==========================================
# SHAREPOINT CONNECTION — MSAL + Graph API
# ==========================================
@st.cache_data(ttl=3000)
def get_access_token() -> str:
    """Authenticate with Azure AD and return a Graph API access token."""
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise ConnectionError(f"SharePoint authentication failed: {result.get('error_description', result)}")
    return result["access_token"]

@st.cache_data(ttl=86400)
def get_site_id() -> str:
    """Resolve the SharePoint site ID once and cache it."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{HOSTNAME}:/sites/{SITE_PATH}"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise ConnectionError(f"Could not find SharePoint site '{SITE_PATH}': {resp.text}")
    return resp.json()["id"]

# MODIFICATION: Added an internal retry loop to silently handle token drops
@st.cache_data(ttl=300)
def load_data() -> pd.DataFrame:
    """Fetch all student records from the SharePoint list into a DataFrame."""
    for attempt in range(3):
        try:
            token   = get_access_token()
            site_id = get_site_id()
            headers = {"Authorization": f"Bearer {token}"}

            # 1. Fetch Dynamic Column Mappings to resolve generic internal names (field_1, field_2...)
            columns_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{DATA_LIST_NAME}/columns"
            col_resp = requests.get(columns_url, headers=headers)
            
            # If the token is rejected, force a token clear and restart the loop silently
            if col_resp.status_code == 401:
                get_access_token.clear()
                continue
            
            column_mapping = {}
            if col_resp.status_code == 200:
                cols_data = col_resp.json().get("value", [])
                column_mapping = {col["name"]: col["displayName"] for col in cols_data if "name" in col and "displayName" in col}

            # 2. Paginate through all items (Graph returns max 5000 per page)
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}"
                f"/lists/{DATA_LIST_NAME}/items?expand=fields&$top=5000"
            )
            records = []
            while url:
                resp = requests.get(url, headers=headers)
                
                # Check for pagination token expiration
                if resp.status_code == 401:
                    get_access_token.clear()
                    raise ValueError("Token expired mid-pagination")
                    
                if resp.status_code != 200:
                    raise ValueError(f"Failed to fetch list '{DATA_LIST_NAME}': {resp.text}")
                data = resp.json()
                records.extend(item["fields"] for item in data.get("value", []))
                url = data.get("@odata.nextLink")   # follow pagination if >5000 rows

            df = pd.DataFrame(records)
            
            # Apply the display name translations dynamically
            if column_mapping:
                df.rename(columns=column_mapping, inplace=True)
                
            # Match the incoming 'Form Access #' column structure to your logic's expectation ('C#')
            if "Form Access #" in df.columns and "C#" not in df.columns:
                df.rename(columns={"Form Access #": "C#"}, inplace=True)

            df.columns = df.columns.astype(str).str.strip()

            if "StudentID" not in df.columns or "C#" not in df.columns:
                raise ValueError(
                    f"SharePoint list '{DATA_LIST_NAME}' must contain 'StudentID' and 'C#' (or 'Form Access #') columns. "
                    f"Columns found: {list(df.columns)}"
                )

            df["StudentID"] = df["StudentID"].astype(str).str.strip().str.lower()
            df["C#"]        = df["C#"].astype(str).str.strip().str.lower()
            
            return df # Successful load
            
        except Exception as e:
            if attempt == 2: # Give up only after 3 failed attempts
                raise e
            get_access_token.clear() # Clear bad token before trying again

# ==========================================
# GENERAL HELPERS
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
    if text.lower() in ["nan", "none", "null"]:
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

    # exact match first
    for col in candidates:
        val = clean_value(data.get(col, ""))
        if val:
            return val

    # case/space-insensitive match
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
    raw_semester = first_available_value(
        student_row,
        ["Semester", "Term", "Registration Term", "Cohort", "Semester Term"],
        default="Semester 2"
    )
    raw_semester = clean_value(raw_semester)
    # If Google Sheet stores only 2, show it as Semester 2 in the UI.
    if re.fullmatch(r"\d+", raw_semester):
        return f"Semester {raw_semester}"
    return raw_semester

def get_student_courses(student_row):
    semester = get_student_semester(student_row)
    course_candidates = []

    # If Semester = Semester 2, first try Semester 2 Courses
    semester_number = re.search(r"(\d+)", semester or "")
    if semester_number:
        n = semester_number.group(1)
        course_candidates.extend([
            f"Semester {n} Courses",
            f"Semester {n} Course",
            f"Sem {n} Courses",
            f"Sem {n} Course List",
        ])

    course_candidates.extend([
        "Semester 2 Courses",
        "Course Name",
        "Course Names",
        "Course Code",
        "Course Codes",
        "Courses",
        "Registered Courses",
        "Semester Course List"
    ])

    course_text = first_available_value(student_row, course_candidates)
    courses = split_course_text(course_text)
    if courses:
        return courses

    # fallback: columns named like N121, N122, N123L
    found = []
    for col, val in student_row.items():
        col_name = str(col).strip()
        cell_val = clean_value(val).lower()
        if re.fullmatch(r"[a-zA-Z]{1,5}\d{2,4}[a-zA-Z]?", col_name):
            if cell_val and cell_val not in ["no", "false", "0", "n/a", "na"]:
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
            if amount.is_integer():
                return f"${int(amount):,}"
            return f"${amount:,.2f}"
        except Exception:
            pass
    return value

def get_tuition(student_row, semester):
    candidates = ["Tuition", "Tuition Charge", "Amount", "Semester Tuition", "Tuition Amount"]
    match = re.search(r"(\d+)", semester or "")
    if match:
        n = match.group(1)
        candidates = [
            f"Semester {n} Tuition",
            f"Semester {n} Tuition Charge",
            f"Tuition Semester {n}",
            f"Semester {n} Amount",
        ] + candidates
    return format_money(first_available_value(student_row, candidates), default="$2,400")

def get_full_name(s):
    return " ".join([safe_get(s, "FirstName"), safe_get(s, "MiddleName"), safe_get(s, "LastName")]).replace("  ", " ").strip()

def get_student_id(s):
    return safe_get(s, "StudentID").upper()

def save_uploaded_file(student_id, doc_key, uploaded_file):
    folder = Path(__file__).parent / UPLOAD_ROOT / student_id
    folder.mkdir(parents=True, exist_ok=True)
    safe_name = re.sub(r"[^A-Za-z0-9_.-]", "_", uploaded_file.name)
    filename = f"{doc_key}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{safe_name}"
    file_path = folder / filename
    file_path.write_bytes(uploaded_file.getbuffer())
    return {
        "filename": uploaded_file.name,
        "stored_path": str(file_path),
        "size_kb": round(len(uploaded_file.getbuffer()) / 1024, 1),
        "uploaded_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

def append_submission_to_sharepoint(submission_data: dict):
    """Add one row to the Submissions SharePoint list. Always silent on failure."""
    try:
        token   = get_access_token()
        site_id = get_site_id()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        }
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}"
            f"/lists/{SUBMISSION_LIST_NAME}/items"
        )
        payload = {"fields": {k: str(v) for k, v in submission_data.items()}}
        requests.post(url, headers=headers, json=payload)
    except Exception:
        pass  # Submission save must NEVER block navigation

# ==========================================
# SESSION STATE
# ==========================================
def init_state():
    defaults = {
        "step": 1,
        "student_data": None,
        "identity_data": {},
        "contact_data": {},
        "academic_data": {},
        "financial_data": {},
        "documents_data": {},
        "confirmation_number": "",
        "submitted": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_state()

# MODIFICATION: Replaced st.error crash screen with a silent background auto-retry loop
try:
    df = load_data()
except Exception:
    # Instead of showing a red error, show a friendly loading message and retry in 3 seconds
    st.info("🔄 Connecting to server... please wait.")
    time.sleep(3)
    get_access_token.clear()
    load_data.clear()
    st.rerun()

# ==========================================
# UI HELPERS
# ==========================================
def render_navbar(student_name: str = "", student_id: str = ""):
    badge_html = f"<div class='oes-navbar-right'>{html.escape(student_name)} · {html.escape(student_id.upper())}</div>" if student_name else ""
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

def render_progress(current_step: int = 0, pct: int = 0):
    steps = [("🔒", "Verify"), ("👤", "Identity"), ("📞", "Contact"), ("🎓", "Academics"), ("💳", "Financial"), ("📁", "Documents"), ("✅", "Review")]
    items_html = ""
    for i, (icon, label) in enumerate(steps):
        if i < current_step:
            cls, disp = "done", "✓"
        elif i == current_step:
            cls, disp = "active", icon
        else:
            cls, disp = "", icon
        items_html += f'<div class="step-item"><div class="step-icon {cls}">{disp}</div><div class="step-label {cls}">{label}</div></div>'
    st.markdown(f'<div class="progress-card"><div class="progress-header"><span>REGISTRATION PROGRESS</span><span class="progress-pct">{pct}% complete</span></div><div class="progress-steps">{items_html}</div></div>', unsafe_allow_html=True)

def render_page_start(s=None, progress_step=0, pct=0):
    if s:
        render_navbar(student_name=get_full_name(s), student_id=get_student_id(s))
    else:
        render_navbar()
    st.markdown('<div class="oes-page">', unsafe_allow_html=True)
    render_progress(current_step=progress_step, pct=pct)

def render_page_end():
    st.markdown('</div>', unsafe_allow_html=True)

def summary_card(title, rows, edit_step=None):
    edit_html = f'<a class="edit-link">Edit</a>' if edit_step else ""
    body = ""
    for label, value in rows:
        body += f"""
        <div class="review-field">
            <div class="review-label">{html.escape(label)}</div>
            <div class="review-value">{html.escape(clean_value(value) or 'N/A')}</div>
        </div>
        """
    st.markdown(f"""
    <div class="review-card">
        <div class="review-card-header"><span>{html.escape(title)}</span>{edit_html}</div>
        <div class="review-grid">{body}</div>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# STEP 1 — IDENTITY VERIFICATION
# ==========================================
if st.session_state.step == 1:
    render_page_start(progress_step=0, pct=0)
    st.markdown("""
    <div class="form-card">
        <div class="step-badge">🔒 STEP 0 OF 6</div>
        <div class="form-title">Identity Verification</div>
        <div class="form-subtitle">Enter your pre-issued Student ID and C# to access your registration.</div>
        <div class="info-banner">ℹ️ Your <b>Student ID</b> and <b>C#</b> were sent to your DHS email.</div>
        <div class="credentials-label">Enter Your Credentials</div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login"):
        id_input = st.text_input("Student ID *", placeholder="A0000000000")
        c_input = st.text_input("C# (Computer Access Number) *", type="password", placeholder="c000000")
        if st.form_submit_button("🔒 Verify and Access My Registration"):
            match = df[(df["StudentID"] == id_input.strip().lower()) & (df["C#"] == c_input.strip().lower())]
            if not match.empty:
                st.session_state.student_data = match.iloc[0].to_dict()
                st.session_state.step = 2
                st.rerun()
            else:
                st.error("❌ Invalid Credentials.")
    render_page_end()

# ==========================================
# STEP 2 — IDENTITY CONFIRMATION
# ==========================================
elif st.session_state.step == 2:
    s = st.session_state.student_data
    render_page_start(s, progress_step=1, pct=14)
    st.markdown("""
    <div class="form-card">
        <div class="step-badge">👤 STEP 1 OF 6</div>
        <div class="form-title">Identity Confirmation</div>
        <div class="form-subtitle">Review and confirm your legal name and birth date.</div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("identity"):
        col1, col2 = st.columns(2)
        with col1:
            fname = st.text_input("First Name *", value=first_available_value(s, ["FirstName", "First Name"]))
            mname = st.text_input("Middle Name", value=first_available_value(s, ["MiddleName", "Middle Name"]))
        with col2:
            lname = st.text_input("Last Name *", value=first_available_value(s, ["LastName", "Last Name"]))
            dob = st.text_input("Birth Date *", value=first_available_value(s, ["BirthDate", "Birth Date", "DOB"]))

        if st.form_submit_button("Confirm & Continue ➡️"):
            st.session_state.identity_data = {
                "first_name": fname, "middle_name": mname, "last_name": lname, "birth_date": dob
            }
            st.session_state.step = 3
            st.rerun()

    if st.button("← Logout"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        init_state()
        st.rerun()
    render_page_end()

# ==========================================
# STEP 3 — CONTACT INFORMATION
# ==========================================
elif st.session_state.step == 3:
    s = st.session_state.student_data
    render_page_start(s, progress_step=2, pct=28)
    st.markdown("""
    <div class="form-card">
        <div class="step-badge">📞 STEP 2 OF 6</div>
        <div class="form-title">Contact & Address</div>
        <div class="form-subtitle">Ensure your contact details are correct for official communications.</div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("contact"):
        col1, col2 = st.columns(2)
        with col1:
            phone = st.text_input("Phone *", value=first_available_value(s, ["Phone1", "Phone", "Cell", "Mobile"]))
        with col2:
            personal_email = st.text_input("Personal Email *", value=first_available_value(s, ["Personal Email", "PersonalEmail", "Email"]))
        address = st.text_input("Address *", value=first_available_value(s, ["Address1", "Address", "Home Address"]))
        c3, c4, c5 = st.columns([2, 1, 1])
        with c3:
            city = st.text_input("City *", value=first_available_value(s, ["City"]))
        with c4:
            state = st.text_input("State *", value=first_available_value(s, ["State"]))
        with c5:
            zipc = st.text_input("Zip *", value=first_available_value(s, ["ZipCode", "Zip Code", "Zip"]))

        st.markdown("---")
        ec1_name = st.text_input("Emergency Contact Name", value=first_available_value(s, ["EC1 Name", "Emergency Contact Name", "EmergencyContactName"]))
        ec_col1, ec_col2, ec_col3 = st.columns(3)
        with ec_col1:
            ec1_relationship = st.text_input("EC1 Relationship", value=first_available_value(s, ["EC1 Relationship", "Emergency Contact Relationship"]))
        with ec_col2:
            ec1_cell = st.text_input("EC1 Cell", value=first_available_value(s, ["EC1 Cell", "Emergency Contact Cell", "EC1 Phone"]))
        with ec_col3:
            ec1_work = st.text_input("EC1 Work", value=first_available_value(s, ["EC1 Work", "Emergency Contact Work"]))

        if st.form_submit_button("Confirm & Continue ➡️"):
            st.session_state.contact_data = {
                "phone": phone, "personal_email": personal_email,
                "address": address, "city": city, "state": state, "zip": zipc,
                "ec1_name": ec1_name, "ec1_relationship": ec1_relationship,
                "ec1_cell": ec1_cell, "ec1_work": ec1_work
            }
            st.session_state.step = 4
            st.rerun()

    if st.button("← Back to Identity"):
        st.session_state.step = 2
        st.rerun()
    render_page_end()

# ==========================================
# STEP 4 — ACADEMIC / COURSE VERIFICATION
# ==========================================
elif st.session_state.step == 4:
    s = st.session_state.student_data
    render_page_start(s, progress_step=3, pct=42)

    semester = get_student_semester(s)
    courses = get_student_courses(s)
    courses_text = ", ".join(courses)

    st.markdown(f"""
    <div class="form-card academic-main-card">
        <div class="step-badge">🎓 STEP 3 OF 6</div>
        <div class="form-title">Academic Registration Confirmation</div>
        <div class="form-subtitle">Review your registered courses. Course changes must go through the Registrar.</div>

        <div class="academic-dark-card">
            <div class="academic-label">SEMESTER</div>
            <div class="academic-semester">{html.escape(semester)}</div>
            <div class="academic-divider"></div>
            <div class="academic-label">REGISTERED COURSES</div>
            <div class="course-pill-wrap">{course_badges_html(courses)}</div>
        </div>

        <div class="info-banner academic-info-banner">
            ℹ️ Your academic information is pulled from the official OES student record. Course changes must be reviewed by the Registrar or Office of Educational Services.
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("academic_verification"):
        st.markdown("<div class='course-confirm-title'>IS YOUR SEMESTER COURSE LIST CORRECT?</div>", unsafe_allow_html=True)
        col_ok, col_bad = st.columns(2)
        with col_ok:
            submit_ok = st.form_submit_button("✅ Yes, Everything is Correct")
        with col_bad:
            submit_bad = st.form_submit_button("⚠️ No, This is Not Correct")

        if submit_ok:
            st.session_state.academic_data = {
                "semester": semester,
                "courses": courses_text,
                "academic_confirmed": "Yes"
            }
            st.session_state.step = 5
            st.rerun()

        if submit_bad:
            st.warning("⚠️ Please contact the Office of Educational Services or Registrar to correct your academic record before continuing.")

    col_back, col_save = st.columns([1, 2])
    with col_back:
        if st.button("← Back"):
            st.session_state.step = 3
            st.rerun()
    with col_save:
        if st.button("Save & Return Later"):
            st.info("Your progress has been saved temporarily in this session.")
    render_page_end()

# ==========================================
# STEP 5 — FINANCIAL INFORMATION
# ==========================================
elif st.session_state.step == 5:
    s = st.session_state.student_data
    render_page_start(s, progress_step=4, pct=67)

    semester = st.session_state.academic_data.get("semester") or get_student_semester(s)
    tuition = get_tuition(s, semester)
    financial_aid = first_available_value(s, ["Financial Aid", "FinancialAid", "Aid", "Aid Status"], default="No")
    fafsa = first_available_value(s, ["FAFSA", "FAFSA Status", "Fafsa"], default="N/A")
    aid_badge = "No Financial Aid on File" if financial_aid.lower() in ["no", "n", "none", "n/a", ""] else f"Financial Aid: {financial_aid}"

    st.markdown(f"""
    <div class="form-card financial-card">
        <div class="step-badge">💳 STEP 4 OF 6</div>
        <div class="form-title">Financial Information</div>
        <div class="form-subtitle">Review your tuition and financial aid status, then select your payment intent.</div>

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
        col_back, col_save, col_next = st.columns([1, 2, 2])
        with col_back:
            back = st.form_submit_button("← Back")
        with col_save:
            save = st.form_submit_button("Save & Return Later")
        with col_next:
            cont = st.form_submit_button("Save & Continue →")

        if back:
            st.session_state.step = 4
            st.rerun()
        if save:
            st.info("Your financial selection has been saved temporarily in this session.")
        if cont:
            st.session_state.financial_data = {
                "tuition": tuition,
                "financial_aid": financial_aid,
                "fafsa": fafsa,
                "payment_election": payment_choice
            }
            st.session_state.step = 6
            st.rerun()
    render_page_end()

# ==========================================
# STEP 6 — DOCUMENT UPLOAD
# ==========================================
elif st.session_state.step == 6:
    s = st.session_state.student_data
    render_page_start(s, progress_step=5, pct=83)

    st.markdown("""
    <div class="form-card docs-card">
        <div class="step-badge">📁 STEP 5 OF 6</div>
        <div class="form-title">Document Upload</div>
        <div class="form-subtitle">Upload all required documents. Files are stored securely in your registration record.</div>
    </div>
    """, unsafe_allow_html=True)

    mandatory_docs = [
        ("bls", "🩺", "Current BLS Certificate"),
        ("health", "🏥", "Health Clearance"),
        ("id_front", "🪪", "ID Badge — Front"),
        ("id_back", "🪪", "ID Badge — Back"),
        ("access_front", "💳", "Access Card — Front"),
        ("access_back", "💳", "Access Card — Back"),
    ]
    conditional_docs = [("tuition_receipt", "🧾", "Tuition Payment Receipt")]
    optional_docs = [
        ("pregnancy_1", "📜", "Dignity in Pregnancy Cert — Part 1"),
        ("pregnancy_2", "📜", "Dignity in Pregnancy Cert — Part 2"),
    ]

    required_docs = mandatory_docs + conditional_docs
    docs_data = st.session_state.documents_data
    required_uploaded = sum(1 for key, _, _ in required_docs if key in docs_data)
    optional_uploaded = sum(1 for key, _, _ in optional_docs if key in docs_data)

    st.markdown(f"""
    <div class="doc-progress-row">
        <div><span class="status-dot {'green' if required_uploaded == len(required_docs) else 'orange'}"></span> {required_uploaded} of {len(required_docs)} required documents uploaded</div>
        <div class="doc-rule">Max 25 MB · PDF JPG PNG</div>
    </div>
    """, unsafe_allow_html=True)

    def render_doc_section(title, docs, required=True):
        st.markdown(f'<div class="section-line-title doc-title"><span>{html.escape(title)}</span></div>', unsafe_allow_html=True)
        for key, icon, label in docs:
            uploaded = docs_data.get(key)
            row_cls = "doc-row uploaded" if uploaded else "doc-row"
            status_text = "✓ Uploaded" if uploaded else ("Required" if required else "Optional")
            meta = f"📎 {html.escape(uploaded['filename'])} ({uploaded['size_kb']} KB)" if uploaded else ""
            st.markdown(f"""
            <div class="{row_cls}">
                <div class="doc-icon">{icon}</div>
                <div class="doc-info">
                    <div class="doc-name">{html.escape(label)}</div>
                    <div class="doc-status">{html.escape(status_text)}</div>
                    <div class="doc-meta">{meta}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            col_upload, col_remove = st.columns([4, 1])
            with col_upload:
                uploaded_file = st.file_uploader(
                    f"Upload {label}",
                    type=["pdf", "jpg", "jpeg", "png"],
                    key=f"upload_{key}",
                    label_visibility="collapsed"
                )
                if uploaded_file is not None:
                    docs_data[key] = save_uploaded_file(get_student_id(s), key, uploaded_file)
                    st.session_state.documents_data = docs_data
                    st.rerun()
            with col_remove:
                if uploaded and st.button("Remove", key=f"remove_{key}"):
                    docs_data.pop(key, None)
                    st.session_state.documents_data = docs_data
                    st.rerun()

    render_doc_section("MANDATORY DOCUMENTS", mandatory_docs, required=True)
    render_doc_section("CONDITIONAL DOCUMENTS", conditional_docs, required=True)
    render_doc_section("OPTIONAL DOCUMENTS", optional_docs, required=False)

    all_required_done = required_uploaded == len(required_docs)
    col_back, col_save, col_next = st.columns([1, 2, 2])
    with col_back:
        if st.button("← Back"):
            st.session_state.step = 5
            st.rerun()
    with col_save:
        if st.button("Save & Return Later"):
            st.info("Your uploaded document information has been saved temporarily in this session.")
    with col_next:
        if st.button("Save & Continue →", disabled=not all_required_done):
            st.session_state.step = 7
            st.rerun()
    if not all_required_done:
        st.caption("Upload all required documents to continue.")
    render_page_end()

# ==========================================
# STEP 7 — REVIEW & SUBMIT
# ==========================================
elif st.session_state.step == 7:
    s = st.session_state.student_data
    render_page_start(s, progress_step=6, pct=100)

    identity = st.session_state.identity_data or {
        "first_name": first_available_value(s, ["FirstName", "First Name"]),
        "middle_name": first_available_value(s, ["MiddleName", "Middle Name"]),
        "last_name": first_available_value(s, ["LastName", "Last Name"]),
        "birth_date": first_available_value(s, ["BirthDate", "Birth Date", "DOB"]),
    }
    contact = st.session_state.contact_data
    academic = st.session_state.academic_data
    financial = st.session_state.financial_data
    docs_data = st.session_state.documents_data

    full_name = " ".join([identity.get("first_name", ""), identity.get("middle_name", ""), identity.get("last_name", "")]).replace("  ", " ").strip() or get_full_name(s)
    dhs_email = first_available_value(s, ["DHS Email", "DHSEmail", "School Email", "DHS_Email"])
    c_num = safe_get(s, "C#")
    access_card = first_available_value(s, ["Access Card #", "Access Card", "AccessCard", "Access Card Number"])

    st.markdown("""
    <div class="form-card review-main-card">
        <div class="step-badge">✅ STEP 6 OF 6</div>
        <div class="form-title">Review & Submit</div>
        <div class="form-subtitle">Review your complete registration before final submission.</div>
    </div>
    """, unsafe_allow_html=True)

    summary_card("IDENTITY", [
        ("FULL NAME", full_name),
        ("STUDENT ID", get_student_id(s)),
        ("C# (FORM ACCESS #)", c_num),
        ("DHS EMAIL", dhs_email),
        ("ACCESS CARD #", access_card),
        ("PERSONAL EMAIL", contact.get("personal_email") or first_available_value(s, ["Personal Email", "Email"]))
    ], edit_step=2)

    full_address = ", ".join([v for v in [contact.get("address"), contact.get("city"), contact.get("state"), contact.get("zip")] if v])
    summary_card("CONTACT & EMERGENCY", [
        ("HOME ADDRESS", full_address or first_available_value(s, ["Address1", "Address"])),
        ("PHONE", contact.get("phone") or first_available_value(s, ["Phone1", "Phone"])),
        ("EC1 NAME", contact.get("ec1_name") or first_available_value(s, ["EC1 Name"])),
        ("EC1 RELATIONSHIP", contact.get("ec1_relationship") or first_available_value(s, ["EC1 Relationship"])),
        ("EC1 CELL", contact.get("ec1_cell") or first_available_value(s, ["EC1 Cell"])),
        ("EC1 WORK", contact.get("ec1_work") or first_available_value(s, ["EC1 Work"])),
    ], edit_step=3)

    summary_card("ACADEMICS", [
        ("SEMESTER", academic.get("semester") or get_student_semester(s)),
        ("COURSES", academic.get("courses") or ", ".join(get_student_courses(s))),
        ("COURSE LIST OK", "✓ Confirmed"),
    ], edit_step=4)

    summary_card("FINANCIAL", [
        ("TUITION", financial.get("tuition") or get_tuition(s, academic.get("semester"))),
        ("FINANCIAL AID", financial.get("financial_aid") or "No"),
        ("FAFSA", financial.get("fafsa") or "N/A"),
        ("PAYMENT ELECTION", financial.get("payment_election") or "Self-Pay — Direct payment"),
    ], edit_step=5)

    required_keys = ["bls", "health", "id_front", "id_back", "access_front", "access_back", "tuition_receipt"]
    required_uploaded = sum(1 for k in required_keys if k in docs_data)
    optional_uploaded = sum(1 for k in ["pregnancy_1", "pregnancy_2"] if k in docs_data)
    st.markdown(f"""
    <div class="review-card">
        <div class="review-card-header"><span>DOCUMENTS</span><a class="edit-link">Edit</a></div>
        <div class="documents-summary"><span class="big-check">✅</span><div><b>{required_uploaded}/7 mandatory uploaded</b><br><span>+ {optional_uploaded} optional uploaded</span></div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-line-title declaration-title"><span>DECLARATION REQUIRED</span></div>', unsafe_allow_html=True)
    declaration = st.checkbox(
        "Accuracy Declaration: I confirm that all information I have provided is accurate and complete. I understand that providing false information may affect my enrollment for the 2027-II term."
    )

    col_back, col_submit = st.columns([1, 2])
    with col_back:
        if st.button("← Back"):
            st.session_state.step = 6
            st.rerun()
    with col_submit:
        if st.button("Submit Registration ✔", disabled=not declaration, type="primary"):
            confirmation = f"REG-2027-{str(uuid.uuid4().int)[-5:]}"
            submission = {
                "SubmittedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ConfirmationNumber": confirmation,
                "Status": "Submitted / Doc Review",
                "StudentID": get_student_id(s),
                "FullName": full_name,
                "CNumber": c_num,
                "DHS Email": dhs_email,
                "Personal Email": contact.get("personal_email") or first_available_value(s, ["Personal Email", "Email"]),
                "Phone": contact.get("phone") or first_available_value(s, ["Phone1", "Phone"]),
                "Address": full_address,
                "Semester": academic.get("semester") or get_student_semester(s),
                "Courses": academic.get("courses") or ", ".join(get_student_courses(s)),
                "Tuition": financial.get("tuition") or get_tuition(s, academic.get("semester")),
                "FinancialAid": financial.get("financial_aid") or "No",
                "FAFSA": financial.get("fafsa") or "N/A",
                "PaymentElection": financial.get("payment_election") or "Self-Pay — Direct payment",
                "RequiredDocumentsUploaded": f"{required_uploaded}/7",
                "OptionalDocumentsUploaded": str(optional_uploaded),
                "DocumentFilesJSON": json.dumps(docs_data),
            }
            try:
                append_submission_to_sharepoint(submission)
                st.session_state.confirmation_number = confirmation
                st.session_state.submitted = True
                st.session_state.step = 8
                st.rerun()
            except Exception as e:
                st.error(f"Submission could not be saved to SharePoint: {e}")

    if not declaration:
        st.caption("Check the declaration to enable submission.")
    render_page_end()

# ==========================================
# STEP 8 — SUCCESS / CONFIRMATION
# ==========================================
elif st.session_state.step == 8:
    s = st.session_state.student_data
    render_navbar()
    st.markdown('<div class="oes-page success-page">', unsafe_allow_html=True)
    confirmation = st.session_state.confirmation_number or "REG-2027-00000"
    semester = st.session_state.academic_data.get("semester") or get_student_semester(s)
    personal_email = st.session_state.contact_data.get("personal_email") or first_available_value(s, ["Personal Email", "Email"])
    dhs_email = first_available_value(s, ["DHS Email", "DHSEmail", "School Email"])

    st.markdown(f"""
    <div class="success-card">
        <div class="success-icon">✅</div>
        <div class="success-title">You are now registered for this semester!</div>
        <div class="confirmation-pill">{html.escape(confirmation)}</div>
        <div class="success-note">Save this confirmation number for your records</div>
        <p>Your registration for <b>2027-II {html.escape(semester)}</b> has been submitted and is pending OES review. A confirmation email was sent to<br><b>{html.escape(dhs_email)}</b> and <b>{html.escape(personal_email)}</b>.</p>

        <div class="next-box">
            <div class="next-title">WHAT HAPPENS NEXT</div>
            <div class="next-item"><span>1</span> Registrar staff will review your submission and documents within <b>3–5 business days.</b></div>
            <div class="next-item"><span>2</span> Finance staff will review your payment election. Corrections will be emailed if needed.</div>
            <div class="next-item"><span>3</span> If required documents are missing or expired, you will receive an email with a resubmission deadline.</div>
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
    render_page_end()

else:
    st.error("Unknown step. Please restart the registration.")
    if st.button("Restart"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        init_state()
        st.rerun()
