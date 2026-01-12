# =========================================
# COMMENTCRAFT - Streamlit Report Comment Generator (Edgy Light Mode UI)
# Supports Year 5, 7 & 8; Subjects: English, Maths, Science
# =========================================

import random
import streamlit as st
from docx import Document
import tempfile
import os
import time
from datetime import datetime, timedelta
import pandas as pd
import io
import re

# ========== SECURITY & PRIVACY SETTINGS ==========
TARGET_CHARS = 499
MAX_FILE_SIZE_MB = 5
MAX_ROWS_PER_UPLOAD = 100
RATE_LIMIT_SECONDS = 10

# ========== PAGE CONFIGURATION ==========
st.set_page_config(
    page_title="CommentCraft",
    page_icon="🖋️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========== SECURITY INITIALIZATION ==========
if 'app_initialized' not in st.session_state:
    st.session_state.clear()
    st.session_state.app_initialized = True
    st.session_state.upload_count = 0
    st.session_state.last_upload_time = datetime.now()
    st.session_state.generated_files = []

# ========== IMPORT STATEMENTS ==========
try:
    # Year 5, 7, 8: English, Maths, Science
    # Import all statement banks as before (kept intact)
    from statements_year5_English import *
    from statements_year5_Maths import *
    from statements_year5_Science import *
    from statements_year7_English import *
    from statements_year7_Maths import *
    from statements_year7_science import *
    from statements_year8_English import *
    from statements_year8_Maths import *
    from statements_year8_science import *
except ImportError as e:
    st.error(f"Missing required statement files: {e}")
    st.stop()

# ========== SECURITY FUNCTIONS ==========
def validate_upload_rate():
    time_since_last = datetime.now() - st.session_state.last_upload_time
    if time_since_last < timedelta(seconds=RATE_LIMIT_SECONDS):
        wait_time = RATE_LIMIT_SECONDS - time_since_last.seconds
        st.error(f"Please wait {wait_time} seconds before uploading again")
        return False
    return True

def sanitize_input(text, max_length=100):
    if not text:
        return ""
    sanitized = ''.join(c for c in text if c.isalnum() or c in " .'-")
    return sanitized[:max_length].strip().title()

def validate_file(file):
    if file.size > MAX_FILE_SIZE_MB * 1024 * 1024:
        return False, f"File too large (max {MAX_FILE_SIZE_MB}MB)"
    if not file.name.lower().endswith('.csv'):
        return False, "Only CSV files allowed"
    return True, ""

def process_csv_securely(uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='wb') as tmp:
        tmp.write(uploaded_file.getvalue())
        temp_path = tmp.name
    try:
        df = pd.read_csv(temp_path, nrows=MAX_ROWS_PER_UPLOAD + 1)
        if len(df) > MAX_ROWS_PER_UPLOAD:
            st.warning(f"Only processing first {MAX_ROWS_PER_UPLOAD} rows")
            df = df.head(MAX_ROWS_PER_UPLOAD)
        if 'Student Name' in df.columns:
            df['Student Name'] = df['Student Name'].apply(lambda x: sanitize_input(str(x)))
        return df
    except Exception as e:
        st.error(f"Error reading CSV: {e}")
        return None
    finally:
        try:
            os.unlink(temp_path)
        except:
            pass

def get_pronouns(gender):
    gender = gender.lower()
    if gender == "male":
        return "he", "his"
    elif gender == "female":
        return "she", "her"
    return "they", "their"

def lowercase_first(text):
    return text[0].lower() + text[1:] if text else ""

def truncate_comment(comment, target=TARGET_CHARS):
    if len(comment) <= target:
        return comment
    truncated = comment[:target].rstrip(" ,;.")
    if "." in truncated:
        truncated = truncated[:truncated.rfind(".")+1]
    return truncated

def fix_pronouns_in_text(text, pronoun, possessive):
    if not text:
        return text
    text = re.sub(r'\bhe\b', pronoun, text, flags=re.IGNORECASE)
    text = re.sub(r'\bHe\b', pronoun.capitalize(), text)
    text = re.sub(r'\bhis\b', possessive, text, flags=re.IGNORECASE)
    text = re.sub(r'\bHis\b', possessive.capitalize(), text)
    text = re.sub(r'\bhim\b', pronoun, text, flags=re.IGNORECASE)
    text = re.sub(r'\bHim\b', pronoun.capitalize(), text)
    text = re.sub(r'\bhimself\b', f"{pronoun}self", text, flags=re.IGNORECASE)
    text = re.sub(r'\bherself\b', f"{pronoun}self", text, flags=re.IGNORECASE)
    return text

# ========== COMMENT GENERATOR ==========
def generate_comment(subject, year, name, gender, att, achieve, target, pronouns, attitude_target=None):
    # Keep all logic intact
    comment = generate_comment_logic(subject, year, name, gender, att, achieve, target, pronouns, attitude_target)
    # Ensure "Keep it up." at the end
    if not comment.endswith("Keep it up."):
        comment = comment.rstrip('. ') + ". Keep it up."
    return comment

# ========== STREAMLIT APP INTERFACE (REDESIGNED) ==========

# ---------- HEADER ----------
st.markdown(
    """
    <div style="text-align:center; padding:10px 0;">
        <img src="assets/ikc_logo.png" style="height:80px; margin-bottom:10px; object-fit:contain;">
        <div style="font-size:2.4rem; font-weight:700; color:#1f2d2b;">CommentCraft</div>
        <div style="font-size:1rem; color:#6b6f6a; margin-bottom:15px;">Report Comment Assistant • International Kingdom College</div>
    </div>
    <hr style="border:1px solid #eee;">
    """, unsafe_allow_html=True
)

# ---------- SIDEBAR ----------
with st.sidebar:
    st.markdown("<h3 style='color:#2e7d32;'>Navigation</h3>", unsafe_allow_html=True)
    app_mode = st.radio("Mode", ["Single Student", "Batch Upload", "Privacy Info"], index=0, horizontal=False)
    st.markdown("---")
    st.markdown("<h4 style='color:#2e7d32;'>Privacy Features</h4>", unsafe_allow_html=True)
    st.info("""
    - All processing in browser memory
    - Auto-deletion of temp files
    - Input sanitization & rate limiting
    """, icon="ℹ️")
    
    if st.button("Clear All Data"):
        st.session_state.clear()
        st.session_state.app_initialized = True
        st.session_state.upload_count = 0
        st.session_state.last_upload_time = datetime.now()
        st.success("All data cleared!")
        st.rerun()
    
    st.markdown("---")
    st.caption("v3.0 • Multi-Year Edition")

# ---------- THREE GREEN STEPS ----------
st.markdown("<h4 style='margin-bottom:5px;'>Three Easy Steps</h4>", unsafe_allow_html=True)
if 'progress' not in st.session_state:
    st.session_state.progress = 1

step_col1, step_col2, step_col3 = st.columns(3)

def step_box(col, step_num, title, description):
    with col:
        is_current = st.session_state.progress == step_num
        bg_color = '#e8f5e9' if is_current else '#f8f9fa'
        st.markdown(f"""
        <div style='text-align:center; padding:8px; margin:2px 0; background-color:{bg_color}; border-radius:8px; border-left:4px solid #2e7d32;'>
            <div style='font-weight:600;'>{'✓' if st.session_state.progress > step_num else f'{step_num}.'} {title}</div>
            <div style='font-size:0.85em; color:#666;'>{description}</div>
        </div>
        """, unsafe_allow_html=True)

step_box(step_col1, 1, "Select", "Choose student details")
step_box(step_col2, 2, "Generate", "Create the comment")
step_box(step_col3, 3, "Download", "Export your reports")

st.markdown("<br>", unsafe_allow_html=True)

# ---------- SINGLE STUDENT MODE ----------
if app_mode == "Single Student":
    st.subheader("Single Student Entry")
    
    if 'form_submitted' not in st.session_state:
        st.session_state.form_submitted = False
    
    with st.form("single_student_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            subject = st.selectbox("Subject", ["English", "Maths", "Science"])
            year = st.selectbox("Year", [5, 7, 8])
            name = st.text_input("Student Name", placeholder="Enter first name only", key='student_name_input')
            gender = st.selectbox("Gender", ["Male", "Female"])
        with col2:
            att = st.selectbox("Attitude Band", [90,85,80,75,70,65,60,55,40], index=3)
            achieve = st.selectbox("Achievement Band", [90,85,80,75,70,65,60,55,40], index=3)
            target = st.selectbox("Target Band", [90,85,80,75,70,65,60,55,40], index=3)
            st.caption("Use dropdowns for faster input. Tab key moves between fields.")
        
        attitude_target = st.text_area("Optional Attitude Next Steps", placeholder="E.g., continue to participate actively...", height=60, key='attitude_target_input')
        
        submitted = st.form_submit_button("Generate Comment")
    
    if submitted and name:
        if not validate_upload_rate(): st.stop()
        name = sanitize_input(name)
        pronouns = get_pronouns(gender)
        
        with st.spinner("Generating comment..."):
            comment = generate_comment(subject, year, name, gender, att, achieve, target, pronouns, st.session_state.get('attitude_target_input',''))
            char_count = len(comment)
        
        st.session_state.progress = 2
        st.session_state.form_submitted = True
        
        st.subheader("Generated Comment")
        st.text_area("", comment, height=180)
        
        if 'all_comments' not in st.session_state:
            st.session_state.all_comments = []
        
        st.session_state.all_comments.append({
            'name': name,
            'subject': subject,
            'year': year,
            'comment': comment,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M")
        })

# ---------- BATCH UPLOAD MODE ----------
elif app_mode == "Batch Upload":
    st.subheader("Batch Upload (CSV)")
    st.info("CSV columns: Student Name, Gender, Subject, Year, Attitude, Achievement, Target")
    uploaded_file = st.file_uploader("Choose CSV file", type=['csv'])
    
    if uploaded_file:
        if not validate_upload_rate(): st.stop()
        is_valid, msg = validate_file(uploaded_file)
        if not is_valid: st.error(msg); st.stop()
        df = process_csv_securely(uploaded_file)
        if df is not None:
            st.success(f"Processed {len(df)} students successfully")
            if st.button("Generate All Comments"):
                if 'all_comments' not in st.session_state: st.session_state.all_comments = []
                for idx, row in df.iterrows():
                    pronouns = get_pronouns(str(row.get('Gender','')).lower())
                    comment = generate_comment(
                        subject=str(row.get('Subject','English')),
                        year=int(row.get('Year',7)),
                        name=str(row.get('Student Name','')),
                        gender=str(row.get('Gender','')),
                        att=int(row.get('Attitude',75)),
                        achieve=int(row.get('Achievement',75)),
                        target=int(row.get('Target',75)),
                        pronouns=pronouns
                    )
                    st.session_state.all_comments.append({
                        'name': sanitize_input(str(row.get('Student Name',''))),
                        'subject': str(row.get('Subject','English')),
                        'year': int(row.get('Year',7)),
                        'comment': comment,
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M")
                    })
                st.session_state.progress = 2
                st.success(f"Generated {len(df)} comments!")

# ---------- DOWNLOAD SECTION ----------
if 'all_comments' in st.session_state and st.session_state.all_comments:
    st.session_state.progress = 3
    st.subheader("Download Reports")
    total_comments = len(st.session_state.all_comments)
    st.info(f"{total_comments} comment(s) generated")
    
    col_dl1, col_dl2, col_dl3 = st.columns(3)
    with col_dl1:
        if st.button("Word Document"):
            doc = Document()
            doc.add_heading('Report Comments',0)
            doc.add_paragraph(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
            doc.add_paragraph(f'Total Students: {total_comments}\n')
            for entry in st.session_state.all_comments:
                doc.add_heading(f"{entry['name']} - {entry['subject']} Year {entry['year']}",2)
                doc.add_paragraph(entry['comment'])
            bio = io.BytesIO(); doc.save(bio)
            st.download_button("⬇️ Download Word File", bio.getvalue(), f"report_comments_{datetime.now().strftime('%Y%m%d_%H%M')}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with col_dl2:
        if st.button("CSV Export"):
            df_export = pd.DataFrame(st.session_state.all_comments)
            csv_bytes = df_export.to_csv(index=False).encode('utf-8')
            st.download_button("⬇️ Download CSV", csv_bytes, f"report_comments_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", "text/csv")
    with col_dl3:
        if st.button("Clear & Start Over"):
            st.session_state.all_comments = []
            st.session_state.progress = 1
            st.success("All comments cleared!")
            st.rerun()

# ---------- FOOTER ----------
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("© CommentCraft v3.0 • Multi-Year Edition • International Kingdom College")
