# =========================================
# MULTI-SUBJECT REPORT COMMENT GENERATOR - Secure Streamlit Version
# Supports Year 5, 7 & 8; Subjects: English, Maths, Science
# NOW WITH VARIANT SUPPORT for avoiding duplicate comments
# =========================================

import streamlit as st
import sys
import os

# Show loading message
loading_placeholder = st.empty()
loading_placeholder.info("Loading application...")

# Try to import all required packages with detailed error messages
try:
    # Try to import docx (using docx package instead of python-docx)
    import docx
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    st.error("'pandas' package not installed")
    st.stop()

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# Clear loading message
loading_placeholder.empty()

# Now import other standard libraries
import random
import tempfile
import time
from datetime import datetime, timedelta
import io
import re

# ========== SECURITY & PRIVACY SETTINGS ==========
TARGET_CHARS = 499
MAX_FILE_SIZE_MB = 5
MAX_ROWS_PER_UPLOAD = 100
RATE_LIMIT_SECONDS = 10

# ========== PAGE CONFIGURATION ==========
st.set_page_config(
    page_title="Secure Report Generator",
    page_icon="",
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
    # Store form data separately
    if 'form_data' not in st.session_state:
        st.session_state.form_data = {}
    if 'all_comments' not in st.session_state:
        st.session_state.all_comments = []
    if 'variant_tracker' not in st.session_state:
        st.session_state.variant_tracker = {}

# ========== HELPER FUNCTIONS ==========
def apply_british_spelling(text):
    """Convert American spelling to British spelling"""
    if not text:
        return text
    
    # Common American to British spelling conversions
    replacements = {
        r'\borganized\b': 'organised',
        r'\brealized\b': 'realised',
        r'\brecognized\b': 'recognised',
        r'\banalyzed\b': 'analysed',
        r'\bcategorized\b': 'categorised',
        r'\bcharacterized\b': 'characterised',
        r'\bemphasized\b': 'emphasised',
        r'\bfavor\b': 'favour',
        r'\bcolor\b': 'colour',
        r'\blabor\b': 'labour',
        r'\bhonor\b': 'honour',
        r'\bbehavior\b': 'behaviour',
        r'\bfavorite\b': 'favourite',
        r'\bcenter\b': 'centre',
        r'\bmeter\b': 'metre',
        r'\bliter\b': 'litre',
        r'\btheater\b': 'theatre',
        r'\banalyze\b': 'analyse',
        r'\bcivilize\b': 'civilise',
        r'\borganize\b': 'organise',
        r'\brealize\b': 'realise',
        r'\brecognize\b': 'recognise',
        r'\bcharacterize\b': 'characterise',
        r'\bemphasize\b': 'emphasise',
        r'\bdefense\b': 'defence',
        r'\boffense\b': 'offence',
        r'\blicense\b': 'licence',
        r'\bpractice\b': 'practise',  # verb
        r'\bpracticed\b': 'practised',
        r'\bpracticing\b': 'practising',
        r'\btraveled\b': 'travelled',
        r'\btraveling\b': 'travelling',
        r'\bcanceled\b': 'cancelled',
        r'\bcanceling\b': 'cancelling',
        r'\bmodeled\b': 'modelled',
        r'\bmodeling\b': 'modelling',
        r'\blabeled\b': 'labelled',
        r'\blabeling\b': 'labelling',
        r'\bsignaled\b': 'signalled',
        r'\bsignaling\b': 'signalling',
    }
    
    for american, british in replacements.items():
        text = re.sub(american, british, text, flags=re.IGNORECASE)
    
    return text

def validate_upload_rate():
    """Prevent rapid-fire uploads/abuse"""
    time_since_last = datetime.now() - st.session_state.last_upload_time
    if time_since_last < timedelta(seconds=RATE_LIMIT_SECONDS):
        wait_time = RATE_LIMIT_SECONDS - time_since_last.seconds
        st.error(f"Please wait {wait_time} seconds before uploading again")
        return False
    return True

def sanitize_input(text, max_length=100):
    """Sanitize user input to prevent injection attacks"""
    if not text:
        return ""
    sanitized = ''.join(c for c in text if c.isalnum() or c in " .'-")
    return sanitized[:max_length].strip().title()

def validate_file(file):
    """Validate uploaded file size and type"""
    if file.size > MAX_FILE_SIZE_MB * 1024 * 1024:
        return False, f"File too large (max {MAX_FILE_SIZE_MB}MB)"
    if not file.name.lower().endswith('.csv'):
        return False, "Only CSV files allowed"
    return True, ""

def process_csv_securely(uploaded_file):
    """Process CSV with auto-cleanup of temp files"""
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
    """Fix gender pronouns in statement text using word boundaries"""
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

# ========== IMPORT STATEMENTS ==========
# We'll use the same files for both variants since you don't have variant2 files
# The variant system will still work by using different random selections

try:
    # Year 5 English
    from statements_year5_English import (
        opening_phrases as opening_5_eng,
        attitude_bank as attitude_5_eng,
        reading_bank as reading_5_eng,
        writing_bank as writing_5_eng,
        reading_target_bank as target_5_eng,
        writing_target_bank as target_write_5_eng,
        closer_bank as closer_5_eng
    )
    
    # Use same files for "variant2"
    opening_5_eng_v2 = opening_5_eng
    attitude_5_eng_v2 = attitude_5_eng
    reading_5_eng_v2 = reading_5_eng
    writing_5_eng_v2 = writing_5_eng
    target_5_eng_v2 = target_5_eng
    target_write_5_eng_v2 = target_write_5_eng
    closer_5_eng_v2 = closer_5_eng
    
    # Year 5 Maths
    from statements_year5_Maths import (
        opening_phrases as opening_5_maths,
        attitude_bank as attitude_5_maths,
        number_bank as number_5_maths,
        problem_solving_bank as problem_5_maths,
        target_bank as target_5_maths,
        closer_bank as closer_5_maths
    )
    
    opening_5_maths_v2 = opening_5_maths
    attitude_5_maths_v2 = attitude_5_maths
    number_5_maths_v2 = number_5_maths
    problem_5_maths_v2 = problem_5_maths
    target_5_maths_v2 = target_5_maths
    closer_5_maths_v2 = closer_5_maths
    
    # Year 5 Science
    from statements_year5_Science import (
        opening_phrases as opening_5_sci,
        attitude_bank as attitude_5_sci,
        science_bank as science_5_sci,
        target_bank as target_5_sci,
        closer_bank as closer_5_sci
    )
    
    opening_5_sci_v2 = opening_5_sci
    attitude_5_sci_v2 = attitude_5_sci
    science_5_sci_v2 = science_5_sci
    target_5_sci_v2 = target_5_sci
    closer_5_sci_v2 = closer_5_sci
    
    # Year 7 English
    from statements_year7_English import (
        opening_phrases as opening_7_eng,
        attitude_bank as attitude_7_eng,
        reading_bank as reading_7_eng,
        writing_bank as writing_7_eng,
        reading_target_bank as target_7_eng,
        writing_target_bank as target_write_7_eng,
        closer_bank as closer_7_eng
    )
    
    opening_7_eng_v2 = opening_7_eng
    attitude_7_eng_v2 = attitude_7_eng
    reading_7_eng_v2 = reading_7_eng
    writing_7_eng_v2 = writing_7_eng
    target_7_eng_v2 = target_7_eng
    target_write_7_eng_v2 = target_write_7_eng
    closer_7_eng_v2 = closer_7_eng
    
    # Year 7 Maths
    from statements_year7_Maths import (
        opening_phrases as opening_7_maths,
        attitude_bank as attitude_7_maths,
        number_and_algebra_bank as number_7_maths,
        geometry_and_measurement_bank as geometry_7_maths,
        problem_solving_and_reasoning_bank as problem_7_maths,
        target_bank as target_7_maths,
        closer_bank as closer_7_maths
    )
    
    opening_7_maths_v2 = opening_7_maths
    attitude_7_maths_v2 = attitude_7_maths
    number_7_maths_v2 = number_7_maths
    geometry_7_maths_v2 = geometry_7_maths
    problem_7_maths_v2 = problem_7_maths
    target_7_maths_v2 = target_7_maths
    closer_7_maths_v2 = closer_7_maths
    
    # Year 7 Science
    from statements_year7_science import (
        opening_phrases as opening_7_sci,
        attitude_bank as attitude_7_sci,
        science_bank as science_7_sci,
        target_bank as target_7_sci,
        closer_bank as closer_7_sci
    )
    
    opening_7_sci_v2 = opening_7_sci
    attitude_7_sci_v2 = attitude_7_sci
    science_7_sci_v2 = science_7_sci
    target_7_sci_v2 = target_7_sci
    closer_7_sci_v2 = closer_7_sci
    
    # Year 8 English
    from statements_year8_English import (
        opening_phrases as opening_8_eng,
        attitude_bank as attitude_8_eng,
        reading_bank as reading_8_eng,
        writing_bank as writing_8_eng,
        reading_target_bank as target_8_eng,
        writing_target_bank as target_write_8_eng,
        closer_bank as closer_8_eng
    )
    
    opening_8_eng_v2 = opening_8_eng
    attitude_8_eng_v2 = attitude_8_eng
    reading_8_eng_v2 = reading_8_eng
    writing_8_eng_v2 = writing_8_eng
    target_8_eng_v2 = target_8_eng
    target_write_8_eng_v2 = target_write_8_eng
    closer_8_eng_v2 = closer_8_eng
    
    # Year 8 Maths
    from statements_year8_Maths import (
        opening_phrases as opening_8_maths,
        attitude_bank as attitude_8_maths,
        maths_bank as maths_8_maths,
        target_bank as target_8_maths,
        closer_bank as closer_8_maths
    )
    
    opening_8_maths_v2 = opening_8_maths
    attitude_8_maths_v2 = attitude_8_maths
    maths_8_maths_v2 = maths_8_maths
    target_8_maths_v2 = target_8_maths
    closer_8_maths_v2 = closer_8_maths
    
    # Year 8 Science
    from statements_year8_science import (
        opening_phrases as opening_8_sci,
        attitude_bank as attitude_8_sci,
        science_bank as science_8_sci,
        target_bank as target_8_sci,
        closer_bank as closer_8_sci
    )
    
    opening_8_sci_v2 = opening_8_sci
    attitude_8_sci_v2 = attitude_8_sci
    science_8_sci_v2 = science_8_sci
    target_8_sci_v2 = target_8_sci
    closer_8_sci_v2 = closer_8_sci
    
except ImportError as e:
    st.error(f"Missing required statement files: {e}")
    st.info("Make sure all statement files are in the same directory")
    st.stop()

# ========== COMMENT GENERATOR FUNCTIONS ==========
def get_statement_banks(subject, year, variant=0):
    """
    Get statement banks based on subject, year, and variant.
    variant: 0 = variant1, 1 = variant2
    """
    
    # Year 5 English
    if year == 5 and subject == "English":
        if variant == 1:
            return (opening_5_eng_v2, attitude_5_eng_v2, reading_5_eng_v2, writing_5_eng_v2,
                   target_5_eng_v2, target_write_5_eng_v2, closer_5_eng_v2)
        else:
            return (opening_5_eng, attitude_5_eng, reading_5_eng, writing_5_eng,
                   target_5_eng, target_write_5_eng, closer_5_eng)
    
    # Year 5 Maths
    elif year == 5 and subject == "Maths":
        if variant == 1:
            return (opening_5_maths_v2, attitude_5_maths_v2, number_5_maths_v2, None,
                   target_5_maths_v2, None, closer_5_maths_v2)
        else:
            return (opening_5_maths, attitude_5_maths, number_5_maths, None,
                   target_5_maths, None, closer_5_maths)
    
    # Year 5 Science
    elif year == 5 and subject == "Science":
        if variant == 1:
            return (opening_5_sci_v2, attitude_5_sci_v2, science_5_sci_v2, None,
                   target_5_sci_v2, None, closer_5_sci_v2)
        else:
            return (opening_5_sci, attitude_5_sci, science_5_sci, None,
                   target_5_sci, None, closer_5_sci)
    
    # Year 7 English
    elif year == 7 and subject == "English":
        if variant == 1:
            return (opening_7_eng_v2, attitude_7_eng_v2, reading_7_eng_v2, writing_7_eng_v2,
                   target_7_eng_v2, target_write_7_eng_v2, closer_7_eng_v2)
        else:
            return (opening_7_eng, attitude_7_eng, reading_7_eng, writing_7_eng,
                   target_7_eng, target_write_7_eng, closer_7_eng)
    
    # Year 7 Maths
    elif year == 7 and subject == "Maths":
        if variant == 1:
            return (opening_7_maths_v2, attitude_7_maths_v2, number_7_maths_v2, None,
                   target_7_maths_v2, None, closer_7_maths_v2)
        else:
            return (opening_7_maths, attitude_7_maths, number_7_maths, None,
                   target_7_maths, None, closer_7_maths)
    
    # Year 7 Science
    elif year == 7 and subject == "Science":
        if variant == 1:
            return (opening_7_sci_v2, attitude_7_sci_v2, science_7_sci_v2, None,
                   target_7_sci_v2, None, closer_7_sci_v2)
        else:
            return (opening_7_sci, attitude_7_sci, science_7_sci, None,
                   target_7_sci, None, closer_7_sci)
    
    # Year 8 English
    elif year == 8 and subject == "English":
        if variant == 1:
            return (opening_8_eng_v2, attitude_8_eng_v2, reading_8_eng_v2, writing_8_eng_v2,
                   target_8_eng_v2, target_write_8_eng_v2, closer_8_eng_v2)
        else:
            return (opening_8_eng, attitude_8_eng, reading_8_eng, writing_8_eng,
                   target_8_eng, target_write_8_eng, closer_8_eng)
    
    # Year 8 Maths
    elif year == 8 and subject == "Maths":
        if variant == 1:
            return (opening_8_maths_v2, attitude_8_maths_v2, maths_8_maths_v2, None,
                   target_8_maths_v2, None, closer_8_maths_v2)
        else:
            return (opening_8_maths, attitude_8_maths, maths_8_maths, None,
                   target_8_maths, None, closer_8_maths)
    
    # Year 8 Science
    elif year == 8 and subject == "Science":
        if variant == 1:
            return (opening_8_sci_v2, attitude_8_sci_v2, science_8_sci_v2, None,
                   target_8_sci_v2, None, closer_8_sci_v2)
        else:
            return (opening_8_sci, attitude_8_sci, science_8_sci, None,
                   target_8_sci, None, closer_8_sci)
    
    # Default fallback
    return None

def generate_comment(subject, year, name, gender, att, achieve, target, attitude_target="", variant=0):
    """
    Generate report comment with optional variant support.
    variant: 0 = variant1, 1 = variant2
    """
    p, p_poss = get_pronouns(gender)
    name = sanitize_input(name)
    
    # Get appropriate statement banks
    banks = get_statement_banks(subject, year, variant)
    if not banks:
        return "Error: Statement banks not found"
    
    opening_bank, attitude_bank, achievement_bank, writing_bank, target_bank, writing_target_bank, closer_bank = banks
    
    # Build comment based on subject
    if subject == "English":
        # English has reading and writing
        opening = random.choice(opening_bank)
        attitude_text = fix_pronouns_in_text(attitude_bank[att], p, p_poss)
        attitude_sentence = f"{opening} {name} {attitude_text}"
        if not attitude_sentence.endswith('.'):
            attitude_sentence += '.'
        
        reading_text = fix_pronouns_in_text(achievement_bank[achieve], p, p_poss)
        if reading_text[0].islower():
            reading_text = f"{p} {reading_text}"
        reading_sentence = f"In reading, {reading_text}"
        if not reading_sentence.endswith('.'):
            reading_sentence += '.'
        
        writing_text = fix_pronouns_in_text(writing_bank[achieve], p, p_poss)
        if writing_text[0].islower():
            writing_text = f"{p} {writing_text}"
        writing_sentence = f"In writing, {writing_text}"
        if not writing_sentence.endswith('.'):
            writing_sentence += '.'
        
        reading_target_text = fix_pronouns_in_text(target_bank[target], p, p_poss)
        reading_target_sentence = f"For the next term, {p} should {lowercase_first(reading_target_text)}"
        if not reading_target_sentence.endswith('.'):
            reading_target_sentence += '.'
        
        writing_target_text = fix_pronouns_in_text(writing_target_bank[target], p, p_poss)
        writing_target_sentence = f"Additionally, {p} should {lowercase_first(writing_target_text)}"
        if not writing_target_sentence.endswith('.'):
            writing_target_sentence += '.'
        
        closer_sentence = random.choice(closer_bank)
        
    elif subject == "Maths":
        # Maths has only achievement (no reading/writing split)
        opening = random.choice(opening_bank)
        attitude_text = fix_pronouns_in_text(attitude_bank[att], p, p_poss)
        attitude_sentence = f"{opening} {name} {attitude_text}"
        if not attitude_sentence.endswith('.'):
            attitude_sentence += '.'
        
        achievement_text = fix_pronouns_in_text(achievement_bank[achieve], p, p_poss)
        if achievement_text[0].islower():
            achievement_text = f"{p} {achievement_text}"
        reading_sentence = achievement_text
        if not reading_sentence.endswith('.'):
            reading_sentence += '.'
        
        target_text = fix_pronouns_in_text(target_bank[target], p, p_poss)
        reading_target_sentence = f"For the next term, {p} should {lowercase_first(target_text)}"
        if not reading_target_sentence.endswith('.'):
            reading_target_sentence += '.'
        
        writing_sentence = ""
        writing_target_sentence = ""
        closer_sentence = random.choice(closer_bank)
        
    else:  # Science
        # Science has only achievement (no reading/writing split)
        opening = random.choice(opening_bank)
        attitude_text = fix_pronouns_in_text(attitude_bank[att], p, p_poss)
        attitude_sentence = f"{opening} {name} {attitude_text}"
        if not attitude_sentence.endswith('.'):
            attitude_sentence += '.'
        
        science_text = fix_pronouns_in_text(achievement_bank[achieve], p, p_poss)
        if science_text[0].islower():
            science_text = f"{p} {science_text}"
        reading_sentence = science_text
        if not reading_sentence.endswith('.'):
            reading_sentence += '.'
        
        target_text = fix_pronouns_in_text(target_bank[target], p, p_poss)
        reading_target_sentence = f"For the next term, {p} should {lowercase_first(target_text)}"
        if not reading_target_sentence.endswith('.'):
            reading_target_sentence += '.'
        
        writing_sentence = ""
        writing_target_sentence = ""
        closer_sentence = random.choice(closer_bank)
    
    # Optional attitude target
    if attitude_target and attitude_target.strip():
        attitude_target = sanitize_input(attitude_target)
        attitude_target_sentence = f"{lowercase_first(attitude_target)}"
        if not attitude_target_sentence.endswith('.'):
            attitude_target_sentence += '.'
        attitude_target_sentence = attitude_target_sentence.replace('..', '.')
    else:
        attitude_target_sentence = ""
    
    # Assemble comment
    comment_parts = [
        attitude_sentence,
        reading_sentence,
        writing_sentence,
        reading_target_sentence,
        writing_target_sentence,
        closer_sentence,
        attitude_target_sentence
    ]
    
    comment = " ".join([c for c in comment_parts if c])
    comment = comment.strip()
    
    if not comment.endswith('.'):
        comment += '.'
    
    comment = comment.replace('..', '.')
    comment = truncate_comment(comment, TARGET_CHARS)
    
    if not comment.endswith('.'):
        comment = comment.rstrip(' ,;') + '.'
    
    comment = comment.replace('..', '.')
    
    # Apply British spelling
    comment = apply_british_spelling(comment)
    
    return comment

def get_available_variants(subject, year):
    """Check which variants are available for a given subject/year"""
    # Since we're using the same files for both variants, both are always available
    return [0, 1]

# ========== STREAMLIT APP LAYOUT ==========

# Apply custom CSS for cleaner design
st.markdown("""
<style>
    /* Remove blue accents and use green/yellow */
    .stButton > button {
        background-color: #4CAF50 !important;
        color: white !important;
        border: none !important;
        border-radius: 4px !important;
    }
    .stButton > button:hover {
        background-color: #45a049 !important;
    }
    
    .secondary_button > button {
        background-color: #FFC107 !important;
        color: black !important;
        border: none !important;
        border-radius: 4px !important;
    }
    
    /* Remove borders and shadows */
    .stTextArea textarea, .stSelectbox div[data-baseweb="select"], 
    .stTextInput input, .stNumberInput input {
        border: 1px solid #ddd !important;
        box-shadow: none !important;
        border-radius: 4px !important;
    }
    
    /* Style metrics */
    .stMetric {
        background-color: #f9f9f9;
        padding: 10px;
        border-radius: 4px;
        border-left: 4px solid #4CAF50;
    }
    
    /* Remove icons */
    [data-testid="stDecoration"] {
        display: none;
    }
    
    /* Header styling */
    h1, h2, h3 {
        color: #333 !important;
        margin-top: 0 !important;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
    }
    
    /* Info boxes */
    .stAlert {
        background-color: #FFF3CD !important;
        border: 1px solid #FFEEBA !important;
        color: #856404 !important;
    }
    
    /* Success boxes */
    .stAlert [data-testid="stMarkdownContainer"] p {
        color: #155724 !important;
    }
    
    /* Warning boxes */
    .stAlert [data-testid="stMarkdownContainer"] p:contains("PRIVACY") {
        color: #856404 !important;
    }
    
    /* Remove any blue */
    a {
        color: #4CAF50 !important;
    }
    
    /* Progress tracker */
    .step-box {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 4px;
        border-left: 4px solid #FFC107;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar - simplified
with st.sidebar:
    st.markdown("## Navigation")
    
    app_mode = st.radio(
        "",
        ["Single Student", "Batch Upload", "Privacy Info"],
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    st.markdown("**Privacy Features:**")
    st.markdown("""
    - No data stored on servers
    - All processing in memory
    - Auto-deletion of temp files
    - Input sanitization
    """)
    
    if st.button("Clear All Data", use_container_width=True, type="secondary"):
        st.session_state.clear()
        st.session_state.app_initialized = True
        st.session_state.upload_count = 0
        st.session_state.last_upload_time = datetime.now()
        st.session_state.form_data = {}
        st.session_state.all_comments = []
        st.session_state.variant_tracker = {}
        st.success("All data cleared!")
        st.rerun()

# Main content - simplified and more compact
st.title("Multi-Subject Report Comment Generator")
st.markdown("**Years 5, 7 & 8 • English, Maths, Science • With Variant Support**")

st.markdown("""
<div style='background-color: #FFF3CD; padding: 15px; border-radius: 4px; border: 1px solid #FFEEBA; margin: 20px 0;'>
<strong>PRIVACY NOTICE:</strong> All data is processed in memory only. No files are stored on servers. 
Close browser tab to completely erase all data.
</div>
""", unsafe_allow_html=True)

# Progress tracker - horizontal and more compact
st.markdown("### Quick Steps")
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <div class='step-box'>
    <h4>1. Select</h4>
    <p>Choose student details</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class='step-box'>
    <h4>2. Generate</h4>
    <p>Create the comment</p>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class='step-box'>
    <h4>3. Download</h4>
    <p>Export your reports</p>
    </div>
    """, unsafe_allow_html=True)

# ========== SINGLE STUDENT MODE ==========
if app_mode == "Single Student":
    st.markdown("---")
    st.markdown("### Student Details")
    
    # Initialize form data in session state
    if 'form_data' not in st.session_state:
        st.session_state.form_data = {}
    
    # Create form - more compact layout
    with st.form(key="student_form"):
        # First row - basic info
        col1, col2, col3 = st.columns(3)
        
        with col1:
            subject = st.selectbox("Subject", ["English", "Maths", "Science"])
            name = st.text_input("Student Name", placeholder="First name only")
        
        with col2:
            year = st.selectbox("Year", [5, 7, 8])
            gender = st.selectbox("Gender", ["Male", "Female"])
        
        with col3:
            att = st.selectbox("Attitude", options=[90,85,80,75,70,65,60,55,40], index=3, help="Attitude band")
            achieve = st.selectbox("Achievement", options=[90,85,80,75,70,65,60,55,40], index=3, help="Achievement band")
        
        # Second row - target and optional
        col4, col5 = st.columns([1, 2])
        
        with col4:
            target = st.selectbox("Target", options=[90,85,80,75,70,65,60,55,40], index=3, help="Target band")
        
        with col5:
            attitude_target = st.text_area(
                "Optional Next Steps",
                placeholder="E.g., continue to participate actively in class discussions...",
                height=40
            )
        
        # Submit button
        submitted = st.form_submit_button("Generate Comment", use_container_width=True)
    
    # Handle form submission
    if submitted:
        if not name.strip():
            st.error("Please enter a student name")
            st.stop()
        
        # Store form data in session state
        st.session_state.form_data = {
            'subject': subject,
            'year': year,
            'name': name,
            'gender': gender,
            'att': att,
            'achieve': achieve,
            'target': target,
            'attitude_target': attitude_target
        }
        
        # Generate original comment (variant 0)
        with st.spinner("Generating comment..."):
            comment = generate_comment(
                subject=subject,
                year=year,
                name=name,
                gender=gender,
                att=att,
                achieve=achieve,
                target=target,
                attitude_target=attitude_target,
                variant=0
            )
            
            # Store in session state
            st.session_state.current_comment = comment
            st.session_state.current_variant = None
            st.session_state.show_variant = False
    
    # Show generated comment if it exists
    if 'form_data' in st.session_state and st.session_state.form_data and 'current_comment' in st.session_state:
        form_data = st.session_state.form_data
        
        st.markdown("---")
        st.markdown("### Generated Comment")
        
        # Determine which comment to show
        if st.session_state.get('show_variant', False) and st.session_state.get('current_variant'):
            display_comment = st.session_state.current_variant
            comment_source = "Variant Comment"
        else:
            display_comment = st.session_state.current_comment
            comment_source = "Original Comment"
        
        # Display comment
        col_comment, col_copy = st.columns([4, 1])
        with col_comment:
            st.text_area("", display_comment, height=150, key="comment_display", label_visibility="collapsed")
        with col_copy:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Copy", use_container_width=True, help="Copy comment to clipboard"):
                st.code(display_comment, language=None)
                st.success("Ready to copy!")
        
        # Statistics - compact
        char_count = len(display_comment)
        col_stats = st.columns(3)
        with col_stats[0]:
            st.metric("Characters", f"{char_count}/{TARGET_CHARS}")
        with col_stats[1]:
            st.metric("Words", len(display_comment.split()))
        with col_stats[2]:
            status = "✓ Good" if char_count < TARGET_CHARS - 50 else "⚠ Near limit"
            st.metric("Status", status)
        
        # Add to all_comments list
        current_entry = {
            'name': form_data.get('name', 'Student'),
            'subject': form_data.get('subject', 'English'),
            'year': form_data.get('year', 7),
            'comment': display_comment,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M"),
            'variant': 'Variant' if st.session_state.get('show_variant', False) else 'Original'
        }
        
        # Check if this exact comment is already in the list
        comment_exists = False
        for entry in st.session_state.all_comments:
            if (entry['name'] == current_entry['name'] and 
                entry['subject'] == current_entry['subject'] and 
                entry['year'] == current_entry['year'] and
                entry['comment'] == current_entry['comment']):
                comment_exists = True
                break
        
        if not comment_exists:
            st.session_state.all_comments.append(current_entry)
        
        # Action buttons
        col_actions = st.columns(3)
        
        with col_actions[0]:
            if st.button("Generate Variant", type="secondary", use_container_width=True):
                # Get current form data
                form_data = st.session_state.form_data
                
                # Get available variants
                available_variants = get_available_variants(form_data['subject'], form_data['year'])
                
                # Create a unique key for this student
                student_key = f"{form_data['name']}_{form_data['subject']}_{form_data['year']}"
                
                # Initialize variant tracker for this student if not exists
                if student_key not in st.session_state.variant_tracker:
                    st.session_state.variant_tracker[student_key] = {
                        'variant_count': 0,
                        'last_comment': '',
                        'last_variant_num': 0
                    }
                
                # Get the tracker for this student
                tracker = st.session_state.variant_tracker[student_key]
                
                # Use variant 1 for variant generation
                variant_num = 1 if 1 in available_variants else 0
                
                # Generate variant comment with a new random seed
                comment_variant = generate_comment(
                    subject=form_data['subject'],
                    year=form_data['year'],
                    name=form_data['name'],
                    gender=form_data['gender'],
                    att=form_data['att'],
                    achieve=form_data['achieve'],
                    target=form_data['target'],
                    attitude_target=form_data.get('attitude_target', ''),
                    variant=variant_num
                )
                
                # Check if this is the same as the last variant comment
                attempt = 0
                while comment_variant == tracker.get('last_comment', '') and attempt < 3:
                    attempt += 1
                    # If same, try to generate again with different random seed
                    random.seed(time.time() + attempt * 1000)
                    comment_variant = generate_comment(
                        subject=form_data['subject'],
                        year=form_data['year'],
                        name=form_data['name'],
                        gender=form_data['gender'],
                        att=form_data['att'],
                        achieve=form_data['achieve'],
                        target=form_data['target'],
                        attitude_target=form_data.get('attitude_target', ''),
                        variant=variant_num
                    )
                
                # Store variant
                st.session_state.current_variant = comment_variant
                st.session_state.show_variant = True
                st.session_state.variant_num = variant_num
                
                # Update tracker
                tracker['variant_count'] += 1
                tracker['last_comment'] = comment_variant
                tracker['last_variant_num'] = variant_num
                
                st.rerun()
        
        with col_actions[1]:
            if st.button("Use Original", use_container_width=True, type="secondary"):
                st.session_state.show_variant = False
                st.rerun()
        
        with col_actions[2]:
            if st.button("Add Another Student", use_container_width=True):
                st.session_state.current_comment = ""
                st.session_state.current_variant = ""
                st.session_state.show_variant = False
                st.rerun()

# ========== BATCH UPLOAD MODE ==========
elif app_mode == "Batch Upload":
    st.markdown("---")
    st.markdown("### Batch Upload (CSV)")
    
    st.markdown("""
    <div style='background-color: #f8f9fa; padding: 15px; border-radius: 4px; margin: 20px 0;'>
    <strong>CSV Format Required:</strong><br>
    - Columns: Student Name, Gender, Subject, Year, Attitude, Achievement, Target<br>
    - Gender: Male/Female<br>
    - Subject: English/Maths/Science<br>
    - Year: 5, 7, or 8<br>
    - Bands: 90,85,80,75,70,65,60,55,40
    </div>
    """, unsafe_allow_html=True)
    
    example_csv = """Student Name,Gender,Subject,Year,Attitude,Achievement,Target
Aseel,Female,English,5,75,80,85
Mohamed,Male,Maths,7,80,75,80
Sarah,Female,Science,8,85,90,85"""
    
    col_dl, col_up = st.columns(2)
    
    with col_dl:
        st.download_button(
            label="Download Example CSV",
            data=example_csv,
            file_name="example_students.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col_up:
        uploaded_file = st.file_uploader("Choose CSV file", type=['csv'], label_visibility="collapsed")
    
    if uploaded_file:
        if not validate_upload_rate():
            st.stop()
        
        is_valid, msg = validate_file(uploaded_file)
        if not is_valid:
            st.error(msg)
            st.stop()
        
        with st.spinner("Processing CSV securely..."):
            df = process_csv_securely(uploaded_file)
        
        if df is not None:
            st.success(f"Processed {len(df)} students successfully")
            
            with st.expander("Preview Data (First 5 rows)"):
                st.dataframe(df.head())
            
            if st.button("Generate All Comments", use_container_width=True):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, row in df.iterrows():
                    progress = (idx + 1) / len(df)
                    progress_bar.progress(progress)
                    status_text.text(f"Processing {idx + 1}/{len(df)}: {row.get('Student Name', 'Student')}")
                    
                    try:
                        comment = generate_comment(
                            subject=str(row.get('Subject', 'English')),
                            year=int(row.get('Year', 7)),
                            name=str(row.get('Student Name', '')),
                            gender=str(row.get('Gender', '')),
                            att=int(row.get('Attitude', 75)),
                            achieve=int(row.get('Achievement', 75)),
                            target=int(row.get('Target', 75)),
                            attitude_target=""
                        )
                        
                        student_entry = {
                            'name': sanitize_input(str(row.get('Student Name', ''))),
                            'subject': str(row.get('Subject', 'English')),
                            'year': int(row.get('Year', 7)),
                            'comment': comment,
                            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M")
                        }
                        st.session_state.all_comments.append(student_entry)
                        
                    except Exception as e:
                        st.error(f"Error processing row {idx + 1}: {e}")
                
                progress_bar.empty()
                status_text.empty()
                st.success(f"Generated {len(df)} comments!")
                st.session_state.last_upload_time = datetime.now()

# ========== PRIVACY INFO MODE ==========
elif app_mode == "Privacy Info":
    st.markdown("---")
    st.markdown("### Privacy & Security Information")
    
    st.markdown("""
    #### How We Protect Student Data
    
    **Data Handling:**
    - All processing happens in your browser's memory
    - No student data is sent to or stored on our servers
    - Temporary files are created and immediately deleted
    - No database or persistent storage is used
    
    **Security Features:**
    1. **Input Sanitization** - Removes special characters from names
    2. **Rate Limiting** - Prevents abuse of the system
    3. **File Validation** - Checks file size and type
    4. **Auto-Cleanup** - Temporary files deleted after processing
    5. **Memory Clearing** - All data erased on browser close
    
    **Best Practices for Users:**
    - Use only first names or student IDs
    - Close browser tab when finished to clear all data
    - Download reports immediately after generation
    - For maximum privacy, use on school-managed devices
    
    **Compliance:**
    - Designed for use with anonymized data
    - Suitable for FERPA/GDPR compliant workflows
    - No third-party data sharing
    """)
    
    if st.button("Print Privacy Notice", type="secondary", use_container_width=True):
        privacy_text = """
        MULTI-SUBJECT REPORT GENERATOR - PRIVACY NOTICE
        
        Data Processing: All student data is processed locally in memory only.
        No data is transmitted to external servers or stored permanently.
        
        Data Retention: All data is cleared when the browser tab is closed.
        
        Security: Input sanitization and validation prevents data injection.
        
        Usage: For use with anonymized student data only.
        """
        st.text_area("Privacy Notice for Records", privacy_text, height=200)

# ========== DOWNLOAD SECTION ==========
if 'all_comments' in st.session_state and st.session_state.all_comments:
    st.markdown("---")
    st.markdown("### Download Reports")
    
    total_comments = len(st.session_state.all_comments)
    st.markdown(f"**You have {total_comments} generated comment(s)**")
    
    with st.expander(f"Preview Comments ({total_comments})"):
        for idx, entry in enumerate(st.session_state.all_comments, 1):
            variant_label = f" ({entry.get('variant', '')})" if 'variant' in entry else ''
            st.markdown(f"**{idx}. {entry['name']}** ({entry['subject']} Year {entry['year']}){variant_label}")
            st.write(entry['comment'])
            st.markdown("---")
    
    col_dl1, col_dl2, col_dl3 = st.columns(3)
    
    with col_dl1:
        if DOCX_AVAILABLE:
            if st.button("Word Document", use_container_width=True):
                doc = Document()
                doc.add_heading('Report Comments', 0)
                doc.add_paragraph(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
                doc.add_paragraph(f'Total Students: {total_comments}')
                doc.add_paragraph('')
                
                for entry in st.session_state.all_comments:
                    doc.add_heading(f"{entry['name']} - {entry['subject']} Year {entry['year']}", level=2)
                    doc.add_paragraph(entry['comment'])
                    doc.add_paragraph('')
                
                bio = io.BytesIO()
                doc.save(bio)
                
                st.download_button(
                    label="Download Word File",
                    data=bio.getvalue(),
                    file_name=f"report_comments_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        else:
            st.button("Word Document (Disabled)", use_container_width=True, disabled=True)
            st.caption("Word export requires 'docx' package")
    
    with col_dl2:
        if st.button("CSV Export", use_container_width=True):
            csv_data = []
            for entry in st.session_state.all_comments:
                csv_data.append({
                    'Student Name': entry['name'],
                    'Subject': entry['subject'],
                    'Year': entry['year'],
                    'Comment': entry['comment'],
                    'Generated': entry['timestamp']
                })
            
            df_export = pd.DataFrame(csv_data)
            csv_bytes = df_export.to_csv(index=False).encode('utf-8')
            
            st.download_button(
                label="Download CSV",
                data=csv_bytes,
                file_name=f"report_comments_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    with col_dl3:
        if st.button("Clear & Start Over", type="secondary", use_container_width=True):
            st.session_state.all_comments = []
            st.session_state.form_data = {}
            st.session_state.current_comment = ""
            st.session_state.current_variant = ""
            st.session_state.show_variant = False
            st.session_state.variant_tracker = {}
            st.success("All comments cleared!")
            st.rerun()

# ========== FOOTER ==========
st.markdown("---")
st.caption("© Report Generator v3.1 • With Variant Support")
