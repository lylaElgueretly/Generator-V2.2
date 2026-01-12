# report_comment_app.py

import streamlit as st

# ---------------------------
# IMPORT STATEMENT FILES
# ---------------------------

# Year 5
import statements_year5_English_variant1 as statements_year5_English
import statements_year5_Maths_variant1 as statements_year5_Maths
import statements_year5_Science_variant1 as statements_year5_Science

# Year 7
import statements_year7_English_variant1 as statements_year7_English
import statements_year7_Maths_variant1 as statements_year7_Maths
import statements_year7_science_variant1 as statements_year7_science

# Year 8
import statements_year8_English_variant1 as statements_year8_English
import statements_year8_Maths_variant1 as statements_year8_Maths
import statements_year8_science_variant1 as statements_year8_science

# ---------------------------
# FORCE LIGHT MODE & STYLING
# ---------------------------

st.set_page_config(
    page_title="CommentCraft",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    body, .css-18e3th9 {background-color: #ffffff; color: #000000;}
    .stTextInput>div>div>input, .stTextArea>div>div>textarea {font-size: 16px;}
    .stButton>button {height: 40px; font-size: 16px;}
    .css-1v0mbdj, .css-1v3fvcr {display:none;} /* remove all icons */
    .css-1d391kg, .css-1avcm0n {overflow: visible !important;} /* single-window fit */
    .stMarkdown {margin-bottom: 5px;}
    </style>
""", unsafe_allow_html=True)

# ---------------------------
# APP HEADER
# ---------------------------

st.title("CommentCraft")
st.markdown("**International Kingdom College**")

st.markdown("""
### 1. Select
Choose student details

### 2. Generate
Create the comment

### 3. Download
Export your reports
""")

# ---------------------------
# STEP 1: Student Selection
# ---------------------------

col1, col2, col3 = st.columns([1,1,1])

with col1:
    year = st.selectbox("Select Year Group:", ["Year 5", "Year 7", "Year 8"])
with col2:
    subject = st.selectbox("Select Subject:", ["English", "Maths", "Science"])
with col3:
    student_name = st.text_input("Student Name:")

# ---------------------------
# STEP 2: Teacher Input (optional)
# ---------------------------

teacher_comment = st.text_area(
    "Teacher Comment (optional):",
    placeholder="Enter any comment..."
)

# ---------------------------
# COMMENT GENERATION FUNCTION
# ---------------------------

def generate_comment(year, subject, name, teacher_text):
    # Pick statement module based on year and subject
    if year == "Year 5":
        statements_module = {"English": statements_year5_English,
                             "Maths": statements_year5_Maths,
                             "Science": statements_year5_Science}[subject]
    elif year == "Year 7":
        statements_module = {"English": statements_year7_English,
                             "Maths": statements_year7_Maths,
                             "Science": statements_year7_science}[subject]
    elif year == "Year 8":
        statements_module = {"English": statements_year8_English,
                             "Maths": statements_year8_Maths,
                             "Science": statements_year8_science}[subject]

    # Generate comment
    if not teacher_text.strip():
        comment = statements_module.default_comment()
    else:
        teacher_text = teacher_text.strip()
        if not teacher_text.endswith("."):
            teacher_text += "."
        # Append optional encouragement dynamically
        comment = f"{teacher_text} Keep it up."
    return comment

# ---------------------------
# STEP 2 BUTTON: GENERATE COMMENT
# ---------------------------

if st.button("Generate Comment"):
    if not student_name.strip():
        st.warning("Please enter the student's name.")
    else:
        final_comment = generate_comment(year, subject, student_name, teacher_comment)
        st.subheader("Generated Comment")
        st.success(f"{student_name}: {final_comment}")

# ---------------------------
# STEP 3: DOWNLOAD BUTTON
# ---------------------------

st.download_button(
    label="Download Comment",
    data=f"{student_name}: {generate_comment(year, subject, student_name, teacher_comment)}",
    file_name=f"{student_name}_comment.txt",
    mime="text/plain"
)
