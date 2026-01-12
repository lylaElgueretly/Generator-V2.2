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
# FORCE LIGHT MODE
# ---------------------------

st.set_page_config(
    page_title="CommentCraft",
    layout="wide",
    initial_sidebar_state="auto"
)

# Custom CSS to force light mode and remove scrollbars
st.markdown("""
    <style>
    body, .css-18e3th9 {
        background-color: #ffffff;
        color: #000000;
    }
    .css-1d391kg, .css-1v3fvcr, .css-1avcm0n {
        overflow: visible !important;
    }
    .css-1v0mbdj {
        display: none; /* hides all icons in the app */
    }
    .stTextInput>div>div>input {
        height: 35px;
        font-size: 16px;
    }
    .stTextArea>div>div>textarea {
        height: 80px;
        font-size: 16px;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------
# APP HEADER
# ---------------------------
st.title("CommentCraft")
st.write("Effortlessly generate polished student comments.")

# ---------------------------
# SIDEBAR SELECTIONS
# ---------------------------

year = st.selectbox("Select Year Group:", ["Year 5", "Year 7", "Year 8"])
subject = st.selectbox("Select Subject:", ["English", "Maths", "Science"])
student_name = st.text_input("Student Name:")

# ---------------------------
# TEACHER INPUT
# ---------------------------
st.subheader("Teacher Comment Input")
teacher_comment = st.text_area(
    "Enter or paste your comment for the student (optional):",
    placeholder="You may write anything..."
)

# ---------------------------
# COMMENT GENERATION
# ---------------------------

def generate_comment(year, subject, name, teacher_text):
    # Get statement module
    if year == "Year 5":
        statements_module = {
            "English": statements_year5_English,
            "Maths": statements_year5_Maths,
            "Science": statements_year5_Science
        }[subject]
    elif year == "Year 7":
        statements_module = {
            "English": statements_year7_English,
            "Maths": statements_year7_Maths,
            "Science": statements_year7_science
        }[subject]
    elif year == "Year 8":
        statements_module = {
            "English": statements_year8_English,
            "Maths": statements_year8_Maths,
            "Science": statements_year8_science
        }[subject]
    
    # Pick default comment if teacher_text empty
    if not teacher_text.strip():
        comment = statements_module.default_comment()
    else:
        # Ensure full stop before the optional ending text
        teacher_text = teacher_text.strip()
        if not teacher_text.endswith("."):
            teacher_text += "."
        comment = f"{teacher_text} Keep it up."
    return comment

# ---------------------------
# GENERATE BUTTON
# ---------------------------
if st.button("Generate Comment"):
    if not student_name.strip():
        st.warning("Please enter the student's name.")
    else:
        final_comment = generate_comment(year, subject, student_name, teacher_comment)
        st.subheader("Generated Comment")
        st.success(f"{student_name}: {final_comment}")

# ---------------------------
# OPTIONAL: DOWNLOAD COMMENT
# ---------------------------
st.subheader("Download Comment")
st.download_button(
    label="Download as TXT",
    data=f"{student_name}: {generate_comment(year, subject, student_name, teacher_comment)}",
    file_name=f"{student_name}_comment.txt",
    mime="text/plain"
)
