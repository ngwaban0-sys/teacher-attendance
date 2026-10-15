import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

# ==================== CONFIGURATION ====================
SPREADSHEET_ID = "1aSdwGcoPT4PZl_laM4xfg418l7uKeTCEhoSbHqGHhxg"

CLASSES_CONFIG = {
    "ALTVE": {
        "sheet_name": "ALTVE",
        "first_row": 4,
        "last_row": 31,
        "subjects": {
            "SURV2026": {"name": "Surveying, soils and Materials (7220)", "start_col": 2, "num_papers": 3},
            "BUILD2026": {"name": "Building Construction Technology and Practice (7201)", "start_col": 5, "num_papers": 3},
            "DRAW2026": {"name": "Building Construction Drawing (7206)", "start_col": 8, "num_papers": 3},
            "PROJ2026": {"name": "Building Construction Project Management (7211)", "start_col": 11, "num_papers": 3},
            "MECH2026": {"name": "Building Construction Applied Mechanics (7216)", "start_col": 14, "num_papers": 3},
            "MATH2026": {"name": "Mathematics (7160)", "start_col": 17, "num_papers": 3},
            "ENGSCI2026": {"name": "Engineering Science (7155)", "start_col": 20, "num_papers": 3},
            "COMP2026": {"name": "Industrial Computing (7650)", "start_col": 23, "num_papers": 3},
            "ENG2026": {"name": "English language (7105)", "start_col": 26, "num_papers": 3},
            "LAW2026": {"name": "Law (7110)", "start_col": 29, "num_papers": 3},
            "ENT2026": {"name": "Entrepreneurship (7055)", "start_col": 32, "num_papers": 3},
        }
    },
    "ILTVE": {
        "sheet_name": "ILTVE",
        "first_row": 4,
        "last_row": 61,
        "subjects": {
            "DRAW5225": {"name": "BUILDING CONSTRUCTION DRAWING (5225)", "start_col": 2, "num_papers": 3},
            "SURV5230": {"name": "SURVEYING AND SOIL MECHANICS (5230)", "start_col": 5, "num_papers": 3},
            "PROC5200": {"name": "CONSTRUCTION PROCESSES AND PRACTICE (5200)", "start_col": 8, "num_papers": 3},
            "CITIZEN0562": {"name": "CITIZENSHIP AND MORAL EDUCATION (0562)", "start_col": 11, "num_papers": 2},
            "ENGSCI5155": {"name": "ENGINEERING SCIENCE (5155)", "start_col": 13, "num_papers": 2},
            "ENG5105": {"name": "ENGLISH LANGUAGE (5105)", "start_col": 15, "num_papers": 2},
            "ENT5055": {"name": "ENTREPRENEURSHIP (5055)", "start_col": 17, "num_papers": 2},
            "FRENCH0545": {"name": "FRENCH LANGUAGE (0545)", "start_col": 19, "num_papers": 2},
            "COMP5650": {"name": "INDUSTRIAL COMPUTING (5650)", "start_col": 21, "num_papers": 3},
            "LAW5120": {"name": "LAW AND GOVERNMENT (5120)", "start_col": 24, "num_papers": 2},
            "MATH5440": {"name": "MATHEMATICS (5440)", "start_col": 26, "num_papers": 2},
            "QHSE5420": {"name": "QUALITY HYGIENE SAFETY AND ENVIRONMENT (5420)", "start_col": 28, "num_papers": 2},
        }
    },
    "CAPI": {
        "sheet_name": "CAPI",
        "first_row": 4,
        "last_row": 32,
        "subjects": {
            "ENG_CAPI": {"name": "ENGLISH LANGUAGE", "start_col": 2, "num_papers": 1},
            "FR_CAPI": {"name": "FRENCH LANGUAGE", "start_col": 3, "num_papers": 1},
            "PHYS_CAPI": {"name": "PHYSICS AND CHEMISTRY", "start_col": 4, "num_papers": 1},
            "MATH_CAPI": {"name": "MATHEMATICS", "start_col": 5, "num_papers": 1},
            "CS_CAPI": {"name": "COMPUTER SCIENCE", "start_col": 6, "num_papers": 1},
            "CITIZEN_CAPI": {"name": "CITIZENSHIP AND MORAL EDUCATION", "start_col": 7, "num_papers": 1},
            "QHSE_CAPI": {"name": "QUALITY HYGIENE, SAFETY AND ENVIRONMENT", "start_col": 8, "num_papers": 1},
            "ENT_CAPI": {"name": "ENTREPRENEURSHIP", "start_col": 9, "num_papers": 1},
            "MATERIAL_CAPI": {"name": "MATERIAL, EQUIPMENT AND REALISATION PROCESSES", "start_col": 10, "num_papers": 1},
            "DRAW_CAPI": {"name": "TECHNICAL DRAWING", "start_col": 11, "num_papers": 1},
            "SPORT_CAPI": {"name": "SPORTS AND PHYSICAL EDUCATION", "start_col": 12, "num_papers": 1},
        }
    }
}

# =======================================================

def col_letter(col_num):
    """Convert 1‑based column number to Excel letter (e.g., 1→A, 27→AA)."""
    result = ""
    while col_num > 0:
        col_num -= 1
        result = chr(ord('A') + col_num % 26) + result
        col_num //= 26
    return result

@st.cache_resource
def get_sheet(sheet_name):
    creds_dict = json.loads(st.secrets["GCP_SERVICE_ACCOUNT"])
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)

def get_candidates(sheet, first_row, last_row):
    values = sheet.get(f"A{first_row}:A{last_row}")
    return [row[0] for row in values if row and row[0].strip()]

def load_scores(sheet, start_col, num_papers, first_row, last_row, num_rows):
    if num_rows == 0:
        return []
    col_start = col_letter(start_col)
    range_str = f"{col_start}{first_row}:{col_letter(start_col + num_papers - 1)}{last_row}"
    data = sheet.get(range_str)
    result = []
    for i in range(num_rows):
        row_data = [0] * num_papers
        if i < len(data) and data[i]:
            for p in range(num_papers):
                val = data[i][p] if p < len(data[i]) and data[i][p] else 0
                row_data[p] = int(val)
        result.append(row_data)
    return result

def save_scores(sheet, start_col, num_papers, first_row, last_row, scores):
    col_start = col_letter(start_col)
    range_str = f"{col_start}{first_row}:{col_letter(start_col + num_papers - 1)}{last_row}"
    sheet.update(range_str, scores)

st.set_page_config(page_title="Mark Entry System", layout="centered")
st.set_page_config(page_title="Mark Entry System", layout="centered")

# === HIDE MANAGE APP BUTTON ===
hide_manage = """
<style>
button[title="Manage app"] { display: none !important; }
footer { visibility: hidden; }
</style>
"""
st.markdown(hide_manage, unsafe_allow_html=True)
# ==============================
# ... rest of your app
st.title("📝 Mark Entry System")

# Session state
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.classe = None
    st.session_state.subject_key = None
    st.session_state.subject_config = None

if not st.session_state.authenticated:
    with st.form("login_form"):
        classe = st.selectbox("Select class", list(CLASSES_CONFIG.keys()))
        code = st.text_input("Subject code", placeholder="e.g., SURV2026")
        submitted = st.form_submit_button("Access")
        if submitted:
            code = code.strip().upper()
            if code in CLASSES_CONFIG[classe]["subjects"]:
                st.session_state.authenticated = True
                st.session_state.classe = classe
                st.session_state.subject_key = code
                st.session_state.subject_config = CLASSES_CONFIG[classe]["subjects"][code]
                st.rerun()
            else:
                st.error("Invalid subject code for this class")
else:
    classe = st.session_state.classe
    config = CLASSES_CONFIG[classe]
    subject = st.session_state.subject_config
    sheet_name = config["sheet_name"]
    first_row = config["first_row"]
    last_row = config["last_row"]
    num_papers = subject["num_papers"]
    
    st.subheader(f"Class: {classe}  |  Subject: {subject['name']}")
    st.caption(f"Number of papers: {num_papers}  |  Candidates: rows {first_row} to {last_row}")
    
    try:
        sheet = get_sheet(sheet_name)
        candidates = get_candidates(sheet, first_row, last_row)
        if not candidates:
            st.error(f"No candidates found in sheet '{sheet_name}' (column A, rows {first_row} to {last_row}).")
            if st.button("Change class / subject"):
                st.session_state.authenticated = False
                st.rerun()
            st.stop()
        
        scores = load_scores(sheet, subject["start_col"], num_papers, first_row, last_row, len(candidates))
        col_names = [f"P{i+1}" for i in range(num_papers)]
        df = pd.DataFrame(scores, columns=col_names, index=candidates)
        df.index.name = "Candidate"
        
        # Configure columns: Paper 1 max 50, others max 100
        col_config = {}
        for i, col in enumerate(col_names):
            max_val = 50 if i == 0 and num_papers >= 1 else 100
            col_config[col] = st.column_config.NumberColumn(col, min_value=0, max_value=max_val, step=1)
        
        edited_df = st.data_editor(df, use_container_width=True, column_config=col_config)
        
        if st.button("💾 Save all marks"):
            new_scores = edited_df.values.tolist()
            save_scores(sheet, subject["start_col"], num_papers, first_row, last_row, new_scores)
            st.success("Marks saved successfully!")
            st.rerun()
        
        if st.button("🚪 Change class / subject"):
            st.session_state.authenticated = False
            st.rerun()
            
    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check that the Google Sheet exists and columns are correct.")
