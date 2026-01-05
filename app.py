import streamlit as st
import pandas as pd
import openpyxl
from firebase_admin import credentials, firestore
import firebase_admin
import os
import json

if not firebase_admin._apps:
    # 1. Pull the secret and force it into a standard Python Dictionary
    # Use .to_dict() or dict() to ensure it's not a string or AttrDict
    cred_info = dict(st.secrets["firebase"])
    
    # 2. Safety check: Ensure the private key handles newlines correctly
    if "\\n" in cred_info["private_key"]:
        cred_info["private_key"] = cred_info["private_key"].replace("\\n", "\n")
    
    # 3. Pass the actual dictionary object to the credentials
    cred = credentials.Certificate(cred_info)
    firebase_admin.initialize_app(cred)

db = firestore.client()
EXCEL_FILE = "P.I - Tool Kit.xlsx"
LOGO_FILE = "image_a242cc.jpg"

# --- 2. COLOR EXTRACTION ---
def get_excel_color_matrix(sheet_name):
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        ws = wb[sheet_name]
        full_matrix = []
        for row in ws.iter_rows():
            row_colors = []
            for cell in row:
                color = cell.fill.start_color.index
                if isinstance(color, str) and len(color) == 8:
                    hex_val = f"#{color[2:]}"
                elif isinstance(color, str) and len(color) == 6:
                    hex_val = f"#{color}"
                else:
                    hex_val = "#FFFFFF"
                row_colors.append(hex_val)
            full_matrix.append(row_colors)
        return full_matrix
    except:
        return None

# --- 3. UI & CSS INJECTION ---
st.set_page_config(page_title="Erudite Tool Kit", layout="wide")

def inject_excel_ui_styles(header_colors):
    css = "<style>"
    for i, color in enumerate(header_colors):
        if color != "#FFFFFF":
            css += f"""
            [data-testid="stDataFrame"] th:nth-child({i+1}) {{
                background-color: {color} !important;
                color: {'white' if color != '#FFFFFF' else 'black'} !important;
            }}
            """
    css += """
    [data-testid='stElementToolbar'] { display: none !important; }
    .stDataFrame td { vertical-align: middle !important; }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# --- 4. APP LOGIC ---
if 'seed' not in st.session_state:
    if os.path.exists(LOGO_FILE): st.image(LOGO_FILE, width=150)
    st.title("Erudite P.I. Tool Kit")
    seed_input = st.text_input("Enter 5-digit Toolkit Code:", max_chars=5).upper()
    if st.button("Enter Dashboard"):
        if len(seed_input) == 5:
            st.session_state.seed = seed_input
            st.rerun()
        else:
            st.error("Please enter exactly 5 characters.")
else:
    seed = st.session_state.seed
    student_ref = db.collection("students").document(seed)
    
    s_doc = student_ref.get()
    current_name = s_doc.to_dict().get('name', 'New Student') if s_doc.exists else "New Student"
    st.sidebar.title(f"Code: {seed}")
    new_name = st.sidebar.text_input("Student Name:", value=current_name)
    if new_name != current_name:
        student_ref.set({'name': new_name, 'seed': seed}, merge=True)

    tabs = st.tabs(["Master", "Work Experience", "Personal Questions", "Time Lines", "Story Line"])

    def render_tab(tab_obj, sheet_name, is_dynamic=False):
        with tab_obj:
            doc_ref = student_ref.collection("sheets").document(sheet_name)
            fb_doc = doc_ref.get()
            
            # FIXED: Unique space-based names for blank columns to satisfy Streamlit requirements
            hardcoded_headers = [
                " ", # Col 0
                "Some Important Work Experience Questions", # Col 1
                "  ", # Col 2 (Two spaces)
                "Have Complete Clarity", # Col 3
                "Need to work", # Col 4
                "No Idea", # Col 5
                "   ", # Col 6 (Three spaces)
                "Story 1", "Story 2", "Story 3", "Story 4", "Story 5"
            ]

            if fb_doc.exists:
                df = pd.DataFrame(fb_doc.to_dict()['data'])
            else:
                df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
                if sheet_name == "Work Experience":
                    # Map unique names to columns
                    df.columns = hardcoded_headers[:len(df.columns)]
                    # Remove the extra title row from Excel if it exists
                    if df.shape[0] > 0 and "Work Experience" in str(df.iloc[0, 1]):
                        df = df.iloc[1:].reset_index(drop=True)

            # CLEANING: Fix 1.000000 formatting and remove 'nan'
            def clean_val(v):
                if pd.isna(v) or str(v).lower() in ['nan', 'none']: return ""
                if isinstance(v, float) and v.is_integer(): return str(int(v))
                return str(v)
            
            df = df.applymap(clean_val)

            # COLORING
            color_matrix = get_excel_color_matrix(sheet_name)
            if color_matrix:
                if sheet_name == "Work Experience":
                    header_colors = color_matrix[1] 
                    body_colors = color_matrix[2:]
                else:
                    header_colors = color_matrix[0]
                    body_colors = color_matrix[1:]
                
                inject_excel_ui_styles(header_colors)

                def apply_matrix_styling(data):
                    style_df = pd.DataFrame('', index=data.index, columns=data.columns)
                    for r in range(len(data)):
                        for c in range(len(data.columns)):
                            try:
                                color = body_colors[r][c]
                                if color != "#FFFFFF":
                                    style_df.iloc[r, c] = f'background-color: {color}; color: {"white" if color != "#FFFFFF" else "inherit"}'
                            except: pass
                    return style_df

                styled_df = df.style.apply(apply_matrix_styling, axis=None)
            else:
                styled_df = df

            # CONFIG: Matching dropdowns to hardcoded names
            config = {}
            for col in df.columns:
                if any(x in col for x in ["Clarity", "work", "Idea"]):
                    config[col] = st.column_config.SelectboxColumn(label=col, options=["YES", "NO", "In Progress"])
                else:
                    config[col] = st.column_config.Column(label=col)

            st.subheader(sheet_name)
            edited_df = st.data_editor(
                styled_df,
                column_config=config,
                width='stretch',
                hide_index=True,
                num_rows="dynamic" if is_dynamic else "fixed",
                key=f"editor_{sheet_name}"
            )

            if st.button(f"Save {sheet_name}"):
                data_to_save = pd.DataFrame(edited_df).to_dict(orient='records')
                doc_ref.set({"data": data_to_save})
                st.success("Changes saved successfully!")

    render_tab(tabs[0], "Master")
    render_tab(tabs[1], "Work Experience")
    render_tab(tabs[2], "Personal Questions")
    render_tab(tabs[3], "time lines", is_dynamic=True)
    with tabs[4]:
        df_story = pd.read_excel(EXCEL_FILE, sheet_name="Building Story Line")
        st.dataframe(df_story.astype(str).replace('nan', ''), width='stretch', hide_index=True)

    if st.sidebar.button("Logout"):
        del st.session_state.seed
        st.rerun()
