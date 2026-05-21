import streamlit as st
import pandas as pd
import re
from datetime import datetime

# ========================
# PAGE CONFIGURATION
# ========================
st.set_page_config(
    page_title="JEE Student Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ========================
# LOAD DATA WITH CACHING
# ========================
@st.cache_data(ttl=3600)
def load_students():
    """Load students from CSV file"""
    df = pd.read_csv("students.csv")
    # Handle different column name possibilities
    if 'name' in df.columns:
        name_col = 'name'
    elif 'STUDENT NAME' in df.columns:
        name_col = 'STUDENT NAME'
    elif 'Name' in df.columns:
        name_col = 'Name'
    else:
        # Use first column as name
        name_col = df.columns[0]
    
    if 'mother_name' in df.columns:
        mother_col = 'mother_name'
    elif 'MOTHER NAME' in df.columns:
        mother_col = 'MOTHER NAME'
    elif 'Mother' in df.columns:
        mother_col = 'Mother'
    else:
        # Use second column as mother name
        mother_col = df.columns[1]
    
    # Rename for consistency
    df = df.rename(columns={name_col: 'name', mother_col: 'mother_name'})
    return df

@st.cache_data(ttl=3600)
def load_all_test_sheets():
    """Load and process all test sheets from Excel"""
    xl = pd.ExcelFile("StudentMarks.xlsx")
    
    # Get all test sheets
    test_sheets = []
    for sheet in xl.sheet_names:
        sheet_upper = sheet.upper()
        if any(x in sheet_upper for x in ["BATCH", "GRAND", "BTEST", "BRTEST"]):
            test_sheets.append(sheet)
    
    all_test_data = {}
    
    for sheet_name in test_sheets:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        
        # Find header row
        header_row = None
        for i, row in df.iterrows():
            if "TOTAL RANK" in str(row.iloc[0]):
                header_row = i
                break
        
        if header_row is None:
            continue
        
        headers = df.iloc[header_row].tolist()
        
        try:
            name_col = headers.index("STUDENT NAME")
            phy_col = headers.index("PHY")
            chem_col = headers.index("CHEM")
            maths_col = headers.index("MATHS")
            total_col = headers.index("TOTAL")
            rank_col = headers.index("TOTAL RANK")
        except ValueError:
            continue
        
        data_df = df.iloc[header_row + 1:].reset_index(drop=True)
        
        records = []
        for i in range(len(data_df)):
            student_name = data_df.iloc[i, name_col]
            if pd.notna(student_name):
                records.append({
                    "STUDENT NAME": str(student_name).strip(),
                    "PHY": data_df.iloc[i, phy_col] if pd.notna(data_df.iloc[i, phy_col]) else 0,
                    "CHEM": data_df.iloc[i, chem_col] if pd.notna(data_df.iloc[i, chem_col]) else 0,
                    "MATHS": data_df.iloc[i, maths_col] if pd.notna(data_df.iloc[i, maths_col]) else 0,
                    "TOTAL": data_df.iloc[i, total_col] if pd.notna(data_df.iloc[i, total_col]) else 0,
                    "RANK": data_df.iloc[i, rank_col] if pd.notna(data_df.iloc[i, rank_col]) else "-",
                })
        
        if records:
            df_clean = pd.DataFrame(records)
            is_brtest = "BRTEST" in sheet_name.upper()
            max_phy_chem = 50 if is_brtest else 100
            
            all_test_data[sheet_name] = {
                "data": df_clean,
                "max_phy_chem": max_phy_chem,
                "is_brtest": is_brtest
            }
    
    return all_test_data

def get_student_data(student_name, all_test_data):
    """Get all test data for a specific student"""
    student_records = []
    
    for sheet_name, sheet_info in all_test_data.items():
        df = sheet_info["data"]
        student_row = df[df["STUDENT NAME"].str.lower() == student_name.lower()]
        
        if not student_row.empty:
            row = student_row.iloc[0]
            student_records.append({
                "test_name": sheet_name[:45],
                "phy": row["PHY"],
                "chem": row["CHEM"],
                "maths": row["MATHS"],
                "total": row["TOTAL"],
                "rank": row["RANK"],
                "max_phy_chem": sheet_info["max_phy_chem"]
            })
    
    return student_records

# ========================
# CUSTOM CSS
# ========================
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .main-header {
        background: white;
        padding: 2rem;
        border-radius: 20px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .main-header h1 {
        color: #1e3c72;
        margin-bottom: 0.5rem;
    }
    .main-header p {
        color: #666;
    }
    .stat-card {
        background: white;
        padding: 1.2rem;
        border-radius: 15px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        transition: transform 0.2s;
    }
    .stat-card:hover {
        transform: translateY(-3px);
    }
    .stat-value {
        font-size: 2rem;
        font-weight: bold;
        color: #1e3c72;
    }
    .stat-label {
        color: #666;
        font-size: 0.8rem;
        margin-top: 0.5rem;
    }
    </style>
""", unsafe_allow_html=True)

# ========================
# SESSION STATE
# ========================
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.student_name = ""

# ========================
# LOGIN PAGE
# ========================
if not st.session_state.logged_in:
    st.markdown("""
        <div class="main-header">
            <h1>🔐 Student Portal</h1>
            <p>Login to view your performance dashboard</p>
        </div>
    """, unsafe_allow_html=True)
    
    with st.container():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            try:
                students_df = load_students()
                
                selected_name = st.selectbox(
                    "📝 Select Your Name",
                    options=sorted(students_df['name'].tolist()),
                    index=None,
                    placeholder="Type or select your name..."
                )
                
                mother_name = st.text_input(
                    "👩 Mother's Name (Password)",
                    type="password",
                    placeholder="Enter mother's name in lowercase"
                )
                
                if st.button("🔐 Login", type="primary", use_container_width=True):
                    if selected_name and mother_name:
                        student_row = students_df[students_df['name'] == selected_name].iloc[0]
                        if student_row['mother_name'].lower() == mother_name.lower():
                            st.session_state.logged_in = True
                            st.session_state.student_name = selected_name
                            st.rerun()
                        else:
                            st.error("❌ Invalid mother's name")
                    else:
                        st.warning("Please select your name and enter password")
            except Exception as e:
                st.error(f"Error loading student data: {e}")
                st.info("Please make sure students.csv has columns: 'name' and 'mother_name'")
    
    st.markdown("""
        <div style="text-align: center; margin-top: 2rem; color: rgba(255,255,255,0.7); font-size: 12px;">
            ⓘ Password is your mother's name in <strong>small letters</strong>
        </div>
    """, unsafe_allow_html=True)

# ========================
# DASHBOARD
# ========================
else:
    with st.spinner("Loading your dashboard..."):
        all_tests = load_all_test_sheets()
        student_records = get_student_data(st.session_state.student_name, all_tests)
    
    if not student_records:
        st.error(f"❌ No test records found for {st.session_state.student_name}")
        if st.button("← Back to Login"):
            st.session_state.logged_in = False
            st.rerun()
    else:
        st.markdown(f"""
            <div class="main-header">
                <h1>📊 Student Performance Dashboard</h1>
                <p><strong>{st.session_state.student_name}</strong></p>
            </div>
        """, unsafe_allow_html=True)
        
        df = pd.DataFrame(student_records)
        
        df['phy_percent'] = (df['phy'] / df['max_phy_chem']) * 100
        df['chem_percent'] = (df['chem'] / df['max_phy_chem']) * 100
        df['maths_percent'] = (df['maths'] / 100) * 100
        df['total_percent'] = (df['total'] / (df['max_phy_chem'] * 2 + 100)) * 100
        
        # Statistics
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        
        with col1:
            st.metric("Tests Taken", len(df))
        with col2:
            st.metric("Avg Physics", f"{round(df['phy_percent'].mean())}%")
        with col3:
            st.metric("Avg Chemistry", f"{round(df['chem_percent'].mean())}%")
        with col4:
            st.metric("Avg Maths", f"{round(df['maths_percent'].mean())}%")
        with col5:
            valid_ranks = df[df['rank'].apply(lambda x: isinstance(x, (int, float)) and x > 0)]['rank']
            best_rank = int(valid_ranks.min()) if not valid_ranks.empty else "—"
            st.metric("Best Rank", best_rank)
        with col6:
            st.metric("Avg Total", f"{round(df['total_percent'].mean())}%")
        
        st.markdown("---")
        
        # Charts
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📈 Subject-wise Progress")
            chart_data = pd.DataFrame({
                "Test": df['test_name'],
                "Physics": df['phy_percent'],
                "Chemistry": df['chem_percent'],
                "Maths": df['maths_percent']
            })
            st.line_chart(chart_data.set_index("Test"), height=350)
        
        with col2:
            st.subheader("🏆 Overall Performance")
            st.bar_chart(pd.DataFrame({
                "Test": df['test_name'],
                "Total %": df['total_percent']
            }).set_index("Test"), height=350)
        
        # Test History
        st.markdown("---")
        st.subheader("📋 Detailed Test History")
        
        display_df = df[['test_name', 'phy', 'chem', 'maths', 'total', 'rank']].copy()
        display_df.columns = ['Test Name', 'Physics', 'Chemistry', 'Maths', 'Total', 'Rank']
        display_df['Physics'] = df.apply(lambda x: f"{int(x['phy'])}/{x['max_phy_chem']}", axis=1)
        display_df['Chemistry'] = df.apply(lambda x: f"{int(x['chem'])}/{x['max_phy_chem']}", axis=1)
        display_df['Maths'] = df['maths'].astype(int).astype(str) + "/100"
        
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        # Logout
        st.markdown("---")
        if st.button("🚪 Logout", type="secondary", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.student_name = ""
            st.rerun()
        
        st.caption("⚠️ **Note:** BRTEST format has Physics & Chemistry out of 50 marks. Other tests have all subjects out of 100.")
