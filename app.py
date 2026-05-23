import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import re
from collections import Counter
import os
import warnings
warnings.filterwarnings('ignore')

# ========================
# PAGE CONFIGURATION
# ========================
st.set_page_config(
    page_title="Student Performance Dashboard",
    page_icon="📊",
    layout="wide"
)

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
    .student-header {
        background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460);
        padding: 20px;
        border-radius: 15px;
        margin-bottom: 20px;
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
# FILE PATHS
# ========================
EXCEL_FILE_PATH = "StudentMarks.xlsx"
CSV_FILE_PATH = "students.csv"

# ========================
# HELPER FUNCTIONS
# ========================
def detect_test_type(sheet_name):
    if sheet_name.upper().startswith("BRTEST"):
        return "BRTEST"
    return "BTEST"

def normalize_name(name):
    if pd.isna(name):
        return None
    name = str(name).upper().strip()
    name = re.sub(r'\s+', ' ', name)
    name = name.replace('KILKARNI', 'KULKARNI')
    if len(name) < 3:
        return None
    return name

def get_weakest_subject(phy_rank, chem_rank, math_rank):
    if pd.isna(phy_rank) or pd.isna(chem_rank) or pd.isna(math_rank):
        return "Absent"
    p, c, m = float(phy_rank), float(chem_rank), float(math_rank)
    if max(p, c, m) - min(p, c, m) <= 30:
        return "Balanced"
    ranks = {'Physics': p, 'Chemistry': c, 'Maths': m}
    weakest = max(ranks, key=ranks.get)
    other_avg = (sum(ranks.values()) - ranks[weakest]) / 2
    if ranks[weakest] > 1.5 * other_avg:
        return weakest
    return "Balanced"

@st.cache_data(ttl=3600)
def load_students():
    """Load students from CSV file"""
    try:
        df = pd.read_csv(CSV_FILE_PATH)
        # Handle different column name possibilities
        if 'name' in df.columns:
            name_col = 'name'
        elif 'STUDENT NAME' in df.columns:
            name_col = 'STUDENT NAME'
        elif 'Name' in df.columns:
            name_col = 'Name'
        else:
            name_col = df.columns[0]
        
        if 'mother_name' in df.columns:
            mother_col = 'mother_name'
        elif 'MOTHER NAME' in df.columns:
            mother_col = 'MOTHER NAME'
        elif 'Mother' in df.columns:
            mother_col = 'Mother'
        else:
            mother_col = df.columns[1]
        
        df = df.rename(columns={name_col: 'name', mother_col: 'mother_name'})
        return df
    except Exception as e:
        st.error(f"Error loading students.csv: {e}")
        return pd.DataFrame(columns=['name', 'mother_name'])

@st.cache_data(ttl=3600)
def load_excel_data():
    """Load all test data from Excel file"""
    all_student_data = {}
    test_metadata = {}
    
    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            st.error(f"Excel file not found: {EXCEL_FILE_PATH}")
            return None, None
        
        xl = pd.ExcelFile(EXCEL_FILE_PATH)
        
        # Get test sheets
        test_sheets = []
        for s in xl.sheet_names:
            s_upper = s.upper()
            if "SUMMARY" not in s_upper and "ANALYSIS" not in s_upper and "SHEET20" not in s_upper and "SHEET1" not in s_upper:
                test_sheets.append(s)
        
        for idx, sheet in enumerate(test_sheets, start=1):
            test_type = detect_test_type(sheet)
            
            if test_type == "BRTEST":
                max_phy, max_chem, max_math = 50, 50, 100
                max_total = 200
            else:
                max_phy, max_chem, max_math = 100, 100, 100
                max_total = 300

            test_metadata[sheet] = {
                "type": test_type, "index": idx,
                "max_phy": max_phy, "max_chem": max_chem,
                "max_math": max_math, "max_total": max_total
            }

            try:
                # Read with skiprows=7 (adjust if needed)
                df = pd.read_excel(xl, sheet_name=sheet, skiprows=7)
                
                if len(df.columns) >= 10:
                    overall_rank_col = df.columns[0]
                    name_col = df.columns[1]
                    roll_col = df.columns[2]
                    phy_col = df.columns[3]
                    phy_rank_col = df.columns[4]
                    chem_col = df.columns[5]
                    chem_rank_col = df.columns[6]
                    math_col = df.columns[7]
                    math_rank_col = df.columns[8]
                    total_col = df.columns[9]
                    
                    for _, row in df.iterrows():
                        if pd.isna(row[name_col]) and pd.isna(row[roll_col]):
                            continue
                        
                        student_name = normalize_name(row[name_col])
                        if not student_name:
                            continue
                        
                        roll_no = pd.to_numeric(row[roll_col], errors='coerce')
                        total = pd.to_numeric(row[total_col], errors='coerce')
                        overall_rank = pd.to_numeric(row[overall_rank_col], errors='coerce')
                        
                        if pd.isna(total) or total < 0:
                            continue
                        
                        if student_name not in all_student_data:
                            all_student_data[student_name] = {
                                "name": student_name, "roll_numbers": set(), "tests": {}
                            }
                        
                        if pd.notna(roll_no):
                            all_student_data[student_name]["roll_numbers"].add(int(roll_no))
                        
                        phy = pd.to_numeric(row[phy_col], errors='coerce')
                        chem = pd.to_numeric(row[chem_col], errors='coerce')
                        math_score = pd.to_numeric(row[math_col], errors='coerce')
                        phy_rank = pd.to_numeric(row[phy_rank_col], errors='coerce')
                        chem_rank = pd.to_numeric(row[chem_rank_col], errors='coerce')
                        math_rank = pd.to_numeric(row[math_rank_col], errors='coerce')
                        
                        all_student_data[student_name]["tests"][sheet] = {
                            "phy": phy if pd.notna(phy) else 0,
                            "chem": chem if pd.notna(chem) else 0,
                            "math": math_score if pd.notna(math_score) else 0,
                            "phy_rank": phy_rank if pd.notna(phy_rank) else None,
                            "chem_rank": chem_rank if pd.notna(chem_rank) else None,
                            "math_rank": math_rank if pd.notna(math_rank) else None,
                            "total": total,
                            "overall_rank": overall_rank if pd.notna(overall_rank) else None,
                            "type": test_type
                        }
                        
            except Exception as e:
                continue
        
        return all_student_data, test_metadata
        
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None, None

# ========================
# SESSION STATE
# ========================
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.student_name = ""
    st.session_state.student_data = None

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
            students_df = load_students()
            
            if students_df.empty:
                st.error("No student data found. Please check students.csv file.")
                st.stop()
            
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
    
    st.markdown("""
        <div style="text-align: center; margin-top: 2rem; color: rgba(255,255,255,0.7); font-size: 12px;">
            ⓘ Password is your mother's name in <strong>small letters</strong>
        </div>
    """, unsafe_allow_html=True)

# ========================
# DASHBOARD
# ========================
else:
    with st.spinner("🔄 Loading your data..."):
        all_student_data, test_metadata = load_excel_data()
    
    if all_student_data is None or len(all_student_data) == 0:
        st.error("❌ No student data found. Please check the Excel file.")
        if st.button("← Back to Login"):
            st.session_state.logged_in = False
            st.rerun()
        st.stop()
    
    student = all_student_data.get(st.session_state.student_name.upper())
    
    if not student:
        st.error(f"❌ No test records found for {st.session_state.student_name}")
        st.info("Possible reasons: The name might be spelled differently in the Excel file.")
        if st.button("← Back to Login"):
            st.session_state.logged_in = False
            st.rerun()
        st.stop()
    
    # Process student data
    btest_results = []
    brtest_results = []
    
    all_overall_ranks = []
    all_phy_ranks = []
    all_chem_ranks = []
    all_math_ranks = []
    
    for sheet, meta in test_metadata.items():
        if sheet not in student["tests"]:
            continue
        
        marks = student["tests"][sheet]
        pct = round((marks["total"] / meta["max_total"]) * 100, 1)
        
        overall_rank = marks.get("overall_rank")
        
        if overall_rank is None or pd.isna(overall_rank):
            all_scores = []
            for s_data in all_student_data.values():
                if sheet in s_data["tests"]:
                    all_scores.append(s_data["tests"][sheet]["total"])
            overall_rank = sum(score > marks["total"] for score in all_scores) + 1
        else:
            overall_rank = int(overall_rank)
        
        all_overall_ranks.append(overall_rank)
        if marks.get("phy_rank") is not None and not pd.isna(marks.get("phy_rank")):
            all_phy_ranks.append(int(marks.get("phy_rank")))
        if marks.get("chem_rank") is not None and not pd.isna(marks.get("chem_rank")):
            all_chem_ranks.append(int(marks.get("chem_rank")))
        if marks.get("math_rank") is not None and not pd.isna(marks.get("math_rank")):
            all_math_ranks.append(int(marks.get("math_rank")))
        
        weakest = get_weakest_subject(marks.get("phy_rank"), marks.get("chem_rank"), marks.get("math_rank"))
        
        phy_pct = round((marks["phy"] / meta["max_phy"]) * 100, 1) if marks["phy"] > 0 else 0
        chem_pct = round((marks["chem"] / meta["max_chem"]) * 100, 1) if marks["chem"] > 0 else 0
        math_pct = round((marks["math"] / meta["max_math"]) * 100, 1) if marks["math"] > 0 else 0
        
        result = {
            "S.No.": meta["index"],
            "Test Name": sheet,
            "Type": meta["type"],
            "Physics": marks["phy"],
            "Physics %": phy_pct,
            "Phy Rank": marks.get("phy_rank") if marks.get("phy_rank") is not None else '-',
            "Chemistry": marks["chem"],
            "Chemistry %": chem_pct,
            "Chem Rank": marks.get("chem_rank") if marks.get("chem_rank") is not None else '-',
            "Maths": marks["math"],
            "Maths %": math_pct,
            "Math Rank": marks.get("math_rank") if marks.get("math_rank") is not None else '-',
            "Total": f"{marks['total']:.0f}/{meta['max_total']}",
            "%": f"{pct}%",
            "Overall Rank": overall_rank,
            "Weakest Subject": weakest
        }
        
        if meta["type"] == "BTEST":
            btest_results.append(result)
        else:
            brtest_results.append(result)
    
    btest_results.sort(key=lambda x: x["S.No."])
    brtest_results.sort(key=lambda x: x["S.No."])
    
    all_tests = btest_results + brtest_results
    avg_pct = round(np.mean([float(t["%"].replace("%", "")) for t in all_tests]), 1) if all_tests else 0
    
    best_overall_rank = min(all_overall_ranks) if all_overall_ranks else 'N/A'
    worst_overall_rank = max(all_overall_ranks) if all_overall_ranks else 'N/A'
    best_phy_rank = min(all_phy_ranks) if all_phy_ranks else 'N/A'
    worst_phy_rank = max(all_phy_ranks) if all_phy_ranks else 'N/A'
    best_chem_rank = min(all_chem_ranks) if all_chem_ranks else 'N/A'
    worst_chem_rank = max(all_chem_ranks) if all_chem_ranks else 'N/A'
    best_math_rank = min(all_math_ranks) if all_math_ranks else 'N/A'
    worst_math_rank = max(all_math_ranks) if all_math_ranks else 'N/A'
    
    roll_numbers_display = ", ".join([str(r) for r in sorted(student["roll_numbers"])])
    
    weak_subjects = [t["Weakest Subject"] for t in all_tests if t["Weakest Subject"] not in ["Balanced", "Absent"]]
    weak_count = Counter(weak_subjects)
    
    # ============================================================
    # STUDENT INFO HEADER
    # ============================================================
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#1a1a2e,#16213e,#0f3460);
                padding:20px;
                border-radius:15px;
                margin-bottom:20px;">
        <h2 style="color:white; margin:0;">🎓 {st.session_state.student_name}</h2>
        <p style="color:white; margin:5px 0 0 0;">Roll Number(s): {roll_numbers_display}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # ============================================================
    # TEST SUMMARY METRICS
    # ============================================================
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📚 Tests Attempted", f"{len(all_tests)}/{len(test_metadata)}")
    with col2:
        st.metric("📊 Average Score", f"{avg_pct}%")
    with col3:
        st.metric("🏆 Best Overall Rank", best_overall_rank)
    with col4:
        st.metric("⚠️ Worst Overall Rank", worst_overall_rank)
    
    st.markdown("---")
    
    # ============================================================
    # SUBJECT RANK SUMMARY
    # ============================================================
    st.subheader("📊 Subject-wise Rank Summary")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.info("**🔬 Physics**")
        st.metric("Best Rank", best_phy_rank, delta="⭐")
        st.metric("Worst Rank", worst_phy_rank, delta="⚠️")
    
    with col2:
        st.info("**⚗️ Chemistry**")
        st.metric("Best Rank", best_chem_rank, delta="⭐")
        st.metric("Worst Rank", worst_chem_rank, delta="⚠️")
    
    with col3:
        st.info("**📐 Mathematics**")
        st.metric("Best Rank", best_math_rank, delta="⭐")
        st.metric("Worst Rank", worst_math_rank, delta="⚠️")
    
    st.markdown("---")
    
    # ============================================================
    # WEAK SUBJECT SUMMARY
    # ============================================================
    if weak_count:
        st.subheader("⚠️ Weak Subject Analysis")
        weak_cols = st.columns(len(weak_count))
        for i, (subject, count) in enumerate(weak_count.items()):
            with weak_cols[i]:
                st.warning(f"⚠️ {subject}")
                st.metric("Weak in", f"{count} test(s)")
        st.markdown("---")
    
    # ============================================================
    # BTEST RESULTS TABLE
    # ============================================================
    if btest_results:
        with st.expander("📘 BTEST/GRAND TESTS (JEE Format - 300 marks)", expanded=True):
            display_cols = ['Test Name', 'Physics', 'Phy Rank', 'Chemistry', 'Chem Rank', 
                           'Maths', 'Math Rank', 'Total', '%', 'Overall Rank', 'Weakest Subject']
            btest_df = pd.DataFrame(btest_results)
            st.dataframe(btest_df[display_cols], use_container_width=True)
    
    # ============================================================
    # BRTEST RESULTS TABLE
    # ============================================================
    if brtest_results:
        with st.expander("📘 BRTEST TESTS (CET Format - 200 marks)", expanded=True):
            display_cols = ['Test Name', 'Physics', 'Phy Rank', 'Chemistry', 'Chem Rank', 
                           'Maths', 'Math Rank', 'Total', '%', 'Overall Rank', 'Weakest Subject']
            brtest_df = pd.DataFrame(brtest_results)
            st.dataframe(brtest_df[display_cols], use_container_width=True)
    
    st.markdown("---")
    
    # ============================================================
    # SUBJECT MARKS TRENDS - BTEST
    # ============================================================
    if btest_results:
        st.subheader("📊 Subject Marks Trends - BTEST Tests")
        
        btest_names = [t['Test Name'][:20] for t in btest_results]
        
        fig_marks = go.Figure()
        fig_marks.add_trace(go.Scatter(x=btest_names, y=[t['Physics'] for t in btest_results], 
                                        mode='lines+markers', name='Physics', 
                                        line=dict(color='#3498DB', width=2), marker=dict(size=8)))
        fig_marks.add_trace(go.Scatter(x=btest_names, y=[t['Chemistry'] for t in btest_results], 
                                        mode='lines+markers', name='Chemistry', 
                                        line=dict(color='#9B59B6', width=2), marker=dict(size=8)))
        fig_marks.add_trace(go.Scatter(x=btest_names, y=[t['Maths'] for t in btest_results], 
                                        mode='lines+markers', name='Mathematics', 
                                        line=dict(color='#F1C40F', width=2), marker=dict(size=8)))
        fig_marks.add_hline(y=75, line_dash="dash", line_color="green", annotation_text="75% Target")
        fig_marks.update_layout(title="Subject Marks Comparison (BTEST)", height=400)
        st.plotly_chart(fig_marks, use_container_width=True)
    
    # ============================================================
    # SUBJECT RANK TRENDS - BTEST
    # ============================================================
    if btest_results:
        st.subheader("🏆 Subject Rank Trends - BTEST Tests (Lower is Better)")
        
        fig_ranks_btest = go.Figure()
        
        btest_names_rank = [t['Test Name'][:20] for t in btest_results]
        phy_ranks = [t['Phy Rank'] if t['Phy Rank'] != '-' else None for t in btest_results]
        chem_ranks = [t['Chem Rank'] if t['Chem Rank'] != '-' else None for t in btest_results]
        math_ranks = [t['Math Rank'] if t['Math Rank'] != '-' else None for t in btest_results]
        
        fig_ranks_btest.add_trace(go.Scatter(x=btest_names_rank, y=phy_ranks, 
                                              mode='lines+markers', name='Physics Rank',
                                              line=dict(color='#3498DB', width=2), marker=dict(size=8)))
        fig_ranks_btest.add_trace(go.Scatter(x=btest_names_rank, y=chem_ranks, 
                                              mode='lines+markers', name='Chemistry Rank',
                                              line=dict(color='#9B59B6', width=2), marker=dict(size=8)))
        fig_ranks_btest.add_trace(go.Scatter(x=btest_names_rank, y=math_ranks, 
                                              mode='lines+markers', name='Mathematics Rank',
                                              line=dict(color='#F1C40F', width=2), marker=dict(size=8)))
        
        fig_ranks_btest.update_layout(title="Subject Rank Comparison (BTEST)", 
                                      yaxis=dict(autorange="reversed"), height=400)
        st.plotly_chart(fig_ranks_btest, use_container_width=True)
    
    # ============================================================
    # OVERALL RANK TREND - BTEST
    # ============================================================
    if btest_results:
        st.subheader("🏆 Overall Rank Trend - BTEST Tests (Lower is Better)")
        
        fig_rank_btest = go.Figure()
        btest_overall_ranks = [t['Overall Rank'] for t in btest_results]
        
        fig_rank_btest.add_trace(go.Scatter(x=btest_names_rank, y=btest_overall_ranks,
                                             mode='lines+markers', name='Overall Rank',
                                             line=dict(color='#FF6B6B', width=3), marker=dict(size=10)))
        best_btest_rank = min(btest_overall_ranks) if btest_overall_ranks else None
        if best_btest_rank:
            fig_rank_btest.add_hline(y=best_btest_rank, line_dash="dash", line_color="green", 
                                    annotation_text=f"Best: {best_btest_rank}")
        
        fig_rank_btest.update_layout(title="Overall Rank Performance (BTEST)", 
                                     yaxis=dict(autorange="reversed"), height=400)
        st.plotly_chart(fig_rank_btest, use_container_width=True)
    
    # ============================================================
    # SUBJECT MARKS TRENDS - BRTEST
    # ============================================================
    if brtest_results:
        st.subheader("📊 Subject Marks Trends - BRTEST Tests (CET Format)")
        
        brtest_names = [t['Test Name'][:20] for t in brtest_results]
        
        fig_marks_brtest = go.Figure()
        fig_marks_brtest.add_trace(go.Scatter(x=brtest_names, y=[t['Physics'] for t in brtest_results], 
                                               mode='lines+markers', name='Physics (max 50)', 
                                               line=dict(color='#3498DB', width=2), marker=dict(size=8)))
        fig_marks_brtest.add_trace(go.Scatter(x=brtest_names, y=[t['Chemistry'] for t in brtest_results], 
                                               mode='lines+markers', name='Chemistry (max 50)', 
                                               line=dict(color='#9B59B6', width=2), marker=dict(size=8)))
        fig_marks_brtest.add_trace(go.Scatter(x=brtest_names, y=[t['Maths'] for t in brtest_results], 
                                               mode='lines+markers', name='Mathematics (max 100)', 
                                               line=dict(color='#F1C40F', width=2), marker=dict(size=8)))
        fig_marks_brtest.update_layout(title="Subject Marks Comparison (BRTEST - CET Format)", height=400)
        st.plotly_chart(fig_marks_brtest, use_container_width=True)
    
    # ============================================================
    # SUBJECT RANK TRENDS - BRTEST
    # ============================================================
    if brtest_results:
        st.subheader("🏆 Subject Rank Trends - BRTEST Tests (Lower is Better)")
        
        fig_ranks_brtest = go.Figure()
        
        brtest_names_rank = [t['Test Name'][:20] for t in brtest_results]
        phy_ranks_br = [t['Phy Rank'] if t['Phy Rank'] != '-' else None for t in brtest_results]
        chem_ranks_br = [t['Chem Rank'] if t['Chem Rank'] != '-' else None for t in brtest_results]
        math_ranks_br = [t['Math Rank'] if t['Math Rank'] != '-' else None for t in brtest_results]
        
        fig_ranks_brtest.add_trace(go.Scatter(x=brtest_names_rank, y=phy_ranks_br, 
                                               mode='lines+markers', name='Physics Rank',
                                               line=dict(color='#3498DB', width=2), marker=dict(size=8)))
        fig_ranks_brtest.add_trace(go.Scatter(x=brtest_names_rank, y=chem_ranks_br, 
                                               mode='lines+markers', name='Chemistry Rank',
                                               line=dict(color='#9B59B6', width=2), marker=dict(size=8)))
        fig_ranks_brtest.add_trace(go.Scatter(x=brtest_names_rank, y=math_ranks_br, 
                                               mode='lines+markers', name='Mathematics Rank',
                                               line=dict(color='#F1C40F', width=2), marker=dict(size=8)))
        
        fig_ranks_brtest.update_layout(title="Subject Rank Comparison (BRTEST)", 
                                       yaxis=dict(autorange="reversed"), height=400)
        st.plotly_chart(fig_ranks_brtest, use_container_width=True)
    
    # ============================================================
    # OVERALL RANK TREND - BRTEST
    # ============================================================
    if brtest_results:
        st.subheader("🏆 Overall Rank Trend - BRTEST Tests (Lower is Better)")
        
        fig_rank_brtest = go.Figure()
        brtest_overall_ranks = [t['Overall Rank'] for t in brtest_results]
        
        fig_rank_brtest.add_trace(go.Scatter(x=brtest_names_rank, y=brtest_overall_ranks,
                                              mode='lines+markers', name='Overall Rank',
                                              line=dict(color='#FF6B6B', width=3), marker=dict(size=10)))
        best_brtest_rank = min(brtest_overall_ranks) if brtest_overall_ranks else None
        if best_brtest_rank:
            fig_rank_brtest.add_hline(y=best_brtest_rank, line_dash="dash", line_color="green", 
                                     annotation_text=f"Best: {best_brtest_rank}")
        
        fig_rank_brtest.update_layout(title="Overall Rank Performance (BRTEST)", 
                                      yaxis=dict(autorange="reversed"), height=400)
        st.plotly_chart(fig_rank_brtest, use_container_width=True)
    
    # ============================================================
    # OVERALL PERCENTAGE TREND
    # ============================================================
    st.subheader("📈 Overall Percentage Trend")
    all_names = [t['Test Name'][:20] for t in all_tests]
    all_pcts = [float(t["%"].replace("%", "")) for t in all_tests]
    
    fig_trend = go.Figure()
    
    btest_indices = [i for i, t in enumerate(all_tests) if t['Type'] == 'BTEST']
    brtest_indices = [i for i, t in enumerate(all_tests) if t['Type'] == 'BRTEST']
    
    if btest_indices:
        fig_trend.add_trace(go.Scatter(
            x=[all_names[i] for i in btest_indices],
            y=[all_pcts[i] for i in btest_indices],
            mode='lines+markers', name='BTEST (300 marks)',
            line=dict(color='#3498DB', width=3), marker=dict(size=10)
        ))
    
    if brtest_indices:
        fig_trend.add_trace(go.Scatter(
            x=[all_names[i] for i in brtest_indices],
            y=[all_pcts[i] for i in brtest_indices],
            mode='lines+markers', name='BRTEST (200 marks)',
            line=dict(color='#E67E22', width=3), marker=dict(size=10, symbol='diamond')
        ))
    
    fig_trend.add_hline(y=75, line_dash="dash", line_color="green", annotation_text="Target (75%)")
    fig_trend.update_layout(title="Percentage Score Across All Tests", height=400)
    st.plotly_chart(fig_trend, use_container_width=True)
    
    # ============================================================
    # DETAILED WEAKNESS INSIGHTS
    # ============================================================
    st.subheader("📊 Detailed Weakness Analysis")
    
    subject_weakness = {'Physics': 0, 'Chemistry': 0, 'Maths': 0}
    total_weak = len(weak_subjects)
    
    for ws in weak_subjects:
        if ws in subject_weakness:
            subject_weakness[ws] += 1
    
    if total_weak > 0:
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Subject-wise weakness breakdown:**")
            for subject, count in sorted(subject_weakness.items(), key=lambda x: x[1], reverse=True):
                pct = (count / total_weak) * 100
                st.write(f"• {subject}: {count} times ({pct:.0f}%)")
                st.progress(int(pct))
        
        with col2:
            weak_tests_list = [t for t in all_tests if t["Weakest Subject"] not in ["Balanced", "Absent"]]
            if weak_tests_list:
                st.write("**Tests with weakness:**")
                for t in weak_tests_list[:5]:
                    st.write(f"• {t['Test Name'][:35]} → {t['Weakest Subject']}")
    
    balanced_tests_list = [t for t in all_tests if t["Weakest Subject"] == "Balanced"]
    if balanced_tests_list:
        st.write(f"**✅ Balanced performance in {len(balanced_tests_list)} tests:**")
        for t in balanced_tests_list[:5]:
            st.write(f"• {t['Test Name'][:50]}")
        if len(balanced_tests_list) > 5:
            st.write(f"... and {len(balanced_tests_list) - 5} more")
    
    # Missing tests
    attempted = set([t['Test Name'] for t in all_tests])
    missing = [t for t in test_metadata.keys() if t not in attempted]
    if missing:
        st.write(f"**⚠️ ABSENT/NO DATA for {len(missing)} tests:**")
        for m in missing[:10]:
            st.write(f"• {m} ({test_metadata[m]['type']})")
    
    # ============================================================
    # LOGOUT BUTTON
    # ============================================================
    st.markdown("---")
    if st.button("🚪 Logout", type="secondary", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.student_name = ""
        st.rerun()
    
    st.markdown("---")
    st.caption("✅ Dashboard Complete | Data Source: Master Sheet Excel | Rank Analysis: Lower number = Better performance")
