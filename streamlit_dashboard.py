import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import warnings
import traceback

# Suppress warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="Complaint Management Dashboard",
    page_icon="📊",
    layout="wide"
)

# Initialize session state for file history
if 'file_history' not in st.session_state:
    st.session_state.file_history = []

if 'current_file_index' not in st.session_state:
    st.session_state.current_file_index = None

# --- CRITICAL FIX: Robust Arrow Compatibility ---
def fix_dataframe_for_arrow(df):
    """Safely convert columns to Arrow-compatible types to prevent AttributeError."""
    if df is None or df.empty:
        return df
    
    df_fixed = df.copy()
    for col in df_fixed.columns:
        try:
            # Check for datetime safely
            if pd.api.types.is_datetime64_any_dtype(df_fixed[col]):
                continue
            
            # Convert objects/mixed types to string and handle NaN
            df_fixed[col] = df_fixed[col].astype(str).replace(['nan', 'NaT', 'None', '<NA>'], '')
        except:
            continue
    return df_fixed

# Title
st.title("📊 Complaint Management Analytics Dashboard")
st.markdown("---")

# File upload
uploaded_file = st.file_uploader(
    "📁 Upload your Excel file",
    type=['xlsx', 'xls', 'xlsm']
)

# File History Logic
if uploaded_file is not None:
    file_info = {'name': uploaded_file.name, 'size': uploaded_file.size, 
                 'upload_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 'file_obj': uploaded_file}
    file_names = [f['name'] for f in st.session_state.file_history]
    if uploaded_file.name not in file_names:
        st.session_state.file_history.append(file_info)
        st.session_state.current_file_index = len(st.session_state.file_history) - 1

# Sidebar - History
st.sidebar.header("📁 File History")
if len(st.session_state.file_history) > 0:
    for idx, file_info in enumerate(st.session_state.file_history):
        is_current = (idx == st.session_state.current_file_index)
        if st.sidebar.button(f"{'✅' if is_current else '📄'} {file_info['name']}", key=f"f_{idx}"):
            st.session_state.current_file_index = idx
            st.rerun()

# Processing
current_file = None
if st.session_state.current_file_index is not None and len(st.session_state.file_history) > 0:
    current_file = st.session_state.file_history[st.session_state.current_file_index]['file_obj']

if current_file is not None:
    try:
        excel_file = pd.ExcelFile(current_file)
        sheet_names = excel_file.sheet_names
        
        selected_sheet = st.sidebar.selectbox("Select Sheet", sheet_names)
        df = pd.read_excel(excel_file, sheet_name=selected_sheet)
        df.columns = df.columns.str.strip()
        
        # Standard Date Conversion
        for col in ['Call Received Date', 'Call Close Date', 'Engineer Visit Date']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        df_analysis = df.copy()

        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📈 Key Insights", "⚠️ Data Quality", "🔮 Repetition Analysis", "📋 Raw Data"])

        # ==================== TAB 1: KEY INSIGHTS ====================
        with tab1:
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Records", len(df_analysis))
            if 'Call Status' in df_analysis.columns:
                closed = len(df_analysis[df_analysis['Call Status'].str.contains('Close', case=False, na=False)])
                col2.metric("Closed", closed)
                col3.metric("Pending", len(df_analysis) - closed)
            
            # Simple Chart for Insights
            if 'Nature Of Fault' in df_analysis.columns:
                fault_dist = df_analysis['Nature Of Fault'].value_counts()
                fig_fault = px.bar(fault_dist, title="Fault Distribution")
                st.plotly_chart(fig_fault, use_container_width=True)

        # ==================== TAB 2: DATA QUALITY ====================
        with tab2:
            st.subheader("Missing Data & Duplicates")
            st.write(df_analysis.isnull().sum().rename("Missing Values"))
            
            if 'Complaint No.' in df_analysis.columns:
                dupes = df_analysis[df_analysis.duplicated('Complaint No.', keep=False)]
                if not dupes.empty:
                    st.warning(f"Found {len(dupes)} duplicates")
                    st.dataframe(fix_dataframe_for_arrow(dupes))

        # ==================== TAB 3: ALL REPETITIVE SOL IDs ====================
        with tab3:
            st.header("🔍 Comprehensive Repetition Report")
            
            # Detection logic for Sol ID and Nature of Fault
            sol_col = next((c for c in df_analysis.columns if 'sol' in c.lower()), None)
            fault_col = 'Nature Of Fault' if 'Nature Of Fault' in df_analysis.columns else None

            if sol_col:
                # 1. Count ALL Repetitive Sol IDs (Count > 1)
                all_counts = df_analysis[sol_col].value_counts().reset_index()
                all_counts.columns = [sol_col, 'Total Occurrences']
                
                # Filter to keep ONLY those that are repetitive (more than 1)
                repetitive_only = all_counts[all_counts['Total Occurrences'] > 1]
                
                col1, col2 = st.columns([1, 1])
                
                with col1:
                    st.subheader(f"📋 All Repetitive {sol_col}s")
                    st.write(f"Found **{len(repetitive_only)}** unique Sol IDs with recurring complaints.")
                    st.dataframe(fix_dataframe_for_arrow(repetitive_only), use_container_width=True)

                with col2:
                    st.subheader("🔧 Nature of Fault Recursion")
                    if fault_col:
                        recursion_issue = df_analysis[fault_col].value_counts().reset_index()
                        recursion_issue.columns = ['Fault Type', 'Frequency']
                        st.dataframe(fix_dataframe_for_arrow(recursion_issue), use_container_width=True)
                    else:
                        st.info("Nature of Fault column not found.")

                st.markdown("---")
                # 3. High Risk Table
                st.subheader("🚨 Details of Open Complaints at Repetitive Sites")
                if 'Call Status' in df_analysis.columns:
                    # Filter for Pending/Open
                    open_mask = ~df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False)
                    # Filter for Repetitive Sol IDs
                    rep_list = repetitive_only[sol_col].tolist()
                    risk_mask = df_analysis[sol_col].isin(rep_list)
                    
                    high_risk_df = df_analysis[open_mask & risk_mask]
                    
                    if not high_risk_df.empty:
                        st.error(f"Attention: There are {len(high_risk_df)} open complaints at sites with a history of failure.")
                        st.dataframe(fix_dataframe_for_arrow(high_risk_df), use_container_width=True)
                    else:
                        st.success("No open complaints found for the repetitive Sol IDs.")
            else:
                st.error("Sol ID column not detected. Please ensure your Excel has a 'Sol ID' column.")

        # ==================== TAB 4: RAW DATA ====================
        with tab4:
            st.dataframe(fix_dataframe_for_arrow(df_analysis), use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")
        st.code(traceback.format_exc())

st.markdown("---")
st.caption("Dashboard v3.2 - Focused Repetition Analysis")
