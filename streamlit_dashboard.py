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

# --- IMPROVED HELPER FUNCTION ---
def fix_dataframe_for_arrow(df):
    """Convert all problematic columns to Arrow-compatible types safely"""
    if df is None or df.empty:
        return df
    df_fixed = df.copy()
    
    for col in df_fixed.columns:
        # Check for datetime safely using pandas api
        if pd.api.types.is_datetime64_any_dtype(df_fixed[col]):
            continue
            
        # Convert object/mixed types to string and handle nulls
        df_fixed[col] = df_fixed[col].astype(str).replace(['nan', 'NaT', 'None', '<NA>'], '')
    
    return df_fixed

# Title
st.title("📊 Complaint Management Analytics Dashboard")
st.markdown("---")

# File upload
uploaded_file = st.file_uploader(
    "📁 Upload your Excel file (Supports: .xlsx, .xls, .xlsm with multi-sheet workbooks)",
    type=['xlsx', 'xls', 'xlsm'],
    help="Your file can contain multiple sheets. Macros will be ignored."
)

# Add uploaded file to history logic
if uploaded_file is not None:
    file_info = {
        'name': uploaded_file.name,
        'size': uploaded_file.size,
        'upload_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'file_obj': uploaded_file
    }
    file_names = [f['name'] for f in st.session_state.file_history]
    if uploaded_file.name not in file_names:
        st.session_state.file_history.append(file_info)
        st.session_state.current_file_index = len(st.session_state.file_history) - 1
    else:
        existing_idx = file_names.index(uploaded_file.name)
        st.session_state.file_history[existing_idx] = file_info
        st.session_state.current_file_index = existing_idx

# Sidebar - History
st.sidebar.header("📁 Uploaded Files History")
if len(st.session_state.file_history) > 0:
    for idx, file_info in enumerate(st.session_state.file_history):
        is_current = (idx == st.session_state.current_file_index)
        with st.sidebar.expander(f"{'✅ ' if is_current else '📄 '}{file_info['name']}", expanded=is_current):
            if not is_current:
                if st.button(f"Load", key=f"load_{idx}"):
                    st.session_state.current_file_index = idx
                    st.rerun()
            if st.button(f"Delete", key=f"remove_{idx}"):
                st.session_state.file_history.pop(idx)
                st.rerun()

# Processing Data
current_file = None
if st.session_state.current_file_index is not None and len(st.session_state.file_history) > 0:
    current_file = st.session_state.file_history[st.session_state.current_file_index]['file_obj']

if current_file is not None:
    try:
        excel_file = pd.ExcelFile(current_file)
        sheet_names = excel_file.sheet_names
        
        st.sidebar.header("📋 Sheet Selection")
        combine_option = st.sidebar.radio("Data Source:", ["Select Single Sheet", "Combine All Sheets", "Select Multiple Sheets"])
        
        df = None
        if combine_option == "Select Single Sheet":
            selected_sheet = st.sidebar.selectbox("Select Sheet", sheet_names)
            df = pd.read_excel(excel_file, sheet_name=selected_sheet)
        elif combine_option == "Combine All Sheets":
            df = pd.concat([pd.read_excel(excel_file, sheet_name=s).assign(Source_Sheet=s) for s in sheet_names], ignore_index=True)
        else:
            sel_sheets = st.sidebar.multiselect("Select Sheets", sheet_names, default=[sheet_names[0]])
            df = pd.concat([pd.read_excel(excel_file, sheet_name=s).assign(Source_Sheet=s) for s in sel_sheets], ignore_index=True)

        df.columns = df.columns.str.strip()
        df_analysis = df.copy()

        # Standard Date Conversion
        date_cols = ['Call Received Date', 'Call Close Date', 'Engineer Visit Date']
        for col in date_cols:
            if col in df_analysis.columns:
                df_analysis[col] = pd.to_datetime(df_analysis[col], errors='coerce')

        # TABS
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Key Insights", "⚠️ Data Quality", "🔮 Repetition Analysis", "📋 Raw Data", "📊 Sheet Comparison"])

        with tab1:
            st.metric("Total Records", len(df_analysis))
            if 'Nature Of Fault' in df_analysis.columns:
                fig = px.pie(df_analysis, names='Nature Of Fault', title="Fault Distribution")
                st.plotly_chart(fig, use_container_width=True)

        with tab2:
            st.subheader("Data Health")
            st.write(df_analysis.isnull().sum())

        # ==================== TAB 3: UPDATED REPETITION LOGIC ====================
        with tab3:
            st.header("🔍 Comprehensive Repetition Report")
            
            # Detect Sol ID column (handles variations like Sol.Id, Sol ID, sol_id)
            sol_col = next((c for c in df_analysis.columns if 'sol' in c.lower()), None)
            fault_col = 'Nature Of Fault' if 'Nature Of Fault' in df_analysis.columns else None

            if sol_col:
                # Calculate counts for ALL IDs
                all_counts = df_analysis[sol_col].value_counts().reset_index()
                all_counts.columns = [sol_col, 'Total Complaints']
                
                # Filter: Keep ALL that have more than 1 occurrence
                repetitive_only = all_counts[all_counts['Total Complaints'] > 1]
                
                col1, col2 = st.columns([1, 1])
                
                with col1:
                    st.subheader(f"📋 All Repetitive {sol_col}s")
                    st.write(f"Found **{len(repetitive_only)}** unique Sol IDs with recurring issues.")
                    st.dataframe(fix_dataframe_for_arrow(repetitive_only), use_container_width=True)

                with col2:
                    st.subheader("🔧 Fault Recursion Analysis")
                    if fault_col:
                        fault_rec = df_analysis[fault_col].value_counts().reset_index()
                        fault_rec.columns = ['Fault Type', 'Frequency']
                        st.dataframe(fix_dataframe_for_arrow(fault_rec), use_container_width=True)
                
                st.markdown("---")
                st.subheader("🚨 High Priority: Open Complaints at Repeat Sites")
                if 'Call Status' in df_analysis.columns:
                    # Filter for Open/Pending and for IDs in our repetitive list
                    open_mask = ~df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False)
                    repeat_list = repetitive_only[sol_col].tolist()
                    risk_df = df_analysis[open_mask & df_analysis[sol_col].isin(repeat_list)]
                    
                    if not risk_df.empty:
                        st.error(f"Alert: {len(risk_df)} open complaints are active at sites with history of repeated failure.")
                        st.dataframe(fix_dataframe_for_arrow(risk_df), use_container_width=True)
                    else:
                        st.success("No active open complaints at repetitive sites.")
            else:
                st.warning("Sol ID column not detected. Please ensure your data has a 'Sol ID' column.")

        with tab4:
            st.dataframe(fix_dataframe_for_arrow(df_analysis), use_container_width=True)

        with tab5:
            st.info("Sheet comparison metrics are displayed here.")

    except Exception as e:
        st.error(f"Error: {e}")
        st.code(traceback.format_exc())
else:
    st.info("Please upload a file to begin.")
