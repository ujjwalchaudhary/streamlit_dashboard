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

# --- TARGETED FIX FOR HELPER FUNCTION ---
def fix_dataframe_for_arrow(df):
    """Safely convert columns to Arrow-compatible types without triggering AttributeError."""
    if df is None or df.empty:
        return df
    df_fixed = df.copy()
    for col in df_fixed.columns:
        # Use pandas api to check type safely
        if pd.api.types.is_datetime64_any_dtype(df_fixed[col]):
            continue
        # Convert objects to string and handle NaN
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

# Add uploaded file to history
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

# Sidebar - File History
st.sidebar.header("📁 Uploaded Files History")
if len(st.session_state.file_history) > 0:
    for idx, file_info in enumerate(st.session_state.file_history):
        is_current = (idx == st.session_state.current_file_index)
        with st.sidebar.expander(f"{'✅ ' if is_current else '📄 '}{file_info['name']}", expanded=is_current):
            col1, col2 = st.columns(2)
            with col1:
                if not is_current and st.button(f"📂 Load", key=f"load_{idx}"):
                    st.session_state.current_file_index = idx
                    st.rerun()
            with col2:
                if st.button(f"🗑️ Delete", key=f"remove_{idx}"):
                    st.session_state.file_history.pop(idx)
                    st.session_state.current_file_index = None
                    st.rerun()

# Get current file
current_file = None
if st.session_state.current_file_index is not None and len(st.session_state.file_history) > 0:
    current_file = st.session_state.file_history[st.session_state.current_file_index]['file_obj']
elif uploaded_file is not None:
    current_file = uploaded_file

if current_file is not None:
    try:
        excel_file = pd.ExcelFile(current_file)
        sheet_names = excel_file.sheet_names
        
        st.sidebar.header("📋 Sheet Selection")
        combine_option = st.sidebar.radio("Data Source:", ["Select Single Sheet", "Combine All Sheets"])
        
        if combine_option == "Combine All Sheets":
            df = pd.concat([pd.read_excel(excel_file, sheet_name=s).assign(Source_Sheet=s) for s in sheet_names], ignore_index=True)
        else:
            selected_sheet = st.sidebar.selectbox("Select Sheet", sheet_names)
            df = pd.read_excel(excel_file, sheet_name=selected_sheet)

        df.columns = df.columns.str.strip()
        date_columns = ['Call Received Date', 'Call Close Date']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        df_analysis = df.copy()

        tab1, tab2, tab3, tab4 = st.tabs(["📈 Key Insights", "⚠️ Data Quality", "🔮 Future Predictions", "📋 Raw Data"])

        # TAB 1 & 2 remain as per your original logic...
        with tab1:
            st.metric("Total Records", len(df_analysis))
        with tab2:
            st.subheader("Data Quality Checks")

        # ==================== TAB 3: UPDATED FOR CAPABILITIES 1, 2 & 3 ====================
        with tab3:
            st.header("🔮 Advanced Repetition & Risk Analysis")

            # Capability 1 & 2 Logic: Targeted Column Detection
            # Looking for "Sol ID" or "Sol.Id" specifically
            sol_id_candidates = [c for c in df_analysis.columns if 'sol' in c.lower()]
            sol_col = sol_id_candidates[0] if sol_id_candidates else 'Branch'
            fault_col = 'Nature Of Fault' if 'Nature Of Fault' in df_analysis.columns else 'Call Status'

            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader(f"🔁 Capability 1: Repetitive {sol_col}s")
                rep_counts = df_analysis[sol_col].value_counts().reset_index()
                rep_counts.columns = [sol_col, 'Frequency']
                repetitive_only = rep_counts[rep_counts['Frequency'] > 1]
                
                if not repetitive_only.empty:
                    fig1 = px.bar(repetitive_only.head(10), x=sol_col, y='Frequency', color='Frequency')
                    st.plotly_chart(fig1, use_container_width=True)
                    st.write(f"**Data as Repetitive {sol_col}:**")
                    st.dataframe(fix_dataframe_for_arrow(repetitive_only), use_container_width=True)
                else:
                    st.info("No repetitive IDs found.")

            with col2:
                st.subheader("🔧 Capability 2: Most Recurring Issue")
                if fault_col in df_analysis.columns:
                    rec_issues = df_analysis[fault_col].value_counts().reset_index()
                    rec_issues.columns = ['Issue', 'Occurrences']
                    fig2 = px.pie(rec_issues.head(10), names='Issue', values='Occurrences', hole=0.4)
                    st.plotly_chart(fig2, use_container_width=True)

            st.markdown("---")
            st.subheader("🚨 Capability 3: Risk Analysis for Open Complaints")
            st.write("Cross-referencing Open Calls with repetitive IDs and common issues.")

            if 'Call Status' in df_analysis.columns:
                # Filter for Open
                df_open = df_analysis[~df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False)].copy()
                
                if not df_open.empty:
                    # Mark if the open call belongs to a repetitive ID
                    if not repetitive_only.empty:
                        df_open['Is Repetitive Site'] = df_open[sol_col].isin(repetitive_only[sol_col].tolist())
                    
                    # Sort to bring problematic sites and old calls to the top
                    sort_order = ['Is Repetitive Site'] if 'Is Repetitive Site' in df_open.columns else []
                    if 'Call Received Date' in df_open.columns: sort_order.append('Call Received Date')
                    
                    df_open_display = df_open.sort_values(by=sort_order, ascending=[False, True])
                    st.dataframe(fix_dataframe_for_arrow(df_open_display), use_container_width=True)
                else:
                    st.success("No open complaints found!")

        with tab4:
            st.dataframe(fix_dataframe_for_arrow(df_analysis), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        with st.expander("🔍 Debug Information"):
            st.code(traceback.format_exc())
else:
    st.info("Please upload a file to begin.")
