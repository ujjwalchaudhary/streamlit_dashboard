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

# --- FIXED HELPER FUNCTION (Resolves the AttributeError) ---
def fix_dataframe_for_arrow(df):
    """Safely convert columns to Arrow-compatible types for Streamlit display."""
    if df is None or df.empty:
        return df
   
    df_fixed = df.copy()
    for col in df_fixed.columns:
        try:
            # Safely check for datetime types using pandas API
            if pd.api.types.is_datetime64_any_dtype(df_fixed[col]):
                continue
           
            # Convert objects/mixed types to string to avoid Arrow serialization errors
            df_fixed[col] = df_fixed[col].astype(str).replace(['nan', 'NaT', 'None', '<NA>'], '')
        except Exception:
            # Fallback for unexpected column structures
            continue
    return df_fixed

# Title
st.title("📊 Complaint Management Analytics Dashboard")
st.markdown("---")

# File upload
uploaded_file = st.file_uploader(
    "📁 Upload your Excel file (Supports: .xlsx, .xls, .xlsm)",
    type=['xlsx', 'xls', 'xlsm'],
    help="Your file can contain multiple sheets."
)

# File History Logic
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

# Sidebar - History UI
st.sidebar.header("📁 Uploaded Files History")
if len(st.session_state.file_history) > 0:
    for idx, file_info in enumerate(st.session_state.file_history):
        is_current = (idx == st.session_state.current_file_index)
        with st.sidebar.expander(f"{'✅ ' if is_current else '📄 '}{file_info['name']}", expanded=is_current):
            if st.button(f"📂 Load File", key=f"load_{idx}"):
                st.session_state.current_file_index = idx
                st.rerun()
            if st.button(f"🗑️ Remove", key=f"del_{idx}"):
                st.session_state.file_history.pop(idx)
                st.session_state.current_file_index = None
                st.rerun()

# Processing Current File
current_file = None
if st.session_state.current_file_index is not None and len(st.session_state.file_history) > 0:
    current_file = st.session_state.file_history[st.session_state.current_file_index]['file_obj']

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
       
        # Standard Date Conversion
        date_cols = ['Call Received Date', 'Call Close Date', 'Engineer Visit Date']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        df_analysis = df.copy()

        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📈 Key Insights", "⚠️ Data Quality", "🔮 Future Predictions & Repetition", "📋 Raw Data"])

        # ==================== TAB 1: INSIGHTS ====================
        with tab1:
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Records", len(df_analysis))
            if 'Call Status' in df_analysis.columns:
                closed = len(df_analysis[df_analysis['Call Status'].str.contains('Close', case=False, na=False)])
                col2.metric("Closed", closed)
                col3.metric("Pending", len(df_analysis) - closed)

        # ==================== TAB 2: QUALITY ====================
        with tab2:
            st.subheader("🔍 Duplicate Complaint Numbers")
            if 'Complaint No.' in df_analysis.columns:
                dupes = df_analysis[df_analysis.duplicated('Complaint No.', keep=False)]
                if not dupes.empty:
                    st.warning(f"Found {len(dupes)} duplicate complaint entries.")
                    st.dataframe(fix_dataframe_for_arrow(dupes), use_container_width=True)
                else:
                    st.success("No duplicates found.")

        # ==================== TAB 3: ENHANCED REPETITION ====================
        with tab3:
            st.header("🔮 Advanced Repetition & Recurring Issues")
           
            # Map columns for Capability 1 & 2
            # Identify Sol ID (using 'Branch' or 'Sol ID' depending on column availability)
            sol_col = 'Sol ID' if 'Sol ID' in df_analysis.columns else ('Branch' if 'Branch' in df_analysis.columns else None)
            fault_col = 'Nature Of Fault' if 'Nature Of Fault' in df_analysis.columns else None

            if sol_col and fault_col:
                c1, c2 = st.columns(2)

                with c1:
                    st.subheader("🔁 Capability 1: Repetitive Sol IDs")
                    rep_counts = df_analysis[sol_col].value_counts().reset_index()
                    rep_counts.columns = [sol_col, 'Frequency']
                    repetitive_list = rep_counts[rep_counts['Frequency'] > 1]
                   
                    if not repetitive_list.empty:
                        fig1 = px.bar(repetitive_list.head(10), x=sol_col, y='Frequency', color='Frequency', title="Top 10 Problematic Sites")
                        st.plotly_chart(fig1, use_container_width=True)
                        st.write("**Data as Repetitive Sol ID:**")
                        st.dataframe(fix_dataframe_for_arrow(repetitive_list), use_container_width=True)
                    else:
                        st.info("No sites with multiple complaints found.")

                with c2:
                    st.subheader("🔧 Capability 2: Most Recurring Issue")
                    recurring_issue = df_analysis[fault_col].value_counts().reset_index()
                    recurring_issue.columns = ['Issue', 'Occurrences']
                    fig2 = px.pie(recurring_issue.head(10), names='Issue', values='Occurrences', hole=0.4)
                    st.plotly_chart(fig2, use_container_width=True)

                st.markdown("---")
               
                # Capability 3: High-Risk Open Complaints
                st.subheader("🚨 Capability 3: Risk Analysis for Open Complaints")
                st.write("Cross-referencing Open Calls with repetitive Sol IDs and common issues.")
               
                if 'Call Status' in df_analysis.columns:
                    # Filter for Pending/Open
                    df_open = df_analysis[~df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False)].copy()
                   
                    if not df_open.empty:
                        # Identify if the Sol ID has a history of repetition
                        repeat_ids = repetitive_list[sol_col].tolist()
                        df_open['Is Repetitive Site'] = df_open[sol_col].isin(repeat_ids)
                       
                        # Rank by repetition and age of complaint
                        sort_list = ['Is Repetitive Site']
                        if 'Call Received Date' in df_open.columns: sort_list.append('Call Received Date')
                       
                        df_open_risk = df_open.sort_values(by=sort_list, ascending=[False, True])
                       
                        st.dataframe(fix_dataframe_for_arrow(df_open_risk), use_container_width=True)
                    else:
                        st.success("Great news! There are no open complaints.")
            else:
                st.error("Missing 'Sol ID' or 'Nature Of Fault' columns required for Capability 1, 2, and 3.")

        # ==================== TAB 4: RAW DATA ====================
        with tab4:
            st.dataframe(fix_dataframe_for_arrow(df_analysis), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error loading file: {str(e)}")
        with st.expander("🔍 Debug Information"):
            st.code(traceback.format_exc())
else:
    st.info("👆 Please upload an Excel file to see the dashboard analysis.")

st.markdown("---")
st.markdown("<div style='text-align: center;'>Complaint Management Dashboard v3.5</div>", unsafe_allow_html=True)
