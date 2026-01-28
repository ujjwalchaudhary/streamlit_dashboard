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

# Helper function to fix dataframe for Arrow compatibility
def fix_dataframe_for_arrow(df):
    """Convert all problematic columns to Arrow-compatible types"""
    df_fixed = df.copy()
    for col in df_fixed.columns:
        if pd.api.types.is_datetime64_any_dtype(df_fixed[col]):
            continue
        if df_fixed[col].dtype == 'object':
            df_fixed[col] = df_fixed[col].astype(str)
            df_fixed[col] = df_fixed[col].replace(['nan', 'NaT', 'None', '<NA>'], '')
            df_fixed[col] = df_fixed[col].replace('', None)
    return df_fixed

# Title
st.title("📊 Complaint Management Analytics Dashboard")
st.markdown("---")

# File upload
uploaded_file = st.file_uploader(
    "📁 Upload your Excel file (Supports: .xlsx, .xls, .xlsm)",
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

# Sidebar - File History
st.sidebar.header("📁 Uploaded Files History")
if len(st.session_state.file_history) > 0:
    for idx, file_info in enumerate(st.session_state.file_history):
        is_current = (idx == st.session_state.current_file_index)
        with st.sidebar.expander(f"{'✅ ' if is_current else '📄 '}{file_info['name']}", expanded=is_current):
            st.write(f"**Uploaded:** {file_info['upload_time']}")
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
    if st.sidebar.button("🗑️ Clear All History"):
        st.session_state.file_history = []
        st.session_state.current_file_index = None
        st.rerun()
else:
    st.sidebar.info("📭 No files uploaded yet")

# Load current file
current_file = None
if st.session_state.current_file_index is not None and len(st.session_state.file_history) > 0:
    current_file = st.session_state.file_history[st.session_state.current_file_index]['file_obj']
elif uploaded_file is not None:
    current_file = uploaded_file

if current_file is not None:
    try:
        excel_file = pd.ExcelFile(current_file)
        sheet_names = excel_file.sheet_names
        st.success(f"✅ File loaded successfully! Found {len(sheet_names)} sheet(s)")

        # Sheet selection logic
        st.sidebar.header("📋 Sheet Selection")
        combine_option = st.sidebar.radio("Data Source:", ["Select Single Sheet", "Combine All Sheets", "Select Multiple Sheets"])
       
        df = pd.DataFrame()
        selected_sheets = []

        if combine_option == "Select Single Sheet":
            selected_sheet = st.sidebar.selectbox("Select Sheet", sheet_names)
            df = pd.read_excel(excel_file, sheet_name=selected_sheet)
            selected_sheets = [selected_sheet]
        elif combine_option == "Combine All Sheets":
            df = pd.concat([pd.read_excel(excel_file, sheet_name=s).assign(Source_Sheet=s) for s in sheet_names], ignore_index=True)
            selected_sheets = sheet_names
        else:
            selected_sheets = st.sidebar.multiselect("Select Sheets", sheet_names, default=[sheet_names[0]])
            if selected_sheets:
                df = pd.concat([pd.read_excel(excel_file, sheet_name=s).assign(Source_Sheet=s) for s in selected_sheets], ignore_index=True)

        if not df.empty:
            df.columns = df.columns.str.strip()
            date_cols = ['Call Received Date', 'Tentative Date', 'Engineer Visit Date', 'Quote Sent', 'Call Close Date']
            for col in date_cols:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')

            df_analysis = df.copy()

            # Filters
            if 'State' in df_analysis.columns:
                states = ['All'] + sorted(df_analysis['State'].dropna().unique().tolist())
                sel_state = st.sidebar.selectbox("Filter State", states)
                if sel_state != 'All':
                    df_analysis = df_analysis[df_analysis['State'] == sel_state]

            tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Key Insights", "⚠️ Data Quality", "🔮 Predictions", "📋 Raw Data", "📊 Comparison"])

            with tab1:
                # KPI Metrics
                c1, c2, c3 = st.columns(3)
                c1.metric("Total Records", len(df_analysis))
                if 'Call Status' in df_analysis.columns:
                    closed = len(df_analysis[df_analysis['Call Status'].str.contains('Close', case=False, na=False)])
                    c2.metric("Closed", closed)
                    c3.metric("Pending", len(df_analysis) - closed)
               
                # Chart: State distribution
                if 'State' in df_analysis.columns:
                    st.subheader("📍 Complaints by State")
                    fig = px.bar(df_analysis['State'].value_counts().head(10))
                    st.plotly_chart(fig, use_container_width=True)

            with tab2:
                st.subheader("Missing Values")
                missing = df_analysis.isnull().sum()
                st.write(missing[missing > 0])

            with tab3:
                st.info("Forecasting and priority analysis would appear here based on date trends.")

            with tab4:
                st.dataframe(fix_dataframe_for_arrow(df_analysis), use_container_width=True)

            with tab5:
                if 'Source_Sheet' in df_analysis.columns:
                    sheet_counts = df_analysis['Source_Sheet'].value_counts()
                    st.plotly_chart(px.pie(names=sheet_counts.index, values=sheet_counts.values, title="Data Split by Sheet"))

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        with st.expander("🔍 Traceback"):
            st.code(traceback.format_exc())
else:
    st.info("👆 Please upload an Excel file to get started")
    # This part shows the landing page instructions
    st.subheader("📝 How to Use")
    st.write("1. Upload file. 2. Use Sidebar to filter. 3. Explore tabs.")
   
    sample_data = pd.DataFrame({
        'Complaint No.': ['C001', 'C002'],
        'Branch': ['Mumbai', 'Delhi'],
        'Call Status': ['Closed', 'Pending']
    })
    st.write("**Required Format Example:**")
    st.dataframe(sample_data, use_container_width=True)

st.markdown("---")
st.markdown("Made with ❤️ using Streamlit", unsafe_allow_html=True)

 
