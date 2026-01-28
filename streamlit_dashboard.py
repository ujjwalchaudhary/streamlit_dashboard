import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import warnings

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
   
    # Convert all columns to appropriate types
    for col in df_fixed.columns:
        # Skip if already datetime
        if pd.api.types.is_datetime64_any_dtype(df_fixed[col]):
            continue
           
        # If object type, convert to string first
        if df_fixed[col].dtype == 'object':
            df_fixed[col] = df_fixed[col].astype(str)
            # Replace string representations of NaN/NaT
            df_fixed[col] = df_fixed[col].replace(['nan', 'NaT', 'None', '<NA>'], '')
            df_fixed[col] = df_fixed[col].replace('', None)
   
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
   
    # Check if file already in history (by name)
    file_names = [f['name'] for f in st.session_state.file_history]
   
    if uploaded_file.name not in file_names:
        st.session_state.file_history.append(file_info)
        st.session_state.current_file_index = len(st.session_state.file_history) - 1
    else:
        # Update existing file
        existing_idx = file_names.index(uploaded_file.name)
        st.session_state.file_history[existing_idx] = file_info
        st.session_state.current_file_index = existing_idx

# Sidebar - File History
st.sidebar.header("📁 Uploaded Files History")

if len(st.session_state.file_history) > 0:
    st.sidebar.write(f"**Total Files:** {len(st.session_state.file_history)}")
    st.sidebar.markdown("---")
   
    # Display each file with details
    for idx, file_info in enumerate(st.session_state.file_history):
        is_current = (idx == st.session_state.current_file_index)
       
        with st.sidebar.expander(
            f"{'✅ ' if is_current else '📄 '}{file_info['name']}",
            expanded=is_current
        ):
            st.write(f"**Size:** {file_info['size'] / 1024:.2f} KB")
            st.write(f"**Uploaded:** {file_info['upload_time']}")
           
            col1, col2 = st.columns(2)
            with col1:
                if not is_current:
                    if st.button(f"📂 Load", key=f"load_{idx}"):
                        st.session_state.current_file_index = idx
                        st.rerun()
                else:
                    st.success("Current")
           
            with col2:
                if st.button(f"🗑️ Delete", key=f"remove_{idx}"):
                    st.session_state.file_history.pop(idx)
                    if st.session_state.current_file_index == idx:
                        st.session_state.current_file_index = None
                    elif st.session_state.current_file_index and st.session_state.current_file_index > idx:
                        st.session_state.current_file_index -= 1
                    st.rerun()
   
    # Clear all button
    st.sidebar.markdown("---")
    if st.sidebar.button("🗑️ Clear All History"):
        st.session_state.file_history = []
        st.session_state.current_file_index = None
        st.rerun()
else:
    st.sidebar.info("📭 No files uploaded yet")

st.sidebar.markdown("---")

# Get current file to process
current_file = None
if st.session_state.current_file_index is not None and len(st.session_state.file_history) > 0:
    current_file = st.session_state.file_history[st.session_state.current_file_index]['file_obj']
elif uploaded_file is not None:
    current_file = uploaded_file

if current_file is not None:
    try:
        # Read all sheets from Excel file
        excel_file = pd.ExcelFile(current_file)
        sheet_names = excel_file.sheet_names
       
        st.success(f"✅ File loaded successfully! Found {len(sheet_names)} sheet(s)")
       
        # Display available sheets
        with st.expander("📑 Available Sheets in Workbook", expanded=False):
            col1, col2 = st.columns([1, 3])
            with col1:
                st.write("**Sheet Name**")
                for sheet in sheet_names:
                    st.write(f"• {sheet}")
            with col2:
                st.write("**Preview (First 3 rows)**")
                for sheet in sheet_names:
                    preview_df = pd.read_excel(excel_file, sheet_name=sheet, nrows=3)
                    preview_df = fix_dataframe_for_arrow(preview_df)
                    st.write(f"**{sheet}:**")
                    st.dataframe(preview_df, width='stretch')
                    st.markdown("---")
       
        # Sheet selection
        st.sidebar.header("📋 Sheet Selection")
       
        # Option to combine all sheets or select specific one
        combine_option = st.sidebar.radio(
            "Data Source:",
            ["Select Single Sheet", "Combine All Sheets", "Select Multiple Sheets"]
        )
       
        df = None
        selected_sheets = []
       
        if combine_option == "Select Single Sheet":
            selected_sheet = st.sidebar.selectbox("Select Sheet to Analyze", sheet_names)
            df = pd.read_excel(excel_file, sheet_name=selected_sheet)
            selected_sheets = [selected_sheet]
            st.info(f"📊 Analyzing: **{selected_sheet}**")
           
        elif combine_option == "Combine All Sheets":
            st.info(f"📊 Combining all {len(sheet_names)} sheets into one dataset")
           
            all_dfs = []
            for sheet in sheet_names:
                temp_df = pd.read_excel(excel_file, sheet_name=sheet)
                temp_df['Source_Sheet'] = sheet
                all_dfs.append(temp_df)
           
            df = pd.concat(all_dfs, ignore_index=True, sort=False)
            selected_sheets = sheet_names
            st.success(f"✅ Combined {len(all_dfs)} sheets with {len(df)} total records")
           
        else:  # Select Multiple Sheets
            selected_sheets = st.sidebar.multiselect(
                "Select Sheets to Combine",
                sheet_names,
                default=[sheet_names[0]]
            )
           
            if selected_sheets:
                st.info(f"📊 Combining {len(selected_sheets)} selected sheet(s)")
               
                all_dfs = []
                for sheet in selected_sheets:
                    temp_df = pd.read_excel(excel_file, sheet_name=sheet)
                    temp_df['Source_Sheet'] = sheet
                    all_dfs.append(temp_df)
               
                df = pd.concat(all_dfs, ignore_index=True, sort=False)
                st.success(f"✅ Combined {len(selected_sheets)} sheets with {len(df)} total records")
            else:
                st.warning("⚠️ Please select at least one sheet")
                st.stop()
       
        # Clean column names
        df.columns = df.columns.str.strip()
       
        # Store original date columns for analysis
        date_columns = ['Call Received Date', 'Tentative Date', 'Engineer Visit Date',
                       'Quote Sent', 'Call Close Date']
       
        # Convert date columns
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
       
        # Create a working copy for analysis
        df_analysis = df.copy()
       
        # Show data info
        with st.expander("ℹ️ Data Information", expanded=False):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Records", f"{len(df):,}")
            with col2:
                st.metric("Total Columns", f"{len(df.columns)}")
            with col3:
                st.metric("Sheets Used", f"{len(selected_sheets)}")
           
            st.write("**Available Columns:**")
            st.write(", ".join(df.columns.tolist()))
       
        # Sidebar filters
        st.sidebar.header("🔍 Filters")
       
        # Sheet filter
        if 'Source_Sheet' in df.columns and len(selected_sheets) > 1:
            sheet_filter = st.sidebar.multiselect(
                "Filter by Source Sheet",
                options=df['Source_Sheet'].unique().tolist(),
                default=df['Source_Sheet'].unique().tolist()
            )
            df_analysis = df_analysis[df_analysis['Source_Sheet'].isin(sheet_filter)].copy()
       
        # State filter
        if 'State' in df.columns:
            states = ['All'] + sorted(df['State'].dropna().unique().tolist())
            selected_state = st.sidebar.selectbox("Select State", states)
            if selected_state != 'All':
                df_analysis = df_analysis[df_analysis['State'] == selected_state].copy()
       
        # Branch filter
        if 'Branch' in df.columns:
            branches = ['All'] + sorted(df['Branch'].dropna().unique().tolist())
            selected_branch = st.sidebar.selectbox("Select Branch", branches)
            if selected_branch != 'All':
                df_analysis = df_analysis[df_analysis['Branch'] == selected_branch].copy()
       
        # Date range filter
        if 'Call Received Date' in df.columns:
            min_date = df_analysis['Call Received Date'].min()
            max_date = df_analysis['Call Received Date'].max()
            if pd.notna(min_date) and pd.notna(max_date):
                date_range = st.sidebar.date_input(
                    "Date Range",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date
                )
                if len(date_range) == 2:
                    df_analysis = df_analysis[
                        (df_analysis['Call Received Date'] >= pd.Timestamp(date_range[0])) &
                        (df_analysis['Call Received Date'] <= pd.Timestamp(date_range[1]))
                    ].copy()
       
        # Show filtered records count
        st.sidebar.markdown("---")
        st.sidebar.metric("Filtered Records", f"{len(df_analysis):,}")
       
        # Create tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📈 Key Insights",
            "⚠️ Data Quality",
            "🔮 Future Predictions",
            "📋 Raw Data",
            "📊 Sheet Comparison"
        ])
       
        # ==================== TAB 1: KEY INSIGHTS ====================
        with tab1:
            # KPI Metrics
            col1, col2, col3, col4 = st.columns(4)
           
            with col1:
                total_complaints = len(df_analysis)
                st.metric("Total Complaints", f"{total_complaints:,}")
           
            with col2:
                if 'Call Status' in df_analysis.columns:
                    closed = len(df_analysis[df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False)])
                    st.metric("Closed Complaints", f"{closed:,}",
                             f"{(closed/total_complaints*100):.1f}%" if total_complaints > 0 else "0%")
           
            with col3:
                if 'Call Status' in df_analysis.columns:
                    pending = len(df_analysis[~df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False)])
                    st.metric("Pending Complaints", f"{pending:,}",
                             f"{(pending/total_complaints*100):.1f}%" if total_complaints > 0 else "0%")
           
            with col4:
                if 'Call Received Date' in df_analysis.columns and 'Call Close Date' in df_analysis.columns:
                    closed_df = df_analysis[df_analysis['Call Close Date'].notna()].copy()
                    if len(closed_df) > 0:
                        avg_resolution = (closed_df['Call Close Date'] - closed_df['Call Received Date']).dt.days.mean()
                        st.metric("Avg Resolution Time", f"{avg_resolution:.1f} days")
                    else:
                        st.metric("Avg Resolution Time", "N/A")
           
            st.markdown("---")
           
            # Two columns for charts
            col1, col2 = st.columns(2)
           
            with col1:
                if 'State' in df_analysis.columns:
                    st.subheader("📍 Complaints by State")
                    state_counts = df_analysis['State'].value_counts().head(10)
                    fig = px.bar(x=state_counts.index, y=state_counts.values,
                                labels={'x': 'State', 'y': 'Number of Complaints'},
                                color=state_counts.values,
                                color_continuous_scale='Blues')
                    fig.update_layout(showlegend=False, height=400)
                    st.plotly_chart(fig)
           
            with col2:
                if 'Nature Of Fault' in df_analysis.columns:
                    st.subheader("🔧 Nature of Fault Distribution")
                    fault_counts = df_analysis['Nature Of Fault'].value_counts().head(10)
                    fig = px.pie(values=fault_counts.values, names=fault_counts.index, hole=0.4)
                    fig.update_traces(textposition='inside', textinfo='percent+label')
                    fig.update_layout(height=400)
                    st.plotly_chart(fig)
           
            # Trend Analysis
            if 'Call Received Date' in df_analysis.columns:
                st.subheader("📈 Monthly Complaint Trend")
                df_temp = df_analysis.copy()
                df_temp['Month'] = df_temp['Call Received Date'].dt.to_period('M').astype(str)
                monthly_trend = df_temp.groupby('Month').size().reset_index(name='Count')
               
                fig = px.line(monthly_trend, x='Month', y='Count', markers=True, line_shape='spline')
                fig.update_layout(height=400, xaxis_tickangle=-45)
                st.plotly_chart(fig)
           
            # Branch Performance
            if 'Branch' in df_analysis.columns and 'Call Status' in df_analysis.columns:
                st.subheader("🏢 Branch Performance")
                branch_status = pd.crosstab(df_analysis['Branch'],
                                           df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False))
                branch_status.columns = ['Pending', 'Closed']
               
                fig = go.Figure(data=[
                    go.Bar(name='Closed', x=branch_status.index, y=branch_status['Closed'], marker_color='green'),
                    go.Bar(name='Pending', x=branch_status.index, y=branch_status['Pending'], marker_color='orange')
                ])
                fig.update_layout(barmode='stack', height=400, xaxis_tickangle=-45)
                st.plotly_chart(fig)
       
        # ==================== TAB 2: DATA QUALITY ====================
        with tab2:
            st.header("⚠️ Data Quality Analysis")
           
            st.subheader("📊 Missing Data Overview")
           
            missing_data = pd.DataFrame({
                'Column': df_analysis.columns,
                'Missing Count': df_analysis.isnull().sum().values,
                'Missing %': (df_analysis.isnull().sum().values / len(df_analysis) * 100).round(2)
            })
            missing_data = missing_data[missing_data['Missing Count'] > 0].sort_values('Missing Count', ascending=False)
           
            if len(missing_data) > 0:
                fig = px.bar(missing_data, x='Column', y='Missing %',
                            text='Missing %', color='Missing %',
                            color_continuous_scale='Reds')
                fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                fig.update_layout(height=400, xaxis_tickangle=-45)
                st.plotly_chart(fig)
               
                # Fix missing_data for display
                missing_data_display = fix_dataframe_for_arrow(missing_data)
                st.dataframe(missing_data_display, width='stretch')
            else:
                st.success("✅ No missing data found!")
           
            st.markdown("---")
           
            col1, col2 = st.columns(2)
           
            with col1:
                st.subheader("🚨 Critical Issues")
               
                issues = []
               
                if 'Call Status' in df_analysis.columns and 'Call Close Date' in df_analysis.columns:
                    closed_no_date = len(df_analysis[
                        (df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False)) &
                        (df_analysis['Call Close Date'].isna())
                    ])
                    if closed_no_date > 0:
                        issues.append(f"❌ {closed_no_date} closed complaints without close date")
               
                if 'Engineer Visit Date' in df_analysis.columns and 'Quote Sent' in df_analysis.columns:
                    visit_no_quote = len(df_analysis[
                        (df_analysis['Engineer Visit Date'].notna()) &
                        (df_analysis['Quote Sent'].isna())
                    ])
                    if visit_no_quote > 0:
                        issues.append(f"⚠️ {visit_no_quote} complaints with engineer visit but no quote")
               
                if 'Call Received Date' in df_analysis.columns and 'Engineer Visit Date' in df_analysis.columns:
                    df_temp = df_analysis.copy()
                    df_temp['Visit Delay'] = (df_temp['Engineer Visit Date'] - df_temp['Call Received Date']).dt.days
                    delayed_visits = len(df_temp[df_temp['Visit Delay'] > 7])
                    if delayed_visits > 0:
                        issues.append(f"⏰ {delayed_visits} complaints with >7 days visit delay")
               
                if issues:
                    for issue in issues:
                        st.warning(issue)
                else:
                    st.success("✅ No critical issues found!")
           
            with col2:
                st.subheader("📌 Recommendations")
               
                st.info("💡 **Data Completeness:**")
                st.write("- Ensure all closed complaints have close dates")
                st.write("- Track material status for all complaints")
                st.write("- Standardize fault categories")
               
                st.info("💡 **Process Improvement:**")
                st.write("- Reduce engineer visit delays")
                st.write("- Improve quote turnaround time")
                st.write("- Implement automated status updates")
           
           st.subheader("🔍 Duplicate Analysis")

if 'Complaint No.' in df_analysis.columns:
    duplicates = df_analysis[
        df_analysis.duplicated(subset=['Complaint No.'], keep=False)
    ].copy()

    if len(duplicates) > 0:
        st.warning(f"⚠️ Found {len(duplicates)} duplicate complaint records")

        duplicates_display = fix_dataframe_for_arrow(duplicates)
        st.dataframe(duplicates_display, width='stretch')

    else:
        st.success("✅ No duplicate complaints found")

else:
    st.info("ℹ️ 'Complaint No.' column not available for duplicate check") 
       
        # ==================== TAB 3: FUTURE PREDICTIONS ====================
        with tab3:
            st.header("🔮 Future Predictions & Insights")
           
            if 'Call Received Date' in df_analysis.columns:
                st.subheader("📊 Complaint Volume Forecast")
               
                monthly_data = df_analysis.groupby(df_analysis['Call Received Date'].dt.to_period('M')).size()
                monthly_data.index = monthly_data.index.to_timestamp()
               
                if len(monthly_data) >= 3:
                    last_3_months_avg = monthly_data.tail(3).mean()
                    trend = (monthly_data.tail(3).values[-1] - monthly_data.tail(3).values[0]) / 2
                    next_month_prediction = last_3_months_avg + trend
                   
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Last Month", f"{int(monthly_data.iloc[-1]):,}")
                    with col2:
                        st.metric("Predicted Next Month", f"{int(next_month_prediction):,}",
                                 f"{((next_month_prediction - monthly_data.iloc[-1])/monthly_data.iloc[-1]*100):.1f}%")
                    with col3:
                        st.metric("3-Month Average", f"{int(last_3_months_avg):,}")
                   
                    future_months = pd.date_range(monthly_data.index[-1], periods=4, freq='M')[1:]
                    prediction_df = pd.DataFrame({
                        'Month': list(monthly_data.index) + list(future_months),
                        'Complaints': list(monthly_data.values) + [next_month_prediction]*3,
                        'Type': ['Actual']*len(monthly_data) + ['Predicted']*3
                    })
                   
                    fig = px.line(prediction_df, x='Month', y='Complaints', color='Type',
                                 markers=True, line_dash='Type')
                    fig.update_layout(height=400)
                    st.plotly_chart(fig)
           
            st.markdown("---")
           
            col1, col2 = st.columns(2)
           
            with col1:
                st.subheader("🎯 High Priority Areas")
               
                if 'State' in df_analysis.columns and 'Call Status' in df_analysis.columns:
                    pending_by_state = df_analysis[
                        ~df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False)
                    ].groupby('State').size().sort_values(ascending=False).head(5)
                   
                    fig = px.bar(x=pending_by_state.index, y=pending_by_state.values,
                                labels={'x': 'State', 'y': 'Pending Complaints'},
                                color=pending_by_state.values,
                                color_continuous_scale='Reds')
                    fig.update_layout(height=350, showlegend=False)
                    st.plotly_chart(fig)
           
            with col2:
                st.subheader("⚡ Action Items")
               
                if 'Call Received Date' in df_analysis.columns and 'Call Close Date' in df_analysis.columns:
                    pending_df = df_analysis[df_analysis['Call Close Date'].isna()].copy()
                    if len(pending_df) > 0:
                        pending_df['Days Open'] = (datetime.now() - pending_df['Call Received Date']).dt.days
                        old_complaints = len(pending_df[pending_df['Days Open'] > 30])
                       
                        st.error(f"🚨 {old_complaints} complaints open >30 days")
                        st.warning(f"⚠️ {len(pending_df)} total pending complaints")
                       
                        st.write("**Oldest Open Complaints:**")
                        display_cols = ['Complaint No.', 'Branch', 'State', 'Days Open', 'Nature Of Fault']
                        if 'Source_Sheet' in df_analysis.columns:
                            display_cols.append('Source_Sheet')
                        oldest = pending_df.nlargest(5, 'Days Open')[display_cols]
                        oldest_display = fix_dataframe_for_arrow(oldest)
                        st.dataframe(oldest_display, width='stretch')
       
        # ==================== TAB 4: RAW DATA ====================
        with tab4:
            st.header("📋 Raw Data")
           
            search = st.text_input("🔍 Search in data", "")
           
            # Fix dataframe for display
            df_display = fix_dataframe_for_arrow(df_analysis)
           
            if search:
                mask = df_display.astype(str).apply(lambda x: x.str.contains(search, case=False, na=False)).any(axis=1)
                filtered_df = df_display[mask]
                st.write(f"Found {len(filtered_df)} matching records")
                st.dataframe(filtered_df, width='stretch')
            else:
                st.dataframe(df_display, width='stretch')
           
            st.subheader("💾 Download Options")
           
            col1, col2 = st.columns(2)
            with col1:
                csv = df_analysis.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Filtered Data as CSV",
                    data=csv,
                    file_name='complaint_data_filtered.csv',
                    mime='text/csv',
                )
           
            with col2:
                summary = df_analysis.describe(include='all').to_csv().encode('utf-8')
                st.download_button(
                    label="📊 Download Summary Report",
                    data=summary,
                    file_name='summary_report.csv',
                    mime='text/csv',
                )
       
        # ==================== TAB 5: SHEET COMPARISON ====================
        with tab5:
            st.header("📊 Sheet Comparison Analysis")
           
            if 'Source_Sheet' in df_analysis.columns and len(selected_sheets) > 1:
                st.subheader("📈 Sheet-wise Statistics")
               
                sheet_stats = df_analysis.groupby('Source_Sheet').agg({
                    df_analysis.columns[0]: 'count'
                }).rename(columns={df_analysis.columns[0]: 'Total Records'})
               
                if 'Call Status' in df_analysis.columns:
                    closed_by_sheet = df_analysis[
                        df_analysis['Call Status'].str.contains('Close|Closed', case=False, na=False)
                    ].groupby('Source_Sheet').size()
                    sheet_stats['Closed'] = closed_by_sheet
                    sheet_stats['Pending'] = sheet_stats['Total Records'] - sheet_stats['Closed'].fillna(0)
                    sheet_stats['Closure Rate %'] = (sheet_stats['Closed'] / sheet_stats['Total Records'] * 100).round(2)
               
                sheet_stats_display = fix_dataframe_for_arrow(sheet_stats.reset_index())
                st.dataframe(sheet_stats_display, width='stretch')
               
                col1, col2 = st.columns(2)
               
                with col1:
                    st.subheader("Records per Sheet")
                    fig = px.bar(sheet_stats, y='Total Records',
                                color='Total Records',
                                color_continuous_scale='Viridis')
                    fig.update_layout(height=400, showlegend=False)
                    st.plotly_chart(fig)
               
                with col2:
                    if 'Closure Rate %' in sheet_stats.columns:
                        st.subheader("Closure Rate by Sheet")
                        fig = px.bar(sheet_stats, y='Closure Rate %',
                                    color='Closure Rate %',
                                    color_continuous_scale='RdYlGn')
                        fig.update_layout(height=400, showlegend=False)
                        st.plotly_chart(fig)
               
                if 'Call Received Date' in df_analysis.columns:
                    st.subheader("📅 Monthly Trends Comparison")
                   
                    df_temp = df_analysis.copy()
                    df_temp['Month'] = df_temp['Call Received Date'].dt.to_period('M').astype(str)
                    monthly_by_sheet = df_temp.groupby(['Month', 'Source_Sheet']).size().reset_index(name='Count')
                   
                    fig = px.line(monthly_by_sheet, x='Month', y='Count',
                                 color='Source_Sheet', markers=True)
                    fig.update_layout(height=400, xaxis_tickangle=-45)
                    st.plotly_chart(fig)
               
            else:
                st.info("ℹ️ Sheet comparison is available when multiple sheets are combined or selected")
                st.write("To use this feature:")
                st.write("1. Select 'Combine All Sheets' or 'Select Multiple Sheets'")
                st.write("2. The dashboard will show comparison metrics across sheets")
       
    except Exception as e:
        st.error(f"❌ Error loading file: {str(e)}")
        with st.expander("🔍 Debug Information"):
            st.write(f"Error type: {type(e).__name__}")
            st.write(f"Error message: {str(e)}")
import traceback
st.code(traceback.format_exc())
else:
st.info("👆 Please upload an Excel file to get started")
st.subheader("📝 How to Use")
st.write("""
1. **Upload your Excel file** - Supports both single and multi-sheet workbooks
2. **View file history** - See all uploaded files in the sidebar
3. **Select data source** - Single sheet, combine all, or select multiple
4. **Apply filters** - State, Branch, Date range
5. **Explore insights** - Key metrics, trends, predictions
6. **Download reports** - Export filtered data or summaries
""")

st.subheader("📋 Expected Data Format")
sample_data = pd.DataFrame({
    'Complaint No.': ['C001', 'C002', 'C003'],
    'Branch': ['Mumbai', 'Delhi', 'Bangalore'],
    'State': ['Maharashtra', 'Delhi', 'Karnataka'],
    'Call Received Date': ['2024-01-15', '2024-01-16', '2024-01-17'],
    'Nature Of Fault': ['Hardware', 'Software', 'Network'],
    'Call Status': ['Closed', 'Pending', 'Closed']
})
st.dataframe(sample_data, width='stretch')
st.markdown("---")
st.markdown(
"""

Made with ❤️ using Streamlit | Multi-Sheet Dashboard with File History v3.1

""",
unsafe_allow_html=True
) 
