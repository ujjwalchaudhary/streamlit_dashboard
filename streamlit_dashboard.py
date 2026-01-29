import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import warnings
import traceback

warnings.filterwarnings("ignore")

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Complaint Management Dashboard",
    page_icon="📊",
    layout="wide"
)

# ---------------- SESSION STATE ----------------
if "file_history" not in st.session_state:
    st.session_state.file_history = []

if "current_file_index" not in st.session_state:
    st.session_state.current_file_index = None


# ---------------- HELPER ----------------
def fix_dataframe_for_arrow(df):
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).replace(
                ["nan", "NaT", "None", "<NA>"], ""
            )
    return df


# ---------------- TITLE ----------------
st.title("📊 Complaint Management Analytics Dashboard")
st.markdown("---")

# ---------------- FILE UPLOAD ----------------
uploaded_file = st.file_uploader(
    "📁 Upload Excel file",
    type=["xlsx", "xls", "xlsm"]
)

if uploaded_file:
    file_info = {
        "name": uploaded_file.name,
        "size": uploaded_file.size,
        "upload_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_obj": uploaded_file,
    }

    names = [f["name"] for f in st.session_state.file_history]
    if uploaded_file.name not in names:
        st.session_state.file_history.append(file_info)
        st.session_state.current_file_index = len(
            st.session_state.file_history
        ) - 1
    else:
        idx = names.index(uploaded_file.name)
        st.session_state.file_history[idx] = file_info
        st.session_state.current_file_index = idx

# ---------------- SIDEBAR ----------------
st.sidebar.header("📁 File History")

if st.session_state.file_history:
    for idx, f in enumerate(st.session_state.file_history):
        with st.sidebar.expander(
            f"{'✅' if idx == st.session_state.current_file_index else '📄'} {f['name']}",
            expanded=(idx == st.session_state.current_file_index),
        ):
            st.write(f"Size: {f['size']/1024:.1f} KB")
            st.write(f"Uploaded: {f['upload_time']}")

            if idx != st.session_state.current_file_index:
                if st.button("Load", key=f"load_{idx}"):
                    st.session_state.current_file_index = idx
                    st.rerun()

# ---------------- CURRENT FILE ----------------
current_file = None
if st.session_state.current_file_index is not None:
    current_file = st.session_state.file_history[
        st.session_state.current_file_index
    ]["file_obj"]

# ---------------- MAIN LOGIC ----------------
if current_file:
    try:
        excel = pd.ExcelFile(current_file)
        sheet_names = excel.sheet_names

        st.success(f"Loaded {len(sheet_names)} sheet(s)")

        st.sidebar.header("📋 Sheet Selection")
        mode = st.sidebar.radio(
            "Choose data source",
            ["Single Sheet", "Combine All Sheets"],
        )

        if mode == "Single Sheet":
            sheet = st.sidebar.selectbox("Select Sheet", sheet_names)
            df = pd.read_excel(excel, sheet_name=sheet)
        else:
            dfs = []
            for s in sheet_names:
                temp = pd.read_excel(excel, sheet_name=s)
                temp["Source_Sheet"] = s
                dfs.append(temp)
            df = pd.concat(dfs, ignore_index=True)

        df.columns = df.columns.str.strip()

        # Date columns
        date_cols = [
            "Call Received Date",
            "Engineer Visit Date",
            "Call Close Date",
        ]
        for c in date_cols:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")

        df_analysis = df.copy()

        # ---------------- FILTERS ----------------
        st.sidebar.header("🔍 Filters")

        if "State" in df_analysis.columns:
            state = st.sidebar.selectbox(
                "State",
                ["All"] + sorted(df_analysis["State"].dropna().unique()),
            )
            if state != "All":
                df_analysis = df_analysis[df_analysis["State"] == state]

        if "Branch" in df_analysis.columns:
            branch = st.sidebar.selectbox(
                "Branch",
                ["All"] + sorted(df_analysis["Branch"].dropna().unique()),
            )
            if branch != "All":
                df_analysis = df_analysis[df_analysis["Branch"] == branch]

        # ---------------- TABS ----------------
        tab1, tab2, tab3, tab4 = st.tabs(
            ["📈 Insights", "⚠️ Data Quality", "🔮 Future & Recurrence", "📋 Raw Data"]
        )

        # ================= TAB 1 =================
        with tab1:
            st.metric("Total Complaints", len(df_analysis))

            if "Call Status" in df_analysis.columns:
                closed = df_analysis[
                    df_analysis["Call Status"]
                    .str.contains("close", case=False, na=False)
                ]
                st.metric("Closed", len(closed))

        # ================= TAB 2 =================
        with tab2:
            missing = (
                df_analysis.isnull().sum()
                .reset_index()
                .rename(columns={"index": "Column", 0: "Missing"})
            )
            missing = missing[missing["Missing"] > 0]
            st.dataframe(missing, use_container_width=True)

        # ================= TAB 3 (UPDATED) =================
        with tab3:
            st.header("🔮 Forecast & Repetitive Issue Intelligence")

            # ---- Forecast ----
            if "Call Received Date" in df_analysis.columns:
                monthly = (
                    df_analysis.groupby(
                        df_analysis["Call Received Date"].dt.to_period("M")
                    )
                    .size()
                )

                if len(monthly) >= 3:
                    st.line_chart(monthly)

            st.markdown("---")
            st.subheader("🔁 Repetitive Issue Intelligence")

            col1, col2 = st.columns(2)

          # ---- REPEATED SOL ID (ACTUAL RECURSION) ----
with col1:
    if 'SOL ID' in df_analysis.columns:
        st.write("🏦 Repeated SOL IDs (Recursion Analysis)")

        sol_repeat_df = (
            df_analysis['SOL ID']
            .dropna()
            .astype(str)
            .value_counts()
            .reset_index()
        )
        sol_repeat_df.columns = ['SOL ID', 'Total Complaints']

        # 🔥 keep only repeated SOL IDs
        sol_repeat_df = sol_repeat_df[sol_repeat_df['Total Complaints'] > 1]

        if not sol_repeat_df.empty:
            sol_repeat_df = sol_repeat_df.sort_values(
                'Total Complaints', ascending=False
            ).head(15)

            fig = px.bar(
                sol_repeat_df,
                x='SOL ID',
                y='Total Complaints',
                color='Total Complaints',
                color_continuous_scale='Blues'
            )
            fig.update_layout(showlegend=False, height=350)
            st.plotly_chart(fig, use_container_width=True)

            st.dataframe(
                fix_dataframe_for_arrow(sol_repeat_df),
                use_container_width=True
            )
        else:
            st.success("✅ No repeated SOL IDs found")


            # ---- Repeated Nature of Fault ----
            with col2:
                if "Nature Of Fault" in df_analysis.columns:
                    fault_df = (
                        df_analysis["Nature Of Fault"]
                        .dropna()
                        .astype(str)
                        .value_counts()
                        .head(10)
                        .reset_index()
                    )
                    fault_df.columns = ["Nature Of Fault", "Count"]

                    if not fault_df.empty:
                        st.write("🔧 Most Recurring Faults")
                        st.plotly_chart(
                            px.bar(
                                fault_df,
                                x="Nature Of Fault",
                                y="Count",
                                color="Count",
                            ),
                            use_container_width=True,
                        )
                        st.dataframe(
                            fix_dataframe_for_arrow(fault_df),
                            use_container_width=True,
                        )

            # ---- SOL × Fault Hotspots ----
            if (
                "SOL ID" in df_analysis.columns
                and "Nature Of Fault" in df_analysis.columns
            ):
                st.subheader("🔥 Chronic SOL × Fault Hotspots")

                hotspot = (
                    df_analysis.dropna(
                        subset=["SOL ID", "Nature Of Fault"]
                    )
                    .groupby(["SOL ID", "Nature Of Fault"])
                    .size()
                    .reset_index(name="Repeat Count")
                    .sort_values("Repeat Count", ascending=False)
                )

                hotspot = hotspot[hotspot["Repeat Count"] > 1].head(15)

                if not hotspot.empty:
                    st.warning(
                        "Repeated combinations indicate **structural / chronic issues**"
                    )
                    st.dataframe(
                        fix_dataframe_for_arrow(hotspot),
                        use_container_width=True,
                    )
                else:
                    st.success("No recurring SOL-Fault patterns found")

        # ================= TAB 4 =================
        with tab4:
            st.dataframe(
                fix_dataframe_for_arrow(df_analysis),
                use_container_width=True,
            )

    except Exception as e:
        st.error("❌ Error loading file")
        st.code(traceback.format_exc())

else:
    st.info("👆 Upload an Excel file to begin")

st.markdown("---")
st.caption("Built with ❤️ using Streamlit")

