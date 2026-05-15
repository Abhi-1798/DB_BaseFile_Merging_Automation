# pyrefly: ignore [missing-import]
import streamlit as st
import os
import json
import pandas as pd
from scripts.engine import ExcelEngine

st.set_page_config(page_title="LAB DB Merging App", layout="wide")

# Initialize Session State
if 'merge_queue' not in st.session_state:
    st.session_state['merge_queue'] = []
if 'merge_results' not in st.session_state:
    st.session_state.merge_results = None

# --- Initialize Paths ---
BASE_DIR = r"D:\DB Merging Automation\DB Merging Automation"
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
# Use Excel Mapper instead of JSON
MAPPER_PATH = os.path.join(BASE_DIR, "Column_Mapper.xlsx")

# --- Dashboard Mode Selection ---
st.sidebar.title("🛠️ Project Settings")
mode = st.sidebar.radio("Select Dashboard Mode", ["Export", "Import"])

# --- Load Config & Engine ---
if not os.path.exists(CONFIG_PATH) or not os.path.exists(MAPPER_PATH):
    st.error("Project files missing. Please run the setup scripts.")
    st.stop()

# Initialize engine based on selected mode
engine = ExcelEngine(MAPPER_PATH, CONFIG_PATH, mode=mode)

st.title(f"🚀 LAB DB Merging - {mode} Dashboard")

tab1, tab2, tab3 = st.tabs(["📊 Merge Data", "🗺️ Column Mapper", "⚙️ Settings"])

with tab1:
    st.header("Merge Monthly Data")
    
    col1, col2 = st.columns(2)
    with col1:
        # File selector for base files
        base_folder = engine.config['base_files_folder']
        if os.path.exists(base_folder):
            files = [f for f in os.listdir(base_folder) if f.endswith('.xlsx') and not f.startswith('~$')]
            selected_file = st.selectbox("Select Base File to Import", files)
            file_path = os.path.join(base_folder, selected_file) if selected_file else None
        else:
            st.warning(f"Base folder not found: {base_folder}")
            file_path = None

    # Identifying Mapping Selection (Needed before date discovery)
    mapping_keys = list(engine.mappings.keys())
    country_key = st.selectbox("Assign mapping for this file", mapping_keys)

    with col2:
        # --- Calendar Range Selection ---
        import datetime
        today = datetime.date.today()
        # Default to a 2-month range
        default_start = (today.replace(day=1) - datetime.timedelta(days=1)).replace(day=1)
        
        date_range = st.date_input(
            "Select Date Range",
            value=[default_start, today],
            min_value=datetime.date(2021, 1, 1),
            max_value=datetime.date(2030, 12, 31),
            format="DD-MM-YYYY"
        )

        valid_pairs = []
        if isinstance(date_range, (list, tuple)):
            if len(date_range) == 2:
                start, end = date_range
                # Formatted preview in the user's preferred style
                st.caption(f"Selected Range: :blue[{start.strftime('%d-%b-%Y')}] to :blue[{end.strftime('%d-%b-%Y')}]")
                
                # Generate all months within the range
                periods = pd.period_range(start=start, end=end, freq='M')
                valid_pairs = [(p.year, p.strftime('%b')) for p in periods]
                
                # Display summary
                unique_years = sorted(list(set(p.year for p in periods)))
                st.info(f"📊 Range covers {len(valid_pairs)} months from {len(unique_years)} year(s).")
            elif len(date_range) == 1:
                start = date_range[0]
                st.caption(f"Selected: :blue[{start.strftime('%d-%b-%Y')}]")
                valid_pairs = [(start.year, start.strftime('%b'))]
                st.info(f"📊 Single month selected: {start.strftime('%b %Y')}")
        
        # For compatibility with engine calling signature
        # We pass the list of (Year, Month) pairs into the target_months argument
        target_months = valid_pairs
        target_years = [] # Not needed when passing pairs

    # --- Merge Buttons ---
    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("Run Single Merge", type="secondary", use_container_width=True):
            if file_path and country_key and target_months:
                with st.spinner("Executing surgical merge... Processing large dataset (100MB+), Excel is working in background."):
                    success, msg = engine.merge_file(file_path, country_key, target_months, target_years)
                    if success: st.success(msg)
                    else: st.error(f"Merge Failed: {msg}")
            else:
                st.warning("Please select a file, mapping, and at least one month/year.")
    
    with col_btn2:
        if st.button("➕ Add to Batch Queue", type="primary", use_container_width=False):
            if file_path and country_key and target_months:
                st.session_state['merge_queue'].append({
                    "file": selected_file,
                    "path": file_path,
                    "key": country_key,
                    "month": target_months,
                    "year": target_years
                })
                st.success(f"Added {selected_file} to queue ({len(target_months)} months).")
            else:
                st.error("Select file, mapping, and dates first.")

    # --- Queue Display ---
    if st.session_state['merge_queue']:
        st.divider()
        st.subheader(f"📋 Pending Batch Queue ({len(st.session_state['merge_queue'])} files)")
        
        df_queue = pd.DataFrame(st.session_state['merge_queue'])
        st.table(df_queue[["file", "key", "month", "year"]])
        
        q_col1, q_col2 = st.columns([1, 4])
        with q_col1:
            if st.button("🚀 Process All Queued Files", type="primary"):
                with st.spinner("Processing batch... Handling massive datasets (1M rows), please wait."):
                    results = engine.merge_files_batch(st.session_state['merge_queue'])
                    st.session_state.merge_results = results
                    st.session_state['merge_queue'] = []
                    st.rerun()
        with q_col2:
            if st.button("🗑️ Clear Queue"):
                st.session_state['merge_queue'] = []
                st.rerun()

    # --- Batch Results ---
    if st.session_state.merge_results:
        st.divider()
        st.subheader("🏁 Last Batch Result Summary")
        res_df = pd.DataFrame(st.session_state.merge_results)
        
        # Color coding for status
        def color_status(val):
            color = 'red' if val == 'Failed' else 'green'
            return f'color: {color}'
        
        st.dataframe(res_df.style.applymap(color_status, subset=['status']), use_container_width=True)
        if st.button("Clear Results"):
            st.session_state.merge_results = None
            st.rerun()

with tab2:
    st.header("Excel Column Mapper")
    st.success("To change mappings, simply open the **Column_Mapper.xlsx** file in Excel, edit it, and save. The app will use your changes automatically.")
    
    # Show the current mappings in a nice table
    df_map = pd.read_excel(MAPPER_PATH, sheet_name=mode)
    st.dataframe(df_map, use_container_width=True)
    
    if st.button("Open Folder to Edit Mapper"):
        os.startfile(BASE_DIR)

with tab3:
    st.header("Global Settings")
    st.write(f"**Current Dashboard:** {engine.dashboard_path}")
    st.write(f"**Base Files Folder:** {engine.config['base_files_folder']}")
    st.write(f"**Mapping File:** {MAPPER_PATH}")
    
    if st.button("Reset Config to Defaults"):
        # Logic to reset
        pass
