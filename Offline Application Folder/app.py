# pyrefly: ignore [missing-import]
import streamlit as st
import os
import json
import pandas as pd
from scripts.engine import ExcelEngine

st.set_page_config(page_title="DB Merging Application", layout="wide")

def inject_custom_css():
    st.markdown("""
        <style>
        /* Smooth, dark theme gradient for titles */
        h1 {
            background: linear-gradient(90deg, #A8BFFF, #884DFF);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 2.8rem !important;
            font-weight: 800 !important;
        }
        
        /* Glassmorphism for Streamlit tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 24px;
            padding-bottom: 10px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            border-radius: 4px 4px 0px 0px;
            padding: 10px 16px;
            background-color: transparent;
            transition: all 0.3s;
        }
        
        /* Soft, premium button animations */
        div.stButton > button {
            border-radius: 8px;
            transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
            font-weight: 500;
        }
        div.stButton > button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(255, 255, 255, 0.08);
            border-color: #884DFF;
            color: #A8BFFF;
        }
        
        /* Primary Button glow */
        div.stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: none;
            color: white;
        }
        div.stButton > button[kind="primary"]:hover {
            background: linear-gradient(135deg, #764ba2 0%, #667eea 100%);
            transform: translateY(-2px);
            box-shadow: 0 6px 15px rgba(118, 75, 162, 0.4);
            border: none;
            color: white;
        }
        </style>
    """, unsafe_allow_html=True)

inject_custom_css()

# Initialize Session State
if 'merge_queue' not in st.session_state:
    st.session_state['merge_queue'] = []
if 'merge_results' not in st.session_state:
    st.session_state.merge_results = None

# --- Initialize Paths ---
# Dynamically resolve BASE_DIR to the folder where app.py lives.
# This makes the app fully portable — no hardcoded paths needed.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
MAPPER_PATH = os.path.join(BASE_DIR, "Column_Mapper.xlsx")

# --- Load Config & Engine ---
if not os.path.exists(CONFIG_PATH) or not os.path.exists(MAPPER_PATH):
    st.error("Project files missing. Please run the setup scripts.")
    st.stop()


def load_config():
    with open(CONFIG_PATH, 'r') as f:
        return json.load(f)


def save_config(cfg):
    with open(CONFIG_PATH, 'w') as f:
        json.dump(cfg, f, indent=4)


def browse_file(title="Select File", filetypes=None):
    """Open a native Windows file picker and return the chosen path (or empty string)."""
    import subprocess
    if filetypes is None:
        filter_str = "Excel files (*.xlsx;*.xlsb;*.xlsm;*.xls)|*.xlsx;*.xlsb;*.xlsm;*.xls|All files (*.*)|*.*"
    else:
        filters = []
        for name, ext in filetypes:
            ext_ps = ext.replace(" ", ";")
            filters.append(f"{name} ({ext})|{ext_ps}")
        filter_str = "|".join(filters)
        
    ps_script = f"""
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $f = New-Object System.Windows.Forms.OpenFileDialog
    $f.Title = '{title}'
    $f.Filter = '{filter_str}'
    $f.AutoUpgradeEnabled = $true
    if ($f.ShowDialog() -eq 'OK') {{
        Write-Output $f.FileName
    }}
    """
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    res = subprocess.run(["powershell", "-Sta", "-NoProfile", "-Command", ps_script], 
                         capture_output=True, text=True, startupinfo=startupinfo)
    if res.stderr:
        print(f"PowerShell Error (browse_file): {res.stderr}")
    return res.stdout.strip()


def browse_folder(title="Select Folder"):
    """Open a modern Windows 10/11 folder picker via IFileOpenDialog COM interface."""
    import subprocess
    # Use IFileOpenDialog COM interface — the same API Windows Explorer uses.
    # FolderBrowserDialog (even with AutoUpgradeEnabled) does not reliably
    # show the modern dialog from a subprocess context.
    ps_script = """
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport]
[ClassInterface(ClassInterfaceType.None)]
[Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
class FileOpenDialogRCW {}

[ComImport]
[Guid("42F85136-DB7E-439C-85F1-E4075D135FC8")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IFileOpenDialog {
    [PreserveSig] int Show(IntPtr hwnd);
    void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
    void SetFileTypeIndex(uint iFileType);
    void GetFileTypeIndex(out uint piFileType);
    void Advise(IntPtr pfde, out uint pdwCookie);
    void Unadvise(uint dwCookie);
    void SetOptions(uint fos);
    void GetOptions(out uint pfos);
    void SetDefaultFolder(IShellItem psi);
    void SetFolder(IShellItem psi);
    void GetFolder(out IShellItem ppsi);
    void GetCurrentSelection(out IShellItem ppsi);
    void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
    void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
    void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
    void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
    void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
    void GetResult(out IShellItem ppsi);
    void AddPlace(IShellItem psi, uint fdap);
    void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
    void Close(int hr);
    void SetClientGuid([In] ref Guid guid);
    void ClearClientData();
    void SetFilter(IntPtr pFilter);
    void GetResults(out IntPtr ppenum);
    void GetSelectedItems(out IntPtr ppsai);
}

[ComImport]
[Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IShellItem {
    void BindToHandler(IntPtr pbc, [In] ref Guid bhid, [In] ref Guid riid, out IntPtr ppv);
    void GetParent(out IShellItem ppsi);
    void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
    void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    void Compare(IShellItem psi, uint hint, out int piOrder);
}

public class FolderPicker {
    const uint FOS_PICKFOLDERS = 0x00000020;
    const uint FOS_FORCEFILESYSTEM = 0x00000040;
    const uint SIGDN_FILESYSPATH = 0x80058000;

    public static string Show(string title) {
        var dialog = (IFileOpenDialog)new FileOpenDialogRCW();
        try {
            uint options;
            dialog.GetOptions(out options);
            dialog.SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
            dialog.SetTitle(title);
            int hr = dialog.Show(IntPtr.Zero);
            if (hr == 0) {
                IShellItem item;
                dialog.GetResult(out item);
                string path;
                item.GetDisplayName(SIGDN_FILESYSPATH, out path);
                Marshal.ReleaseComObject(item);
                return path;
            }
        } finally {
            Marshal.ReleaseComObject(dialog);
        }
        return null;
    }
}
"@ -Language CSharp

$path = [FolderPicker]::Show('FOLDER_TITLE')
if ($path) { Write-Output $path }
""".replace('FOLDER_TITLE', title)
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    res = subprocess.run(["powershell", "-Sta", "-NoProfile", "-Command", ps_script],
                         capture_output=True, text=True, startupinfo=startupinfo)
    if res.stderr:
        print(f"PowerShell Error (browse_folder): {res.stderr}")
    return res.stdout.strip()



# ── Sidebar ─────────────────────────────────────────────────────────────────
st.sidebar.title("Project Settings")
mode = st.sidebar.radio("Select Dashboard Mode", ["Export", "Import"])

# Dashboard file picker (per mode)
cfg = load_config()
db_key = "import_dashboard_path" if mode == "Import" else "export_dashboard_path"
current_db_path = cfg.get(db_key, "")

st.sidebar.markdown("---")
st.sidebar.markdown(f"**{mode} Dashboard File**")
if current_db_path:
    st.sidebar.markdown(f"📄 **{os.path.basename(current_db_path)}**")
    st.sidebar.code(os.path.dirname(current_db_path), language="plaintext")
else:
    st.sidebar.caption("📄 *(not set)*")

if st.sidebar.button(f"📂 Browse {mode} DB File…", use_container_width=True, key="sidebar_browse_db"):
    picked = browse_file(title=f"Select {mode} Dashboard File")
    if picked:
        cfg[db_key] = picked.replace("\\", "/")
        save_config(cfg)
        st.rerun()

# ── Engine ───────────────────────────────────────────────────────────────────
# Always re-read config after potential browse updates
engine = ExcelEngine(MAPPER_PATH, CONFIG_PATH, mode=mode)

st.title(f"DB Merging App - {mode} Dashboard")

tab1, tab2, tab3 = st.tabs(["📊 Merge Data", "🗺️ Column Mapper", "⚙️ Settings"])

# ── Tab 1: Merge Data ─────────────────────────────────────────────────────────
with tab1:
    st.header("Merge Monthly Data")

    col1, col2 = st.columns(2)
    with col1:
        base_folder = engine.config.get('base_files_folder', '')

        # Base folder picker row
        folder_label_col, folder_btn_col = st.columns([7, 3])
        with folder_label_col:
            st.markdown("**Base Files Folder**")
            if base_folder:
                st.code(base_folder, language="plaintext")
            else:
                st.info("*(not set — click Browse)*")
        with folder_btn_col:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📁 Browse", key="tab1_browse_folder", use_container_width=True):
                picked_folder = browse_folder(title="Select Base Files Folder")
                if picked_folder:
                    cfg2 = load_config()
                    cfg2['base_files_folder'] = picked_folder.replace("\\", "/")
                    save_config(cfg2)
                    st.rerun()

        # File list from chosen folder
        if base_folder and os.path.exists(base_folder):
            files = [
                f for f in os.listdir(base_folder)
                if f.endswith(('.xlsx', '.xlsb', '.xlsm', '.xls')) and not f.startswith('~$')
            ]
            if files:
                selected_file = st.selectbox("Select Base File to Merge", files)
                file_path = os.path.join(base_folder, selected_file) if selected_file else None
            else:
                st.warning("No Excel files found in the selected folder.")
                selected_file = None
                file_path = None
        else:
            if base_folder:
                st.warning(f"Folder not found: `{base_folder}`")
            else:
                st.info("Click **📁 Browse…** to choose the folder containing base files.")
            selected_file = None
            file_path = None

    # Mapping selection
    mapping_keys = list(engine.mappings.keys())
    country_key = st.selectbox("Assign mapping for this file", mapping_keys)

    with col2:
        import datetime
        today = datetime.date.today()
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
                st.caption(f"Selected Range: :blue[{start.strftime('%d-%b-%Y')}] to :blue[{end.strftime('%d-%b-%Y')}]")
                periods = pd.period_range(start=start, end=end, freq='M')
                valid_pairs = [(p.year, p.strftime('%b')) for p in periods]
                unique_years = sorted(list(set(p.year for p in periods)))
                st.info(f"📊 Range covers {len(valid_pairs)} months from {len(unique_years)} year(s).")
            elif len(date_range) == 1:
                start = date_range[0]
                st.caption(f"Selected: :blue[{start.strftime('%d-%b-%Y')}]")
                valid_pairs = [(start.year, start.strftime('%b'))]
                st.info(f"📊 Single month selected: {start.strftime('%b %Y')}")

        target_months = valid_pairs
        target_years = []

    # Merge buttons
    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("Run Single Merge", type="secondary", use_container_width=True):
            if file_path and country_key and target_months:
                with st.spinner("Executing surgical merge... Excel is working in the background."):
                    success, msg = engine.merge_file(file_path, country_key, target_months, target_years)
                    if success:
                        st.success(msg)
                    else:
                        st.error(f"Merge Failed: {msg}")
            else:
                st.warning("Please select a file, mapping, and at least one month/year.")

    with col_btn2:
        if st.button("➕ Add to Batch Queue", type="primary"):
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

    # Queue display
    if st.session_state['merge_queue']:
        st.divider()
        st.subheader(f"📋 Pending Batch Queue ({len(st.session_state['merge_queue'])} files)")

        # Column headers
        h1, h2, h3, h4 = st.columns([4, 2, 4, 1])
        with h1:
            st.markdown("**📄 File**")
        with h2:
            st.markdown("**🗺️ Mapping**")
        with h3:
            st.markdown("**📅 Months**")
        with h4:
            st.markdown("**Del**")
        st.markdown("<hr style='margin:4px 0 8px 0; border-color: rgba(255,255,255,0.1)'>", unsafe_allow_html=True)

        # Per-item rows with individual delete buttons
        for i, item in enumerate(list(st.session_state['merge_queue'])):
            c1, c2, c3, c4 = st.columns([4, 2, 4, 1])
            with c1:
                st.markdown(f"{item['file']}")
            with c2:
                st.markdown(f"{item['key']}")
            with c3:
                months_str = ", ".join([f"{m} {y}" for y, m in item['month']]) if item['month'] else "—"
                st.caption(months_str)
            with c4:
                if st.button("🗑️", key=f"del_queue_{i}", help=f"Remove '{item['file']}' from queue"):
                    st.session_state['merge_queue'].pop(i)
                    st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        q_col1, q_col2 = st.columns([1, 4])
        with q_col1:
            if st.button("🚀 Process All Queued Files", type="primary"):
                with st.spinner("Processing batch... please wait."):
                    results = engine.merge_files_batch(st.session_state['merge_queue'])
                    st.session_state.merge_results = results
                    st.session_state['merge_queue'] = []
                    st.rerun()
        with q_col2:
            if st.button("🗑️ Clear Queue"):
                st.session_state['merge_queue'] = []
                st.rerun()

    # Batch results
    if st.session_state.merge_results:
        st.divider()
        st.subheader("🏁 Last Batch Result Summary")
        res_df = pd.DataFrame(st.session_state.merge_results)

        def color_status(val):
            color = 'red' if val == 'Failed' else 'green'
            return f'color: {color}'

        st.dataframe(res_df.style.map(color_status, subset=['status']), use_container_width=True)
        if st.button("Clear Results"):
            st.session_state.merge_results = None
            st.rerun()

# ── Tab 2: Column Mapper ──────────────────────────────────────────────────────
with tab2:
    st.header("Excel Column Mapper")
    st.success("To change mappings, open **Column_Mapper.xlsx** in Excel, edit it, and save. The app picks up changes automatically.")
    df_map = pd.read_excel(MAPPER_PATH, sheet_name=mode)
    st.dataframe(df_map, use_container_width=True)
    if st.button("Open Folder to Edit Mapper"):
        os.startfile(BASE_DIR)

# ── Tab 3: Settings ───────────────────────────────────────────────────────────
with tab3:
    st.header("⚙️ Global Settings")
    cfg_live = load_config()

    st.subheader("Dashboard Files")

    # Export dashboard
    c1, c2 = st.columns([5, 1])
    with c1:
        export_path = cfg_live.get('export_dashboard_path', 'not set')
        st.markdown(f"**Export Dashboard**")
        if export_path != 'not set':
            st.code(export_path, language="plaintext")
        else:
            st.info(export_path)
    with c2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📂 Browse…", key="settings_browse_export"):
            picked = browse_file(title="Select Export Dashboard File")
            if picked:
                cfg_live['export_dashboard_path'] = picked.replace("\\", "/")
                save_config(cfg_live)
                st.rerun()

    # Import dashboard
    c3, c4 = st.columns([5, 1])
    with c3:
        import_path = cfg_live.get('import_dashboard_path', 'not set')
        st.markdown(f"**Import Dashboard**")
        if import_path != 'not set':
            st.code(import_path, language="plaintext")
        else:
            st.info(import_path)
    with c4:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📂 Browse…", key="settings_browse_import"):
            picked = browse_file(title="Select Import Dashboard File")
            if picked:
                cfg_live['import_dashboard_path'] = picked.replace("\\", "/")
                save_config(cfg_live)
                st.rerun()

    st.divider()
    st.subheader("Base Files Folder")

    c5, c6 = st.columns([5, 1])
    with c5:
        base_path = cfg_live.get('base_files_folder', 'not set')
        st.markdown("**Base Files Folder**")
        if base_path != 'not set':
            st.code(base_path, language="plaintext")
        else:
            st.info(base_path)
    with c6:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📁 Browse…", key="settings_browse_base"):
            picked_folder = browse_folder(title="Select Base Files Folder")
            if picked_folder:
                cfg_live['base_files_folder'] = picked_folder.replace("\\", "/")
                save_config(cfg_live)
                st.rerun()

    st.divider()
    st.markdown("**Mapping File:**")
    st.code(MAPPER_PATH, language="plaintext")
