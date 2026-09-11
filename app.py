import io
import os
import glob
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import streamlit as st

# Page setup
st.set_page_config(
    page_title="SanSan Data Matcher",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        color: #4B5563;
        font-size: 1.05rem;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-title {
        font-size: 0.9rem;
        color: #64748B;
        font-weight: 600;
        text-transform: uppercase;
    }
    .metric-val {
        font-size: 1.8rem;
        font-weight: 700;
        margin-top: 4px;
    }
    .badge-match {
        color: #15803D;
        background-color: #DCFCE7;
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 600;
    }
    .badge-mismatch {
        color: #B91C1C;
        background-color: #FEE2E2;
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 600;
    }
    .badge-missing {
        color: #B45309;
        background-color: #FEF3C7;
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Pre-defined standard fields configuration
DEFAULT_FIELDS = {
    'Item Select':   {'excel_name': 'Item Select',   'default_col': 'Item Select',   'type': 'item_select'},
    'Date':          {'excel_name': 'Date',          'default_col': 'Date',          'type': 'bcard', 'csv_temp': 'Date'},
    'Email':         {'excel_name': 'Email',         'default_col': 'Email',         'type': 'bcard', 'csv_temp': 'Email'},
    'Fax':           {'excel_name': 'Fax',           'default_col': 'Fax',           'type': 'bcard', 'csv_temp': 'FAX'},
    'Full Name':     {'excel_name': 'Full Name',     'default_col': 'Full Name',     'type': 'bcard', 'csv_temp': 'full name'},
    'MOBILE':        {'excel_name': 'MOBILE',        'default_col': 'MOBILE',        'type': 'bcard', 'csv_temp': 'mobile'},
    'TEL':           {'excel_name': 'TEL',           'default_col': 'TEL',           'type': 'bcard', 'csv_temp': 'TEL'},
    'URL':           {'excel_name': 'URL',           'default_col': 'URL',           'type': 'bcard', 'csv_temp': 'URL'},
    'Address':       {'excel_name': 'Address',       'default_col': 'Address',       'type': 'bcard', 'csv_temp': 'Address'},
    'Company Name':  {'excel_name': 'Company Name',  'default_col': 'Company Name',  'type': 'bcard', 'csv_temp': 'company'},
    'Division Name': {'excel_name': 'Division Name', 'default_col': 'Division Name', 'type': 'bcard', 'csv_temp': 'division'},
    'Position Name': {'excel_name': 'Position Name', 'default_col': 'Position Name', 'type': 'bcard', 'csv_temp': 'position'},
}

def to_float_or_str(val):
    if val is None or val == '' or (isinstance(val, str) and val.strip() == ''):
        return None
    try:
        return round(float(val), 2)
    except (ValueError, TypeError):
        return str(val).strip()

def generate_excel_report(df_results):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comparison Summary"

    headers = list(df_results.columns)
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    for col_num, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    match_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # light green
    mismatch_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # light red/orange
    missing_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # light yellow

    for row_data in df_results.values.tolist():
        ws.append(row_data)
        curr_row = ws.max_row
        status = str(row_data[-1])
        if status == "MATCH":
            row_fill = match_fill
        elif status == "MISMATCH":
            row_fill = mismatch_fill
        else:
            row_fill = missing_fill

        for c in range(1, len(row_data) + 1):
            ws.cell(row=curr_row, column=c).fill = row_fill

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# Sidebar - Source Selection
st.sidebar.title("📁 File Source")
data_source = st.sidebar.radio("Choose Input Mode:", ["Use Local Workspace Files", "Upload New Files"])

excel_file = None
csv_files = []

if data_source == "Use Local Workspace Files":
    local_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$") and "Report" not in f]
    local_csvs = [f for f in glob.glob("*.csv") if not os.path.basename(f).startswith("~$")]
    
    if local_xlsx:
        selected_xlsx = st.sidebar.selectbox("Select Excel File:", local_xlsx)
        excel_file = open(selected_xlsx, "rb")
        st.sidebar.caption(f"Excel: `{os.path.basename(selected_xlsx)}`")
    else:
        st.sidebar.warning("No .xlsx file found in local directory.")

    if local_csvs:
        selected_csvs = st.sidebar.multiselect("Select CSV Files to Include:", local_csvs, default=local_csvs)
        csv_files = [open(f, "rb") for f in selected_csvs]
        st.sidebar.caption(f"Selected {len(csv_files)} CSV files.")
    else:
        st.sidebar.warning("No .csv files found in local directory.")

else:
    uploaded_xlsx = st.sidebar.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
    uploaded_csvs = st.sidebar.file_uploader("Upload CSV File(s)", type=["csv"], accept_multiple_files=True)
    if uploaded_xlsx:
        excel_file = uploaded_xlsx
    if uploaded_csvs:
        csv_files = uploaded_csvs

# Main Header
st.markdown('<div class="main-header">SanSan Operation Records vs Excel Matcher</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Automated data verification tool to check match rates across all user identifiers</div>', unsafe_allow_html=True)

if not excel_file or not csv_files:
    st.info("👋 Please ensure both an Excel file and at least one CSV file are selected in the sidebar to begin.")
    st.stop()

# Read Excel Sheets
try:
    excel_wb = openpyxl.load_workbook(excel_file, data_only=True)
    sheet_names = excel_wb.sheetnames
except Exception as e:
    st.error(f"Error reading Excel workbook: {e}")
    st.stop()

# Layout: Configuration
with st.expander("⚙️ Excel & Mapping Configuration", expanded=True):
    col1, col2, col3 = st.columns([2, 1, 2])
    with col1:
        default_idx = sheet_names.index("Monthly Record") if "Monthly Record" in sheet_names else 0
        selected_sheet = st.selectbox("Excel Sheet to Check:", sheet_names, index=default_idx)
    with col2:
        header_row_idx = st.number_input("Header Row (1-indexed):", min_value=1, max_value=20, value=3)
    
    ws = excel_wb[selected_sheet]
    
    # Read headers from the chosen header row
    excel_headers = {}
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row_idx, column=col_idx).value
        if val is not None and str(val).strip() != '':
            excel_headers[str(val).strip()] = col_idx
            
    header_list = list(excel_headers.keys())
    
    with col3:
        # Default user identifier column in Excel
        default_id_col = "User Identifier" if "User Identifier" in header_list else header_list[0]
        selected_excel_id_col = st.selectbox("Excel User Identifier Column:", header_list, index=header_list.index(default_id_col))

    st.markdown("#### 🎯 Target Fields to Verify")
    
    # Let user select which fields to verify
    avail_fields = list(DEFAULT_FIELDS.keys())
    
    col_sel1, col_sel2 = st.columns([1, 4])
    with col_sel1:
        select_all = st.checkbox("Select All Fields", value=True)
    
    field_selection = {}
    field_cols = st.columns(4)
    for i, f_name in enumerate(avail_fields):
        with field_cols[i % 4]:
            is_checked = st.checkbox(f_name, value=select_all, key=f"chk_{f_name}")
            field_selection[f_name] = is_checked

# Parse CSV files
csv_records = {} # (uid, field_name) -> float
csv_users = {}   # uid -> name
csv_summary_info = []

for c_file in csv_files:
    c_name = getattr(c_file, 'name', 'CSV File')
    c_file.seek(0)
    try:
        df_curr = pd.read_csv(c_file)
        c_cols = list(df_curr.columns)
        
        # Detect if it's Item Select or Business Card
        is_item_select = "項目選択" in c_name or "Item" in c_name or ('temp' not in c_cols and 'matchRate' in c_cols)
        
        row_count = len(df_curr)
        csv_summary_info.append(f"• **{os.path.basename(c_name)}**: {row_count} rows ({'Item Selection' if is_item_select else 'Business Card items'})")

        if is_item_select and 'userIdentifier' in df_curr.columns and 'matchRate' in df_curr.columns:
            for _, r in df_curr.iterrows():
                uid = str(r['userIdentifier']).strip()
                if uid not in csv_users:
                    csv_users[uid] = ''
                rate = to_float_or_str(r['matchRate'])
                if rate is not None:
                    csv_records[(uid, 'Item Select')] = rate

        else:
            # Business card or general format with 'temp'
            for _, r in df_curr.iterrows():
                uid = str(r['userIdentifier']).strip() if 'userIdentifier' in r else ''
                name = str(r['name']).strip() if ('name' in r and pd.notna(r['name'])) else ''
                if uid and name:
                    csv_users[uid] = name

                temp_val = str(r['temp']).strip() if ('temp' in r and pd.notna(r['temp'])) else ''
                rate = to_float_or_str(r.get('matchRate'))
                
                # Check against DEFAULT_FIELDS
                for f_key, conf in DEFAULT_FIELDS.items():
                    if conf.get('csv_temp') and conf['csv_temp'] == temp_val and rate is not None:
                        csv_records[(uid, f_key)] = rate

    except Exception as e:
        st.error(f"Error parsing CSV file {c_name}: {e}")

st.markdown("##### 📄 Loaded CSV Files:")
for info in csv_summary_info:
    st.markdown(info)

# Parse Excel User Rows
excel_users = {}
excel_id_col_idx = excel_headers.get(selected_excel_id_col, 2)
op_name_col_idx = excel_headers.get("Operator name", 3)
for h_k, h_v in excel_headers.items():
    if "Operator" in h_k:
        op_name_col_idx = h_v
        break

for r in range(header_row_idx + 1, ws.max_row + 1):
    uid_cell = ws.cell(row=r, column=excel_id_col_idx).value
    if uid_cell is not None and str(uid_cell).strip() != '':
        uid = str(uid_cell).strip()
        op_name = ws.cell(row=r, column=op_name_col_idx).value or ''
        excel_users[uid] = (r, str(op_name).strip())

# Run Comparison
active_fields = [f for f, checked in field_selection.items() if checked]

if not active_fields:
    st.warning("Please select at least one field to check.")
    st.stop()

all_uids = sorted(list(set(list(csv_users.keys()) + list(excel_users.keys()))))
comparison_results = []
match_count = 0
mismatch_count = 0
missing_in_excel_count = 0
missing_in_csv_count = 0

for uid in all_uids:
    in_excel = uid in excel_users
    row_num, op_name = excel_users.get(uid, (None, csv_users.get(uid, '')))

    for f_name in active_fields:
        target_col_name = DEFAULT_FIELDS[f_name]['default_col']
        col_idx = excel_headers.get(target_col_name)

        c_val = csv_records.get((uid, f_name))
        e_val = None

        if in_excel and row_num and col_idx:
            raw_e = ws.cell(row=row_num, column=col_idx).value
            e_val = to_float_or_str(raw_e)

        if c_val is None and e_val is None:
            continue

        status = "MATCH"
        diff = 0
        if c_val is not None and e_val is not None:
            if c_val == e_val:
                status = "MATCH"
                match_count += 1
            else:
                status = "MISMATCH"
                mismatch_count += 1
                try:
                    diff = round(c_val - e_val, 2)
                except:
                    diff = "N/A"
        elif c_val is not None and e_val is None:
            status = "MISSING_IN_EXCEL"
            missing_in_excel_count += 1
        elif c_val is None and e_val is not None:
            status = "MISSING_IN_CSV"
            missing_in_csv_count += 1

        comparison_results.append({
            'User Identifier': uid,
            'Operator Name': op_name,
            'Excel Row': row_num if row_num else 'Not Found',
            'Field': f_name,
            'CSV matchRate': c_val if c_val is not None else '(empty)',
            'Excel Value': e_val if e_val is not None else '(empty)',
            'Difference (CSV - Excel)': diff,
            'Status': status
        })

df_results = pd.DataFrame(comparison_results)

# Display Metrics Cards
st.markdown("---")
mcol1, mcol2, mcol3, mcol4 = st.columns(4)

total_checked = len(df_results)
match_pct = (match_count / total_checked * 100) if total_checked > 0 else 0

with mcol1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">Total Values Checked</div>
        <div class="metric-val" style="color: #1E3A8A;">{total_checked:,}</div>
    </div>
    """, unsafe_allow_html=True)

with mcol2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">Matched Values</div>
        <div class="metric-val" style="color: #15803D;">{match_count:,} ({match_pct:.1f}%)</div>
    </div>
    """, unsafe_allow_html=True)

with mcol3:
    color = "#B91C1C" if mismatch_count > 0 else "#64748B"
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">Mismatches</div>
        <div class="metric-val" style="color: {color};">{mismatch_count:,}</div>
    </div>
    """, unsafe_allow_html=True)

with mcol4:
    color = "#B45309" if (missing_in_excel_count + missing_in_csv_count) > 0 else "#64748B"
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">Missing / Unmatched</div>
        <div class="metric-val" style="color: {color};">{missing_in_excel_count + missing_in_csv_count:,}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

if mismatch_count == 0 and missing_in_excel_count == 0 and missing_in_csv_count == 0:
    st.success("🎉 **PERFECT MATCH!** All values in CSV files match the Excel file exactly.")
else:
    st.error(f"⚠️ **Attention Needed:** Found {mismatch_count} value mismatch(es) and {missing_in_excel_count + missing_in_csv_count} missing value(s).")

# Interactive Filter & Search
st.markdown("### 📋 Comparison Details")
fcol1, fcol2, fcol3 = st.columns([2, 2, 3])

with fcol1:
    status_filter = st.selectbox("Filter by Status:", ["All", "Only Mismatches", "Only Missing", "Matches Only"])
with fcol2:
    field_filter = st.selectbox("Filter by Field:", ["All Fields"] + active_fields)
with fcol3:
    search_query = st.text_input("🔍 Search User Identifier or Operator Name:", "")

filtered_df = df_results.copy()

if status_filter == "Only Mismatches":
    filtered_df = filtered_df[filtered_df['Status'] == 'MISMATCH']
elif status_filter == "Only Missing":
    filtered_df = filtered_df[filtered_df['Status'].isin(['MISSING_IN_EXCEL', 'MISSING_IN_CSV'])]
elif status_filter == "Matches Only":
    filtered_df = filtered_df[filtered_df['Status'] == 'MATCH']

if field_filter != "All Fields":
    filtered_df = filtered_df[filtered_df['Field'] == field_filter]

if search_query:
    q = search_query.strip().lower()
    filtered_df = filtered_df[
        filtered_df['User Identifier'].str.lower().str.contains(q) |
        filtered_df['Operator Name'].str.lower().str.contains(q)
    ]

# Styling Function for Streamlit Dataframe
def highlight_status(row):
    status = row['Status']
    if status == 'MATCH':
        return ['background-color: #E2EFDA; color: #1E4620'] * len(row)
    elif status == 'MISMATCH':
        return ['background-color: #FCE4D6; color: #8A1F11; font-weight: bold'] * len(row)
    else:
        return ['background-color: #FFF2CC; color: #7F6000'] * len(row)

st.write(f"Showing **{len(filtered_df):,}** entries:")
if len(filtered_df) > 0:
    styled_df = filtered_df.style.apply(highlight_status, axis=1)
    st.dataframe(styled_df, use_container_width=True, height=450)
else:
    st.info("No records match the selected filter.")

# Export Section
st.markdown("---")
dcol1, dcol2 = st.columns([3, 1])
with dcol1:
    st.markdown("#### 📥 Download Complete Comparison Report")
    st.caption("Includes all checked fields, formulas, status labels, and color formatting.")
with dcol2:
    report_bytes = generate_excel_report(df_results)
    st.download_button(
        label="Download Excel Report (.xlsx)",
        data=report_bytes,
        file_name="Match_Comparison_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
