import streamlit as st
import pandas as pd
import io
import msoffcrypto

# --- 1. UI CONFIG ---
st.set_page_config(page_title="Validations Pro", page_icon="📂", layout="wide")

st.markdown("""
    <style>
    .header-container {
        text-align: center;
        padding: 1.5rem;
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        color: white;
        border-radius: 15px;
        margin-bottom: 2rem;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. HELPER FUNCTIONS ---

def load_file(file):
    if file is None: return None
    try:
        return pd.read_excel(file, engine='openpyxl')
    except:
        try:
            file.seek(0)
            return pd.read_csv(file)
        except: return None

def load_encrypted_xlsx(file, password):
    if file is None: return None
    if not password: 
        return load_file(file)
    try:
        decrypted_workbook = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_workbook)
        return pd.read_excel(decrypted_workbook, engine='openpyxl')
    except:
        return load_file(file)

def clean_string_id(df, col_name):
    if df is not None and col_name in df.columns:
        df[col_name] = df[col_name].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df[col_name] = df[col_name].replace('nan', None)
    return df

def find_and_rename_id(df, baseline_name='Employee ID'):
    if df is None: return None
    variations = [
        'Users Sys Id', 'Employee Code', 'User/Employee ID', 'Employee ID', 
        'Emp Code', 'Emp ID', 'Employee ID ', 'Employee ID  ', 'Employee No', 'ID', 'SNO'
    ]
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    for var in variations:
        if var in df.columns:
            df.rename(columns={var: baseline_name}, inplace=True)
            return df
    return df

def identity_check(val1, val2, is_date=False):
    if is_date:
        if pd.isna(val1) or pd.isna(val2) or str(val1).strip() == "" or str(val2).strip().lower() == "nan":
            return "No"
        try:
            d1 = pd.to_datetime(str(val1).strip(), errors='coerce').strftime('%d-%m-%Y')
            d2 = pd.to_datetime(str(val2).strip(), errors='coerce').strftime('%d-%m-%Y')
            if d1 == "NaT" or d2 == "NaT" or pd.isna(d1) or pd.isna(d2):
                return "No"
            return "Yes" if d1 == d2 else f"No ({d1})"
        except: 
            return "No"
            
    v1 = str(val1).strip().lower() if not pd.isna(val1) else ""
    v2 = str(val2).strip().lower() if not pd.isna(val2) else ""
    if v1 == "" or v1 == "none" or v2 == "" or v2 == "none":
        return "No"
    return "Yes" if v1 == v2 else "No"

# --- 3. MAIN APP ---

st.markdown('<div class="header-container"><h1>📂 Checkpoint Portal </h1></div>', unsafe_allow_html=True)

observations_registry = {
    'Missing Employee Codes': [],
    'DOJ Mismatch': [],
    'DOL Mismatch': [],
    'Designation Mismatch': [],
    'State Mismatch': []
}

with st.sidebar:
    st.subheader("🔑 Core Master Record (Required)")
    sales_file = st.file_uploader("1. Sales Register (Master Payout)", type=['xlsx'])
    sales_pw = st.text_input("Sales Register Password (If protected)", type="password")
    
    st.markdown("---")
    st.subheader("📁 Reference Files (Optional)")
    
    input_file = st.file_uploader("2. Input Sheet", type=['xlsx', 'csv'])
    hc_file = st.file_uploader("3. HC Report", type=['xlsx', 'csv'])
    
    run_audit = st.button("🚀 RUN MATRIX AUDIT", use_container_width=True)

if run_audit:
    if sales_file is None:
        st.error("❌ The primary 'Sales Register' file must be uploaded to establish the audit base.")
    else:
        with st.status("🔍 Processing Active Reference Audits...") as status:
            
            df_sales = load_encrypted_xlsx(sales_file, sales_pw) if sales_pw else load_file(sales_file)
            if df_sales is None:
                st.error("Decryption or loading failed for the Sales Register. Verify your file and master password.")
                st.stop()
                
            df_sales = clean_string_id(df_sales, 'Emp Code')
            df_sales = df_sales.dropna(subset=['Emp Code'])
            df_sales = df_sales[df_sales['Emp Code'].astype(str).str.strip() != ""]
            
            audit_df = df_sales.copy()

            df_input = load_file(input_file) if input_file else None
            df_hc = load_file(hc_file) if hc_file else None

            df_input = find_and_rename_id(df_input, 'Employee ID')
            df_hc = find_and_rename_id(df_hc, 'Employee ID')

            if df_input is not None: df_input = clean_string_id(df_input, 'Employee ID')
            if df_hc is not None: df_hc = clean_string_id(df_hc, 'Employee ID')

            # --- TARGETED CHECK: Cross-check if all input sheet employee records are present in register ---
            if df_input is not None and 'Employee ID' in df_input.columns:
                sales_ids_set = set(audit_df['Emp Code'].dropna().astype(str).str.strip())
                for idx, input_row in df_input.iterrows():
                    inp_id = str(input_row['Employee ID']).strip()
                    if inp_id and inp_id != "None" and inp_id not in sales_ids_set:
                        observations_registry['Missing Employee Codes'].append({
                            "Employee ID": inp_id,
                            "Value as per Register": "Missing from Payout Register",
                            "Actual Lookup": "Present in Input Sheet"
                        })

            # --- PART 1: BASIC DETAILS PIPELINES ---

            # Pipeline A: Input Sheet Lookups
            if df_input is not None and 'Employee ID' in df_input.columns:
                input_lookup_df = df_input[['Employee ID']].copy()
                input_lookup_df.rename(columns={'Employee ID': 'Lookup_Emp_Code'}, inplace=True)
                
                audit_df = audit_df.merge(input_lookup_df, left_on='Emp Code', right_on='Lookup_Emp_Code', how='left')
                audit_df['Check_Emp_Code'] = audit_df.apply(lambda x: identity_check(x['Emp Code'], x['Lookup_Emp_Code']), axis=1)
            else:
                audit_df['Lookup_Emp_Code'] = None
                audit_df['Check_Emp_Code'] = "No"

            # Pipeline B: Headcount Report Lookups
            if df_hc is not None and 'Employee ID' in df_hc.columns:
                hc_cols = ['Employee ID', 'Employment Details Group Date of Joining', 'Employment Details Actual Exit Date', 'Position Title', 'State']
                hc_existing = [c for c in hc_cols if c in df_hc.columns]
                audit_df = audit_df.merge(df_hc[hc_existing], left_on='Emp Code', right_on='Employee ID', how='left', suffixes=('', '_hc'))
                
                audit_df['Check_DOJ'] = audit_df.apply(lambda x: identity_check(x.get('Date of Joining'), x.get('Employment Details Group Date of Joining'), True), axis=1)
                audit_df['Check_DOL'] = audit_df.apply(lambda x: identity_check(x.get('DOL'), x.get('Employment Details Actual Exit Date'), True), axis=1)
                audit_df['Check_Designation'] = audit_df.apply(lambda x: identity_check(x.get('DESIGNATION'), x.get('Position Title')), axis=1)
                audit_df['Check_State'] = audit_df.apply(lambda x: identity_check(x.get('STATE'), x.get('State')), axis=1)
            else:
                audit_df['Employment Details Group Date of Joining'] = None
                audit_df['Check_DOJ'] = None
                audit_df['Employment Details Actual Exit Date'] = None
                audit_df['Check_DOL'] = None
                audit_df['Position Title'] = None
                audit_df['Check_Designation'] = None
                audit_df['State'] = None
                audit_df['Check_State'] = None

            # --- GENERATING THE OBSERVATIONS LOGIC MATRIX ---
            for idx, row in audit_df.iterrows():
                emp_id = str(row['Emp Code']).strip()
                
                if row['Check_DOJ'] == "No" or (isinstance(row['Check_DOJ'], str) and row['Check_DOJ'].startswith("No")):
                    observations_registry['DOJ Mismatch'].append({
                        "Employee ID": emp_id, "Value as per Register": str(row.get('Date of Joining')), "Actual Lookup": str(row.get('Employment Details Group Date of Joining'))
                    })
                
                text_checks = {
                    'DOL Mismatch': ('DOL', 'Employment Details Actual Exit Date', 'Check_DOL'),
                    'Designation Mismatch': ('DESIGNATION', 'Position Title', 'Check_Designation'),
                    'State Mismatch': ('STATE', 'State', 'Check_State')
                }
                for heading, (reg_c, lkp_c, chk_c) in text_checks.items():
                    if row[chk_c] == "No" or (isinstance(row[chk_c], str) and row[chk_c].startswith("No")):
                        observations_registry[heading].append({
                            "Employee ID": emp_id, "Value as per Register": str(row.get(reg_c)), "Actual Lookup": str(row.get(lkp_c))
                        })

            # --- EXPLICIT FIELD ORDERING & STRUCTURE ASSEMBLY ---
            ordered_headers = [
                'SNO', 'Emp Code', 'Lookup_Emp_Code', 'Check_Emp_Code', 'Emp Name', 'DOB', 'Date of Joining',
                'Employment Details Group Date of Joining', 'Check_DOJ', 'DOL', 'Employment Details Actual Exit Date', 'Check_DOL',
                'DEPARTMENT', 'DESIGNATION', 'Position Title', 'Check_Designation', 'Cost Centre', 'Grade', 'GENDER', 'PAN',
                'LOCATION CODE', 'LOCATION', 'STATE', 'State', 'Check_State'
            ]

            final_headers = [h for h in ordered_headers if h in audit_df.columns]
            result_df = audit_df.reindex(columns=final_headers)

            final_renaming = {
                'Lookup_Emp_Code': 'Lookup', 'Check_Emp_Code': 'Check',
                'Employment Details Group Date of Joining': 'Lookup ', 'Check_DOJ': 'Check ',
                'Employment Details Actual Exit Date': 'Lookup  ', 'Check_DOL': 'Check  ',
                'Position Title': 'Lookup   ', 'Check_Designation': 'Check   ',
                'State': 'Lookup    ', 'Check_State': 'Check    '
            }
            
            display_df = result_df.rename(columns=final_renaming)
            status.update(label="✅ Run Completed Successfully!", state="complete")

    st.dataframe(display_df, use_container_width=True)

    st.markdown("## 📋 Real-Time Audit Observations Sheet Log")
    active_observations_found = False
    for heading, rows in observations_registry.items():
        if rows:
            active_observations_found = True
            st.markdown(f"### 📌 {heading}")
            section_df = pd.DataFrame(rows)[["Employee ID", "Value as per Register", "Actual Lookup"]]
            st.dataframe(section_df, use_container_width=True)
            
    if not active_observations_found:
        st.success("🎉 Perfect Run! No audit exceptions triggered on this data subset.")

    # --- 5. EXCEL EXPORT ENGINE ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        display_df.to_excel(writer, index=False, sheet_name='FnF_Master_Audit')
        workbook  = writer.book
        worksheet = writer.sheets['FnF_Master_Audit']
        
        normal_border_fmt = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0F172A', 'font_color': 'white', 'border': 1})

        for col_num, col_name in enumerate(display_df.columns):
            if 'lookup' in col_name.lower() or 'check' in col_name.lower():
                worksheet.set_column(col_num, col_num, 16, normal_border_fmt)
            else:
                worksheet.set_column(col_num, col_num, 18)
            worksheet.write(0, col_num, col_name, header_fmt)

        obs_rows = []
        for heading, rows in observations_registry.items():
            if rows:
                obs_rows.append({"Employee ID": f"--- {heading.upper()} ---", "Value as per Register": "", "Actual Lookup": ""})
                for r in rows:
                    obs_rows.append(r)
                    
        obs_df = pd.DataFrame(obs_rows) if obs_rows else pd.DataFrame(columns=["Employee ID", "Value as per Register", "Actual Lookup"])
        obs_df.to_excel(writer, index=False, sheet_name='Audit_Observations')
        obs_worksheet = writer.sheets['Audit_Observations']
        
        obs_worksheet.set_column(0, 0, 35) 
        obs_worksheet.set_column(1, 1, 30) 
        obs_worksheet.set_column(2, 2, 45) 
        
        section_title_fmt = workbook.add_format({'bold': True, 'bg_color': '#E2E8F0', 'font_color': '#0F172A', 'border': 1})
        
        for col_num, col_name in enumerate(obs_df.columns):
            obs_worksheet.write(0, col_num, col_name, header_fmt)
            
        for row_num, row_data in enumerate(obs_rows, start=1):
            if str(row_data["Employee ID"]).startswith("--- "):
                obs_worksheet.write(row_num, 0, row_data["Employee ID"], section_title_fmt)
                obs_worksheet.write(row_num, 1, "", section_title_fmt)
                obs_worksheet.write(row_num, 2, "", section_title_fmt)

    st.download_button("📥 DOWNLOAD AUDIT REPORT & OBSERVATIONS WORKBOOK", output.getvalue(), "FnF_Audit_Package_Final.xlsx", use_container_width=True)
