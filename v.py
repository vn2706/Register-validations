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

def clean_num(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return 0.0
        s = str(val).replace(',', '').replace('(', '-').replace(')', '').strip()
        return round(float(s), 3)
    except: return 0.0

def clean_string_id(df, col_name):
    if df is not None and col_name in df.columns:
        df[col_name] = df[col_name].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return df

def find_and_rename_id(df, baseline_name='Employee ID'):
    """Dynamically locates common identifier columns and maps them to a single uniform primary key label"""
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

def find_and_rename_sales_rates(df):
    """Normalizes Sales Register rate columns to ensure exact matches regardless of system case or characters"""
    if df is None: return None
    mapping_rules = {
        'BASIC SALARY RATE': ['BASIC SALARY RATE', 'BASIC SALARY', 'BASIC RATE', 'BASIC_SALARY_RATE', 'Basic Salary Rate'],
        'MOBILE ALLOWANCE RATE': ['MOBILE ALLOWANCE RATE', 'MOBILE ALLOWANCE', 'MOBILE RATE', 'MOBILE_ALLOWANCE_RATE', 'Mobile Allowance Rate'],
        'CONSISTENCY ALLOWANCE RATE': ['CONSISTENCY ALLOWANCE RATE', 'CONSISTENCY ALLOWANCE', 'CONSISTENCY RATE', 'CONSISTENCY_ALLOWANCE_RATE', 'Consistency Allowance Rate'],
        'SALES LINKED COMMISSION RATE': ['SALES LINKED COMMISSION RATE', 'SALES LINKED COMMISSION', 'SALES COMMISSION', 'SALES_LINKED_COMMISSION_RATE', 'Sales Linked Commission Rate'],
        'HOUSE RENT ALLOWANCE RATE': ['HOUSE RENT ALLOWANCE RATE', 'HOUSE RENT ALLOWANCE', 'HRA RATE', 'HOUSE_RENT_ALLOWANCE_RATE', 'House Rent Allowance Rate'],
        'STATUTORY BONUS RATE': ['STATUTORY BONUS RATE', 'STATUTORY BONUS', 'BONUS RATE', 'STATUTORY_BONUS_RATE', 'Statutory Bonus Rate']
    }
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    for target, variants in mapping_rules.items():
        for var in variants:
            if var in df.columns:
                df.rename(columns={var: target}, inplace=True)
                break
    return df

def variance_check(val1, val2):
    n1 = clean_num(val1)
    n2 = clean_num(val2)
    diff = round(n1 - n2, 3)
    return 0 if abs(diff) < 0.001 else diff

def identity_check(val1, val2, is_date=False):
    if pd.isna(val1) or pd.isna(val2) or str(val1).strip() == "" or str(val2).strip() == "": 
        return "No"
    if is_date:
        try:
            d1 = pd.to_datetime(val1).strftime('%d-%m-%Y')
            d2 = pd.to_datetime(val2).strftime('%d-%m-%Y')
            return "Yes" if d1 == d2 else f"No ({d1})"
        except: return "No"
    v1 = str(val1).strip().lower()
    v2 = str(val2).strip().lower()
    return "Yes" if v1 == v2 else "No"

# --- 3. MAIN APP ---

st.markdown('<div class="header-container"><h1>📂 Checkpoint Portal </h1></div>', unsafe_allow_html=True)

missing_reports_log = {}

# Registry to log rows group by group cleanly for dynamic table compilation
observations_registry = {
    'Missing Columns': [],
    'Missing Employee Codes': [],
    'DOJ Mismatch': [],
    'DOL Mismatch': [],
    'Designation Mismatch': [],
    'State Mismatch': [],
    'Bank Account Mismatch': [],
    'IFSC Mismatch': [],
    'Inventory Recovery Variance': [],
    'Notice Period Variance': [],
    'Facility Recovery Variance': [],
    'Leave Encashment Variance': [],
    'All Allowance Rates Mismatch': []
}

with st.sidebar:
    st.subheader("🔑 Core Master Record (Required)")
    sales_file = st.file_uploader("1. Sales Register (Master Payout)", type=['xlsx'])
    sales_pw = st.text_input("Sales Register Password (If protected)", type="password")
    
    st.markdown("---")
    st.subheader("📁 Reference Files (Optional)")
    
    input_file = st.file_uploader("2. Input Sheet", type=['xlsx', 'csv'])
    hc_file = st.file_uploader("3. HC Report", type=['xlsx', 'csv'])
    bank_file = st.file_uploader("4. Bank Report", type=['xlsx', 'csv'])
    
    ctc_file = st.file_uploader("5. CTC Report", type=['xlsx', 'csv'])
    ctc_pw = st.text_input("CTC Report Password (If protected)", type="password")
    
    run_audit = st.button("🚀 RUN MATRIX AUDIT", use_container_width=True)

if run_audit:
    if sales_file is None:
        st.error("❌ The primary 'Sales Register' file must be uploaded to establish the audit base.")
    else:
        with st.status("🔍 Processing Active Reference Audits...") as status:
            
            # 1. Ingest Master Sales Record & Map headers safely
            df_sales = load_encrypted_xlsx(sales_file, sales_pw) if sales_pw else load_file(sales_file)
            if df_sales is None:
                st.error("Decryption or loading failed for the Sales Register. Verify your file and master password.")
                st.stop()
                
            df_sales = find_and_rename_sales_rates(df_sales)
            df_sales = clean_string_id(df_sales, 'Emp Code')
            
            # Drop trailing rows completely if they don't contain an employee code
            df_sales = df_sales.dropna(subset=['Emp Code'])
            df_sales = df_sales[df_sales['Emp Code'].astype(str).str.strip() != ""]
            
            audit_df = df_sales.copy()

            # GLOBAL DEFINITION: Target Arrear headers
            arrear_columns = [
                'Basic Salary(Arrear)', 'Mobile Allowance(Arrear)', 'Consistency Allowance(Arrear)',
                'Sales Linked Commission(Arrear)', 'House Rent Allowance(Arrear)', 'Advance Statutory Bonus(Arrear)'
            ]

            # 2. Ingest Optional Supporting Files safely
            df_input = load_file(input_file) if input_file else None
            df_hc = load_file(hc_file) if hc_file else None
            df_bank = load_file(bank_file) if bank_file else None
            df_ctc = load_encrypted_xlsx(ctc_file, ctc_pw) if ctc_file else None

            # Standardize identifier columns across ALL dataframes dynamically to prevent KeyErrors
            df_input = find_and_rename_id(df_input, 'Employee ID')
            df_hc = find_and_rename_id(df_hc, 'Employee ID')
            df_bank = find_and_rename_id(df_bank, 'Employee ID')
            df_ctc = find_and_rename_id(df_ctc, 'Employee ID')

            # Clean and sanitize row strings to avoid center-alignment spacing issues
            if df_input is not None: df_input = clean_string_id(df_input, 'Employee ID')
            if df_hc is not None: df_hc = clean_string_id(df_hc, 'Employee ID')
            if df_bank is not None: df_bank = clean_string_id(df_bank, 'Employee ID')
            if df_ctc is not None: df_ctc = clean_string_id(df_ctc, 'Employee ID')

            # Fault-Tolerance Check: If rate columns as a whole are missing from Sales Register, declare placeholders
            required_register_rates = ['BASIC SALARY RATE', 'MOBILE ALLOWANCE RATE', 'CONSISTENCY ALLOWANCE RATE', 'SALES LINKED COMMISSION RATE', 'HOUSE RENT ALLOWANCE RATE', 'STATUTORY BONUS RATE']
            for rate_col in required_register_rates:
                if rate_col not in audit_df.columns:
                    audit_df[rate_col] = 0.0
                    observations_registry['Missing Columns'].append({
                        "Employee ID": "N/A", 
                        "Value as per Register": "Missing Column", 
                        "Actual Lookup": f"Column '{rate_col}' absent from Register"
                    })

            # --- EXECUTE SEGMENTED PIPELINES BASED ON UPLOADS ---

            # Pipeline A: Input Sheet Lookups
            if df_input is not None and 'Employee ID' in df_input.columns:
                target_cols = ['Employee ID', 'Leave Encashment', 'Inventory Recovery', 'NP recovery', 'Facility Recovery']
                existing_cols = [c for c in target_cols if c in df_input.columns]
                audit_df = audit_df.merge(df_input[existing_cols], left_on='Emp Code', right_on='Employee ID', how='left', suffixes=('', '_input'))
                audit_df.rename(columns={'Employee ID': 'Lookup_Emp_Code'}, inplace=True)
                audit_df['Check_Emp_Code'] = audit_df.apply(lambda x: identity_check(x['Lookup_Emp_Code'], x['Emp Code']), axis=1)
                
                inv_rec_col = 'Inventory Recovery_input' if 'Inventory Recovery_input' in audit_df.columns else 'Inventory Recovery'
                audit_df['Check_Inventory_Recovery'] = audit_df.apply(lambda x: variance_check(x.get('Inventory Recovery'), x.get(inv_rec_col)), axis=1)
            else:
                audit_df['Lookup_Emp_Code'] = None
                audit_df['Check_Emp_Code'] = None
                audit_df['Leave Encashment'] = None
                audit_df['Inventory Recovery_input'] = None
                audit_df['Check_Inventory_Recovery'] = None
                audit_df['NP recovery'] = None
                audit_df['Facility Recovery_input'] = None
                missing_reports_log["Input Sheet"] = ["Lookup", "Check (Emp Code Match)", "LE Days", "Inventory Recovery Lookup/Check", "NP_Lookup", "Facility Recovery Lookup/Check"]

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
                missing_reports_log["HC Report"] = ["DOJ Lookup/Check", "DOL Lookup/Check", "Designation Lookup/Check", "State Lookup/Check"]

            # Pipeline C: Bank Report Lookups
            if df_bank is not None and 'Employee ID' in df_bank.columns:
                bank_acc_col = 'accountNumber' if 'accountNumber' in df_bank.columns else ('Account Number' if 'Account Number' in df_bank.columns else None)
                bank_ifsc_col = 'IFSC Code' if 'IFSC Code' in df_bank.columns else ('IFSC' if 'IFSC' in df_bank.columns else None)
                
                slice_cols = ['Employee ID']
                if bank_acc_col: slice_cols.append(bank_acc_col)
                if bank_ifsc_col: slice_cols.append(bank_ifsc_col)
                
                audit_df = audit_df.merge(df_bank[slice_cols], left_on='Emp Code', right_on='Employee ID', how='left', suffixes=('', '_bank'))
                
                actual_acc = audit_df[bank_acc_col] if bank_acc_col else None
                actual_ifsc = audit_df[bank_ifsc_col] if bank_ifsc_col else None
                
                audit_df['Check_Bank_Acc'] = audit_df.apply(lambda x: identity_check(x.get('Bank_Acc_No'), actual_acc.loc[x.name] if bank_acc_col else None), axis=1)
                audit_df['Check_IFSC'] = audit_df.apply(lambda x: identity_check(x.get('IFSC'), actual_ifsc.loc[x.name] if bank_ifsc_col else None), axis=1)
                
                audit_df['accountNumber'] = actual_acc
                audit_df['IFSC Code'] = actual_ifsc
            else:
                audit_df['accountNumber'] = None
                audit_df['Check_Bank_Acc'] = None
                audit_df['IFSC Code'] = None
                audit_df['Check_IFSC'] = None
                missing_reports_log["Bank Report"] = ["Bank Account Lookup/Check", "IFSC Lookup/Check"]

            # Pipeline D: CTC Report Lookups
            if df_ctc is not None and 'Employee ID' in df_ctc.columns:
                ctc_cols = ['Employee ID', 'Basic Pay Sales Master', 'Mobile Allow Sales Master', 'Consistency Allowance', 'Sales Linked Comm. Master', 'HRA Sales Master', 'Adv Stt Bonus SalesMaster', 'IBP']
                ctc_existing = [c for c in ctc_cols if c in df_ctc.columns]
                audit_df = audit_df.merge(df_ctc[ctc_existing], left_on='Emp Code', right_on='Employee ID', how='left', suffixes=('', '_ctc'))
                
                # Monthly rate lookups mapped side-by-side with /12 adjustments
                audit_df['Final Basic pay'] = audit_df['Basic Pay Sales Master'].apply(lambda x: round(clean_num(x) / 12, 3))
                audit_df['Final Mobile allowance'] = audit_df['Mobile Allow Sales Master'].apply(lambda x: round(clean_num(x) / 12, 3))
                audit_df['Final Const. Bonus'] = audit_df['Consistency Allowance'].apply(lambda x: round(clean_num(x) / 12, 3))
                audit_df['Final Sales linked'] = audit_df['Sales Linked Comm. Master'].apply(lambda x: round(clean_num(x) / 12, 3))
                audit_df['Final HRA'] = audit_df['HRA Sales Master'].apply(lambda x: round(clean_num(x) / 12, 3))
                audit_df['Final Adv. stat bonus'] = audit_df['Adv Stt Bonus SalesMaster'].apply(lambda x: round(clean_num(x) / 12, 3))
                
                # CRITICAL CALCULATION SPECIFICATION: IBP monthly rate factor strictly requires /12 AND /2 logic parameters
                ibp_col_key = 'IBP' if 'IBP' in audit_df.columns else ('IBP_ctc' if 'IBP_ctc' in audit_df.columns else None)
                if ibp_col_key:
                    audit_df['Calculated_IBP'] = audit_df[ibp_col_key].apply(lambda x: round((clean_num(x) / 12) / 2, 3))
                else:
                    audit_df['Calculated_IBP'] = None

                audit_df['Check_Basic_Rate'] = audit_df.apply(lambda x: variance_check(x.get('BASIC SALARY RATE'), x.get('Final Basic pay')), axis=1)
                audit_df['Check_Mobile_Rate'] = audit_df.apply(lambda x: variance_check(x.get('MOBILE ALLOWANCE RATE'), x.get('Final Mobile allowance')), axis=1)
                audit_df['Check_Consistency_Rate'] = audit_df.apply(lambda x: variance_check(x.get('CONSISTENCY ALLOWANCE RATE'), x.get('Final Const. Bonus')), axis=1)
                audit_df['Check_Sales_Rate'] = audit_df.apply(lambda x: variance_check(x.get('SALES LINKED COMMISSION RATE'), x.get('Final Sales linked')), axis=1)
                audit_df['Check_HRA_Rate'] = audit_df.apply(lambda x: variance_check(x.get('HOUSE RENT ALLOWANCE RATE'), x.get('Final HRA')), axis=1)
                audit_df['Check_Statutory_Rate'] = audit_df.apply(lambda x: variance_check(x.get('STATUTORY BONUS RATE'), x.get('Final Adv. stat bonus')), axis=1)
            else:
                rate_cols = ['Final Basic pay', 'Check_Basic_Rate', 'Final Mobile allowance', 'Check_Mobile_Rate', 
                             'Final Const. Bonus', 'Check_Consistency_Rate', 'Final Sales linked', 'Check_Sales_Rate', 
                             'Final HRA', 'Check_HRA_Rate', 'Final Adv. stat bonus', 'Check_Statutory_Rate', 'Basic Pay Sales Master', 'Calculated_IBP']
                for c in rate_cols: audit_df[c] = None
                missing_reports_log["CTC Report"] = ["All Salary Rate Lookup fields", "All Component Verification checks", "Basic Pay Sales Master", "Calculated IBP field"]

            # Pipeline F: Dependent Formulas & Cross-Checks
            le_input_col = 'Leave Encashment_input' if 'Leave Encashment_input' in audit_df.columns else 'Leave Encashment'
            np_input_col = 'NP recovery_input' if 'NP recovery_input' in audit_df.columns else 'NP recovery'
            
            if df_ctc is not None and df_input is not None:
                # Leave Encashment and NP Recovery mathematical architectures map sequentially to the newly divided Calculated_IBP
                audit_df['Calc_LE_Payout'] = audit_df.apply(lambda x: round((max(clean_num(x.get('Basic Pay Sales Master')), clean_num(x.get('Calculated_IBP'))) / 30) * clean_num(x.get(le_input_col)), 3), axis=1)
                audit_df['Check_LE_Days_Variance'] = audit_df.apply(lambda x: variance_check(x.get('Calc_LE_Payout'), x.get('LE Days')), axis=1)
                
                audit_df['Calc_NP_Recovery'] = audit_df.apply(lambda x: round((max(clean_num(x.get('Basic Pay Sales Master')), clean_num(x.get('Calculated_IBP'))) / 30) * clean_num(x.get(np_input_col)), 3), axis=1)
                audit_df['Check_NP_Variance'] = audit_df.apply(lambda x: variance_check(x.get('Notice Recovery'), x.get('Calc_NP_Recovery')), axis=1)
            else:
                audit_df['Calc_LE_Payout'] = None
                audit_df['Check_LE_Days_Variance'] = None
                audit_df['Calc_NP_Recovery'] = None
                audit_df['Check_NP_Variance'] = None
                missing_reports_log["Cross-Calculations (LE / NP)"] = ["Calc Payouts", "Calculated Notice Period Recovery", "Variance Checks"]

            # Independent Static Lookups/Placeholders safely tracking Input Sheet suffixes
            audit_df['LE Days'] = audit_df[le_input_col] if df_input is not None else None
            audit_df['Lookup_PTax'] = None
            audit_df['Check_PTax'] = None
            
            fac_rec_col = 'Facility Recovery_input' if 'Facility Recovery_input' in audit_df.columns else 'Facility Recovery'
            audit_df['Check_Facility_Recovery'] = audit_df.apply(lambda x: variance_check(x.get('Facility Recovery'), x.get(fac_rec_col)), axis=1) if df_input is not None else None

            # --- GENERATING THE OBSERVATIONS LOGIC MATRIX (ONE CONSOLIDATED TABLE PER SECTION) ---
            for idx, row in audit_df.iterrows():
                emp_id = str(row['Emp Code']).strip()
                
                # Exception 1: Unmapped/Missing Profile
                if pd.isna(row['Lookup_Emp_Code']) or str(row['Lookup_Emp_Code']).strip() == "" or row['Lookup_Emp_Code'] is None:
                    observations_registry['Missing Employee Codes'].append({
                        "Employee ID": emp_id, "Value as per Register": emp_id, "Actual Lookup": "Missing reference entry"
                    })
                
                # Exception 2: Date of Joining Mismatches
                if row['Check_DOJ'] == "No" or (isinstance(row['Check_DOJ'], str) and row['Check_DOJ'].startswith("No")):
                    observations_registry['DOJ Mismatch'].append({
                        "Employee ID": emp_id, "Value as per Register": str(row.get('Date of Joining')), "Actual Lookup": str(row.get('Employment Details Group Date of Joining'))
                    })
                
                # Exception 3: Standard Metadata String Mismatches (Yes / No Flags)
                text_checks = {
                    'DOL Mismatch': ('DOL', 'Employment Details Actual Exit Date', 'Check_DOL'),
                    'Designation Mismatch': ('DESIGNATION', 'Position Title', 'Check_Designation'),
                    'State Mismatch': ('STATE', 'State', 'Check_State'),
                    'Bank Account Mismatch': ('Bank_Acc_No', 'accountNumber', 'Check_Bank_Acc'),
                    'IFSC Mismatch': ('IFSC', 'IFSC Code', 'Check_IFSC')
                }
                for heading, (reg_c, lkp_c, chk_c) in text_checks.items():
                    if row[chk_c] == "No" or (isinstance(row[chk_c], str) and row[chk_c].startswith("No")):
                        observations_registry[heading].append({
                            "Employee ID": emp_id, "Value as per Register": str(row.get(reg_c)), "Actual Lookup": str(row.get(lkp_c))
                        })

                # Exception 4: Numerical Variances (Threshold filter set to >= 1.0)
                numeric_checks = {
                    'Inventory Recovery Variance': ('Inventory Recovery', inv_rec_col, 'Check_Inventory_Recovery'),
                    'Notice Period Variance': ('Notice Recovery', 'Calc_NP_Recovery', 'Check_NP_Variance'),
                    'Facility Recovery Variance': ('Facility Recovery', fac_rec_col, 'Check_Facility_Recovery'),
                    'Leave Encashment Variance': ('Leave Encashment', 'Calc_LE_Payout', 'Check_LE_Days_Variance')
                }
                for heading, (reg_c, lkp_c, chk_c) in numeric_checks.items():
                    v_val = abs(clean_num(row.get(chk_c)))
                    if v_val >= 1.0:
                        observations_registry[heading].append({
                            "Employee ID": emp_id, "Value as per Register": str(row.get(reg_c)), "Actual Lookup": str(row.get(lkp_c))
                        })

                # Exception 5: Allowance Rates Mismatch Loop (Triggers only if all six rate checks >= 1.0 simultaneously)
                rate_checks_list = ['Check_Basic_Rate', 'Check_Mobile_Rate', 'Check_Consistency_Rate', 'Check_Sales_Rate', 'Check_HRA_Rate', 'Check_Statutory_Rate']
                if all(abs(clean_num(row.get(c))) >= 1.0 for c in rate_checks_list):
                    observations_registry['All Allowance Rates Mismatch'].append({
                        "Employee ID": emp_id, 
                        "Value as per Register": f"Basic Rate: {row.get('BASIC SALARY RATE')}", 
                        "Actual Lookup": f"Reference Monthly Basic: {row.get('Final Basic pay')}"
                    })

            # --- ARREARS CONDITIONAL DELETION ENGINE ---
            for col in arrear_columns:
                if col in audit_df.columns:
                    if audit_df[col].apply(clean_num).eq(0).all():
                        audit_df.drop(columns=[col], inplace=True)

            # --- EXPLICIT FIELD ORDERING & STRUCTURE ASSEMBLY ---
            ordered_headers = [
                'SNO', 'Emp Code', 'Lookup_Emp_Code', 'Check_Emp_Code', 'Emp Name', 'DOB', 'Date of Joining',
                'Employment Details Group Date of Joining', 'Check_DOJ', 'DOL', 'Employment Details Actual Exit Date', 'Check_DOL',
                'DEPARTMENT', 'DESIGNATION', 'Position Title', 'Check_Designation', 'Cost Centre', 'Grade', 'GENDER', 'PAN',
                'LOCATION CODE', 'LOCATION', 'STATE', 'State', 'Check_State', 'Entity', 'Bank_Name', 'Bank_Acc_No',
                'accountNumber', 'Check_Bank_Acc', 'IFSC', 'IFSC Code', 'Check_IFSC', 'STATUS_DESC', 'BASIC SALARY RATE',
                'Final Basic pay', 'Check_Basic_Rate', 'MOBILE ALLOWANCE RATE', 'Final Mobile allowance', 'Check_Mobile_Rate',
                'CONSISTENCY ALLOWANCE RATE', 'Final Const. Bonus', 'Check_Consistency_Rate', 'SALES LINKED COMMISSION RATE',
                'Final Sales linked', 'Check_Sales_Rate', 'HOUSE RENT ALLOWANCE RATE', 'Final HRA', 'Check_HRA_Rate',
                'STATUTORY BONUS RATE', 'Final Adv. stat bonus', 'Check_Statutory_Rate', 'Rated_Gross', 'Leave Encashment',
                'LE Days', 'Basic Pay Sales Master', 'Calculated_IBP', 'Calc_LE_Payout', 'Check_LE_Days_Variance',
                'Gross_Salary', 'GROSSARREAR', 'NETGROSSTOTAL', 'PF', 'VPF', 'ESI', 'Professional_Tax', 'Lookup_PTax', 'Check_PTax',
                'Income_Tax', 'Hold Release Salary', 'Neg Salary Brought Forward', 'Inventory Recovery', 'Inventory Recovery_input',
                'Check_Inventory_Recovery', 'Emp LWF', 'Notice Recovery', 'NP recovery', 'Calc_NP_Recovery', 'Check_NP_Variance',
                'Facility Recovery', 'Facility Recovery_input', 'Check_Facility_Recovery', 'Emp LWF(Arrear)', 'Gross_Deduction', 'Net_Salary'
            ]

            final_headers = [h for h in ordered_headers if h in audit_df.columns]
            result_df = audit_df.reindex(columns=final_headers)

            # Map into final layouts
            final_renaming = {
                'Lookup_Emp_Code': 'Lookup', 'Check_Emp_Code': 'Check',
                'Employment Details Group Date of Joining': 'Lookup ', 'Check_DOJ': 'Check ',
                'Employment Details Actual Exit Date': 'Lookup  ', 'Check_DOL': 'Check  ',
                'Position Title': 'Lookup   ', 'Check_Designation': 'Check   ',
                'State': 'Lookup    ', 'Check_State': 'Check    ',
                'accountNumber': 'Lookup     ', 'Check_Bank_Acc': 'Check     ',
                'IFSC Code': 'Lookup      ', 'Check_IFSC': 'Check      ',
                'Final Basic pay': 'Lookup       ', 'Check_Basic_Rate': 'Check       ',
                'Final Mobile allowance': 'Lookup ........', 'Check_Mobile_Rate': 'Check        ',
                'Final Const. Bonus': 'Lookup         ', 'Check_Consistency_Rate': 'Check         ',
                'Final Sales linked': 'Lookup          ', 'Check_Sales_Rate': 'Check          ',
                'Final HRA': 'Lookup           ', 'Check_HRA_Rate': 'Check           ',
                'Final Adv. stat bonus': 'Lookup            ', 'Check_Statutory_Rate': 'Check            ',
                'Leave Encashment': 'Leave Encashment', 'LE Days': 'LE Days',
                'Basic Pay Sales Master': 'Basic', 'Calculated_IBP': 'IBP',
                'Calc_LE_Payout': 'Calc', 'Check_LE_Days_Variance': 'Check             ',
                'Lookup_PTax': 'Lookup              ', 'Check_PTax': 'check',
                'Inventory Recovery_input': 'Lookup               ', 'Check_Inventory_Recovery': 'Check                ',
                'NP recovery': 'NP_Lookup', 'Calc_NP_Recovery': 'Calc ', 'Check_NP_Variance': 'check ',
                'Facility Recovery_input': 'Lookup                ', 'Check_Facility_Recovery': 'check  '
            }
            
            display_df = result_df.rename(columns=final_renaming)
            status.update(label="✅ Run Completed Successfully!", state="complete")

    # Render Sheets
    st.dataframe(display_df, use_container_width=True)

    # Render Interface Summary Logs
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

    # --- 5. EXCEL EXPORT WORKBOOK GENERATION ENGINE (STRICTLY UNHIGHLIGHTED) ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        display_df.to_excel(writer, index=False, sheet_name='FnF_Master_Audit')
        workbook  = writer.book
        worksheet = writer.sheets['FnF_Master_Audit']
        
        normal_border_fmt = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0F172A', 'font_color': 'white', 'border': 1})

        for col_num, col_name in enumerate(display_df.columns):
            if 'lookup' in col_name.lower() or 'check' in col_name.lower() or 'calc' in col_name.lower():
                worksheet.set_column(col_num, col_num, 16, normal_border_fmt)
            else:
                worksheet.set_column(col_num, col_num, 18)
            worksheet.write(0, col_num, col_name, header_fmt)

        # Build Sheet 2 exception observations sequential matrix rows
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
