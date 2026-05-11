import streamlit as st
import pandas as pd
import io
import msoffcrypto
import re

# --- 1. UI CONFIG ---
st.set_page_config(page_title="Checkpoint Portal", page_icon="📂", layout="wide")

st.markdown("""
    <style>
    .header-container {
        text-align: center;
        padding: 1.5rem;
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
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

def load_input_file(file):
    if file is None: return None
    try:
        all_sheets = pd.read_excel(file, sheet_name=None, engine='openpyxl')
        sheets = list(all_sheets.values())
        main_df = sheets[0]
        id_col = main_df.columns[0]
        main_df[id_col] = main_df[id_col].astype(str).str.strip()
        for i in range(1, len(sheets)):
            other_sheet = sheets[i]
            other_id = other_sheet.columns[0]
            other_sheet[other_id] = other_sheet[other_id].astype(str).str.strip()
            main_df = main_df.merge(other_sheet, left_on=id_col, right_on=other_id, how='left', suffixes=('', f'_s{i}'))
        return main_df
    except: return None

def load_encrypted_xlsx(file, password):
    try:
        decrypted_workbook = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_workbook)
        return pd.read_excel(decrypted_workbook, engine='openpyxl')
    except: return None

def get_col(df, options_str):
    if df is None: return None
    options = [opt.strip() for opt in options_str.split('/')]
    for opt in options:
        if opt in df.columns: return opt
    return None

def clean_num(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return 0.0
        s = str(val).replace(',', '').replace('(', '-').replace(')', '').strip()
        return float(s)
    except: return 0.0

def compare_logic(val1, val2, is_date=False):
    """
    If Numeric: returns difference (rounded to 2).
    If Text: returns Yes/No.
    """
    if is_date:
        try:
            if pd.isna(val1) or pd.isna(val2): return "No"
            d1 = pd.to_datetime(val1).strftime('%d-%m-%Y')
            d2 = pd.to_datetime(val2).strftime('%d-%m-%Y')
            return "Yes" if d1 == d2 else f"No ({d1})"
        except: return "No"

    # Check if values are numeric
    s1 = str(val1).replace(',', '').strip()
    s2 = str(val2).replace(',', '').strip()
    
    try:
        # Attempt numeric comparison
        float(s1)
        float(s2)
        n1, n2 = clean_num(val1), clean_num(val2)
        diff = round(n1 - n2, 2)
        return 0 if abs(diff) < 0.01 else diff
    except ValueError:
        # Fallback to Text comparison
        v1 = str(val1).strip().lower()
        v2 = str(val2).strip().lower()
        if v1 == v2: return "Yes"
        return "No"

# --- 3. MAIN APP ---

st.markdown('<div class="header-container"><h1>📂 Welcome to Checkpoint Portal</h1></div>', unsafe_allow_html=True)

with st.sidebar:
    input_file = st.file_uploader("1. Input Sheet", type=['xlsx'])
    hc_file = st.file_uploader("2. HC Report", type=['xlsx', 'csv'])
    bank_file = st.file_uploader("3. Bank Account Details", type=['xlsx', 'csv'])
    salary_file = st.file_uploader("4. Salary Register (Protected)", type=['xlsx'])
    salary_pw = st.text_input("Decryption Password", type="password")
    run_validation = st.button("🚀 RUN AUDIT")

if run_validation and all([input_file, hc_file, bank_file, salary_file]):
    with st.status("🔍 Processing Master Files...") as status:
        df_input = load_input_file(input_file)
        df_hc = load_file(hc_file)
        df_bank = load_file(bank_file)
        df_salary = load_encrypted_xlsx(salary_file, salary_pw)

        if df_salary is not None:
            df_salary['Employee No'] = df_salary['Employee No'].astype(str).str.strip()
            hc_id_col = get_col(df_hc, "User/Employee ID / Employee ID") or df_hc.columns[0]
            bank_id_col = get_col(df_bank, "Users Sys Id / Account Number") or df_bank.columns[0]
            
            df_hc[hc_id_col] = df_hc[hc_id_col].astype(str).str.strip()
            df_bank[bank_id_col] = df_bank[bank_id_col].astype(str).str.strip()
            df_input['Employee ID'] = df_input.iloc[:, 0].astype(str).str.strip()

            final_df = df_salary.copy()
            final_df = final_df.merge(df_hc, left_on='Employee No', right_on=hc_id_col, how='left')
            final_df = final_df.merge(df_bank, left_on='Employee No', right_on=bank_id_col, how='left')
            final_df = final_df.merge(df_input, left_on='Employee No', right_on='Employee ID', how='left', suffixes=('','_inp'))

            # Logic
            hc_join = get_col(df_hc, "Employment Details Group Date of Joining / Group Date of Joining")
            hc_exit = get_col(df_hc, "Employment Details Actual Exit Date / Actual Exit Date")
            hc_resig = get_col(df_hc, "Employment Details Date of Resignation / Resignation Date")
            bank_ifsc = get_col(df_bank, "IFSC Code / IFSC")
            bank_acc = get_col(df_bank, "accountNumber / Account Number")

            final_df['Lookup_Entity Name'] = final_df['Company'].astype(str).str.split('-').str[-1].str.strip()
            final_df['chk1'] = final_df.apply(lambda x: compare_logic(x.get('Entity Name Description'), x['Lookup_Entity Name']), axis=1)
            final_df['Lookup_GroupJoinedOn'] = final_df[hc_join] if hc_join else None
            final_df['chk2'] = final_df.apply(lambda x: compare_logic(x.get('GroupJoinedOn'), x['Lookup_GroupJoinedOn'], True), axis=1)
            final_df['Lookup_Left Date'] = final_df[hc_exit] if hc_exit else None
            final_df['chk3'] = final_df.apply(lambda x: compare_logic(x.get('Left Date'), x['Lookup_Left Date'], True), axis=1)
            final_df['lookup_Designation'] = final_df['Position Title']
            final_df['chk4'] = final_df.apply(lambda x: compare_logic(x.get('Designation'), x['lookup_Designation']), axis=1)
            final_df['lookup_State Name'] = final_df['State'].astype(str).str.split('-').str[-1].str.strip()
            final_df['chk5'] = final_df.apply(lambda x: compare_logic(x.get('State Name'), x['lookup_State Name']), axis=1)
            final_df['lookup_IFSC'] = final_df[bank_ifsc] if bank_ifsc else None
            final_df['chk6'] = final_df.apply(lambda x: compare_logic(x.get('Primary Bank IFSC'), x['lookup_IFSC']), axis=1)
            final_df['Lookup_acc_no'] = final_df[bank_acc].astype(str).str.replace('.0', '', regex=False) if bank_acc else ""
            final_df['chk7'] = final_df.apply(lambda x: compare_logic(x.get('Primary Bank Account No'), x['Lookup_acc_no']), axis=1)
            final_df['Lookup_ResignedDate'] = final_df[hc_resig] if hc_resig else None
            final_df['chk8'] = final_df.apply(lambda x: compare_logic(x.get('ResignedDate'), x['Lookup_ResignedDate'], True), axis=1)

            # Leave Encashment Logic
            inp_le_days = get_col(df_input, "Leave Encashment / Total Leave Encashment / Leave Encashment Days")
            final_df['Lookup_Leave Encash'] = final_df[inp_le_days] if inp_le_days else 0
            final_df['chk9'] = final_df.apply(lambda x: compare_logic(x.get('Leave Encashment Days'), x['Lookup_Leave Encash']), axis=1)
            final_df['calc_le'] = final_df.apply(lambda x: (max(clean_num(x.get('Basic')), 0.5 * clean_num(x.get('IBP'))) / 30) * clean_num(x.get('Leave Encashment Days')), axis=1)
            final_df['chk10'] = final_df.apply(lambda x: compare_logic(x.get('LeaveEncas'), x['calc_le']), axis=1)

            # Arrears
            arr_cfg = {'basic': ('Ar-BASIC', 'Arr Basic Pay'), 'HRA': ('ArrHRA', 'Arr HRA'), 'consy': ('Consy_Bn_S', 'Consistency Bonus'), 'stt': ('Stt_Bo_S_A', 'Arr Adv Stt Bonus'), 'sales': ('Sales_Cm_A', 'Sales linked'), 'mob': ('Moble_S_Ar', 'Mobile Allow')}
            for k, (s_o, i_o) in arr_cfg.items():
                sc = get_col(df_salary, s_o); ic = get_col(df_input, i_o)
                final_df[f'v_{k}'] = final_df[sc] if sc else 0
                final_df[f'l_{k}'] = final_df[ic] if ic else 0
                final_df[f'c_{k}'] = final_df.apply(lambda x: compare_logic(x[f'v_{k}'], x[f'l_{k}']), axis=1)

            # --- COLUMN ORDERING ---
            final_order = [
                "Employee No", "Display Name", "Entity Name Description", "Lookup_Entity Name", "chk1",
                "Joined Date", "GroupJoinedOn", "Lookup_GroupJoinedOn", "chk2",
                "Left Date", "Lookup_Left Date", "chk3",
                "Status Description", "Email Id", "Period", "Payroll Code", "Designation",
                "lookup_Designation", "chk4", "PAN No", "Gender", "Birth Date",
                "Full Location Code", "Full Location Description", "State Name", "lookup_State Name", "chk5",
                "Department Code", "Department Description", "Cost Centre Code", "Cost Center Description",
                "Category Description", "B.A.Code", "B.A.Description", "Company PF Reg.Code",
                "Provident Fund No", "UAN Number", "PRAN", "Primary Bank Name", "Primary Bank IFSC",
                "lookup_IFSC", "chk6", "Primary Bank Account No", "Lookup_acc_no", "chk7",
                "Division Code", "Division Name", "CalcPFOn", "CTC", "ResignedDate", "Lookup_ResignedDate", "chk8",
                "Inactive Reason", "Total Days", "Leave Encashment Days", "Lookup_Leave Encash", "chk9",
                "LeaveEncas", "chk10", "Basic", "IBP", "calc_le", # Moved chk10 here as requested
                "v_basic", "l_basic", "c_basic", "v_HRA", "l_HRA", "c_HRA", "v_consy", "l_consy", "c_consy",
                "v_stt", "l_stt", "c_stt", "v_sales", "l_sales", "c_sales", "v_mob", "l_mob", "c_mob",
                "GROSS EARNING", "P.TAX"
            ]

            result = final_df.reindex(columns=final_order)
            mapping = {
                "chk1":"Check", "chk2":"Check ", "chk3":"Check  ", "chk6":"Check   ", "chk9":"Check    ", "chk10":"check",
                "chk4":"check ", "chk5":"check  ", "chk7":"check   ", "chk8":"check    ",
                "v_basic":"Basic_S_Ar", "l_basic":"lookup_basic", "c_basic":"check     ",
                "v_HRA":"HRA_S_Arrs", "l_HRA":"lookup_HRA", "c_HRA":"check      ",
                "v_consy":"Consy_Bn_S", "l_consy":"lookup_consy", "c_consy":"check       ",
                "v_stt":"Stt_Bo_S_A", "l_stt":"lookup_stt_bo", "c_stt":"check        ",
                "v_sales":"Sales_Cm_A", "l_sales":"lookup_sales_cm", "c_sales":"check         ",
                "v_mob":"Mobile_S_Ar", "l_mob":"lookup_mobile_s", "c_mob":"check          ",
                "calc_le":"calc", "LeaveEncas":"Leave Encas"
            }
            display_df = result.rename(columns=mapping)

            st.dataframe(display_df, use_container_width=True)

            # --- EXCEL EXPORT WITH COLORS ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                display_df.to_excel(writer, index=False, sheet_name='Audit_Report')
                workbook  = writer.book
                worksheet = writer.sheets['Audit_Report']
                
                # Formats
                yellow = workbook.add_format({'bg_color': '#FFF9C4', 'border': 1})
                green = workbook.add_format({'bg_color': '#C8E6C9', 'border': 1})
                header = workbook.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1})

                for col_num, value in enumerate(display_df.columns):
                    worksheet.write(0, col_num, value, header)
                    if 'lookup' in value.lower():
                        worksheet.set_column(col_num, col_num, 18, yellow)
                    elif 'check' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, green)
                    else:
                        worksheet.set_column(col_num, col_num, 18)

            st.download_button("📥 DOWNLOAD COLORED REPORT", output.getvalue(), "FnF_Audit_Final.xlsx", use_container_width=True)
        else:
            st.error("Decryption failed.")
