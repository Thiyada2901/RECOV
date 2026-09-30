import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import re

st.set_page_config(page_title="AI Reinsurance System", layout="wide")
st.title("🤖 ระบบประมวลผล Reinsurance Bordereaux (4-Step Complete Workflow)")

# ---------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------
def clean_num(val):
    if pd.isna(val) or val is None:
        return 0.0
    s = str(val).strip()
    if not s or s.lower() in ['nan', 'none', 'null', '-', 'n/a']:
        return 0.0
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    s = re.sub(r'[^0-9.-]', '', s)
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0

def find_col_by_keywords(df, keywords):
    for kw in keywords:
        for col in df.columns:
            if kw.lower() in str(col).strip().lower():
                return col
    return None

# Initialization Session State
if 'df_master_raw' not in st.session_state:
    st.session_state.df_master_raw = None

# =========================================================
# Step 1: Upload File & Sheet Detection
# =========================================================
st.header("📌 Step 1: อัปโหลดไฟล์ Original Data (Excel)")
uploaded_file = st.file_uploader("เลือกไฟล์ Original Data (.xlsx)", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    
    # อ่าน Sheet Settle และ Incurred
    sheet_set = next((s for s in xl.sheet_names if 'settle' in s.lower() and 'pivot' not in s.lower()), xl.sheet_names[0])
    sheet_inc = next((s for s in xl.sheet_names if ('incurred' in s.lower() or 'reserve' in s.lower()) and 'pivot' not in s.lower()), xl.sheet_names[-1])

    df_set_raw = pd.read_excel(uploaded_file, sheet_name=sheet_set)
    df_inc_raw = pd.read_excel(uploaded_file, sheet_name=sheet_inc)

    df_set_raw.columns = [str(c).strip() for c in df_set_raw.columns]
    df_inc_raw.columns = [str(c).strip() for c in df_inc_raw.columns]

    # ค้นหา คอลัมน์สำคัญ
    col_claim_set = find_col_by_keywords(df_set_raw, ['เลขที่สินไหม', 'claim no', 'claim_no', 'claim']) or df_set_raw.columns[0]
    col_sg = find_col_by_keywords(df_set_raw, ['settle gross', 'paid gross', 'gross loss', 'gross amount', 'gross'])
    col_sn = find_col_by_keywords(df_set_raw, ['settle net', 'paid net', 'net loss', 'retention', 'net'])
    col_status_set = find_col_by_keywords(df_set_raw, ['status', 'remark', 'coverage', 'หมายเหตุ', 'สถานะ'])

    col_claim_inc = find_col_by_keywords(df_inc_raw, ['เลขที่สินไหม', 'claim no', 'claim_no', 'claim']) or df_inc_raw.columns[0]
    col_rg = find_col_by_keywords(df_inc_raw, ['reserve gross', 'estimated gross', 'gross reserve', 'incurred gross', 'gross'])
    col_rn = find_col_by_keywords(df_inc_raw, ['reserve net', 'estimated net', 'net reserve', 'retention', 'net'])

    # =========================================================
    # Step 2: Mapping & Processing
    # =========================================================
    st.markdown("---")
    st.header("📌 Step 2: ประมวลผลและเชื่อมโยงข้อมูล (Data Mapping)")

    # Settle Data
    df_set = pd.DataFrame()
    df_set['Claim No.'] = df_set_raw[col_claim_set].astype(str).str.strip().str.upper()
    df_set['Settle Gross Loss'] = df_set_raw[col_sg].apply(clean_num) if col_sg else 0.0
    df_set['Settle Net Loss Retention'] = df_set_raw[col_sn].apply(clean_num) if col_sn else df_set['Settle Gross Loss']
    df_set['Status_Raw'] = df_set_raw[col_status_set].astype(str).str.strip() if col_status_set else 'Normal'

    df_set = df_set[~df_set['Claim No.'].isin(['NAN', 'NONE', '', 'NULL', 'TOTAL', 'ยอดรวม'])]
    grp_set = df_set.groupby('Claim No.', as_index=False).agg({
        'Settle Gross Loss': 'sum',
        'Settle Net Loss Retention': 'sum',
        'Status_Raw': 'first'
    })

    # Reserve Data
    df_inc = pd.DataFrame()
    df_inc['Claim No.'] = df_inc_raw[col_claim_inc].astype(str).str.strip().str.upper()
    df_inc['Reserve Gross Loss'] = df_inc_raw[col_rg].apply(clean_num) if col_rg else 0.0
    df_inc['Reserve Net Loss Retention'] = df_inc_raw[col_rn].apply(clean_num) if col_rn else df_inc['Reserve Gross Loss']

    df_inc = df_inc[~df_inc['Claim No.'].isin(['NAN', 'NONE', '', 'NULL', 'TOTAL', 'ยอดรวม'])]
    grp_inc = df_inc.groupby('Claim No.', as_index=False).agg({
        'Reserve Gross Loss': 'sum',
        'Reserve Net Loss Retention': 'sum'
    })

    # Merge Raw Master
    df_merged = pd.merge(grp_set, grp_inc, on='Claim No.', how='outer').fillna(0.0)
    
    for c in ['Row Labels', 'Sub Class', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด']:
        df_merged[c] = ''
    df_merged['Status'] = 'Closed'

    st.session_state.df_master_raw = df_merged
    st.success("✅ โหลดข้อมูลพื้นฐานเข้าสู่ระบบเรียบร้อยแล้ว")

    # =========================================================
    # Step 3: Verify & Adjust Claims (คัดกรองรายการเคลมที่ไม่คุ้มครอง)
    # =========================================================
    st.markdown("---")
    st.header("📌 Step 3: ตรวจสอบและตัดเคลมที่ไม่คุ้มครอง (Claim Verification & Filtering)")
    st.info("💡 สามารถเลือกตัดสถานะเคลมที่ไม่คุ้มครองออก หรือค้นหา Claim No. เพื่อตัดออกทีละรายการเพื่อให้ยอดตรงกับ Master")

    df_working = st.session_state.df_master_raw.copy()

    # 3.1 Filter ตาม Status/Remark ที่ดึงมาจากไฟล์
    all_statuses = list(df_working['Status_Raw'].unique())
    selected_statuses = st.multiselect(
        "เลือกสถานะเคลมที่ต้องการนำมาคิดคำนวณ (เอาติ๊กออกเพื่อตัดเคลมที่ไม่คุ้มครองทิ้ง):",
        options=all_statuses,
        default=all_statuses
    )
    df_filtered = df_working[df_working['Status_Raw'].isin(selected_statuses)].copy()

    # 3.2 Exclude Claim No. ระบุเฉพาะ
    exclude_claims_text = st.text_area("ระบุ Claim No. ที่ต้องการตัดออกจากการคำนวณ (คั่นด้วยเครื่องหมายจุลภาค , หรือขึ้นบรรทัดใหม่):", "")
    if exclude_claims_text.strip():
        ex_list = [c.strip().upper() for c in re.split(r'[\n,]', exclude_claims_text) if c.strip()]
        df_filtered = df_filtered[~df_filtered['Claim No.'].isin(ex_list)]

    # =========================================================
    # Step 4: Summary & Export Master Bordereaux
    # =========================================================
    st.markdown("---")
    st.header("📌 Step 4: สรุปยอดรวมสุทธิและดาวน์โหลด (Export Master Bordereaux)")

    cols_order = ['Row Labels', 'Sub Class', 'Claim No.', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด', 
                  'Settle Gross Loss', 'Settle Net Loss Retention', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']
    
    df_final = df_filtered[cols_order].copy()

    # คำนวณยอดรวมสุทธิหลังกรอง Step 3
    tot_sg = float(df_final['Settle Gross Loss'].sum())
    tot_sn = float(df_final['Settle Net Loss Retention'].sum())
    tot_rg = float(df_final['Reserve Gross Loss'].sum())
    tot_rn = float(df_final['Reserve Net Loss Retention'].sum())

    # แสดงผลตัวเลขสรุป
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Settle Gross", f"{tot_sg:,.2f}")
    c2.metric("Total Settle Net", f"{tot_sn:,.2f}")
    c3.metric("Total Reserve Gross", f"{tot_rg:,.2f}")
    c4.metric("Total Reserve Net", f"{tot_rn:,.2f}")

    st.markdown("#### ตารางผลลัพธ์ Master Bordereaux (หลังตัดเคลมไม่อยู่ในเงื่อนไข)")
    st.dataframe(df_final, use_container_width=True)

    # สร้างไฟล์ Excel สำหรับ Download
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Details Claim"

    fill_settle = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    fill_reserve = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid")
    fill_total = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    font_bold = Font(name="Aptos", size=11, bold=True)
    font_regular = Font(name="Aptos", size=11)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    total_border = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    ws.append([])
    ws.append(['', '', '', '', '', '', '', 'Settle 12/2025', '', 'Reserve as at 31/12/2025', '', ''])
    ws.merge_cells('H2:I2')
    ws.merge_cells('J2:K2')

    ws['H2'].fill = fill_settle
    ws['H2'].font = font_bold
    ws['H2'].alignment = Alignment(horizontal='center')
    ws['J2'].fill = fill_reserve
    ws['J2'].font = font_bold
    ws['J2'].alignment = Alignment(horizontal='center')

    ws.append(cols_order)
    for c in range(1, 13):
        cell = ws.cell(row=3, column=c)
        cell.font = font_bold
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
        if c in [8, 9]: cell.fill = fill_settle
        elif c in [10, 11]: cell.fill = fill_reserve

    r_idx = 4
    for r in df_final.itertuples(index=False):
        ws.append(list(r))
        for c in range(1, 13):
            cell = ws.cell(row=r_idx, column=c)
            cell.font = font_regular
            cell.border = thin_border
            if c in [8, 9, 10, 11]:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')
        r_idx += 1

    ws.append(['', '', '', '', '', '', '', tot_sg, tot_sn, tot_rg, tot_rn, ''])
    for c in range(1, 13):
        cell = ws.cell(row=r_idx, column=c)
        cell.font = font_bold
        cell.fill = fill_total
        cell.border = total_border
        if c in [8, 9, 10, 11]:
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(horizontal='right')

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)

    st.download_button(
        label="📥 ดาวน์โหลด Bordereaux Master (Excel)",
        data=excel_buffer.getvalue(),
        file_name="Bordereaux_Master_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
