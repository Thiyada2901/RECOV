import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="AI Reinsurance System", layout="wide")
st.title("🤖 ระบบ AI ประมวลผลและตรวจสอบ Reinsurance Claims (Master Structure & Styling)")

# ---------------------------------------------------------
# Helper Functions สำหรับแก้ปัญหาตัวเลขกลายเป็น 0
# ---------------------------------------------------------
def clean_numeric_series(series):
    """ แปลงข้อความ/Text ให้เป็น Float ตัวเลขที่คำนวณได้จริง """
    if series is None:
        return pd.Series(0.0)
    
    # แปลงเป็น String แล้วลบสัญลักษณ์ที่มักติดมากับตัวเลข
    s_clean = (
        series.astype(str)
        .str.replace('฿', '', regex=False)
        .str.replace('$', '', regex=False)
        .str.replace(',', '', regex=False)
        .str.replace(' ', '', regex=False)
        .str.strip()
    )
    # แปลงค่าที่เป็นข้อความเปล่า หรือตัวขีด '-' ให้เป็น 0
    s_clean = s_clean.replace(['-', 'N/A', 'nan', 'None', 'null', ''], '0')
    
    # แปลงเป็น numeric
    return pd.to_numeric(s_clean, errors='coerce').fillna(0.0)

def find_col_smart(df, possible_keywords):
    """ ค้นหาคอลัมน์จากคีย์เวิร์ดแบบยืดหยุ่น """
    # 1. ค้นหาแบบตรงตัวก่อน
    for name in possible_keywords:
        for col in df.columns:
            if name.lower() == str(col).strip().lower():
                return col
    # 2. ค้นหาแบบ Partial Match (มีคำนั้นผสมอยู่)
    for name in possible_keywords:
        for col in df.columns:
            if name.lower() in str(col).strip().lower():
                return col
    return None

# ---------------------------------------------------------
# Step 1: Upload File
# ---------------------------------------------------------
st.header("📌 Step 1: โยนไฟล์ Original Data (Excel)")
uploaded_file = st.file_uploader("เลือกไฟล์ Original Data (.xlsx)", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    
    # Auto detect sheet names
    inc_sheet = [s for s in xl.sheet_names if 'Incurred' in s and 'Pivot' not in s][0]
    set_sheet = [s for s in xl.sheet_names if 'Settle' in s and 'Pivot' not in s][0]
    
    df_inc_ori = pd.read_excel(uploaded_file, sheet_name=inc_sheet)
    df_set_ori = pd.read_excel(uploaded_file, sheet_name=set_sheet)
    
    st.success("✅ AI อ่านและทำความเข้าใจโครงสร้าง Ori Data เรียบร้อยแล้ว!")
    
    # ---------------------------------------------------------
    # Step 2: Merge & Build Master Bordereaux (Details Claim)
    # ---------------------------------------------------------
    st.markdown("---")
    st.header("📋 Step 2: ตรวจสอบ Bordereaux Master (Details Claim - ตกแต่งสีสันตาม Master)")
    
    # Clean whitespace ในชื่อคอลัมน์
    df_inc_ori.columns = [str(c).strip() for c in df_inc_ori.columns]
    df_set_ori.columns = [str(c).strip() for c in df_set_ori.columns]

    # Mapping คอลัมน์ Settlement Data
    col_claim_set = find_col_smart(df_set_ori, ['เลขที่สินไหม', 'Claim No', 'Claim No.', 'เลขสินไหม', 'Claim_No'])
    col_branch_set = find_col_smart(df_set_ori, ['สาขา', 'Branch'])
    col_subclass_set = find_col_smart(df_set_ori, ['Sub Class', 'SubClass', 'Class'])
    col_policy_set = find_col_smart(df_set_ori, ['เลขที่กรมธรรม์', 'Policy No', 'Policy No.'])
    col_date_set = find_col_smart(df_set_ori, ['วันที่เกิดเหตุ', 'Loss Date', 'Date of Loss'])
    col_insured_set = find_col_smart(df_set_ori, ['ชื่อผู้เอาประกัน', 'Insured Name', 'Insured'])
    col_province_set = find_col_smart(df_set_ori, ['จังหวัด', 'Province'])
    
    # ค้นหาคอลัมน์การเงินของ Settle
    col_paid_set = find_col_smart(df_set_ori, ['ค่าสินไหมจ่าย', 'ค่าสินไหม', 'Settle Amount', 'Paid Amount', 'Settle Gross', 'Paid Gross', 'Amount'])
    col_ret_set = find_col_smart(df_set_ori, ['Retention Amount', 'Retention', 'Settle Net', 'Paid Net'])

    if not col_claim_set:
        col_claim_set = df_set_ori.columns[0]

    # Clean Claim No. ป้องกัน Merge แล้วจับคู่ไม่เจอ
    df_set_ori[col_claim_set] = df_set_ori[col_claim_set].astype(str).str.strip().str.upper()

    # แปลงคอลัมน์การเงินให้เป็น float ชัวร์ๆ (แก้ปัญหาเลข 0)
    if col_paid_set:
        df_set_ori[col_paid_set] = clean_numeric_series(df_set_ori[col_paid_set])
    if col_ret_set:
        df_set_ori[col_ret_set] = clean_numeric_series(df_set_ori[col_ret_set])

    agg_dict_set = {}
    rename_dict_set = {col_claim_set: 'Claim No.'}

    if col_branch_set: agg_dict_set[col_branch_set] = 'first'; rename_dict_set[col_branch_set] = 'Row Labels'
    if col_subclass_set: agg_dict_set[col_subclass_set] = 'first'; rename_dict_set[col_subclass_set] = 'Sub Class'
    if col_policy_set: agg_dict_set[col_policy_set] = 'first'; rename_dict_set[col_policy_set] = 'Policy No.'
    if col_date_set: agg_dict_set[col_date_set] = 'first'; rename_dict_set[col_date_set] = 'Loss Date'
    if col_insured_set: agg_dict_set[col_insured_set] = 'first'; rename_dict_set[col_insured_set] = 'Insured Name'
    if col_province_set: agg_dict_set[col_province_set] = 'first'; rename_dict_set[col_province_set] = 'จังหวัด'
    if col_paid_set: agg_dict_set[col_paid_set] = 'sum'; rename_dict_set[col_paid_set] = 'Settle Gross Loss'
    if col_ret_set: agg_dict_set[col_ret_set] = 'sum'; rename_dict_set[col_ret_set] = 'Settle Net Loss Retention'

    df_set_grp = df_set_ori.groupby(col_claim_set, as_index=False).agg(agg_dict_set).rename(columns=rename_dict_set)

    # Mapping คอลัมน์ Incurred Data
    col_claim_inc = find_col_smart(df_inc_ori, ['เลขที่สินไหม', 'Claim No', 'Claim No.', 'เลขสินไหม', 'Claim_No'])
    col_branch_inc = find_col_smart(df_inc_ori, ['สาขา', 'Branch'])
    col_subclass_inc = find_col_smart(df_inc_ori, ['Sub Class', 'SubClass'])
    col_policy_inc = find_col_smart(df_inc_ori, ['เลขที่กรมธรรม์', 'Policy No'])
    col_date_inc = find_col_smart(df_inc_ori, ['วันที่เกิดเหตุ', 'Loss Date'])
    col_insured_inc = find_col_smart(df_inc_ori, ['ชื่อผู้เอาประกัน', 'Insured Name'])
    col_province_inc = find_col_smart(df_inc_ori, ['จังหวัด', 'Province'])
    
    # ค้นหาคอลัมน์การเงินของ Incurred
    col_est_inc = find_col_smart(df_inc_ori, ['ประมาณการค่าสินไหม', 'Reserve Amount', 'Estimated Loss', 'Reserve Gross', 'Incurred Amount', 'Estimate'])
    col_ret_inc = find_col_smart(df_inc_ori, ['Retention (By type of loss)', 'Retention', 'Reserve Net', 'Net Reserve'])
    col_status_inc = find_col_smart(df_inc_ori, ['สถานะ', 'Status'])

    if not col_claim_inc:
        col_claim_inc = df_inc_ori.columns[0]

    # Clean Claim No. ฝั่ง Incurred
    df_inc_ori[col_claim_inc] = df_inc_ori[col_claim_inc].astype(str).str.strip().str.upper()

    # แปลงคอลัมน์การเงินให้เป็น float ชัวร์ๆ (แก้ปัญหาเลข 0)
    if col_est_inc:
        df_inc_ori[col_est_inc] = clean_numeric_series(df_inc_ori[col_est_inc])
    if col_ret_inc:
        df_inc_ori[col_ret_inc] = clean_numeric_series(df_inc_ori[col_ret_inc])

    agg_dict_inc = {}
    rename_dict_inc = {col_claim_inc: 'Claim No.'}

    if col_branch_inc: agg_dict_inc[col_branch_inc] = 'first'; rename_dict_inc[col_branch_inc] = 'Row Labels'
    if col_subclass_inc: agg_dict_inc[col_subclass_inc] = 'first'; rename_dict_inc[col_subclass_inc] = 'Sub Class'
    if col_policy_inc: agg_dict_inc[col_policy_inc] = 'first'; rename_dict_inc[col_policy_inc] = 'Policy No.'
    if col_date_inc: agg_dict_inc[col_date_inc] = 'first'; rename_dict_inc[col_date_inc] = 'Loss Date'
    if col_insured_inc: agg_dict_inc[col_insured_inc] = 'first'; rename_dict_inc[col_insured_inc] = 'Insured Name'
    if col_province_inc: agg_dict_inc[col_province_inc] = 'first'; rename_dict_inc[col_province_inc] = 'จังหวัด'
    if col_est_inc: agg_dict_inc[col_est_inc] = 'sum'; rename_dict_inc[col_est_inc] = 'Reserve Gross Loss'
    if col_ret_inc: agg_dict_inc[col_ret_inc] = 'sum'; rename_dict_inc[col_ret_inc] = 'Reserve Net Loss Retention'
    if col_status_inc: agg_dict_inc[col_status_inc] = 'first'; rename_dict_inc[col_status_inc] = 'Status'

    df_inc_grp = df_inc_ori.groupby(col_claim_inc, as_index=False).agg(agg_dict_inc).rename(columns=rename_dict_inc)

    # รับประกันคอลัมน์ครบถ้วน
    for required_col in ['Claim No.', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']:
        if required_col not in df_inc_grp.columns:
            if required_col in ['Reserve Gross Loss', 'Reserve Net Loss Retention']:
                df_inc_grp[required_col] = 0.0
            elif required_col == 'Status':
                df_inc_grp[required_col] = 'Closed'
            else:
                df_inc_grp[required_col] = ''

    # Merge Data
    df_bor = pd.merge(df_set_grp, df_inc_grp[['Claim No.', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']], on='Claim No.', how='outer')

    for c in ['Row Labels', 'Sub Class', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด']:
        if c not in df_bor.columns:
            df_bor[c] = ''

    # Clean ข้อมูลหลัง Merge อีกรอบเพื่อความชัวร์
    df_bor['Settle Gross Loss'] = clean_numeric_series(df_bor.get('Settle Gross Loss'))
    df_bor['Settle Net Loss Retention'] = clean_numeric_series(df_bor.get('Settle Net Loss Retention'))
    df_bor['Reserve Gross Loss'] = clean_numeric_series(df_bor.get('Reserve Gross Loss'))
    df_bor['Reserve Net Loss Retention'] = clean_numeric_series(df_bor.get('Reserve Net Loss Retention'))
    df_bor['Status'] = df_bor['Status'].fillna('Closed')

    cols_order = ['Row Labels', 'Sub Class', 'Claim No.', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด', 
                  'Settle Gross Loss', 'Settle Net Loss Retention', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']
    df_bor = df_bor[cols_order]

    # รวมผลลัพธ์
    s_settle_gross = float(df_bor['Settle Gross Loss'].sum())
    s_settle_net = float(df_bor['Settle Net Loss Retention'].sum())
    s_res_gross = float(df_bor['Reserve Gross Loss'].sum())
    s_res_net = float(df_bor['Reserve Net Loss Retention'].sum())

    st.dataframe(df_bor.head(15), use_container_width=True)

    # แสดงผลยอดรวมบนหน้าจอ Streamlit เพื่อเช็กว่าเลขมาถูกต้องไหม
    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Total Settle Gross", f"{s_settle_gross:,.2f}")
    col_b.metric("Total Settle Net", f"{s_settle_net:,.2f}")
    col_c.metric("Total Reserve Gross", f"{s_res_gross:,.2f}")
    col_d.metric("Total Reserve Net", f"{s_res_net:,.2f}")

    # ---------------------------------------------------------
    # Styling Bordereaux Workbook (ส่วนที่เหลือทำงานปกติ)
    # ---------------------------------------------------------
    bor_wb = openpyxl.Workbook()
    ws_bor = bor_wb.active
    ws_bor.title = "Details Claim"

    fill_settle = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    fill_reserve = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid")
    fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    font_bold = Font(name="Aptos", size=11, bold=True)
    font_header = Font(name="Aptos", size=11, bold=True)
    font_regular = Font(name="Aptos", size=11)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    total_border = Border(
        top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000')
    )

    ws_bor.append([])
    ws_bor.append(['', '', '', '', '', '', '', 'Settle 12/2025', '', 'Reserve as at 31/12/2025', '', ''])
    ws_bor.merge_cells('H2:I2')
    ws_bor.merge_cells('J2:K2')

    ws_bor['H2'].fill = fill_settle
    ws_bor['H2'].font = font_bold
    ws_bor['H2'].alignment = Alignment(horizontal='center', vertical='center')

    ws_bor['J2'].fill = fill_reserve
    ws_bor['J2'].font = font_bold
    ws_bor['J2'].alignment = Alignment(horizontal='center', vertical='center')

    ws_bor.append(cols_order)
    for col_num in range(1, 13):
        cell = ws_bor.cell(row=3, column=col_num)
        cell.font = font_header
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
        if col_num in [8, 9]:
            cell.fill = fill_settle
        elif col_num in [10, 11]:
            cell.fill = fill_reserve

    current_row = 4
    for r in df_bor.itertuples(index=False):
        row_vals = list(r)
        if pd.notnull(row_vals[4]):
            try:
                row_vals[4] = pd.to_datetime(row_vals[4]).strftime('%Y-%m-%d')
            except:
                pass
        ws_bor.append(row_vals)
        
        for c in range(1, 13):
            cell = ws_bor.cell(row=current_row, column=c)
            cell.font = font_regular
            cell.border = thin_border
            if c in [8, 9, 10, 11]:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif c == 5:
                cell.alignment = Alignment(horizontal='center')
        current_row += 1

    tot_row_vals = ['', '', '', '', '', '', '', s_settle_gross, s_settle_net, s_res_gross, s_res_net, '']
    ws_bor.append(tot_row_vals)
    for c in range(1, 13):
        cell = ws_bor.cell(row=current_row, column=c)
        cell.font = font_bold
        cell.fill = fill_total
        cell.border = total_border
        if c in [8, 9, 10, 11]:
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(horizontal='right')

    for col in ws_bor.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_bor.column_dimensions[col_letter].width = max(max_len + 3, 12)

    bor_buffer = io.BytesIO()
    bor_wb.save(bor_buffer)

    st.download_button(
        label="📥 ดาวน์โหลด Bordereaux Master (มีสีสันตาม Master Excel)",
        data=bor_buffer.getvalue(),
        file_name="Bordereaux_Claim_Master_Styled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("✅ อนุมัติ Bordereaux (Approve & Proceed)"):
        st.session_state['approved_step1'] = True
