import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

st.set_page_config(page_title="Reinsurance Automated System", layout="wide")
st.title("🤖 ระบบประมวลผล Reinsurance Bordereaux อัตโนมัติ 100%")

# ==========================================
# CORE ENGINE: อัลกอริทึมวิเคราะห์อัตโนมัติในตัวเอง
# ==========================================

def clean_to_float(val):
    """ คลีนทุกรูปแบบตัวเลข/ข้อความ/วงเล็บ ให้เป็น Float ปลอดภัย 100% """
    if pd.isna(val) or val is None:
        return 0.0
    s = str(val).strip()
    if not s or s.lower() in ['nan', 'none', 'null', '-', 'n/a']:
        return 0.0
    # จัดการตัวเลขติดลบในวงเล็บ เช่น (1,000.00) -> -1000.00
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    # ลบสัญลักษณ์ที่ไม่ใช่ตัวเลข
    s = re.sub(r'[^0-9.-]', '', s)
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0

def process_sheet_fully_auto(uploaded_file, sheet_keyword):
    """ สแกนและดึงข้อมูลจาก Sheet ที่ต้องการแบบอัตโนมัติโดยไม่ต้องให้ผู้ใช้ระบุหัวตาราง """
    xl = pd.ExcelFile(uploaded_file)
    
    # 1. ค้นหา Sheet ที่ถูกต้องอัตโนมัติ
    target_sheet = None
    for s in xl.sheet_names:
        if sheet_keyword.lower() in s.lower() and 'pivot' not in s.lower():
            target_sheet = s
            break
    if not target_sheet:
        target_sheet = xl.sheet_names[0]

    # อ่านข้อมูลสแกน 25 แถวแรก
    df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=None, nrows=25)
    
    # 2. ค้นหาแถวที่เป็น Header จริง
    header_row = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v).lower() for v in row.values if pd.notnull(v)])
        if any(k in row_str for k in ['claim', 'สินไหม', 'policy', 'gross', 'net', 'amount', 'loss']):
            header_row = idx
            break
            
    # อ่าน Dataframe ตาม Header ที่พบ
    df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    # 3. ระบุคอลัมน์ Claim No.
    claim_col = None
    for c in df.columns:
        if any(k in c.lower() for k in ['claim', 'สินไหม', 'เลขที่', 'no']):
            claim_col = c
            break
    if not claim_col:
        claim_col = df.columns[0] # Fallback คอลัมน์แรก

    # 4. ระบุคอลัมน์ Gross และ Net ด้วยการสแกนประเภทข้อมูลยอดรวม
    numeric_scores = []
    for c in df.columns:
        # แปลงข้อมูลทั้งคอลัมน์เป็นตัวเลข
        col_floats = df[c].apply(clean_to_float)
        total_val = col_floats.sum()
        if total_val > 0:
            numeric_scores.append((c, total_val, col_floats))

    # เรียงลำดับคอลัมน์ที่มีมูลค่าตัวเลขสูงสุด
    numeric_scores.sort(key=lambda x: x[1], reverse=True)

    gross_col_name, net_col_name = None, None
    
    # ค้นหาจากชื่อก่อน
    for c, val, floats in numeric_scores:
        c_lower = c.lower()
        if 'gross' in c_lower and not gross_col_name:
            gross_col_name = c
        elif 'net' in c_lower and not net_col_name:
            net_col_name = c

    # ถ้าชื่อไม่ชัดเจน ให้ใช้คอลัมน์ที่มีมูลค่าเงินสูงสุดเป็น Gross และอันดับสองเป็น Net
    if not gross_col_name and len(numeric_scores) > 0:
        gross_col_name = numeric_scores[0][0]
    if not net_col_name and len(numeric_scores) > 1:
        net_col_name = numeric_scores[1][0]
    elif not net_col_name and gross_col_name:
        net_col_name = gross_col_name

    # สร้าง Clean DataFrame
    res_df = pd.DataFrame()
    res_df['Claim No.'] = df[claim_col].astype(str).str.strip().str.upper()
    res_df['Gross'] = df[gross_col_name].apply(clean_to_float) if gross_col_name else 0.0
    res_df['Net'] = df[net_col_name].apply(clean_to_float) if net_col_name else res_df['Gross']

    # กรองเฉพาะแถวที่มีเลข Claim จริง
    res_df = res_df[~res_df['Claim No.'].isin(['NAN', 'NONE', '', 'NULL', 'TOTAL', 'ยอดรวม'])]
    
    return res_df

# ==========================================
# STEP 1: UPLOAD & AUTOMATIC EXECUTION
# ==========================================
st.header("📌 อัปโหลดไฟล์ Original Data")
file_input = st.file_uploader("ลากไฟล์ Excel (.xlsx) มาวางที่นี่ระบบจะทำงานให้อัตโนมัติทันที", type=["xlsx"])

if file_input:
    with st.spinner("🤖 ระบบกำลังวิเคราะห์โครงสร้างไฟล์และประมวลผลให้อัตโนมัติ..."):
        # ประมวลผล Sheet Settle และ Incurred/Reserve อัตโนมัติในตัวเอง
        df_settle = process_sheet_fully_auto(file_input, 'Settle')
        df_reserve = process_sheet_fully_auto(file_input, 'Incurred')

        # จัดกลุ่มและรวมยอดตาม Claim No.
        grp_settle = df_settle.groupby('Claim No.', as_index=False).agg({
            'Gross': 'sum',
            'Net': 'sum'
        }).rename(columns={'Gross': 'Settle Gross Loss', 'Net': 'Settle Net Loss Retention'})

        grp_reserve = df_reserve.groupby('Claim No.', as_index=False).agg({
            'Gross': 'sum',
            'Net': 'sum'
        }).rename(columns={'Gross': 'Reserve Gross Loss', 'Net': 'Reserve Net Loss Retention'})

        # รวมข้อมูล 2 ฝั่งเข้าด้วยกัน (Outer Join)
        master_df = pd.merge(grp_settle, grp_reserve, on='Claim No.', how='outer').fillna(0.0)

        # เติม คอลัมน์มาตรฐาน Bordereaux
        for col in ['Row Labels', 'Sub Class', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด']:
            master_df[col] = ''
        master_df['Status'] = 'Closed'

        # จัดเรียง คอลัมน์ตาม Format มาตรฐาน
        final_cols = ['Row Labels', 'Sub Class', 'Claim No.', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด', 
                      'Settle Gross Loss', 'Settle Net Loss Retention', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']
        master_df = master_df[final_cols]

        # คำนวณ ยอดรวม
        tot_sg = float(master_df['Settle Gross Loss'].sum())
        tot_sn = float(master_df['Settle Net Loss Retention'].sum())
        tot_rg = float(master_df['Reserve Gross Loss'].sum())
        tot_rn = float(master_df['Reserve Net Loss Retention'].sum())

    st.success("✅ ประมวลผลสำเร็จ! ข้อมูลถูกดึงและคำนวณเรียบร้อยโดยไม่ต้องเลือกคอลัมน์")

    # Display Metrics
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Settle Gross", f"{tot_sg:,.2f}")
    m2.metric("Total Settle Net", f"{tot_sn:,.2f}")
    m3.metric("Total Reserve Gross", f"{tot_rg:,.2f}")
    m4.metric("Total Reserve Net", f"{tot_rn:,.2f}")

    st.markdown("---")
    st.subheader("📋 ตัวอย่าง Bordereaux Master Data")
    st.dataframe(master_df.head(15), use_container_width=True)

    # ==========================================
    # BUILD EXCEL WORKBOOK (STYLING)
    # ==========================================
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Details Claim"

    # Styles Setup
    fill_settle = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    fill_reserve = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid")
    fill_total = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    font_bold = Font(name="Aptos", size=11, bold=True)
    font_main = Font(name="Aptos", size=11)
    
    border_grid = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    border_total = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    # Line 1: Empty, Line 2: Super Headers
    ws.append([])
    ws.append(['', '', '', '', '', '', '', 'Settle 12/2025', '', 'Reserve as at 31/12/2025', '', ''])
    ws.merge_cells('H2:I2')
    ws.merge_cells('J2:K2')

    ws['H2'].fill, ws['H2'].font, ws['H2'].alignment = fill_settle, font_bold, Alignment(horizontal='center')
    ws['J2'].fill, ws['J2'].font, ws['J2'].alignment = fill_reserve, font_bold, Alignment(horizontal='center')

    # Line 3: Column Headers
    ws.append(final_cols)
    for c in range(1, 13):
        cell = ws.cell(row=3, column=c)
        cell.font = font_bold
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_grid
        if c in [8, 9]: cell.fill = fill_settle
        elif c in [10, 11]: cell.fill = fill_reserve

    # Data Rows
    r_idx = 4
    for row in master_df.itertuples(index=False):
        ws.append(list(row))
        for c in range(1, 13):
            cell = ws.cell(row=r_idx, column=c)
            cell.font = font_main
            cell.border = border_grid
            if c in [8, 9, 10, 11]:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')
        r_idx += 1

    # Total Row
    ws.append(['', '', '', '', '', '', '', tot_sg, tot_sn, tot_rg, tot_rn, ''])
    for c in range(1, 13):
        cell = ws.cell(row=r_idx, column=c)
        cell.font = font_bold
        cell.fill = fill_total
        cell.border = border_total
        if c in [8, 9, 10, 11]:
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(horizontal='right')

    excel_out = io.BytesIO()
    wb.save(excel_out)

    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Bordereaux Master (Excel)",
        data=excel_out.getvalue(),
        file_name="Bordereaux_Master_Automated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
