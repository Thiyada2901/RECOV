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
st.title("🤖 ระบบ AI ประมวลผลและตรวจสอบ Reinsurance Claims (Fully Automated)")

# ---------------------------------------------------------
# Helper Functions สำหรับจัดการตัวเลขและการสแกนโครงสร้าง Excel
# ---------------------------------------------------------
def clean_numeric_series(series):
    """ แปลงข้อความ/Text/Comma/วงเล็บ ให้เป็น Float ตัวเลขที่นำไปคำนวณได้จริง """
    if series is None:
        return pd.Series(0.0)
    
    s_clean = (
        series.astype(str)
        .str.replace('฿', '', regex=False)
        .str.replace('$', '', regex=False)
        .str.replace(',', '', regex=False)
        .str.replace(' ', '', regex=False)
        .str.replace('(', '-', regex=False)
        .str.replace(')', '', regex=False)
        .str.strip()
    )
    s_clean = s_clean.replace(['-', 'N/A', 'nan', 'None', 'null', ''], '0')
    return pd.to_numeric(s_clean, errors='coerce').fillna(0.0)

def read_excel_smart(uploaded_file, sheet_name):
    """ อ่านไฟล์ Excel และค้นหา Header Row อัตโนมัติ ป้องกันปัญหา Header ซ้อนหรือ Merge Cells """
    # อ่าน 20 แถวแรกเพื่อหาแถวที่มีชื่อคอลัมน์หลัก
    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, nrows=20)
    header_idx = 0
    
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(val).lower() for val in row.values if pd.notnull(val)])
        if any(k in row_str for k in ['claim', 'สินไหม', 'gross', 'net', 'policy', 'กรมธรรม์']):
            header_idx = idx
            break
            
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_idx)
    # ลบชื่อคอลัมน์ที่เป็น Unnamed หรือค่าว่าง
    df.columns = [str(c).strip() for c in df.columns]
    return df

def find_best_financial_columns(df, is_settle=True):
    """ AI สแกนหาคอลัมน์ Gross และ Net จากชื่อและข้อมูลตัวเลขจริง """
    cols = df.columns
    gross_col = None
    net_col = None
    
    # 1. ค้นหาจากคีย์เวิร์ดในชื่อคอลัมน์
    g_keywords = ['settle gross', 'paid gross', 'reserve gross', 'estimated gross', 'gross loss', 'gross amount', 'ค่าสินไหมรวม', 'gross']
    n_keywords = ['settle net', 'paid net', 'reserve net', 'estimated net', 'net loss', 'retention', 'ค่าสินไหมสุทธิ', 'net']
    
    for c in cols:
        c_lower = str(c).lower()
        if any(k in c_lower for k in g_keywords) and not gross_col:
            if clean_numeric_series(df[c]).sum() > 0:
                gross_col = c
        if any(k in c_lower for k in n_keywords) and not net_col:
            if clean_numeric_series(df[c]).sum() > 0:
                net_col = c

    # 2. หากยังหา Gross/Net ไม่เจอ ให้สแกนคอลัมน์ที่เป็นตัวเลขทั้งหมดแล้วจัดอันดับยอดรวมสูงสุด
    numeric_cols = []
    for c in cols:
        s = clean_numeric_series(df[c])
        s_sum = s.sum()
        if s_sum > 0:
            numeric_cols.append((c, s_sum))
            
    numeric_cols.sort(key=lambda x: x[1], reverse=True)
    
    if not gross_col and len(numeric_cols) > 0:
        gross_col = numeric_cols[0][0]
    if not net_col and len(numeric_cols) > 1:
        net_col = numeric_cols[1][0]
    elif not net_col and gross_col:
        net_col = gross_col

    return gross_col, net_col

def find_claim_column(df):
    """ สแกนหาคอลัมน์ Claim No. อัตโนมัติ """
    for c in df.columns:
        c_lower = str(c).lower()
        if any(k in c_lower for k in ['claim', 'สินไหม', 'เลขที่']):
            return c
    return df.columns[0]

# ---------------------------------------------------------
# Step 1: Upload File & Automatic Processing
# ---------------------------------------------------------
st.header("📌 Step 1: โยนไฟล์ Original Data (Excel)")
uploaded_file = st.file_uploader("เลือกไฟล์ Original Data (.xlsx)", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    
    # Auto detect sheet names
    inc_sheets = [s for s in xl.sheet_names if 'Incurred' in s and 'Pivot' not in s]
    set_sheets = [s for s in xl.sheet_names if 'Settle' in s and 'Pivot' not in s]
    
    inc_sheet = inc_sheets[0] if inc_sheets else xl.sheet_names[0]
    set_sheet = set_sheets[0] if set_sheets else (xl.sheet_names[1] if len(xl.sheet_names) > 1 else xl.sheet_names[0])
    
    df_inc_ori = read_excel_smart(uploaded_file, inc_sheet)
    df_set_ori = read_excel_smart(uploaded_file, set_sheet)

    # Auto Detect Columns
    claim_col_set = find_claim_column(df_set_ori)
    claim_col_inc = find_claim_column(df_inc_ori)

    col_sg, col_sn = find_best_financial_columns(df_set_ori, is_settle=True)
    col_rg, col_rn = find_best_financial_columns(df_inc_ori, is_settle=False)

    st.success("✅ AI ทำการวิเคราะห์โครงสร้างไฟล์และจับคู่คอลัมน์การเงินให้อัตโนมัติเรียบร้อยแล้ว!")

    # ---------------------------------------------------------
    # Step 2: Merge & Build Master Bordereaux
    # ---------------------------------------------------------
    st.markdown("---")
    st.header("📋 Step 2: ตรวจสอบ Bordereaux Master (Details Claim)")

    # Processing Settle Data
    df_set_ori['Claim_Clean'] = df_set_ori[claim_col_set].astype(str).str.strip().str.upper()
    df_set_ori['Settle Gross Loss'] = clean_numeric_series(df_set_ori[col_sg]) if col_sg else 0.0
    df_set_ori['Settle Net Loss Retention'] = clean_numeric_series(df_set_ori[col_sn]) if col_sn else df_set_ori['Settle Gross Loss']

    df_set_grp = df_set_ori.groupby('Claim_Clean', as_index=False).agg({
        'Settle Gross Loss': 'sum',
        'Settle Net Loss Retention': 'sum'
    }).rename(columns={'Claim_Clean': 'Claim No.'})

    # Processing Reserve Data
    df_inc_ori['Claim_Clean'] = df_inc_ori[claim_col_inc].astype(str).str.strip().str.upper()
    df_inc_ori['Reserve Gross Loss'] = clean_numeric_series(df_inc_ori[col_rg]) if col_rg else 0.0
    df_inc_ori['Reserve Net Loss Retention'] = clean_numeric_series(df_inc_ori[col_rn]) if col_rn else df_inc_ori['Reserve Gross Loss']

    df_inc_grp = df_inc_ori.groupby('Claim_Clean', as_index=False).agg({
        'Reserve Gross Loss': 'sum',
        'Reserve Net Loss Retention': 'sum'
    }).rename(columns={'Claim_Clean': 'Claim No.'})

    # Merge Data
    df_bor = pd.merge(df_set_grp, df_inc_grp, on='Claim No.', how='outer').fillna(0.0)

    # Fill default text columns
    for c in ['Row Labels', 'Sub Class', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด']:
        df_bor[c] = ''
    df_bor['Status'] = 'Closed'

    cols_order = ['Row Labels', 'Sub Class', 'Claim No.', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด', 
                  'Settle Gross Loss', 'Settle Net Loss Retention', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']
    df_bor = df_bor[cols_order]

    # Calculate Totals
    s_settle_gross = float(df_bor['Settle Gross Loss'].sum())
    s_settle_net = float(df_bor['Settle Net Loss Retention'].sum())
    s_res_gross = float(df_bor['Reserve Gross Loss'].sum())
    s_res_net = float(df_bor['Reserve Net Loss Retention'].sum())

    st.dataframe(df_bor.head(10), use_container_width=True)

    # Display Summary Metrics
    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Total Settle Gross", f"{s_settle_gross:,.2f}")
    col_b.metric("Total Settle Net", f"{s_settle_net:,.2f}")
    col_c.metric("Total Reserve Gross", f"{s_res_gross:,.2f}")
    col_d.metric("Total Reserve Net", f"{s_res_net:,.2f}")

    # Export Excel Button
    bor_wb = openpyxl.Workbook()
    ws_bor = bor_wb.active
    ws_bor.title = "Details Claim"

    fill_settle = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    fill_reserve = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid")
    fill_total = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
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
        ws_bor.append(row_vals)
        for c in range(1, 13):
            cell = ws_bor.cell(row=current_row, column=c)
            cell.font = font_regular
            cell.border = thin_border
            if c in [8, 9, 10, 11]:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')
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

    bor_buffer = io.BytesIO()
    bor_wb.save(bor_buffer)

    st.download_button(
        label="📥 ดาวน์โหลด Bordereaux Master (Excel)",
        data=bor_buffer.getvalue(),
        file_name="Bordereaux_Claim_Master_Styled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
