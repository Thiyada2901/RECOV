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
st.title("🤖 ระบบ AI ประมวลผลและตรวจสอบ Reinsurance Claims")

# ---------------------------------------------------------
# Helper Function สำหรับทำความสะอาดตัวเลข
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

# ---------------------------------------------------------
# Step 1: Upload File
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
    
    st.info(f"📄 อ่านข้อมูลจาก Sheet: **Settle = {set_sheet}** | **Incurred/Reserve = {inc_sheet}**")

    df_inc_ori = pd.read_excel(uploaded_file, sheet_name=inc_sheet)
    df_set_ori = pd.read_excel(uploaded_file, sheet_name=set_sheet)
    
    df_inc_ori.columns = [str(c).strip() for c in df_inc_ori.columns]
    df_set_ori.columns = [str(c).strip() for c in df_set_ori.columns]

    # ---------------------------------------------------------
    # UI ให้ผู้ใช้เลือกคอลัมน์การเงินด้วยตนเอง (แก้ปัญหาจับคอลัมน์ผิด 100%)
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("⚙️ กรุณาเลือกชื่อคอลัมน์การเงินให้ตรงกับไฟล์ Excel ของคุณ")
    
    cols_set = list(df_set_ori.columns)
    cols_inc = list(df_inc_ori.columns)

    col_ui1, col_ui2 = st.columns(2)

    with col_ui1:
        st.write("🟢 **ฝั่ง Settlement Data:**")
        sel_claim_set = st.selectbox("คอลัมน์ เลขที่สินไหม (Claim No):", cols_set, index=0, key="c_set")
        sel_sg = st.selectbox("คอลัมน์ Settle Gross:", cols_set, index=min(1, len(cols_set)-1), key="sg")
        sel_sn = st.selectbox("คอลัมน์ Settle Net:", cols_set, index=min(2, len(cols_set)-1), key="sn")

    with col_ui2:
        st.write("🔵 **ฝั่ง Reserve / Incurred Data:**")
        sel_claim_inc = st.selectbox("คอลัมน์ เลขที่สินไหม (Claim No):", cols_inc, index=0, key="c_inc")
        sel_rg = st.selectbox("คอลัมน์ Reserve Gross:", cols_inc, index=min(1, len(cols_inc)-1), key="rg")
        sel_rn = st.selectbox("คอลัมน์ Reserve Net:", cols_inc, index=min(2, len(cols_inc)-1), key="rn")

    # ---------------------------------------------------------
    # Step 2: Merge & Build Master Bordereaux
    # ---------------------------------------------------------
    st.markdown("---")
    st.header("📋 Step 2: ตรวจสอบ Bordereaux Master (Details Claim)")

    # Processing Settle
    df_set_ori['Claim_Clean'] = df_set_ori[sel_claim_set].astype(str).str.strip().str.upper()
    df_set_ori['Settle Gross Loss'] = clean_numeric_series(df_set_ori[sel_sg])
    df_set_ori['Settle Net Loss Retention'] = clean_numeric_series(df_set_ori[sel_sn])

    df_set_grp = df_set_ori.groupby('Claim_Clean', as_index=False).agg({
        'Settle Gross Loss': 'sum',
        'Settle Net Loss Retention': 'sum'
    }).rename(columns={'Claim_Clean': 'Claim No.'})

    # Processing Reserve
    df_inc_ori['Claim_Clean'] = df_inc_ori[sel_claim_inc].astype(str).str.strip().str.upper()
    df_inc_ori['Reserve Gross Loss'] = clean_numeric_series(df_inc_ori[sel_rg])
    df_inc_ori['Reserve Net Loss Retention'] = clean_numeric_series(df_inc_ori[sel_rn])

    df_inc_grp = df_inc_ori.groupby('Claim_Clean', as_index=False).agg({
        'Reserve Gross Loss': 'sum',
        'Reserve Net Loss Retention': 'sum'
    }).rename(columns={'Claim_Clean': 'Claim No.'})

    # Merge Both
    df_bor = pd.merge(df_set_grp, df_inc_grp, on='Claim No.', how='outer').fillna(0.0)

    # Clean & Reorder
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

    # Display Results
    st.dataframe(df_bor.head(10), use_container_width=True)

    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Total Settle Gross", f"{s_settle_gross:,.2f}")
    col_b.metric("Total Settle Net", f"{s_settle_net:,.2f}")
    col_c.metric("Total Reserve Gross", f"{s_res_gross:,.2f}")
    col_d.metric("Total Reserve Net", f"{s_res_net:,.2f}")

    # Export Excel
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
