import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import re

st.set_page_config(page_title="AI Reinsurance System", layout="wide")
st.title("🤖 ระบบ AI ประมวลผลและตรวจสอบ Reinsurance Claims (Automated Filtering)")

# ---------------------------------------------------------
# Helper Functions สำหรับทำความสะอาดข้อมูลและการกรอง
# ---------------------------------------------------------
def clean_num(val):
    """ แปลงข้อความ/Text/Comma/วงเล็บ ให้เป็น Float ตัวเลขที่นำไปคำนวณได้จริง """
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

def is_excluded_row(row):
    """ สแกนหาคำระบุสถานะเคลมที่ไม่คุ้มครองเพื่อตัดทิ้งทั้งบรรทัด """
    row_str = " ".join([str(val).lower().strip() for val in row.values if pd.notnull(val)])
    
    # คำคีย์เวิร์ดบ่งบอกว่าเคลมนี้ไม่อยู่ในความคุ้มครอง / ปฏิเสธจ่าย
    exclude_keywords = [
        'not cover', 'non-cover', 'excluded', 'decline', 'declined', 
        'reject', 'rejected', 'non-reinsurance', 'out of coverage',
        'void', 'cancel', 'cancelled', 'uncovered', 'ex-gratia'
    ]
    
    for kw in exclude_keywords:
        if kw in row_str:
            return True
    return False

def find_best_col(df, keywords):
    """ ค้นหาคอลัมน์ที่ตรงกับ Keyword """
    for kw in keywords:
        for col in df.columns:
            if kw.lower() in str(col).strip().lower():
                return col
    return None

def process_sheet_with_filter(uploaded_file, sheet_keyword):
    xl = pd.ExcelFile(uploaded_file)
    
    # 1. ค้นหา Sheet
    target_sheet = next((s for s in xl.sheet_names if sheet_keyword.lower() in s.lower() and 'pivot' not in s.lower()), xl.sheet_names[0])
    
    # อ่าน Header
    df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=None, nrows=20)
    header_row = 0
    for idx, row in df_raw.iterrows():
        r_str = " ".join([str(v).lower() for v in row.values if pd.notnull(v)])
        if any(k in r_str for k in ['claim', 'สินไหม', 'policy', 'gross', 'net', 'loss']):
            header_row = idx
            break
            
    df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    # 2. กรองแถวที่ไม่คุ้มครองออก (Excluded Filter)
    excluded_mask = df.apply(is_excluded_row, axis=1)
    df_filtered = df[~excluded_mask].copy()

    # 3. ระบุคอลัมน์ Claim No.
    col_claim = find_best_col(df_filtered, ['เลขที่สินไหม', 'claim no', 'claim_no', 'claim']) or df_filtered.columns[0]
    
    # 4. ระบุคอลัมน์ Gross และ Net
    col_gross = find_best_col(df_filtered, ['gross loss', 'settle gross', 'reserve gross', 'estimated gross', 'gross amount', 'gross'])
    col_net = find_best_col(df_filtered, ['net loss', 'settle net', 'reserve net', 'retention', 'net amount', 'net'])

    # Fallback ค้นหาคอลัมน์ตัวเลขถ้าหาจากชื่อไม่เจอ
    if not col_gross or not col_net:
        num_cols = []
        for c in df_filtered.columns:
            s_val = df_filtered[c].apply(clean_num).sum()
            if s_val > 0:
                num_cols.append((c, s_val))
        num_cols.sort(key=lambda x: x[1], reverse=True)
        if not col_gross and len(num_cols) > 0: col_gross = num_cols[0][0]
        if not col_net and len(num_cols) > 1: col_net = num_cols[1][0]
        elif not col_net and col_gross: col_net = col_gross

    # สรุปข้อมูล
    res = pd.DataFrame()
    res['Claim No.'] = df_filtered[col_claim].astype(str).str.strip().str.upper()
    res['Gross'] = df_filtered[col_gross].apply(clean_num) if col_gross else 0.0
    res['Net'] = df_filtered[col_net].apply(clean_num) if col_net else res['Gross']

    # ลบแถวที่เป็นค่าว่างหรือตัวอักษรรวมยอด
    res = res[~res['Claim No.'].isin(['NAN', 'NONE', '', 'NULL', 'TOTAL', 'ยอดรวม', 'SUBTOTAL'])]
    
    return res

# ---------------------------------------------------------
# Step 1: Upload & Auto Processing
# ---------------------------------------------------------
st.header("📌 Step 1: โยนไฟล์ Original Data (Excel)")
uploaded_file = st.file_uploader("เลือกไฟล์ Original Data (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("🤖 ระบบกำลังประมวลผล คัดกรองรายการเคลมที่ไม่คุ้มครองออกให้อัตโนมัติ..."):
        # อ่านข้อมูลพร้อมกรองเคลมไม่อยู่ในความคุ้มครอง
        df_set = process_sheet_with_filter(uploaded_file, 'Settle')
        df_inc = process_sheet_with_filter(uploaded_file, 'Incurred')

        # Groupby ตาม Claim No.
        grp_set = df_set.groupby('Claim No.', as_index=False).agg({
            'Gross': 'sum',
            'Net': 'sum'
        }).rename(columns={'Gross': 'Settle Gross Loss', 'Net': 'Settle Net Loss Retention'})

        grp_inc = df_inc.groupby('Claim No.', as_index=False).agg({
            'Gross': 'sum',
            'Net': 'sum'
        }).rename(columns={'Gross': 'Reserve Gross Loss', 'Net': 'Reserve Net Loss Retention'})

        # Merge รวมตาราง
        df_master = pd.merge(grp_set, grp_inc, on='Claim No.', how='outer').fillna(0.0)

        # เติมคอลัมน์มาตรฐาน
        for c in ['Row Labels', 'Sub Class', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด']:
            df_master[c] = ''
        df_master['Status'] = 'Closed'

        cols_order = ['Row Labels', 'Sub Class', 'Claim No.', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด', 
                      'Settle Gross Loss', 'Settle Net Loss Retention', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']
        df_master = df_master[cols_order]

        # คำนวณ ยอดรวม
        tot_sg = float(df_master['Settle Gross Loss'].sum())
        tot_sn = float(df_master['Settle Net Loss Retention'].sum())
        tot_rg = float(df_master['Reserve Gross Loss'].sum())
        tot_rn = float(df_master['Reserve Net Loss Retention'].sum())

    st.success("✅ กรองรายการเคลมที่ไม่คุ้มครองออก และประมวลผลตาราง Master เรียบร้อยแล้ว!")

    # Display Summary Metrics
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Settle Gross", f"{tot_sg:,.2f}")
    col2.metric("Total Settle Net", f"{tot_sn:,.2f}")
    col3.metric("Total Reserve Gross", f"{tot_rg:,.2f}")
    col4.metric("Total Reserve Net", f"{tot_rn:,.2f}")

    st.markdown("---")
    st.dataframe(df_master.head(15), use_container_width=True)

    # ---------------------------------------------------------
    # Export Excel File (Master Styled)
    # ---------------------------------------------------------
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
    for r in df_master.itertuples(index=False):
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
        file_name="Bordereaux_Master_Filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
