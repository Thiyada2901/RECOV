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
# Helper Functions สำหรับจัดการตัวเลขและการค้นหาคอลัมน์
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

def find_col_smart(df, possible_keywords):
    """ ค้นหา index/ชื่อคอลัมน์จากคีย์เวิร์ดแบบยืดหยุ่น """
    for name in possible_keywords:
        for col in df.columns:
            if name.lower() == str(col).strip().lower():
                return col
    for name in possible_keywords:
        for col in df.columns:
            if name.lower() in str(col).strip().lower():
                return col
    return df.columns[0] if len(df.columns) > 0 else None

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
    
    df_inc_ori = pd.read_excel(uploaded_file, sheet_name=inc_sheet)
    df_set_ori = pd.read_excel(uploaded_file, sheet_name=set_sheet)
    
    df_inc_ori.columns = [str(c).strip() for c in df_inc_ori.columns]
    df_set_ori.columns = [str(c).strip() for c in df_set_ori.columns]

    st.success("✅ AI อ่านโครงสร้างไฟล์เรียบร้อยแล้ว!")

    # ---------------------------------------------------------
    # Mapping Selection UI (ป้องกันคอลัมน์ผิดพลาด)
    # ---------------------------------------------------------
    st.subheader("⚙️ ตรวจสอบการเลือกคอลัมน์การเงิน (หากตัวเลขไม่ตรง สามารถปรับเลือกชื่อคอลัมน์ได้ที่นี่)")
    
    cols_set_list = list(df_set_ori.columns)
    cols_inc_list = list(df_inc_ori.columns)

    def_sg = find_col_smart(df_set_ori, ['Settle Gross Loss', 'Paid Gross', 'Gross Loss', 'Gross Amount', 'ค่าสินไหมรวม', 'Gross'])
    def_sn = find_col_smart(df_set_ori, ['Settle Net Loss Retention', 'Paid Net', 'Net Loss', 'Retention', 'ค่าสินไหมสุทธิ', 'Net'])
    
    def_rg = find_col_smart(df_inc_ori, ['Reserve Gross Loss', 'Estimated Gross', 'Gross Reserve', 'Incurred Gross', 'Gross Loss', 'Gross'])
    def_rn = find_col_smart(df_inc_ori, ['Reserve Net Loss Retention', 'Estimated Net', 'Net Reserve', 'Retention', 'Net Loss', 'Net'])

    idx_sg = cols_set_list.index(def_sg) if def_sg in cols_set_list else 0
    idx_sn = cols_set_list.index(def_sn) if def_sn in cols_set_list else min(1, len(cols_set_list)-1)
    idx_rg = cols_inc_list.index(def_rg) if def_rg in cols_inc_list else 0
    idx_rn = cols_inc_list.index(def_rn) if def_rn in cols_inc_list else min(1, len(cols_inc_list)-1)

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    sel_set_gross = col_m1.selectbox("Settle Gross Column:", cols_set_list, index=idx_sg)
    sel_set_net = col_m2.selectbox("Settle Net Column:", cols_set_list, index=idx_sn)
    sel_inc_gross = col_m3.selectbox("Reserve Gross Column:", cols_inc_list, index=idx_rg)
    sel_inc_net = col_m4.selectbox("Reserve Net Column:", cols_inc_list, index=idx_rn)

    # ---------------------------------------------------------
    # Step 2: Merge & Build Master Bordereaux (Details Claim)
    # ---------------------------------------------------------
    st.markdown("---")
    st.header("📋 Step 2: ตรวจสอบ Bordereaux Master (Details Claim)")

    # Mapping คอลัมน์พื้นฐาน Settlement
    col_claim_set = find_col_smart(df_set_ori, ['เลขที่สินไหม', 'Claim No', 'Claim No.', 'เลขสินไหม'])
    col_branch_set = find_col_smart(df_set_ori, ['สาขา', 'Branch'])
    col_subclass_set = find_col_smart(df_set_ori, ['Sub Class', 'SubClass', 'Class'])
    col_policy_set = find_col_smart(df_set_ori, ['เลขที่กรมธรรม์', 'Policy No', 'Policy No.'])
    col_date_set = find_col_smart(df_set_ori, ['วันที่เกิดเหตุ', 'Loss Date', 'Date of Loss'])
    col_insured_set = find_col_smart(df_set_ori, ['ชื่อผู้เอาประกัน', 'Insured Name', 'Insured'])
    col_province_set = find_col_smart(df_set_ori, ['จังหวัด', 'Province'])

    df_set_ori['Claim_Clean'] = df_set_ori[col_claim_set].astype(str).str.strip().str.upper()
    df_set_ori['Settle Gross Loss'] = clean_numeric_series(df_set_ori[sel_set_gross])
    df_set_ori['Settle Net Loss Retention'] = clean_numeric_series(df_set_ori[sel_set_net])

    agg_dict_set = {'Settle Gross Loss': 'sum', 'Settle Net Loss Retention': 'sum'}
    rename_dict_set = {'Claim_Clean': 'Claim No.'}

    if col_branch_set: agg_dict_set[col_branch_set] = 'first'; rename_dict_set[col_branch_set] = 'Row Labels'
    if col_subclass_set: agg_dict_set[col_subclass_set] = 'first'; rename_dict_set[col_subclass_set] = 'Sub Class'
    if col_policy_set: agg_dict_set[col_policy_set] = 'first'; rename_dict_set[col_policy_set] = 'Policy No.'
    if col_date_set: agg_dict_set[col_date_set] = 'first'; rename_dict_set[col_date_set] = 'Loss Date'
    if col_insured_set: agg_dict_set[col_insured_set] = 'first'; rename_dict_set[col_insured_set] = 'Insured Name'
    if col_province_set: agg_dict_set[col_province_set] = 'first'; rename_dict_set[col_province_set] = 'จังหวัด'

    df_set_grp = df_set_ori.groupby('Claim_Clean', as_index=False).agg(agg_dict_set).rename(columns=rename_dict_set)

    # Mapping คอลัมน์พื้นฐาน Reserve / Incurred
    col_claim_inc = find_col_smart(df_inc_ori, ['เลขที่สินไหม', 'Claim No', 'Claim No.', 'เลขสินไหม'])
    col_branch_inc = find_col_smart(df_inc_ori, ['สาขา', 'Branch'])
    col_subclass_inc = find_col_smart(df_inc_ori, ['Sub Class', 'SubClass'])
    col_policy_inc = find_col_smart(df_inc_ori, ['เลขที่กรมธรรม์', 'Policy No'])
    col_date_inc = find_col_smart(df_inc_ori, ['วันที่เกิดเหตุ', 'Loss Date'])
    col_insured_inc = find_col_smart(df_inc_ori, ['ชื่อผู้เอาประกัน', 'Insured Name'])
    col_province_inc = find_col_smart(df_inc_ori, ['จังหวัด', 'Province'])
    col_status_inc = find_col_smart(df_inc_ori, ['สถานะ', 'Status'])

    df_inc_ori['Claim_Clean'] = df_inc_ori[col_claim_inc].astype(str).str.strip().str.upper()
    df_inc_ori['Reserve Gross Loss'] = clean_numeric_series(df_inc_ori[sel_inc_gross])
    df_inc_ori['Reserve Net Loss Retention'] = clean_numeric_series(df_inc_ori[sel_inc_net])

    agg_dict_inc = {'Reserve Gross Loss': 'sum', 'Reserve Net Loss Retention': 'sum'}
    rename_dict_inc = {'Claim_Clean': 'Claim No.'}

    if col_branch_inc: agg_dict_inc[col_branch_inc] = 'first'; rename_dict_inc[col_branch_inc] = 'Row Labels'
    if col_subclass_inc: agg_dict_inc[col_subclass_inc] = 'first'; rename_dict_inc[col_subclass_inc] = 'Sub Class'
    if col_policy_inc: agg_dict_inc[col_policy_inc] = 'first'; rename_dict_inc[col_policy_inc] = 'Policy No.'
    if col_date_inc: agg_dict_inc[col_date_inc] = 'first'; rename_dict_inc[col_date_inc] = 'Loss Date'
    if col_insured_inc: agg_dict_inc[col_insured_inc] = 'first'; rename_dict_inc[col_insured_inc] = 'Insured Name'
    if col_province_inc: agg_dict_inc[col_province_inc] = 'first'; rename_dict_inc[col_province_inc] = 'จังหวัด'
    if col_status_inc: agg_dict_inc[col_status_inc] = 'first'; rename_dict_inc[col_status_inc] = 'Status'

    df_inc_grp = df_inc_ori.groupby('Claim_Clean', as_index=False).agg(agg_dict_inc).rename(columns=rename_dict_inc)

    if 'Status' not in df_inc_grp.columns:
        df_inc_grp['Status'] = 'Closed'

    # Merge Data
    df_bor = pd.merge(df_set_grp, df_inc_grp[['Claim No.', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']], on='Claim No.', how='outer')

    for c in ['Row Labels', 'Sub Class', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด']:
        if c not in df_bor.columns:
            df_bor[c] = ''

    df_bor['Settle Gross Loss'] = clean_numeric_series(df_bor.get('Settle Gross Loss'))
    df_bor['Settle Net Loss Retention'] = clean_numeric_series(df_bor.get('Settle Net Loss Retention'))
    df_bor['Reserve Gross Loss'] = clean_numeric_series(df_bor.get('Reserve Gross Loss'))
    df_bor['Reserve Net Loss Retention'] = clean_numeric_series(df_bor.get('Reserve Net Loss Retention'))
    df_bor['Status'] = df_bor['Status'].fillna('Closed')

    cols_order = ['Row Labels', 'Sub Class', 'Claim No.', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด', 
                  'Settle Gross Loss', 'Settle Net Loss Retention', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']
    df_bor = df_bor[cols_order]

    # คำนวณยอดรวม
    s_settle_gross = float(df_bor['Settle Gross Loss'].sum())
    s_settle_net = float(df_bor['Settle Net Loss Retention'].sum())
    s_res_gross = float(df_bor['Reserve Gross Loss'].sum())
    s_res_net = float(df_bor['Reserve Net Loss Retention'].sum())

    st.dataframe(df_bor.head(10), use_container_width=True)

    # Card แสดงผลยอดรวม 4 ช่องตรงเป๊ะตาม Master Image
    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Total Settle Gross", f"{s_settle_gross:,.2f}")
    col_b.metric("Total Settle Net", f"{s_settle_net:,.2f}")
    col_c.metric("Total Reserve Gross", f"{s_res_gross:,.2f}")
    col_d.metric("Total Reserve Net", f"{s_res_net:,.2f}")

    # ตรวจสอบความถูกต้องกับรูปภาพ[cite: 5]
    if abs(s_settle_gross - 446018209.84) < 1 and abs(s_res_gross - 1236790997.87) < 1:
        st.success("🎯 ยอดรวมการเงินตรงกับรูป Master 100% เรียบร้อยแล้ว!")
    else:
        st.warning("⚠️ ยอดรวมยังไม่ตรงกับรูป Master โปรดเปลี่ยนตัวเลือกในเมนู 'ตรวจสอบการเลือกคอลัมน์การเงิน' ด้านบนให้ตรงกับชื่อคอลัมน์ใน Excel")

    # ---------------------------------------------------------
    # Styling Bordereaux Workbook (Excel Export)
    # ---------------------------------------------------------
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
        label="📥 ดาวน์โหลด Bordereaux Master (Excel)",
        data=bor_buffer.getvalue(),
        file_name="Bordereaux_Claim_Master_Styled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("✅ อนุมัติ Bordereaux (Approve & Proceed)"):
        st.session_state['approved_step1'] = True

    # ---------------------------------------------------------
    # Step 3: Build Full Master Summary Claim (By Layer)
    # ---------------------------------------------------------
    if st.session_state.get('approved_step1'):
        st.markdown("---")
        st.header("📊 Step 3: AI สร้างตาราง Summary Claim (By Layer)")

        gross_pla = s_settle_gross + s_res_gross
        net_pla = s_settle_net + s_res_net

        def calc_layer_row(layer_name, limit, excess, gross, net):
            under_xl = min(max(net - excess, 0.0), limit) if net > excess else 0.0
            return {
                "Layer": layer_name, "Gross 100%": gross, "Net Loss": net,
                "Limit": limit, "Excess Point": excess, "Under XL": under_xl,
                "IRMC 40%": under_xl * 0.40, "Lockton 36%": under_xl * 0.36,
                "TQR 15%": under_xl * 0.15, "Aon 9%": under_xl * 0.09,
                "Total": under_xl
            }

        sum_wb = openpyxl.Workbook()
        ws_sum = sum_wb.active
        ws_sum.title = "P&E XL"

        fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        ws_sum.append(["Claim Flood 2025 as at 31/12/2025"])
        ws_sum['A1'].font = Font(name="Aptos", size=12, bold=True)

        headers = ["Section", "Layer", "Gross 100%", "Net Loss", "Limit", "Excess Point", "", "Under XL", "IRMC 40%", "Lockton 36%", "TQR 15%", "Aon 9%", "Total"]

        def add_summary_section(section_title, rows_data):
            ws_sum.append([])
            h_row = [section_title if i == 0 else headers[i] for i in range(len(headers))]
            ws_sum.append(h_row)

            r_idx = ws_sum.max_row
            for c in range(1, 14):
                cell = ws_sum.cell(row=r_idx, column=c)
                cell.fill = fill_yellow
                cell.font = font_bold
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')

            pct_row = ["", "", "", "", "", "", "", "", 0.40, 0.36, 0.15, 0.09, ""]
            ws_sum.append(pct_row)
            r_pct = ws_sum.max_row
            for c in range(9, 13):
                cell = ws_sum.cell(row=r_pct, column=c)
                cell.number_format = '0%'
                cell.alignment = Alignment(horizontal='right')

            for rd in rows_data:
                row_vals = [
                    "", rd["Layer"], rd["Gross 100%"], rd["Net Loss"], rd["Limit"],
                    rd["Excess Point"], "", rd["Under XL"], rd["IRMC 40%"],
                    rd["Lockton 36%"], rd["TQR 15%"], rd["Aon 9%"], rd["Total"]
                ]
                ws_sum.append(row_vals)
                r_curr = ws_sum.max_row
                for c in range(1, 14):
                    cell = ws_sum.cell(row=r_curr, column=c)
                    cell.font = font_regular
                    cell.border = thin_border
                    if c in [3, 4, 5, 6, 8, 9, 10, 11, 12, 13]:
                        cell.number_format = '#,##0.00'
                        cell.alignment = Alignment(horizontal='right')

        layer_2nd = calc_layer_row("2nd Layer", 220000000, 120000000, gross_pla, net_pla)
        pla_rows = [
            layer_2nd,
            calc_layer_row("3rd Layer", 1060000000, 340000000, gross_pla, net_pla),
            calc_layer_row("4th Layer", 2100000000, 1400000000, gross_pla, net_pla)
        ]
        add_summary_section("PLA XL", pla_rows)

        lsa_rows = [
            calc_layer_row("2nd Layer", 220000000, 120000000, s_settle_gross, s_settle_net),
            calc_layer_row("3rd Layer", 1060000000, 340000000, s_settle_gross, s_settle_net),
            calc_layer_row("4th Layer", 2100000000, 1400000000, s_settle_gross, s_settle_net)
        ]
        add_summary_section("LSA (1st Interim Payment)", lsa_rows)

        for col in ws_sum.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_sum.column_dimensions[col_letter].width = max(max_len + 3, 14)

        sum_buffer = io.BytesIO()
        sum_wb.save(sum_buffer)

        st.download_button(
            label="📥 ดาวน์โหลด Summary Claim By Layer Master (Excel)",
            data=sum_buffer.getvalue(),
            file_name="Summary_Claim_By_Layer_Master_Styled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.button("✅ อนุมัติ Summary By Layer (Approve & Generate PDF)"):
            st.session_state['approved_step2'] = True

    # ---------------------------------------------------------
    # Step 4: Generate Official PDF
    # ---------------------------------------------------------
    if st.session_state.get('approved_step2'):
        st.markdown("---")
        st.header("📄 Step 4: ออกเอกสารรายงาน PDF (PLA Advice)")

        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_buffer,
            pagesize=A4,
            leftMargin=40,
            rightMargin=40,
            topMargin=35,
            bottomMargin=35
        )

        styles = getSampleStyleSheet()

        style_th_address = ParagraphStyle('TH_Address', fontName='Helvetica', fontSize=8, leading=10, alignment=0)
        style_en_address = ParagraphStyle('EN_Address', fontName='Helvetica', fontSize=8, leading=10, alignment=0)
        style_company_th = ParagraphStyle('CompanyTH', fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1)
        style_title = ParagraphStyle('TitleStyle', fontName='Helvetica-Bold', fontSize=11, leading=13, alignment=1)
        style_body = ParagraphStyle('BodyStyle', fontName='Helvetica', fontSize=9, leading=12)

        elements = []

        # Header Area
        left_addr = """สำนักงานใหญ่ตั้งอยู่เลขที่<br/>
1115 ถนนพระราม 3 แขวงช่องนนทรี<br/>
เขตยานนาวา กรุงเทพฯ 10120<br/>
โทรศัพท์. 1736, 0 2239 2200<br/><br/>
เลขประจำตัวผู้เสียภาษี<br/>
0107538000533"""

        right_addr = """<b>HEAD OFFICE ADDRESS :-</b><br/>
1115 RAMA 3 ROAD, Chong Nonsi,<br/>
Yannawa, Bangkok 10120<br/>
TEL. 1736, 0 2239 2200"""

        p_left = Paragraph(left_addr, style_th_address)
        p_right = Paragraph(right_addr, style_en_address)

        logo_filename = "logo_dhipaya.jpg"
        if os.path.exists(logo_filename):
            img_logo = Image(logo_filename, width=65, height=65)
        else:
            img_logo = Paragraph("", style_body)

        center_comp = Paragraph("<b>บริษัท ทิพยประกันภัย จำกัด (มหาชน)</b><br/><font size=8.5><b>DHIPAYA INSURANCE PUBLIC COMPANY LIMITED</b></font>", style_company_th)

        header_data = [
            [p_left, img_logo, p_right],
            ['', center_comp, '']
        ]

        t_head = Table(header_data, colWidths=[160, 195, 160])
        t_head.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('ALIGN', (1,0), (1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        elements.append(t_head)
        elements.append(Spacer(1, 5))

        elements.append(Paragraph("Fire XL-2nd Layer 2025", ParagraphStyle('Sub', fontName='Helvetica', fontSize=9, alignment=2)))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph("<b>PRELIMINARY LOSS ADVICE</b>", style_title))
        elements.append(Spacer(1, 15))

        to_date_data = [
            [Paragraph("<b>To :</b> Aon Re (Thailand) Co., Ltd.", style_body), Paragraph("<b>Date :</b> 14/01/2026", ParagraphStyle('R', fontName='Helvetica', fontSize=9, alignment=2))]
        ]
        t_to_date = Table(to_date_data, colWidths=[330, 185])
        t_to_date.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        elements.append(t_to_date)
        elements.append(Spacer(1, 10))

        elements.append(Paragraph("Dear Sirs,", style_body))
        elements.append(Spacer(1, 4))
        elements.append(Paragraph("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;We regret to inform you that we have received the loss advice from the claimant as per following detail.", style_body))
        elements.append(Spacer(1, 10))

        layer_2nd = calc_layer_row("2nd Layer", 220000000, 120000000, gross_pla, net_pla)

        details_rows = [
            ("CLAIM NO.", ": Please see Attachment", "EVENT NO. : E2026-0005"),
            ("POLICY NO.", ": Please see Attachment", ""),
            ("INSURED", ": Please see Attachment", ""),
            ("LOCATION", ": Please see Attachment", ""),
            ("NATURE OF LOSS", ": Flood 2025", ""),
            ("DATE OF LOSS", ": 19/11/2025 - 30/11/2025", ""),
            ("SUM INSURED (100%)", ": Please see Attachment", ""),
            ("OUR GROSS RETENTION", ": Please see Attachment", ""),
            ("LOSS ESTIMATE", f": BHT. {gross_pla:,.2f}", ""),
            ("LOSS OF GROSS RETENTION", f": BHT. {net_pla:,.2f}", ""),
            ("EXCESS POINT", f": BHT. {layer_2nd['Excess Point']:,.2f}", ""),
            ("ESTIMATE UNDER XOL TREATY", f": BHT. {layer_2nd['Under XL']:,.2f}", ""),
            ("YOUR SHARE OF ESTIMATE", f": BHT. {layer_2nd['Aon 9%']:,.2f} (Second Layer)", "")
        ]

        table_body_data = []
        for lbl, val, extra in details_rows:
            p_lbl = Paragraph(f"<b>{lbl}</b>", style_body)
            p_val = Paragraph(val, style_body)
            p_extra = Paragraph(f"<b>{extra}</b>" if extra else "", style_body)
            table_body_data.append([p_lbl, p_val, p_extra])

        t_details = Table(table_body_data, colWidths=[175, 215, 125])
        t_details.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))

        elements.append(t_details)
        elements.append(Spacer(1, 15))

        elements.append(Paragraph("Kindly reserve the above captioned amount pending for further advice of each call from us.", style_body))
        elements.append(Spacer(1, 20))

        p_note = Paragraph("This is a computer print out, therefore no signature is required", ParagraphStyle('Note', fontName='Helvetica', fontSize=8.5, alignment=2))
        t_note = Table([[Paragraph("", style_body), p_note]], colWidths=[200, 315])
        t_note.setStyle(TableStyle([('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
        elements.append(t_note)
        elements.append(Spacer(1, 20))

        p_sign = Paragraph("Please Sign and return copy here of", style_body)
        p_handled = Paragraph("Handled by: -", ParagraphStyle('H', fontName='Helvetica', fontSize=9, alignment=2))

        t_foot = Table([[p_sign, p_handled]], colWidths=[300, 215])
        t_foot.setStyle(TableStyle([('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
        elements.append(t_foot)

        doc.build(elements)

        st.success("🎉 ออกเอกสาร PDF สำเร็จ!")
        st.download_button(
            label="📄 ดาวน์โหลดเอกสารรายงาน PDF (PLA Advice)",
            data=pdf_buffer.getvalue(),
            file_name="Final_PLA_Advice.pdf",
            mime="application/pdf"
        )
