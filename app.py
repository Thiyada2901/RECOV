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
    
    # Aggregate Settle Data
    df_set_grp = df_set_ori.groupby('เลขที่สินไหม', as_index=False).agg({
        'สาขา': 'first', 'Sub Class': 'first', 'เลขที่กรมธรรม์': 'first',
        'วันที่เกิดเหตุ': 'first', 'ชื่อผู้เอาประกัน/ชื่อบริษัทประกันภัย': 'first',
        'จังหวัด': 'first', 'ค่าสินไหม': 'sum', 'Retention': 'sum'
    }).rename(columns={
        'สาขา': 'Row Labels', 'เลขที่สินไหม': 'Claim No.', 'เลขที่กรมธรรม์': 'Policy No.',
        'วันที่เกิดเหตุ': 'Loss Date', 'ชื่อผู้เอาประกัน/ชื่อบริษัทประกันภัย': 'Insured Name',
        'ค่าสินไหม': 'Settle Gross Loss', 'Retention': 'Settle Net Loss Retention'
    })
    
    # Aggregate Reserve Data
    df_inc_grp = df_inc_ori.groupby('เลขที่สินไหม', as_index=False).agg({
        'สาขา': 'first', 'Sub Class': 'first', 'เลขที่กรมธรรม์': 'first',
        'วันที่เกิดเหตุ': 'first', 'ชื่อผู้เอาประกัน': 'first',
        'จังหวัด': 'first', 'ประมาณการค่าสินไหม': 'sum', 'Retention (By type of loss)': 'sum',
        'สถานะ': 'first'
    }).rename(columns={
        'สาขา': 'Row Labels', 'เลขที่สินไหม': 'Claim No.', 'เลขที่กรมธรรม์': 'Policy No.',
        'วันที่เกิดเหตุ': 'Loss Date', 'ชื่อผู้เอาประกัน': 'Insured Name',
        'ประมาณการค่าสินไหม': 'Reserve Gross Loss', 'Retention (By type of loss)': 'Reserve Net Loss Retention',
        'สถานะ': 'Status'
    })
    
    # Outer Merge on Claim No.
    df_bor = pd.merge(df_set_grp, df_inc_grp[['Claim No.', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']], on='Claim No.', how='outer')
    
    df_bor['Settle Gross Loss'] = df_bor['Settle Gross Loss'].fillna(0)
    df_bor['Settle Net Loss Retention'] = df_bor['Settle Net Loss Retention'].fillna(0)
    df_bor['Reserve Gross Loss'] = df_bor['Reserve Gross Loss'].fillna(0)
    df_bor['Reserve Net Loss Retention'] = df_bor['Reserve Net Loss Retention'].fillna(0)
    df_bor['Status'] = df_bor['Status'].fillna('Closed')
    
    cols_order = ['Row Labels', 'Sub Class', 'Claim No.', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด', 
                  'Settle Gross Loss', 'Settle Net Loss Retention', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']
    df_bor = df_bor[cols_order]
    
    # Totals
    s_settle_gross = df_bor['Settle Gross Loss'].sum()
    s_settle_net = df_bor['Settle Net Loss Retention'].sum()
    s_res_gross = df_bor['Reserve Gross Loss'].sum()
    s_res_net = df_bor['Reserve Net Loss Retention'].sum()
    
    st.dataframe(df_bor.head(15), use_container_width=True)
    
    # ---------------------------------------------------------
    # Styling Bordereaux Workbook
    # ---------------------------------------------------------
    bor_wb = openpyxl.Workbook()
    ws_bor = bor_wb.active
    ws_bor.title = "Details Claim"
    
    # Styles
    fill_settle = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid") # Green
    fill_reserve = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid") # Blue
    fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") # Light Grey
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
    
    # Row 1: Empty
    ws_bor.append([])
    
    # Row 2: Group Header
    ws_bor.append(['', '', '', '', '', '', '', 'Settle 12/2025', '', 'Reserve as at 31/12/2025', '', ''])
    ws_bor.merge_cells('H2:I2')
    ws_bor.merge_cells('J2:K2')
    
    ws_bor['H2'].fill = fill_settle
    ws_bor['H2'].font = font_bold
    ws_bor['H2'].alignment = Alignment(horizontal='center', vertical='center')
    
    ws_bor['J2'].fill = fill_reserve
    ws_bor['J2'].font = font_bold
    ws_bor['J2'].alignment = Alignment(horizontal='center', vertical='center')
    
    # Row 3: Column Headers
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

    # Data Rows
    current_row = 4
    for r in df_bor.itertuples(index=False):
        row_vals = list(r)
        # Format Loss Date to string YYYY-MM-DD
        if pd.notnull(row_vals[4]):
            row_vals[4] = pd.to_datetime(row_vals[4]).strftime('%Y-%m-%d')
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

    # Total Row
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

    # Auto-fit column widths
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

    # ---------------------------------------------------------
    # Step 3: Build Full Master Summary Claim (By Layer with Yellow Headers)
    # ---------------------------------------------------------
    if st.session_state.get('approved_step1'):
        st.markdown("---")
        st.header("📊 Step 3: AI สร้างตาราง Summary Claim (By Layer) ตกแต่งหัวตารางสีเหลืองสด")
        
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

        # Build workbook with openpyxl for exact colors
        sum_wb = openpyxl.Workbook()
        ws_sum = sum_wb.active
        ws_sum.title = "P&E XL"
        
        fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # Master Yellow Header
        
        ws_sum.append(["Claim Flood 2025 as at 31/12/2025"])
        ws_sum['A1'].font = Font(name="Aptos", size=12, bold=True)
        
        headers = ["Section", "Layer", "Gross 100%", "Net Loss", "Limit", "Excess Point", "", "Under XL", "IRMC 40%", "Lockton 36%", "TQR 15%", "Aon 9%", "Total"]
        
        # Section Generator Helper
        def add_summary_section(section_title, rows_data):
            ws_sum.append([]) # Blank row
            # Header Row
            h_row = [section_title if i == 0 else headers[i] for i in range(len(headers))]
            ws_sum.append(h_row)
            
            r_idx = ws_sum.max_row
            for c in range(1, 14):
                cell = ws_sum.cell(row=r_idx, column=c)
                cell.fill = fill_yellow
                cell.font = font_bold
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Sub-header Share percentages
            pct_row = ["", "", "", "", "", "", "", "", 0.40, 0.36, 0.15, 0.09, ""]
            ws_sum.append(pct_row)
            r_pct = ws_sum.max_row
            for c in range(9, 13):
                cell = ws_sum.cell(row=r_pct, column=c)
                cell.number_format = '0%'
                cell.alignment = Alignment(horizontal='right')
                
            # Data Rows
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

        # Add Sections matching Master
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
        
        # Auto-fit Column Widths
        for col in ws_sum.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_sum.column_dimensions[col_letter].width = max(max_len + 3, 14)
            
        sum_buffer = io.BytesIO()
        sum_wb.save(sum_buffer)
        
        st.download_button(
            label="📥 ดาวน์โหลด Summary Claim By Layer Master (พร้อมหัวตารางสีเหลืองสด)",
            data=sum_buffer.getvalue(),
            file_name="Summary_Claim_By_Layer_Master_Styled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        if st.button("✅ อนุมัติ Summary By Layer (Approve & Generate PDF)"):
            st.session_state['approved_step2'] = True

    # ---------------------------------------------------------
    # Step 4: Generate Official PDF (การแต่งรูปแบบเต็ม)
    # ---------------------------------------------------------
    if st.session_state.get('approved_step2'):
        st.markdown("---")
        st.header("📄 Step 4: ออกเอกสารรายงาน PDF (PLA / SLA Advice)")
        
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

        # 1. Header Area
        left_addr = """สำนักงานใหญ่ตั้งอยู่เลขที่<br/>
1115 ถนนพระราม 3 แขวงช่องนนทรี<br/>
เขตยานนาวา กรุงเทพฯ 10120<br/>
โทรศัพท์. 1736, 0 2239 2200<br/><br/>
เลขประจำตัวผู้เสียภาษี<br/>
0107538000533"""

        right_addr = """<b>HEAD OFFICE ADDRESS :-</b><br/>
115 RAMA 3 ROAD, Chong Nonsi,<br/>
Yannawa, Bangkok 10120<br/>
TEL. 1736, 0 2239 2200"""

        p_left = Paragraph(left_addr, style_th_address)
        p_right = Paragraph(right_addr, style_en_address)

        # ตรวจสอบโลโก้
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

        # Subhead & Title
        elements.append(Paragraph("Fire XL-2nd Layer 2025", ParagraphStyle('Sub', fontName='Helvetica', fontSize=9, alignment=2)))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph("<b>PRELIMINARY LOSS ADVICE</b>", style_title))
        elements.append(Spacer(1, 15))

        # To & Date
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

        # Salutation
        elements.append(Paragraph("Dear Sirs,", style_body))
        elements.append(Spacer(1, 4))
        elements.append(Paragraph("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;We regret to inform you that we have received the loss advice from the claimant as per following detail.", style_body))
        elements.append(Spacer(1, 10))

        # Dynamic Details Table (ดึงค่าคำนวณจริงจาก Step 2 & 3)
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

        # Footer Notes
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

        st.success("🎉 ระบบประมวลผลคำนวณตัวเลขและออกเอกสาร PDF ฉบับสมบูรณ์เรียบร้อยแล้ว!")
        st.download_button(
            label="📄 ดาวน์โหลดเอกสารรายงาน PDF (PLA Advice)",
            data=pdf_buffer.getvalue(),
            file_name="Final_PLA_Advice.pdf",
            mime="application/pdf"
        )
