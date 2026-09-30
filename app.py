import streamlit as st
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def generate_pdf(output_filename="PLA_XOL_2nd_Layer_Flood_2025.pdf"):
    logo_filename = "logo_dhipaya.jpg"
    
    doc = SimpleDocTemplate(
        output_filename,
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
    style_company_en = ParagraphStyle('CompanyEN', fontName='Helvetica-Bold', fontSize=8.5, leading=11, alignment=1)
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

    # ใส่รูปโลโก้ ถ้าหาไฟล์ไม่พบจะเว้นไว้
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

    # Layer Subhead
    elements.append(Paragraph("Fire XL-2nd Layer 2025", ParagraphStyle('Sub', fontName='Helvetica', fontSize=9, alignment=2)))
    elements.append(Spacer(1, 10))

    # Document Title
    elements.append(Paragraph("<b>PRELIMINARY LOSS ADVICE</b>", style_title))
    elements.append(Spacer(1, 15))

    # Recipient & Date
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

    # Details Table
    details_rows = [
        ("CLAIM NO.", ": Please see Attachment", "EVENT NO. : E2026-0005"),
        ("POLICY NO.", ": Please see Attachment", ""),
        ("INSURED", ": Please see Attachment", ""),
        ("LOCATION", ": Please see Attachment", ""),
        ("NATURE OF LOSS", ": Flood 2025", ""),
        ("DATE OF LOSS", ": 19/11/2025 - 30/11/2025", ""),
        ("SUM INSURED (100%)", ": Please see Attachment", ""),
        ("OUR GROSS RETENTION", ": Please see Attachment", ""),
        ("LOSS ESTIMATE", ": BHT. 1,682,809,207.71", ""),
        ("LOSS OF GROSS RETENTION", ": BHT. 1,057,357,105.82", ""),
        ("EXCESS POINT", ": BHT. 120,000,000.00", ""),
        ("ESTIMATE UNDER XOL TREATY", ": BHT. 220,000,000.00", ""),
        ("YOUR SHARE OF ESTIMATE", ": BHT. 19,800,000.00 (Second Layer)", "")
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
    return output_filename

# UI Streamlit
st.title("ระบบสร้างเอกสาร PDF (Preliminary Loss Advice)")

if st.button("สร้างเอกสาร PDF"):
    pdf_path = generate_pdf()
    with open(pdf_path, "rb") as f:
        st.download_button(
            label="คลิกเพื่อดาวน์โหลด PDF",
            data=f,
            file_name="PLA_XOL_2nd_Layer_Flood_2025.pdf",
            mime="application/pdf"
        )
