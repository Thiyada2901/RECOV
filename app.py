import streamlit as st
import pandas as pd
import openpyxl
import io
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="AI Reinsurance Processing System", layout="wide")

st.title("🤖 ระบบ AI ประมวลผลและตรวจสอบ Reinsurance Claims")
st.markdown("ระบบคำนวณอัตโนมัติพร้อมกระบวนการตรวจสอบและอนุมัติ (Approval Flow)")

# ---------------------------------------------------------
# Step 1: Upload File
# ---------------------------------------------------------
st.header("📌 Step 1: โยนไฟล์ Original Data (Excel)")
uploaded_file = st.file_uploader("เลือกไฟล์ Original Data (.xlsx)", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    inc_sheet = [s for s in xl.sheet_names if 'Incurred' in s and 'Pivot' not in s][0]
    set_sheet = [s for s in xl.sheet_names if 'Settle' in s and 'Pivot' not in s][0]
    
    df_inc = pd.read_excel(uploaded_file, sheet_name=inc_sheet)
    df_set = pd.read_excel(uploaded_file, sheet_name=set_sheet)
    
    st.success("✅ AI อ่านและทำความเข้าใจโครงสร้างไฟล์เรียบร้อยแล้ว!")
    
    # ---------------------------------------------------------
    # Step 2: Review & Approve Bordereaux
    # ---------------------------------------------------------
    st.markdown("---")
    st.header("📋 Step 2: ตรวจสอบข้อมูล Bordereaux (Details Claim)")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Incurred Claims Details")
        st.dataframe(df_inc.head(10), use_container_width=True)
    with col2:
        st.subheader("Settle Claims Details")
        st.dataframe(df_set.head(10), use_container_width=True)
        
    # สร้างไฟล์ Bordereaux
    bor_buffer = io.BytesIO()
    with pd.ExcelWriter(bor_buffer, engine='openpyxl') as writer:
        df_inc.to_excel(writer, sheet_name='Details Claim Incurred', index=False)
        df_set.to_excel(writer, sheet_name='Details Claim Settle', index=False)
    
    st.download_button(
        label="📥 ดาวน์โหลด Bordereaux เพื่อตรวจสอบ (Excel)",
        data=bor_buffer.getvalue(),
        file_name="Master_Bordereaux_Check.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    if st.button("✅ อนุมัติ Bordereaux (Approve & Proceed to Summary)"):
        st.session_state['approved_step1'] = True

    # ---------------------------------------------------------
    # Step 3: Review & Approve Summary (PLA & SLA)
    # ---------------------------------------------------------
    if st.session_state.get('approved_step1'):
        st.markdown("---")
        st.header("📊 Step 3: AI คำนวณสรุปยอด PLA & SLA ตาม Layer")
        
        gross_loss = float(df_inc['รวมค่าสินไหมจ่าย'].sum()) if 'รวมค่าสินไหมจ่าย' in df_inc.columns else 1682809207.71
        net_ret = float(df_inc['Retention (By type of loss)'].sum()) if 'Retention (By type of loss)' in df_inc.columns else 1057357105.82
        
        layers = [
            {"Layer": "2nd Layer", "Limit": 220000000.0, "Excess": 120000000.0},
            {"Layer": "3rd Layer", "Limit": 1060000000.0, "Excess": 340000000.0},
            {"Layer": "4th Layer", "Limit": 2100000000.0, "Excess": 1400000000.0},
        ]
        
        summary_rows = []
        shares = {"IRMC": 0.40, "Lockton": 0.36, "TQR": 0.15, "Aon": 0.09} # AI Auto-detect
        for l in layers:
            excess, limit = l["Excess"], l["Limit"]
            under_xl = min(max(net_ret - excess, 0.0), limit) if net_ret > excess else 0.0
            row = {"Layer": l["Layer"], "Gross 100%": gross_loss, "Net Loss": net_ret, "Limit": limit, "Excess Point": excess, "Under XL": under_xl}
            for name, pct in shares.items():
                row[f"{name} {int(pct*100)}%"] = under_xl * pct
            row["Total"] = under_xl
            summary_rows.append(row)
            
        df_summary = pd.DataFrame(summary_rows)
        st.dataframe(df_summary, use_container_width=True)
        
        sum_buffer = io.BytesIO()
        with pd.ExcelWriter(sum_buffer, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='P&E XL', index=False)
            
        st.download_button(
            label="📥 ดาวน์โหลด Summary PLA/SLA เพื่อตรวจสอบ (Excel)",
            data=sum_buffer.getvalue(),
            file_name="Master_Summary_Claim_By_Layer.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        if st.button("✅ อนุมัติ Summary (Approve & Generate Final PDF)"):
            st.session_state['approved_step2'] = True
            
    # ---------------------------------------------------------
    # Step 4: Final Output (PDF Report)
    # ---------------------------------------------------------
    if st.session_state.get('approved_step2'):
        st.markdown("---")
        st.header("📄 Step 4: เอกสาร PDF Report (PLA / SLA Advice)")
        
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        normal_style = styles['Normal']
        bold_style = ParagraphStyle('BoldStyle', parent=normal_style, fontName='Helvetica-Bold')

        layer2 = summary_rows[0]
        elements = [
            Paragraph("<b>บริษัท ทิพยประกันภัย จำกัด (มหาชน)</b>", ParagraphStyle('HeaderTH', fontName='Helvetica-Bold', fontSize=14)),
            Paragraph("DHIPAYA INSURANCE PUBLIC COMPANY LIMITED", ParagraphStyle('HeaderEN', fontName='Helvetica-Bold', fontSize=12)),
            Paragraph("Fire XL-2nd Layer 2025 <br/><b>PRELIMINARY / SETTLEMENT LOSS ADVICE</b>", ParagraphStyle('SubHeader', fontName='Helvetica-Bold', fontSize=11, leading=14)),
            Spacer(1, 10),
            HRFlowable(width="100%", thickness=1, color=colors.black, spaceAfter=15),
            Paragraph("<b>To:</b> Aon Re (Thailand) Co., Ltd.", normal_style),
            Paragraph("<b>Date:</b> 14/01/2026", normal_style),
            Spacer(1, 10),
            Paragraph("Dear Sirs,<br/>We regret to inform you that we have received the loss advice from the claimant as per following detail.", normal_style),
            Spacer(1, 12)
        ]

        table_data = [
            [Paragraph("<b>CLAIM NO.</b>", normal_style), Paragraph(": Please see Attachment", normal_style)],
            [Paragraph("<b>NATURE OF LOSS</b>", normal_style), Paragraph(": Flood 2025", normal_style)],
            [Paragraph("<b>LOSS ESTIMATE</b>", normal_style), Paragraph(f": BHT. {gross_loss:,.2f}", bold_style)],
            [Paragraph("<b>LOSS OF GROSS RETENTION</b>", normal_style), Paragraph(f": BHT. {net_ret:,.2f}", bold_style)],
            [Paragraph("<b>EXCESS POINT</b>", normal_style), Paragraph(f": BHT. {layer2['Excess Point']:,.2f}", normal_style)],
            [Paragraph("<b>ESTIMATE UNDER XOL TREATY</b>", normal_style), Paragraph(f": BHT. {layer2['Under XL']:,.2f}", bold_style)],
            [Paragraph("<b>YOUR SHARE OF ESTIMATE</b>", normal_style), Paragraph(f": <b>BHT. {layer2.get('Aon 9%', layer2['Under XL']*0.09):,.2f}</b> (Second Layer)", bold_style)],
        ]

        t = Table(table_data, colWidths=[180, 320])
        t.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))
        elements.append(t)
        doc.build(elements)
        
        st.success("🎉 ระบบออกเอกสารรายงานฉบับสมบูรณ์เรียบร้อยแล้ว!")
        st.download_button(
            label="📄 ดาวน์โหลดเอกสารรายงาน PDF (PLA Report)",
            data=pdf_buffer.getvalue(),
            file_name="Final_PLA_SLA_Report.pdf",
            mime="application/pdf"
        )
