import os
import weasyprint

def generate_pdf(output_filename="PLA_XOL_2nd_Layer_Flood_2025.pdf"):
    # ตรวจสอบว่ามีไฟล์โลโก้ในโฟลเดอร์เดียวกับ app.py หรือไม่
    logo_filename = "logo_dhipaya.jpg"  # ชื่อไฟล์โลโก้ที่คุณบันทึกไว้
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
      @page {{
        size: A4;
        margin: 1.2cm 1.5cm;
      }}
      body {{
        font-family: 'Helvetica', 'Arial', sans-serif;
        font-size: 9.5pt;
        line-height: 1.3;
        color: #000;
      }}
      .header-table {{
        width: 100%;
        border-collapse: collapse;
      }}
      .header-table td {{
        vertical-align: top;
      }}
      .left-addr {{
        width: 32%;
        font-size: 7.5pt;
        line-height: 1.2;
      }}
      .center-logo {{
        width: 36%;
        text-align: center;
      }}
      .center-logo img {{
        width: 80px;
        height: auto;
      }}
      .right-addr {{
        width: 32%;
        text-align: left;
        font-size: 7.5pt;
        line-height: 1.2;
        padding-left: 15px;
      }}
      .company-th {{
        font-size: 11pt;
        font-weight: bold;
      }}
      .company-en {{
        font-size: 8.5pt;
        font-weight: bold;
      }}
      .layer-sub {{
        text-align: right;
        font-size: 9pt;
        margin-top: 5px;
        margin-bottom: 5px;
      }}
      .doc-type {{
        text-align: center;
        margin-bottom: 20px;
      }}
      .doc-type h2 {{
        font-size: 12pt;
        margin: 0;
        letter-spacing: 0.5px;
      }}
      .recipient-table {{
        width: 100%;
        margin-bottom: 12px;
      }}
      .salutation {{
        margin-bottom: 12px;
      }}
      .details-table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 5px;
        margin-bottom: 15px;
      }}
      .details-table td {{
        padding: 2.5px 0;
        vertical-align: top;
      }}
      .col-label {{
        width: 34%;
        font-weight: bold;
      }}
      .col-colon {{
        width: 2%;
      }}
      .col-val {{
        width: 42%;
      }}
      .col-event {{
        width: 22%;
        font-weight: bold;
      }}
      .closing {{
        margin-top: 15px;
        margin-bottom: 25px;
      }}
      .computer-print {{
        text-align: right;
        font-size: 8.5pt;
        margin-bottom: 25px;
      }}
      .footer-table {{
        width: 100%;
      }}
    </style>
    </head>
    <body>

    <table class="header-table">
      <tr>
        <td class="left-addr">
          สำนักงานใหญ่ตั้งอยู่เลขที่<br>
          1115 ถนนพระราม 3 แขวงช่องนนทรี<br>
          เขตยานนาวา กรุงเทพฯ 10120<br>
          โทรศัพท์. 1736, 0 2239 2200<br><br>
          เลขประจำตัวผู้เสียภาษี<br>
          0107538000533
        </td>
        <td class="center-logo">
          <img src="{logo_filename}" alt="Dhipaya Logo"><br>
          <div style="margin-top: 5px;">
            <span class="company-th">บริษัท ทิพยประกันภัย จำกัด (มหาชน)</span><br>
            <span class="company-en">DHIPAYA INSURANCE PUBLIC COMPANY LIMITED</span>
          </div>
        </td>
        <td class="right-addr">
          <b>HEAD OFFICE ADDRESS :-</b><br>
          115 RAMA 3 ROAD, Chong Nonsi,<br>
          Yannawa, Bangkok 10120<br>
          TEL. 1736, 0 2239 2200
        </td>
      </tr>
    </table>

    <div class="layer-sub">
      Fire XL-2nd Layer 2025
    </div>

    <div class="doc-type">
      <h2>PRELIMINARY LOSS ADVICE</h2>
    </div>

    <table class="recipient-table">
      <tr>
        <td style="width: 70%;"><b>To :</b> Aon Re (Thailand) Co., Ltd.</td>
        <td style="text-align: right;"><b>Date :</b> 14/01/2026</td>
      </tr>
    </table>

    <div class="salutation">
      <p style="margin: 0;">Dear Sirs,</p>
      <p style="margin: 4px 0 0 0; text-indent: 40px;">We regret to inform you that we have received the loss advice from the claimant as per following detail.</p>
    </div>

    <table class="details-table">
      <tr>
        <td class="col-label">CLAIM NO.</td>
        <td class="col-colon">:</td>
        <td class="col-val">Please see Attachment</td>
        <td class="col-event">EVENT NO. : E2026-0005</td>
      </tr>
      <tr>
        <td class="col-label">POLICY NO.</td>
        <td class="col-colon">:</td>
        <td class="col-val">Please see Attachment</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">INSURED</td>
        <td class="col-colon">:</td>
        <td class="col-val">Please see Attachment</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">LOCATION</td>
        <td class="col-colon">:</td>
        <td class="col-val">Please see Attachment</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">NATURE OF LOSS</td>
        <td class="col-colon">:</td>
        <td class="col-val">Flood 2025</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">DATE OF LOSS</td>
        <td class="col-colon">:</td>
        <td class="col-val">19/11/2025 - 30/11/2025</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">SUM INSURED (100%)</td>
        <td class="col-colon">:</td>
        <td class="col-val">Please see Attachment</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">OUR GROSS RETENTION</td>
        <td class="col-colon">:</td>
        <td class="col-val">Please see Attachment</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">LOSS ESTIMATE</td>
        <td class="col-colon">:</td>
        <td class="col-val">BHT. 1,682,809,207.71</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">LOSS OF GROSS RETENTION</td>
        <td class="col-colon">:</td>
        <td class="col-val">BHT. 1,057,357,105.82</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">EXCESS POINT</td>
        <td class="col-colon">:</td>
        <td class="col-val">BHT. 120,000,000.00</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">ESTIMATE UNDER XOL TREATY</td>
        <td class="col-colon">:</td>
        <td class="col-val">BHT. 220,000,000.00</td>
        <td></td>
      </tr>
      <tr>
        <td class="col-label">YOUR SHARE OF ESTIMATE</td>
        <td class="col-colon">:</td>
        <td class="col-val">BHT. 19,800,000.00 (Second Layer)</td>
        <td></td>
      </tr>
    </table>

    <div class="closing">
      Kindly reserve the above captioned amount pending for further advice of each call from us.
    </div>

    <div class="computer-print">
      This is a computer print out, therefore no signature is required
    </div>

    <table class="footer-table">
      <tr>
        <td style="width: 60%; padding-left: 30px;">Please Sign and return copy here of</td>
        <td style="text-align: right;">Handled by: -</td>
      </tr>
    </table>

    </body>
    </html>
    """

    weasyprint.HTML(string=html_content, base_url=os.path.dirname(__file__)).write_pdf(output_filename)
    print(f"สร้าง PDF สำเร็จ: {output_filename}")

if __name__ == "__main__":
    generate_pdf()
