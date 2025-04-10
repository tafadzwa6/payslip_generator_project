import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
import time

# Load environment variables from .env file
load_dotenv()

SENDER_EMAIL = ""
SENDER_PASSWORD = ""  

if not SENDER_EMAIL or not SENDER_PASSWORD:
    raise ValueError("Missing EMAIL_USER or EMAIL_PASS in .env file.")

SMTP_SERVER = "smtp.gmail.com"
PAYSILP_DIR = "payslips"

# Load employee data from Excel
def load_employees(file_path):
    try:
        data = pd.read_excel(file_path)
        data.columns = data.columns.str.strip().str.upper()
        return data
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        return None

# Generate payslip PDF
def generate_payslip(row):
    if not os.path.exists(PAYSILP_DIR):
        os.makedirs(PAYSILP_DIR)

    filename = f"{PAYSILP_DIR}/{row['EMPLOYEE ID']}_payslip.pdf"
    net_salary = row['BASIC PAY'] + row['ALLOWANCE'] - row['DEDUCTIONS']

    c = canvas.Canvas(filename, pagesize=letter)

    # Header with modern dark blue
    c.setFillColor(colors.HexColor("#1D3557"))
    c.rect(0, 750, 600, 50, fill=1)  # Header background
    c.setFont("Helvetica-Bold", 22)
    c.setFillColor(colors.white)
    c.drawString(40, 770, "TM MOTORS")

    # Employee Information
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.HexColor("#333333"))  # Dark Grey Text
    c.drawString(40, 720, f"Employee ID: {row['EMPLOYEE ID']}")
    c.setFont("Helvetica", 12)
    c.drawString(40, 705, f"Name: {row['NAME']}")
    c.drawString(40, 690, f"Email: {row['EMAIL']}")

    # Salary Table with soft teal header and light grey borders
    data = [
        ['Description', 'Amount'],
        ['Basic Pay', f"${row['BASIC PAY']:,.2f}"],
        ['Allowances', f"${row['ALLOWANCE']:,.2f}"],
        ['Deductions', f"${row['DEDUCTIONS']:,.2f}"],
        ['Net Salary', f"${net_salary:,.2f}"]
    ]
    
    table = Table(data, colWidths=[300, 150])
    table.setStyle(TableStyle([ 
        ('TEXTCOLOR', (0, 0), (1, 0), colors.white),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#457B9D")),  # Soft teal header
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E4E4E4")),  # Light grey borders
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, 0), 12)
    ]))

    table.wrapOn(c, 40, 600)
    table.drawOn(c, 40, 580)

    # Footer with a muted orange color
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.HexColor("#F1A208"))  # Muted Orange Footer
    c.drawString(40, 30, "Thank you for your hard work! Wishing you a great month ahead.")

    c.save()
    return filename

# Send email with attachment
def send_payslip_email(row, payslip_path):
    if not os.path.exists(payslip_path):
        print(f"❌ Payslip file not found for {row['NAME']}. Skipping email.")
        return

    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = row['EMAIL']
        msg['Subject'] = "Your Monthly Payslip"

        msg.attach(MIMEText("Hi there,\n\nPlease find your payslip attached.\n\nRegards,\nPayroll Team", 'plain'))

        with open(payslip_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(payslip_path)}')
            msg.attach(part)

        # Debug: Print email details
        print(f"Attempting to send email to {row['EMAIL']}...")

        server = smtplib.SMTP(SMTP_SERVER, 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()

        print(f"✅ Payslip sent to {row['NAME']} at {row['EMAIL']}")
    except Exception as e:
        print(f"❌ Failed to send to {row['NAME']}: {e}")

# Main logic
def main():
    data = load_employees("employees2.xlsx")
    if data is None:
        return

    for _, row in data.iterrows():
        # Update deductions manually
        if row['EMPLOYEE ID'] == 'A001':
            row['DEDUCTIONS'] = 150
        elif row['EMPLOYEE ID'] == 'A002':
            row['DEDUCTIONS'] = 200
        elif row['EMPLOYEE ID'] == 'A003':
            row['DEDUCTIONS'] = 180
        elif row['EMPLOYEE ID'] == 'A004':
            row['DEDUCTIONS'] = 320

        payslip_path = generate_payslip(row)
        send_payslip_email(row, payslip_path)

        # Add a delay of 1 second between emails
        time.sleep(1)

if __name__ == "__main__":
    main()
