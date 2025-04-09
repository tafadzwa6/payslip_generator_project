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

# Load environment variables from .env file
load_dotenv()

# 🔥 PUT YOUR EMAIL HERE in the .env file as EMAIL_USER
SENDER_EMAIL = os.getenv("munhuharaswit@gmail.com")  # e.g., 'your_email@gmail.com'

# 🔥 PUT YOUR APP PASSWORD HERE in the .env file as EMAIL_PASS
SENDER_PASSWORD = os.getenv("0775774419mom")  # e.g., 'your_app_password'

SMTP_SERVER = "smtp.gmail.com"  # SMTP server for Gmail (you can change this if using a different email service)
PAYSILP_DIR = "payslips"  # Folder where the payslips will be saved

# Load employee data from Excel
def load_employees(file_path):
    try:
        data = pd.read_excel(file_path)  # Load the Excel file
        data.columns = data.columns.str.strip().str.upper()  # Normalize column names (Remove spaces and convert to uppercase)
        return data
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        return None

# Generate a beautiful and professional payslip PDF file
def generate_payslip(row):
    if not os.path.exists(PAYSILP_DIR):
        os.makedirs(PAYSILP_DIR)

    # 🔥 The file will be saved in the 'payslips' directory with the employee's ID
    filename = f"{PAYSILP_DIR}/{row['EMPLOYEE ID']}_payslip.pdf"
    
    # 🔥 Calculation of net salary (Basic Pay + Allowance - Deductions)
    net_salary = row['BASIC PAY'] + row['ALLOWANCE'] - row['DEDUCTIONS']

    # Create a PDF document for the payslip
    c = canvas.Canvas(filename, pagesize=letter)
    
    # Page border
    c.setStrokeColor(colors.black)
    c.setLineWidth(2)
    c.rect(10, 10, 580, 770)  # Outer border for the page

    # Header Gradient Background
    c.setFillColor(colors.HexColor("#4CAF50"))
    c.rect(10, 730, 580, 40, fill=1)  # Header background

    # Title in the Header
    c.setFont("Helvetica-Bold", 18)
    c.setFillColor(colors.white)
    c.drawString(20, 745, "PAYSLIP")  # Title

    # Employee Information Section
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.black)
    c.drawString(20, 700, f"Employee ID: {row['EMPLOYEE ID']}")
    c.setFont("Helvetica", 12)
    c.drawString(150, 700, f"Name: {row['NAME']}")
    c.drawString(20, 680, f"Email: {row['EMAIL']}")

    # Create Salary Breakdown Table
    data = [
        ['Description', 'Amount'],
        ['Basic Pay', f"${row['BASIC PAY']:,.2f}"],
        ['Allowances', f"${row['ALLOWANCE']:,.2f}"],
        ['Deductions', f"${row['DEDUCTIONS']:,.2f}"],
        ['Net Salary', f"${net_salary:,.2f}"]
    ]
    
    table = Table(data, colWidths=[300, 150])
    
    # Table styling
    table.setStyle(TableStyle([ 
        ('TEXTCOLOR', (0, 0), (1, 0), colors.white),  # White header
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4CAF50")),  # Header background color
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Table grid (lines between cells)
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, 0), 12)
    ]))
    
    # Table position on the page
    table.wrapOn(c, 40, 620)
    table.drawOn(c, 40, 600)

    # Footer with contact information
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.HexColor("#4CAF50"))
    c.drawString(20, 30, "For queries, contact Payroll Team at payroll@tmmotors.com")

    # Save the PDF
    c.save()

    return filename

# Send payslip via email
def send_payslip_email(row, payslip_path):
    try:
        # Creating the email message
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = row['EMAIL']  # 🔥 This is where the recipient's email is placed (Excel column 'EMAIL')
        msg['Subject'] = "Your Monthly Payslip"

        msg.attach(MIMEText("Hi there,\n\nPlease find your payslip attached.\n\nRegards,\nPayroll Team", 'plain'))

        # Attaching the payslip PDF file
        with open(payslip_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(payslip_path)}')
            msg.attach(part)

        # Sending the email via SMTP server
        server = smtplib.SMTP(SMTP_SERVER, 587)
        server.starttls()  # Secure the connection
        server.login(SENDER_EMAIL, SENDER_PASSWORD)  # 🔥 This is where you log in with your email and password
        server.send_message(msg)
        server.quit()

        print(f"✅ Payslip sent to {row['NAME']} at {row['EMAIL']}")
    except Exception as e:
        print(f"❌ Failed to send to {row['NAME']}: {e}")

# Main process
def main():
    # 🔥 Specify the path to your Excel file here
    data = load_employees("employees2.xlsx")  # Make sure 'employees2.xlsx' is in the correct directory or provide full path
    if data is None:
        return

    # Iterating through each row in the Excel file and generating/sending payslips
    for _, row in data.iterrows():
        payslip = generate_payslip(row)  # Generate the payslip
        send_payslip_email(row, payslip)  # Send the payslip via email

if __name__ == "__main__":
    main()

