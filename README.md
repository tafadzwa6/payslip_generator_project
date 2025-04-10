
# Payslip Generator Script

This Python script automates the process of generating payslips in PDF format and sending them to employees via email. It integrates with Gmail to send emails, and uses libraries like `pandas` for data handling, `smtplib` for email functionality, and `reportlab` for PDF creation.

## Prerequisites

Before you can run the script, ensure that you have the following:

1. **Python 3.x** installed on your machine.
2. A **Gmail account** with **2-Step Verification** enabled.
3. An **App Password** for Gmail.
4. The necessary Python libraries installed.

## Setup Instructions

### 1. Install Dependencies

First, install the required libraries using `pip`:

```bash
pip install pandas smtplib reportlab python-dotenv
```

Alternatively, you can create a `requirements.txt` file with the following content:

```
pandas
smtplib
reportlab
python-dotenv
```

And install dependencies with:

```bash
pip install -r requirements.txt
```

### 2. Set Up Gmail Credentials

To send emails, you need to configure Gmail with **2-Step Verification** and generate an **App Password**.

- [Enable 2-Step Verification on Google](https://support.google.com/accounts/answer/185833?hl=en).
- Generate an **App Password** to use with the script (don’t use your main Gmail password).

After you have generated the App Password, create a `.env` file in the root directory of your project and add your Gmail credentials like this:

```env
EMAIL_USER=your_gmail_address@gmail.com
EMAIL_PASS=your_generated_app_password
```

### 3. Prepare the Employee Data File

The script expects an Excel file `employees2.xlsx` with the following columns:

- `EMPLOYEE ID`
- `NAME`
- `EMAIL`
- `BASIC PAY`
- `ALLOWANCE`
- `DEDUCTIONS`

You can add employee records to this Excel file, and the script will process them accordingly. The `DEDUCTIONS` column can be updated manually in the script or set according to each employee ID.

### 4. Folder Setup

The script will save generated payslips in a folder called `payslips`. This folder will be created automatically if it doesn't already exist.

### 5. Script Overview

The script does the following:

1. Loads employee data from `employees2.xlsx`.
2. Generates a payslip in PDF format for each employee.
3. Sends the payslip as an email attachment to each employee's email address.

### 6. Running the Script

To run the script, ensure that you have the following files in the working directory:

- `payslip_project.py` (the script itself).
- `.env` file with Gmail credentials.
- `employees2.xlsx` containing employee data.

Then, simply run the script with:

```bash
python payslip_project.py
```

The script will process each employee, generate a payslip PDF, and send it via email.

### 7. Troubleshooting

- **Missing Libraries**: If you encounter errors related to missing libraries, run `pip install -r requirements.txt` or manually install any missing libraries.
- **Gmail Issues**: Ensure that 2-Step Verification is enabled and that you've generated an App Password.
- **Excel File Errors**: Ensure that the Excel file (`employees2.xlsx`) has the correct structure and data format.

### Example Output

Once the script completes successfully, you will see output like:

```
✅ Payslip sent to John Doe at john.doe@example.com
✅ Payslip sent to Jane Smith at jane.smith@example.com
```

Additionally, the payslips will be saved in the `payslips` folder in PDF format.

## License

This script is released under the [MIT License](LICENSE).

## Contact

For further inquiries or assistance, feel free to contact the developer at munhuharaswit@gmail.com.
