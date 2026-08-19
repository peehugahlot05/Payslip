# Paysly - Excel to PDF Consultant Payslip Generator

Paysly is a Flask web application that automates the generation of professional PDF payslips from Excel spreadsheets. Designed for HR and finance teams, it allows you to upload an Excel sheet and download a ZIP file containing individual payslips for each consultant.

## 🌐 Live Demo
[https://payslip-mdn5.onrender.com/](https://payslip-mdn5.onrender.com/)

## 🔧 Features
- Upload `.xlsx` or `.xls` consultant payout sheets
- Automatically extract and map data fields
- Generate professional PDF payslips using a Jinja2 HTML template
- Password-protect each payslip PDF with the consultant's PAN
- Background processing with a live progress page, so large workbooks (100s of consultants) don't time out
- Download all payslips as a ZIP file
- Downloadable default Excel template to get started quickly
- Secure file handling via temporary storage
- Clean Tailwind CSS-based UI

## 🚀 Built With
- Python + Flask + Gunicorn
- Pandas + openpyxl + WeasyPrint + Jinja2 + pikepdf
- Background jobs via Python `threading`
- HTML/CSS (Tailwind)
- Docker, deployed on Render




