import os
import sys
import threading
import uuid
from itertools import zip_longest

import pikepdf
import pandas as pd
from flask import Flask, render_template, request, send_file, redirect, url_for, jsonify, abort
import zipfile
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from num2words import num2words


# If running from PyInstaller bundle
if getattr(sys, 'frozen', False):
    os.environ['PATH'] = sys._MEIPASS + os.pathsep + os.environ['PATH']


base_path = os.environ.get("FLASK_APP_BASE", os.path.abspath("."))

app = Flask(__name__, template_folder=os.path.join(base_path, "templates"))

UPLOAD_FOLDER = os.path.join(base_path, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# In-memory job tracking for background payslip generation.
JOBS = {}
JOBS_LOCK = threading.Lock()


def load_excel(file_path):
    company_info = pd.read_excel(file_path, header=None, nrows=4)
    df = pd.read_excel(file_path, dtype=str, header=3)
    df.columns = df.columns.map(str)
    return company_info, df


def count_valid_vendors(df, columns, month_col, net_payment_col, total_payable_col, month_str, company_name, company_address):
    count = 0
    for _, row in df.iterrows():
        filled_data = standardize(
            row, columns, month_str, month_col, net_payment_col,
            total_payable_col, company_name, company_address
        )
        vendor_name = str(filled_data['vendor_name']).strip()
        if not vendor_name or vendor_name.lower() in ["total", "nan"]:
            continue
        count += 1
    return count


def find_column(columns, search):
    search = search.lower().replace(" ", "")
    for col in columns:
        col_clean = col.lower().replace(" ", "")
        if search in col_clean:
            return col
    return None


def format_joining_date(date_str):
    try:
        dt = pd.to_datetime(date_str)
        return dt.strftime('%d %B %Y')
    except Exception:
        return date_str or ""


def net_payment_words_inr(amount):
    try:
        n = float(str(amount).replace(",", "")) if amount else 0
        words = num2words(n, lang='en_IN').replace(' and', '').replace('-', ' ')
        words = words.replace(',', '')
        words = words.title()
        if words == "Zero":
            return "Rupees Zero Only"
        return f"Rupees {words} Only"
    except Exception:
        return "Rupees Zero Only"


def is_zero_or_blank(val):
    if val is None:
        return True
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return True
    try:
        return float(s.replace(",", "")) == 0
    except ValueError:
        return False


def build_payment_rows(monthly_fee, total_fee, tds_10, bonus, other_deduction,
                        travel_reimbursement, financial_pendency,
                        new_area_allowance, advance_recovery,
                        rent_deduction, notice_recovery):
    gross_items = []
    if not is_zero_or_blank(total_fee):
        gross_items.append(("Fee", monthly_fee, total_fee))
    if not is_zero_or_blank(bonus):
        gross_items.append(("Other Payments", "", bonus))
    if not is_zero_or_blank(travel_reimbursement):
        gross_items.append(("Travel Reimbursement", "", travel_reimbursement))
    if not is_zero_or_blank(new_area_allowance):
        gross_items.append(("New Area Allowance", "", new_area_allowance))

    deduction_items = []
    if not is_zero_or_blank(tds_10):
        deduction_items.append(("TDS", tds_10))
    if not is_zero_or_blank(other_deduction):
        deduction_items.append(("Other Deduction", other_deduction))
    if not is_zero_or_blank(financial_pendency):
        deduction_items.append(("Financial Pendency", financial_pendency))
    if not is_zero_or_blank(advance_recovery):
        deduction_items.append(("Advance", advance_recovery))
    if not is_zero_or_blank(rent_deduction):
        deduction_items.append(("Rent Deduction", rent_deduction))
    if not is_zero_or_blank(notice_recovery):
        deduction_items.append(("Notice Recovery", notice_recovery))

    rows = []
    for g, d in zip_longest(gross_items, deduction_items):
        rows.append({
            "g_label": g[0] if g else "",
            "g_master": g[1] if g else "",
            "g_amount": g[2] if g else "",
            "d_label": d[0] if d else "",
            "d_amount": d[1] if d else "",
        })
    return rows


def standardize(row, columns, month_str, month_col, net_payment_col, total_payable_col, company_name, company_address):
    def get_val(key):
        col = find_column(columns, key)
        return row.get(col, "") if col else ""

    total_fee = row.get(month_col, "")
    net_payment = row.get(net_payment_col, "")
    total_payable = row.get(total_payable_col, "")
    contract_start_date = format_joining_date(get_val("Contract Start Date"))
    contract_end_date = format_joining_date(get_val("Contract End Date"))

    return {
        "company_name": company_name,
        "company_address": company_address,
        "month": month_str,
        "vendor_name": get_val("Vendor's Name"),
        "contract_start_date": contract_start_date,
        "contract_end_date": contract_end_date,
        "location": get_val("Location"),
        "pay_days": get_val("Pay Days"),
        "LOP_days": get_val("LOP days") or 0,
        "vendor_code": get_val("Vendor's Code"),
        "bank_name": get_val("Consultant’s  Bank Name"),
        "bank_account": get_val("Consultant’s  Bank A/c No."),
        "pan_no": get_val("PAN No."),
        "monthly_fee": get_val("Monthly Fee"),
        "total_fee": total_fee,
        "bonus": get_val("Incentive/Bonus") or 0,
        "travel_reimbursement": get_val("Travel Reimbursement") or 0,
        "new_area_allowance": get_val("New Area Allowance") or 0,
        "tds_10": get_val("TDS@10%"),
        "other_deduction": get_val("Other Deduction") or 0,
        "financial_pendency": get_val("Financial Pendency") or 0,
        "advance_recovery": get_val("Advance Recovery") or 0,
        "rent_deduction": get_val("Rent Deduction") or 0,
        "notice_recovery": get_val("Notice Recovery") or 0,
        "total_gross": get_val("Total Gross"),
        "total_deductions": get_val("Total Deduction"),
        "net_payment": net_payment,
        "net_payment_words": net_payment_words_inr(net_payment),
        "total_payable": total_payable,
        "payment_rows": build_payment_rows(
            get_val("Monthly Fee"), total_fee, get_val("TDS@10%"),
            get_val("Incentive/Bonus") or 0, get_val("Other Deduction") or 0,
            get_val("Travel Reimbursement") or 0, get_val("Financial Pendency") or 0,
            get_val("New Area Allowance") or 0, get_val("Advance Recovery") or 0,
            get_val("Rent Deduction") or 0, get_val("Notice Recovery") or 0,
        ),
    }


def add_pdf_password(input_pdf, pan_no):
    """
    Adds password to PDF (overwrites same file safely)
    Password = PAN No. in lowercase
    """
    if not pan_no or str(pan_no).lower() == "nan":
        return  # keep PDF as-is

    password = str(pan_no).strip().lower()

    with pikepdf.open(input_pdf, allow_overwriting_input=True) as pdf:
        pdf.save(
            input_pdf,
            encryption=pikepdf.Encryption(
                user=password,
                owner=password
            )
        )


def process_job(job_id, file_path, df, columns, company_name, company_address,
                 month_col, net_payment_col, total_payable_col, month_str):
    job = JOBS[job_id]
    try:
        output_folder = os.path.join("/tmp", f"generated_pdfs_{job_id}")
        os.makedirs(output_folder, exist_ok=True)

        env = Environment(loader=FileSystemLoader(os.path.join(base_path, "templates")))
        template = env.get_template('payslip_template.html')

        vendor_counter = 0
        for _, row in df.iterrows():
            filled_data = standardize(
                row, columns, month_str, month_col, net_payment_col,
                total_payable_col, company_name, company_address
            )

            vendor_name = str(filled_data['vendor_name']).strip()
            vendor_name_lower = vendor_name.lower()
            if not vendor_name or vendor_name_lower in ["total", "nan"]:
                continue

            vendor_counter += 1
            print(f"{vendor_counter}: {vendor_name.upper()}", flush=True)

            html_out = template.render(data=filled_data)
            consultant_code = filled_data['vendor_code']

            safe_month = month_str.replace("'", "")  # avoids Windows issues
            file_name = f"{consultant_code}_Consultants Payment of {safe_month}.pdf"
            output_path = os.path.join(output_folder, file_name)

            HTML(string=html_out).write_pdf(output_path)

            pan_no = filled_data.get("pan_no", "")
            add_pdf_password(output_path, pan_no)

            job["current"] = vendor_counter

        zip_path = os.path.join("/tmp", f"Payslips_{job_id}.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for filename in os.listdir(output_folder):
                zipf.write(os.path.join(output_folder, filename), arcname=filename)

        job["zip_path"] = zip_path
        job["status"] = "done"
    except Exception as exc:
        job["error"] = str(exc)
        job["status"] = "error"
    finally:
        try:
            os.remove(file_path)
        except OSError:
            pass


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/download-template')
def download_template():
    template_path = os.path.join(base_path, "static", "downloads", "Paysly_Default_Template.xlsx")
    return send_file(template_path, as_attachment=True, download_name="Payslip_Default_Template.xlsx")


@app.route('/upload', methods=['POST'])
def upload():
    uploaded_file = request.files['excel_file']
    if uploaded_file.filename == '':
        return 'No file selected.'

    job_id = uuid.uuid4().hex
    file_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_{uploaded_file.filename}")
    uploaded_file.save(file_path)

    company_info, df = load_excel(file_path)
    columns = list(df.columns)

    company_name = str(company_info.iloc[0, 0]) if not company_info.empty else ""
    company_address = "<br>".join([
        str(company_info.iloc[i, 0]) for i in range(1, company_info.shape[0])
        if pd.notna(company_info.iloc[i, 0])
        and str(company_info.iloc[i, 0]).strip()
        and "consultants pay-out sheet" not in str(company_info.iloc[i, 0]).lower()
        and "s. no." not in str(company_info.iloc[i, 0]).lower()
    ])

    month_col = find_column(columns, "Total Fee in")
    net_payment_col = find_column(columns, "Net Payment for")
    total_payable_col = find_column(columns, "Total Payable in")

    if not month_col or not net_payment_col or not total_payable_col:
        return 'Required dynamic columns not found in Excel.'

    month_str = month_col.split("in")[-1].strip().replace("'", "")
    if len(month_str) > 2 and not month_str[-3] == "'":
        month_str = month_str[:-2] + "'" + month_str[-2:]

    total_vendors = count_valid_vendors(
        df, columns, month_col, net_payment_col, total_payable_col,
        month_str, company_name, company_address
    )

    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "processing",
            "current": 0,
            "total": total_vendors,
            "zip_path": None,
            "error": None,
        }

    thread = threading.Thread(
        target=process_job,
        args=(job_id, file_path, df, columns, company_name, company_address,
              month_col, net_payment_col, total_payable_col, month_str),
        daemon=True,
    )
    thread.start()

    return redirect(url_for('processing', job_id=job_id))


@app.route('/processing/<job_id>')
def processing(job_id):
    if job_id not in JOBS:
        abort(404)
    return render_template('processing.html', job_id=job_id)


@app.route('/status/<job_id>')
def status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"status": "error", "error": "Job not found."}), 404
    return jsonify({
        "status": job["status"],
        "current": job["current"],
        "total": job["total"],
        "error": job["error"],
    })


@app.route('/success/<job_id>')
def success(job_id):
    job = JOBS.get(job_id)
    if not job or job["status"] != "done":
        abort(404)
    return render_template('success.html', job_id=job_id)


@app.route('/download/<job_id>')
def download_zip(job_id):
    job = JOBS.get(job_id)
    if not job or job["status"] != "done" or not job["zip_path"]:
        return "No zip file found."
    return send_file(job["zip_path"], as_attachment=True)


if __name__ == "__main__":
    # Disable auto-reloader when packaged into .exe
    use_reloader = "--no-reload" not in sys.argv
    app.run(debug=False, use_reloader=use_reloader)
