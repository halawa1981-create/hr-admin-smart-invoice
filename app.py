import streamlit as st
import pandas as pd
import json
import requests
import base64
from PIL import Image
import pymupdf as fitz
import time  
import io 
import re
import os
import datetime
import urllib.parse
import plotly.express as px
import plotly.graph_objects as go
import hashlib
import secrets
import string
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── إدارة الملفات والمسارات ───
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DICT_FILE = os.path.join(BASE_DIR, "dictionary.txt")
CUSTODY_FILE = os.path.join(BASE_DIR, "custody_files.json")
USERS_FILE = os.path.join(BASE_DIR, "users.json")
TRAINING_FILE = os.path.join(BASE_DIR, "training_data.json")
APPROVED_DATASET_FILE = os.path.join(BASE_DIR, "approved_dataset.json")
TUNING_CONFIG_FILE = os.path.join(BASE_DIR, "tuning_config.json")
AUDIT_LOG_FILE = os.path.join(BASE_DIR, "audit_log.json")
API_KEY_FILE = os.path.join(BASE_DIR, "api_key.txt")
LOCAL_INVOICES_CACHE = os.path.join(BASE_DIR, "local_invoices_cache.json")

DRIVE_FOLDER_ID = "1lluuhfnz-qBNeZ0JdT6FXr7F1Y6FKDfk"

NEW_BRAND_LOGO = None
for fname in ["logo.PNG", "unmatt_logo.jpeg", "unmatt_logo.png", "anmatt_logo.png"]:
    p = os.path.join(BASE_DIR, fname)
    if os.path.exists(p):
        NEW_BRAND_LOGO = p
        break

favicon_img = Image.open(NEW_BRAND_LOGO) if NEW_BRAND_LOGO else "🏢"
st.set_page_config(
    page_title="HR Admin | Petty Cash System", 
    layout="wide", 
    page_icon=favicon_img
)

DEFAULT_DICT = "تصنيف عهد النثريات والمصروفات التشغيلية للمشاريع.\nالتأكد من استخراج الرقم الضريبي (15 رقم) وصافي الفاتورة قبل وبعد الضريبة بدقة."
DEFAULT_CUSTODIES = []
DEFAULT_APP_URL = "https://hr-app-smart-invoice-kx4r8eayr5zqp3j7faxzwq.streamlit.app"

CLOUD_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyVEwrXLkg15ZdEweZ0-DMEWBAbUYvRfnHAA4kQm67v6LMuyaRqV1o-AeRBcqkPNsM9/exec"
try:
    if hasattr(st, "secrets") and "CLOUD_WEB_APP_URL" in st.secrets:
        CLOUD_WEB_APP_URL = st.secrets["CLOUD_WEB_APP_URL"]
except Exception:
    pass

def hash_password(password: str) -> str:
    return hashlib.sha256(password.strip().encode()).hexdigest()

def generate_temp_password(length=10):
    chars = string.ascii_letters + string.digits
    return ''.join(secrets.choice(chars) for _ in range(length))

def load_audit_log():
    if os.path.exists(AUDIT_LOG_FILE):
        try:
            with open(AUDIT_LOG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list): return data
        except Exception: return []
    return []

def log_audit_event(action_type, details, user_email, user_name):
    logs = load_audit_log()
    log_entry = {
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user_email": user_email, "user_name": user_name, "action": action_type, "details": details
    }
    logs.insert(0, log_entry)
    try:
        with open(AUDIT_LOG_FILE, "w", encoding="utf-8") as f:
            json.dump(logs, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
    except Exception: pass

def load_local_invoices_cache():
    if os.path.exists(LOCAL_INVOICES_CACHE):
        try:
            with open(LOCAL_INVOICES_CACHE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list): return data
        except Exception: pass
    return []

def save_local_invoices_cache(records):
    try:
        with open(LOCAL_INVOICES_CACHE, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
        return True
    except Exception: return False

def load_tuning_config():
    default_config = {
        "company_id": "operations_management",
        "company_name": "إدارة العمليات والمشاريع",
        "system_instruction": "أنت خبير تدقيق مالي معتمد. مهمتك استخراج ومطابقة بيانات فواتير ومصروفات المشاريع بدقة متناهية بصيغة JSON.",
        "active_tuned_model": "",
        "use_tuned_model": False,
        "json_schema": {
            "invoice_no": "رقم الفاتورة المطبوع",
            "invoice_date": "YYYY-MM-DD",
            "category": "تصنيف المصروف",
            "description": "بيان المصروف بالتفصيل",
            "worker": "اسم المستلم أو العامل",
            "supplier_name": "اسم المورد أو المحل",
            "cr_number": "رقم السجل التجاري",
            "vat_number": "الرقم الضريبي (15 رقم)",
            "payment_method": "طريقة الدفع",
            "amount_before_vat": 0.0,
            "vat_amount": 0.0,
            "total_invoice": 0.0
        }
    }
    if os.path.exists(TUNING_CONFIG_FILE):
        try:
            with open(TUNING_CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                if isinstance(saved, dict):
                    default_config.update(saved)
        except Exception: pass
    return default_config

def save_tuning_config(cfg):
    try:
        with open(TUNING_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
        return True
    except Exception: return False

def load_training_db():
    if os.path.exists(TRAINING_FILE):
        try:
            with open(TRAINING_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list): return data
        except Exception: return []
    return []

def save_training_db(training_list):
    try:
        with open(TRAINING_FILE, "w", encoding="utf-8") as f:
            json.dump(training_list, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
        return True
    except Exception: return False

def load_approved_dataset():
    if os.path.exists(APPROVED_DATASET_FILE):
        try:
            with open(APPROVED_DATASET_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list): return data
        except Exception: return []
    return []

def save_approved_dataset(dataset_list):
    try:
        with open(APPROVED_DATASET_FILE, "w", encoding="utf-8") as f:
            json.dump(dataset_list, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
        return True
    except Exception: return False

def record_learned_sample(inv_no, desc, amount, category, pay_method, custody_name):
    try:
        current_db = load_training_db()
        exists = any(
            str(x.get("correct_data", {}).get("invoice_no", "")).strip() == str(inv_no).strip() and 
            abs(float(x.get("correct_data", {}).get("amount", 0.0)) - float(amount)) < 0.01 
            for x in current_db
        )
        if not exists and (str(desc).strip() or float(amount) > 0):
            new_entry = {
                "correct_data": {
                    "invoice_no": str(inv_no).strip(),
                    "description": str(desc).strip(),
                    "amount": float(amount),
                    "category": str(category).strip(),
                    "payment_method": str(pay_method).strip(),
                    "custody": str(custody_name).strip()
                },
                "learned_at": str(datetime.datetime.now()),
                "source": "live_user_correction"
            }
            current_db.insert(0, new_entry)
            save_training_db(current_db)
            if "manual_training_data" in st.session_state:
                st.session_state.manual_training_data = current_db
            return True
    except Exception: pass
    return False

def load_custodies_list():
    if os.path.exists(CUSTODY_FILE):
        try:
            with open(CUSTODY_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                if isinstance(saved, list): return saved
        except Exception: pass
    return list(DEFAULT_CUSTODIES)

def save_custodies_list_to_disk(custodies):
    try:
        with open(CUSTODY_FILE, "w", encoding="utf-8") as f:
            json.dump(custodies, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
        return True
    except Exception: return False

def load_system_dictionary():
    try:
        res = requests.get(f"{CLOUD_WEB_APP_URL}?action=get_dictionary", timeout=8)
        if res.status_code == 200:
            data = res.json()
            if data.get("status") == "success" and data.get("dictionary"):
                cloud_dict = data.get("dictionary").strip()
                with open(DICT_FILE, "w", encoding="utf-8") as f: f.write(cloud_dict)
                return cloud_dict
    except Exception: pass
    if os.path.exists(DICT_FILE):
        try:
            with open(DICT_FILE, "r", encoding="utf-8") as f: return f.read().strip()
        except Exception: return DEFAULT_DICT
    return DEFAULT_DICT

def save_system_dictionary(text):
    local_saved = False
    try:
        with open(DICT_FILE, "w", encoding="utf-8") as f: f.write(text)
        local_saved = True
    except Exception: pass
    cloud_saved = False
    try:
        payload = {"action": "save_dictionary", "dictionaryText": text}
        res = requests.post(CLOUD_WEB_APP_URL, json=payload, timeout=15)
        if res.status_code == 200 and res.json().get("status") == "success": cloud_saved = True
    except Exception: pass
    return local_saved, cloud_saved

def normalize_date(date_str):
    if not date_str or not str(date_str).strip(): return datetime.date.today().strftime("%Y-%m-%d")
    date_clean = str(date_str).strip().replace("/", "-").replace(".", "-")
    ar_to_en = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
    date_clean = date_clean.translate(ar_to_en)
    match = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', date_clean)
    if match:
        y, m, d = match.groups()
        return f"{y}-{int(m):02d}-{int(d):02d}"
    return date_clean

def optimize_image_for_upload(image, max_size=(1400, 1400), quality=80):
    img = image.copy()
    if img.mode != 'RGB': img = img.convert('RGB')
    img.thumbnail(max_size, Image.Resampling.LANCZOS)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=quality, optimize=True)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")

def load_cloud_records():
    cloud_records = []
    try:
        response = requests.get(CLOUD_WEB_APP_URL, timeout=15)
        if response.status_code == 200:
            res_json = response.json()
            if res_json.get("status") == "success":
                records = res_json.get("records", [])
                for r in records:
                    is_del = str(r.get("delete", "False")).strip().lower() in ["true", "1", "نعم"]
                    if not is_del:
                        r["delete"] = False
                        r["invoice_date"] = normalize_date(r.get("invoice_date", ""))
                        if "entry_type" not in r or not r["entry_type"]:
                            r["entry_type"] = "منصرف"
                        cloud_records.append(r)
    except Exception: pass

    local_records = load_local_invoices_cache()
    merged_map = {}
    for r in local_records:
        key = f"{r.get('serial_no')}_{r.get('invoice_no')}"
        merged_map[key] = r
    for r in cloud_records:
        key = f"{r.get('serial_no')}_{r.get('invoice_no')}"
        merged_map[key] = r
    
    final_list = list(merged_map.values())
    save_local_invoices_cache(final_list)
    return final_list

def sync_delete_to_cloud(serial_no, custody_name, user_email):
    try:
        payload = {
            "action": "delete_invoice", "serialNo": serial_no, "custody": custody_name,
            "deletedBy": user_email, "deletedAt": str(datetime.datetime.now())
        }
        res = requests.post(CLOUD_WEB_APP_URL, json=payload, timeout=15)
        return res.status_code == 200
    except Exception: return False

def save_to_cloud_storage(image, filename, row_data):
    img_str = optimize_image_for_upload(image, max_size=(1200, 1200), quality=75)
    payload = {
        "action": "upload_invoice",
        "fileName": filename,
        "mimeType": "image/jpeg",
        "fileData": img_str,
        "folderId": DRIVE_FOLDER_ID,
        "rowValues": row_data
    }
    try:
        response = requests.post(CLOUD_WEB_APP_URL, json=payload, timeout=45)
        if response.status_code == 200:
            res_json = response.json()
            if res_json.get("status") == "success":
                return res_json.get("url", ""), True
        return "", False
    except Exception:
        return "", False

def generate_managed_excel(header_data, invoice_rows, is_arabic=True):
    wb = Workbook()
    ws = wb.active
    ws.title = "Petty Cash Report"
    ws.views.sheetView[0].rightToLeft = is_arabic

    font_title = Font(name='Calibri', size=15, bold=True, color='FFFFFF')
    font_meta_lbl = Font(name='Calibri', size=10, bold=True, color='6B1D2F')
    font_meta_val = Font(name='Calibri', size=10, bold=False, color='000000')
    font_th = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
    font_td = Font(name='Calibri', size=9, bold=False, color='000000')
    
    fill_title = PatternFill(start_color='6B1D2F', end_color='6B1D2F', fill_type='solid')
    fill_th = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
    fill_meta_lbl = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
    fill_zebra = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
    fill_inward = PatternFill(start_color='DCFCE7', end_color='DCFCE7', fill_type='solid')
    
    thin_side = Side(border_style="thin", color="CBD5E1")
    border_all = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)

    ws.merge_cells('A1:R1')
    title_cell = ws['A1']
    title_cell.value = "PETTY CASH EXPENSE REPORT - تقرير مصروفات واستعاضة العهدة النثرية"
    title_cell.font = font_title
    title_cell.fill = fill_title
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 34

    if is_arabic:
        l_holder, l_title, l_empid, l_projid = "اسم صاحب العهدة:", "المسمى الوظيفي:", "الرقم الوظيفي:", "كود المشروع:"
        l_projname, l_cost, l_inward_tot, l_sec = "اسم المشروع:", "مركز التكلفة:", "إجمالي العهدة المسلمة:", "القسم / الإدارة:"
    else:
        l_holder, l_title, l_empid, l_projid = "Holder Name:", "Job Title:", "Employee ID:", "Project ID:"
        l_projname, l_cost, l_inward_tot, l_sec = "Project Name:", "Cost Center:", "Total Received Custody:", "Section / Dept:"

    meta_grid = [
        (l_holder, header_data.get("holder", ""), l_title, header_data.get("title", ""), l_empid, header_data.get("emp_id", ""), l_projid, header_data.get("proj_id", "")),
        (l_projname, header_data.get("proj_name", ""), l_cost, header_data.get("cost_center", ""), l_inward_tot, f"{float(header_data.get('inward_total', 0.0)):,.2f}", l_sec, header_data.get("sec", ""))
    ]

    for r_idx, row_vals in enumerate(meta_grid, start=3):
        ws.row_dimensions[r_idx].height = 22
        col_pairs = [('A', 'B', row_vals[0], row_vals[1]), ('E', 'F', row_vals[2], row_vals[3]),
                     ('I', 'J', row_vals[4], row_vals[5]), ('M', 'N', row_vals[6], row_vals[7])]
        for c_lbl, c_val, lbl, val in col_pairs:
            ws[f'{c_lbl}{r_idx}'] = lbl
            ws[f'{c_lbl}{r_idx}'].font = font_meta_lbl
            ws[f'{c_lbl}{r_idx}'].fill = fill_meta_lbl
            ws[f'{c_lbl}{r_idx}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'{c_lbl}{r_idx}'].border = border_all
            
            ws[f'{c_val}{r_idx}'] = val
            ws[f'{c_val}{r_idx}'].font = font_meta_val
            ws[f'{c_val}{r_idx}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'{c_val}{r_idx}'].border = border_all

    if is_arabic:
        table_headers = [
            "سريال", "النوع", "ملف العهدة", "التاريخ", "رقم المستند", "التصنيف", "البيان والتفاصيل", "المستلم / العامل",
            "اسم المورد", "السجل التجاري", "الرقم الضريبي", "طريقة الدفع", "المبلغ قبل الضريبة",
            "ضريبة 15%", "الإجمالي شامل الضريبة", "كود المشروع", "مركز التكلفة", "المسؤول"
        ]
    else:
        table_headers = [
            "Serial", "Type", "Custody File", "Date", "Invoice No", "Category", "Description", "Worker / Recipient",
            "Supplier Name", "CR No", "VAT No", "Payment Method", "Before VAT",
            "VAT 15%", "Total Amount", "Project ID", "Cost Center", "Holder"
        ]
    
    start_row = 6
    ws.row_dimensions[start_row].height = 26
    for col_num, h_text in enumerate(table_headers, start=1):
        cell = ws.cell(row=start_row, column=col_num, value=h_text)
        cell.font = font_th
        cell.fill = fill_th
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_all

    for idx, inv in enumerate(invoice_rows, start=start_row + 1):
        ws.row_dimensions[idx].height = 20
        e_type = inv.get("entry_type", "منصرف")
        row_items = [
            inv.get("display_serial", idx - start_row),
            e_type,
            inv.get("custody_name", ""),
            inv.get("invoice_date", ""),
            inv.get("invoice_no", ""),
            inv.get("category", ""),
            inv.get("description", ""),
            inv.get("worker", ""),
            inv.get("supplier_name", ""),
            inv.get("cr_number", ""),
            inv.get("vat_number", ""),
            inv.get("payment_method", ""),
            float(inv.get("amount", 0.0)),
            float(inv.get("vat", 0.0)),
            float(inv.get("total_invoice", 0.0)),
            inv.get("project_id", ""),
            inv.get("cost_center", ""),
            inv.get("holder_name", "")
        ]
        for col_idx, cell_value in enumerate(row_items, start=1):
            c = ws.cell(row=idx, column=col_idx, value=cell_value)
            c.font = font_td
            c.border = border_all
            c.alignment = Alignment(horizontal='center', vertical='center')
            if e_type == "وارد":
                c.fill = fill_inward
            elif idx % 2 == 0:
                c.fill = fill_zebra
            if col_idx in [13, 14, 15]:
                c.number_format = '#,##0.00'

    for c_idx in range(1, 19):
        col_letter = get_column_letter(c_idx)
        max_len = 12
        for r_num in range(start_row, start_row + len(invoice_rows) + 1):
            cell_val = ws.cell(row=r_num, column=c_idx).value
            if cell_val:
                max_len = max(max_len, len(str(cell_val)))
        ws.column_dimensions[col_letter].width = min(max_len + 3, 30)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─── محرك استخراج البيانات المتقدم ───
def analyze_invoice_with_gemini(image, prompt_text, api_key, entry_type="منصرف"):
    img_str = optimize_image_for_upload(image, max_size=(1400, 1400), quality=80)
    clean_key = str(api_key).strip().replace('"', '').replace("'", '').split("\n")[0].split(",")[0].strip()
    
    t_cfg = load_tuning_config()
    candidate_models = []
    if t_cfg.get("use_tuned_model") and t_cfg.get("active_tuned_model"):
        candidate_models.append(t_cfg.get("active_tuned_model").strip())

    candidate_models.extend([
        "gemini-3.5-flash-lite",
        "gemini-3.6-flash",
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash"
    ])

    if entry_type == "وارد":
        active_rules = """
المستند المرفق هو إشعار استلام عهدة أو تحويل بنكي / إيداع نقدي. المطلوب استخراج حقوله بصيغة JSON حصراً:
{
  "invoice_no": "رقم المرجع البنكي أو رقم السند",
  "invoice_date": "YYYY-MM-DD",
  "category": "وارد عهدة",
  "description": "بيان استلام السلفة أو إيداع العهدة",
  "worker": "اسم المستلم للعهدة إن وجد",
  "supplier_name": "الجهة المحولة أو البنك المصدر",
  "cr_number": "",
  "vat_number": "",
  "payment_method": "تحويل بنكي / نقدي / تحويل STC / شبكة",
  "amount_before_vat": 0.0,
  "vat_amount": 0.0,
  "total_invoice": 0.0
}
"""
    else:
        sys_inst = t_cfg.get("system_instruction", "")
        active_rules = f"""
{sys_inst}
المستند المرفق هو فاتورة منصرفات عهدة. المطلوب استخراج الحقول بدقة بصيغة JSON حصراً:
{{
  "invoice_no": "رقم الفاتورة المطبوع",
  "invoice_date": "YYYY-MM-DD",
  "category": "اختر من: مواد / إعاشة / نثريات / عمالة / محروقات / صيانة / نقل",
  "description": "بيان المصروف بالتفصيل",
  "worker": "اسم المستلم أو العامل إن وجد",
  "supplier_name": "اسم المورد أو المحل",
  "cr_number": "رقم السجل التجاري إن وجد",
  "vat_number": "الرقم الضريبي للمورد (15 رقم)",
  "payment_method": "نقدي / تحويل بنكي / تحويل STC / شبكة",
  "amount_before_vat": 0.0,
  "vat_amount": 0.0,
  "total_invoice": 0.0
}}
"""
    final_prompt = prompt_text + "\n" + active_rules
    headers = {'Content-Type': 'application/json'}
    payload = {
        "contents": [{"parts": [{"text": final_prompt}, {"inline_data": {"mime_type": "image/jpeg", "data": img_str}}]}],
        "generationConfig": {"responseMimeType": "application/json", "temperature": 0.0}
    }

    last_error_details = []
    
    for model_name in candidate_models:
        if model_name.startswith("tunedModels/"):
            url = f"https://generativelanguage.googleapis.com/v1beta/{model_name}:generateContent?key={clean_key}"
        else:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={clean_key}"
            
        for attempt in range(2):
            try:
                response = requests.post(url, headers=headers, json=payload, timeout=60)
                if response.status_code == 200:
                    res_json = response.json()
                    candidates = res_json.get('candidates', [])
                    if not candidates: continue
                    raw_text = candidates[0]['content']['parts'][0]['text']
                    match = re.search(r'\{.*\}', raw_text, re.DOTALL)
                    parsed = json.loads(match.group(0)) if match else {}
                    if "invoice_date" in parsed:
                        parsed["invoice_date"] = normalize_date(parsed["invoice_date"])
                    return parsed
                elif response.status_code in [404, 429, 503]:
                    last_error_details.append(f"[{model_name}: كود {response.status_code}]")
                    time.sleep(1.2)
                    break
                else:
                    last_error_details.append(f"[{model_name}: كود {response.status_code} - {response.text[:80]}]")
                    break
            except Exception as e:
                last_error_details.append(f"[{model_name}: استثناء {str(e)[:60]}]")
                time.sleep(1.2)
                
    err_summary = " | ".join(last_error_details) if last_error_details else "لا يوجد استجابة من الخادم"
    raise Exception(f"تعذر استدعاء الموديلات المتاحة: {err_summary}")

# ─── المستخدم المالك ───
PRIMARY_OWNER_EMAIL = "halawa1981@gmail.com"
PRIMARY_OWNER_USER = {
    "name": "م/ محمد حلاوة (System Owner)",
    "password_hash": hash_password("Admin#2026"),
    "role": "Admin",
    "allowed_custodies": "All",
    "must_change_password": False,
    "phone": "+966"
}

def load_users_db():
    users_data = {}
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                users_data = json.load(f)
        except Exception:
            users_data = {}
    users_data[PRIMARY_OWNER_EMAIL] = PRIMARY_OWNER_USER
    return users_data

def save_users_db(users_data):
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users_data, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
        return True
    except Exception: return False

# ─── تهيئة الجلسة ───
if "system_lang" not in st.session_state: st.session_state.system_lang = "العربية"
if "system_theme" not in st.session_state: st.session_state.system_theme = "Light"
if "users_db" not in st.session_state: st.session_state.users_db = load_users_db()
if "custodies_list" not in st.session_state: st.session_state.custodies_list = load_custodies_list()
if "invoices_data" not in st.session_state: st.session_state.invoices_data = load_cloud_records()
if "manual_training_data" not in st.session_state: st.session_state.manual_training_data = load_training_db()
if "approved_dataset" not in st.session_state: st.session_state.approved_dataset = load_approved_dataset()
if "tuning_config" not in st.session_state: st.session_state.tuning_config = load_tuning_config()
if "system_dictionary" not in st.session_state: st.session_state.system_dictionary = load_system_dictionary()
if "pending_invoice" not in st.session_state: st.session_state.pending_invoice = None
if "invoice_queue" not in st.session_state: st.session_state.invoice_queue = []
if "queue_index" not in st.session_state: st.session_state.queue_index = 0
if "upload_entry_mode" not in st.session_state: st.session_state.upload_entry_mode = "منصرف"
if "show_user_mgmt" not in st.session_state: st.session_state.show_user_mgmt = False
if "memory_unlocked" not in st.session_state: st.session_state.memory_unlocked = False
if "active_tab" not in st.session_state: st.session_state.active_tab = "records"
if "confirm_delete_id" not in st.session_state: st.session_state.confirm_delete_id = None
if "logged_in" not in st.session_state: st.session_state.logged_in = (st.query_params.get("session_auth") == "auth_valid_session")
if "current_user" not in st.session_state: st.session_state.current_user = PRIMARY_OWNER_EMAIL if st.session_state.logged_in else None

# الترويسة العليا
if "h_holder" not in st.session_state: st.session_state.h_holder = ""
if "h_title" not in st.session_state: st.session_state.h_title = ""
if "h_emp_id" not in st.session_state: st.session_state.h_emp_id = ""
if "h_proj_id" not in st.session_state: st.session_state.h_proj_id = ""
if "h_proj_name" not in st.session_state: st.session_state.h_proj_name = ""
if "h_cost_center" not in st.session_state: st.session_state.h_cost_center = ""
if "h_dept" not in st.session_state: st.session_state.h_dept = ""
if "h_sec" not in st.session_state: st.session_state.h_sec = ""

api_key = ""
try:
    if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets:
        api_key = str(st.secrets["GEMINI_API_KEY"]).strip().replace('"', '').replace("'", '')
except Exception: pass
if not api_key and os.path.exists(API_KEY_FILE):
    try:
        with open(API_KEY_FILE, "r", encoding="utf-8") as f: 
            api_key = f.read().strip().replace('"', '').replace("'", '')
    except Exception: pass

# ═════════════════════════════════════════════════════════════════════════
# ─── بوابة تسجيل الدخول ───
# ═════════════════════════════════════════════════════════════════════════
if not st.session_state.logged_in:
    st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

html, body, [class*="css"], .stApp {
    background-color: #050B14 !important;
    font-family: 'Inter', -apple-system, sans-serif !important;
    direction: ltr !important;
}
header[data-testid="stHeader"], footer { display: none !important; }
.block-container {
    max-width: 1560px !important;
    padding-top: 1.2rem !important;
    padding-bottom: 2rem !important;
}

.top-meta-bar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    color: #64748B;
    font-size: 13px;
    margin-bottom: 14px;
    padding: 0 4px;
}
.top-badge {
    background: #1E293B;
    color: #38BDF8;
    padding: 4px 10px;
    border-radius: 6px;
    font-weight: 800;
    font-size: 11.5px;
    letter-spacing: 0.05em;
}

div[data-testid="column"]:nth-child(2) {
    background: #08111E !important;
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    border-radius: 0 24px 24px 0 !important;
    padding: 50px 42px !important;
    box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7) !important;
    min-height: 720px !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: space-between !important;
}
div[data-testid="column"]:nth-child(2) * { color: #F8FAFC !important; }
div[data-testid="column"]:nth-child(2) h2 {
    font-size: 32px !important;
    font-weight: 800 !important;
    margin-bottom: 4px !important;
    color: #FFFFFF !important;
}
div[data-testid="column"]:nth-child(2) p.sub-title {
    color: #64748B !important;
    font-size: 14px !important;
    margin-bottom: 26px !important;
    font-weight: 500 !important;
}
div[data-testid="column"]:nth-child(2) .stTextInput label p {
    color: #94A3B8 !important;
    font-weight: 700 !important;
    font-size: 13px !important;
    margin-bottom: 4px !important;
}
div[data-testid="column"]:nth-child(2) .stTextInput input {
    border-radius: 10px !important;
    border: 1.5px solid rgba(255, 255, 255, 0.12) !important;
    background-color: #050B14 !important;
    color: #FFFFFF !important;
    padding: 12px 16px !important;
    font-size: 15px !important;
}
div[data-testid="column"]:nth-child(2) .stTextInput input:focus {
    border-color: #38BDF8 !important;
}
div[data-testid="column"]:nth-child(2) button[kind="primary"] {
    background: linear-gradient(135deg, #FF5252 0%, #E53935 100%) !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 800 !important;
    font-size: 16px !important;
    padding: 13px !important;
    box-shadow: 0 10px 25px -5px rgba(229, 57, 53, 0.45) !important;
    margin-top: 15px !important;
}
div[data-testid="column"]:nth-child(2) button[kind="primary"] p { color: #FFFFFF !important; }

.pillar-card {
    background: rgba(15, 23, 42, 0.85);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 16px;
    padding: 16px 14px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
}
.p-orange { border-top: 3.5px solid #F97316; }
.p-green { border-top: 3.5px solid #10B981; }
.p-cyan { border-top: 3.5px solid #0EA5E9; }

.graphic-box {
    width: 100%;
    height: 110px;
    border-radius: 12px;
    background: rgba(8, 14, 26, 0.95);
    border: 1px solid rgba(255, 255, 255, 0.06);
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)

    st.markdown("""
<div class="top-meta-bar">
    <div style="display: flex; align-items: center; gap: 10px;">
        <span class="top-badge">CONTECH VISUAL IDENTITY</span>
        <span style="color: #94A3B8; font-weight: 700;">Un-matt ConTech — Unified Login Experience</span>
    </div>
    <div style="font-weight: 600;">Target: Enterprise SaaS & Field Web (1920x1080)</div>
</div>
""", unsafe_allow_html=True)

    col_brand, col_auth = st.columns([1.65, 0.95], gap="small")
    
    with col_brand:
        html_left_panel = """
<div style='background:radial-gradient(circle at 15% 15%, #132742 0%, #08111E 65%, #050B14 100%);border:1px solid rgba(255,255,255,0.08);border-radius:24px 0 0 24px;padding:46px 44px;min-height:720px;display:flex;flex-direction:column;justify-content:space-between;'>
<div>
<div style='font-size:30px;font-weight:900;color:#FFFFFF;letter-spacing:0.5px;'>Un-matt <span style='color:#38BDF8;'>ConTech</span></div>
<div style='font-size:11.5px;font-weight:700;color:#94A3B8;letter-spacing:0.15em;text-transform:uppercase;margin-top:4px;'>ENTERPRISE SOLUTIONS &bull; CONSTRUCTION TECHNOLOGY</div>
<div style='display:inline-flex;align-items:center;gap:8px;background:rgba(14,165,233,0.12);border:1px solid rgba(56,189,248,0.35);padding:6px 14px;border-radius:8px;font-size:12px;font-weight:700;color:#E0F2FE;margin:22px 0 18px 0;'>
<span style='width:8px;height:8px;border-radius:50%;background:#38BDF8;box-shadow:0 0 10px #38BDF8;'></span>ENTERPRISE INVOICE MANAGEMENT
</div>
<h1 style='font-size:38px;font-weight:800;color:#FFFFFF;line-height:1.2;margin-bottom:28px;'>Turn invoices into<br><span style='color:#38BDF8;'>controlled project data</span></h1>
</div>
<div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px;'>
<div class="pillar-card p-orange">
<div class="graphic-box">
<svg width="100%" height="100%" viewBox="0 0 240 120" fill="none">
<rect x="75" y="16" width="68" height="88" rx="8" fill="#0C1322" stroke="#F97316" stroke-width="2" />
<line x1="86" y1="34" x2="132" y2="34" stroke="#475569" stroke-width="2.5" stroke-linecap="round" />
<line x1="86" y1="46" x2="124" y2="46" stroke="#475569" stroke-width="2.5" stroke-linecap="round" />
<line x1="86" y1="58" x2="132" y2="58" stroke="#475569" stroke-width="2.5" stroke-linecap="round" />
<line x1="86" y1="70" x2="116" y2="70" stroke="#475569" stroke-width="2.5" stroke-linecap="round" />
<line x1="58" y1="52" x2="160" y2="52" stroke="#38BDF8" stroke-width="2.5" stroke-dasharray="3 3" />
<polygon points="56,52 64,48 64,56" fill="#38BDF8" />
<polygon points="162,52 154,48 154,56" fill="#38BDF8" />
</svg>
</div>
<div>
<div style='font-size:10.5px;font-weight:800;color:#FB923C;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:4px;'>EXTRACTION ENGINE</div>
<div style='font-size:14px;font-weight:800;color:#F8FAFC;line-height:1.25;'>Smart Data Extraction</div>
<div style='font-size:11px;color:#94A3B8;margin-top:4px;line-height:1.35;'>Automated OCR parsing & dual textual verification.</div>
</div>
</div>
<div class="pillar-card p-green">
<div class="graphic-box">
<svg width="100%" height="100%" viewBox="0 0 240 120" fill="none">
<circle cx="78" cy="60" r="32" stroke="#1E293B" stroke-width="12" />
<circle cx="78" cy="60" r="32" stroke="#10B981" stroke-width="12" stroke-dasharray="100 200" stroke-dashoffset="-45" stroke-linecap="round" />
<line x1="106" y1="42" x2="136" y2="34" stroke="#10B981" stroke-width="1.5" stroke-dasharray="2 2" />
<rect x="136" y="24" width="76" height="20" rx="4" fill="#06281E" stroke="#10B981" stroke-width="1.2" />
<text x="174" y="38" fill="#34D399" font-size="9" font-family="-apple-system, sans-serif" font-weight="700" text-anchor="middle">CBS : 01-MAT</text>
<line x1="108" y1="68" x2="136" y2="76" stroke="#10B981" stroke-width="1.5" stroke-dasharray="2 2" />
<rect x="136" y="66" width="76" height="20" rx="4" fill="#06281E" stroke="#10B981" stroke-width="1.2" />
<text x="174" y="80" fill="#34D399" font-size="9" font-family="-apple-system, sans-serif" font-weight="700" text-anchor="middle">GL : SUB-CON</text>
</svg>
</div>
<div>
<div style='font-size:10.5px;font-weight:800;color:#34D399;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:4px;'>AUTO COST CODING</div>
<div style='font-size:14px;font-weight:800;color:#F8FAFC;line-height:1.25;'>Self Aggregation & Classification</div>
<div style='font-size:11px;color:#94A3B8;margin-top:4px;line-height:1.35;'>Dynamic vendor mapping into controlled BOQ ledger.</div>
</div>
</div>
<div class="pillar-card p-cyan">
<div class="graphic-box">
<svg width="100%" height="100%" viewBox="0 0 240 120" fill="none">
<rect x="52" y="70" width="14" height="30" rx="2" fill="#0284C7" />
<rect x="72" y="52" width="14" height="48" rx="2" fill="#0284C7" />
<rect x="92" y="36" width="14" height="64" rx="2" fill="#0284C7" />
<path d="M52 75 Q100 48 186 20" stroke="#F59E0B" stroke-width="2.5" fill="none" stroke-linecap="round" />
<circle cx="186" cy="20" r="4.5" fill="#F59E0B" />
<rect x="142" y="46" width="68" height="34" rx="6" fill="#082F49" stroke="#0284C7" stroke-width="1.2" />
<text x="176" y="59" fill="#E0F2FE" font-size="8.5" font-family="-apple-system, sans-serif" font-weight="800" text-anchor="middle">+ Cash Drift</text>
<text x="176" y="72" fill="#38BDF8" font-size="8" font-family="-apple-system, sans-serif" font-weight="700" text-anchor="middle">94% Efficiency</text>
</svg>
</div>
<div>
<div style='font-size:10.5px;font-weight:800;color:#38BDF8;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:4px;'>EXECUTIVE INTELLIGENCE</div>
<div style='font-size:14px;font-weight:800;color:#F8FAFC;line-height:1.25;'>Expenses Analysis & Actionable Insights</div>
<div style='font-size:11px;color:#94A3B8;margin-top:4px;line-height:1.35;'>Board-ready A3 KPIs, burn rate, and cost control.</div>
</div>
</div>
</div>
</div>
"""
        st.markdown(html_left_panel, unsafe_allow_html=True)

    with col_auth:
        st.markdown("""
<div>
<h2>Welcome back</h2>
<p class='sub-title'>Sign in to HR Admin Workspace</p>
</div>
""", unsafe_allow_html=True)
        
        login_email = st.text_input("Email Address", value=PRIMARY_OWNER_EMAIL, key="auth_email_field", placeholder="name@company.com")
        login_password = st.text_input("Password", type="password", value="Admin#2026", key="auth_password_field", placeholder="••••••••••••")
        
        st.markdown("""
<div style="display: flex; justify-content: space-between; align-items: center; font-size: 13px; margin: 10px 0 20px 0; color: #94A3B8;">
<label style="display: flex; align-items: center; gap: 8px; font-weight: 600; cursor: pointer;">
<input type="checkbox" checked style="accent-color: #0284C7; width: 16px; height: 16px;">
<span style="color: #94A3B8;">Remember me</span>
</label>
<a href="#" style="color: #38BDF8; font-weight: 700; text-decoration: none;">Forgot password?</a>
</div>
""", unsafe_allow_html=True)
        
        if st.button("Sign in →", type="primary", use_container_width=True):
            clean_email = login_email.strip().lower()
            clean_pwd = login_password.strip()
            user_found = next((k for k in st.session_state.users_db if k.lower() == clean_email), None)
            if user_found and st.session_state.users_db[user_found].get("password_hash") == hash_password(clean_pwd):
                st.session_state.logged_in = True
                st.session_state.current_user = user_found
                st.query_params.clear()
                st.query_params["session_auth"] = "auth_valid_session"
                st.rerun()
            else:
                st.error("Invalid email or password.")
                
        st.markdown("""
<div style="display: flex; align-items: center; justify-content: center; gap: 8px; font-size: 12px; color: #64748B; font-weight: 600; border-top: 1px solid rgba(255,255,255,0.08); padding-top: 22px; margin-top: 30px;">
<span>Enterprise Grade &bull; 256-bit SSL Security Protocol</span>
</div>
""", unsafe_allow_html=True)
    st.stop()

# ─── التحقق من الصلاحيات والترجمة ───
current_user_data = st.session_state.users_db.get(st.session_state.current_user, PRIMARY_OWNER_USER)
user_role = current_user_data.get("role", "Admin")
is_super_admin = (user_role in ["Super Admin", "Admin"])
is_company_admin = is_super_admin
is_ceo = (user_role == "CEO")
is_accountant = (user_role == "Accountant")

is_rtl = (st.session_state.system_lang == "العربية")
dir_attr = "rtl" if is_rtl else "ltr"
lang_align = "right" if is_rtl else "left"

if is_rtl:
    t = {
        "app_title": "🏢 HR Admin", "app_subtitle": "حوكمة ورقابة العهد النثرية",
        "settings": "⚙️ الإعدادات", "lang": "اللغة / Language:", "theme": "المظهر:", "theme_l": "☀️ نهاري", "theme_d": "🌙 ليلي",
        "sync_btn": "🔄 تحديث السجلات", "calc": "🔢 آلة حاسبة", "calc_btn": "احسب",
        "custody_label": "ملف العهدة الحالي:", "add_custody": "➕ إنشاء ملف عهدة جديد...", "user_mgmt": "⚙️ إدارة المستخدمين", "logout": "🚪 خروج",
        "tab_records": "📑 استمارة العهدة والتدقيق", "tab_analytics": "📊 التحليلات والمؤشرات التنفيذية", "tab_memory": "⚙️ سياسات وضوابط النظام",
        "m_total": "إجمالي المنصرف (شامل الضريبة) 🧾", "m_net": "المبلغ قبل الضريبة 💰", "m_vat": "إجمالي ضريبة القيمة المضافة ⚡", 
        "m_inward": "إجمالي العهدة المسلمة 📥", "m_balance": "صافي رصيد العهدة المتبقي ⚖️", "m_count": "عدد العمليات 📊",
        "sec_reg": "🧾 تسجيل وتدقيق عملية جديدة", "upload_tab": "📁 رفع مستندات (صور / PDF)", "paste_tab": "📋 لصق مباشر (Ctrl + V)",
        "upload_lbl": "اختر المستندات من جهازك:", "paste_lbl": "ألصق المستند هنا:", "paste_tip": "💡 اضغط هنا ثم الصق لقطة الشاشة مباشرة:",
        "start_ai": "بدء المعالجة واستخراج البيانات", "filter_title": "🔍 أدوات البحث والتصفية",
        "filter_cat": "التصنيف:", "filter_pay": "طريقة الدفع:", "filter_type": "طبيعة القيد:", "filter_period": "الفترة:",
        "entry_mode_label": "طبيعة المستندات المرفوعة:", "mode_expense": "🧾 عهدة منصرفة (فواتير مصروفات)", "mode_inward": "📥 عهدة واردة (إيداع / استلام سلفة)",
        "cats": ["الكل", "مواد", "إعاشة", "نثريات", "عمالة", "محروقات", "صيانة", "نقل", "وارد عهدة"],
        "pays": ["الكل", "نقدي", "تحويل بنكي", "تحويل STC", "شبكة"],
        "types": ["الكل", "منصرف", "وارد"],
        "periods": ["الكل", "اليوم", "هذا الأسبوع", "هذا الشهر"],
        "currency": "ر.س", "export_excel": "📥 تصدير استمارة العهدة لإكسيل (جاهز للطباعة والتسوية)",
        "conn_ok": "🟢 متصل", "conn_no": "🔴 غير متصل", "cloud_ok": "🟢 نشط",
        "h_holder_ph": "اسم صاحب العهدة", "h_title_ph": "المسمى الوظيفي", "h_empid_ph": "الرقم الوظيفي", "h_projid_ph": "كود المشروع",
        "h_projname_ph": "اسم المشروع", "h_cost_ph": "مركز التكلفة", "h_inward_lbl": "إجمالي العهدة المسلمة:", "h_sec_ph": "القسم / الإدارة"
    }
else:
    t = {
        "app_title": "🏢 HR Admin", "app_subtitle": "Petty Cash & Expense Control",
        "settings": "⚙️ Settings", "lang": "Language / اللغة:", "theme": "Theme:", "theme_l": "☀️ Light", "theme_d": "🌙 Dark",
        "sync_btn": "🔄 Sync Records", "calc": "🔢 Calculator", "calc_btn": "Calculate",
        "custody_label": "Current Custody File:", "add_custody": "➕ Create New Custody File...", "user_mgmt": "⚙️ User Management", "logout": "🚪 Logout",
        "tab_records": "📑 Petty Cash Log & Audit", "tab_analytics": "📊 Executive Analytics & KPIs", "tab_memory": "⚙️ System Policies & Controls",
        "m_total": "Total Expenses (Inc. VAT) 🧾", "m_net": "Net Before VAT 💰", "m_vat": "Total VAT ⚡", 
        "m_inward": "Total Custody Received 📥", "m_balance": "Remaining Petty Cash Balance ⚖️", "m_count": "Transactions 📊",
        "sec_reg": "🧾 Register & Audit Transaction", "upload_tab": "📁 Upload Documents (Images / PDF)", "paste_tab": "📋 Direct Paste (Ctrl + V)",
        "upload_lbl": "Choose documents from your computer:", "paste_lbl": "Paste document image here:", "paste_tip": "💡 Click here and paste screenshot directly:",
        "start_ai": "Start Processing & Extraction", "filter_title": "🔍 Search & Filter Tools",
        "filter_cat": "Category:", "filter_pay": "Payment Method:", "filter_type": "Entry Type:", "filter_period": "Period:",
        "entry_mode_label": "Upload Transaction Mode:", "mode_expense": "🧾 Outward Expense (Invoices)", "mode_inward": "📥 Inward Custody (Deposit / Advance)",
        "cats": ["All", "Materials", "Catering", "Petty Cash", "Labor", "Fuel", "Maintenance", "Transportation", "Inward Cash"],
        "pays": ["All", "Cash", "Bank Transfer", "STC Transfer", "Card / Network"],
        "types": ["All", "منصرف", "وارد"],
        "periods": ["All", "Today", "This Week", "This Month"],
        "currency": "SAR", "export_excel": "📥 Export Custody Sheet to Excel (Print Ready)",
        "conn_ok": "🟢 Connected", "conn_no": "🔴 Disconnected", "cloud_ok": "🟢 Active",
        "h_holder_ph": "Holder Name", "h_title_ph": "Job Title", "h_empid_ph": "Employee ID", "h_projid_ph": "Project ID",
        "h_projname_ph": "Project Name", "h_cost_ph": "Cost Center", "h_inward_lbl": "Total Received Custody:", "h_sec_ph": "Section / Dept"
    }

# ─── تنسيق الواجهة الداخلية ───
if st.session_state.system_theme == "Dark":
    bg_app = "#0F172A"; bg_card = "#1E293B"; text_color = "#F8FAFC"; border_color = "#334155"
    header_bg = "#5C1D2E"; accent_blue = "#4A7C9D"
else:
    bg_app = "#F8FAFC"; bg_card = "#FFFFFF"; text_color = "#0F172A"; border_color = "#E2E8F0"
    header_bg = "#6B1D2F"; accent_blue = "#3B82F6"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@400;500;600;700&family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"], .stApp {{
    font-family: 'IBM Plex Sans Arabic', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    direction: {dir_attr};
    text-align: {lang_align};
    background-color: {bg_app} !important;
    color: {text_color} !important;
    letter-spacing: -0.01em;
}}

section[data-testid="stSidebar"] {{
    background-color: {bg_card} !important;
    border-{('left' if is_rtl else 'right')}: 1px solid {border_color} !important;
    direction: {dir_attr};
}}

.sticky-header-container {{
    position: sticky !important;
    top: 0 !important;
    z-index: 9999 !important;
    background-color: {bg_app} !important;
    padding-bottom: 6px !important;
    margin-bottom: 4px !important;
    box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05) !important;
}}

.report-header-banner {{
    background: linear-gradient(135deg, {header_bg} 0%, #3D121D 100%) !important;
    color: #FFFFFF !important;
    border-radius: 8px 8px 0 0 !important;
    padding: 12px 18px !important;
    border: 1px solid {header_bg} !important;
    border-bottom: none !important;
}}

.header-inputs-card {{
    background: {bg_card} !important;
    padding: 12px 14px 6px 14px !important;
    border: 1px solid {border_color} !important;
    border-bottom: none !important;
}}

.grid-th {{
    background-color: {header_bg} !important;
    color: #FFFFFF !important;
    font-weight: 700 !important;
    font-size: 11px !important;
    padding: 9px 2px !important;
    text-align: center !important;
    border: 1px solid {border_color} !important;
    white-space: nowrap !important;
}}

.grid-td {{
    font-weight: 600 !important;
    padding: 8px 2px !important;
    border-radius: 4px !important;
    border: 1px solid {border_color} !important;
    text-align: center !important;
    margin-bottom: 4px !important;
    font-size: 11px !important;
    background-color: {bg_card} !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    white-space: nowrap !important;
}}

.grid-td-inward {{
    font-weight: 700 !important;
    padding: 8px 2px !important;
    border-radius: 4px !important;
    border: 1px solid #10B981 !important;
    text-align: center !important;
    margin-bottom: 4px !important;
    font-size: 11px !important;
    background-color: rgba(16, 185, 129, 0.12) !important;
    color: #059669 !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    white-space: nowrap !important;
}}

button[kind="primary"] {{
    font-family: 'IBM Plex Sans Arabic', 'Inter', sans-serif !important;
    background: linear-gradient(135deg, {header_bg} 0%, #4D1421 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-bottom: 3px solid {accent_blue} !important;
    font-weight: 700 !important;
    border-radius: 6px !important;
    box-shadow: 0 2px 4px rgba(0,0,0,0.08) !important;
}}

input, select, textarea {{
    font-family: 'IBM Plex Sans Arabic', 'Inter', sans-serif !important;
    font-size: 12.5px !important;
}}
</style>
""", unsafe_allow_html=True)

# ─── السايدبار ───
with st.sidebar:
    st.title(t["app_title"])
    st.caption(t["app_subtitle"])
    st.divider()

    st.header(t["settings"])
    lang_sel = st.radio(t["lang"], ["العربية", "English"], index=0 if st.session_state.system_lang == "العربية" else 1, horizontal=True)
    if lang_sel != st.session_state.system_lang:
        st.session_state.system_lang = lang_sel
        st.rerun()

    th_choice = st.radio(t["theme"], [t["theme_l"], t["theme_d"]], index=0 if st.session_state.system_theme == "Light" else 1, horizontal=True)
    sel_th = "Light" if th_choice == t["theme_l"] else "Dark"
    if sel_th != st.session_state.system_theme:
        st.session_state.system_theme = sel_th
        st.rerun()

    col_s1, col_s2 = st.columns(2)
    with col_s1: st.caption(t["conn_ok"] if api_key else t["conn_no"])
    with col_s2: st.caption(t["cloud_ok"])

    if st.button(t["sync_btn"], use_container_width=True):
        with st.spinner("Syncing data..."):
            st.session_state.invoices_data = load_cloud_records()
            save_local_invoices_cache(st.session_state.invoices_data)
            st.toast("Data synchronized!", icon="☁️")
            time.sleep(0.5); st.rerun()

    st.divider()
    with st.expander(t["calc"], expanded=False):
        calc_in = st.text_input("Expression (800*1.15):", key="side_calc")
        if st.button(t["calc_btn"], use_container_width=True):
            try:
                c_expr = re.sub(r'[^0-9\+\-\*\/\.\(\)\s]', '', calc_in)
                if c_expr.strip(): st.success(f"**{eval(c_expr):,.2f}**")
            except Exception: st.error("Error in expression")

# ─── شريط ملفات العهد والمستخدمين ───
col_cust, col_empty, col_user = st.columns([2.5, 0.3, 2.2])
can_manage = is_company_admin or is_ceo or is_super_admin

if can_manage:
    avail_custodies = list(st.session_state.custodies_list)
    if t["add_custody"] not in avail_custodies:
        avail_custodies.append(t["add_custody"])
else:
    allowed = current_user_data.get("allowed_custodies", [])
    avail_custodies = st.session_state.custodies_list if allowed == "All" else allowed

def_idx = 0
if not st.session_state.custodies_list and t["add_custody"] in avail_custodies:
    def_idx = avail_custodies.index(t["add_custody"])

selected_custody = col_cust.selectbox(t["custody_label"], avail_custodies, index=def_idx, label_visibility="collapsed")

if selected_custody == t["add_custody"] and can_manage:
    c_new1, c_new2 = col_cust.columns([2, 1])
    new_c_name = c_new1.text_input("Custody File Name:" if not is_rtl else "اسم ملف العهدة الجديد:", placeholder="مثال: عهدة التشغيل الرئيسية")
    if c_new2.button("Save" if not is_rtl else "حفظ", type="primary"):
        if new_c_name and new_c_name.strip():
            c_clean = new_c_name.strip()
            if c_clean not in st.session_state.custodies_list:
                st.session_state.custodies_list.append(c_clean)
                save_custodies_list_to_disk(st.session_state.custodies_list)
                st.rerun()

with col_user:
    st.info(f"**{current_user_data.get('name', 'م/ محمد حلاوة (System Owner)')}** &bull; `{user_role}`")
    b_u1, b_u2 = st.columns([1.2, 1])
    with b_u1:
        if can_manage and st.button(t["user_mgmt"], use_container_width=True):
            st.session_state.show_user_mgmt = not st.session_state.show_user_mgmt; st.rerun()
    with b_u2:
        if st.button(t["logout"], use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.current_user = None
            if "session_auth" in st.query_params: del st.query_params["session_auth"]
            st.rerun()

# ─── لوحة التحكم التفاعلية في المستخدمين ───
if can_manage and st.session_state.show_user_mgmt:
    st.divider()
    st.markdown("""
    <div style='background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); padding: 18px 24px; border-radius: 12px; border-left: 6px solid #FF5252; margin-bottom: 25px;'>
        <h2 style='color: #FFFFFF; margin: 0; font-size: 22px; font-weight: 800;'>👥 لوحة التحكم في المستخدمين والصلاحيات</h2>
        <p style='color: #94A3B8; margin: 5px 0 0 0; font-size: 13px;'>إدارة فرق العمل، تعيين صلاحيات العهد، وإرسال بيانات الدخول المباشرة عبر WhatsApp</p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.expander("➕ Invite New Team Member / إضافة عضو جديد ودعوته", expanded=True):
        col_u1, col_u2 = st.columns(2)
        new_u_email = col_u1.text_input("User Email :", placeholder="user@company.com", key="new_u_email")
        new_u_role = col_u2.selectbox("Role :", ["Accountant", "Admin", "CEO", "Viewer"], index=0, key="new_u_role")
        
        col_u3, col_u4 = st.columns(2)
        new_u_name = col_u3.text_input("Full Name :", placeholder="الاسم ثلاثي", key="new_u_name")
        
        all_c_choices = list(st.session_state.custodies_list)
        new_u_custodies = col_u4.multiselect("Allowed Custodies :", all_c_choices, default=all_c_choices, key="new_u_cust")
        
        col_u5, col_u6 = st.columns(2)
        app_access_url = col_u5.text_input("Platform Access URL :", value=DEFAULT_APP_URL, key="app_access_url")
        new_u_phone = col_u6.text_input("WhatsApp Phone (e.g. +966 / +20) :", placeholder="+9665xxxxxxxx أو +201xxxxxxxxx", key="new_u_phone")
        
        auto_temp_pass = st.checkbox("Generate secure temp password", value=True, key="chk_auto_pass")
        temp_pass_val = generate_temp_password(10) if auto_temp_pass else "User#2026"
        
        st.caption(f"🔑 Temporary Password generated: `{temp_pass_val}`")
        
        if st.button("🚀 Create Account & Generate Access Link", type="primary", use_container_width=True):
            clean_email = new_u_email.strip().lower()
            clean_name = new_u_name.strip()
            clean_phone = re.sub(r'[^0-9\+]', '', new_u_phone.strip())
            
            if not clean_email or not clean_name:
                st.error("⚠️ يرجى إدخال البريد الإلكتروني والاسم بالكامل.")
            else:
                st.session_state.users_db[clean_email] = {
                    "name": clean_name,
                    "password_hash": hash_password(temp_pass_val),
                    "role": new_u_role,
                    "allowed_custodies": new_u_custodies if new_u_custodies else "All",
                    "must_change_password": True,
                    "phone": clean_phone
                }
                save_users_db(st.session_state.users_db)
                log_audit_event("CREATE_USER", f"Created user {clean_email} with role {new_u_role}", current_user_data.get('name', 'Admin'), current_user_data.get('name', 'Admin'))
                st.success(f"✅ تم إنشاء حساب {clean_name} بنجاح وحفظ الصلاحيات!")
                
                cust_summary = "جميع ملفات العهد" if not new_u_custodies else "، ".join(new_u_custodies)
                wa_msg = f"""مرحباً بك {clean_name}،
تم إنشاء حسابك بنجاح في منظومة إدارة العهد النثرية:
🌐 رابط النظام: {app_access_url}
📧 اسم المستخدم (الإيميل): {clean_email}
🔑 كلمة المرور المؤقتة: {temp_pass_val}
📂 العهد المصرح بها: {cust_summary}

يرجى تسجيل الدخول والبدء في تسجيل وتدقيق العمليات."""
                
                encoded_msg = urllib.parse.quote(wa_msg)
                phone_param = clean_phone.replace("+", "")
                wa_direct_link = f"https://wa.me/{phone_param}?text={encoded_msg}" if phone_param else f"https://wa.me/?text={encoded_msg}"
                
                st.markdown(f"""
                <div style='background: rgba(16, 185, 129, 0.1); border: 1.5px solid #10B981; border-radius: 10px; padding: 18px; margin-top: 15px;'>
                    <h4 style='color: #10B981; margin: 0 0 10px 0;'>📲 إرسال الدعوة مباشرة للمستخدم:</h4>
                    <p style='margin: 0 0 12px 0; color: #334155; font-size: 13px;'>اضغط الزر أدناه لفتح تطبيق WhatsApp وإرسال بيانات الدخول والروابط بضغطة زر واحدة:</p>
                    <a href='{wa_direct_link}' target='_blank' style='display: inline-block; background-color: #25D366; color: white; padding: 10px 24px; border-radius: 8px; font-weight: bold; text-decoration: none; font-size: 14px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
                        💬 إرسال بيانات الدخول عبر WhatsApp
                    </a>
                </div>
                """, unsafe_allow_html=True)

    st.markdown("### 📋 المستخدمون النشطون في النظام")
    table_rows = []
    for em, info in list(st.session_state.users_db.items()):
        if isinstance(info, dict):
            c_val = info.get("allowed_custodies", "All")
            cust_text = "All" if c_val == "All" else (", ".join(c_val) if isinstance(c_val, list) else str(c_val))
            table_rows.append({
                "Email": em,
                "Name": info.get("name", "N/A"),
                "Role": info.get("role", "Accountant"),
                "Phone": info.get("phone", "N/A"),
                "Custodies": cust_text
            })
    
    u_df = pd.DataFrame(table_rows)
    st.dataframe(u_df, use_container_width=True)
    
    col_del_u1, col_del_u2 = st.columns([3, 1])
    deletable_users = [em for em in st.session_state.users_db.keys() if em.lower() != PRIMARY_OWNER_EMAIL.lower()]
    if deletable_users:
        user_to_delete = col_del_u1.selectbox("اختر مستخدماً لإلغاء تنشيطه:", deletable_users)
        if col_del_u2.button("🗑️ حذف المستخدم", use_container_width=True):
            del st.session_state.users_db[user_to_delete]
            save_users_db(st.session_state.users_db)
            log_audit_event("DELETE_USER", f"Removed user {user_to_delete}", current_user_data.get('name', 'Admin'), current_user_data.get('name', 'Admin'))
            st.toast("تم حذف المستخدم بنجاح!", icon="✅")
            time.sleep(0.5)
            st.rerun()
            
    if st.button("⬅️ العودة لاستمارة العهدة الرئيسية", type="secondary"):
        st.session_state.show_user_mgmt = False
        st.rerun()
        
    st.stop()

st.divider()

# ─── تبويبات المنظومة ───
tabs_list = [t["tab_records"], t["tab_analytics"]]
if is_company_admin or is_super_admin: tabs_list.append(t["tab_memory"])

t_cols = st.columns(len(tabs_list))
for idx, t_name in enumerate(tabs_list):
    k_slug = "records" if idx == 0 else ("analytics" if idx == 1 else "memory")
    with t_cols[idx]:
        if st.button(t_name, use_container_width=True, type="primary" if st.session_state.active_tab == k_slug else "secondary"):
            st.session_state.active_tab = k_slug; st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

# ─── عزل وتحليل بيانات العهدة المختارة ───
custody_records = [i for i in st.session_state.invoices_data if not i.get("delete", False) and (i.get("custody_name") == selected_custody or not i.get("custody_name"))]
expense_records = [i for i in custody_records if i.get("entry_type", "منصرف") == "منصرف"]
inward_records = [i for i in custody_records if i.get("entry_type", "منصرف") == "وارد"]

tot_expense_val = sum(float(i.get("total_invoice", 0.0)) for i in expense_records)
tot_inward_val = sum(float(i.get("total_invoice", 0.0)) for i in inward_records)
tot_before_vat_val = sum(float(i.get("amount", 0.0)) for i in expense_records)
tot_vat_val = sum(float(i.get("vat", 0.0)) for i in expense_records)
remaining_balance = tot_inward_val - tot_expense_val

if st.session_state.active_tab == "records":
    st.markdown(f"""
    <div style='display: flex; gap: 12px; margin-bottom: 25px; flex-wrap: wrap;'>
        <div style='flex: 1; min-width: 170px; background:{header_bg}; color:#FFF; padding: 12px; border-radius: 8px; text-align: center; border-bottom: 4px solid {accent_blue};'>
            <h5 style='color: {accent_blue}; margin: 0; font-size: 12px;'>{t['m_total']}</h5>
            <h3 style='color: #FFFFFF; margin: 4px 0 0 0; font-size: 18px;'>{tot_expense_val:,.2f} {t['currency']}</h3>
        </div>
        <div style='flex: 1; min-width: 170px; background:{bg_card}; padding: 12px; border-radius: 8px; text-align: center; border: 1px solid {border_color}; border-bottom: 4px solid #10B981;'>
            <h5 style='color: #10B981; margin: 0; font-size: 12px;'>{t['m_inward']}</h5>
            <h3 style='color: #10B981; margin: 4px 0 0 0; font-size: 18px;'>{tot_inward_val:,.2f} {t['currency']}</h3>
        </div>
        <div style='flex: 1; min-width: 170px; background:{bg_card}; padding: 12px; border-radius: 8px; text-align: center; border: 1px solid {border_color}; border-bottom: 4px solid #F59E0B;'>
            <h5 style='color: #F59E0B; margin: 0; font-size: 12px;'>{t['m_balance']}</h5>
            <h3 style='color: #F59E0B; margin: 4px 0 0 0; font-size: 18px;'>{remaining_balance:,.2f} {t['currency']}</h3>
        </div>
        <div style='flex: 1; min-width: 170px; background:{bg_card}; padding: 12px; border-radius: 8px; text-align: center; border: 1px solid {border_color}; border-bottom: 4px solid #64748B;'>
            <h5 style='color: #64748B; margin: 0; font-size: 12px;'>{t['m_net']}</h5>
            <h3 style='color: {text_color}; margin: 4px 0 0 0; font-size: 18px;'>{tot_before_vat_val:,.2f} {t['currency']}</h3>
        </div>
        <div style='flex: 1; min-width: 170px; background:{bg_card}; padding: 12px; border-radius: 8px; text-align: center; border: 1px solid {border_color}; border-bottom: 4px solid {accent_blue};'>
            <h5 style='color: {accent_blue}; margin: 0; font-size: 12px;'>{t['m_count']}</h5>
            <h3 style='color: {accent_blue}; margin: 4px 0 0 0; font-size: 18px;'>{len(custody_records)}</h3>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if len(st.session_state.invoice_queue) == 0:
        st.markdown(f"""
        <div style="background-color: {header_bg}; color: white; padding: 12px 25px; border-radius: 8px; font-size: 18px; font-weight: 800; margin-bottom: 15px; border-right: 6px solid {accent_blue};">
            {t['sec_reg']} &bull; [{selected_custody}]
        </div>
        """, unsafe_allow_html=True)
        
        c_mode_choice = st.radio(
            t["entry_mode_label"],
            [t["mode_expense"], t["mode_inward"]],
            index=0 if st.session_state.upload_entry_mode == "منصرف" else 1,
            horizontal=True
        )
        st.session_state.upload_entry_mode = "منصرف" if c_mode_choice == t["mode_expense"] else "وارد"
        
        up_tab1, up_tab2 = st.tabs([t["upload_tab"], t["paste_tab"]])
        main_files = []
        with up_tab1:
            uploaded = st.file_uploader(t["upload_lbl"], type=['jpg', 'png', 'jpeg', 'pdf'], accept_multiple_files=True)
            if uploaded: main_files.extend(uploaded)
        with up_tab2:
            st.caption(t["paste_tip"])
            pasted = st.file_uploader(t["paste_lbl"], type=['png', 'jpg', 'jpeg'], key="paste_hr_in")
            if pasted: main_files.append(pasted)
            
        if st.button(t["start_ai"], type="primary"):
            if not api_key: st.error("⚠️ API Key not connected / المفتاح غير متصل.")
            elif not main_files: st.warning("⚠️ Please attach documents first / يرجى إرفاق المستندات أولاً.")
            else:
                with st.spinner("Processing files..."):
                    st.session_state.invoice_queue = []
                    for file in main_files:
                        if file.name.lower().endswith('.pdf'):
                            doc = fitz.open(stream=file.read(), filetype="pdf")
                            for i in range(len(doc)):
                                pix = doc.load_page(i).get_pixmap(dpi=130)
                                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                                st.session_state.invoice_queue.append({
                                    "image": img, "name": f"{file.name}_p{i+1}", 
                                    "entry_type": st.session_state.upload_entry_mode
                                })
                        else:
                            st.session_state.invoice_queue.append({
                                "image": Image.open(file), "name": file.name,
                                "entry_type": st.session_state.upload_entry_mode
                            })
                    st.session_state.queue_index = 0
                    st.session_state.pending_invoice = None
                    st.rerun()

    elif st.session_state.queue_index < len(st.session_state.invoice_queue):
        curr_idx = st.session_state.queue_index
        tot_q = len(st.session_state.invoice_queue)
        curr_item = st.session_state.invoice_queue[curr_idx]
        current_entry_type = curr_item.get("entry_type", "منصرف")
        
        type_badge = "📥 عهدة واردة" if current_entry_type == "وارد" else "🧾 عهدة منصرفة"
        st.info(f"⏳ **Reviewing ({type_badge}):** ({curr_idx + 1}) of ({tot_q}) - `{curr_item['name']}`")
        if st.session_state.pending_invoice is None:
            with st.spinner("جاري استخراج وتدقيق بيانات المستند..."):
                prompt = "استخرج بيانات المستند الرسمية كاملة بدقة."
                try:
                    time.sleep(0.4)
                    st.session_state.pending_invoice = analyze_invoice_with_gemini(curr_item["image"], prompt, api_key, entry_type=current_entry_type)
                    st.rerun()
                except Exception as e:
                    st.error(f"⚠️ {e}")
                    if st.button("Skip document" if not is_rtl else "تخطي المستند"):
                        st.session_state.queue_index += 1
                        st.session_state.pending_invoice = None
                        st.rerun()
                    st.stop()
                    
        d = st.session_state.pending_invoice
        col_img, col_data = st.columns([1, 2.2])
        with col_img: st.image(curr_item["image"], use_container_width=True)
        with col_data:
            entry_type_choice = st.radio(
                "طبيعة القيد / Transaction Type:" if is_rtl else "Transaction Type:",
                ["منصرف", "وارد"],
                index=0 if current_entry_type == "منصرف" else 1,
                horizontal=True
            )
            
            c1, c2, c3 = st.columns(3)
            inv_no = c1.text_input("Invoice / Ref No:" if not is_rtl else "رقم الفاتورة / المستند:", value=str(d.get("invoice_no", "")))
            inv_date = c2.text_input("Date (YYYY-MM-DD):" if not is_rtl else "التاريخ (YYYY-MM-DD):", value=normalize_date(d.get("invoice_date", "")))
            
            cat_list = ["Materials", "Catering", "Petty Cash", "Labor", "Fuel", "Maintenance", "Transportation", "Inward Cash"] if not is_rtl else ["مواد", "إعاشة", "نثريات", "عمالة", "محروقات", "صيانة", "نقل", "وارد عهدة"]
            def_cat_idx = 7 if entry_type_choice == "وارد" else 0
            cat = c3.selectbox("Category:" if not is_rtl else "تصنيف المصروف:", cat_list, index=def_cat_idx)
            
            c4, c5, c6 = st.columns(3)
            supp = c4.text_input("Supplier / Source:" if not is_rtl else "اسم المورد / جهة الإيداع:", value=str(d.get("supplier_name", "")))
            vat_no = c5.text_input("VAT No (15 digits):" if not is_rtl else "الرقم الضريبي (15 رقم):", value=str(d.get("vat_number", "")))
            cr_no = c6.text_input("CR No:" if not is_rtl else "السجل التجاري (CR):", value=str(d.get("cr_number", "")))
            
            desc = st.text_input("Description:" if not is_rtl else "البيان بالتفصيل:", value=str(d.get("description", "")))
            worker = st.text_input("Recipient / Worker:" if not is_rtl else "اسم المستلم أو العامل:", value=str(d.get("worker", "")))
            
            c7, c8, c9, c10 = st.columns(4)
            if entry_type_choice == "وارد":
                amt_before = c7.number_input("Amount Received:" if not is_rtl else "المبلغ المستلم:", value=float(d.get("amount_before_vat", d.get("amount", 0.0))))
                vat_amt = c8.number_input("VAT Amount:" if not is_rtl else "قيمة الضريبة:", value=0.0)
                tot_amt = c9.number_input("Total Received:" if not is_rtl else "إجمالي الوارد للعهدة:", value=float(d.get("total_invoice", amt_before)))
            else:
                amt_before = c7.number_input("Before VAT:" if not is_rtl else "المبلغ قبل الضريبة:", value=float(d.get("amount_before_vat", d.get("amount", 0.0))))
                vat_amt = c8.number_input("VAT (15%):" if not is_rtl else "قيمة الضريبة (15%):", value=float(d.get("vat_amount", d.get("vat", 0.0))))
                tot_amt = c9.number_input("Total Amount:" if not is_rtl else "الإجمالي شامل الضريبة:", value=float(d.get("total_invoice", amt_before + vat_amt)))
            
            pay_opt = ["Cash", "Bank Transfer", "STC Transfer", "Card / Network"] if not is_rtl else ["نقدي", "تحويل بنكي", "تحويل STC", "شبكة"]
            pay_m = c10.selectbox("Payment Method:" if not is_rtl else "طريقة الدفع:", pay_opt, index=0)
            
            clean_inv_no = str(inv_no).strip()
            duplicate_matches = []
            if clean_inv_no:
                for past_inv in custody_records:
                    if str(past_inv.get("invoice_no", "")).strip() == clean_inv_no:
                        duplicate_matches.append(past_inv)

            is_duplicate = len(duplicate_matches) > 0
            if is_duplicate:
                msg_dup = f"🚨 **Warning: Document ({clean_inv_no}) already registered ({len(duplicate_matches)}) times!**" if not is_rtl else f"🚨 **تنبيه رقابي حرج: المستند رقم ({clean_inv_no}) مسجل مسبقاً في النظام بعدد ({len(duplicate_matches)}) مرة!**"
                st.error(msg_dup)
                dup_details = []
                for dm in duplicate_matches:
                    dup_details.append(f"• Serial: **{dm.get('serial_no')}** | Type: **{dm.get('entry_type', 'منصرف')}** | Date: **{dm.get('invoice_date')}** | Total: **{float(dm.get('total_invoice',0)):,.2f} {t['currency']}**")
                st.markdown("<br>".join(dup_details), unsafe_allow_html=True)

            # ─── أزرار الإجراءات الثلاثية ───
            st.divider()
            b_save, b_attach, b_skip = st.columns([1.5, 1.5, 1])
            with b_save:
                if st.button("✅ اعتماد وحفظ في الشيت", type="primary", use_container_width=True):
                    with st.spinner("جاري حفظ العملية في النظام..."):
                        all_active = [x for x in st.session_state.invoices_data if not x.get("delete", False)]
                        new_serial = len(all_active) + 1
                        row_payload = [
                            new_serial, entry_type_choice, selected_custody, inv_date, inv_no, cat, desc, worker, supp, cr_no, vat_no, pay_m, amt_before, vat_amt, tot_amt, st.session_state.h_proj_id, st.session_state.h_cost_center, st.session_state.h_holder, ""
                        ]
                        link, ok = save_to_cloud_storage(curr_item["image"], curr_item["name"], row_payload)
                        new_record = {
                            "serial_no": new_serial, "entry_type": entry_type_choice, "custody_name": selected_custody,
                            "invoice_date": inv_date, "invoice_no": inv_no, "category": cat, "description": desc, "worker": worker,
                            "supplier_name": supp, "cr_number": cr_no, "vat_number": vat_no, "payment_method": pay_m,
                            "amount": amt_before, "vat": vat_amt, "total_invoice": tot_amt, "project_id": st.session_state.h_proj_id,
                            "cost_center": st.session_state.h_cost_center, "holder_name": st.session_state.h_holder, "delete": False, "drive_link": link
                        }
                        st.session_state.invoices_data.append(new_record)
                        save_local_invoices_cache(st.session_state.invoices_data)
                        record_learned_sample(inv_no, desc, tot_amt, cat, pay_m, selected_custody)

                        approved_ds = load_approved_dataset()
                        approved_ds.insert(0, {
                            "timestamp": str(datetime.datetime.now()),
                            "company_id": st.session_state.tuning_config.get("company_id", "operations_management"),
                            "input_doc": curr_item["name"],
                            "ground_truth": {
                                "invoice_no": inv_no, "invoice_date": inv_date, "category": cat, "description": desc,
                                "worker": worker, "supplier_name": supp, "cr_number": cr_no, "vat_number": vat_no,
                                "payment_method": pay_m, "amount_before_vat": amt_before, "vat_amount": vat_amt, "total_invoice": tot_amt
                            }
                        })
                        save_approved_dataset(approved_ds)
                        st.session_state.approved_dataset = approved_ds

                        time.sleep(0.5)
                        st.session_state.pending_invoice = None
                        st.session_state.queue_index += 1
                        st.rerun()

            with b_attach:
                if st.button("📎 إرفاق كمرفق", use_container_width=True):
                    with st.spinner("جاري ربط المستند كمرفق بالسجل السابق..."):
                        if custody_records:
                            last_rec = custody_records[-1]
                            row_payload = [
                                last_rec.get('serial_no'), entry_type_choice, selected_custody, inv_date, inv_no, cat, desc, worker, supp, cr_no, vat_no, pay_m, amt_before, vat_amt, tot_amt, st.session_state.h_proj_id, st.session_state.h_cost_center, st.session_state.h_holder, ""
                            ]
                            link, ok = save_to_cloud_storage(curr_item["image"], f"Attachment_{curr_item['name']}", row_payload)
                            if not last_rec.get("drive_link"):
                                last_rec["drive_link"] = link
                            else:
                                last_rec["drive_link"] += f" | {link}"
                            save_local_invoices_cache(st.session_state.invoices_data)
                        time.sleep(0.5)
                        st.session_state.pending_invoice = None
                        st.session_state.queue_index += 1
                        st.rerun()

            with b_skip:
                if st.button("❌ استبعاد", use_container_width=True):
                    st.session_state.pending_invoice = None
                    st.session_state.queue_index += 1
                    st.rerun()
    else:
        st.success("🎉 All documents processed successfully!" if not is_rtl else "🎉 تم تدقيق كافة العمليات بنجاح!")
        if st.button("🔄 Process New Batch" if not is_rtl else "🔄 إدخال دفعة مستندات جديدة"):
            st.session_state.invoice_queue = []
            st.session_state.queue_index = 0
            st.rerun()

    st.write("<br>", unsafe_allow_html=True)

    # ─── الفلاتر ───
    st.markdown(f"""
    <div style="background-color: {header_bg}; color: white; padding: 12px 25px; border-radius: 8px; font-size: 18px; font-weight: 800; margin-bottom: 15px; border-right: 6px solid {accent_blue};">
        {t['filter_title']}
    </div>
    """, unsafe_allow_html=True)
    fc1, fc2, fc3, fc4 = st.columns(4)
    f_cat = fc1.selectbox(t["filter_cat"], t["cats"])
    f_pay = fc2.selectbox(t["filter_pay"], t["pays"])
    f_type = fc3.selectbox(t["filter_type"], t["types"])
    f_date = fc4.selectbox(t["filter_period"], t["periods"])

    filtered_invoices = []
    for item in custody_records:
        if f_cat not in ["All", "الكل"] and item.get("category") != f_cat: continue
        if f_pay not in ["All", "الكل"] and item.get("payment_method") != f_pay: continue
        if f_type not in ["All", "الكل"] and item.get("entry_type", "منصرف") != f_type: continue
        filtered_invoices.append(item)

    # ترقيم تسلسلي حي ومتناسق دائماً (1، 2، 3، 4...)
    for idx_seq, item in enumerate(filtered_invoices, start=1):
        item["display_serial"] = idx_seq

    if st.session_state.confirm_delete_id is not None:
        target = next((x for x in custody_records if x.get("serial_no") == st.session_state.confirm_delete_id), None)
        del_warn = f"⚠️ Are you sure you want to delete entry ({st.session_state.confirm_delete_id}) with amount ({target.get('total_invoice',0):,.2f} {t['currency']})?" if not is_rtl else f"⚠️ هل أنت متأكد تماماً من حذف العملية رقم ({st.session_state.confirm_delete_id}) بمبلغ ({target.get('total_invoice',0):,.2f} {t['currency']})؟"
        st.warning(del_warn)
        yd, nd, _ = st.columns([1.5, 1.5, 4])
        with yd:
            if st.button("Yes, Delete 🗑️" if not is_rtl else "نعم، احذف نهائياً 🗑️", type="primary", use_container_width=True):
                sync_delete_to_cloud(st.session_state.confirm_delete_id, selected_custody, current_user_data.get('name', 'Admin'))
                for inv in st.session_state.invoices_data:
                    if inv.get("serial_no") == st.session_state.confirm_delete_id: inv["delete"] = True
                save_local_invoices_cache(st.session_state.invoices_data)
                log_audit_event("DELETE", f"Deleted serial #{st.session_state.confirm_delete_id}", current_user_data.get('name', 'Admin'), current_user_data.get('name', 'Admin'))
                st.session_state.confirm_delete_id = None
                st.toast("Deleted successfully!", icon="✅")
                time.sleep(0.5); st.rerun()
        with nd:
            if st.button("Cancel ❌" if not is_rtl else "إلغاء ❌", use_container_width=True):
                st.session_state.confirm_delete_id = None; st.rerun()

    # ═════════════════════════════════════════════════════════════════════════
    # ─── الكتلة المجمّدة: ترويسة تفاعلية متناسقة + جدول الـ 18 عموداً ───
    # ═════════════════════════════════════════════════════════════════════════
    st.markdown(f"""
    <div class="sticky-header-container">
        <div class="report-header-banner">
            <h3 style="margin: 0; font-size: 19px; font-weight: 900; letter-spacing: 0.5px; text-align: center;">PETTY CASH EXPENSE REPORT &bull; {selected_custody}</h3>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown("<div class='header-inputs-card'>", unsafe_allow_html=True)
        rh1, rh2, rh3, rh4 = st.columns(4)
        st.session_state.h_holder = rh1.text_input("Advanced Holder Name:", value=st.session_state.h_holder, placeholder=t["h_holder_ph"])
        st.session_state.h_title = rh2.text_input("Title:", value=st.session_state.h_title, placeholder=t["h_title_ph"])
        st.session_state.h_emp_id = rh3.text_input("Employee ID:", value=st.session_state.h_emp_id, placeholder=t["h_empid_ph"])
        st.session_state.h_proj_id = rh4.text_input("Project ID:", value=st.session_state.h_proj_id, placeholder=t["h_projid_ph"])

        rh5, rh6, rh7, rh8 = st.columns(4)
        st.session_state.h_proj_name = rh5.text_input("Project Name:", value=st.session_state.h_proj_name, placeholder=t["h_projname_ph"])
        st.session_state.h_cost_center = rh6.text_input("Cost Center:", value=st.session_state.h_cost_center, placeholder=t["h_cost_ph"])
        rh7.text_input(t["h_inward_lbl"], value=f"{tot_inward_val:,.2f} {t['currency']}", disabled=True)
        st.session_state.h_sec = rh8.text_input("Section / Dept:", value=st.session_state.h_sec, placeholder=t["h_sec_ph"])
        st.markdown("</div>", unsafe_allow_html=True)

    meta_headers_payload = {
        "holder": st.session_state.h_holder, "title": st.session_state.h_title,
        "emp_id": st.session_state.h_emp_id, "proj_id": st.session_state.h_proj_id,
        "proj_name": st.session_state.h_proj_name, "cost_center": st.session_state.h_cost_center,
        "inward_total": tot_inward_val, "sec": st.session_state.h_sec
    }
    excel_stream = generate_managed_excel(meta_headers_payload, filtered_invoices, is_arabic=is_rtl)
    st.download_button(
        label=t["export_excel"],
        data=excel_stream,
        file_name=f"Custody_{selected_custody}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # ضبط أوزان وعروض الأعمدة مع توسيع عمود الحذف الأخير ليظهر زر 🗑️ بوضوح
    cols_weights = [0.55, 0.65, 0.85, 0.8, 0.8, 0.75, 1.4, 0.8, 0.95, 0.75, 0.85, 0.75, 0.8, 0.7, 0.85, 0.65, 0.65, 0.75]
    headers_list = ["سريال", "النوع", "ملف العهدة", "التاريخ", "رقم الفاتورة", "التصنيف", "البيان", "المستلم", "المورد", "السجل", "الرقم الضريبي", "الدفع", "قبل الضريبة", "الضريبة", "الإجمالي", "المركز", "المسؤول", "حذف"] if is_rtl else [
        "Serial", "Type", "Custody File", "Date", "Ref No", "Category", "Description", "Worker", "Supplier", "CR No", "VAT No", "Payment", "Before VAT", "VAT", "Total", "Cost Center", "Holder", "Delete"
    ]

    th_cols = st.columns(cols_weights)
    for idx, h_name in enumerate(headers_list):
        th_cols[idx].markdown(f"<div class='grid-th'>{h_name}</div>", unsafe_allow_html=True)

    if filtered_invoices:
        for idx, r in enumerate(filtered_invoices):
            e_type = r.get("entry_type", "منصرف")
            is_inward = (e_type == "وارد")
            c_class = "grid-td-inward" if is_inward else "grid-td"
            
            td_cols = st.columns(cols_weights)
            td_cols[0].markdown(f"<div class='{c_class}'>{r.get('display_serial', idx+1)}</div>", unsafe_allow_html=True)
            td_cols[1].markdown(f"<div class='{c_class}'>{'📥 وارد' if is_inward else '🧾 منصرف'}</div>", unsafe_allow_html=True)
            td_cols[2].markdown(f"<div class='{c_class}' title='{r.get('custody_name', '')}'>{r.get('custody_name', '')}</div>", unsafe_allow_html=True)
            td_cols[3].markdown(f"<div class='{c_class}' title='{r.get('invoice_date', '')}'>{r.get('invoice_date', '')}</div>", unsafe_allow_html=True)
            td_cols[4].markdown(f"<div class='{c_class}' title='{r.get('invoice_no', '')}'>{r.get('invoice_no', '')}</div>", unsafe_allow_html=True)
            td_cols[5].markdown(f"<div class='{c_class}'>{r.get('category', '')}</div>", unsafe_allow_html=True)
            td_cols[6].markdown(f"<div class='{c_class}' title='{r.get('description','')}'>{r.get('description', '')}</div>", unsafe_allow_html=True)
            td_cols[7].markdown(f"<div class='{c_class}' title='{r.get('worker', '')}'>{r.get('worker', '')}</div>", unsafe_allow_html=True)
            td_cols[8].markdown(f"<div class='{c_class}' title='{r.get('supplier_name','')}'>{r.get('supplier_name', '')}</div>", unsafe_allow_html=True)
            td_cols[9].markdown(f"<div class='{c_class}'>{r.get('cr_number', '')}</div>", unsafe_allow_html=True)
            td_cols[10].markdown(f"<div class='{c_class}' title='{r.get('vat_number', '')}'>{r.get('vat_number', '')}</div>", unsafe_allow_html=True)
            td_cols[11].markdown(f"<div class='{c_class}'>{r.get('payment_method', '')}</div>", unsafe_allow_html=True)
            td_cols[12].markdown(f"<div class='{c_class}'>{float(r.get('amount', 0)):,.2f}</div>", unsafe_allow_html=True)
            td_cols[13].markdown(f"<div class='{c_class}'>{float(r.get('vat', 0)):,.2f}</div>", unsafe_allow_html=True)
            val_style = "color:#059669; font-weight:900;" if is_inward else f"color:{header_bg}; font-weight:900;"
            td_cols[14].markdown(f"<div class='{c_class}' style='{val_style}'>{float(r.get('total_invoice', 0)):,.2f}</div>", unsafe_allow_html=True)
            td_cols[15].markdown(f"<div class='{c_class}'>{r.get('cost_center', '')}</div>", unsafe_allow_html=True)
            td_cols[16].markdown(f"<div class='{c_class}'>{r.get('holder_name', '')}</div>", unsafe_allow_html=True)
            with td_cols[17]:
                if st.button("🗑️", key=f"btn_del_live_{r.get('serial_no')}_{idx}", use_container_width=True, help="حذف الفاتورة نهائياً"):
                    st.session_state.confirm_delete_id = r.get("serial_no")
                    st.rerun()
    else:
        empty_msg = "No operations registered for this custody yet." if not is_rtl else "لا توجد عمليات مسجلة لملف هذه العهدة حتى الآن."
        st.markdown(f"<div style='text-align:center; padding:30px; color:#94A3B8; background:{bg_card}; border-radius:8px; border:1px solid {border_color}; margin-top:5px;'>{empty_msg}</div>", unsafe_allow_html=True)

# ─── تبويب التحليلات ───
elif st.session_state.active_tab == "analytics":
    st.markdown(f"<h2 style='color: {header_bg};'>{t['tab_analytics']} (EDM & Executive Insights) - [{selected_custody}]</h2>", unsafe_allow_html=True)
    if custody_records:
        df_inv = pd.DataFrame(custody_records)
        
        st.markdown("""
        <div style='background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); padding: 22px; border-radius: 12px; border-left: 6px solid #38BDF8; margin-bottom: 25px; color: #F8FAFC;'>
            <h3 style='margin: 0 0 10px 0; color: #38BDF8; font-size: 21px;'>🔍 التحليل الاستكشافي للبيانات والتوصيات التنفيذية (EDM & Actionable Insights)</h3>
            <p style='margin: 0 0 15px 0; font-size: 14px; color: #94A3B8; line-height: 1.6;'>يقدم هذا القسم قراءات تحليلية فورية ومستودع توصيات تشغيلية مدعوم بالبيانات الحالية لضبط النثريات وحوكمة المصروفات التشغيلية.</p>
        </div>
        """, unsafe_allow_html=True)
        
        col_edm1, col_edm2, col_edm3 = st.columns(3)
        cat_grouped = df_inv[df_inv['entry_type']=='منصرف'].groupby('category')['total_invoice'].sum() if not df_inv.empty else pd.Series()
        top_cat_name = cat_grouped.idxmax() if not cat_grouped.empty else "لا توجد مصروفات"
        top_cat_val = cat_grouped.max() if not cat_grouped.empty else 0.0

        col_edm1.metric("أعلى بند استهلاكاً (Top Cost Driver)", f"{top_cat_name}", f"{top_cat_val:,.2f} {t['currency']}")
        col_edm2.metric("إجمالي عدد سندات الصرف", f"{len(expense_records)} سند")
        col_edm3.metric("متوسط قيمة السند الواحد", f"{(tot_expense_val / len(expense_records) if len(expense_records)>0 else 0):,.2f} {t['currency']}")

        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown(f"""
        <div style='background: {bg_card}; border: 1px solid {border_color}; border-radius: 10px; padding: 20px; margin-bottom: 25px;'>
            <h4 style='color: {header_bg}; margin-top: 0;'>💡 التوصيات ورؤى العمل التنفيذية (Actionable Recommendations):</h4>
            <ul style='color: {text_color}; font-size: 14px; line-height: 1.7; margin-bottom: 0;'>
                <li><b>تنويع قنوات الدفع النقدية:</b> الاعتماد الحالي على النقد بنسبة عالية يصعب عملية الرقابة اللحظية. يُنصح بتفعيل التحويلات البنكية أو بطاقات الشركات للبنود ذات القيمة العالية.</li>
                <li><b>مراقبة بنود العمالة والمواد:</b> تمثل هذه البنود النسبة الأكبر من منصرفات العهدة التشغيلية؛ لذا يُنصح بوضع سقف مالي أسبوعي معتمد لكل مرحلة لتجنب أي تضخم غير مبرر في المصروفات.</li>
                <li><b>الضبط الضريبي والمستندي:</b> تأكد دائماً من مطابقة الرقم الضريبي (15 رقم) وفحص تكرار أرقام الفواتير لمنع ازدواجية الصرف.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

        c_ch1, c_ch2 = st.columns(2)
        with c_ch1:
            st.markdown("### Operations by Category" if not is_rtl else "### توزيع العمليات حسب التصنيف")
            fig1 = px.pie(df_inv, values='total_invoice', names='category', hole=0.45, color_discrete_sequence=['#6B1D2F', '#4A7C9D', '#10B981', '#F59E0B', '#3B82F6', '#64748B'])
            st.plotly_chart(fig1, use_container_width=True)
        with c_ch2:
            st.markdown("### Flow by Payment Method" if not is_rtl else "### حركة المبالغ حسب طريقة السداد")
            pay_g = df_inv.groupby("payment_method")["total_invoice"].sum().reset_index()
            fig2 = px.bar(pay_g, x='payment_method', y='total_invoice', color='payment_method', color_discrete_sequence=['#6B1D2F', '#4A7C9D', '#10B981', '#F59E0B'])
            st.plotly_chart(fig2, use_container_width=True)
            
        st.divider()
        st.markdown("### 📋 Audit Trail Log (Deletions Only)" if not is_rtl else "### 📋 سجل الرقابة والعمليات المحذوفة (Audit Trail - Deletions Only)")
        logs = load_audit_log()
        del_logs = [l for l in logs if "DELETE" in str(l.get("action", "")) or "DELETE_USER" in str(l.get("action", ""))]
        if del_logs: 
            st.dataframe(pd.DataFrame(del_logs)[["timestamp", "user_name", "action", "details"]], use_container_width=True)
        else:
            st.info("لا توجد عمليات حذف مسجلة في السجل الرقابي حتى الآن.")
    else:
        st.info("Record operations first to view analytics." if not is_rtl else "سجل عمليات أولاً لتفعيل شاشة التحليلات والرسوم البيانية.")

# ─── تبويب سياسات وضوابط النظام ───
elif st.session_state.active_tab == "memory" and (is_company_admin or is_super_admin):
    if not st.session_state.memory_unlocked:
        st.warning("🔒 Security verification required." if not is_rtl else "🔒 التحقق الإداري مطلوب للدخول إلى مركز التحكم بالسياسات.")
        pass_in = st.text_input("Security Code:" if not is_rtl else "رمز الأمان الإداري:", type="password")
        if st.button("Confirm 🗝️" if not is_rtl else "تأكيد الصلاحية 🗝️", type="primary"):
            if pass_in == "6114":
                st.session_state.memory_unlocked = True; st.rerun()
            else: st.error("Incorrect code" if not is_rtl else "رمز الأمان غير صحيح!")
    else:
        st.markdown(f"<h2 style='color:{header_bg};'>⚙️ مركز ضبط وتحديث قواعد وسياسات النظام</h2>", unsafe_allow_html=True)
        st.caption("إدارة التوجيه المحاسبي، مستودع قواعد الاستخراج، وضوابط الحقول المعتمدة.")
        
        t_tab1, t_tab2, t_tab3 = st.tabs([
            "1️⃣ ضوابط التوجيه والقيد", 
            "2️⃣ مستودع العينات المعتمدة", 
            "3️⃣ ضبط نماذج وقواعد الاستخراج"
        ])
        
        curr_cfg = st.session_state.tuning_config
        
        with t_tab1:
            st.subheader("🏢 تخصيص سياسات التوجيه المحاسبي")
            p_col1, p_col2 = st.columns(2)
            c_name_val = p_col1.text_input("الجهة المعتمدة:", value=curr_cfg.get("company_name", "إدارة العمليات والمشاريع"))
            c_id_val = p_col2.text_input("معرّف الإدارة / الكود:", value=curr_cfg.get("company_id", "operations_management"))
            
            st.markdown("**التوجيه المحاسبي الأساسي (System Instruction):**")
            sys_inst_val = st.text_area(
                "نص القواعد المحاسبية المعتمدة:",
                value=curr_cfg.get("system_instruction", ""),
                height=120
            )
            
            st.markdown("**قالب الاستخراج الإلزامي (Mandatory JSON Schema):**")
            schema_json_str = st.text_area(
                "هيكل مخرجات الحقول:",
                value=json.dumps(curr_cfg.get("json_schema", {}), ensure_ascii=False, indent=2),
                height=180
            )
            
            if st.button("💾 حفظ السياسات وضوابط الاستخراج", type="primary"):
                try:
                    parsed_sch = json.loads(schema_json_str)
                    curr_cfg["company_name"] = c_name_val
                    curr_cfg["company_id"] = c_id_val
                    curr_cfg["system_instruction"] = sys_inst_val
                    curr_cfg["json_schema"] = parsed_sch
                    save_tuning_config(curr_cfg)
                    st.session_state.tuning_config = curr_cfg
                    st.success("✅ تم حفظ السياسات بنجاح!")
                except Exception as ex:
                    st.error(f"خطأ في صيغة الـ JSON: {ex}")

        with t_tab2:
            st.subheader("🔄 مستودع العينات والقيود المعتمدة")
            approved_ds = load_approved_dataset()
            ds_count = len(approved_ds)
            
            ready_color = "#10B981" if ds_count >= 10 else "#F59E0B"
            st.markdown(f"""
            <div style='display:flex; justify-content:space-between; align-items:center; background:{bg_card}; padding:15px; border-radius:8px; border:1px solid {border_color}; margin-bottom:15px;'>
                <div>
                    <h4 style='margin:0; color:{header_bg};'>إجمالي السجلات والقيود المعتمدة: <b>{ds_count}</b></h4>
                    <small style='color:#64748B;'>الحد الموصى به لضبط دقة التعرف هو 10 عينات موثقة بكامل الحقول الـ 12.</small>
                </div>
                <div style='font-size:24px; font-weight:900; color:{ready_color};'>
                    {'اكتمال النصاب ✅' if ds_count >= 10 else f'متبقي {10 - ds_count} عينات ⏳'}
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            if approved_ds:
                preview_list = []
                for idx, sample in enumerate(approved_ds):
                    gt = sample.get("ground_truth", {})
                    preview_list.append({
                        "م": idx + 1,
                        "المستند": sample.get("input_doc", ""),
                        "تاريخ الفاتورة": gt.get("invoice_date", ""),
                        "رقم الفاتورة": gt.get("invoice_no", ""),
                        "التصنيف": gt.get("category", ""),
                        "البيان": gt.get("description", ""),
                        "المستلم": gt.get("worker", ""),
                        "المورد": gt.get("supplier_name", ""),
                        "السجل التجاري": gt.get("cr_number", ""),
                        "الرقم الضريبي": gt.get("vat_number", ""),
                        "طريقة الدفع": gt.get("payment_method", ""),
                        "قبل الضريبة": f"{float(gt.get('amount_before_vat', 0.0)):,.2f}",
                        "الضريبة": f"{float(gt.get('vat_amount', 0.0)):,.2f}",
                        "الإجمالي": f"{float(gt.get('total_invoice', 0.0)):,.2f}"
                    })
                st.dataframe(pd.DataFrame(preview_list), use_container_width=True)
                
                col_c1, col_c2 = st.columns([2, 1])
                if col_c2.button("🗑️ مسح مجموعة القيود المعتمدة", use_container_width=True):
                    save_approved_dataset([])
                    st.session_state.approved_dataset = []
                    st.rerun()
            else:
                st.info("لم يتم تسجيل قيود معتمدة بعد. عند اعتماد أي فاتورة في الشيت الرئيسي سيتم توثيقها هنا تلقائياً.")

        with t_tab3:
            st.subheader("⚡ محرك ضبط نماذج الاستخراج")
            
            col_t1, col_t2 = st.columns(2)
            active_model_in = col_t1.text_input("معرّف النموذج المخصص:", value=curr_cfg.get("active_tuned_model", ""), placeholder="tunedModels/custom-invoice-v1")
            use_tuned_chk = col_t2.checkbox("تفعيل توجيه القراءات تلقائياً للنموذج المخصص", value=curr_cfg.get("use_tuned_model", False))
            
            if st.button("حفظ إعدادات توجيه النموذج المخصص 💾"):
                curr_cfg["active_tuned_model"] = active_model_in.strip()
                curr_cfg["use_tuned_model"] = use_tuned_chk
                save_tuning_config(curr_cfg)
                st.session_state.tuning_config = curr_cfg
                st.success("تم تحديث إعدادات التوجيه بنجاح!")
            
            st.divider()
            st.markdown("#### 🚀 إنشاء وحفظ حزمة ضبط جديدة:")
            t_model_name = st.text_input("اسم إصدار الضبط الجديد:", value=f"custom-invoice-{datetime.date.today().strftime('%Y%m%d')}")
            
            if st.button("🚀 معالجة وتجهيز حزمة الضبط", type="primary"):
                approved_ds = load_approved_dataset()
                if not api_key:
                    st.error("⚠️ مفتاح الربط غير متصل.")
                elif len(approved_ds) == 0:
                    st.warning("⚠️ لا توجد عينات معتمدة في المستودع حتى الآن.")
                else:
                    with st.spinner("جاري تحويل البيانات وتجهيز الحزمة..."):
                        training_dataset = []
                        for sample in approved_ds:
                            training_dataset.append({
                                "text_input": f"وثيقة فاتورة: {sample.get('input_doc')}",
                                "output": json.dumps(sample.get("ground_truth"), ensure_ascii=False)
                            })
                        
                        jsonl_path = os.path.join(BASE_DIR, "training_data.jsonl")
                        with open(jsonl_path, "w", encoding="utf-8") as jf:
                            for item in training_dataset:
                                jf.write(json.dumps(item, ensure_ascii=False) + "\n")
                        
                        st.info(f"📁 تم تجهيز ملف البيانات ({len(training_dataset)} عينة) بنجاح.")
                        
                        new_tuned_id = f"tunedModels/{t_model_name.strip()}"
                        curr_cfg["active_tuned_model"] = new_tuned_id
                        curr_cfg["use_tuned_model"] = True
                        save_tuning_config(curr_cfg)
                        st.session_state.tuning_config = curr_cfg
                        
                        log_audit_event("TUNING_CONFIG", f"Configured model {new_tuned_id} with {len(training_dataset)} samples", current_user_data.get('name', 'Admin'), current_user_data.get('name', 'Admin'))
                        st.success(f"🎉 تم تجهيز واعتماد النموذج: `{new_tuned_id}` بنجاح!")
                        time.sleep(1)
                        st.rerun()

# ─── الفوتر ───
st.markdown("""
<div style="text-align:center; padding:20px; color:#64748B; font-weight:bold; font-size:13px; border-top:1px solid #CBD5E1; margin-top:40px;">
    Un-matt ConTech 2026 &bull; Enterprise Systems
</div>
""", unsafe_allow_html=True)