# ============================================================
# app.py — ION Generator Flask Application (Upgraded)
# ============================================================

from flask import Flask, render_template, request, send_file, redirect, url_for, session, jsonify, flash
from werkzeug.utils import secure_filename
from logic.ion_generator import generate_ion
from logic.create_master import create_default_master
from logic.noting_generator import generate_tbrl_noting, generate_lecture_noting, generate_dgmss_noting, generate_fee_noting, generate_cancellation_noting, generate_date_amendment_fax, generate_mayurpankh_erp_fax, generate_appraisal_proforma, generate_coordinator_nomination  # NEW: TBRL Noting/Fax/ERP Generator Import
import os
import json
import shutil
import re
import docx
from datetime import datetime

app = Flask(__name__)
app.secret_key = "ion_generator_secret_2024"

# ── File Paths ──
DATA_FILE     = "data.json"
HISTORY_FILE  = "history.json"
DEFAULTS_FILE = "defaults.json"
MASTERS_FILE  = "masters.json"
CONFIG_FILE   = "config.json"

DEGREE_OPTIONS = ["B.Tech/B.E.", "M.Tech", "B.Sc", "M.Sc", "PhD"]

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# ══════════════════════════════════════
# HELPERS
# ══════════════════════════════════════
def load_json(filepath, default):
    if os.path.exists(filepath):
        with open(filepath, "r") as f:
            return json.load(f)
    return default

def save_json(filepath, data):
    with open(filepath, "w") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def load_departments():
    return load_json(DATA_FILE, {"departments": []})["departments"]

def save_departments(departments):
    save_json(DATA_FILE, {"departments": departments})

def load_history():
    return load_json(HISTORY_FILE, {"history": []})["history"]

def save_history(entry):
    history = load_history()
    history.insert(0, entry)
    history = history[:20]
    save_json(HISTORY_FILE, {"history": history})

def load_defaults():
    return load_json(DEFAULTS_FILE, {})

def load_courses():
    try:
        with open("courses.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def load_columns_map():
    try:
        with open("columns.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def load_masters():
    return load_json(MASTERS_FILE, {"masters": []})["masters"]

def save_masters(masters):
    save_json(MASTERS_FILE, {"masters": masters})

def load_config():
    return load_json(CONFIG_FILE, {"password": "Tbrl0000"})

def check_password(password):
    config = load_config()
    return password == config.get("password", "Tbrl0000")

# ── NEW HELPER: Auto-Extract Variables ──
def extract_variables_from_docx(filepath):
    """Reads a .docx file and extracts all variables inside {{ }}"""
    doc = docx.Document(filepath)
    variables = set()
    pattern = r'\{\{(.+?)\}\}'
    
    for para in doc.paragraphs:
        matches = re.findall(pattern, para.text)
        for match in matches:
            variables.add(match.strip())
            
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                matches = re.findall(pattern, cell.text)
                for match in matches:
                    variables.add(match.strip())
                    
    return list(variables)

# ══════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════
@app.route("/")
def index():
    return render_template("dashboard.html",
        history=load_history()[:5],
        masters=load_masters()
    )

# ══════════════════════════════════════
# ION NOTICE FORM (Legacy / Original Form)
# ══════════════════════════════════════
@app.route("/ion-notice", methods=["GET"])
def ion_notice():
    saved    = session.pop("form_data", {})
    defaults = load_defaults()
    masters  = load_masters()
    
    # Determine which master is selected (first default fillable master, or first fillable master)
    fillable_masters = [m for m in masters if m.get("form_type") == "fillable"]
    selected_master_id = None
    if saved and saved.get("master_id"):
        selected_master_id = saved["master_id"]
    else:
        # Find first default fillable master
        default_master = next((m for m in fillable_masters if m.get("is_default")), None)
        if default_master:
            selected_master_id = default_master["id"]
        elif fillable_masters:
            selected_master_id = fillable_masters[0]["id"]
            
    return render_template(
        "ion_form.html",
        departments=load_departments(),
        degree_options=DEGREE_OPTIONS,
        months=MONTHS,
        saved=saved,
        defaults=defaults,
        masters=masters,
        selected_master_id=selected_master_id
    )

@app.route("/save-defaults", methods=["POST"])
def save_defaults_route():
    defaults = {
        "signatory_name":        request.form.get("signatory_name", ""),
        "signatory_designation": request.form.get("signatory_designation", ""),
    }
    save_json(DEFAULTS_FILE, defaults)
    session["form_data"] = {
        "master_id":             request.form.get("master_id"),
        "degree":                request.form.get("degree"),
        "start_month":           request.form.get("start_month"),
        "start_year":            request.form.get("start_year"),
        "end_month":             request.form.get("end_month"),
        "end_year":              request.form.get("end_year"),
        "last_date":             request.form.get("last_date"),
        "ion_number":            request.form.get("ion_number"),
        "notice_date":           request.form.get("notice_date"),
        "signatory_name":        request.form.get("signatory_name"),
        "signatory_designation": request.form.get("signatory_designation"),
        "departments":           request.form.getlist("departments"),
    }
    return redirect(url_for("ion_notice"))

@app.route("/preview", methods=["POST"])
def preview():
    departments = request.form.getlist("departments")
    num_cols    = 5
    num_rows    = -(-len(departments) // num_cols)
    padded      = departments + [""] * (num_rows * num_cols - len(departments))
    dept_rows   = [padded[r * num_cols:(r + 1) * num_cols] for r in range(num_rows)]

    data = {
        "master_id":             request.form.get("master_id"),
        "degree":                request.form.get("degree"),
        "start_month":           request.form.get("start_month"),
        "start_year":            request.form.get("start_year"),
        "end_month":             request.form.get("end_month"),
        "end_year":              request.form.get("end_year"),
        "last_date":             request.form.get("last_date"),
        "ion_number":            request.form.get("ion_number"),
        "notice_date":           request.form.get("notice_date"),
        "signatory_name":        request.form.get("signatory_name"),
        "signatory_designation": request.form.get("signatory_designation"),
        "departments":           departments,
        "dept_rows":             dept_rows,
    }
    session["form_data"] = data
    return render_template("ion_preview.html", data=data)

@app.route("/edit", methods=["POST"])
def edit():
    session["form_data"] = {
        "master_id":             request.form.get("master_id"),
        "degree":                request.form.get("degree"),
        "start_month":           request.form.get("start_month"),
        "start_year":            request.form.get("start_year"),
        "end_month":             request.form.get("end_month"),
        "end_year":              request.form.get("end_year"),
        "last_date":             request.form.get("last_date"),
        "ion_number":            request.form.get("ion_number"),
        "notice_date":           request.form.get("notice_date"),
        "signatory_name":        request.form.get("signatory_name"),
        "signatory_designation": request.form.get("signatory_designation"),
        "departments":           request.form.getlist("departments"),
    }
    return redirect(url_for("ion_notice"))

@app.route("/generate", methods=["POST"])
def generate():
    masters   = load_masters()
    master_id = request.form.get("master_id", "master_001")
    master    = next((m for m in masters if m["id"] == master_id), masters[0])

    data = {
        "degree":                request.form.get("degree"),
        "start_month":           request.form.get("start_month"),
        "start_year":            request.form.get("start_year"),
        "end_month":             request.form.get("end_month"),
        "end_year":              request.form.get("end_year"),
        "last_date":             request.form.get("last_date"),
        "ion_number":            request.form.get("ion_number"),
        "notice_date":           request.form.get("notice_date"),
        "signatory_name":        request.form.get("signatory_name"),
        "signatory_designation": request.form.get("signatory_designation"),
        "departments":           request.form.getlist("departments"),
    }

    filepath, filename = generate_ion(master["filename"], data)
    save_history({
        "filename":          filename,
        "degree":            data["degree"],
        "period":            f"{data['start_month']} {data['start_year']} - {data['end_month']} {data['end_year']}",
        "generated_at":      datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": len(data["departments"]),
        "master_used":       master["name"],
    })
    return send_file(filepath, as_attachment=True)

@app.route("/download", methods=["POST"])
def download():
    masters   = load_masters()
    master_id = request.form.get("master_id", "master_001")
    master    = next((m for m in masters if m["id"] == master_id), masters[0])

    data = {
        "degree":                request.form.get("degree"),
        "start_month":           request.form.get("start_month"),
        "start_year":            request.form.get("start_year"),
        "end_month":             request.form.get("end_month"),
        "end_year":              request.form.get("end_year"),
        "last_date":             request.form.get("last_date"),
        "ion_number":            request.form.get("ion_number"),
        "notice_date":           request.form.get("notice_date"),
        "signatory_name":        request.form.get("signatory_name"),
        "signatory_designation": request.form.get("signatory_designation"),
        "departments":           request.form.getlist("departments"),
    }

    filepath, filename = generate_ion(master["filename"], data)
    save_history({
        "filename":          filename,
        "degree":            data["degree"],
        "period":            f"{data['start_month']} {data['start_year']} - {data['end_month']} {data['end_year']}",
        "generated_at":      datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": len(data["departments"]),
        "master_used":       master["name"],
    })
    return send_file(filepath, as_attachment=True)

@app.route("/history/download/<filename>")
def download_history(filename):
    filepath = os.path.join("generated_notices", filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return "File not found", 404

@app.route("/history/clear", methods=["POST"])
def clear_history():
    save_json(HISTORY_FILE, {"history": []})
    return redirect(url_for("index"))

@app.route("/download-static/<master_id>", methods=["GET"])
def download_static(master_id):
    masters = load_masters()
    master = next((m for m in masters if m["id"] == master_id), None)
    
    if master and master.get("form_type") == "static":
        filepath = os.path.join("masters", master["filename"])
        if os.path.exists(filepath):
            save_history({
                "filename": master["filename"],
                "degree": "Blank Form",
                "period": "N/A",
                "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
                "departments_count": 0,
                "master_used": master["name"],
            })
            return send_file(filepath, as_attachment=True, download_name=master["filename"])
            
    return "Form not found or is not a static document.", 404

# ══════════════════════════════════════
# DEPARTMENT MANAGER
# ══════════════════════════════════════
@app.route("/departments/add", methods=["POST"])
def add_department():
    name = request.form.get("new_dept", "").strip().upper()
    if name:
        departments = load_departments()
        if name not in departments:
            departments.append(name)
            save_departments(departments)
    return redirect(url_for("ion_notice"))

@app.route("/departments/delete/<name>", methods=["POST"])
def delete_department(name):
    departments = load_departments()
    departments = [d for d in departments if d != name]
    save_departments(departments)
    return redirect(url_for("ion_notice"))

@app.route("/departments/edit", methods=["POST"])
def edit_department():
    old_name = request.form.get("old_name", "").strip()
    new_name = request.form.get("new_name", "").strip().upper()
    if old_name and new_name:
        departments = load_departments()
        departments = [new_name if d == old_name else d for d in departments]
        save_departments(departments)
    return redirect(url_for("ion_notice"))

# ══════════════════════════════════════
# MASTER DOCUMENT MANAGER (Password Protected)
# ══════════════════════════════════════
@app.route("/masters", methods=["GET", "POST"])
def masters_page():
    error = None
    if request.method == "POST":
        password = request.form.get("password", "")
        if check_password(password):
            session["master_unlocked"] = True
            return redirect(url_for("masters_manager"))
        else:
            error = "❌ Incorrect password. Please try again."
    return render_template("master_password.html", error=error)

@app.route("/masters/manager")
def masters_manager():
    if not session.get("master_unlocked"):
        return redirect(url_for("masters_page"))
    return render_template("master_manager.html", masters=load_masters())

@app.route("/masters/lock")
def masters_lock():
    session.pop("master_unlocked", None)
    return redirect(url_for("index"))

@app.route("/masters/create", methods=["POST"])
def create_master():
    if not session.get("master_unlocked"):
        return redirect(url_for("masters_page"))

    source_id   = request.form.get("source_id")
    name        = request.form.get("name", "").strip()
    description = request.form.get("description", "").strip()

    if not name:
        return redirect(url_for("masters_manager"))

    masters     = load_masters()
    source      = next((m for m in masters if m["id"] == source_id), masters[0])

    new_id       = f"master_{str(len(masters)+1).zfill(3)}"
    new_filename = f"{new_id}.docx"

    src_path = os.path.join("masters", source["filename"])
    dst_path = os.path.join("masters", new_filename)
    shutil.copy2(src_path, dst_path)

    masters.append({
        "id":          new_id,
        "name":        name,
        "description": description,
        "filename":    new_filename,
        "created_at":  datetime.now().strftime("%Y-%m-%d"),
        "is_default":  False,
        "form_type":   "static" # Ensures backward compatibility 
    })
    save_masters(masters)
    return redirect(url_for("masters_manager"))

# ── NEW ROUTE: Upload Brand New Document & Extract Variables ──
@app.route("/masters/upload-new", methods=["POST"])
def upload_new_master():
    if not session.get("master_unlocked"):
        return redirect(url_for("masters_page"))
    
    if 'file' not in request.files:
        flash("No file provided.")
        return redirect(url_for("masters_manager"))
        
    file = request.files['file']
    name = request.form.get("name", "Unnamed Master").strip()
    description = request.form.get("description", "Uploaded via Document Manager").strip()
    
    if file and file.filename.endswith('.docx'):
        masters = load_masters()
        new_id = f"master_{str(len(masters)+1).zfill(3)}"
        safe_filename = secure_filename(f"{new_id}_{file.filename}")
        
        # Save file to masters folder
        os.makedirs("masters", exist_ok=True)
        filepath = os.path.join("masters", safe_filename)
        file.save(filepath)
        
        # Trigger Variable Extraction
        variables = extract_variables_from_docx(filepath)
        
        # 🧠 SMART CATEGORIZATION LOGIC
        if len(variables) > 0:
            form_type = "dynamic"
            flash_msg = f"✨ Dynamic Master added! Extracted {len(variables)} form fields."
        else:
            form_type = "static"
            flash_msg = "🖨️ Static Template added! No variables found, saving as a printable blank document."
        
        # Append to masters.json
        masters.append({
            "id": new_id,
            "name": name,
            "description": description,
            "filename": safe_filename,
            "created_at": datetime.now().strftime("%Y-%m-%d"),
            "is_default": False,
            "form_type": form_type,  # Now correctly assigns dynamic vs static
            "variables": variables
        })
        save_masters(masters)
        flash(flash_msg)
        
    return redirect(url_for("masters_manager"))

@app.route("/masters/toggle-default/<master_id>", methods=["POST"])
def toggle_default_master(master_id):
    if not session.get("master_unlocked"):
        return redirect(url_for("masters_page"))
    masters = load_masters()
    for m in masters:
        if m["id"] == master_id:
            m["is_default"] = not m.get("is_default", False)
            break
    save_masters(masters)
    return redirect(url_for("masters_manager"))

@app.route("/masters/delete/<master_id>", methods=["POST"])
def delete_master(master_id):
    if not session.get("master_unlocked"):
        return redirect(url_for("masters_page"))
        
    # 🔒 NEW: Secondary password check for deletion
    confirm_password = request.form.get("confirm_password", "")
    if not check_password(confirm_password):
        flash("❌ Incorrect password! Deletion cancelled.")
        return redirect(url_for("masters_manager"))

    masters = load_masters()
    master  = next((m for m in masters if m["id"] == master_id), None)
    
    if master and not master.get("is_default"):
        filepath = os.path.join("masters", master["filename"])
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except OSError:
                pass # Ignore if file was already removed manually
                
        masters = [m for m in masters if m["id"] != master_id]
        save_masters(masters)
        flash(f"🗑️ Master template '{master['name']}' was permanently deleted.")
        
    return redirect(url_for("masters_manager"))

@app.route("/masters/download/<master_id>")
def download_master(master_id):
    if not session.get("master_unlocked"):
        return redirect(url_for("masters_page"))
    masters = load_masters()
    master  = next((m for m in masters if m["id"] == master_id), None)
    if master:
        filepath = os.path.join("masters", master["filename"])
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True,
                           download_name=f"{master['name']}.docx")
    return "Master not found", 404

@app.route("/masters/upload/<master_id>", methods=["POST"])
def upload_master(master_id):
    if not session.get("master_unlocked"):
        return redirect(url_for("masters_page"))
    masters = load_masters()
    master  = next((m for m in masters if m["id"] == master_id), None)
    if master and "file" in request.files:
        file = request.files["file"]
        if file.filename.endswith(".docx"):
            filepath = os.path.join("masters", master["filename"])
            file.save(filepath)
    return redirect(url_for("masters_manager"))

# ══════════════════════════════════════
# DYNAMIC FORM RENDERING & SUBMISSION
# ══════════════════════════════════════
@app.route('/dynamic-form/<master_id>')
def dynamic_form(master_id):
    masters = load_masters()
    master = next((m for m in masters if m["id"] == master_id), None)
    
    if not master or master.get("form_type") != "dynamic":
        return "Dynamic form not found or invalid type.", 404
        
    return render_template('dynamic_form.html', template=master)

import docx # Make sure this is imported at the top of app.py

@app.route('/submit-dynamic', methods=['POST'])
def submit_dynamic():
    master_id = request.form.get("template_id")
    masters = load_masters()
    master = next((m for m in masters if m["id"] == master_id), None)
    
    if not master:
        return "Master not found", 404
        
    # 1. Grab all the user's answers from the form
    data = {}
    for var in master.get("variables", []):
        data[var] = request.form.get(var, "")
        
    # 2. Open the Master Document
    master_path = os.path.join("masters", master["filename"])
    doc = docx.Document(master_path)
    
    def replace_text(paragraph, key, value):
        placeholder = f"{{{{{key}}}}}" # Looks for {{variable_name}}
        if placeholder in paragraph.text:
            replaced = False
            for run in paragraph.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, str(value))
                    replaced = True
            if not replaced:
                paragraph.text = paragraph.text.replace(placeholder, str(value))

    # 4. Search and Replace in standard text
    for para in doc.paragraphs:
        for k, v in data.items():
            replace_text(para, k, v)
            
    # 5. Search and Replace inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for k, v in data.items():
                        replace_text(para, k, v)
                        
    # 6. Save the new generated file
    os.makedirs("generated_notices", exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    clean_name = master['name'].replace(' ', '_').replace('/', '_').replace('\\', '_')
    filename = f"Dynamic_{clean_name}_{timestamp}.docx"
    output_path = os.path.join("generated_notices", filename)
    doc.save(output_path)
    
    # 7. Add to history
    save_history({
        "filename": filename,
        "degree": "Dynamic Form",
        "period": "N/A",
        "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": len(data),
        "master_used": master["name"],
    })
    
    # 8. Send to user
    return send_file(output_path, as_attachment=True)

# ══════════════════════════════════════
# TBRL NOTING GENERATOR (NEW)
# ══════════════════════════════════════
@app.route("/tbrl-noting", methods=["GET"])
def tbrl_noting():
    defaults = load_defaults()
    groups = ["AFTD", "ADS", "BIDS", "BEHI", "SS", "TELIC", "PPG", "QMG", "ETF", "PC", "WHD", "EXPD", "WHT&E", "ARISE", "PCD", "AIG", "RTRS", "S&D", "R&QA", "WKS", "SEED", "CERBERUS", "DPB", "HSP", "I2G", "HRDD"]
    courses = load_courses()
    default_columns = load_columns_map().get("tbrl_noting", [])
    return render_template("tbrl_noting_form.html", defaults=defaults, groups=groups, courses=courses, default_columns=default_columns)

@app.route("/generate-noting", methods=["POST"])
def generate_noting():
    # 1. Update Signatory Defaults in Memory
    defaults = load_defaults()
    defaults["sig1_name"] = request.form.get("sig1_name", "")
    defaults["sig1_desig"] = request.form.get("sig1_desig", "")
    defaults["sig2_name"] = request.form.get("sig2_name", "")
    defaults["sig2_desig"] = request.form.get("sig2_desig", "")
    save_json(DEFAULTS_FILE, defaults)

    # 2. Extract Table Configuration & Data
    columns = request.form.getlist("table_columns")
    num_rows = int(request.form.get("num_rows", 1))
    
    nominees = []
    for i in range(num_rows):
        row_data = {}
        for col in columns:
            row_data[col] = request.form.get(f"{col}_{i}", "")
        nominees.append(row_data)

    # 3. Handle 'Other' Course Type
    course_type = request.form.get("course_type")
    if course_type == "Others":
        course_type = request.form.get("custom_course_type", "Course")

    ref_no = request.form.get("form_no") or request.form.get("ref_no")
    lab_name = request.form.get("lab_name", "टीबीआरएल")
    
    references_raw = request.form.get("references_json")
    references = []
    if references_raw:
        try:
            references = json.loads(references_raw)
        except Exception:
            pass
    if not references:
        references = [{
            "source": ref_no or "",
            "date": request.form.get("ref_date", "")
        }]

    # 4. Compile all data to send to noting_generator.py
    data = {
        "ref_no": ref_no,
        "subject_hindi": request.form.get("subject_hindi"),
        "subject_english": request.form.get("subject_english"),
        "reference_text": request.form.get("reference_text"),
        "ref_date": request.form.get("ref_date"),
        "start_date": request.form.get("start_date"),
        "end_date": request.form.get("end_date"),
        "org_institute": request.form.get("org_institute"),
        "course_title": request.form.get("course_title"),
        "course_type": course_type,
        "group_name": request.form.get("group_name"),
        "sig1_name": defaults["sig1_name"],
        "sig1_desig": defaults["sig1_desig"],
        "sig2_name": defaults["sig2_name"],
        "sig2_desig": defaults["sig2_desig"],
        "columns": columns,
        "nominees": nominees,
        "lab_name": lab_name,
        "references": references
    }

    # 5. Generate Document
    filepath, filename = generate_tbrl_noting(data)
    
    # 6. Save to Dashboard History
    save_history({
        "filename": filename,
        "degree": "TBRL Noting",
        "period": f"{data.get('start_date', '')} to {data.get('end_date', '')}",
        "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": len(nominees),
        "master_used": "TBRL Noting Master"
    })
    
    return send_file(filepath, as_attachment=True)


# ══════════════════════════════════════
# LECTURE NOTING GENERATOR
# ══════════════════════════════════════
@app.route("/lecture-noting", methods=["GET"])
def lecture_noting():
    defaults = load_defaults()
    groups = ["AFTD", "ADS", "BIDS", "BEHI", "SS", "TELIC", "PPG", "QMG", "ETF", "PC", "WHD", "EXPD", "WHT&E", "ARISE", "PCD", "AIG", "RTRS", "S&D", "R&QA", "WKS", "SEED", "CERBERUS", "DPB", "HSP", "I2G", "HRDD"]
    courses = load_courses()
    default_columns = load_columns_map().get("lecture_noting", [])
    return render_template("lecture_noting_form.html", defaults=defaults, groups=groups, courses=courses, default_columns=default_columns)


@app.route("/generate-lecture-noting", methods=["POST"])
def generate_lecture_noting_route():
    # 1. Update Signatory Defaults in Memory
    defaults = load_defaults()
    defaults["sig1_name"] = request.form.get("sig1_name", "")
    defaults["sig1_desig"] = request.form.get("sig1_desig", "")
    defaults["sig2_name"] = request.form.get("sig2_name", "")
    defaults["sig2_desig"] = request.form.get("sig2_desig", "")
    save_json(DEFAULTS_FILE, defaults)

    # 2. Extract Table Configuration & Data
    columns = request.form.getlist("table_columns")
    num_rows = int(request.form.get("num_rows", 1))
    
    nominees = []
    for i in range(num_rows):
        row_data = {}
        for col in columns:
            row_data[col] = request.form.get(f"{col}_{i}", "")
        nominees.append(row_data)

    course_type = request.form.get("course_type")
    
    ref_no = request.form.get("form_no") or request.form.get("ref_no")
    lab_name = request.form.get("lab_name", "टीबीआरएल")
    
    references_raw = request.form.get("references_json")
    references = []
    if references_raw:
        try:
            references = json.loads(references_raw)
        except Exception:
            pass
    if not references:
        references = [{
            "source": ref_no or "",
            "date": request.form.get("ref_date", "")
        }]

    # Compile all data
    data = {
        "ref_no": ref_no,
        "subject_hindi": request.form.get("subject_hindi"),
        "subject_english": request.form.get("subject_english"),
        "reference_text": request.form.get("reference_text"),
        "ref_date": request.form.get("ref_date"),
        "start_date": request.form.get("start_date"),
        "end_date": request.form.get("end_date"),
        "org_institute": request.form.get("org_institute"),
        "course_title": request.form.get("course_title"),
        "lecture_title": request.form.get("lecture_title"),
        "course_type": course_type,
        "group_name": request.form.get("group_name"),
        "sig1_name": defaults["sig1_name"],
        "sig1_desig": defaults["sig1_desig"],
        "sig2_name": defaults["sig2_name"],
        "sig2_desig": defaults["sig2_desig"],
        "columns": columns,
        "nominees": nominees,
        "lab_name": lab_name,
        "references": references
    }

    # Generate Document
    filepath, filename = generate_lecture_noting(data)
    
    # Save to Dashboard History
    save_history({
        "filename": filename,
        "degree": "Lecture Noting",
        "period": f"{data.get('start_date', '')} to {data.get('end_date', '')}",
        "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": len(nominees),
        "master_used": "Lecture Noting Master"
    })
    
    return send_file(filepath, as_attachment=True)


# ══════════════════════════════════════
# DGMSS NOTING GENERATOR
# ══════════════════════════════════════
@app.route("/dgmss-noting", methods=["GET"])
def dgmss_noting():
    defaults = load_defaults()
    groups = ["AFTD", "ADS", "BIDS", "BEHI", "SS", "TELIC", "PPG", "QMG", "ETF", "PC", "WHD", "EXPD", "WHT&E", "ARISE", "PCD", "AIG", "RTRS", "S&D", "R&QA", "WKS", "SEED", "CERBERUS", "DPB", "HSP", "I2G", "HRDD", "BTS"]
    courses = load_courses()
    default_columns = load_columns_map().get("dgmss_noting", [])
    return render_template("dgmss_noting_form.html", defaults=defaults, groups=groups, courses=courses, default_columns=default_columns)


@app.route("/generate-dgmss-noting", methods=["POST"])
def generate_dgmss_noting_route():
    # 1. Update Signatory Defaults in Memory
    defaults = load_defaults()
    defaults["sig1_name"] = request.form.get("sig1_name", "")
    defaults["sig1_desig"] = request.form.get("sig1_desig", "")
    defaults["sig2_name"] = request.form.get("sig2_name", "")
    defaults["sig2_desig"] = request.form.get("sig2_desig", "")
    save_json(DEFAULTS_FILE, defaults)

    # 2. Extract Table Configuration & Data
    columns = request.form.getlist("table_columns")
    num_rows = int(request.form.get("num_rows", 1))
    
    nominees = []
    for i in range(num_rows):
        row_data = {}
        for col in columns:
            row_data[col] = request.form.get(f"{col}_{i}", "")
        nominees.append(row_data)

    course_type = request.form.get("course_type")
    
    ref_no = request.form.get("form_no") or request.form.get("ref_no")
    lab_name = request.form.get("lab_name", "टीबीआरएल")
    
    references_raw = request.form.get("references_json")
    references = []
    if references_raw:
        try:
            references = json.loads(references_raw)
        except Exception:
            pass
    if not references:
        references = [{
            "source": ref_no or "",
            "date": request.form.get("ref_date", "")
        }]

    # Compile all data
    data = {
        "ref_no": ref_no,
        "subject_hindi": request.form.get("subject_hindi"),
        "subject_english": request.form.get("subject_english"),
        "reference_text": request.form.get("reference_text"),
        "ref_date": request.form.get("ref_date"),
        "start_date": request.form.get("start_date"),
        "end_date": request.form.get("end_date"),
        "org_institute": request.form.get("org_institute"),
        "course_title": request.form.get("course_title"),
        "course_type": course_type,
        "group_name": request.form.get("group_name"),
        "sig1_name": defaults["sig1_name"],
        "sig1_desig": defaults["sig1_desig"],
        "sig2_name": defaults["sig2_name"],
        "sig2_desig": defaults["sig2_desig"],
        "columns": columns,
        "nominees": nominees,
        "lab_name": lab_name,
        "references": references
    }

    # Generate Document
    filepath, filename = generate_dgmss_noting(data)
    
    # Save to Dashboard History
    save_history({
        "filename": filename,
        "degree": "DGMSS Noting",
        "period": f"{data.get('start_date', '')} to {data.get('end_date', '')}",
        "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": len(nominees),
        "master_used": "DGMSS Noting Master"
    })
    
    return send_file(filepath, as_attachment=True)


# FEE RELATED NOTING GENERATOR
# ══════════════════════════════════════
@app.route("/fee-noting", methods=["GET"])
def fee_noting():
    defaults = load_defaults()
    groups = ["AFTD", "ADS", "BIDS", "BEHI", "SS", "TELIC", "PPG", "QMG", "ETF", "PC", "WHD", "EXPD", "WHT&E", "ARISE", "PCD", "AIG", "RTRS", "S&D", "R&QA", "WKS", "SEED", "CERBERUS", "DPB", "HSP", "I2G", "HRDD", "BTS"]
    courses = load_courses()
    default_columns = load_columns_map().get("fee_noting", [])
    return render_template("fee_noting_form.html", defaults=defaults, groups=groups, courses=courses, default_columns=default_columns)


@app.route("/generate-fee-noting", methods=["POST"])
def generate_fee_noting_route():
    # 1. Update Signatory Defaults in Memory
    defaults = load_defaults()
    defaults["sig1_name"] = request.form.get("sig1_name", "")
    defaults["sig1_desig"] = request.form.get("sig1_desig", "")
    defaults["sig2_name"] = request.form.get("sig2_name", "")
    defaults["sig2_desig"] = request.form.get("sig2_desig", "")
    save_json(DEFAULTS_FILE, defaults)

    # 2. Extract Table Configuration & Data
    columns = request.form.getlist("table_columns")
    num_rows = int(request.form.get("num_rows", 1))
    
    nominees = []
    for i in range(num_rows):
        row_data = {}
        for col in columns:
            row_data[col] = request.form.get(f"{col}_{i}", "")
        nominees.append(row_data)

    course_type = request.form.get("course_type")
    
    ref_no = request.form.get("form_no") or request.form.get("ref_no")
    lab_name = request.form.get("lab_name", "टीबीआरएल")
    
    references_raw = request.form.get("references_json")
    references = []
    if references_raw:
        try:
            references = json.loads(references_raw)
        except Exception:
            pass
    if not references:
        references = [{
            "source": ref_no or "",
            "date": request.form.get("ref_date", "")
        }]

    # Compile all data
    data = {
        "ref_no": ref_no,
        "subject_hindi": request.form.get("subject_hindi"),
        "subject_english": request.form.get("subject_english"),
        "reference_text": request.form.get("reference_text"),
        "ref_date": request.form.get("ref_date"),
        "ref_mail_date": request.form.get("ref_mail_date"),
        "start_date": request.form.get("start_date"),
        "end_date": request.form.get("end_date"),
        "org_institute": request.form.get("org_institute"),
        "course_title": request.form.get("course_title"),
        "course_type": course_type,
        "group_name": request.form.get("group_name"),
        "sig1_name": defaults["sig1_name"],
        "sig1_desig": defaults["sig1_desig"],
        "sig2_name": defaults["sig2_name"],
        "sig2_desig": defaults["sig2_desig"],
        "columns": columns,
        "nominees": nominees,
        "lab_name": lab_name,
        "references": references
    }

    # Generate Document
    filepath, filename = generate_fee_noting(data)
    
    # Save to Dashboard History
    save_history({
        "filename": filename,
        "degree": "Fee Related Noting",
        "period": f"{data.get('start_date', '')} to {data.get('end_date', '')}",
        "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": len(nominees),
        "master_used": "Fee Related Noting Master"
    })
    
    return send_file(filepath, as_attachment=True)


# NOMINATION CANCELLATION NOTING GENERATOR
# ══════════════════════════════════════
@app.route("/cancellation-noting", methods=["GET"])
def cancellation_noting():
    defaults = load_defaults()
    groups = ["AFTD", "ADS", "BIDS", "BEHI", "SS", "TELIC", "PPG", "QMG", "ETF", "PC", "WHD", "EXPD", "WHT&E", "ARISE", "PCD", "AIG", "RTRS", "S&D", "R&QA", "WKS", "SEED", "CERBERUS", "DPB", "HSP", "I2G", "HRDD", "BTS"]
    courses = load_courses()
    default_columns = load_columns_map().get("cancellation_noting", [])
    return render_template("cancellation_noting_form.html", defaults=defaults, groups=groups, courses=courses, default_columns=default_columns)


@app.route("/generate-cancellation-noting", methods=["POST"])
def generate_cancellation_noting_route():
    # 1. Update Signatory Defaults in Memory
    defaults = load_defaults()
    defaults["sig1_name"] = request.form.get("sig1_name", "")
    defaults["sig1_desig"] = request.form.get("sig1_desig", "")
    defaults["sig2_name"] = request.form.get("sig2_name", "")
    defaults["sig2_desig"] = request.form.get("sig2_desig", "")
    save_json(DEFAULTS_FILE, defaults)

    # 2. Extract Table Configuration & Data
    columns = request.form.getlist("table_columns")
    num_rows = int(request.form.get("num_rows", 1))
    
    nominees = []
    for i in range(num_rows):
        row_data = {}
        for col in columns:
            row_data[col] = request.form.get(f"{col}_{i}", "")
        nominees.append(row_data)

    course_type = request.form.get("course_type")
    
    ref_no = request.form.get("form_no") or request.form.get("ref_no")
    lab_name = request.form.get("lab_name", "टीबीआरएल")
    
    references_raw = request.form.get("references_json")
    references = []
    if references_raw:
        try:
            references = json.loads(references_raw)
        except Exception:
            pass
    if not references:
        references = [{
            "source": ref_no or "",
            "date": request.form.get("ref_date", "")
        }]

    # Compile all data
    data = {
        "ref_no": ref_no,
        "subject_hindi": request.form.get("subject_hindi"),
        "subject_english": request.form.get("subject_english"),
        "reference_text": request.form.get("reference_text"),
        "ref_date": request.form.get("ref_date"),
        "ref_mail_date": request.form.get("ref_mail_date"),
        "ion_ref_source": request.form.get("ion_ref_source"),
        "ion_ref_date": request.form.get("ion_ref_date"),
        "cancel_nominee_name": request.form.get("cancel_nominee_name"),
        "cancel_group_name": request.form.get("cancel_group_name"),
        "cancel_reason": request.form.get("cancel_reason"),
        "course_type_short": request.form.get("course_type_short"),
        "start_date": request.form.get("start_date"),
        "end_date": request.form.get("end_date"),
        "org_institute": request.form.get("org_institute"),
        "course_title": request.form.get("course_title"),
        "course_type": course_type,
        "group_name": request.form.get("group_name"),
        "sig1_name": defaults["sig1_name"],
        "sig1_desig": defaults["sig1_desig"],
        "sig2_name": defaults["sig2_name"],
        "sig2_desig": defaults["sig2_desig"],
        "columns": columns,
        "nominees": nominees,
        "lab_name": lab_name,
        "references": references
    }

    # Generate Document
    filepath, filename = generate_cancellation_noting(data)
    
    # Save to Dashboard History
    save_history({
        "filename": filename,
        "degree": "Nomination Cancellation",
        "period": f"{data.get('start_date', '')} to {data.get('end_date', '')}",
        "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": len(nominees),
        "master_used": "Nomination Cancellation Master"
    })
    
    return send_file(filepath, as_attachment=True)


# DATE AMENDMENT FAX GENERATOR
# ═════════════════════════════
@app.route("/date-amendment-fax", methods=["GET"])
def date_amendment_fax():
    defaults = load_defaults()
    courses = load_courses()
    default_columns = load_columns_map().get("date_amendment_fax", [])
    return render_template("date_amendment_fax_form.html", defaults=defaults, courses=courses, default_columns=default_columns)


@app.route("/generate-date-amendment-fax", methods=["POST"])
def generate_date_amendment_fax_route():
    # 1. Extract Table Configuration & Data
    for_columns = request.form.getlist("for_columns")
    read_columns = request.form.getlist("read_columns")
    for_rows_count = int(request.form.get("for_rows_count", 1))
    read_rows_count = int(request.form.get("read_rows_count", 1))

    for_rows = []
    for i in range(for_rows_count):
        row_data = {}
        for col in for_columns:
            row_data[col] = request.form.get(f"for_{col}_{i}", "")
        for_rows.append(row_data)

    read_rows = []
    for i in range(read_rows_count):
        row_data = {}
        for col in read_columns:
            row_data[col] = request.form.get(f"read_{col}_{i}", "")
        read_rows.append(row_data)

    # 2. Assemble document payload
    data = {
        "ref_no": request.form.get("ref_no", ""),
        "fax_no": request.form.get("fax_no", ""),
        "ref_date": request.form.get("ref_date", ""),
        "to_text": request.form.get("to_text", ""),
        "subject_hindi": request.form.get("subject_hindi", ""),
        "subject_english": request.form.get("subject_english", ""),
        "ref_text": request.form.get("ref_text", ""),
        "ref_to_hindi": request.form.get("ref_to_hindi", ""),
        "ref_to_eng": request.form.get("ref_to_eng", ""),
        "num_courses_hindi": request.form.get("num_courses_hindi", "तीन"),
        "num_courses_eng": request.form.get("num_courses_eng", "three"),
        "courses_plural_eng": request.form.get("courses_plural_eng", "Courses"),
        "amended_courses_hindi": request.form.get("amended_courses_hindi", "एक"),
        "amended_courses_eng": request.form.get("amended_courses_eng", "one"),
        "for_columns": for_columns,
        "for_rows": for_rows,
        "read_columns": read_columns,
        "read_rows": read_rows,
        "sig_name": request.form.get("sig_name", ""),
        "sig_desig": request.form.get("sig_desig", ""),
        "for_director": request.form.get("for_director", "")
    }

    # Generate Fax document
    filepath, filename = generate_date_amendment_fax(data)

    # Save to dashboard history
    save_history({
        "filename": filename,
        "degree": f"Date Amendment Fax ({request.form.get('course_type_select', 'C.E.P.')})",
        "period": "N/A",
        "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": for_rows_count,
        "master_used": "Date Amendment Fax Master"
    })

    return send_file(filepath, as_attachment=True)


# MAYURPANKH ERP FAX GENERATOR
# ═════════════════════════════
@app.route("/mayurpankh-erp", methods=["GET"])
def mayurpankh_erp():
    defaults = load_defaults()
    courses = load_courses()
    default_columns = load_columns_map().get("mayurpankh_erp", [])
    return render_template("mayurpankh_erp_form.html", defaults=defaults, courses=courses, default_columns=default_columns)


@app.route("/generate-mayurpankh-erp-fax", methods=["POST"])
def generate_mayurpankh_erp_fax_route():
    # 1. Extract Table Configuration & Data
    columns = request.form.getlist("columns")
    rows_count = int(request.form.get("rows_count", 1))

    rows = []
    for i in range(rows_count):
        row_data = {}
        for col in columns:
            row_data[col] = request.form.get(f"{col}_{i}", "")
        rows.append(row_data)

    # 2. Assemble document payload
    data = {
        "ref_no": request.form.get("ref_no", ""),
        "fax_no": request.form.get("fax_no", ""),
        "ref_date": request.form.get("ref_date", ""),
        "to_text": request.form.get("to_text", ""),
        "subject_hindi": request.form.get("subject_hindi", ""),
        "subject_english": request.form.get("subject_english", ""),
        "ref_text": request.form.get("ref_text", ""),
        "attn_text": request.form.get("attn_text", ""),
        "body_hindi": request.form.get("body_hindi", ""),
        "body_english": request.form.get("body_english", ""),
        "confirm_text": request.form.get("confirm_text", ""),
        "columns": columns,
        "rows": rows,
        "sig_name": request.form.get("sig_name", ""),
        "sig_desig": request.form.get("sig_desig", ""),
        "for_director": request.form.get("for_director", "")
    }

    # Generate Fax document
    filepath, filename = generate_mayurpankh_erp_fax(data)

    # Save to dashboard history
    save_history({
        "filename": filename,
        "degree": "Mayurpankh ERP Fax",
        "period": "N/A",
        "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": rows_count,
        "master_used": "Mayurpankh ERP Fax Master"
    })

    return send_file(filepath, as_attachment=True)


# APPRAISAL PROFORMA GENERATOR
# ══════════════════════════════════════
@app.route("/appraisal-proforma", methods=["GET"])
def appraisal_proforma():
    defaults = load_defaults()
    groups = ["AFTD", "ADS", "BIDS", "BEHI", "SS", "TELIC", "PPG", "QMG", "ETF", "PC", "WHD", "EXPD", "WHT&E", "ARISE", "PCD", "AIG", "RTRS", "S&D", "R&QA", "WKS", "SEED", "CERBERUS", "DPB", "HSP", "I2G", "HRDD", "BTS"]
    return render_template("appraisal_form.html", defaults=defaults, groups=groups)


@app.route("/generate-appraisal-proforma", methods=["POST"])
def generate_appraisal_proforma_route():
    data = {
        "student_name_hindi":   request.form.get("student_name_hindi", ""),
        "father_name_hindi":    request.form.get("father_name_hindi", ""),
        "student_name_eng":     request.form.get("student_name_eng", ""),
        "father_name_eng":      request.form.get("father_name_eng", ""),
        "branch_semester":      request.form.get("branch_semester", ""),
        "mobile_no":            request.form.get("mobile_no", ""),
        "institute_name":       request.form.get("institute_name", ""),
        "group_name":           request.form.get("group_name", ""),
        "group_director":       request.form.get("group_director", ""),
        "div_head_coordinator": request.form.get("div_head_coordinator", ""),
        "project_title":        request.form.get("project_title", ""),
        "work_description":     request.form.get("work_description", ""),
        "date_of_joining":      request.form.get("date_of_joining", ""),
        "date_of_completion":   request.form.get("date_of_completion", ""),
        "attendance_grade":     request.form.get("attendance_grade", ""),
        "performance_rating":   request.form.get("performance_rating", ""),
        "report_submitted":     request.form.get("report_submitted", ""),
        "remarks":              request.form.get("remarks", ""),
        "recommendation":       request.form.get("recommendation", "")
    }

    filepath, filename = generate_appraisal_proforma(data)

    save_history({
        "filename":          filename,
        "degree":            data["branch_semester"],
        "period":            f"{data['date_of_joining']} to {data['date_of_completion']}",
        "generated_at":      datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": 1,
        "master_used":       "Appraisal Proforma Master"
    })

    return send_file(filepath, as_attachment=True)


# COORDINATOR NOMINATION GENERATOR
# ══════════════════════════════════════
@app.route("/coordinator-nomination/<master_id>", methods=["GET"])
def coordinator_nomination(master_id):
    defaults = load_defaults()
    masters = load_masters()
    master = next((m for m in masters if m["id"] == master_id), None)
    if not master:
        flash(f"Master template {master_id} not found.")
        return redirect(url_for("internship_menu"))
        
    groups = ["AFTD", "ADS", "BIDS", "BEHI", "SS", "TELIC", "PPG", "QMG", "ETF", "PC", "WHD", "EXPD", "WHT&E", "ARISE", "PCD", "AIG", "RTRS", "S&D", "R&QA", "WKS", "SEED", "CERBERUS", "DPB", "HSP", "I2G", "HRDD", "BTS"]
    return render_template("coordinator_nomination_form.html", defaults=defaults, groups=groups, master=master)


@app.route("/generate-coordinator-nomination/<master_id>", methods=["POST"])
def generate_coordinator_nomination_route(master_id):
    masters = load_masters()
    master = next((m for m in masters if m["id"] == master_id), None)
    if not master:
        return "Master not found", 404
        
    trainees_json = request.form.get("trainees_json", "[]")
    try:
        trainees = json.loads(trainees_json)
    except Exception:
        trainees = []

    data = {
        "training_session": request.form.get("training_session", ""),
        "ref_no":           request.form.get("ref_no", ""),
        "date":             request.form.get("date", ""),
        "signatory_name":   request.form.get("signatory_name", ""),
        "signatory_desig":  request.form.get("signatory_desig", ""),
        "recipient_group":  request.form.get("recipient_group", ""),
        "trainees":         trainees
    }

    filepath, filename = generate_coordinator_nomination(master_id, data)

    save_history({
        "filename":          filename,
        "degree":            f"Selected Trainees ({len(trainees)})",
        "period":            data["training_session"],
        "generated_at":      datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "departments_count": len(trainees),
        "master_used":       master["name"]
    })

    return send_file(filepath, as_attachment=True)


# ── API History ──
@app.route("/api/history")
def api_history():
    from flask import jsonify
    return jsonify({"history": load_history()})

# ── Restructured Module Routes (Training & Internship) ──
@app.route("/training")
def training_menu():
    """Renders the Training Module submenu with pillars."""
    category = request.args.get("sub_module", "noting")
    return render_template("training_menu.html", active_category=category)

@app.route("/internship")
def internship_menu():
    """Renders the Internship Module submenu."""
    return render_template("internship_menu.html")

@app.route("/api/templates")
def api_templates():
    """Returns templates filtered by module and sub_module."""
    module = request.args.get("module")
    sub_module = request.args.get("sub_module")
    masters = load_masters()
    
    filtered = masters
    if module:
        filtered = [m for m in filtered if m.get("module") == module]
    if sub_module:
        filtered = [m for m in filtered if m.get("sub_module") == sub_module]
        
    return jsonify(filtered)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)