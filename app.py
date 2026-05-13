# ============================================================
# app.py — ION Generator Flask Application
# ============================================================

from flask import Flask, render_template, request, send_file, redirect, url_for, session, jsonify
from logic.ion_generator import generate_ion
from logic.create_master import create_default_master
import os
import json
import shutil
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

def load_masters():
    return load_json(MASTERS_FILE, {"masters": []})["masters"]

def save_masters(masters):
    save_json(MASTERS_FILE, {"masters": masters})

def load_config():
    return load_json(CONFIG_FILE, {"password": "Tbrl0000"})

def check_password(password):
    config = load_config()
    return password == config.get("password", "Tbrl0000")

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
# ION NOTICE FORM
# ══════════════════════════════════════
@app.route("/ion-notice", methods=["GET"])
def ion_notice():
    saved    = session.pop("form_data", {})
    defaults = load_defaults()
    masters  = load_masters()
    return render_template(
        "ion_form.html",
        departments=load_departments(),
        degree_options=DEGREE_OPTIONS,
        months=MONTHS,
        saved=saved,
        defaults=defaults,
        masters=masters
    )

# ── Save Defaults ──
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

# ── Preview ──
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

# ── Edit (back from preview) ──
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

# ── Generate (direct download) ──
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

# ── Download after Preview ──
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

# ── Download from History ──
@app.route("/history/download/<filename>")
def download_history(filename):
    filepath = os.path.join("generated_notices", filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return "File not found", 404

# ── Clear History ──
@app.route("/history/clear", methods=["POST"])
def clear_history():
    save_json(HISTORY_FILE, {"history": []})
    return redirect(url_for("index"))

# ── Download Static Form Directly ──
@app.route("/download-static/<master_id>", methods=["GET"])
def download_static(master_id):
    masters = load_masters()
    master = next((m for m in masters if m["id"] == master_id), None)
    
    if master and master.get("form_type") == "static":
        filepath = os.path.join("masters", master["filename"])
        if os.path.exists(filepath):
            # Save history so HR admins can track who downloaded the blank forms
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

# ── Create New Master from Existing ──
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

    # Generate new ID
    new_id       = f"master_{str(len(masters)+1).zfill(3)}"
    new_filename = f"{new_id}.docx"

    # Copy the source master file
    src_path = os.path.join("masters", source["filename"])
    dst_path = os.path.join("masters", new_filename)
    shutil.copy2(src_path, dst_path)

    # Add to masters list
    masters.append({
        "id":          new_id,
        "name":        name,
        "description": description,
        "filename":    new_filename,
        "created_at":  datetime.now().strftime("%Y-%m-%d"),
        "is_default":  False
    })
    save_masters(masters)
    return redirect(url_for("masters_manager"))

# ── Set Default Master ──
@app.route("/masters/set-default/<master_id>", methods=["POST"])
def set_default_master(master_id):
    if not session.get("master_unlocked"):
        return redirect(url_for("masters_page"))
    masters = load_masters()
    for m in masters:
        m["is_default"] = (m["id"] == master_id)
    save_masters(masters)
    return redirect(url_for("masters_manager"))

# ── Delete Master ──
@app.route("/masters/delete/<master_id>", methods=["POST"])
def delete_master(master_id):
    if not session.get("master_unlocked"):
        return redirect(url_for("masters_page"))
    masters = load_masters()
    master  = next((m for m in masters if m["id"] == master_id), None)
    if master and not master.get("is_default"):
        filepath = os.path.join("masters", master["filename"])
        if os.path.exists(filepath):
            os.remove(filepath)
        masters = [m for m in masters if m["id"] != master_id]
        save_masters(masters)
    return redirect(url_for("masters_manager"))

# ── Download Master ──
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

# ── Upload Updated Master ──
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

# ── API History ──
@app.route("/api/history")
def api_history():
    from flask import jsonify
    return jsonify({"history": load_history()})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)