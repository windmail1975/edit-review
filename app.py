from flask import Flask, request, send_file, jsonify, render_template
import openpyxl
import os
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

EXCEL_PATH = "data/excel.xlsx"
PASSWORD = "your_password"

def init_excel():
    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "修改前"
        ws["B1"] = "修改後"
        ws["C1"] = "時間戳記"
        wb.save(EXCEL_PATH)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/admin")
def admin():
    return render_template("admin.html")

@app.route("/submit", methods=["POST"])
def submit():
    password = request.form.get("password")
    before = request.form.get("before")
    after = request.form.get("after")

    if password != PASSWORD:
        return jsonify({"error": "Wrong password"}), 401
    if not before or not after:
        return jsonify({"error": "Both fields required"}), 400

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

    existing = {(row[0].value, row[1].value) for row in ws.iter_rows(min_row=2, max_col=2) if row[0].value and row[1].value}
    if (before, after) in existing:
        return jsonify({"error": "Duplicate entry"}), 409

    ws.append([before, after, timestamp])
    wb.save(EXCEL_PATH)

    return jsonify({"message": "提交成功"})

@app.route("/upload_excel", methods=["POST"])
def upload_excel():
    password = request.form.get("password")
    file = request.files.get("excel_file")

    if password != PASSWORD:
        return jsonify({"error": "Wrong password"}), 401
    if not file:
        return jsonify({"error": "No file uploaded"}), 400

    filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    main_wb = openpyxl.load_workbook(EXCEL_PATH)
    main_ws = main_wb.active
    existing = {(row[0].value, row[1].value) for row in main_ws.iter_rows(min_row=2, max_col=2) if row[0].value and row[1].value}

    new_wb = openpyxl.load_workbook(filepath)
    new_ws = new_wb.active

    added_count = 0
    for row in new_ws.iter_rows(min_row=2, max_col=2, values_only=True):
        before, after = row
        if not before or not after:
            continue
        if (before, after) not in existing:
            timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            main_ws.append([before, after, timestamp])
            existing.add((before, after))
            added_count += 1

    main_wb.save(EXCEL_PATH)
    os.remove(filepath)

    return jsonify({"message": f"成功合併 {added_count} 筆資料"})

@app.route("/download", methods=["GET"])
def download_excel():
    return send_file(EXCEL_PATH, as_attachment=True)

@app.route("/clear", methods=["POST"])
def clear_excel():
    password = request.form.get("password")
    if password != PASSWORD:
        return jsonify({"error": "Wrong password"}), 401

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "修改前"
    ws["B1"] = "修改後"
    ws["C1"] = "時間戳記"
    wb.save(EXCEL_PATH)

    return jsonify({"message": "Cleared"})

if __name__ == "__main__":
    init_excel()
    app.run(host="0.0.0.0", port=5000)
