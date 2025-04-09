from flask import Flask, request, send_file, jsonify, render_template
import openpyxl
import os
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

EXCEL_PATH = "data/excel.xlsx"

USER_PASSWORD = os.environ.get("USER_PASSWORD")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")

def init_excel():
    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["類別", "修改前", "修改後", "時間戳記"])
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
    category = request.form.get("category")
    before = request.form.get("before")
    after = request.form.get("after")

    if password != USER_PASSWORD:
        return jsonify({"error": "使用者密碼錯誤"}), 401
    if not category or not before or not after:
        return jsonify({"error": "請完整填寫類別、修改前、修改後"}), 400

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

    existing = {(row[0].value, row[1].value, row[2].value) for row in ws.iter_rows(min_row=2, max_col=3)}
    if (category, before, after) in existing:
        return jsonify({"error": "已存在相同資料"}), 409

    ws.append([category, before, after, timestamp])
    wb.save(EXCEL_PATH)

    return jsonify({"message": "提交成功"})

@app.route("/upload_excel", methods=["POST"])
def upload_excel():
    password = request.form.get("password")
    file = request.files.get("excel_file")

    if password != USER_PASSWORD:
        return jsonify({"error": "使用者密碼錯誤"}), 401
    if not file:
        return jsonify({"error": "請選擇檔案"}), 400

    filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    main_wb = openpyxl.load_workbook(EXCEL_PATH)
    main_ws = main_wb.active
    existing = {(row[0].value, row[1].value, row[2].value) for row in main_ws.iter_rows(min_row=2, max_col=3)}

    new_wb = openpyxl.load_workbook(filepath)
    new_ws = new_wb.active

    added_count = 0
    for row in new_ws.iter_rows(min_row=2, max_col=3, values_only=True):
        category, before, after = row
        if not category or not before or not after:
            continue
        if (category, before, after) not in existing:
            timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            main_ws.append([category, before, after, timestamp])
            existing.add((category, before, after))
            added_count += 1

    main_wb.save(EXCEL_PATH)
    os.remove(filepath)

    return jsonify({"message": f"成功合併 {added_count} 筆資料"})

@app.route("/download", methods=["POST"])
def download_excel():
    password = request.form.get("password")
    if password != ADMIN_PASSWORD:
        return jsonify({"error": "管理者密碼錯誤"}), 401
    return send_file(EXCEL_PATH, as_attachment=True)

@app.route("/clear", methods=["POST"])
def clear_excel():
    password = request.form.get("password")
    if password != ADMIN_PASSWORD:
        return jsonify({"error": "管理者密碼錯誤"}), 401

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["類別", "修改前", "修改後", "時間戳記"])
    wb.save(EXCEL_PATH)

    return jsonify({"message": "資料已清空"})

if __name__ == "__main__":
    init_excel()
    app.run(host="0.0.0.0", port=5000)
