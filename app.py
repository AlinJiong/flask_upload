from flask import Flask, request, send_file, render_template
import openpyxl
import os

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"  # 上传文件的文件夹
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)  # 创建文件夹（如果不存在）


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "没有文件上传", 400

    file = request.files["file"]
    if file.filename == "":
        return "没有选择文件", 400

    if file and file.filename.endswith(".xlsx"):
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
        file.save(file_path)

        # 处理 Excel 文件
        output_file_path = process_excel(file_path)

        return send_file(output_file_path, as_attachment=True)

    return "文件格式不正确", 400


def process_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    for row in sheet.iter_rows(min_row=2):  # 从第二行开始
        card_number = row[1].value  # 假设卡号在第二列
        if card_number:
            row[1].value = card_number.replace(" ", "")  # 去除空格

    output_file_path = os.path.join(app.config["UPLOAD_FOLDER"], "processed_" + os.path.basename(file_path))
    workbook.save(output_file_path)
    return output_file_path


if __name__ == "__main__":
    app.run(debug=True)
