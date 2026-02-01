from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
from pdf_processing.pdf_to_xlsx import extract_tables, pdf_to_excel
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)


@app.route("/upload", methods=["POST"])
def upload_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file"}), 400
    file = request.files["file"]

    if not file.filename.endswith(".pdf"):
        return jsonify({"error": "Only PDF allowed"}), 400

    file.save("pdf_processing/file.pdf")

    return jsonify({"message": "PDF uploaded successfully"})
    


@app.route("/convert", methods=["POST"])
def convert_pdf():
    pdf_path = "pdf_processing/file.pdf"
    extract_to_path = "pdf_processing/extracted"
    tables_path = "pdf_processing/extracted/tables"
    save_path = "pdf_processing/output.xlsx"
    template_path = "pdf_processing/template.xlsx"

    try:
        if not os.path.exists(pdf_path):
            return jsonify({"error": "File not found. Please upload again."}), 404


        extract_tables(pdf_path=pdf_path, extract_to=extract_to_path)
        pdf_to_excel(pdf_path=pdf_path, tables_path=tables_path, save_path=save_path, template_path=template_path)

        return jsonify({"message": "Conversion completed successfully! File is ready."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    

@app.route("/download", methods=["GET"])
def download_file():
    target_file = "pdf_processing/output.xlsx"
    if os.path.exists(target_file):
        return send_file(
            target_file,
            as_attachment=True, # Чтобы браузер именно скачивал файл
            download_name="result.xlsx" # Имя файла при скачивании
        )
    else:
        return jsonify({"error": "File not found"}), 404


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)