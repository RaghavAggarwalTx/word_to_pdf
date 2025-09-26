from flask import Flask, request, send_file
import subprocess
import os
import uuid

app = Flask(__name__)

@app.route("/convert", methods=["POST"])
def convert_docx_to_pdf():
    if "file" not in request.files:
        return {"error": "No file uploaded"}, 400

    uploaded_file = request.files["file"]
    input_filename = f"/tmp/{uuid.uuid4()}.docx"
    output_filename = input_filename.replace(".docx", ".pdf")
    uploaded_file.save(input_filename)

    try:
        subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf",
            "--outdir", "/tmp", input_filename
        ], check=True)
    except subprocess.CalledProcessError:
        return {"error": "Conversion failed"}, 500

    return send_file(output_filename, as_attachment=True, download_name="converted.pdf")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
