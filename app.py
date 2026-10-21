from flask import Flask, jsonify, request, send_file, send_from_directory

from extractor import output_filename, pdf_to_xlsx_bytes

app = Flask(__name__, static_folder="static", static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


@app.get("/")
def index():
    return send_from_directory("templates", "index.html")


@app.post("/api/process")
def process():
    file = request.files.get("file")
    if file is None or not file.filename:
        return jsonify(error="Please upload a PDF file."), 400
    if not file.filename.lower().endswith(".pdf"):
        return jsonify(error="Only .pdf files are supported."), 400

    try:
        workbook, count = pdf_to_xlsx_bytes(file.stream)
    except ValueError as exc:
        return jsonify(error=str(exc)), 422
    except Exception:
        app.logger.exception("PDF processing failed")
        return jsonify(error="Could not read this PDF. Is it a Mondiale LCL rate sheet?"), 500

    response = send_file(
        workbook,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=output_filename(),
    )
    response.headers["X-Row-Count"] = str(count)
    return response


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
