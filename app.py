import os
import uuid
from flask import Flask, request, render_template, send_file, jsonify
from formatter import format_document, TEMPLATES
from template_parser import parse_template_from_docx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

custom_templates = {}


@app.route("/")
def index():
    templates_info = {
        k: {"name": v["name"], "description": v["description"]}
        for k, v in TEMPLATES.items()
    }
    return render_template("index.html", templates=templates_info)


@app.route("/parse-template", methods=["POST"])
def parse_template():
    if "file" not in request.files:
        return jsonify({"error": "未上传文件"}), 400

    file = request.files["file"]
    if not file.filename or not file.filename.endswith(".docx"):
        return jsonify({"error": "请上传 .docx 格式的格式要求文档"}), 400

    file_id = str(uuid.uuid4())[:8]
    tmp_path = os.path.join(UPLOAD_DIR, f"{file_id}_tpl.docx")
    file.save(tmp_path)

    try:
        template, raw_json = parse_template_from_docx(tmp_path)
        custom_templates[file_id] = template
        summary = {}
        for role in ("title", "heading1", "heading2", "heading3", "body"):
            info = raw_json.get(role)
            if info:
                summary[role] = {
                    "font": info.get("font_cn", ""),
                    "size": info.get("size", ""),
                    "bold": info.get("bold", False),
                    "alignment": info.get("alignment", ""),
                    "line_spacing": info.get("line_spacing", ""),
                }
        return jsonify({"template_id": file_id, "parsed": summary})
    except Exception as e:
        return jsonify({"error": f"解析失败: {str(e)}"}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.route("/format", methods=["POST"])
def format_file():
    if "file" not in request.files:
        return jsonify({"error": "未上传文件"}), 400

    file = request.files["file"]
    if not file.filename or not file.filename.endswith(".docx"):
        return jsonify({"error": "请上传 .docx 格式的Word文件"}), 400

    template_key = request.form.get("template", "通用论文")
    custom_id = request.form.get("custom_template_id", "")

    if custom_id and custom_id in custom_templates:
        tpl_override = custom_templates[custom_id]
    elif template_key in TEMPLATES:
        tpl_override = None
    else:
        return jsonify({"error": "未知模板"}), 400

    file_id = str(uuid.uuid4())[:8]
    original_name = os.path.splitext(file.filename)[0]
    input_path = os.path.join(UPLOAD_DIR, f"{file_id}_input.docx")
    output_name = f"{original_name}_已排版.docx"
    output_path = os.path.join(OUTPUT_DIR, f"{file_id}_output.docx")

    file.save(input_path)

    try:
        result = format_document(
            input_path, output_path,
            template_key=template_key,
            custom_template=tpl_override,
        )
    except Exception as e:
        return jsonify({"error": f"排版处理失败: {str(e)}"}), 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

    result["file_id"] = file_id
    result["output_name"] = output_name
    return jsonify(result)


@app.route("/download/<file_id>/<filename>")
def download(file_id, filename):
    output_path = os.path.join(OUTPUT_DIR, f"{file_id}_output.docx")
    if not os.path.exists(output_path):
        return jsonify({"error": "文件不存在或已过期"}), 404
    return send_file(output_path, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
