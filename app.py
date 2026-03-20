import os
import json
import uuid
import datetime
from flask import Flask, request, jsonify, render_template, send_file, abort
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", os.urandom(24).hex())

ROLES_FILE = os.path.join(os.path.dirname(__file__), "roles.json")
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB


def load_roles():
    if not os.path.exists(ROLES_FILE):
        return []
    with open(ROLES_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_roles(roles):
    with open(ROLES_FILE, "w", encoding="utf-8") as f:
        json.dump(roles, f, ensure_ascii=False, indent=2)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/roles")
def get_roles():
    return jsonify(load_roles())


@app.route("/review", methods=["POST"])
def review():
    from reviewer import process_document

    role_id = request.form.get("role_id")
    if not role_id:
        return jsonify({"error": "role_id ist erforderlich"}), 400

    roles = load_roles()
    role = next((r for r in roles if r["id"] == role_id), None)
    if not role:
        return jsonify({"error": f"Rolle '{role_id}' nicht gefunden"}), 400

    text = None
    filename = None

    if "file" in request.files and request.files["file"].filename:
        file = request.files["file"]
        if file.content_length and file.content_length > MAX_FILE_SIZE:
            return jsonify({"error": "Datei zu groß (max. 10MB)"}), 413

        original_filename = file.filename
        ext = os.path.splitext(original_filename)[1].lower()
        if ext not in [".docx", ".txt"]:
            return jsonify({"error": "Nur DOCX und TXT Dateien werden unterstützt"}), 400

        tmp_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}{ext}")
        file.save(tmp_path)

        # Check file size after save
        if os.path.getsize(tmp_path) > MAX_FILE_SIZE:
            os.remove(tmp_path)
            return jsonify({"error": "Datei zu groß (max. 10MB)"}), 413

        try:
            if ext == ".docx":
                from docx import Document
                doc = Document(tmp_path)
                text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
                filename = original_filename
            else:
                with open(tmp_path, "r", encoding="utf-8", errors="replace") as f:
                    text = f.read()
                filename = original_filename
        finally:
            os.remove(tmp_path)

    elif request.form.get("text"):
        text = request.form.get("text").strip()
        filename = "eingabe.txt"
    else:
        return jsonify({"error": "Entweder eine Datei oder Text muss angegeben werden"}), 400

    if not text:
        return jsonify({"error": "Das Dokument enthält keinen Text"}), 400

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return jsonify({"error": "ANTHROPIC_API_KEY ist nicht gesetzt"}), 500

    try:
        output_path = process_document(text, role, UPLOAD_FOLDER)
    except Exception as e:
        return jsonify({"error": f"Verarbeitungsfehler: {str(e)}"}), 500

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_role_name = role["name"].replace(" ", "_").replace("/", "_")
    download_name = f"review_{safe_role_name}_{timestamp}.docx"

    response = send_file(
        output_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    @response.call_on_close
    def cleanup():
        try:
            os.remove(output_path)
        except Exception:
            pass

    return response


@app.route("/config")
def config():
    return render_template("config.html")


@app.route("/config/save", methods=["POST"])
def config_save():
    data = request.get_json()
    if not data or not isinstance(data, list):
        return jsonify({"error": "Ungültige Daten"}), 400

    for role in data:
        if not role.get("id") or not role.get("name"):
            return jsonify({"error": "Jede Rolle benötigt id und name"}), 400

    save_roles(data)
    return jsonify({"success": True})


@app.route("/config/delete", methods=["POST"])
def config_delete():
    data = request.get_json()
    role_id = data.get("id") if data else None
    if not role_id:
        return jsonify({"error": "id ist erforderlich"}), 400

    roles = load_roles()
    roles = [r for r in roles if r["id"] != role_id]
    save_roles(roles)
    return jsonify({"success": True})


if __name__ == "__main__":
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("WARNUNG: ANTHROPIC_API_KEY ist nicht gesetzt!")
    app.run(host="0.0.0.0", port=5000, debug=False)
