import os
import json
import uuid
import datetime
import logging
import logging.handlers
from flask import Flask, request, jsonify, render_template, send_file, abort
from dotenv import load_dotenv

load_dotenv()

# --- Logging setup ---
LOG_DIR = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.handlers.RotatingFileHandler(
            os.path.join(LOG_DIR, "app.log"),
            maxBytes=5 * 1024 * 1024,  # 5 MB
            backupCount=5,
            encoding="utf-8",
        ),
    ],
)
logger = logging.getLogger("dokureview.app")

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
        logger.warning("Review-Anfrage ohne role_id abgelehnt")
        return jsonify({"error": "role_id ist erforderlich"}), 400

    roles = load_roles()
    role = next((r for r in roles if r["id"] == role_id), None)
    if not role:
        logger.warning("Unbekannte Rolle angefragt: %s", role_id)
        return jsonify({"error": f"Rolle '{role_id}' nicht gefunden"}), 400

    text = None
    filename = None

    if "file" in request.files and request.files["file"].filename:
        file = request.files["file"]
        if file.content_length and file.content_length > MAX_FILE_SIZE:
            logger.warning("Datei abgelehnt (zu groß): %s", file.filename)
            return jsonify({"error": "Datei zu groß (max. 10MB)"}), 413

        original_filename = file.filename
        ext = os.path.splitext(original_filename)[1].lower()
        if ext not in [".docx", ".txt"]:
            logger.warning("Nicht unterstütztes Dateiformat: %s", ext)
            return jsonify({"error": "Nur DOCX und TXT Dateien werden unterstützt"}), 400

        tmp_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}{ext}")
        file.save(tmp_path)

        # Check file size after save
        actual_size = os.path.getsize(tmp_path)
        if actual_size > MAX_FILE_SIZE:
            os.remove(tmp_path)
            logger.warning("Datei nach dem Speichern zu groß: %s (%d Bytes)", original_filename, actual_size)
            return jsonify({"error": "Datei zu groß (max. 10MB)"}), 413

        logger.info("Datei empfangen: %s (%d Bytes), Rolle: %s", original_filename, actual_size, role_id)

        try:
            if ext == ".docx":
                from docx import Document
                source_doc = Document(tmp_path)
                text = "\n".join(p.text for p in source_doc.paragraphs if p.text.strip())
                filename = original_filename
            else:
                source_doc = None
                with open(tmp_path, "r", encoding="utf-8", errors="replace") as f:
                    text = f.read()
                filename = original_filename
        finally:
            os.remove(tmp_path)

    elif request.form.get("text"):
        text = request.form.get("text").strip()
        source_doc = None
        filename = "eingabe.txt"
        logger.info("Texteingabe empfangen (%d Zeichen), Rolle: %s", len(text), role_id)
    else:
        logger.warning("Review-Anfrage ohne Inhalt (keine Datei, kein Text)")
        return jsonify({"error": "Entweder eine Datei oder Text muss angegeben werden"}), 400

    if not text:
        logger.warning("Dokument enthält keinen Text: %s", filename)
        return jsonify({"error": "Das Dokument enthält keinen Text"}), 400

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        logger.error("ANTHROPIC_API_KEY ist nicht gesetzt")
        return jsonify({"error": "ANTHROPIC_API_KEY ist nicht gesetzt"}), 500

    logger.info("Starte Review: Rolle='%s', Datei='%s', Textlänge=%d", role["name"], filename, len(text))
    try:
        output_path = process_document(text, role, UPLOAD_FOLDER, source_doc=source_doc)
    except Exception as e:
        logger.exception("Verarbeitungsfehler für Datei '%s', Rolle '%s': %s", filename, role_id, e)
        return jsonify({"error": f"Verarbeitungsfehler: {str(e)}"}), 500

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_role_name = role["name"].replace(" ", "_").replace("/", "_")
    download_name = f"review_{safe_role_name}_{timestamp}.docx"
    logger.info("Review abgeschlossen: Download='%s'", download_name)

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
            logger.debug("Temporäre Output-Datei gelöscht: %s", output_path)
        except Exception as e:
            logger.warning("Fehler beim Löschen der temporären Datei %s: %s", output_path, e)

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
    logger.info("Rollen gespeichert (%d Einträge)", len(data))
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
    logger.info("Rolle gelöscht: %s", role_id)
    return jsonify({"success": True})


if __name__ == "__main__":
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        logger.warning("ANTHROPIC_API_KEY ist nicht gesetzt — Reviews werden fehlschlagen!")
    else:
        logger.info("ANTHROPIC_API_KEY gefunden, App startet auf http://0.0.0.0:5000")
    app.run(host="0.0.0.0", port=5000, debug=False)
