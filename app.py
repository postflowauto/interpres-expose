import os
import io
import json
import base64
import zipfile
import re
from flask import Flask, request, jsonify, send_file

app = Flask(__name__)

ALLOWED_TOKEN = os.environ.get("API_TOKEN", "interpres-secret-2026")

def replace_placeholders(xml_content: str, data: dict) -> str:
    """Ersetzt alle {{PLATZHALTER}} im XML mit den echten Werten."""
    for key, value in data.items():
        safe_value = str(value or "")
        # XML-Sonderzeichen escapen
        safe_value = safe_value.replace("&", "&amp;")
        safe_value = safe_value.replace("<", "&lt;")
        safe_value = safe_value.replace(">", "&gt;")
        safe_value = safe_value.replace('"', "&quot;")
        safe_value = safe_value.replace("'", "&apos;")

        # Uppercase Variante: {{PROJEKT_NAME}}
        xml_content = xml_content.replace("{{" + key.upper() + "}}", safe_value)
        # Lowercase Variante: {{projekt_name}}
        xml_content = xml_content.replace("{{" + key.lower() + "}}", safe_value)
        # Gemischte Variante: {{Projekt_Name}}
        xml_content = xml_content.replace("{{" + key + "}}", safe_value)

    return xml_content

def fill_pptx(template_bytes: bytes, data: dict) -> bytes:
    """Füllt ein PPTX-Template mit Daten und gibt die befüllte PPTX zurück."""
    input_buffer = io.BytesIO(template_bytes)
    output_buffer = io.BytesIO()

    with zipfile.ZipFile(input_buffer, "r") as zin:
        with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data_bytes = zin.read(item.filename)

                # Nur XML/rels Dateien bearbeiten
                if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
                    try:
                        content = data_bytes.decode("utf-8")
                        content = replace_placeholders(content, data)
                        data_bytes = content.encode("utf-8")
                    except UnicodeDecodeError:
                        pass  # Binärdateien unverändert lassen

                zout.writestr(item, data_bytes)

    return output_buffer.getvalue()

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "INTERPRES PPTX Filler"})

@app.route("/fill-pptx", methods=["POST"])
def fill_pptx_endpoint():
    # Auth prüfen
    token = request.headers.get("X-API-Token", "")
    if token != ALLOWED_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401

    try:
        body = request.get_json(force=True)
        if not body:
            return jsonify({"error": "Kein JSON Body"}), 400

        # Template als Base64
        template_b64 = body.get("template_base64")
        if not template_b64:
            return jsonify({"error": "template_base64 fehlt"}), 400

        # Exposé-Daten
        expose_data = body.get("expose_data")
        if not expose_data:
            return jsonify({"error": "expose_data fehlt"}), 400

        # Dateiname optional
        filename = body.get("filename", "INTERPRES_Expose.pptx")

        # Template dekodieren
        template_bytes = base64.b64decode(template_b64)

        # PPTX befüllen
        filled_bytes = fill_pptx(template_bytes, expose_data)

        # Als Base64 zurückgeben
        result_b64 = base64.b64encode(filled_bytes).decode("utf-8")

        return jsonify({
            "success": True,
            "pptx_base64": result_b64,
            "filename": filename,
            "size_bytes": len(filled_bytes)
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/fill-pptx-binary", methods=["POST"])
def fill_pptx_binary():
    """Alternativ-Endpoint: gibt die PPTX direkt als Binary zurück."""
    token = request.headers.get("X-API-Token", "")
    if token != ALLOWED_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401

    try:
        body = request.get_json(force=True)
        template_bytes = base64.b64decode(body["template_base64"])
        expose_data = body["expose_data"]
        filename = body.get("filename", "INTERPRES_Expose.pptx")

        filled_bytes = fill_pptx(template_bytes, expose_data)

        return send_file(
            io.BytesIO(filled_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
