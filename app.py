import os
import io
import json
import base64
import zipfile
import requests
from flask import Flask, request, jsonify, send_file

app = Flask(__name__)
API_TOKEN = os.environ.get("API_TOKEN", "interpres-secret-2026")
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY", "")
CLOUDCONVERT_KEY = os.environ.get("CLOUDCONVERT_KEY", "")

TEMPLATE_URL = "https://raw.githubusercontent.com/postflowauto/interpres-expose/main/expose_template_v4.pptx"

PLATZHALTER = {
    "projekt_name": "", "projekt_titel": "", "entwickler_name": "", "entwickler_name_gross": "",
    "stadt": "", "stadt_gross": "", "stadtteil": "", "adresse_lang": "", "plz": "",
    "quartier_name": "", "quartier_history": "", "quartier_ref": "", "stadt_bezeichnung": "",
    "anzahl_we": "", "anzahl_1zi": "", "anzahl_2zi": "", "anzahl_barrierefrei": "",
    "groesse_von": "", "groesse_bis": "", "zimmer_typen": "", "produkt_beschreibung": "",
    "kaufpreis_ab": "", "kfw_darlehen": "150.000", "stellplaetze": "", "kfw_standard": "",
    "energieversorgung": "", "besonderheiten": "",
    "steuerliche_moeglichkeiten": "Dreifach AfA - 5% degressiv §7 Abs.5a EStG + 5% Sonder-AfA §7b EStG + 10% Möbel-AfA",
    "prospekt_datum": "", "text_kapitel_invest": "", "text_kapitel_live": "",
    "text_kapitel_stay": "", "text_kapitel_know": "", "text_intro": "",
    "text_investment_pitch": "", "text_hotel": "", "text_projekt_nachhaltig_1": "",
    "text_projekt_nachhaltig_2": "", "text_greenliving_intro": "", "text_greenliving_1": "",
    "text_greenliving_2": "", "text_ausstattung_intro": "", "text_ausstattung_detail": "",
    "text_ausstattung_kurz": "", "text_ausstattung_lang": "", "text_grundriss_intro": "",
    "text_architektur": "", "text_nachhaltig_1": "", "text_nachhaltig_2": "",
    "text_nachhaltig_3": "", "text_nachhaltig_4": "", "text_standort_1": "", "text_standort_2": "",
    "stadt_einwohner": "", "stadt_bip": "", "stadt_mietsteigerung": "", "stadt_studierende": "",
    "bundesland_bip": "", "text_einwohner_detail": "", "text_bip_detail": "",
    "text_mietsteigerung_detail": "", "text_studierende_detail": "",
    "text_stadt_wachstum_1": "", "text_stadt_wachstum_2": "", "text_stadt_intro": "",
    "text_stadt_wirtschaft_links": "", "text_stadt_wirtschaft_rechts": "",
    "stadt_invest_titel": "", "stadt_invest_label": "", "text_stadt_invest_detail": "",
    "stadt_stat_1_zahl": "", "stadt_stat_1_label": "", "stadt_stat_2_zahl": "",
    "stadt_stat_2_label": "", "stadt_stat_3_zahl": "", "stadt_stat_3_label": "",
    "stadt_branche_titel": "", "text_stadt_branche_1": "", "text_stadt_branche_2": "",
    "quelle_1": "", "quelle_2": "", "quelle_3": "", "quelle_4": "",
    "freizeit_1_name": "", "freizeit_2_name": "", "freizeit_3_name": "", "freizeit_4_name": "",
    "min_freizeit_1": "", "min_freizeit_2": "", "min_freizeit_3": "", "min_freizeit_4": "",
    "min_uni": "", "label_min_uni": "", "min_bahnhof": "", "label_min_bahnhof": "",
    "min_altstadt": "", "label_min_altstadt": "",
    "feature_1_zahl": "", "feature_1_label": "",
    "feature_2_zahl": "100", "feature_2_label": "Prozent möbliert",
    "feature_3_zahl": "24", "feature_3_label": "Stunden Zugang per Smart-Lock-System",
    "amenity_1": "", "amenity_2": "", "amenity_3": "", "amenity_4": "", "amenity_5": "",
    "amenity_6": "", "amenity_7": "", "amenity_8": "", "amenity_9": "",
    "grundriss_1_label": "", "grundriss_2_label": "", "grundriss_3_label": "", "grundriss_4_label": "",
    "ansicht_1_label": "", "ansicht_2_label": "",
    "we_bereich_1": "", "we_bereich_2": "", "we_beispiel_1": "", "we_beispiel_2": "",
    "we_typ_beschreibung": "", "we_flaeche_1": "", "we_flaeche_2": "",
    "we_flaeche_3": "", "we_flaeche_4": "", "we_flaeche_5": ""
}

def extract_pdfs_from_zip(zip_bytes):
    pdfs = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for name in zf.namelist():
            if name.lower().endswith('.pdf') and not name.startswith('__MACOSX'):
                data = zf.read(name)
                pdfs.append({"name": name, "base64": base64.b64encode(data).decode()})
    return pdfs

def analyze_pdfs_with_claude(pdfs):
    content = []
    for pdf in pdfs[:10]:
        content.append({
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": pdf["base64"]},
            "title": pdf["name"]
        })
    content.append({
        "type": "text",
        "text": "Analysiere diese Immobilien-Dokumente und extrahiere alle relevanten Projektdaten. "
                "Antworte NUR mit einem JSON-Objekt mit diesen Feldern: projektname_roh, adresse, "
                "stadt, stadtteil, plz, bautraeger, anzahl_haeuser, we_pro_haus, anzahl_we_gesamt, "
                "kfw_standard, energieversorgung, stellplaetze, groesse_von, groesse_bis, kaufpreis_ab, "
                "besonderheiten, planungsphase. Kein Text davor oder danach, keine Markdown-Backticks."
    })
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": CLAUDE_API_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"},
        json={"model": "claude-opus-4-5-20251101", "max_tokens": 4000, "messages": [{"role": "user", "content": content}]},
        timeout=120
    )
    resp.raise_for_status()
    text = resp.json()["content"][0]["text"]
    text = text.replace("```json", "").replace("```", "").strip()
    return json.loads(text)

def generate_expose_with_claude(projektdaten):
    prompt = (
        "Du bist ein Immobilien-Exposé-Spezialist bei INTERPRÉS GmbH. "
        "Antworte NUR mit einem validen JSON-Objekt. Kein Text davor oder danach. Keine Markdown-Backticks.\n\n"
        f"## PROJEKTDATEN\n{json.dumps(projektdaten, ensure_ascii=False)}\n\n"
        "## RECHERCHE\nNutze web_search für aktuelle Statistiken zur Stadt "
        f"{projektdaten.get('stadt', 'Magdeburg')}: Einwohnerzahl, BIP des Bundeslandes, "
        "Mietsteigerung, Studierende, Top-Arbeitgeber, Freizeiteinrichtungen mit Gehminuten "
        f"von {projektdaten.get('adresse', '')}.\n\n"
        f"## ALLE FELDER AUSFÜLLEN (kein Feld leer lassen):\n{json.dumps(PLATZHALTER, ensure_ascii=False)}"
    )
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": CLAUDE_API_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"},
        json={
            "model": "claude-opus-4-5-20251101",
            "max_tokens": 16000,
            "tools": [{"type": "web_search_20250305", "name": "web_search"}],
            "messages": [{"role": "user", "content": prompt}]
        },
        timeout=300
    )
    resp.raise_for_status()
    content_blocks = resp.json()["content"]
    json_text = ""
    for block in content_blocks:
        if block["type"] == "text":
            json_text = block["text"]
    json_text = json_text.replace("```json", "").replace("```", "").strip()
    return json.loads(json_text)

def fill_pptx(template_bytes, data):
    input_buf = io.BytesIO(template_bytes)
    output_buf = io.BytesIO()
    with zipfile.ZipFile(input_buf, "r") as zin:
        with zipfile.ZipFile(output_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                raw = zin.read(item.filename)
                if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
                    try:
                        content = raw.decode("utf-8")
                        for key, value in data.items():
                            safe = str(value or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                            content = content.replace("{{" + key.upper() + "}}", safe)
                            content = content.replace("{{" + key + "}}", safe)
                        raw = content.encode("utf-8")
                    except Exception:
                        pass
                zout.writestr(item, raw)
    return output_buf.getvalue()

def convert_to_pdf_cloudconvert(pptx_bytes, filename="expose.pptx"):
    # Job erstellen
    job_resp = requests.post(
        "https://sync.api.cloudconvert.com/v2/jobs",
        headers={"Authorization": f"Bearer {CLOUDCONVERT_KEY}", "Content-Type": "application/json"},
        json={"tasks": {
            "upload": {"operation": "import/upload"},
            "convert": {"operation": "convert", "input": "upload", "input_format": "pptx", "output_format": "pdf", "engine": "libreoffice"},
            "export": {"operation": "export/url", "input": "convert"}
        }},
        timeout=30
    )
    job_resp.raise_for_status()
    job = job_resp.json()["data"]
    upload_task = next(t for t in job["tasks"] if t["name"] == "upload")
    form = upload_task["result"]["form"]

    # PPTX hochladen
    files = {"file": (filename, pptx_bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
    up_resp = requests.post(form["url"], data=form.get("parameters", {}), files=files, timeout=60)
    up_resp.raise_for_status()

    # Job-Status abrufen
    status_resp = requests.get(
        f"https://sync.api.cloudconvert.com/v2/jobs/{job['id']}",
        headers={"Authorization": f"Bearer {CLOUDCONVERT_KEY}"},
        timeout=120
    )
    status_resp.raise_for_status()
    tasks = status_resp.json()["data"]["tasks"]
    export_task = next(t for t in tasks if t["name"] == "export")
    pdf_url = export_task["result"]["files"][0]["url"]

    pdf_resp = requests.get(pdf_url, timeout=60)
    pdf_resp.raise_for_status()
    return pdf_resp.content

@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "INTERPRES Full Pipeline"})

@app.route("/generate-expose", methods=["POST"])
def generate_expose():
    token = request.headers.get("X-API-Token", "")
    if token != API_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401

    try:
        # ZIP aus Request
        if request.content_type and "multipart" in request.content_type:
            zip_file = request.files.get("file")
            if not zip_file:
                return jsonify({"error": "Keine Datei im Request"}), 400
            zip_bytes = zip_file.read()
        else:
            zip_b64 = request.get_json().get("zip_base64")
            if not zip_b64:
                return jsonify({"error": "zip_base64 fehlt"}), 400
            zip_bytes = base64.b64decode(zip_b64)

        # 1. PDFs aus ZIP extrahieren
        pdfs = extract_pdfs_from_zip(zip_bytes)
        if not pdfs:
            return jsonify({"error": f"Keine PDFs im ZIP gefunden"}), 400

        # 2. PDFs mit Claude analysieren
        projektdaten = analyze_pdfs_with_claude(pdfs)

        # 3. Exposé-Texte mit Claude generieren
        expose_data = generate_expose_with_claude(projektdaten)

        # 4. Template von GitHub laden
        tmpl_resp = requests.get(TEMPLATE_URL, timeout=30)
        tmpl_resp.raise_for_status()
        template_bytes = tmpl_resp.content

        # 5. PPTX befüllen
        projekt_name = expose_data.get("projekt_name", "Expose").replace(" ", "_")
        pptx_bytes = fill_pptx(template_bytes, expose_data)

        # 6. CloudConvert: PPTX → PDF
        pdf_bytes = convert_to_pdf_cloudconvert(pptx_bytes, f"{projekt_name}.pptx")

        # 7. PDF zurückgeben
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"{projekt_name}_Expose.pdf"
        )

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
