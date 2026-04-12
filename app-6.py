import os, io, json, base64, zipfile, requests, re
from flask import Flask, request, jsonify, send_file

app = Flask(__name__)

from flask import make_response

@app.after_request
def add_cors(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, X-API-Token"
    response.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    return response

@app.route("/generate-expose", methods=["OPTIONS"])
@app.route("/fill-pptx", methods=["OPTIONS"])
@app.route("/health", methods=["OPTIONS"])
def options():
    return make_response("", 204)

API_TOKEN = os.environ.get("API_TOKEN", "interpres-secret-2026")
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY", "")
CLOUDCONVERT_KEY = os.environ.get("CLOUDCONVERT_KEY", "")
UNSPLASH_ACCESS_KEY = os.environ.get("UNSPLASH_ACCESS_KEY", "")
TEMPLATE_URL = "https://raw.githubusercontent.com/postflowauto/interpres-expose/main/urbanunits_Marketing_Expose_v3.pdf-5.pptx"

# Relevante PDF-Typen nach Priorität
PDF_PRIORITY = [
    (1, ["zusammenfassung", "summary"]),
    (1, ["berechnung-bri", "bri-berechnung"]),
    (2, ["grundriss", "floor", "lageplan"]),
    (2, ["wfl-berechnung", "wohnflaeche", "wfl_berechnung"]),
    (3, ["schnitt", "ansicht", "elevation"]),
]

UNSPLASH_QUERIES = {
    "BILD_PROJEKT_AUSSEN": "modern apartment building exterior architecture",
    "BILD_AMENITY_1": "car sharing electric vehicle urban",
    "BILD_AMENITY_2": "solar panels rooftop renewable energy",
    "BILD_AMENITY_3": "gym fitness weights modern",
    "BILD_AMENITY_4": "parcel station locker delivery",
    "BILD_AMENITY_5": "cafe coffee interior modern",
    "BILD_AMENITY_6": "green roof garden urban",
    "BILD_AMENITY_7": "district heating pipes infrastructure",
    "BILD_AMENITY_8": "vintage classic car showroom",
    "BILD_AMENITY_9": "apartment balcony urban view",
    "BILD_GREENLIVING_1": "sustainable green building nature",
    "BILD_GREENLIVING_2": "modern residential building facade",
    "BILD_INTERIOR": "modern bedroom interior minimal design",
    "BILD_AUSSTATTUNG_1": "modern living room interior design",
    "BILD_AUSSTATTUNG_2": "hardwood floor interior",
    "BILD_AUSSTATTUNG_3": "modern bathroom tiles",
    "BILD_AUSSTATTUNG_4": "modern kitchen interior",
    "BILD_AUSSTATTUNG_5": "bedroom furniture design",
    "BILD_AUSSTATTUNG_6": "apartment interior detail",
    "BILD_GRUNDRISS_INTRO_1": "modern apartment living room",
    "BILD_GRUNDRISS_INTRO_2": "modern apartment bedroom",
    "BILD_ANSICHT_1": "apartment building exterior west",
    "BILD_ANSICHT_2": "modern residential building south",
    "BILD_WE_1": "apartment floor plan interior",
    "BILD_WE_2": "modern apartment bedroom interior",
    "BILD_STADT_PRESSE": "newspaper article table coffee",
    "BILD_STADT_BRANCHE": "scientist laboratory research modern",
    "BILD_RECHTLICH_1": "modern residential building exterior",
    "BILD_RECHTLICH_2": "apartment building facade evening",
}

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
    "we_flaeche_3": "", "we_flaeche_4": "", "we_flaeche_5": "",
    "bild_projekt_aussen": "", "bild_amenity_1": "", "bild_amenity_2": "", "bild_amenity_3": "",
    "bild_amenity_4": "", "bild_amenity_5": "", "bild_amenity_6": "", "bild_amenity_7": "",
    "bild_amenity_8": "", "bild_amenity_9": "", "bild_greenliving_1": "", "bild_greenliving_2": "",
    "bild_interior": "", "bild_ausstattung_1": "", "bild_ausstattung_2": "", "bild_ausstattung_3": "",
    "bild_ausstattung_4": "", "bild_ausstattung_5": "", "bild_ausstattung_6": "",
    "bild_grundriss_intro_1": "", "bild_grundriss_intro_2": "",
    "bild_ansicht_1": "", "bild_ansicht_2": "", "bild_we_1": "", "bild_we_2": "",
    "bild_stadt_presse": "", "bild_stadt_branche": "",
    "bild_rechtlich_1": "", "bild_rechtlich_2": "",
}

def get_pdf_priority(filename):
    name_lower = filename.lower()
    for priority, keywords in PDF_PRIORITY:
        if any(kw in name_lower for kw in keywords):
            return priority
    return 99

def extract_pdfs_from_zip(zip_bytes):
    """
    Extrahiert PDFs aus ZIP inkl. verschachtelter Ordner (Haus A/, Haus B/ etc.)
    Gibt max 20 relevanteste PDFs zurück.
    """
    all_pdfs = []

    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if not name.lower().endswith('.pdf'):
                    continue
                if '__MACOSX' in name or name.startswith('.'):
                    continue

                data = zf.read(name)
                if len(data) < 1000:
                    continue

                parts = name.split('/')
                folder = parts[-2] if len(parts) > 1 else "root"
                filename = parts[-1]
                priority = get_pdf_priority(filename)

                all_pdfs.append({
                    "name": filename,
                    "folder": folder,
                    "priority": priority,
                    "base64": base64.b64encode(data).decode(),
                })

    except Exception as e:
        print(f"ZIP Fehler: {e}")

    # Sortieren nach Priorität
    all_pdfs.sort(key=lambda x: (x["priority"], x["folder"]))

    # Auswahl: max 2 Prio-1 PDFs pro Ordner, max 1 Prio-2 pro Ordner, gesamt max 20
    selected = []
    folder_count = {}

    for pdf in all_pdfs:
        if len(selected) >= 20:
            break
        folder = pdf["folder"]
        prio = pdf["priority"]

        if prio == 99:
            continue

        key = f"{folder}_{prio}"
        folder_count[key] = folder_count.get(key, 0) + 1

        limit = 2 if prio == 1 else 1
        if folder_count[key] <= limit:
            selected.append(pdf)

    print(f"PDFs gesamt: {len(all_pdfs)}, ausgewählt: {len(selected)}")
    for p in selected:
        print(f"  [Prio {p['priority']}] {p['folder']} / {p['name']}")

    return selected

def fetch_unsplash_image(query):
    if not UNSPLASH_ACCESS_KEY:
        return ""
    try:
        resp = requests.get(
            "https://api.unsplash.com/photos/random",
            params={"query": query, "orientation": "landscape"},
            headers={"Authorization": f"Client-ID {UNSPLASH_ACCESS_KEY}"},
            timeout=10
        )
        if resp.status_code == 200:
            return resp.json()["urls"]["regular"]
    except Exception as e:
        print(f"Unsplash Fehler für '{query}': {e}")
    return ""

def fill_image_placeholders(data):
    for placeholder_key, query in UNSPLASH_QUERIES.items():
        data_key = placeholder_key.lower()
        if data_key in data:
            url = fetch_unsplash_image(query)
            data[data_key] = url
    return data

def analyze_pdfs_with_claude(pdfs):
    content = []
    for pdf in pdfs:
        content.append({
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": pdf["base64"]},
            "title": f"{pdf['folder']} / {pdf['name']}"
        })
    content.append({
        "type": "text",
        "text": (
            "Analysiere diese Immobilien-Dokumente aus dem Projektdatenraum und extrahiere alle relevanten Projektdaten. "
            "Es kann sein dass du Dokumente von mehreren Häusern (Haus A, B, C...) siehst - das ist ein Gesamtprojekt. "
            "Antworte NUR mit einem JSON-Objekt mit diesen Feldern: "
            "projektname_roh, adresse, stadt, stadtteil, plz, bautraeger, anzahl_haeuser, "
            "we_pro_haus, anzahl_we_gesamt, kfw_standard, energieversorgung, stellplaetze, "
            "groesse_von, groesse_bis, kaufpreis_ab, besonderheiten, planungsphase. "
            "Kein Text davor oder danach, keine Markdown-Backticks."
        )
    })
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": CLAUDE_API_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"},
        json={"model": "claude-opus-4-5-20251101", "max_tokens": 4000,
              "messages": [{"role": "user", "content": content}]},
        timeout=180
    )
    resp.raise_for_status()
    text = resp.json()["content"][0]["text"].replace("```json", "").replace("```", "").strip()
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
            "model": "claude-opus-4-5-20251101", "max_tokens": 16000,
            "tools": [{"type": "web_search_20250305", "name": "web_search"}],
            "messages": [{"role": "user", "content": prompt}]
        },
        timeout=300
    )
    resp.raise_for_status()
    json_text = ""
    for block in resp.json()["content"]:
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

def convert_to_pdf(pptx_bytes, filename):
    job_resp = requests.post(
        "https://sync.api.cloudconvert.com/v2/jobs",
        headers={"Authorization": f"Bearer {CLOUDCONVERT_KEY}", "Content-Type": "application/json"},
        json={"tasks": {
            "upload": {"operation": "import/upload"},
            "convert": {"operation": "convert", "input": "upload", "input_format": "pptx", "output_format": "pdf", "engine": "libreoffice"},
            "export": {"operation": "export/url", "input": "convert"}
        }}, timeout=30
    )
    job_resp.raise_for_status()
    job = job_resp.json()["data"]
    upload_task = next(t for t in job["tasks"] if t["name"] == "upload")
    form = upload_task["result"]["form"]
    files = {"file": (filename, pptx_bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
    requests.post(form["url"], data=form.get("parameters", {}), files=files, timeout=60).raise_for_status()
    status = requests.get(f"https://sync.api.cloudconvert.com/v2/jobs/{job['id']}",
                          headers={"Authorization": f"Bearer {CLOUDCONVERT_KEY}"}, timeout=120)
    status.raise_for_status()
    tasks = status.json()["data"]["tasks"]
    pdf_url = next(t for t in tasks if t["name"] == "export")["result"]["files"][0]["url"]
    return requests.get(pdf_url, timeout=60).content

@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "INTERPRES Full Pipeline v3"})

@app.route("/generate-expose", methods=["POST"])
def generate_expose():
    if request.headers.get("X-API-Token") != API_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401
    try:
        pdfs = []

        if request.content_type and "multipart" in request.content_type:
            uploaded = request.files.getlist("files") or request.files.getlist("file")
            if not uploaded:
                return jsonify({"error": "Keine Dateien im Request"}), 400
            for f in uploaded:
                pdfs.extend(extract_pdfs_from_zip(f.read()))
        else:
            body = request.get_json(force=True) or {}
            if "zip_base64_list" in body:
                for b64 in body["zip_base64_list"]:
                    pdfs.extend(extract_pdfs_from_zip(base64.b64decode(b64)))
            elif "zip_base64" in body:
                pdfs.extend(extract_pdfs_from_zip(base64.b64decode(body["zip_base64"])))
            else:
                return jsonify({"error": "zip_base64 oder zip_base64_list fehlt"}), 400

        if not pdfs:
            return jsonify({"error": "Keine relevanten PDFs gefunden"}), 400

        # Falls mehrere ZIPs hochgeladen: nochmal auf 20 begrenzen
        pdfs = sorted(pdfs, key=lambda x: x["priority"])[:20]

        projektdaten = analyze_pdfs_with_claude(pdfs)
        expose_data = generate_expose_with_claude(projektdaten)
        expose_data = fill_image_placeholders(expose_data)

        tmpl_bytes = requests.get(TEMPLATE_URL, timeout=30).content
        pptx_bytes = fill_pptx(tmpl_bytes, expose_data)

        projekt_name = expose_data.get("projekt_name", "Expose").replace(" ", "_")
        pdf_bytes = convert_to_pdf(pptx_bytes, f"{projekt_name}.pptx")

        return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf",
                         as_attachment=True, download_name=f"{projekt_name}_Expose.pdf")

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
