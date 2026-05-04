"""V2 Editor — schlanke Editor-Schicht über V1.

Konzept:
  - V1 macht die ganze schwere Arbeit (PDF-Analyse, Claude, Bilder, PPTX,
    PDF-Konvertierung, Slide-JPGs). Das Template kommt vom Designer.
  - V2 bietet eine schöne Editor-UI: Folien-Liste links, große Vorschau in
    der Mitte, Properties-Panel rechts mit allen editierbaren Texten und
    Bild-Slots der gerade ausgewählten Folie.
  - Edit → V2-API patcht state.json (expose_data) → triggert V1-Re-Render
    (PPTX neu füllen + PDF + Slide-JPGs) → Editor zeigt neue Vorschau.
  - "PDF herunterladen" → V1-Finalize (gleicher Pfad wie bisher).

Kein eigener Look mehr. V1-Pipeline + dein Template = Quelle der Wahrheit.
"""
from __future__ import annotations
import io
import os
import json
import threading
import time
from pathlib import Path
from flask import (request, send_file, jsonify, send_from_directory,
                   Response, redirect)

V2_DIR      = Path(__file__).parent
STATIC_DIR  = V2_DIR / "static"
EDITOR_HTML = V2_DIR / "editor.html"
JOB_DIR     = "/tmp/interpres_jobs"

# ── Per-Slide-Platzhalter-Scan (gecached) ───────────────────────────────────
# Liest das Original-Template einmal und merkt sich pro Slide alle {{KEY}}-
# Platzhalter (lowercase). Wird im Editor genutzt um pro Folie nur die
# relevanten Edit-Felder anzuzeigen.
_PER_SLIDE_PLACEHOLDERS = None


def _scan_template_placeholders(pptx_bytes: bytes) -> list:
    """Liefert pro Slide die Liste der Platzhalter-Keys (lowercase, dedupliziert).
    Robust gegen Whitespace-Splits, Line-Breaks innerhalb von {{...}} und |Xpt-Hints.
    """
    import io as _io
    import re as _re
    from pptx import Presentation
    PH = _re.compile(r"\{\{(.*?)\}\}", _re.DOTALL)

    def _extract(text):
        out = set()
        for m in PH.finditer(text):
            inner = m.group(1)
            if "|" in inner:
                inner = inner.split("|")[0]
            key = _re.sub(r"\s+", "", inner).lower().replace("-", "")
            if key and _re.match(r"^[a-z][a-z0-9_]*$", key):
                out.add(key)
        return out

    prs = Presentation(_io.BytesIO(pptx_bytes))
    result = []
    for slide in prs.slides:
        keys = set()
        def _scan_tf(tf):
            keys.update(_extract(tf.text))
        for shape in slide.shapes:
            try:
                if shape.has_text_frame:
                    _scan_tf(shape.text_frame)
                if shape.shape_type == 6:
                    for child in shape.shapes:
                        if child.has_text_frame:
                            _scan_tf(child.text_frame)
            except Exception:
                continue
        result.append(sorted(keys))
    return result


def _get_template_placeholders() -> list[list[str]]:
    """Lazy-loads + cached: pro Slide alle Template-Platzhalter."""
    global _PER_SLIDE_PLACEHOLDERS
    if _PER_SLIDE_PLACEHOLDERS is not None:
        return _PER_SLIDE_PLACEHOLDERS
    try:
        import importlib
        appmod = importlib.import_module("app")
        import requests
        tmpl_bytes = requests.get(appmod.TEMPLATE_URL, timeout=30).content
        _PER_SLIDE_PLACEHOLDERS = _scan_template_placeholders(tmpl_bytes)
        print(f"[v2] Template-Scan: {len(_PER_SLIDE_PLACEHOLDERS)} Slides, "
              f"{sum(len(s) for s in _PER_SLIDE_PLACEHOLDERS)} Platzhalter total")
    except Exception as e:
        print(f"[v2] Template-Scan Fehler: {e}")
        _PER_SLIDE_PLACEHOLDERS = []
    return _PER_SLIDE_PLACEHOLDERS


def _v1_state_path(job_id: str) -> str:
    return os.path.join(JOB_DIR, f"work_{job_id}", "state.json")


def _v1_meta_path(job_id: str) -> str:
    return os.path.join(JOB_DIR, f"{job_id}.json")


def _v1_slides_dir(job_id: str) -> str:
    return os.path.join(JOB_DIR, f"work_{job_id}", "slides")


def _v1_uploads_dir(job_id: str) -> str:
    return os.path.join(JOB_DIR, f"work_{job_id}", "uploads")


def _read_state(job_id: str) -> dict:
    """Liest state.json (expose_data + customer_images_files + projekt_name)."""
    with open(_v1_state_path(job_id)) as f:
        return json.load(f)


def _write_state(job_id: str, state: dict):
    path = _v1_state_path(job_id)
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(state, f, ensure_ascii=False)
    os.replace(tmp, path)


def _read_meta(job_id: str) -> dict:
    try:
        with open(_v1_meta_path(job_id)) as f:
            return json.load(f)
    except Exception:
        return {}


def _write_meta(job_id: str, **fields):
    path = _v1_meta_path(job_id)
    try:
        with open(path) as f:
            data = json.load(f)
    except Exception:
        data = {}
    data.update(fields)
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(data, f)
    os.replace(tmp, path)


# ── Slot-Labels (übernommen aus V1, hier im V2-Scope für Editor-Anzeige) ──
SLOT_LABELS = {
    "bild_titel":              "Titelbild (Außenansicht)",
    "bild_projekt_aussen":     "Projekt – Außenansicht",
    "bild_projekt":            "Projekt – Bild",
    "bild_quartier":           "Quartier / Umgebung",
    "bild_greenliving_1":      "Green Living – Bild 1",
    "bild_greenliving_2":      "Green Living – Bild 2",
    "bild_interior":           "Innenraum / Interior",
    "bild_ausstattung_1":      "Ausstattung 1",
    "bild_ausstattung_2":      "Ausstattung 2",
    "bild_ausstattung_3":      "Ausstattung 3",
    "bild_ausstattung_4":      "Ausstattung 4",
    "bild_ausstattung_5":      "Ausstattung 5",
    "bild_ausstattung_6":      "Ausstattung 6",
    "bild_ansicht_1":          "Außenansicht 1",
    "bild_ansicht_2":          "Außenansicht 2",
    "bild_standort_innen":     "Standort innen",
    "bild_standort_aussen":    "Standort außen",
    "bild_lageplan":           "Lageplan",
    "bild_stadt_gross":        "Stadtbild groß",
    "bild_stadt_klein":        "Stadtbild klein",
    "bild_grundriss_intro_1":  "Grundriss-Intro 1",
    "bild_grundriss_intro_2":  "Grundriss-Intro 2",
    "bild_grundriss_intro_3":  "Grundriss-Intro 3",
    "bild_grundriss_1":        "Grundriss 1",
    "bild_grundriss_2":        "Grundriss 2",
    "bild_grundriss_3":        "Grundriss 3",
    "bild_grundriss_4":        "Grundriss 4",
    "bild_collage_1":          "Collage 1",
    "bild_collage_2":          "Collage 2",
    "bild_collage_3":          "Collage 3",
    "bild_collage_4":          "Collage 4",
    "bild_collage_5":          "Collage 5",
    "bild_hotel_1":            "Hotel-Feeling 1",
    "bild_hotel_2":            "Hotel-Feeling 2",
    "bild_rechtlich_1":        "Rechtliches – Bild 1",
    "bild_rechtlich_2":        "Rechtliches – Bild 2",
    "bild_stadt_presse":       "Stadtbild – Presse",
    "bild_stadt_branche":      "Stadtbild – Branche",
}


def _slot_label(key: str) -> str:
    if key in SLOT_LABELS:
        return SLOT_LABELS[key]
    import re
    m = re.match(r"^bild_amenity_(\d+)$", key)
    if m:
        return f"Amenity {m.group(1)}"
    m = re.match(r"^bild_we_(\d+)$", key)
    if m:
        return f"WE-Bild {m.group(1)}"
    return key.replace("_", " ").title()


# Field-Gruppierung — welche expose_data-Keys gehören thematisch zusammen
FIELD_GROUPS = [
    {"name": "Projekt", "keys": [
        ("projekt_titel",         "Projekttitel"),
        ("entwickler_name",       "Entwicklername"),
        ("anzahl_we",             "Anzahl Wohneinheiten"),
        ("groesse_von",           "Größe von (m²)"),
        ("groesse_bis",           "Größe bis (m²)"),
        ("produkt_beschreibung",  "Produkt-Beschreibung"),
        ("kaufpreis_ab",          "Kaufpreis ab (€)"),
        ("kfw_standard",          "KfW-Standard"),
        ("kfw_darlehen",          "KfW-Darlehen (€)"),
        ("energieversorgung",     "Energieversorgung"),
        ("stellplaetze",          "Stellplätze"),
        ("zitat_intro",           "Intro-Zitat"),
    ]},
    {"name": "Stadt & Lage", "keys": [
        ("stadt",                 "Stadt"),
        ("stadtteil",             "Stadtteil"),
        ("adresse_lang",          "Adresse"),
        ("plz",                   "PLZ"),
        ("bundesland",            "Bundesland"),
        ("stadt_einwohner",       "Einwohner"),
        ("bundesland_bip",        "BIP Bundesland"),
        ("stadt_mietsteigerung",  "Mietsteigerung"),
        ("stadt_studierende",     "Studierende"),
    ]},
    {"name": "Texte – Investment", "keys": [
        ("text_intro",                "Intro-Text"),
        ("text_investment_pitch",     "Investment-Pitch"),
        ("text_kapitel_invest_1",     "Kapitel Invest – Lead 1"),
        ("text_kapitel_invest_2",     "Kapitel Invest – Lead 2"),
    ]},
    {"name": "Texte – Standort", "keys": [
        ("text_kapitel_live_1",       "Kapitel Live – Lead 1"),
        ("text_kapitel_live_2",       "Kapitel Live – Lead 2"),
        ("text_standort_1",           "Standort-Text 1"),
        ("text_standort_2",           "Standort-Text 2"),
        ("text_stadt_intro",          "Stadt-Intro"),
        ("text_stadt_wachstum_1",     "Stadt-Wachstum 1"),
        ("text_stadt_wachstum_2",     "Stadt-Wachstum 2"),
    ]},
    {"name": "Texte – Projekt & Stay", "keys": [
        ("text_kapitel_stay_1",       "Kapitel Stay – Lead 1"),
        ("text_kapitel_stay_2",       "Kapitel Stay – Lead 2"),
        ("text_greenliving_1",        "Greenliving 1"),
        ("text_greenliving_2",        "Greenliving 2"),
        ("text_ausstattung_kurz",     "Ausstattung – kurz"),
        ("text_ausstattung_detail",   "Ausstattung – Detail"),
        ("text_architektur",          "Architektur"),
    ]},
    {"name": "Min zu …", "keys": [
        ("min_uni",        "Min – Uni"),
        ("label_min_uni",  "Label – Uni"),
        ("min_bahnhof",    "Min – Bahnhof"),
        ("label_min_bahnhof", "Label – Bahnhof"),
        ("min_altstadt",   "Min – Altstadt"),
        ("label_min_altstadt", "Label – Altstadt"),
    ]},
]


def register_v2(app):
    """Registriert alle /v2/*-Routen an der Flask-App."""

    @app.route("/v2/static/<path:filename>")
    def v2_static(filename):
        return send_from_directory(STATIC_DIR, filename)

    @app.route("/v2/health")
    def v2_health():
        return jsonify({"v2": True, "mode": "editor-on-v1"})

    @app.route("/v2/from-job/<job_id>")
    def v2_from_job(job_id):
        """Öffnet den V2-Editor für einen V1-Job."""
        if not os.path.exists(_v1_state_path(job_id)):
            return Response("""<!doctype html>
<html><head><meta charset="utf-8"><title>Job nicht gefunden</title>
<style>
body{background:#0a1220;color:#e8d9b3;font-family:-apple-system,sans-serif;
     display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;}
.box{max-width:520px;text-align:center;padding:40px;}
h1{font-family:'Playfair Display',serif;color:#C8A96E;font-size:24px;margin-bottom:16px;}
p{color:#c8d1de;line-height:1.6;margin-bottom:24px;}
a.btn{display:inline-block;background:#C8A96E;color:#0a1220;text-decoration:none;
      padding:12px 28px;border-radius:6px;font-weight:600;letter-spacing:0.05em;}
a.btn:hover{background:#d4ba84;}
small{color:#6b7d96;font-size:11px;display:block;margin-top:30px;}
</style></head>
<body><div class="box">
<h1>Dieser Job ist nicht mehr verfügbar</h1>
<p>Der Server hat zwischenzeitlich neu gestartet oder der Job ist abgelaufen.
Erstelle einfach ein neues Exposé — du wirst automatisch in den Editor geleitet.</p>
<a class="btn" href="/">Neues Exposé generieren</a>
<small>Job-ID: """ + job_id + """</small>
</div></body></html>""",
                status=404, mimetype="text/html; charset=utf-8"
            )
        return redirect(f"/v2/editor/{job_id}")

    @app.route("/v2/editor/<job_id>")
    def v2_editor(job_id):
        if not os.path.exists(_v1_state_path(job_id)):
            return Response("Job nicht gefunden.", status=404)
        with open(EDITOR_HTML, encoding="utf-8") as f:
            html = f.read()
        # Job-ID + Token ins Frontend injizieren
        api_token = os.environ.get("API_TOKEN", "interpres-secret-2026")
        inject = (
            f"<script>"
            f"window.JOB_ID = {json.dumps(job_id)};"
            f"window.API_TOKEN = {json.dumps(api_token)};"
            f"</script>"
        )
        html = html.replace("<head>", "<head>\n" + inject)
        return Response(html, mimetype="text/html")

    # ── API: Job-State ───────────────────────────────────────────────────
    @app.route("/v2/api/job/<job_id>")
    def v2_api_job_get(job_id):
        if not os.path.exists(_v1_state_path(job_id)):
            return jsonify({"error": "not found"}), 404
        state = _read_state(job_id)
        meta  = _read_meta(job_id)
        expose = state.get("expose_data", {})
        slides_dir = _v1_slides_dir(job_id)
        slide_count = 0
        if os.path.isdir(slides_dir):
            slide_count = len([n for n in os.listdir(slides_dir)
                              if n.startswith("slide_") and n.endswith(".jpg")])

        # Bild-Slots aus expose_data extrahieren
        bild_slots = []
        for k, v in expose.items():
            if k.startswith("bild_"):
                bild_slots.append({
                    "key":      k,
                    "label":    _slot_label(k),
                    "value":    v if isinstance(v, str) else "",
                    "has_url":  bool(v and isinstance(v, str) and v.startswith("http")),
                })
        bild_slots.sort(key=lambda x: x["key"])

        # Hochgeladene Customer-Images
        uploads_dir = _v1_uploads_dir(job_id)
        uploaded = {}
        if os.path.isdir(uploads_dir):
            for fname in os.listdir(uploads_dir):
                slot = os.path.splitext(fname)[0].lower()
                uploaded[slot] = fname

        # Pro-Slide-Platzhalter aus dem Template (für Editor-Filter)
        slide_placeholders = _get_template_placeholders()

        return jsonify({
            "job_id":      job_id,
            "name":        meta.get("name") or expose.get("projekt_titel", "Expose"),
            "status":      meta.get("status", "unknown"),
            "phase":       meta.get("phase", ""),
            "slide_count": slide_count,
            "expose":      expose,
            "field_groups": FIELD_GROUPS,
            "bild_slots":  bild_slots,
            "uploaded":    uploaded,
            "slide_placeholders": slide_placeholders,
        })

    @app.route("/v2/api/job/<job_id>/text", methods=["PUT", "OPTIONS"])
    def v2_api_text_put(job_id):
        if request.method == "OPTIONS":
            return ("", 204)
        if not os.path.exists(_v1_state_path(job_id)):
            return jsonify({"error": "not found"}), 404
        body = request.get_json(force=True) or {}
        # body = { "key1": "value1", "key2": "value2", ... }
        state = _read_state(job_id)
        expose = state.get("expose_data", {})
        for k, v in body.items():
            # Akzeptiere nur Strings, keine Bild-URLs (die werden separat gehandhabt)
            if k.startswith("bild_"):
                continue
            expose[k] = v if isinstance(v, str) else str(v)
        state["expose_data"] = expose
        _write_state(job_id, state)
        return jsonify({"ok": True, "updated": list(body.keys())})

    @app.route("/v2/api/job/<job_id>/render", methods=["POST", "OPTIONS"])
    def v2_api_render(job_id):
        """Triggert V1-Re-Render: PPTX neu füllen mit aktueller expose_data
        + Customer-Uploads, dann PDF + Slide-JPGs erzeugen.
        Setzt status="processing" + phase, läuft im Background-Thread.
        """
        if request.method == "OPTIONS":
            return ("", 204)
        if not os.path.exists(_v1_state_path(job_id)):
            return jsonify({"error": "not found"}), 404

        # Use V1's own re-render logic by importing app.py's _run_finalize_job
        # → das macht aber nur PDF, keine slide JPGs. Wir machen unseren eigenen
        #   Worker der BEIDE produziert (für Editor-Vorschau + Download).
        _write_meta(job_id, status="processing", phase="V2-Render läuft …")
        t = threading.Thread(target=_v2_render_worker, args=(job_id,), daemon=True)
        t.start()
        return jsonify({"ok": True})

    @app.route("/v2/api/job/<job_id>/render-status")
    def v2_api_render_status(job_id):
        meta = _read_meta(job_id)
        return jsonify({
            "status": meta.get("status", "unknown"),
            "phase":  meta.get("phase", ""),
            "error":  meta.get("error", None),
        })

    @app.route("/v2/api/job/<job_id>/download")
    def v2_api_download(job_id):
        """Triggert Render falls nötig, liefert PDF zurück."""
        meta = _read_meta(job_id)
        pdf_path = meta.get("pdf_path")
        if not pdf_path or not os.path.exists(pdf_path):
            return jsonify({
                "error": "Kein PDF vorhanden – erst 'Aktualisieren' klicken."
            }), 409
        ext = os.path.splitext(pdf_path)[1].lower()
        if ext == ".pdf":
            mt = "application/pdf"
            name = f"{meta.get('name', 'Expose')}.pdf"
        else:
            mt = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            name = f"{meta.get('name', 'Expose')}.pptx"
        return send_file(pdf_path, mimetype=mt, as_attachment=True, download_name=name)


def _v2_render_worker(job_id: str):
    """Background-Worker: lädt expose_data + Customer-Images, ruft V1-fill_pptx,
    konvertiert zu PDF, rendert Slide-JPGs."""
    try:
        # V1-Funktionen importieren (Lazy, um Zirkular-Import zu vermeiden)
        import importlib
        appmod = importlib.import_module("app")

        _write_meta(job_id, status="processing", phase="State + Bilder werden geladen …")
        state = _read_state(job_id)
        expose = state.get("expose_data", {})
        cust_files = state.get("customer_images_files", {}) or {}

        customer_images = {}
        for slot, fpath in cust_files.items():
            try:
                with open(fpath, "rb") as f:
                    customer_images[slot] = f.read()
            except Exception:
                pass

        # User-Uploads übersteuern Auto-Zuweisungen
        uploads_dir = _v1_uploads_dir(job_id)
        if os.path.isdir(uploads_dir):
            for fname in os.listdir(uploads_dir):
                base, ext = os.path.splitext(fname)
                slot = base.lower()
                if not slot.startswith("bild_"):
                    continue
                with open(os.path.join(uploads_dir, fname), "rb") as f:
                    customer_images[slot] = f.read()
                expose[slot] = ""

        _write_meta(job_id, phase="Template wird gefüllt …")
        import requests
        tmpl_url = appmod.TEMPLATE_URL
        tmpl_bytes = requests.get(tmpl_url, timeout=30).content
        pptx_bytes = appmod.fill_pptx(tmpl_bytes, expose, customer_images=customer_images)

        _write_meta(job_id, phase="PDF wird konvertiert …")
        pdf_bytes = None
        if appmod._can_convert_to_pdf():
            try:
                projekt_name = expose.get("projekt_titel", "Expose").replace(" ", "_")
                pdf_bytes = appmod.convert_to_pdf(pptx_bytes, f"{projekt_name}.pptx")
            except Exception as e:
                print(f"[v2] PDF-Konvertierung Fehler: {e}")

        # Output-Pfad (PDF wenn möglich, sonst PPTX)
        if pdf_bytes:
            out_path = os.path.join(JOB_DIR, f"{job_id}.pdf")
            with open(out_path, "wb") as f:
                f.write(pdf_bytes)

            _write_meta(job_id, phase="Slide-Vorschau wird erstellt …")
            try:
                slides_dir = _v1_slides_dir(job_id)
                # Alte JPGs löschen
                if os.path.isdir(slides_dir):
                    for fname in os.listdir(slides_dir):
                        if fname.endswith(".jpg"):
                            try: os.remove(os.path.join(slides_dir, fname))
                            except OSError: pass
                appmod.render_pdf_to_jpgs(pdf_bytes, slides_dir, dpi=150)
            except Exception as e:
                print(f"[v2] Slide-JPG-Render Fehler: {e}")
        else:
            out_path = os.path.join(JOB_DIR, f"{job_id}.pptx")
            with open(out_path, "wb") as f:
                f.write(pptx_bytes)

        projekt_name = expose.get("projekt_titel", "Expose")
        _write_meta(job_id,
                    status="preview",
                    phase="Vorschau aktualisiert",
                    pdf_path=out_path,
                    name=projekt_name)
        print(f"[v2] ✓ Render fertig für Job {job_id}")
    except Exception as e:
        import traceback as tb
        err = f"{e}\n{tb.format_exc()}"
        print(f"[v2] ✗ Render Fehler: {err[:500]}")
        _write_meta(job_id, status="error", phase="Fehler", error=str(e))