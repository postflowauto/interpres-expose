"""V2 Flask-Routen. Wird in app.py via register_v2(app) eingehängt."""
from __future__ import annotations
import io
from pathlib import Path
from flask import request, send_file, jsonify, send_from_directory, Response

from . import renderer, pipeline

V2_DIR     = Path(__file__).parent
STATIC_DIR = V2_DIR / "static"


def register_v2(app):
    """Registriert alle /v2/*-Routen an der bestehenden Flask-App."""

    @app.route("/v2/static/<path:filename>")
    def v2_static(filename):
        return send_from_directory(STATIC_DIR, filename)

    @app.route("/v2/render-sample.pdf")
    def v2_render_sample_pdf():
        """Demo-PDF mit Beispieldaten (kein ZIP nötig). Sofort sichtbar."""
        try:
            specs = pipeline.build_slide_specs(pipeline.sample_expose())
            pdf = renderer.render_to_pdf(specs)
            return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                             as_attachment=False, download_name="V2_Demo.pdf")
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/v2/render-sample.html")
    def v2_render_sample_html():
        """Demo-HTML im Browser (vor PDF-Render zum Inspizieren)."""
        try:
            specs = pipeline.build_slide_specs(pipeline.sample_expose())
            # TOC-Anker → Seitenzahl auflösen wie im Renderer
            anchor_to_page = {}
            for idx, spec in enumerate(specs):
                if spec.get("anchor"):
                    anchor_to_page[spec["anchor"]] = idx + 1
            for spec in specs:
                if spec.get("type") == "toc":
                    for col in spec["data"].get("chapters", []):
                        for it in col.get("items", []):
                            if it.get("anchor") in anchor_to_page:
                                it["page"] = anchor_to_page[it["anchor"]]
            html = renderer.render_html(specs, base_url="")
            return Response(html, mimetype="text/html")
        except Exception as e:
            import traceback
            return Response(f"<pre>{traceback.format_exc()}</pre>",
                            status=500, mimetype="text/html")

    @app.route("/v2/health")
    def v2_health():
        try:
            from playwright.sync_api import sync_playwright  # noqa
            playwright_ok = True
            chromium_ok   = None
            try:
                with sync_playwright() as p:
                    b = p.chromium.launch(args=["--no-sandbox"])
                    b.close()
                    chromium_ok = True
            except Exception as e:
                chromium_ok = f"FEHLER: {e}"
        except ImportError as e:
            playwright_ok = f"nicht installiert: {e}"
            chromium_ok   = None
        return jsonify({
            "v2": True,
            "playwright": playwright_ok,
            "chromium":   chromium_ok,
            "static_dir": str(STATIC_DIR),
            "static_exists": STATIC_DIR.exists(),
        })
