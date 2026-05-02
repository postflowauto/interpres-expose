"""V2 Flask-Routen. Wird in app.py via register_v2(app) eingehängt.

Endpoints:
  /v2/health                                 → Diagnose
  /v2/static/<file>                          → CSS/JS-Assets
  /v2/render-sample.html                     → Demo-HTML
  /v2/render-sample.pdf                      → Demo-PDF
  /v2/editor/<project_id>                    → Editor-UI (HTML)
  /v2/api/project/<pid>                      → GET project (specs+state)
  /v2/api/project/<pid>/slide/<idx>          → PUT slide.data
  /v2/api/project/<pid>/upload               → POST file → URL
  /v2/api/project/<pid>/preview.html         → live HTML-Vorschau
  /v2/api/project/<pid>/render.pdf           → finales PDF
  /v2/api/project/<pid>/uploads/<filename>   → uploaded image
  /v2/demo-project                           → erstellt Demo-Projekt + redirect
"""
from __future__ import annotations
import io
import os
import json
from pathlib import Path
from flask import (request, send_file, jsonify, send_from_directory,
                   Response, redirect, url_for)

from . import renderer, pipeline, db

V2_DIR        = Path(__file__).parent
STATIC_DIR    = V2_DIR / "static"
EDITOR_HTML   = V2_DIR / "editor.html"
UPLOADS_BASE  = Path("/tmp/interpres_v2_uploads")
UPLOADS_BASE.mkdir(exist_ok=True)


def _resolve_anchors(specs: list[dict]) -> None:
    """Setzt TOC-Item-Pages basierend auf Anker-Reihenfolge (in-place)."""
    anchor_to_page = {}
    for idx, spec in enumerate(specs):
        if spec.get("anchor"):
            anchor_to_page[spec["anchor"]] = idx + 1
    for spec in specs:
        if spec.get("type") != "toc":
            continue
        for col in spec["data"].get("chapters", []):
            for it in col.get("items", []):
                if it.get("anchor") in anchor_to_page:
                    it["page"] = anchor_to_page[it["anchor"]]


def register_v2(app):
    """Registriert alle /v2/*-Routen an der bestehenden Flask-App."""

    @app.route("/v2/static/<path:filename>")
    def v2_static(filename):
        return send_from_directory(STATIC_DIR, filename)

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

    # ── Demo / Sample ────────────────────────────────────────────────────
    @app.route("/v2/render-sample.html")
    def v2_render_sample_html():
        try:
            specs = pipeline.build_slide_specs(pipeline.sample_expose())
            _resolve_anchors(specs)
            html = renderer.render_html(specs, base_url="")
            return Response(html, mimetype="text/html")
        except Exception as e:
            import traceback
            return Response(f"<pre>{traceback.format_exc()}</pre>",
                            status=500, mimetype="text/html")

    @app.route("/v2/render-sample.pdf")
    def v2_render_sample_pdf():
        try:
            specs = pipeline.build_slide_specs(pipeline.sample_expose())
            pdf = renderer.render_to_pdf(specs)
            return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                             as_attachment=False, download_name="V2_Demo.pdf")
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/v2/demo-project")
    def v2_demo_project():
        """Erstellt ein Demo-Projekt mit Beispieldaten und leitet zum Editor."""
        expose = pipeline.sample_expose()
        specs = pipeline.build_slide_specs(expose)
        _resolve_anchors(specs)
        pid = db.create_project(specs, expose=expose,
                                name=expose.get("projekt_titel", "Demo"))
        return redirect(f"/v2/editor/{pid}")

    # ── Editor UI ────────────────────────────────────────────────────────
    @app.route("/v2/editor/<pid>")
    def v2_editor(pid):
        proj = db.get_project(pid)
        if not proj:
            return Response(
                "Projekt nicht gefunden. Lege ein neues an unter /v2/demo-project",
                status=404, mimetype="text/plain"
            )
        with open(EDITOR_HTML, encoding="utf-8") as f:
            html = f.read()
        # Inject project id
        html = html.replace("<head>",
                            f"<head>\n<script>window.PROJECT_ID = {json.dumps(pid)};</script>")
        return Response(html, mimetype="text/html")

    # ── API: Project CRUD ────────────────────────────────────────────────
    @app.route("/v2/api/project/<pid>")
    def v2_api_project_get(pid):
        proj = db.get_project(pid)
        if not proj:
            return jsonify({"error": "not found"}), 404
        return jsonify(proj)

    @app.route("/v2/api/project/<pid>/slide/<int:idx>", methods=["PUT", "OPTIONS"])
    def v2_api_slide_put(pid, idx):
        if request.method == "OPTIONS":
            return ("", 204)
        data = request.get_json(force=True) or {}
        if not db.update_slide(pid, idx, data):
            return jsonify({"error": "update failed"}), 400
        # Re-resolve TOC anchors falls sich Slide-Reihenfolge ändert
        proj = db.get_project(pid)
        _resolve_anchors(proj["specs"])
        db.update_specs(pid, proj["specs"])
        return jsonify({"ok": True})

    @app.route("/v2/api/project/<pid>/upload", methods=["POST", "OPTIONS"])
    def v2_api_upload(pid):
        if request.method == "OPTIONS":
            return ("", 204)
        f = request.files.get("file")
        if not f or not f.filename:
            return jsonify({"error": "no file"}), 400
        ext = os.path.splitext(f.filename)[1].lower() or ".jpg"
        if ext not in (".jpg", ".jpeg", ".png", ".webp"):
            return jsonify({"error": "unsupported file type"}), 400
        proj_dir = UPLOADS_BASE / pid
        proj_dir.mkdir(exist_ok=True)
        # Filename: zufällig damit Cache-bust funktioniert
        import uuid as _uuid
        fname = f"{_uuid.uuid4().hex}{ext}"
        target = proj_dir / fname
        f.save(target)
        # Server-relative URL
        url = f"/v2/api/project/{pid}/uploads/{fname}"
        return jsonify({"ok": True, "url": url, "filename": fname})

    @app.route("/v2/api/project/<pid>/uploads/<filename>")
    def v2_api_uploads_get(pid, filename):
        proj_dir = UPLOADS_BASE / pid
        if not (proj_dir / filename).exists():
            return jsonify({"error": "not found"}), 404
        return send_from_directory(proj_dir, filename)

    @app.route("/v2/api/project/<pid>/preview.html")
    def v2_api_preview_html(pid):
        proj = db.get_project(pid)
        if not proj:
            return Response("not found", status=404)
        try:
            html = renderer.render_html(proj["specs"], base_url="")
            return Response(html, mimetype="text/html")
        except Exception as e:
            import traceback
            return Response(f"<pre>{traceback.format_exc()}</pre>",
                            status=500, mimetype="text/html")

    @app.route("/v2/from-job/<job_id>")
    def v2_from_job(job_id):
        """Übernimmt expose_data aus einem fertigen V1-Job in ein V2-Projekt.
        Voraussetzung: V1-Job war im 'preview'-Status und hat job-state.json
        geschrieben (also expose_data + customer_images_files).
        """
        state_path = f"/tmp/interpres_jobs/work_{job_id}/state.json"
        if not os.path.exists(state_path):
            return Response(
                f"V1-Job-State nicht gefunden: {job_id}\n"
                f"Du musst erst /generate-expose nutzen und auf den preview-Status warten.",
                status=404, mimetype="text/plain"
            )
        try:
            with open(state_path) as fh:
                state = json.load(fh)
            expose_data = state.get("expose_data") or {}
            specs = pipeline.build_slide_specs(expose_data)
            _resolve_anchors(specs)
            pid = db.create_project(
                specs,
                expose=expose_data,
                name=expose_data.get("projekt_titel", expose_data.get("projekt_name", "Expose"))
            )
            # Customer-Images aus V1 in V2-Uploads-Verzeichnis kopieren falls vorhanden
            cust_files = state.get("customer_images_files", {}) or {}
            v1_uploads_dir = f"/tmp/interpres_jobs/work_{job_id}/uploads"
            if os.path.isdir(v1_uploads_dir):
                for fname in os.listdir(v1_uploads_dir):
                    src = os.path.join(v1_uploads_dir, fname)
                    proj_dir = UPLOADS_BASE / pid
                    proj_dir.mkdir(exist_ok=True)
                    import shutil
                    shutil.copy(src, proj_dir / fname)
            return redirect(f"/v2/editor/{pid}")
        except Exception as e:
            import traceback
            return Response(f"<pre>{traceback.format_exc()}</pre>",
                            status=500, mimetype="text/html")

    @app.route("/v2/api/project/<pid>/render.pdf")
    def v2_api_render_pdf(pid):
        proj = db.get_project(pid)
        if not proj:
            return jsonify({"error": "not found"}), 404
        try:
            pdf = renderer.render_to_pdf(proj["specs"])
            return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                             as_attachment=True,
                             download_name=f"{proj['name'] or 'Expose'}.pdf")
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500
