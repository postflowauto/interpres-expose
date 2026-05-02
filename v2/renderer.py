"""V2 HTML → PDF Renderer.

Pipeline:
  1. project_data (dict) + slide_specs (Liste von dicts mit type+payload)
     → render_html() baut komplettes HTML-Dokument via Jinja2
  2. render_pdf() schickt HTML an Playwright (Chromium headless), liefert PDF-Bytes
  3. post_process_toc() ersetzt TOC-Platzhalter (...) mit echten Seitenzahlen
     (Heuristik: Anker-Position im PDF → Seite)

Bewusst SYNCHRON gehalten – Playwright-async ist fragiler im Flask-Worker.
Pro Render ein eigener Browser-Context (saubere Isolation).
"""
from __future__ import annotations
import os
import io
import re
from pathlib import Path
from typing import Iterable

from jinja2 import Environment, FileSystemLoader, select_autoescape

V2_DIR     = Path(__file__).parent
SLIDES_DIR = V2_DIR / "slides"
STATIC_DIR = V2_DIR / "static"

_env = Environment(
    loader=FileSystemLoader(SLIDES_DIR),
    autoescape=select_autoescape(default_for_string=True),
    trim_blocks=True,
    lstrip_blocks=True,
)


def render_slide(slide_type: str, payload: dict) -> str:
    """Rendert eine einzelne Folie als HTML-String."""
    tmpl = _env.get_template(f"{slide_type}.html.j2")
    return tmpl.render(**payload)


def render_html(slide_specs: Iterable[dict], base_url: str = "") -> str:
    """Baut das komplette HTML-Dokument aus mehreren Folien-Specs.
    slide_specs: Liste von {"type": "cover", "data": {...}}.
    base_url: für CSS/Asset-Links (z.B. "" für relative oder "file://..." lokal).
    """
    slides_html = []
    for spec in slide_specs:
        try:
            slides_html.append(render_slide(spec["type"], spec.get("data", {})))
        except Exception as e:
            slides_html.append(
                f'<section class="slide" style="padding:20mm; color:#ff8;">'
                f'<h2>Render-Fehler: {spec.get("type", "?")}</h2>'
                f'<pre style="color:#ddd; font-size:8pt;">{e}</pre>'
                f'</section>'
            )
    layout = _env.get_template("_layout.html.j2")
    # Beispieldaten für Layout-Header
    first_data = (slide_specs[0].get("data", {}) if slide_specs else {})
    return layout.render(
        slides=slides_html,
        base_url=base_url,
        projekt_titel=first_data.get("projekt_titel", "Exposé"),
    )


def render_pdf(html: str, *, asset_root: str | None = None) -> bytes:
    """HTML → PDF via Playwright (Chromium headless).
    asset_root: lokales Verzeichnis das via file:// als Basis dient,
                damit relative <link href="..."> CSS finden.
    """
    from playwright.sync_api import sync_playwright

    pw = None
    browser = None
    try:
        pw = sync_playwright().start()
        browser = pw.chromium.launch(args=["--no-sandbox", "--disable-dev-shm-usage"])
        ctx = browser.new_context(viewport={"width": 1123, "height": 794})
        page = ctx.new_page()
        # set_content + wait_for_load_state stellt sicher, dass externe Fonts/Bilder geladen sind
        if asset_root:
            base = Path(asset_root).absolute().as_uri() + "/"
            page.set_content(html.replace('href="/v2/static/', f'href="{base}'),
                             wait_until="networkidle", timeout=30000)
        else:
            page.set_content(html, wait_until="networkidle", timeout=30000)
        # Print-emulation aktivieren – sonst gilt @media screen
        page.emulate_media(media="print")
        pdf_bytes = page.pdf(
            width="297mm",
            height="210mm",
            print_background=True,
            margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
            prefer_css_page_size=True,
        )
        ctx.close()
        return pdf_bytes
    finally:
        if browser is not None:
            try: browser.close()
            except Exception: pass
        if pw is not None:
            try: pw.stop()
            except Exception: pass


def post_process_toc(pdf_bytes: bytes, slide_specs: list[dict]) -> bytes:
    """Setzt TOC-Seitenzahlen ein basierend auf der Reihenfolge der Folien.
    Annahme: 1 Folie = 1 PDF-Seite (Playwright rendert so).

    slide_specs hat Form [{"type": "...", "anchor": "investment", ...}, ...].
    TOC-Platzhalter im HTML enthält data-toc-anchor="<key>" → wir mappen
    anchor → page_idx + 1 und ersetzen die "..." im PDF-Stream.

    Für robustes In-Place-Editing wäre PyMuPDF nötig, das Aufwand und Edge-Cases
    bringt. Stattdessen: TOC-Seitenzahlen werden direkt beim HTML-Build gesetzt
    (siehe build_slide_specs in v2.pipeline). Diese Funktion ist Stub für später.
    """
    return pdf_bytes


def render_to_pdf(slide_specs: list[dict]) -> bytes:
    """One-shot: specs → HTML → PDF. Mit pre-resolved TOC-Pages."""
    # TOC-Anker → Seitenzahl (1-basiert) auflösen, BEVOR HTML gebaut wird
    anchor_to_page = {}
    for idx, spec in enumerate(slide_specs):
        anchor = spec.get("anchor")
        if anchor:
            anchor_to_page[anchor] = idx + 1

    # TOC-Specs mit echten Seitenzahlen patchen
    for spec in slide_specs:
        if spec.get("type") != "toc":
            continue
        data = spec.setdefault("data", {})
        for col in data.get("chapters", []):
            for it in col.get("items", []):
                pg = anchor_to_page.get(it.get("anchor"))
                if pg:
                    it["page"] = pg

    # Template Cover/Items mit page rendern – TOC-Template muss item.page nutzen
    # (statt "...") wenn vorhanden. Fallback "—".
    html = render_html(slide_specs, base_url="")
    return render_pdf(html, asset_root=str(V2_DIR.parent))
