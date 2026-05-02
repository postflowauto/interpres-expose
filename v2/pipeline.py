"""V2 Pipeline – baut slide_specs aus expose_data.

expose_data ist das vorhandene JSON-Format der V1-Pipeline (Claude-Output +
Geocoding + Bilder). Hier wird daraus die Liste der Folien-Specs gebaut, die
der Renderer in HTML+PDF umwandelt.

Designprinzip: Jede Folie ist OPTIONAL und wird nur angelegt wenn die nötigen
Daten vorhanden sind (z.B. WE-Typ-Folien nur für aktive Typen). Keine starre
Slide-Anzahl mehr → kein Page-Resync-Hack mehr nötig.
"""
from __future__ import annotations
from typing import Any


def _g(d: dict, key: str, default: Any = "") -> Any:
    """Safe-get mit Default für leere Strings."""
    v = d.get(key, default)
    if v is None or (isinstance(v, str) and not v.strip()):
        return default
    return v


def build_slide_specs(expose: dict) -> list[dict]:
    """Konstruiert die Folien-Reihenfolge aus dem expose-JSON.
    Reihenfolge orientiert sich am DQN-Standard:
      Cover → TOC → Intro → Kapitel-Trenner Investment → Investmentfacts →
      6 gute Gründe → Kapitel Live → Standort → Stadt-Wachstum → Stadt-Stats →
      Stadt-Branche → Quartier-Proximity → Kapitel Stay → Projekt → Hotel →
      Greenliving → Ausstattung → Grundriss-Intro → Grundrisse → Ansichten →
      WE-Typ-Folien → Kapitel Know → Rechtliches → Schluss
    """
    specs: list[dict] = []

    # ── Cover ─────────────────────────────────────────────────────────────
    specs.append({
        "type": "cover",
        "data": {
            "brand":              "INTERPRÉS",
            "stadt":              _g(expose, "stadt"),
            "anzahl_we":          _g(expose, "anzahl_we"),
            "produkt_beschreibung": _g(expose, "produkt_beschreibung", "Apartments"),
            "projekt_titel":      _g(expose, "projekt_titel", "Exposé"),
            "zitat_intro":        _g(expose, "zitat_intro"),
            "prospekt_datum":     _g(expose, "prospekt_datum", ""),
            "bild_titel":         _g(expose, "bild_titel"),
        },
    })

    # ── TOC ───────────────────────────────────────────────────────────────
    chapters = [
        {
            "icon": "u", "tag": "invest", "title": "Das Investment",
            "items": [
                {"label": _g(expose, "entwickler_name", "Urban Units"), "anchor": "intro"},
                {"label": "Investmentfacts",                            "anchor": "investmentfacts"},
                {"label": "6 gute Gründe",                              "anchor": "sechs_gruende"},
            ],
        },
        {
            "icon": "u", "tag": "live", "title": "Der Standort",
            "items": [
                {"label": "Standort",                                   "anchor": "standort"},
                {"label": _g(expose, "stadt", "Stadt"),                 "anchor": "stadt_wachstum"},
                {"label": "Stadtteil " + _g(expose, "stadtteil", ""),   "anchor": "quartier"},
            ],
        },
        {
            "icon": "u", "tag": "stay", "title": "Das Projekt",
            "items": [
                {"label": "Das Projekt",                                "anchor": "projekt"},
                {"label": "Greenliving",                                "anchor": "greenliving"},
                {"label": "Ausstattung",                                "anchor": "ausstattung"},
                {"label": "Grundrisse",                                 "anchor": "grundrisse"},
            ],
        },
        {
            "icon": "u", "tag": "know", "title": "Das Rechtliche",
            "items": [
                {"label": "Rechtliche Hinweise",                        "anchor": "rechtliches"},
                {"label": "Kontakt",                                    "anchor": "schluss"},
            ],
        },
    ]
    specs.append({"type": "toc", "data": {"chapters": chapters}, "anchor": "toc"})

    # ── Kapitel 01: Investment ────────────────────────────────────────────
    specs.append({
        "type": "kapitel",
        "anchor": "kap_invest",
        "data": {
            "num": 1, "tag": "invest", "title": "Das Investment",
            "lead": _g(expose, "text_kapitel_invest_1") or _g(expose, "text_investment_pitch"),
            "bild": _g(expose, "bild_projekt") or _g(expose, "bild_titel"),
        },
    })

    # ── Investmentfacts ───────────────────────────────────────────────────
    facts = [
        {"label": "Kaufpreis ab",  "value": _g(expose, "kaufpreis_ab", "—"), "unit": "€",
         "desc": "Einstiegspreis je Apartment"},
        {"label": "KfW-Standard",   "value": _g(expose, "kfw_standard", "—"),
         "desc": "Förderfähige Effizienz"},
        {"label": "KfW-Darlehen",   "value": _g(expose, "kfw_darlehen", "150.000"), "unit": "€",
         "desc": "Zinsgünstig pro WE"},
        {"label": "Energie",        "value": _g(expose, "energieversorgung", "—"),
         "desc": "Wärmeversorgung"},
        {"label": "Stellplätze",    "value": _g(expose, "stellplaetze", "—"),
         "desc": "Außen + Tiefgarage"},
        {"label": "AfA",            "value": "3-fach",
         "desc": _g(expose, "steuerliche_moeglichkeiten",
                    "Sonder-AfA + degressive AfA + Möbel-AfA")[:80]},
    ]
    specs.append({
        "type": "investmentfacts",
        "anchor": "investmentfacts",
        "data": {
            "projekt_titel": _g(expose, "projekt_titel"),
            "stadt":         _g(expose, "stadt"),
            "facts":         facts,
        },
    })

    # ── Kapitel 02: Live ──────────────────────────────────────────────────
    specs.append({
        "type": "kapitel",
        "anchor": "kap_live",
        "data": {
            "num": 2, "tag": "live", "title": "Der Standort",
            "lead": _g(expose, "text_kapitel_live_1") or _g(expose, "text_stadt_intro"),
            "bild": _g(expose, "bild_stadt_gross") or _g(expose, "bild_stadt_klein"),
        },
    })

    # ── Stadt-Wachstum ────────────────────────────────────────────────────
    specs.append({
        "type": "stadt_wachstum",
        "anchor": "stadt_wachstum",
        "data": {
            "stadt": _g(expose, "stadt", "Stadt"),
            "stadt_em": "wächst.",
            "lead": _g(expose, "text_stadt_wachstum_1") or _g(expose, "text_stadt_intro"),
            "img": _g(expose, "bild_stadt_klein") or _g(expose, "bild_stadt_gross"),
            "stats": [
                {"v": _g(expose, "stadt_einwohner", "—"),     "l": "Einwohner"},
                {"v": _g(expose, "bundesland_bip", "—") + " €", "l": "BIP Bundesland"},
                {"v": _g(expose, "stadt_mietsteigerung", "—"), "l": "Mietsteigerung seit 2017"},
                {"v": _g(expose, "stadt_studierende", "—"),    "l": "Studierende"},
            ],
        },
    })

    # ── Kapitel 03: Stay ──────────────────────────────────────────────────
    specs.append({
        "type": "kapitel",
        "anchor": "kap_stay",
        "data": {
            "num": 3, "tag": "stay", "title": "Das Projekt",
            "lead": _g(expose, "text_kapitel_stay_1") or _g(expose, "text_intro"),
            "bild": _g(expose, "bild_projekt_aussen") or _g(expose, "bild_projekt"),
        },
    })

    # ── Kapitel 04: Know ──────────────────────────────────────────────────
    specs.append({
        "type": "kapitel",
        "anchor": "kap_know",
        "data": {
            "num": 4, "tag": "know", "title": "Das Rechtliche",
            "lead": _g(expose, "text_kapitel_know_1",
                       "Diese Übersicht dient ausschließlich Informationszwecken. "
                       "Verbindlich sind ausschließlich der notarielle Kaufvertrag "
                       "und der vollständige Verkaufsprospekt."),
            "bild": _g(expose, "bild_rechtlich_1"),
        },
    })

    return specs


def sample_expose() -> dict:
    """Demo-Daten für /v2/render-sample. Damit kann der V2-Look ohne ZIP-
    Upload getestet werden."""
    return {
        "stadt":              "Magdeburg",
        "stadtteil":          "Rothensee",
        "projekt_titel":      "The Central",
        "entwickler_name":    "Urban Units",
        "anzahl_we":          "32",
        "produkt_beschreibung": "1-3 Zi. Apartments",
        "kaufpreis_ab":       "189.000",
        "kfw_standard":       "KfW 40 QNG",
        "kfw_darlehen":       "150.000",
        "energieversorgung":  "Fernwärme + PV",
        "stellplaetze":       "24",
        "stadt_einwohner":    "245.279",
        "bundesland_bip":     "78,4 Mrd.",
        "stadt_mietsteigerung": "+31%",
        "stadt_studierende":  "21.000",
        "zitat_intro":        "»Magdeburg ist Deutschlands neues Technologiezentrum.«",
        "prospekt_datum":     "Stand 2025",
        "text_intro":         "Im »The Central« entstehen 32 möblierte Apartments "
                              "in Magdeburg-Rothensee – modern, smart vernetzt, "
                              "förderfähig nach KfW 40 QNG.",
        "text_investment_pitch":
            "Kleine Einstiegspreise ab 189.000 €, attraktive KfW-Förderung mit "
            "150.000 € Darlehen pro WE und dreifach-AfA bieten ideale "
            "Voraussetzungen für Kapitalanleger.",
        "text_kapitel_invest_1":
            "Kleine Einstiegspreise, attraktive KfW-Förderung und dreifach-AfA "
            "bieten ideale Voraussetzungen für Kapitalanleger, die Wert auf "
            "Effizienz und Stabilität legen.",
        "text_kapitel_live_1":
            "Ein Ort, an dem man das Leben in der Stadt in vollen Zügen genießen "
            "kann – ohne auf die Schönheit der Natur zu verzichten.",
        "text_kapitel_stay_1":
            "The Central denkt Wohnen neu: kompakt, hochwertig, mit Außen-"
            "bereichen – für jeden Lebensabschnitt.",
        "text_stadt_intro":
            "Die Landeshauptstadt wächst. Der Stadtteil Rothensee ist heute einer "
            "der spannendsten Orte Magdeburgs: gewachsen, urban, im Wandel.",
        "text_stadt_wachstum_1":
            "Magdeburg wächst kontinuierlich. Intel investiert 17 Mrd. €, FMC "
            "weitere 3 Mrd. € in Halbleiter-Produktion. Die Mieten stiegen seit "
            "2017 um über 30 %.",
        "bild_titel":         "https://images.unsplash.com/photo-1545324418-cc1a3fa10c00?w=1600",
        "bild_projekt":       "https://images.unsplash.com/photo-1545324418-cc1a3fa10c00?w=1600",
        "bild_projekt_aussen": "https://images.unsplash.com/photo-1486325212027-8081e485255e?w=1600",
        "bild_stadt_klein":   "https://images.unsplash.com/photo-1502920917128-1aa500764cbd?w=1600",
        "bild_stadt_gross":   "https://images.unsplash.com/photo-1499678329028-101435549a4e?w=1600",
        "bild_rechtlich_1":   "https://images.unsplash.com/photo-1516455590571-18256e5bb9ff?w=1600",
    }
