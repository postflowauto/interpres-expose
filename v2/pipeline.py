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
    """Safe-get mit Default für leere Strings/None."""
    v = d.get(key, default)
    if v is None or (isinstance(v, str) and not v.strip()):
        return default
    return v


def _facts_from_expose(expose: dict) -> list[dict]:
    """Baut die 6 Investmentfacts aus den expose-Feldern."""
    return [
        {"label": "Kaufpreis ab",  "value": _g(expose, "kaufpreis_ab", "—"), "unit": "€",
         "desc": "Einstiegspreis je Apartment"},
        {"label": "KfW-Standard",  "value": _g(expose, "kfw_standard", "—"),
         "desc": "Förderfähige Effizienz"},
        {"label": "KfW-Darlehen",  "value": _g(expose, "kfw_darlehen", "150.000"), "unit": "€",
         "desc": "Zinsgünstig pro WE"},
        {"label": "Energie",       "value": _g(expose, "energieversorgung", "—"),
         "desc": "Wärmeversorgung"},
        {"label": "Stellplätze",   "value": _g(expose, "stellplaetze", "—"),
         "desc": "Außen + Tiefgarage"},
        {"label": "AfA",           "value": "3-fach",
         "desc": _g(expose, "steuerliche_moeglichkeiten",
                    "Sonder-AfA + degressive AfA + Möbel-AfA")[:80]},
    ]


def _sechs_gruende(expose: dict) -> list[dict]:
    """Sechs gute Gründe – statisch + dynamisch ergänzt."""
    return [
        {"title": "3-fach Abschreibung",      "desc": "Sonder-AfA, degressive AfA und Möbel-AfA kombinierbar."},
        {"title": "Möblierungskonzept",       "desc": "Komplett ausgestattet, sofort vermietbar."},
        {"title": "KfW-Förderung",            "desc": _g(expose, "kfw_darlehen", "150.000") + " € zinsgünstig pro WE."},
        {"title": "Mietgarantie",             "desc": "Mietausfallsicherung für die ersten Monate."},
        {"title": "Zentrale Lage",            "desc": "In " + _g(expose, "stadt", "der Stadt") + " – kurze Wege, urbanes Leben."},
        {"title": "Wertsteigerung",           "desc": "Mieten in " + _g(expose, "stadt", "der Region") + " seit 2017 " + _g(expose, "stadt_mietsteigerung", "deutlich gestiegen") + "."},
    ]


def _standort_minutes(expose: dict) -> list[dict]:
    out = []
    for key, label_key, fallback_label in [
        ("min_uni",      "label_min_uni",      "Universität"),
        ("min_bahnhof",  "label_min_bahnhof",  "Hauptbahnhof"),
        ("min_altstadt", "label_min_altstadt", "Altstadt"),
    ]:
        m = _g(expose, key, "—")
        l = _g(expose, label_key, fallback_label)
        out.append({"min": m, "label": l})
    return out


def _quartier_categories(expose: dict) -> list[dict]:
    cats = [
        ("Einkaufen",  "🛒", [(f"einkaufen_{i}_name", f"min_einkaufen_{i}") for i in range(1, 5)]),
        ("Ärzte",      "✚",  [(f"arzt_{i}_name",     f"min_arzt_{i}")     for i in range(1, 5)]),
        ("Sport",      "⚽", [(f"sport_{i}_name",    f"min_sport_{i}")    for i in range(1, 5)]),
        ("Bildung",    "📖", [(f"bildung_{i}_name",  f"min_bildung_{i}")  for i in range(1, 5)]),
    ]
    out = []
    for name, icon, items in cats:
        col_items = []
        for nk, mk in items:
            n = _g(expose, nk)
            m = _g(expose, mk)
            if n and m:
                col_items.append({"name": n, "min": m})
        if col_items:
            out.append({"name": name, "icon": icon, "items": col_items})
    return out


def _greenliving_bullets(expose: dict) -> list[dict]:
    return [
        {"title": "Einfach nachhaltiger", "desc": "Kompakte Baukörper – geringe Wärmeverluste."},
        {"title": "Einfach grüner",       "desc": "Balkone & Dachterrassen für Aufenthaltsqualität."},
        {"title": "Einfach effizienter",  "desc": "DIN-277-konforme Flächen, transparente Berechnung."},
        {"title": "Einfach stadtnäher",   "desc": "Urbane Lage – kurze Wege, weniger Autoverkehr."},
    ]


def _we_typ_specs(expose: dict) -> list[dict]:
    """Pro aktivem WE-Typ eine Folie. Aktiv = wenn we_typ_beschreibung
    oder mindestens ein we_flaeche_*-Feld einen Wert hat (oder Typ 1 immer)."""
    specs = []
    letters = "abcdefgh"

    for typ in range(1, 9):
        # Suffix für die Daten-Keys (Typ 1 hat kein Suffix, Typ 2..8 = "_typ2".._typ8")
        suf = "" if typ == 1 else f"_typ{typ}"
        # Beispiel-WEs: links/rechts (1+2 für Typ 1, 3+4 für Typ 2, ...)
        left_n  = typ * 2 - 1
        right_n = typ * 2

        beschr = _g(expose, f"we_typ_beschreibung{suf}")
        any_flaeche = any(_g(expose, f"we_flaeche_{n}{suf}") for n in range(1, 6))

        if typ > 1 and not (beschr or any_flaeche):
            continue  # nicht aktiv

        # Räume aus den 5 Slots (Wohnen/Kochen, Schlafen, Bad, Abstellraum, Balkon)
        raum_namen = ["Wohnen/Kochen", "Schlafen", "Bad", "Abstellraum", "Balkon"]
        raeume = [{"name": rn, "size": _g(expose, f"we_flaeche_{i+1}{suf}", "—")}
                  for i, rn in enumerate(raum_namen)]

        beispiele = []
        for n in (left_n, right_n):
            beispiele.append({
                "nummer":  _g(expose, f"we_beispiel_{n}", f"WE {n:02d}"),
                "bereich": _g(expose, f"we_bereich_{n}", ""),
                "raeume":  raeume,
                "img":     _g(expose, f"bild_we_{n}"),
            })

        specs.append({
            "type":   "we_typ",
            "anchor": f"we_typ_{letters[typ-1]}",
            "data": {
                "typ_letter":       letters[typ-1],
                "typ_beschreibung": beschr or f"Wohnungstyp {typ}",
                "wohnflaeche":      _g(expose, f"we_flaeche_5{suf}"),
                "beispiele":        beispiele,
            },
        })
    return specs


def _rechtliches_paragraphs(expose: dict) -> list[str]:
    """Lange Standard-Paragraphen für Rechtliches (gleich für alle Projekte)."""
    return [
        "Der Verkaufsprospekt ist erst dann als vollständig anzusehen, wenn dem "
        "Investor sowohl der hier vorliegende Prospektteil A als auch der individuelle "
        "Prospektteil B mit dem konkreten Kaufobjekt vorliegen.",
        "Der Prospekt orientiert sich im Aufbau an den Vorgaben der Wirtschaftsprüfer "
        "(IDW S 4) und an den Grundsätzen ordnungsgemäßer Beurteilung von Vermögens-"
        "anlagen-Prospekten.",
        "Maßgeblich für den Erwerb sind die individuellen Verträge mit dem Bauträger "
        "bzw. die unterschriebenen Verträge.",
        "Die hier dargestellten Visualisierungen sind unverbindlich. Abweichungen in "
        "Material, Ausführung und Farbgestaltung sind möglich.",
        "Wirtschaftliche Risiken (Marktentwicklung, Mietausfall, Zinsänderungen) "
        "können den Anlageerfolg beeinflussen. Eine individuelle Prüfung wird empfohlen.",
        "Steuerliche Aussagen basieren auf der zum Erstellungszeitpunkt geltenden "
        "Rechtslage. Änderungen der Steuergesetzgebung können Auswirkungen haben.",
    ]


def build_slide_specs(expose: dict) -> list[dict]:
    """Konstruiert die Folien-Reihenfolge aus dem expose-JSON.
    Reihenfolge orientiert sich am DQN-Standard."""
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
                {"label": _g(expose, "entwickler_name", "Urban Units"), "anchor": "kap_invest"},
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
                {"label": "Grundrisse",                                 "anchor": "we_typ_a"},
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
    specs.append({
        "type": "investmentfacts",
        "anchor": "investmentfacts",
        "data": {
            "projekt_titel": _g(expose, "projekt_titel"),
            "stadt":         _g(expose, "stadt"),
            "facts":         _facts_from_expose(expose),
        },
    })
    specs.append({
        "type": "sechs_gruende",
        "anchor": "sechs_gruende",
        "data": {"gruende": _sechs_gruende(expose)},
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
    specs.append({
        "type": "standort",
        "anchor": "standort",
        "data": {
            "stadtteil": _g(expose, "stadtteil"),
            "lead":      _g(expose, "text_standort_1") or _g(expose, "text_kapitel_live_2"),
            "minutes":   _standort_minutes(expose),
            "img":       _g(expose, "bild_standort_aussen") or _g(expose, "bild_quartier"),
        },
    })
    specs.append({
        "type": "stadt_wachstum",
        "anchor": "stadt_wachstum",
        "data": {
            "stadt":    _g(expose, "stadt", "Stadt"),
            "stadt_em": "wächst.",
            "lead":     _g(expose, "text_stadt_wachstum_1") or _g(expose, "text_stadt_intro"),
            "img":      _g(expose, "bild_stadt_klein") or _g(expose, "bild_stadt_gross"),
            "stats": [
                {"v": _g(expose, "stadt_einwohner", "—"),                       "l": "Einwohner"},
                {"v": _g(expose, "bundesland_bip", "—") + " €",                 "l": "BIP Bundesland"},
                {"v": _g(expose, "stadt_mietsteigerung", "—"),                  "l": "Mietsteigerung seit 2017"},
                {"v": _g(expose, "stadt_studierende", "—"),                     "l": "Studierende"},
            ],
        },
    })
    quartier_cats = _quartier_categories(expose)
    if quartier_cats:
        specs.append({
            "type": "quartier",
            "anchor": "quartier",
            "data": {
                "stadtteil":  _g(expose, "stadtteil"),
                "categories": quartier_cats,
                "lageplan":   _g(expose, "bild_lageplan"),
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
    specs.append({
        "type": "projekt",
        "anchor": "projekt",
        "data": {
            "stadt":         _g(expose, "stadt"),
            "stadtteil":     _g(expose, "stadtteil"),
            "projekt_titel": _g(expose, "projekt_titel"),
            "intro":         _g(expose, "text_intro"),
            "img":           _g(expose, "bild_projekt_aussen") or _g(expose, "bild_projekt"),
            "features": [
                {"value": _g(expose, "anzahl_we", "—"),     "label": "Wohneinheiten"},
                {"value": _g(expose, "groesse_von", "—") + "–" + _g(expose, "groesse_bis", "—") + " m²", "label": "Größe"},
                {"value": _g(expose, "kfw_standard", "—"),  "label": "Effizienz"},
            ],
        },
    })
    specs.append({
        "type": "greenliving",
        "anchor": "greenliving",
        "data": {
            "title":   "Einfach grüner.",
            "lead":    _g(expose, "text_greenliving_1") or _g(expose, "text_greenliving_2"),
            "bullets": _greenliving_bullets(expose),
            "img1":    _g(expose, "bild_greenliving_1"),
            "img2":    _g(expose, "bild_greenliving_2"),
        },
    })
    specs.append({
        "type": "ausstattung",
        "anchor": "ausstattung",
        "data": {
            "lead":   _g(expose, "text_ausstattung_kurz") or "Hochwertige Materialien, durchdachte Möbel, smart vernetzt.",
            "detail": _g(expose, "text_ausstattung_detail") or _g(expose, "text_ausstattung_lang"),
            "images": [_g(expose, f"bild_ausstattung_{i}") for i in range(1, 7)],
        },
    })

    # WE-Typ-Folien (eine pro aktivem Typ)
    for s in _we_typ_specs(expose):
        specs.append(s)

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
    specs.append({
        "type": "rechtliches",
        "anchor": "rechtliches",
        "data": {
            "paragraphs": _rechtliches_paragraphs(expose),
            "img":        _g(expose, "bild_rechtlich_2") or _g(expose, "bild_rechtlich_1"),
        },
    })
    specs.append({
        "type": "schluss",
        "anchor": "schluss",
        "data": {
            "stadt":           _g(expose, "stadt"),
            "intro":           "Über drei Jahrzehnte Erfahrung in der Projektentwicklung hochwertiger Immobilien.",
            "entwickler_name": _g(expose, "entwickler_name"),
            "adresse":         _g(expose, "adresse_lang"),
            "web":             "www.urbanunits.de",
        },
    })

    return specs


def sample_expose() -> dict:
    """Vollständige Demo-Daten für /v2/render-sample. Damit kann der V2-Look
    ohne ZIP-Upload getestet werden."""
    return {
        "stadt":              "Magdeburg",
        "stadtteil":          "Rothensee",
        "adresse_lang":       "Hegelstraße 4, 39104 Magdeburg",
        "projekt_titel":      "The Central",
        "entwickler_name":    "Urban Units",
        "anzahl_we":          "32",
        "groesse_von":        "22",
        "groesse_bis":        "65",
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

        "text_intro":         "Im »The Central« entstehen 32 möblierte Apartments in Magdeburg-Rothensee – modern, smart vernetzt, förderfähig nach KfW 40 QNG.",
        "text_investment_pitch":
            "Kleine Einstiegspreise ab 189.000 €, attraktive KfW-Förderung mit 150.000 € Darlehen pro WE und dreifach-AfA bieten ideale Voraussetzungen für Kapitalanleger.",
        "text_kapitel_invest_1":
            "Kleine Einstiegspreise, attraktive KfW-Förderung und dreifach-AfA bieten ideale Voraussetzungen für Kapitalanleger, die Wert auf Effizienz und Stabilität legen.",
        "text_kapitel_live_1":
            "Ein Ort, an dem man das Leben in der Stadt in vollen Zügen genießen kann – ohne auf die Schönheit der Natur zu verzichten.",
        "text_kapitel_stay_1":
            "The Central denkt Wohnen neu: kompakt, hochwertig, mit Außenbereichen – für jeden Lebensabschnitt.",
        "text_kapitel_know_1":
            "Verbindlich sind ausschließlich der notarielle Kaufvertrag und der vollständige Verkaufsprospekt. Diese Übersicht dient der Erstinformation.",
        "text_stadt_intro":
            "Die Landeshauptstadt wächst. Der Stadtteil Rothensee ist heute einer der spannendsten Orte Magdeburgs: gewachsen, urban, im Wandel.",
        "text_stadt_wachstum_1":
            "Magdeburg wächst kontinuierlich. Intel investiert 17 Mrd. €, FMC weitere 3 Mrd. € in Halbleiter-Produktion. Die Mieten stiegen seit 2017 um über 30 %.",
        "text_standort_1":
            "Rothensee liegt im Norden Magdeburgs – nah an der Elbe, an der Universität und am Hauptbahnhof. Eine Lage, die urbanes Leben mit ruhigem Wohnen verbindet.",
        "text_greenliving_1":
            "The Central setzt auf eine energieeffiziente Gebäudehülle mit modernem Wärmekonzept. Kompakte Baukörper reduzieren Wärmeverluste, niedrige Nebenkosten stärken die Vermietbarkeit – langfristig zukunftssicher.",
        "text_ausstattung_kurz":
            "Hochwertige Materialien, durchdachte Möbel, Smart-Home ready.",
        "text_ausstattung_detail":
            "Die Apartments kommen vollmöbliert mit Designermöbeln, Eichenoptik-Boden, hochwertigen Fliesen, ausgestatteter Küche, Smart-Lock und Glasfaser. Sofort vermietbar.",

        "min_uni":         "5", "label_min_uni":      "Otto-von-Guericke-Uni",
        "min_bahnhof":     "8", "label_min_bahnhof":  "Hauptbahnhof",
        "min_altstadt":    "9", "label_min_altstadt": "Magdeburger Altstadt",

        "einkaufen_1_name": "Bäckerei",     "min_einkaufen_1": "2",
        "einkaufen_2_name": "Supermarkt",   "min_einkaufen_2": "1",
        "einkaufen_3_name": "Drogerie",     "min_einkaufen_3": "3",
        "einkaufen_4_name": "Allee-Center", "min_einkaufen_4": "8",

        "arzt_1_name": "Hausarzt",     "min_arzt_1": "5",
        "arzt_2_name": "Apotheke",     "min_arzt_2": "2",
        "arzt_3_name": "Zahnarzt",     "min_arzt_3": "6",
        "arzt_4_name": "Krankenhaus",  "min_arzt_4": "10",

        "sport_1_name": "Fitnessstudio",  "min_sport_1": "5",
        "sport_2_name": "Schwimmbad",     "min_sport_2": "8",
        "sport_3_name": "Elbufer-Lauf",   "min_sport_3": "3",
        "sport_4_name": "Sportverein",    "min_sport_4": "7",

        "bildung_1_name": "Kita",         "min_bildung_1": "4",
        "bildung_2_name": "Grundschule",  "min_bildung_2": "6",
        "bildung_3_name": "Gymnasium",    "min_bildung_3": "9",
        "bildung_4_name": "Universität",  "min_bildung_4": "5",

        # WE-Typen Demo
        "we_typ_beschreibung":      "1-Zimmer-Wohnung mit Balkon",
        "we_flaeche_1":             "14,20 m²",
        "we_flaeche_2":             "—",
        "we_flaeche_3":             "4,80 m²",
        "we_flaeche_4":             "2,10 m²",
        "we_flaeche_5":             "22,36 m²",
        "we_beispiel_1":            "WE 01",  "we_bereich_1": "EG · Studio",
        "we_beispiel_2":            "WE 09",  "we_bereich_2": "OG · Studio",

        "we_typ_beschreibung_typ2": "2-Zimmer-Wohnung mit Balkon",
        "we_flaeche_1_typ2":        "18,50 m²",
        "we_flaeche_2_typ2":        "11,40 m²",
        "we_flaeche_3_typ2":        "5,20 m²",
        "we_flaeche_4_typ2":        "1,80 m²",
        "we_flaeche_5_typ2":        "37,00 m²",
        "we_beispiel_3":            "WE 03",  "we_bereich_3": "1.OG · 2-Zi",
        "we_beispiel_4":            "WE 11",  "we_bereich_4": "2.OG · 2-Zi",

        # Bilder
        "bild_titel":          "https://images.unsplash.com/photo-1545324418-cc1a3fa10c00?w=1600",
        "bild_projekt":        "https://images.unsplash.com/photo-1545324418-cc1a3fa10c00?w=1600",
        "bild_projekt_aussen": "https://images.unsplash.com/photo-1486325212027-8081e485255e?w=1600",
        "bild_stadt_klein":    "https://images.unsplash.com/photo-1502920917128-1aa500764cbd?w=1600",
        "bild_stadt_gross":    "https://images.unsplash.com/photo-1499678329028-101435549a4e?w=1600",
        "bild_standort_aussen":"https://images.unsplash.com/photo-1480714378408-67cf0d13bc1f?w=1600",
        "bild_greenliving_1":  "https://images.unsplash.com/photo-1518780664697-55e3ad937233?w=1600",
        "bild_greenliving_2":  "https://images.unsplash.com/photo-1542621334-a254cf47733d?w=1600",
        "bild_ausstattung_1":  "https://images.unsplash.com/photo-1556909114-f6e7ad7d3136?w=800",
        "bild_ausstattung_2":  "https://images.unsplash.com/photo-1556909114-44e3e9399a2e?w=800",
        "bild_ausstattung_3":  "https://images.unsplash.com/photo-1502005229762-cf1b2da7c5d6?w=800",
        "bild_ausstattung_4":  "https://images.unsplash.com/photo-1556228720-195a672e8a03?w=800",
        "bild_ausstattung_5":  "https://images.unsplash.com/photo-1554995207-c18c203602cb?w=800",
        "bild_ausstattung_6":  "https://images.unsplash.com/photo-1493809842364-78817add7ffb?w=800",
        "bild_we_1":           "https://images.unsplash.com/photo-1502672260266-1c1ef2d93688?w=1200",
        "bild_we_2":           "https://images.unsplash.com/photo-1493809842364-78817add7ffb?w=1200",
        "bild_we_3":           "https://images.unsplash.com/photo-1554995207-c18c203602cb?w=1200",
        "bild_we_4":           "https://images.unsplash.com/photo-1493809842364-78817add7ffb?w=1200",
        "bild_rechtlich_1":    "https://images.unsplash.com/photo-1516455590571-18256e5bb9ff?w=1600",
        "bild_rechtlich_2":    "https://images.unsplash.com/photo-1486325212027-8081e485255e?w=1600",
    }
