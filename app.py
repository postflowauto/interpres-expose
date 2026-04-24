import os, io, json, base64, zipfile, requests, re, uuid, shutil
from copy import deepcopy
from flask import Flask, request, jsonify, send_file
from pptx import Presentation
from lxml import etree

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB pro Chunk (Render Proxy Limit)

# Chunk-Upload Verzeichnis
CHUNK_DIR = "/tmp/interpres_chunks"
os.makedirs(CHUNK_DIR, exist_ok=True)

from flask import make_response

@app.errorhandler(413)
def request_too_large(e):
    return jsonify({"error": "Datei zu groß (max. 500 MB)"}), 413

@app.after_request
def add_cors(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, X-API-Token"
    response.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    return response

@app.route("/generate-expose", methods=["OPTIONS"])
@app.route("/fill-pptx", methods=["OPTIONS"])
@app.route("/debug-images", methods=["OPTIONS"])
@app.route("/health", methods=["OPTIONS"])
def options():
    return make_response("", 204)

API_TOKEN = os.environ.get("API_TOKEN", "interpres-secret-2026")
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY", "")
CLOUDCONVERT_KEY = os.environ.get("CLOUDCONVERT_KEY", "")
UNSPLASH_ACCESS_KEY = os.environ.get("UNSPLASH_ACCESS_KEY", "")
TEST_MODE = os.environ.get("TEST_MODE", "false").lower() == "true"
TEMPLATE_URL = "https://raw.githubusercontent.com/postflowauto/interpres-expose/main/urbanunits_Marketing_Expose_v3.pdf-15.pptx"

# Dummy-Daten für TEST_MODE (kein Claude-API-Call)
DUMMY_PROJEKTDATEN = {
    "projektname_roh": "Testprojekt Hannover", "adresse": "Lindener Marktplatz 5",
    "stadt": "Hannover", "stadtteil": "Linden", "plz": "30449",
    "bautraeger": "Urban Units GmbH", "anzahl_haeuser": "2",
    "we_pro_haus": "24", "anzahl_we_gesamt": "48", "kfw_standard": "KfW 55 EE",
    "energieversorgung": "Fernwärme + Photovoltaik", "stellplaetze": "24",
    "groesse_von": "28", "groesse_bis": "67", "kaufpreis_ab": "189.000",
    "besonderheiten": "Möbliert, Smart-Lock, Dachterrasse", "planungsphase": "Baugenehmigung erteilt"
}

DUMMY_EXPOSE_DATA = {
    "projekt_name": "Stadtquartier Linden", "projekt_titel": "Leben im Herzen Lindens",
    "entwickler_name": "Urban Units GmbH", "entwickler_name_gross": "URBAN UNITS GMBH",
    "stadt": "Hannover", "stadt_gross": "HANNOVER", "stadtteil": "Linden",
    "adresse_lang": "Lindener Marktplatz 5, 30449 Hannover", "plz": "30449",
    "quartier_name": "Linden-Mitte", "quartier_history": "Lebendiges Gründerzeitviertel mit urbanem Flair",
    "quartier_ref": "Hannover Linden", "stadt_bezeichnung": "Landeshauptstadt",
    "anzahl_we": "48", "anzahl_1zi": "12", "anzahl_2zi": "24", "anzahl_barrierefrei": "12",
    "groesse_von": "28", "groesse_bis": "67", "zimmer_typen": "1-Zimmer und 2-Zimmer",
    "produkt_beschreibung": "Vollmöblierte Mikro-Apartments mit Smart-Lock",
    "kaufpreis_ab": "189.000", "kfw_darlehen": "150.000", "stellplaetze": "24",
    "kfw_standard": "KfW 55 EE", "energieversorgung": "Fernwärme + Photovoltaik",
    "besonderheiten": "Möbliert, Smart-Lock, Dachterrasse, E-Bike-Sharing",
    "steuerliche_moeglichkeiten": "Dreifach AfA - 5% degressiv §7 Abs.5a EStG + 5% Sonder-AfA §7b EStG + 10% Möbel-AfA",
    "prospekt_datum": "April 2026",
    "text_kapitel_invest": "INVEST", "text_kapitel_live": "LIVE",
    "text_kapitel_stay": "STAY", "text_kapitel_know": "KNOW",
    "text_kapitel_hotel": "HOTEL",
    "text_intro": "Hannover wächst – und Linden ist mittendrin.",
    "text_investment_pitch": "Solide Rendite in einem der dynamischsten Stadtteile Hannovers.",
    "text_hotel": "Möbliert, flexibel, sofort vermietbar – ideal für Kurzzeitvermietung.",
    "text_projekt_nachhaltig_1": "KfW 55 EE – höchster Förderstandard.",
    "text_projekt_nachhaltig_2": "Photovoltaik deckt 30% des Allgemeinstrombedarfs.",
    "text_greenliving_intro": "Nachhaltig wohnen in Hannover.",
    "text_greenliving_1": "Fernwärme aus regenerativen Quellen.",
    "text_greenliving_2": "E-Bike-Sharing für alle Bewohner.",
    "text_ausstattung_intro": "Hochwertig. Vollständig. Bezugsfertig.",
    "text_ausstattung_detail": "Designermöbel, Echtholzparkett, moderne Einbauküche.",
    "text_ausstattung_kurz": "Alles inklusive.", "text_ausstattung_lang": "Vom Bett bis zur Kaffeemaschine.",
    "text_grundriss_intro": "Clever geplante Grundrisse für maximale Nutzfläche.",
    "text_architektur": "Zeitloser Klinkerbau trifft moderne Glaselemente.",
    "text_nachhaltig_1": "KfW 55 EE", "text_nachhaltig_2": "Fernwärme",
    "text_nachhaltig_3": "Photovoltaik", "text_nachhaltig_4": "E-Mobilität",
    "text_standort_1": "Zentral in Hannover-Linden.", "text_standort_2": "Alles in Laufnähe.",
    "stadt_einwohner": "535.932", "stadt_bip": "38.500", "stadt_mietsteigerung": "+3,2%",
    "stadt_studierende": "48.000", "bundesland_bip": "310 Mrd.",
    "text_einwohner_detail": "Hannover wächst kontinuierlich.",
    "text_bip_detail": "Niedersachsen – starke Industrie und Dienstleistungen.",
    "text_mietsteigerung_detail": "Stabile Mietsteigerungen über dem Bundesschnitt.",
    "text_studierende_detail": "Universitätsstadt mit hoher Nachfrage.",
    "text_stadt_wachstum_1": "Bevölkerungswachstum seit 2015 konstant.",
    "text_stadt_wachstum_2": "Zuzug aus Ballungsräumen verstärkt Nachfrage.",
    "text_stadt_intro": "Hannover – Niedersachsens Wirtschaftsmotor.",
    "text_stadt_wirtschaft_links": "Messe, Continental, TUI – globale Player vor Ort.",
    "text_stadt_wirtschaft_rechts": "Starker Mittelstand und wachsende Startup-Szene.",
    "stadt_invest_titel": "Investitionsstandort Hannover",
    "stadt_invest_label": "Rendite", "text_stadt_invest_detail": "Attraktive Nettomietrenditen von 4–5%.",
    "stadt_stat_1_zahl": "535.932", "stadt_stat_1_label": "Einwohner",
    "stadt_stat_2_zahl": "48.000", "stadt_stat_2_label": "Studierende",
    "stadt_stat_3_zahl": "+3,2%", "stadt_stat_3_label": "Mietsteigerung p.a.",
    "stadt_branche_titel": "Leitbranchen", "text_stadt_branche_1": "Messe & Kongress",
    "text_stadt_branche_2": "Automobil & Logistik",
    "quelle_1": "Statistik Hannover 2024", "quelle_2": "IHK Hannover 2024",
    "quelle_3": "Wohnmarktreport 2024", "quelle_4": "Bundesagentur für Arbeit 2024",
    # ── Standort-Minuten (Slide 5) ────────────────────────────────────────────
    "min_uni": "18", "label_min_uni": "Leibniz Universität",
    "min_bahnhof": "12", "label_min_bahnhof": "Hauptbahnhof",
    "min_altstadt": "15", "label_min_altstadt": "Altstadt",
    # ── Alles ganz nah (Slide 14): 4 Freizeit-Einträge ───────────────────────
    "freizeit_1_name": "Maschsee", "min_freizeit_1": "8",
    "freizeit_2_name": "Eilenriede", "min_freizeit_2": "12",
    "freizeit_3_name": "Innenstadt", "min_freizeit_3": "15",
    "freizeit_4_name": "Herrenhäuser Gärten", "min_freizeit_4": "20",
    # ── WE-Typen: Original-Slide (Typen 1+2 nebeneinander) ───────────────────
    "we_beispiel_1": "WE 02", "we_bereich_1": "Wohnen & Schlafen",
    "we_beispiel_2": "WE 07", "we_bereich_2": "Wohnen & Schlafen",
    "we_flaeche_1": "23,99 m²",
    "we_flaeche_2": "5,36 m²",
    "we_flaeche_3": "5,34 m²",
    "we_flaeche_4": "2,33 m²",
    "we_flaeche_5": "32,02 m²",
    "we_typ_beschreibung": "1-Zimmer-Wohnung mit Balkon. Optimal für Studierende und Berufspendler.",
    # Duplikat-Slide (Typen 3+4), leer = kein Duplikat
    "we_beispiel_3": "", "we_bereich_3": "",
    "we_beispiel_4": "", "we_bereich_4": "",
    # Duplikat-Slide 2 (Typen 5+6), leer = kein zweites Duplikat
    "we_beispiel_5": "", "we_bereich_5": "",
    "we_beispiel_6": "", "we_bereich_6": "",
    "feature_1_zahl": "48", "feature_1_label": "Wohneinheiten",
    "feature_2_zahl": "100", "feature_2_label": "Prozent möbliert",
    "feature_3_zahl": "24", "feature_3_label": "Stunden Zugang per Smart-Lock-System",
    "amenity_1": "E-Bike-Sharing", "amenity_2": "Solar-Carport",
    "amenity_3": "Fitnessstudio", "amenity_4": "Paketstation",
    "amenity_5": "Café im EG", "amenity_6": "Dachgarten",
    "amenity_7": "Fernwärme", "amenity_8": "Tiefgarage", "amenity_9": "Balkon",
    "grundriss_1_label": "Typ A – 28 m²", "grundriss_2_label": "Typ B – 42 m²",
    "grundriss_3_label": "Typ C – 55 m²", "grundriss_4_label": "Typ D – 67 m²",
    "ansicht_1_label": "Westfassade", "ansicht_2_label": "Südfassade",
    "logo_initial": "S",
    # Kapitel-Seiten Subtext
    "text_kapitel_invest_1": "Nachhaltig investieren in Hannover.",
    "text_kapitel_invest_2": "Maximale Förderung, stabile Rendite.",
    "text_kapitel_live_1": "Die Stadt. Der Standort. Das Quartier.",
    "text_kapitel_live_2": "Hannover – Wirtschaftsmotor Niedersachsens.",
    "text_kapitel_stay_1": "Vollmöbliert. Nachhaltig. Bezugsfertig.",
    "text_kapitel_stay_2": "Design trifft Funktion in Hannover-Linden.",
    "text_kapitel_know_1": "Transparenz und Rechtssicherheit.",
    "text_kapitel_know_2": "Alle Fakten auf einen Blick.",
    # Stadtstatistik Details
    "text_stadt_stat_1_detail": "Hannover wächst kontinuierlich.",
    "text_stadt_stat_2_detail": "Universitätsstadt mit hoher Nachfrage.",
    "text_stadt_stat_3_detail": "Stabile Mietsteigerungen über dem Bundesschnitt.",
    # Grundriss-Bilder
    "bild_grundriss_1": "", "bild_grundriss_2": "", "bild_grundriss_3": "", "bild_grundriss_4": "",
    # Bundesland
    "bundesland": "Niedersachsen",
    "bild_titel": "", "bild_quartier": "",
    "bild_projekt_aussen": "", "bild_amenity_1": "", "bild_amenity_2": "",
    "bild_amenity_3": "", "bild_amenity_4": "", "bild_amenity_5": "",
    "bild_amenity_6": "", "bild_amenity_7": "", "bild_amenity_8": "",
    "bild_amenity_9": "", "bild_greenliving_1": "", "bild_greenliving_2": "",
    "bild_interior": "", "bild_ausstattung_1": "", "bild_ausstattung_2": "",
    "bild_ausstattung_3": "", "bild_ausstattung_4": "", "bild_ausstattung_5": "",
    "bild_ausstattung_6": "", "bild_grundriss_intro_1": "", "bild_grundriss_intro_2": "",
    "bild_ansicht_1": "", "bild_ansicht_2": "",
    **{f"bild_we_{n}": "" for n in range(1, 21)},   # bild_we_1 … bild_we_20 für DUMMY
    "bild_stadt_presse": "", "bild_stadt_branche": "",
    "bild_rechtlich_1": "", "bild_rechtlich_2": "",
    "bild_collage_1": "", "bild_collage_2": "", "bild_collage_3": "",
    "bild_collage_4": "", "bild_collage_5": "",
    "bild_standort_innen": "", "bild_standort_aussen": "",
    "bild_hotel_1": "", "bild_hotel_2": "",
    "bild_stadt_gross": "", "bild_stadt_klein": "",
    "bild_lageplan": "", "bild_grundriss_intro_3": "",
    "bild_projekt": "",
    "produkt_beschreibung": "Vollmöblierte Mikro-Apartments mit Smart-Lock",
    "zitat_intro": "Wohnen neu gedacht.",
}

# Relevante PDF-Typen nach Priorität
PDF_PRIORITY = [
    (1, ["zusammenfassung", "summary"]),
    (1, ["berechnung-bri", "bri-berechnung"]),
    (2, ["grundriss", "floor", "lageplan"]),
    (2, ["wfl-berechnung", "wohnflaeche", "wfl_berechnung"]),
    (3, ["schnitt", "ansicht", "elevation"]),
]

UNSPLASH_QUERIES = {
    "BILD_TITEL": "modern luxury residential building exterior",
    "BILD_QUARTIER": "city neighborhood street urban architecture",
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
    "BILD_WE_1":  "modern apartment interior design minimal furnished",
    "BILD_WE_2":  "studio apartment interior furnished modern design",
    "BILD_WE_3":  "cozy apartment living room interior design",
    "BILD_WE_4":  "modern apartment bedroom minimalist design",
    "BILD_WE_5":  "luxury apartment interior penthouse design",
    "BILD_WE_6":  "compact apartment smart interior design",
    "BILD_WE_7":  "modern apartment open plan living design",
    "BILD_WE_8":  "bright apartment interior contemporary",
    "BILD_WE_9":  "minimalist apartment interior Scandinavian",
    "BILD_WE_10": "modern apartment kitchen dining area",
    "BILD_WE_11": "studio loft apartment modern design",
    "BILD_WE_12": "apartment terrace balcony modern",
    "BILD_WE_13": "penthouse apartment interior luxury",
    "BILD_WE_14": "duplex apartment interior design",
    "BILD_WE_15": "modern apartment bathroom interior",
    "BILD_WE_16": "cozy apartment bedroom interior",
    "BILD_WE_17": "open plan apartment living dining",
    "BILD_WE_18": "modern apartment hallway entrance",
    "BILD_WE_19": "contemporary apartment interior style",
    "BILD_WE_20": "bright apartment modern interior design",
    "BILD_STADT_PRESSE": "newspaper article table coffee",
    "BILD_STADT_BRANCHE": "scientist laboratory research modern",
    "BILD_RECHTLICH_1": "modern residential building exterior",
    "BILD_RECHTLICH_2": "apartment building facade evening",
    "BILD_COLLAGE_1": "modern apartment interior living room",
    "BILD_COLLAGE_2": "food lifestyle dinner table modern",
    "BILD_COLLAGE_3": "rooftop terrace modern apartment",
    "BILD_COLLAGE_4": "modern kitchen interior design",
    "BILD_COLLAGE_5": "apartment building exterior architecture",
    "BILD_STANDORT_INNEN": "modern bedroom interior minimal",
    "BILD_STANDORT_AUSSEN": "residential building exterior street",
    "BILD_HOTEL_1": "hotel bedroom luxury modern",
    "BILD_HOTEL_2": "hotel lobby modern interior",
    "BILD_STADT_GROSS": "city skyline aerial",
    "BILD_STADT_KLEIN": "city street urban",
    "BILD_LAGEPLAN": "city map urban district aerial overview",
    "BILD_GRUNDRISS_INTRO_3": "modern apartment interior living space",
    "BILD_PROJEKT": "modern luxury apartment building exterior night",
    "BILD_GRUNDRISS_1": "apartment floor plan architectural drawing",
    "BILD_GRUNDRISS_2": "apartment floor plan 2 room layout",
    "BILD_GRUNDRISS_3": "apartment floor plan 3 room layout",
    "BILD_GRUNDRISS_4": "apartment floor plan large layout",
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
    "text_kapitel_stay": "", "text_kapitel_know": "", "text_kapitel_hotel": "", "text_intro": "",
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
    # ── Standort-Minuten (Slide 5) ────────────────────────────────────────────
    "min_uni": "", "label_min_uni": "",
    "min_bahnhof": "", "label_min_bahnhof": "",
    "min_altstadt": "", "label_min_altstadt": "",
    # ── Alles ganz nah (Slide 14): 4 Freizeit-Einträge ───────────────────────
    "freizeit_1_name": "", "min_freizeit_1": "",
    "freizeit_2_name": "", "min_freizeit_2": "",
    "freizeit_3_name": "", "min_freizeit_3": "",
    "freizeit_4_name": "", "min_freizeit_4": "",
    # ── WE-Typen ──────────────────────────────────────────────────────────────
    "we_beispiel_1": "", "we_bereich_1": "",
    "we_beispiel_2": "", "we_bereich_2": "",
    "we_flaeche_1": "", "we_flaeche_2": "", "we_flaeche_3": "",
    "we_flaeche_4": "", "we_flaeche_5": "",
    "we_typ_beschreibung": "",
    "we_beispiel_3": "", "we_bereich_3": "",
    "we_beispiel_4": "", "we_bereich_4": "",
    "we_beispiel_5": "", "we_bereich_5": "",
    "we_beispiel_6": "", "we_bereich_6": "",
    "feature_1_zahl": "", "feature_1_label": "",
    "feature_2_zahl": "100", "feature_2_label": "Prozent möbliert",
    "feature_3_zahl": "24", "feature_3_label": "Stunden Zugang per Smart-Lock-System",
    "amenity_1": "", "amenity_2": "", "amenity_3": "", "amenity_4": "", "amenity_5": "",
    "amenity_6": "", "amenity_7": "", "amenity_8": "", "amenity_9": "",
    "grundriss_1_label": "", "grundriss_2_label": "", "grundriss_3_label": "", "grundriss_4_label": "",
    "ansicht_1_label": "", "ansicht_2_label": "",
    "logo_initial": "",
    # Kapitel-Seiten Subtext (2 Zeilen pro Kapitel)
    "text_kapitel_invest_1": "", "text_kapitel_invest_2": "",
    "text_kapitel_live_1": "", "text_kapitel_live_2": "",
    "text_kapitel_stay_1": "", "text_kapitel_stay_2": "",
    "text_kapitel_know_1": "", "text_kapitel_know_2": "",
    # Stadtstatistik Details
    "text_stadt_stat_1_detail": "", "text_stadt_stat_2_detail": "", "text_stadt_stat_3_detail": "",
    # Grundriss-Bilder (direkte Slots, nicht intro)
    "bild_grundriss_1": "", "bild_grundriss_2": "", "bild_grundriss_3": "", "bild_grundriss_4": "",
    # Bundesland (ohne BIP)
    "bundesland": "",
    "bild_titel": "", "bild_quartier": "",
    "bild_projekt_aussen": "", "bild_amenity_1": "", "bild_amenity_2": "", "bild_amenity_3": "",
    "bild_amenity_4": "", "bild_amenity_5": "", "bild_amenity_6": "", "bild_amenity_7": "",
    "bild_amenity_8": "", "bild_amenity_9": "", "bild_greenliving_1": "", "bild_greenliving_2": "",
    "bild_interior": "", "bild_ausstattung_1": "", "bild_ausstattung_2": "", "bild_ausstattung_3": "",
    "bild_ausstattung_4": "", "bild_ausstattung_5": "", "bild_ausstattung_6": "",
    "bild_grundriss_intro_1": "", "bild_grundriss_intro_2": "",
    "bild_ansicht_1": "", "bild_ansicht_2": "",
    **{f"bild_we_{n}": "" for n in range(1, 21)},   # bild_we_1 … bild_we_20
    "bild_stadt_presse": "", "bild_stadt_branche": "",
    "bild_rechtlich_1": "", "bild_rechtlich_2": "",
    "bild_collage_1": "", "bild_collage_2": "", "bild_collage_3": "",
    "bild_collage_4": "", "bild_collage_5": "",
    "bild_standort_innen": "", "bild_standort_aussen": "",
    "bild_hotel_1": "", "bild_hotel_2": "",
    "bild_stadt_gross": "", "bild_stadt_klein": "",
    "bild_lageplan": "", "bild_grundriss_intro_3": "",
    "bild_projekt": "",
    "zitat_intro": "",
}

def generate_logo_initial(projekt_name):
    """Nimmt den ersten markanten Buchstaben des Projektnamens als Logo-Initial."""
    if not projekt_name:
        return "P"
    skip_words = {"das", "der", "die", "ein", "eine", "am", "im", "an", "auf", "the", "a", "an"}
    words = re.split(r'[\s\-_\.]+', projekt_name)
    for word in words:
        cleaned = re.sub(r'[^a-zA-Z\u00c4\u00d6\u00dc\u00e4\u00f6\u00fc]', '', word)
        if cleaned and cleaned.lower() not in skip_words:
            return cleaned[0].upper()
    return projekt_name[0].upper()

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

def extract_images_from_zip(zip_bytes):
    """
    Extrahiert Bilddateien (.jpg/.jpeg/.png/.webp) aus einer ZIP.
    Gibt max. 20 Bilder zurück als Liste von {name, bytes, ext, b64}.
    """
    IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.webp'}
    images = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if '__MACOSX' in name or name.startswith('.'):
                    continue
                ext = os.path.splitext(name.lower())[1]
                if ext not in IMAGE_EXTS:
                    continue
                raw = zf.read(name)
                if len(raw) < 5000:   # Thumbnails/Icons überspringen
                    continue
                images.append({
                    'name': name.split('/')[-1],
                    'bytes': raw,
                    'ext': ext,
                    'b64': base64.b64encode(raw).decode(),
                })
                if len(images) >= 20:
                    break
    except Exception as e:
        print(f"extract_images_from_zip Fehler: {e}")
    print(f"Bilder in ZIP: {len(images)}")
    return images


# Mapping: welche bild_* Slots können Kundenbilder aufnehmen (in Prioritätsreihenfolge)
CUSTOMER_IMAGE_SLOTS = {
    "aussenansicht":  ["bild_titel", "bild_projekt_aussen", "bild_ansicht_1", "bild_ansicht_2",
                       "bild_greenliving_1", "bild_greenliving_2"],
    "grundriss":      ["bild_grundriss_1", "bild_grundriss_2", "bild_grundriss_3", "bild_grundriss_4",
                       "bild_grundriss_intro_1", "bild_grundriss_intro_2", "bild_grundriss_intro_3"],
    "innenansicht":   ["bild_interior", "bild_ausstattung_1", "bild_ausstattung_2", "bild_ausstattung_3",
                       "bild_ausstattung_4", "bild_ausstattung_5", "bild_ausstattung_6",
                       "bild_we_1", "bild_we_2", "bild_hotel_1", "bild_hotel_2"],
    "lageplan":       ["bild_lageplan"],
    "quartier":       ["bild_quartier", "bild_stadt_gross", "bild_stadt_klein",
                       "bild_standort_aussen", "bild_standort_innen"],
    "amenity":        ["bild_amenity_1", "bild_amenity_2", "bild_amenity_3", "bild_amenity_4",
                       "bild_amenity_5", "bild_amenity_6", "bild_amenity_7", "bild_amenity_8", "bild_amenity_9"],
    "collage":        ["bild_collage_1", "bild_collage_2", "bild_collage_3", "bild_collage_4", "bild_collage_5"],
}


def classify_and_assign_customer_images(images):
    """
    Sendet Kundenbilder an Claude Vision und lässt es sie den richtigen bild_*-Slots zuweisen.
    Gibt {bild_key: image_bytes} zurück.
    Fällt auf regelbasierte Zuweisung zurück wenn kein Claude-Key vorhanden.
    """
    if not images:
        return {}

    # Regelbasierter Fallback (anhand Dateiname)
    def _rule_based(images):
        slot_counters = {cat: 0 for cat in CUSTOMER_IMAGE_SLOTS}
        result = {}
        for img in images:
            name_lower = img['name'].lower()
            if any(k in name_lower for k in ('grundriss', 'floor', 'plan', 'gr_')):
                cat = 'grundriss'
            elif any(k in name_lower for k in ('lageplan', 'lage', 'map', 'karte')):
                cat = 'lageplan'
            elif any(k in name_lower for k in ('aussen', 'außen', 'exterior', 'fassade', 'ansicht')):
                cat = 'aussenansicht'
            elif any(k in name_lower for k in ('innen', 'interior', 'zimmer', 'küche', 'bad', 'wohn')):
                cat = 'innenansicht'
            elif any(k in name_lower for k in ('quartier', 'strasse', 'straße', 'stadt', 'neighborhood')):
                cat = 'quartier'
            elif any(k in name_lower for k in ('amenity', 'bike', 'solar', 'gym', 'garten', 'dach')):
                cat = 'amenity'
            else:
                cat = 'aussenansicht'  # Standard-Fallback
            slots = CUSTOMER_IMAGE_SLOTS[cat]
            idx = slot_counters[cat]
            if idx < len(slots):
                result[slots[idx]] = img['bytes']
                slot_counters[cat] += 1
        return result

    if not CLAUDE_API_KEY:
        print("classify_customer_images: kein Claude-Key → regelbasierte Zuweisung")
        result = _rule_based(images)
        print(f"  {len(result)} Bilder zugewiesen (regelbasiert)")
        return result

    # Claude Vision: max. 10 Bilder (Kostenkontrolle)
    batch = images[:10]
    content = []
    for i, img in enumerate(batch):
        if img['ext'] in ('.jpg', '.jpeg'):
            media_type = 'image/jpeg'
        elif img['ext'] == '.png':
            media_type = 'image/png'
        else:
            media_type = 'image/webp'
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": media_type, "data": img['b64']}
        })
        content.append({"type": "text", "text": f"Bild {i+1}: {img['name']}"})

    slot_list = "\n".join(
        f"- {cat}: {', '.join(slots[:4])}" + (" …" if len(slots) > 4 else "")
        for cat, slots in CUSTOMER_IMAGE_SLOTS.items()
    )
    content.append({"type": "text", "text": (
        "Analysiere diese Immobilien-Bilder und weise jedem Bild die passende Kategorie zu.\n\n"
        "Kategorien und ihre Platzhalter:\n" + slot_list + "\n\n"
        "Regeln:\n"
        "- grundriss: NUR echte Grundriss-Zeichnungen mit Raumaufteilung/Maßen\n"
        "- lageplan: NUR Karten, Stadtpläne, Lagepläne\n"
        "- aussenansicht: Gebäude-Außenfotos, Fassade, Architektur\n"
        "- innenansicht: Innenräume, Wohnung, Möbel, Küche, Bad\n"
        "- quartier: Straßen, Stadtteile, Umgebung\n"
        "- amenity: Ausstattungsmerkmale (Fahrrad, Solar, Gym, Garten)\n"
        "- collage: Sonstige/dekorative Bilder\n\n"
        "Antworte NUR mit JSON: {\"1\": \"kategorie\", \"2\": \"kategorie\", ...}\n"
        "Bild-Nummern wie oben angegeben. Jedes Bild bekommt genau eine Kategorie."
    )})

    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": CLAUDE_API_KEY, "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json={"model": "claude-haiku-4-5-20251001", "max_tokens": 300,
                  "messages": [{"role": "user", "content": content}]},
            timeout=60
        )
        resp.raise_for_status()
        json_text = resp.json()["content"][0]["text"]
        json_text = json_text.replace("```json", "").replace("```", "").strip()
        assignments = json.loads(json_text)
    except Exception as e:
        print(f"classify_customer_images Claude-Fehler: {e} → regelbasierter Fallback")
        result = _rule_based(images)
        print(f"  {len(result)} Bilder zugewiesen (regelbasiert)")
        return result

    # Kategorie → nächsten freien Slot zuweisen
    slot_counters = {cat: 0 for cat in CUSTOMER_IMAGE_SLOTS}
    result = {}
    for img_num_str, cat in assignments.items():
        idx = int(img_num_str) - 1
        if not (0 <= idx < len(batch)):
            continue
        cat = cat.strip().lower()
        if cat not in CUSTOMER_IMAGE_SLOTS:
            cat = 'aussenansicht'
        slots = CUSTOMER_IMAGE_SLOTS[cat]
        slot_idx = slot_counters[cat]
        if slot_idx < len(slots):
            result[slots[slot_idx]] = batch[idx]['bytes']
            slot_counters[cat] += 1
            print(f"  Bild {img_num_str} ({batch[idx]['name']}) → {slots[slot_idx]}")

    print(f"classify_customer_images: {len(result)} Bilder zugewiesen (Claude Vision)")
    return result


def fetch_unsplash_image(query):
    """Holt Bild-URL von Unsplash. Bei fehlendem/ungültigem Key → Picsum-Fallback."""
    if UNSPLASH_ACCESS_KEY:
        try:
            resp = requests.get(
                "https://api.unsplash.com/photos/random",
                params={"query": query, "orientation": "landscape"},
                headers={"Authorization": f"Client-ID {UNSPLASH_ACCESS_KEY}"},
                timeout=10
            )
            print(f"  Unsplash [{resp.status_code}] query={query!r}")
            if resp.status_code == 200:
                url = resp.json()["urls"]["regular"]
                print(f"    → {url[:80]}")
                return url
            else:
                print(f"    Fehler-Body: {resp.text[:120]}")
        except Exception as e:
            print(f"  Unsplash Exception für '{query}': {e}")
    else:
        print(f"  UNSPLASH_ACCESS_KEY fehlt – Picsum-Fallback für '{query}'")

    # Picsum-Fallback: deterministisches Bild anhand des Query-Hashes
    seed = abs(hash(query)) % 1000
    url = f"https://picsum.photos/seed/{seed}/1200/800"
    print(f"  Picsum-Fallback → {url}")
    return url

def fill_image_placeholders(data):
    stadt = data.get("stadt", "")
    queries = UNSPLASH_QUERIES.copy()
    if stadt:
        queries["BILD_TITEL"] = f"modern residential building {stadt} urban architecture"
        queries["BILD_QUARTIER"] = f"{stadt} city neighborhood street"
        queries["BILD_PROJEKT_AUSSEN"] = f"modern apartment building {stadt} exterior"
        queries["BILD_GREENLIVING_1"] = f"sustainable green building {stadt}"
        queries["BILD_GREENLIVING_2"] = f"modern residential {stadt} facade"
        queries["BILD_STADT_GROSS"] = f"city skyline aerial {stadt}"
        queries["BILD_STADT_KLEIN"] = f"city street urban {stadt}"
        queries["BILD_LAGEPLAN"] = f"city map district aerial {stadt}"
    filled = 0
    for placeholder_key, query in queries.items():
        data_key = placeholder_key.lower()
        if data_key not in data:
            continue
        # Skip bild_we_N für nicht vorhandene WE-Typen:
        # Paar k (k≥2): left_n = 2k-1, right_n = 2k
        # Nur laden wenn TEXT-Keys des Paares befüllt sind.
        import re as _re
        _m = _re.match(r'^bild_we_(\d+)$', data_key)
        if _m:
            n = int(_m.group(1))
            if n > 2:
                pair_k    = (n + 1) // 2   # welches Paar: n=3→2, n=4→2, n=5→3, ...
                left_n    = pair_k * 2 - 1
                right_n   = pair_k * 2
                has_text  = (data.get(f"we_beispiel_{left_n}") or data.get(f"we_bereich_{left_n}")
                             or data.get(f"we_beispiel_{right_n}") or data.get(f"we_bereich_{right_n}"))
                if not has_text:
                    continue
        url = fetch_unsplash_image(query)
        if url:
            data[data_key] = url
            filled += 1
    print(f"fill_image_placeholders: {filled}/{len(queries)} Bilder befüllt")
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
        json={"model": "claude-haiku-4-5-20251001", "max_tokens": 4000,
              "messages": [{"role": "user", "content": content}]},
        timeout=120
    )
    resp.raise_for_status()
    text = resp.json()["content"][0]["text"].replace("```json", "").replace("```", "").strip()
    return json.loads(text)

def generate_expose_with_claude(projektdaten):
    stadt = projektdaten.get('stadt', 'der Stadt')
    projekt = projektdaten.get('projektname_roh', 'das Projekt')
    bautraeger = projektdaten.get('bautraeger', 'urbanunits')

    prompt = (
        "Du bist ein erfahrener Immobilien-Exposé-Texter bei INTERPRÉS GmbH. "
        "Antworte NUR mit einem validen JSON-Objekt. Kein Text davor oder danach. Keine Markdown-Backticks.\n\n"

        f"## PROJEKTDATEN\n{json.dumps(projektdaten, ensure_ascii=False)}\n\n"

        "## SCHREIBSTIL – REFERENZ (genau so schreiben!)\n"
        "Das Exposé folgt dem Stil eines Premium-Immobilien-Prospekts. Konkret:\n\n"

        "### Slogans (text_kapitel_*, text_hotel, text_*_kurz):\n"
        "Kurze englische Phrasen mit Punkt. Maximal 3-4 Wörter. Beispiele:\n"
        "'feels like a hotel.'  'think green. live smart.'  'naturban.'  'work, life balance.'\n"
        "'designed to stay.'  'stilvoll. durchdacht.'  'simply more.'\n\n"

        "### Fließtexte (text_intro, text_investment_pitch, text_greenliving_*, text_ausstattung_detail, text_hotel, text_architektur):\n"
        "LANG und inhaltsstark. Mindestens 4-6 Sätze. Emotionale Einstiegssatz + konkrete Fakten + Zielgruppe + Ausblick.\n"
        "Beispiel text_intro-Länge:\n"
        "'Inmitten der geschäftigen Kulisse der Neuen Neustadt, einem aufstrebenden Stadtteil, verbindet "
        "\"urbanunits – The Central\" auf einzigartige Weise urbanen Komfort mit dem Gefühl von Rückzug und Ruhe. "
        "Hier entfaltet sich eine grüne Oase, die auf dem südlichen Areal der einstigen Diamant Brauerei entsteht. "
        "Dieses Projekt ist mehr als nur ein Bauvorhaben – es ist ein Ort mit Charakter, geschaffen für Menschen "
        "mit einem modernen Lebensstil. \"urbanunits – The Central\" bietet nicht nur gefragten Wohnraum, "
        "sondern auch eine Verbindung von Stadt und Qualität, die eine besondere Balance schafft. "
        "Ideal für Studierende, Berufstätige, Zweitwohnsitznutzer – und jeden, der flexibel wohnen möchte.'\n\n"

        "### Key-Facts (feature_N_label, amenity_N):\n"
        "Kurze prägnante Substantive oder Kurzsätze. Beispiele:\n"
        "'Fitnessstudio direkt im Quartier'  'E-Bike-Sharing'  'Paketstation'  'Dachbegrünung'\n\n"

        "### Zahlen (feature_N_zahl, min_*, stadtstatistiken):\n"
        "Nur die Zahl, kein Text. Fahrrad-/Gehminuten realistisch für die Stadt.\n\n"

        "### Stadttext (text_stadt_*, text_einwohner_detail etc.):\n"
        "Nutze ECHTES Wissen über die Stadt. Nenne echte Unternehmen, Branchen, Hochschulen, Investitionen.\n"
        "Stil: 'Magdeburg verzeichnet seit Jahren ein kontinuierliches Wachstum: mehr als 245.000 Einwohner, "
        "über 21.000 Studierende, steigende Mieten im Neubausegment und ein deutlich wachsendes BIP.'\n\n"

        "### Investmenttext (text_investment_pitch, text_kapitel_invest_1/2):\n"
        "Konkret: Einstiegspreis nennen, KfW-Darlehen, 3-fach AfA erklären, Rendite/Mietperspektive.\n"
        "Beispiel: 'Kleine Einstiegspreise, attraktive KfW-Förderung und dreifach-AfA bieten ideale "
        "Voraussetzungen für Kapitalanleger, die Wert auf Effizienz und Stabilität legen.'\n\n"

        "### Nachhaltigkeit (text_greenliving_1, text_greenliving_2, text_projekt_nachhaltig_*):\n"
        "text_greenliving_1: 5-6 Sätze über Fernwärme, Photovoltaik, Gründach, E-Ladeinfrastruktur.\n"
        "text_greenliving_2: 4-5 Sätze über Außenbereiche, Bepflanzung, Mikroklima, Lebensqualität.\n\n"

        "### Ausstattung (text_ausstattung_detail, text_ausstattung_intro):\n"
        "text_ausstattung_detail: 4-5 Sätze. Konkret: Bodenbeläge, Fliesen, Fußbodenheizung, "
        "Beleuchtung, Barrierefreiheit, Balkone/Terrassen.\n\n"

        f"## STANDORT-MINUTEN ({stadt} – Slide 5):\n"
        f"min_uni / label_min_uni: Fahrradminuten + Name der nächsten Uni/FH in {stadt}\n"
        f"min_bahnhof / label_min_bahnhof: Fahrradminuten + Hauptbahnhof\n"
        f"min_altstadt / label_min_altstadt: Fahrradminuten + Altstadt/Innenstadt\n"
        f"WICHTIG: 'min_*'-Felder nur die Zahl, z.B. '3'. 'label_min_*' nur den Namen, z.B. 'Leibniz Universität'.\n\n"

        f"## FREIZEIT NAH ({stadt} – Slide 14, 4 Einträge):\n"
        f"freizeit_N_name: ECHTER Name (Park, See, Sehenswürdigkeit) in {stadt}\n"
        f"min_freizeit_N: Gehminuten als Zahl\n\n"

        f"## WOHNUNGSTYPEN:\n"
        f"Analysiere die Grundrisse aus den Projektdaten. Pro WE-Paar (je 2 nebeneinander):\n"
        f"- we_beispiel_N: 'WE 02' (linke Spalte), we_beispiel_N+1: 'WE 07' (rechte Spalte)\n"
        f"- we_bereich_N: Hauptbereich des Typs, z.B. 'Wohnen & Schlafen', 'Wohnen/Kochen'\n"
        f"- we_flaeche_1-5: NUR Quadratmeter als '23,99 m²' (Raumnamen sind hardcoded im Template)\n"
        f"- we_flaeche_5: Gesamtfläche\n"
        f"- we_typ_beschreibung: 2-3 Sätze Beschreibung des Wohnungstyps mit Zielgruppe\n"
        f"Slide 1 (immer): Paar 1 (we_beispiel_1, we_bereich_1, we_beispiel_2, we_bereich_2)\n"
        f"Slide 2 (wenn ≥2 Typen): Paar 2 (we_beispiel_3..4), leere Strings wenn nicht benötigt\n"
        f"Slide 3 (wenn ≥3 Typen): Paar 3 (we_beispiel_5..6)\n"
        f"Und so weiter für weitere Typen.\n\n"

        f"## STADTSTATISTIKEN ({stadt}):\n"
        f"Verwende echte, aktuelle Zahlen für {stadt}:\n"
        f"stadt_einwohner: Einwohnerzahl als formatierte Zahl, z.B. '245.279'\n"
        f"bundesland_bip: BIP des Bundeslandes NUR als Zahl+Einheit OHNE 'EUR'/'Euro', z.B. '310 Mrd.' oder '78,4 Mrd.'\n"
        f"  (Das Template-Label schreibt 'in €' bereits dahinter – niemals doppelt!)\n"
        f"stadt_mietsteigerung: Mietsteigerung seit 2017/2018, z.B. '+31%'\n"
        f"stadt_studierende: Studierende an Hochschulen, z.B. '21.000'\n\n"

        f"## ALLE FELDER – PFLICHT:\n"
        f"Jedes Feld MUSS befüllt werden. Leere Strings sind nicht akzeptabel außer bei\n"
        f"we_beispiel_N/we_bereich_N für nicht vorhandene WE-Typen.\n\n"
        f"{json.dumps(PLATZHALTER, ensure_ascii=False)}"
    )
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": CLAUDE_API_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"},
        json={
            "model": "claude-sonnet-4-6", "max_tokens": 16000,
            "messages": [{"role": "user", "content": prompt}]
        },
        timeout=300
    )
    resp.raise_for_status()
    json_text = ""
    for block in resp.json()["content"]:
        if block.get("type") == "text":
            json_text = block["text"]
    json_text = json_text.replace("```json", "").replace("```", "").strip()
    if not json_text:
        raise ValueError("Claude hat keinen Text zurückgegeben. Stop-Reason: " +
                         str(resp.json().get("stop_reason")))
    return json.loads(json_text)

# Regex: matcht {{KEY}}, {{KEY-SPLIT}}, {{KEY|suffix}}, {{KEY | suffix}}
# Bindestriche im Key werden beim Lookup entfernt (z.B. {{PRODUKT_BESCHREI-BUNG}})
_PH_RE = re.compile(r'\{\{\s*([A-Z0-9_-]+)\s*(?:\|[^}]*)?\}\}', re.IGNORECASE)

def _replace_placeholders(text, data):
    """Ersetzt alle {{KEY}} und {{KEY|suffix}} Platzhalter. Case-insensitiv.
    Bindestriche im Key werden vor dem Lookup entfernt (für Template-Zeilenumbrüche).
    """
    repl_map = {k.upper(): str(v or "") for k, v in data.items()}
    def _sub(m):
        key = m.group(1).upper().strip().replace('-', '')
        return repl_map.get(key, m.group(0))
    return _PH_RE.sub(_sub, text)

# Regex für Cross-Paragraph-Splits mit optionalem Bindestrich am Zeilenende
# z.B. "{{PRODUKT_BESCHREI-\nBUNG}}" → key = "PRODUKT_BESCHREIBUNG"
_PH_SPLIT_RE = re.compile(
    r'\{\{\s*([A-Z0-9_]+-?)\n([A-Z0-9_]*)\s*(?:\|[^}]*)?\}\}',
    re.IGNORECASE
)

def _replace_split_placeholder(text, data):
    """Behandelt Platzhalter die mit Bindestrich über zwei Zeilen gesplittet sind."""
    repl_map = {k.upper(): str(v or "") for k, v in data.items()}
    def _sub(m):
        part1 = m.group(1).rstrip('-')
        part2 = m.group(2)
        key = (part1 + part2).upper().strip()
        return repl_map.get(key, m.group(0))
    return _PH_SPLIT_RE.sub(_sub, text)


def duplicate_slide(prs, slide_index):
    """Duplicates the slide at slide_index and inserts the copy at slide_index + 1."""
    template = prs.slides[slide_index]
    new_slide = prs.slides.add_slide(template.slide_layout)

    # Replace shape tree with a deep copy of the template's
    sp_tree = new_slide.shapes._spTree
    tmpl_sp_tree = template.shapes._spTree

    # Remove all children added by add_slide (keep only nvGrpSpPr + grpSpPr = first 2)
    for child in list(sp_tree)[2:]:
        sp_tree.remove(child)

    # Copy shapes from template (skip first 2 as well, copy the rest)
    for child in list(tmpl_sp_tree)[2:]:
        sp_tree.append(deepcopy(child))

    # Move new slide (currently last) to position slide_index + 1
    sldIdLst = prs.slides._sldIdLst
    moved_el = sldIdLst[-1]
    sldIdLst.remove(moved_el)
    sldIdLst.insert(slide_index + 1, moved_el)

    return prs.slides[slide_index + 1]


def duplicate_we_slides(prs, data):
    """
    Dynamisch: Originale Slide hat Typen 1+2 (a/b).
    Für jedes weitere befüllte Typ-Paar (3+4, 5+6, 7+8, ...) wird ein Duplikat erstellt.
    Unterstützt beliebig viele WE-Typen (nicht nur bis f).
    Dupliziert nur wenn TEXT-Keys (we_beispiel_N / we_bereich_N) befüllt sind —
    Bild-URLs allein triggern keine Duplikation.
    """
    from pptx.oxml import parse_xml
    import string

    # Ermittle Anzahl extra Slides dynamisch:
    # Pair 1 (original): types 1+2, Pair 2: 3+4, Pair 3: 5+6, ...
    # extra_slides = Anzahl der befüllten Paare ab Pair 2
    extra_slides = 0
    pair = 2
    while True:
        left_n  = pair * 2 - 1   # 3, 5, 7, 9, ...
        right_n = pair * 2        # 4, 6, 8, 10, ...
        if data.get(f"we_beispiel_{left_n}") or data.get(f"we_bereich_{left_n}") \
                or data.get(f"we_beispiel_{right_n}") or data.get(f"we_bereich_{right_n}"):
            extra_slides += 1
            pair += 1
        else:
            break

    # Find original WE slide
    we_idx = None
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            txt = ""
            if shape.has_text_frame:
                txt = shape.text_frame.text.upper()
            elif shape.shape_type == 6:
                for c in shape.shapes:
                    if c.has_text_frame:
                        txt += c.text_frame.text.upper()
            if "WE_BEISPIEL_1" in txt or "WE_BEREICH_1" in txt or "BILD_WE_1" in txt:
                we_idx = i
                break
        if we_idx is not None:
            break

    if we_idx is None:
        print("duplicate_we_slides: WE-Slide nicht gefunden – überspringe Buchstaben-Fix")
        return

    print(f"WE-Slide bei Index {we_idx} → {extra_slides} extra Slide(s) "
          f"({2 + extra_slides * 2} WE-Typen insgesamt)")

    # Buchstaben für Dekorativ-Beschriftung: a, b, c, d, e, f, g, h, ...
    # Für Pair k: linker Buchstabe = letters[2*(k-1)], rechter = letters[2*(k-1)+1]
    all_letters = string.ascii_lowercase  # a-z (genug für 13 WE-Paare)

    def _fix_we_letters(xml_str, left_letter, right_letter):
        """Ersetzt die zwei dekorativen 'a'-Buchstaben (links/rechts) im XML."""
        tag = '<a:t>a</a:t>'
        first_pos = xml_str.find(tag)
        if first_pos < 0:
            return xml_str
        last_pos = xml_str.rfind(tag)
        rep_left  = f'<a:t>{left_letter}</a:t>'
        rep_right = f'<a:t>{right_letter}</a:t>'
        if first_pos == last_pos:
            return xml_str[:first_pos] + rep_left + xml_str[first_pos + len(tag):]
        # Letzten (rechts) zuerst ersetzen, damit first_pos valide bleibt
        result = xml_str[:last_pos] + rep_right + xml_str[last_pos + len(tag):]
        fp = result.find(tag)
        if fp >= 0:
            result = result[:fp] + rep_left + result[fp + len(tag):]
        return result

    # Original-XML VOR jeglicher Modifikation speichern
    orig_xml = etree.tostring(prs.slides[we_idx].shapes._spTree, encoding="unicode")

    # Duplikate in UMGEKEHRTER Reihenfolge erstellen:
    # Jedes neue Slide wird bei we_idx+1 eingefügt → vorherige rücken vor
    # → Endergebnis: Pair2 bei we_idx+1, Pair3 bei we_idx+2, ...
    for offset in range(extra_slides, 0, -1):
        pair_num    = offset + 1                    # 2, 3, 4, ...
        left_n      = pair_num * 2 - 1              # 3, 5, 7, ...
        right_n     = pair_num * 2                  # 4, 6, 8, ...
        left_letter  = all_letters[(pair_num - 1) * 2]         # c, e, g, ...
        right_letter = all_letters[(pair_num - 1) * 2 + 1]     # d, f, h, ...

        new_slide = duplicate_slide(prs, we_idx)
        sp_tree   = new_slide.shapes._spTree

        # Immer vom Original-XML ausgehen
        xml_str = orig_xml

        # Platzhalter umbenennen (rechts vor links um Präfix-Kollisionen zu vermeiden)
        xml_str = xml_str.replace("WE_BEISPIEL_2", f"WE_BEISPIEL_{right_n}")
        xml_str = xml_str.replace("WE_BEISPIEL_1", f"WE_BEISPIEL_{left_n}")
        xml_str = xml_str.replace("WE_BEREICH_2",  f"WE_BEREICH_{right_n}")
        xml_str = xml_str.replace("WE_BEREICH_1",  f"WE_BEREICH_{left_n}")
        xml_str = xml_str.replace("BILD_WE_2",     f"BILD_WE_{right_n}")
        xml_str = xml_str.replace("BILD_WE_1",     f"BILD_WE_{left_n}")
        # WE_FLAECHE_1-5 und WE_TYP_BESCHREIBUNG sind pro Slide geteilt → unverändert

        xml_str = _fix_we_letters(xml_str, left_letter, right_letter)

        new_sp_tree = parse_xml(xml_str.encode("utf-8"))
        for child in list(sp_tree):
            sp_tree.remove(child)
        for child in list(new_sp_tree):
            sp_tree.append(child)

    # Original-Slide: rechtes 'a' → 'b' (linkes bleibt 'a')
    orig_sp_tree = prs.slides[we_idx].shapes._spTree
    fixed_xml    = _fix_we_letters(
        etree.tostring(orig_sp_tree, encoding="unicode"), 'a', 'b'
    )
    new_orig_sp_tree = parse_xml(fixed_xml.encode("utf-8"))
    for child in list(orig_sp_tree):
        orig_sp_tree.remove(child)
    for child in list(new_orig_sp_tree):
        orig_sp_tree.append(child)


def fill_pptx(template_bytes, data, customer_images=None):
    """Fill PPTX template using python-pptx: embeds images and replaces text placeholders.
    customer_images: optional dict {bild_key: bytes} – Kundenbilder haben Vorrang vor URLs."""

    prs = Presentation(io.BytesIO(template_bytes))

    # Pre-load images: Kundenbilder zuerst, dann URLs für fehlende Slots
    image_data = {}

    # 1. Kundenbilder direkt einsetzen (höchste Priorität)
    if customer_images:
        for key, raw in customer_images.items():
            image_data[key] = raw
            print(f"  ✓ Kundenbild: {key} ({len(raw)//1024} KB)")
        print(f"  Kundenbilder: {len(customer_images)} eingeladen")

    # 2. Unsplash/Picsum-URLs für alle noch fehlenden bild_* Keys herunterladen
    bild_keys_total = [k for k in data if k.startswith("bild_")]
    bild_keys_with_url = [k for k in bild_keys_total
                          if isinstance(data[k], str) and data[k].startswith("http")
                          and k not in image_data]
    print(f"fill_pptx: {len(bild_keys_total)} bild_* Keys, {len(bild_keys_with_url)} noch per URL zu laden")

    for key in bild_keys_with_url:
        value = data[key]
        try:
            resp = requests.get(value, timeout=15)
            if resp.status_code == 200:
                image_data[key] = resp.content
                print(f"  ✓ Bild geladen: {key} ({len(resp.content)//1024} KB)")
            else:
                print(f"  ✗ Bild HTTP-Fehler {key}: {resp.status_code}  URL={value[:80]}")
        except Exception as e:
            print(f"  ✗ Bild Download-Fehler {key}: {e}")

    print(f"  image_data gesamt: {len(image_data)} Bilder")

    def make_replacement_map(data):
        """Build a case-insensitive lookup: UPPER_KEY -> value."""
        return {k.upper(): str(v or "") for k, v in data.items()}

    REPL_MAP = make_replacement_map(data)

    # Regex that matches {{KEY}}, {{KEY|suffix}}, {{KEY | suffix}} (spaces ok around pipe)
    PLACEHOLDER_RE = re.compile(r'\{\{\s*([A-Z0-9_]+)\s*(?:\|[^}]*)?\}\}', re.IGNORECASE)
    # Matches the |Xpt font-size hint inside a placeholder, e.g. {{MIN_UNI|50pt}}
    _SIZE_HINT_RE = re.compile(r'\|\s*(\d+)\s*pt\b', re.IGNORECASE)

    def replace_text(text):
        """Replace all placeholders in a string using REPL_MAP."""
        def _sub(m):
            key = m.group(1).upper().strip()
            return REPL_MAP.get(key, m.group(0))  # keep original if not found
        return PLACEHOLDER_RE.sub(_sub, text)

    def replace_in_paragraph(para):
        """Replace placeholders in a paragraph, handling split-run placeholders.

        Strategy: reassemble all runs into one string, replace, write back into
        runs[0] preserving its formatting, clear all other runs.
        If the placeholder contains a |Xpt font-size hint (e.g. {{MIN_UNI|50pt}}),
        apply that size to the replacement run — Canva PPTX exports lose font sizes.
        """
        if not para.runs:
            return
        full_text = "".join(r.text for r in para.runs)
        if "{{" not in full_text:
            return
        # Extract font-size hint BEFORE stripping the suffix
        size_hint = None
        sh = _SIZE_HINT_RE.search(full_text)
        if sh:
            size_hint = int(sh.group(1))
        modified = replace_text(full_text)
        if modified != full_text:
            para.runs[0].text = modified
            for run in para.runs[1:]:
                run.text = ""
            # Apply explicit font size if hint was present
            if size_hint is not None:
                from pptx.util import Pt
                para.runs[0].font.size = Pt(size_hint)

    def replace_in_textframe(tf):
        """Replace placeholders across entire text frame, including cross-paragraph splits.

        Some placeholders span two paragraphs (e.g. '{{PRODUKT_BESCHREI-' / 'BUNG}}').
        We join the full text frame, detect these, and replace them in para[0] of the span.
        """
        # First do normal per-paragraph replacement
        for para in tf.paragraphs:
            replace_in_paragraph(para)

        # Then handle cross-paragraph splits: join all paragraph texts and check
        # Build list of (para_index, full_run_text) pairs
        para_texts = ["".join(r.text for r in p.runs) for p in tf.paragraphs]
        joined = "\n".join(para_texts)

        # Find any remaining {{...}} that survived (i.e. were split across paragraphs)
        # by scanning pairs of adjacent paragraphs
        for i in range(len(tf.paragraphs) - 1):
            combined = para_texts[i] + para_texts[i + 1]
            if "{{" in combined and "}}" in combined:
                replaced = replace_text(combined)
                if replaced != combined:
                    # Split the replacement back at the original boundary
                    split_at = len(para_texts[i])
                    new_p0 = replaced[:split_at] if len(replaced) >= split_at else replaced
                    new_p1 = replaced[split_at:] if len(replaced) >= split_at else ""
                    # Write to para i
                    p0 = tf.paragraphs[i]
                    if p0.runs:
                        p0.runs[0].text = new_p0
                        for r in p0.runs[1:]:
                            r.text = ""
                    # Write to para i+1
                    p1 = tf.paragraphs[i + 1]
                    if p1.runs:
                        p1.runs[0].text = new_p1
                        for r in p1.runs[1:]:
                            r.text = ""
                    # Update cache
                    para_texts[i] = new_p0
                    para_texts[i + 1] = new_p1

    def get_abs_coords(group_shape, child_shape):
        """Berechnet absolute Slide-Koordinaten eines Child-Shapes in einer Gruppe.
        Nutzt grpSpPr/xfrm aus dem p:-Namespace (nicht a:-Namespace!).
        """
        NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        grp_el  = group_shape._element
        grpSpPr = grp_el.find(f'{{{NS_P}}}grpSpPr')
        if grpSpPr is None:
            # fallback: simple addition
            return (
                (group_shape.left or 0) + (child_shape.left or 0),
                (group_shape.top  or 0) + (child_shape.top  or 0),
                child_shape.width  or 0,
                child_shape.height or 0,
            )
        xfrm  = grpSpPr.find(f'{{{NS_A}}}xfrm')
        if xfrm is None:
            return (
                (group_shape.left or 0) + (child_shape.left or 0),
                (group_shape.top  or 0) + (child_shape.top  or 0),
                child_shape.width  or 0,
                child_shape.height or 0,
            )
        off   = xfrm.find(f'{{{NS_A}}}off')
        ext   = xfrm.find(f'{{{NS_A}}}ext')
        chOff = xfrm.find(f'{{{NS_A}}}chOff')
        chExt = xfrm.find(f'{{{NS_A}}}chExt')
        if None in (off, ext, chOff, chExt):
            return (
                (group_shape.left or 0) + (child_shape.left or 0),
                (group_shape.top  or 0) + (child_shape.top  or 0),
                child_shape.width  or 0,
                child_shape.height or 0,
            )
        grp_x  = int(off.get('x',  0));  grp_y  = int(off.get('y',  0))
        grp_w  = int(ext.get('cx', 1));  grp_h  = int(ext.get('cy', 1))
        ch_x   = int(chOff.get('x', 0)); ch_y   = int(chOff.get('y', 0))
        ch_w   = int(chExt.get('cx', 1));ch_h   = int(chExt.get('cy', 1))
        scale_x = grp_w / ch_w if ch_w else 1
        scale_y = grp_h / ch_h if ch_h else 1
        abs_left = int(grp_x + ((child_shape.left or 0) - ch_x) * scale_x)
        abs_top  = int(grp_y + ((child_shape.top  or 0) - ch_y) * scale_y)
        abs_w    = int((child_shape.width  or 0) * scale_x)
        abs_h    = int((child_shape.height or 0) * scale_y)
        return abs_left, abs_top, abs_w, abs_h

    # Mindestgröße in EMU: nur für echte Null-/Winzig-Shapes als Fallback
    # (10 000 EMU ≈ 0,028 cm – alles darunter gilt als fehlende Dimension)
    MIN_IMG_SIZE = 10_000

    def insert_at_z(slide, pic_element, z_index):
        """Setzt pic_element an Position z_index im spTree (preserviert Z-Order des Originals).
        Indices 0+1 sind nvGrpSpPr/grpSpPr → min. Index 2.
        """
        sp_tree = slide.shapes._spTree
        sp_tree.remove(pic_element)
        sp_tree.insert(max(2, z_index), pic_element)

    def process_shape(slide, shape, image_data):
        """Ersetzt Text oder bettet Bild via blipFill ein.
        Das Template nutzt zwei Gruppen pro Bildslot:
          - Placeholder-Gruppe: solidFill Freeform + TextBox {{BILD_X}} (sichtbar, oben)
          - Target-Gruppe:      blipFill Freeform, kein Text (dahinter, wartet auf rId)
        Strategie:
          Case A (Target existiert): rId in Target-blipFill eintragen → Placeholder-Gruppe entfernen
          Case B (kein Target):      solidFill der Freeform im Placeholder durch blipFill ersetzen,
                                     TextBox-Kind entfernen
        """
        from lxml import etree

        _NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        _NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        _NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

        # ── Gruppe ──────────────────────────────────────────────────────────────
        if shape.shape_type == 6:
            # Suche TextBox-Kind mit BILD_-Placeholder
            bild_child = None
            for child in shape.shapes:
                if child.has_text_frame:
                    txt = child.text_frame.text.strip()
                    m = PLACEHOLDER_RE.match(txt)
                    if m:
                        key = m.group(1).lower()
                        if key in image_data and image_data[key]:
                            bild_child = (child, key)
                            break

            if bild_child is not None:
                child, key = bild_child
                try:
                    print(f"  → blipFill: key={key!r} in Gruppe {shape.name!r}")

                    # ── Schritt 1: Bild einbetten → rId holen ─────────────────
                    temp_pic = slide.shapes.add_picture(
                        io.BytesIO(image_data[key]), 0, 0, 914400, 914400
                    )
                    temp_el = temp_pic._element
                    blip_in_temp = temp_el.find(f'.//{{{_NS_A}}}blip')
                    new_rid = blip_in_temp.get(f'{{{_NS_R}}}embed') if blip_in_temp is not None else None
                    temp_el.getparent().remove(temp_el)   # p:pic entfernen; Relation bleibt

                    if not new_rid:
                        print(f"  FEHLER: kein rId für {key!r}")
                        return

                    print(f"    rId={new_rid!r}")

                    # ── Schritt 2: Position der Placeholder-Gruppe ─────────────
                    ph_grpSpPr = shape._element.find(f'{{{_NS_P}}}grpSpPr')
                    ph_xfrm    = ph_grpSpPr.find(f'{{{_NS_A}}}xfrm') if ph_grpSpPr is not None else None
                    ph_off     = ph_xfrm.find(f'{{{_NS_A}}}off')     if ph_xfrm    is not None else None
                    ph_x = ph_off.get('x', '0') if ph_off is not None else '0'
                    ph_y = ph_off.get('y', '0') if ph_off is not None else '0'

                    # ── Schritt 3: Target-Gruppe (blipFill, gleiche Position) suchen ──
                    target_info = None   # (target_group_shape, freeform_spPr, blipFill_el)
                    for other in slide.shapes:
                        if other.shape_id == shape.shape_id or other.shape_type != 6:
                            continue
                        o_grpSpPr = other._element.find(f'{{{_NS_P}}}grpSpPr')
                        o_xfrm    = o_grpSpPr.find(f'{{{_NS_A}}}xfrm') if o_grpSpPr is not None else None
                        o_off     = o_xfrm.find(f'{{{_NS_A}}}off')     if o_xfrm    is not None else None
                        if o_off is None:
                            continue
                        if (abs(int(o_off.get('x', 0)) - int(ph_x)) > 50000 or
                                abs(int(o_off.get('y', 0)) - int(ph_y)) > 50000):
                            continue
                        # Gleiche Position → hat diese Gruppe eine blipFill-Freeform?
                        for grp_child_o in other.shapes:
                            sp_pr_o = grp_child_o._element.find(f'{{{_NS_P}}}spPr')
                            if sp_pr_o is None:
                                continue
                            bf = sp_pr_o.find(f'{{{_NS_A}}}blipFill')
                            if bf is not None:
                                target_info = (other, sp_pr_o, bf)
                                break
                        if target_info is not None:
                            break

                    if target_info is not None:
                        # ── Case A: Target-Gruppe aktualisieren ────────────────
                        tgt_grp, tgt_spPr, tgt_blipFill = target_info
                        existing_blip = tgt_blipFill.find(f'{{{_NS_A}}}blip')
                        if existing_blip is not None:
                            existing_blip.set(f'{{{_NS_R}}}embed', new_rid)
                            print(f"  Case A ✓ blip.r:embed={new_rid!r} in {tgt_grp.name!r}")
                        else:
                            new_blip = etree.SubElement(tgt_blipFill, f'{{{_NS_A}}}blip')
                            new_blip.set(f'{{{_NS_R}}}embed', new_rid)
                            print(f"  Case A ✓ neues a:blip in {tgt_grp.name!r}")
                        # Reset fillRect so new image fills the shape without template crop offsets
                        stretch = tgt_blipFill.find(f'{{{_NS_A}}}stretch')
                        if stretch is None:
                            stretch = etree.SubElement(tgt_blipFill, f'{{{_NS_A}}}stretch')
                        fr = stretch.find(f'{{{_NS_A}}}fillRect')
                        if fr is None:
                            etree.SubElement(stretch, f'{{{_NS_A}}}fillRect')
                        else:
                            fr.attrib.clear()   # remove l/t/r/b crop offsets → full fill
                        # Placeholder-Gruppe entfernen (solidFill+TextBox waren oben drüber)
                        shape._element.getparent().remove(shape._element)

                    else:
                        # ── Case B: solidFill in eigener Freeform → blipFill ──
                        print(f"  Case B: kein Target für {key!r} @ ({ph_x},{ph_y})")
                        # TextBox-Placeholder-Kind entfernen
                        for grp_child_b in list(shape.shapes):
                            if grp_child_b.has_text_frame:
                                txt_b = grp_child_b.text_frame.text.strip()
                                if PLACEHOLDER_RE.match(txt_b):
                                    grp_child_b._element.getparent().remove(grp_child_b._element)
                                    break
                        # solidFill in der Freeform durch blipFill ersetzen
                        for grp_child_b in list(shape.shapes):
                            sp_pr_b = grp_child_b._element.find(f'{{{_NS_P}}}spPr')
                            if sp_pr_b is None:
                                continue
                            solid = sp_pr_b.find(f'{{{_NS_A}}}solidFill')
                            if solid is None:
                                continue
                            idx = list(sp_pr_b).index(solid)
                            sp_pr_b.remove(solid)
                            bf_el = etree.Element(f'{{{_NS_A}}}blipFill')
                            bl_el = etree.SubElement(bf_el, f'{{{_NS_A}}}blip')
                            bl_el.set(f'{{{_NS_R}}}embed', new_rid)
                            st_el = etree.SubElement(bf_el, f'{{{_NS_A}}}stretch')
                            etree.SubElement(st_el, f'{{{_NS_A}}}fillRect')
                            sp_pr_b.insert(idx, bf_el)
                            print(f"  Case B ✓ solidFill→blipFill in {grp_child_b.name!r}")
                            break

                except Exception as e:
                    print(f"  blipFill Fehler {shape.name!r}/{key!r}: {e}")
                    import traceback; traceback.print_exc()
                return

            # Keine BILD_-Gruppe → Text in allen Kindern ersetzen
            for child in list(shape.shapes):
                try:
                    if child.has_text_frame:
                        replace_in_textframe(child.text_frame)
                except Exception as e:
                    print(f"  Gruppen-Child Fehler {child.name}: {e}")
            if shape.has_text_frame:
                replace_in_textframe(shape.text_frame)
            return

        # ── Top-Level Non-Group Shape ────────────────────────────────────────
        sp_tree = slide.shapes._spTree
        shape_name_lower = (shape.name or "").lower()

        # 1. Bild per Shape-Name
        if shape_name_lower in image_data and image_data[shape_name_lower]:
            try:
                left   = shape.left   or 0
                top    = shape.top    or 0
                width  = shape.width  or 0
                height = shape.height or 0
                if width  < MIN_IMG_SIZE: width  = 3_000_000
                if height < MIN_IMG_SIZE: height = 2_400_000
                shape_z = list(sp_tree).index(shape._element)
                shape._element.getparent().remove(shape._element)
                _pic = slide.shapes.add_picture(
                    io.BytesIO(image_data[shape_name_lower]), left, top, width, height
                )
                insert_at_z(slide, _pic._element, shape_z)
                return
            except Exception as e:
                print(f"Bild-Ersatz Fehler (name) {shape_name_lower}: {e}")

        # 2. Bild per Text-Inhalt
        if shape.has_text_frame:
            txt = shape.text_frame.text.strip()
            m = PLACEHOLDER_RE.match(txt)
            if m:
                key = m.group(1).lower()
                if key in image_data and image_data[key]:
                    try:
                        # Check if this placeholder TextBox sits inside a solidFill group
                        # (e.g. BILD_LAGEPLAN inside dark left-panel Group) → replace that
                        # group's solidFill with the image instead of inserting a tiny picture.
                        ph_cx = (shape.left or 0) + (shape.width or 0) // 2
                        ph_cy = (shape.top  or 0) + (shape.height or 0) // 2
                        covering_target = None
                        for other in slide.shapes:
                            if other.shape_id == shape.shape_id or other.shape_type != 6:
                                continue
                            g_left  = other.left  or 0
                            g_top   = other.top   or 0
                            g_right = g_left + (other.width  or 0)
                            g_bot   = g_top  + (other.height or 0)
                            if not (g_left <= ph_cx <= g_right and g_top <= ph_cy <= g_bot):
                                continue
                            for gc in other.shapes:
                                sp_pr_gc = gc._element.find(f'{{{_NS_P}}}spPr')
                                if sp_pr_gc is None:
                                    continue
                                solid_gc = sp_pr_gc.find(f'{{{_NS_A}}}solidFill')
                                if solid_gc is None:
                                    continue
                                covering_target = (other, gc, sp_pr_gc, solid_gc)
                                break
                            if covering_target:
                                break

                        if covering_target:
                            # Embed image → get rId
                            grp, grp_child, sp_pr_gc, solid_gc = covering_target
                            temp_pic = slide.shapes.add_picture(
                                io.BytesIO(image_data[key]), 0, 0, 914400, 914400
                            )
                            temp_el = temp_pic._element
                            blip_in_temp = temp_el.find(f'.//{{{_NS_A}}}blip')
                            new_rid = blip_in_temp.get(f'{{{_NS_R}}}embed') if blip_in_temp is not None else None
                            temp_el.getparent().remove(temp_el)
                            if new_rid:
                                idx = list(sp_pr_gc).index(solid_gc)
                                sp_pr_gc.remove(solid_gc)
                                bf_el = etree.Element(f'{{{_NS_A}}}blipFill')
                                bl_el = etree.SubElement(bf_el, f'{{{_NS_A}}}blip')
                                bl_el.set(f'{{{_NS_R}}}embed', new_rid)
                                st_el = etree.SubElement(bf_el, f'{{{_NS_A}}}stretch')
                                etree.SubElement(st_el, f'{{{_NS_A}}}fillRect')
                                sp_pr_gc.insert(idx, bf_el)
                                shape._element.getparent().remove(shape._element)
                                print(f"  Panel fill: {key!r} → group {grp.name!r}")
                                return

                        # Fallback: insert picture at TextBox dimensions
                        left   = shape.left   or 0
                        top    = shape.top    or 0
                        width  = shape.width  or 0
                        height = shape.height or 0
                        if width  < MIN_IMG_SIZE: width  = 3_000_000
                        if height < MIN_IMG_SIZE: height = 2_400_000
                        shape_z = list(sp_tree).index(shape._element)
                        shape._element.getparent().remove(shape._element)
                        _pic = slide.shapes.add_picture(
                            io.BytesIO(image_data[key]), left, top, width, height
                        )
                        insert_at_z(slide, _pic._element, shape_z)
                        return
                    except Exception as e:
                        print(f"Bild-Ersatz Fehler (text) {key}: {e}")

        # 3. Text ersetzen
        if shape.has_text_frame:
            replace_in_textframe(shape.text_frame)

    # Duplicate WE slides BEFORE text/image replacement so placeholders are still intact
    duplicate_we_slides(prs, data)

    for slide in prs.slides:
        for shape in list(slide.shapes):
            try:
                process_shape(slide, shape, image_data)
            except Exception as e:
                print(f"Shape-Fehler slide={slide.slide_id} shape={shape.name}: {e}")

    # Cleanup-Pass: Template-Texte die "in Euro" statt "in €" enthalten korrigieren
    _euro_fixes = [("in Euro", "in €"), (" Euro", " €"), ("in EUR", "in €"), (" EUR", " €")]
    for slide in prs.slides:
        for shape in slide.shapes:
            shapes_to_check = [shape]
            if shape.shape_type == 6:
                shapes_to_check += list(shape.shapes)
            for s in shapes_to_check:
                if not s.has_text_frame:
                    continue
                for para in s.text_frame.paragraphs:
                    for run in para.runs:
                        for old, new in _euro_fixes:
                            if old in run.text:
                                run.text = run.text.replace(old, new)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()

def convert_to_pdf(pptx_bytes, filename):
    import time
    if not CLOUDCONVERT_KEY:
        raise RuntimeError("CLOUDCONVERT_KEY ist nicht gesetzt (leere Umgebungsvariable)")
    print(f"convert_to_pdf: starte CloudConvert für {filename} ({len(pptx_bytes)//1024} KB)")
    cc_headers = {"Authorization": f"Bearer {CLOUDCONVERT_KEY}", "Content-Type": "application/json"}

    # Job erstellen (async API)
    job_resp = requests.post(
        "https://api.cloudconvert.com/v2/jobs",
        headers=cc_headers,
        json={"tasks": {
            "upload":  {"operation": "import/upload"},
            "convert": {"operation": "convert", "input": "upload",
                        "input_format": "pptx", "output_format": "pdf", "engine": "office"},
            "export":  {"operation": "export/url", "input": "convert"}
        }}, timeout=30
    )
    if not job_resp.ok:
        raise RuntimeError(f"CloudConvert Job-Erstellung fehlgeschlagen: {job_resp.status_code} – {job_resp.text[:300]}")
    job = job_resp.json()["data"]
    job_id = job["id"]

    # Datei hochladen
    upload_task = next(t for t in job["tasks"] if t["name"] == "upload")
    form = upload_task["result"]["form"]
    files = {"file": (filename, pptx_bytes,
                      "application/vnd.openxmlformats-officedocument.presentationml.presentation")}
    requests.post(form["url"], data=form.get("parameters", {}), files=files, timeout=60).raise_for_status()

    # Warten bis Job fertig (max 5 Minuten, alle 5s pollen)
    for _ in range(60):
        time.sleep(5)
        status_resp = requests.get(
            f"https://api.cloudconvert.com/v2/jobs/{job_id}",
            headers=cc_headers, timeout=30
        )
        status_resp.raise_for_status()
        job_status = status_resp.json()["data"]["status"]
        if job_status == "finished":
            tasks = status_resp.json()["data"]["tasks"]
            pdf_url = next(t for t in tasks if t["name"] == "export")["result"]["files"][0]["url"]
            return requests.get(pdf_url, timeout=60).content
        if job_status == "error":
            tasks = status_resp.json()["data"]["tasks"]
            err = next((t.get("message","") for t in tasks if t.get("status") == "error"), "Unbekannter Fehler")
            raise RuntimeError(f"CloudConvert Fehler: {err}")

    raise RuntimeError("CloudConvert Timeout nach 5 Minuten")

def assemble_session(session_id):
    """Liest alle Chunks einer Session von /tmp und gibt die assemblierten Bytes zurück."""
    session_dir = os.path.join(CHUNK_DIR, session_id)
    if not os.path.isdir(session_dir):
        raise ValueError(f"Session {session_id} nicht gefunden")
    meta_path = os.path.join(session_dir, "meta.json")
    with open(meta_path) as f:
        meta = json.load(f)
    total = meta["total_chunks"]
    assembled = b""
    for i in range(total):
        chunk_path = os.path.join(session_dir, f"chunk_{i:04d}")
        with open(chunk_path, "rb") as f:
            assembled += f.read()
    shutil.rmtree(session_dir, ignore_errors=True)
    return assembled


@app.route("/upload-chunk", methods=["POST", "OPTIONS"])
def upload_chunk():
    if request.method == "OPTIONS":
        return make_response("", 204)
    if request.headers.get("X-API-Token") != API_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401
    try:
        session_id = request.form.get("session_id") or str(uuid.uuid4())
        chunk_index = int(request.form.get("chunk_index", 0))
        total_chunks = int(request.form.get("total_chunks", 1))
        filename = request.form.get("filename", "upload.zip")

        chunk_file = request.files.get("chunk")
        if not chunk_file:
            return jsonify({"error": "Kein 'chunk' im Request"}), 400

        session_dir = os.path.join(CHUNK_DIR, session_id)
        os.makedirs(session_dir, exist_ok=True)

        chunk_path = os.path.join(session_dir, f"chunk_{chunk_index:04d}")
        chunk_file.save(chunk_path)

        meta_path = os.path.join(session_dir, "meta.json")
        if not os.path.exists(meta_path):
            with open(meta_path, "w") as f:
                json.dump({"total_chunks": total_chunks, "filename": filename}, f)

        received = len([n for n in os.listdir(session_dir) if n.startswith("chunk_")])
        ready = received >= total_chunks
        return jsonify({"session_id": session_id, "chunk": chunk_index,
                        "received": received, "total": total_chunks, "ready": ready})
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/")
def index():
    index_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")
    with open(index_path, encoding="utf-8") as f:
        return f.read(), 200, {"Content-Type": "text/html; charset=utf-8"}

@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "INTERPRES Full Pipeline v3",
                    "test_mode": TEST_MODE})

@app.route("/generate-expose", methods=["POST"])
def generate_expose():
    if request.headers.get("X-API-Token") != API_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401
    try:
        pdfs = []
        customer_image_list = []   # Bilder aus der ZIP (Kundenbilder)

        # --- Chunked-Session Upload ---
        session_ids = request.form.getlist("session_ids")
        if session_ids:
            for sid in session_ids:
                zip_bytes = assemble_session(sid)
                pdfs.extend(extract_pdfs_from_zip(zip_bytes))
                customer_image_list.extend(extract_images_from_zip(zip_bytes))

        elif request.content_type and "multipart" in request.content_type:
            uploaded = request.files.getlist("files") or request.files.getlist("file")
            if not uploaded:
                return jsonify({"error": "Keine Dateien im Request"}), 400
            for f in uploaded:
                raw = f.read()
                pdfs.extend(extract_pdfs_from_zip(raw))
                customer_image_list.extend(extract_images_from_zip(raw))
        else:
            body = request.get_json(force=True) or {}
            if "zip_base64_list" in body:
                for b64 in body["zip_base64_list"]:
                    raw = base64.b64decode(b64)
                    pdfs.extend(extract_pdfs_from_zip(raw))
                    customer_image_list.extend(extract_images_from_zip(raw))
            elif "zip_base64" in body:
                raw = base64.b64decode(body["zip_base64"])
                pdfs.extend(extract_pdfs_from_zip(raw))
                customer_image_list.extend(extract_images_from_zip(raw))
            else:
                return jsonify({"error": "zip_base64 oder zip_base64_list fehlt"}), 400

        if not pdfs:
            return jsonify({"error": "Keine relevanten PDFs gefunden"}), 400

        # Max. 3 PDFs senden (Kostenkontrolle)
        pdfs = sorted(pdfs, key=lambda x: x["priority"])[:3]

        if TEST_MODE:
            print("TEST_MODE aktiv – überspringe Claude API")
            expose_data = DUMMY_EXPOSE_DATA.copy()
        else:
            projektdaten = analyze_pdfs_with_claude(pdfs)
            raw_expose = generate_expose_with_claude(projektdaten)
            # Merge: PLATZHALTER-Defaults sicherstellen, damit alle bild_* Keys existieren
            expose_data = {**PLATZHALTER, **raw_expose}
            expose_data["logo_initial"] = generate_logo_initial(expose_data.get("projekt_name", ""))

        # Kundenbilder klassifizieren (Vorrang vor Unsplash)
        customer_images = {}
        if customer_image_list:
            print(f"generate_expose: {len(customer_image_list)} Kundenbilder → klassifiziere…")
            customer_images = classify_and_assign_customer_images(customer_image_list)

        # Unsplash/Picsum-Fallback nur für Slots ohne Kundenbild
        expose_data = fill_image_placeholders(expose_data)
        bild_count = sum(1 for k, v in expose_data.items()
                         if k.startswith("bild_") and isinstance(v, str) and v.startswith("http"))
        print(f"generate_expose: {bild_count} bild_* URLs, {len(customer_images)} Kundenbilder vor fill_pptx")

        tmpl_bytes = requests.get(TEMPLATE_URL, timeout=30).content
        pptx_bytes = fill_pptx(tmpl_bytes, expose_data, customer_images=customer_images)

        projekt_name = expose_data.get("projekt_name", "Expose").replace(" ", "_")
        pdf_bytes = convert_to_pdf(pptx_bytes, f"{projekt_name}.pptx")

        return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf",
                         as_attachment=True, download_name=f"{projekt_name}_Expose.pdf")

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/debug-images", methods=["GET"])
def debug_images():
    """Testet den kompletten Bild-Pipeline ohne Upload: Unsplash → Download → PPTX.
    Gibt JSON-Bericht zurück. Nur mit korrektem API-Token aufrufbar."""
    if request.headers.get("X-API-Token") != API_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401

    report = {
        "unsplash_key_set": bool(UNSPLASH_ACCESS_KEY),
        "unsplash_key_len": len(UNSPLASH_ACCESS_KEY) if UNSPLASH_ACCESS_KEY else 0,
        "unsplash_test": None,
        "picsum_test": None,
        "image_urls_count": 0,
        "image_downloads_ok": 0,
        "image_downloads_fail": 0,
    }

    # Unsplash Test
    try:
        r = requests.get(
            "https://api.unsplash.com/photos/random",
            params={"query": "modern apartment", "orientation": "landscape"},
            headers={"Authorization": f"Client-ID {UNSPLASH_ACCESS_KEY}"},
            timeout=10
        )
        report["unsplash_test"] = {"status": r.status_code, "body_preview": r.text[:200]}
        if r.status_code == 200:
            report["unsplash_sample_url"] = r.json()["urls"]["regular"]
    except Exception as e:
        report["unsplash_test"] = {"error": str(e)}

    # Picsum Test
    try:
        r = requests.get("https://picsum.photos/seed/42/200/150", timeout=10)
        report["picsum_test"] = {"status": r.status_code, "size_kb": len(r.content) // 1024}
    except Exception as e:
        report["picsum_test"] = {"error": str(e)}

    # Voller Image-Flow mit Dummy-Daten
    data = DUMMY_EXPOSE_DATA.copy()
    data = fill_image_placeholders(data)
    urls = {k: v for k, v in data.items() if k.startswith("bild_") and isinstance(v, str) and v.startswith("http")}
    report["image_urls_count"] = len(urls)
    report["sample_urls"] = dict(list(urls.items())[:3])

    for key, url in urls.items():
        try:
            r = requests.get(url, timeout=10)
            if r.status_code == 200:
                report["image_downloads_ok"] += 1
            else:
                report["image_downloads_fail"] += 1
        except Exception:
            report["image_downloads_fail"] += 1

    return jsonify(report)


@app.route("/fill-pptx", methods=["POST"])
def fill_pptx_endpoint():
    """Debug-Endpunkt: Gibt das rohe PPTX ohne PDF-Konvertierung zurück.
    Body: JSON mit optionalem 'data'-Objekt. Ohne 'data' → DUMMY_EXPOSE_DATA + Unsplash."""
    if request.headers.get("X-API-Token") != API_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401
    try:
        body = request.get_json(force=True) or {}
        data = body.get("data") or DUMMY_EXPOSE_DATA.copy()
        data = {**PLATZHALTER, **data}
        data = fill_image_placeholders(data)

        bild_count = sum(1 for k, v in data.items()
                         if k.startswith("bild_") and isinstance(v, str) and v.startswith("http"))
        print(f"fill-pptx endpoint: {bild_count} bild_* URLs")

        tmpl_bytes = requests.get(TEMPLATE_URL, timeout=30).content
        pptx_bytes = fill_pptx(tmpl_bytes, data)

        projekt_name = data.get("projekt_name", "Debug").replace(" ", "_")
        return send_file(io.BytesIO(pptx_bytes),
                         mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                         as_attachment=True,
                         download_name=f"{projekt_name}_debug.pptx")
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
