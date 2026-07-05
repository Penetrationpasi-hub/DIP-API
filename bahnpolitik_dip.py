#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bahnpolitik-Tracker ueber die DIP-API des Deutschen Bundestags.
Gleiche Machart wie der MEX13-Sammler: laeuft in Pydroid oder GitHub Actions.

Was das Skript tut:
  1. Sucht Drucksachen (Antraege, Gesetzentwuerfe, Kleine/Grosse Anfragen usw.)
     zu Bahn- und Schienenthemen ueber die Titelsuche der DIP-API
  2. Merkt sich, welche Fraktion bzw. welches Organ die Drucksache
     eingebracht hat (Feld "urheber")
  3. Haengt neue Treffer an bahnpolitik.sqlite an
  4. Schreibt aus dem Gesamtbestand die Excel bahnpolitik_gesamt.xlsx
     mit Auswertung je Fraktion, je Drucksachentyp und im Zeitverlauf

Voraussetzungen:
  - pip: requests und openpyxl
  - API-Key der DIP-API. Entweder als Umgebungsvariable DIP_API_KEY
    (fuer GitHub Actions als Secret) oder in zugangsdaten.txt im selben
    Ordner als Zeile:
        DIP_API_KEY=...
    Den aktuellen oeffentlichen Key gibt es hier (er wechselt regelmaessig):
        https://dip.bundestag.de/über-dip/hilfe/api
    Einen persoenlichen Key gibt es formlos per Mail an:
        parlamentsdokumentation@bundestag.de
"""

from __future__ import annotations

import os
import sqlite3
import sys
import time as _time
from datetime import datetime

import requests

# ---------------------------------------------------------------- Einstellungen

BASE_URL = "https://search.dip.bundestag.de/api/v1"

# Suchbegriffe fuer die Titelsuche. Jeder Begriff wird einzeln abgefragt,
# Dubletten werden ueber die Drucksachen-ID entfernt.
# Bewusst KEIN blankes "Bahn" (sonst matcht auch "Autobahn").
KEYWORDS = [
    "Schienenverkehr",
    "Schienennetz",
    "Schieneninfrastruktur",
    "Eisenbahn",
    "Deutsche Bahn",
    "Bahnstrecke",
    "Bahnhof",
    "Nahverkehr",
    "Regionalisierungsmittel",
    "Deutschlandtakt",
    "Deutschlandticket",
]

# Ab wann gesammelt wird (Beginn der 20. Wahlperiode). Anpassbar.
DATE_START = "2021-10-26"

# Nur Bundestags-Drucksachen (BT), keine Bundesrats-Drucksachen (BR)
ZUORDNUNG = "BT"

SLEEP_BETWEEN_CALLS = 0.7   # Sekunden, schont das Rate-Limit
MAX_PAGES_PER_QUERY = 200   # Notbremse gegen Endlosschleifen

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "bahnpolitik.sqlite")
XLSX_PATH = os.path.join(BASE_DIR, "bahnpolitik_gesamt.xlsx")
CRED_PATH = os.path.join(BASE_DIR, "zugangsdaten.txt")

SCRIPT_VERSION = 1


# ---------------------------------------------------------------- Zugangsdaten

def load_api_key() -> str:
    key = os.getenv("DIP_API_KEY", "")
    if not key and os.path.exists(CRED_PATH):
        with open(CRED_PATH, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line.startswith("DIP_API_KEY="):
                    key = line.partition("=")[2].strip()
    return key.strip()


# ---------------------------------------------------------------- API

class DipClient:
    def __init__(self, api_key: str):
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"ApiKey {api_key}",
            "Accept": "application/json",
        })

    def drucksachen(self, keyword: str):
        """Liefert alle Drucksachen-Treffer der Titelsuche fuer ein Keyword.
        Folgt der Cursor-Paginierung, bis sich der Cursor nicht mehr aendert."""
        params = {
            "f.titel": keyword,
            "f.zuordnung": ZUORDNUNG,
            "f.datum.start": DATE_START,
        }
        cursor = None
        docs = []
        for _page in range(MAX_PAGES_PER_QUERY):
            if cursor:
                params["cursor"] = cursor
            resp = self.session.get(f"{BASE_URL}/drucksache",
                                    params=params, timeout=30)
            if resp.status_code in (401, 403):
                body = resp.text[:200].replace("\n", " ")
                print(f"\nFEHLER: API verweigert den Zugriff "
                      f"(HTTP {resp.status_code}): {body}")
                print("-> API-Key pruefen. Der oeffentliche Key wechselt")
                print("   regelmaessig, aktueller Stand unter:")
                print("   https://dip.bundestag.de/über-dip/hilfe/api")
                sys.exit(1)
            if resp.status_code == 429:
                _time.sleep(5)
                continue
            resp.raise_for_status()
            data = resp.json()
            docs.extend(data.get("documents", []))
            new_cursor = data.get("cursor")
            _time.sleep(SLEEP_BETWEEN_CALLS)
            if not new_cursor or new_cursor == cursor:
                break
            cursor = new_cursor
        return docs


# ---------------------------------------------------------------- Aufbereitung

def normalize_urheber(raw: str) -> str:
    """Rohes Urheber-Feld auf einen kurzen Fraktions-/Organnamen eindampfen."""
    low = raw.lower()
    mapping = [
        ("cdu/csu", "CDU/CSU"),
        ("spd", "SPD"),
        ("bündnis 90", "Gruene"),
        ("grünen", "Gruene"),
        ("fdp", "FDP"),
        ("afd", "AfD"),
        ("bsw", "BSW"),
        ("linke", "Linke"),
        ("bundesregierung", "Bundesregierung"),
        ("bundesrat", "Bundesrat"),
    ]
    for needle, name in mapping:
        if needle in low:
            return name
    return raw.strip() or "unbekannt"


def rows_from_docs(docs, keyword: str):
    rows = []
    for d in docs:
        urheber_raw = d.get("urheber") or []
        if isinstance(urheber_raw, list):
            urheber_namen = []
            for u in urheber_raw:
                if isinstance(u, dict):
                    urheber_namen.append(u.get("titel") or u.get("bezeichnung") or "")
                else:
                    urheber_namen.append(str(u))
        else:
            urheber_namen = [str(urheber_raw)]
        urheber_join = "; ".join(n for n in urheber_namen if n)
        fraktion = normalize_urheber(urheber_join) if urheber_join else "unbekannt"
        fundstelle = d.get("fundstelle") or {}
        rows.append({
            "id": str(d.get("id", "")),
            "dokumentnummer": d.get("dokumentnummer", ""),
            "typ": d.get("drucksachetyp", ""),
            "datum": d.get("datum", ""),
            "titel": d.get("titel", ""),
            "urheber": urheber_join,
            "fraktion": fraktion,
            "wahlperiode": d.get("wahlperiode", ""),
            "pdf_url": fundstelle.get("pdf_url", ""),
            "keyword": keyword,
        })
    return rows


# ---------------------------------------------------------------- SQLite

DDL = """
CREATE TABLE IF NOT EXISTS drucksachen (
    id TEXT PRIMARY KEY,
    dokumentnummer TEXT, typ TEXT, datum TEXT, titel TEXT,
    urheber TEXT, fraktion TEXT, wahlperiode TEXT,
    pdf_url TEXT, keyword TEXT
);
"""


def upsert(rows) -> int:
    con = sqlite3.connect(DB_PATH)
    con.execute(DDL)
    n = 0
    for r in rows:
        cur = con.execute(
            """INSERT INTO drucksachen
               (id, dokumentnummer, typ, datum, titel, urheber, fraktion,
                wahlperiode, pdf_url, keyword)
               VALUES (:id, :dokumentnummer, :typ, :datum, :titel, :urheber,
                       :fraktion, :wahlperiode, :pdf_url, :keyword)
               ON CONFLICT(id) DO UPDATE SET
                 titel=excluded.titel, urheber=excluded.urheber,
                 fraktion=excluded.fraktion, datum=excluded.datum""",
            r)
        n += cur.rowcount
    con.commit()
    con.close()
    return n


def load_all():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    rows = [dict(r) for r in con.execute(
        "SELECT * FROM drucksachen ORDER BY datum DESC")]
    con.close()
    return rows


# ---------------------------------------------------------------- Excel

def write_excel(rows):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="1D3C6E")
    wb = Workbook()

    def style_header(ws, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A2"

    def autosize(ws):
        for col in ws.columns:
            width = max((len(str(c.value)) for c in col if c.value is not None),
                        default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = \
                min(width + 2, 60)

    # --- Drucksachen ---
    ws_d = wb.active
    ws_d.title = "Drucksachen"
    ws_d.append(["Datum", "Nummer", "Typ", "Fraktion/Organ", "Titel",
                 "Urheber (roh)", "WP", "PDF"])
    for r in rows:
        ws_d.append([r["datum"], r["dokumentnummer"], r["typ"], r["fraktion"],
                     r["titel"], r["urheber"], r["wahlperiode"], r["pdf_url"]])
    style_header(ws_d, 8)

    # --- Je Fraktion x Typ ---
    typen = sorted({r["typ"] for r in rows if r["typ"]})
    fraktionen = sorted({r["fraktion"] for r in rows})
    matrix = {}
    for r in rows:
        key = (r["fraktion"], r["typ"])
        matrix[key] = matrix.get(key, 0) + 1
    ws_f = wb.create_sheet("Je Fraktion")
    ws_f.append(["Fraktion/Organ"] + typen + ["Gesamt"])
    for f in fraktionen:
        line = [f] + [matrix.get((f, t), 0) for t in typen]
        line.append(sum(line[1:]))
        ws_f.append(line)
    style_header(ws_f, len(typen) + 2)

    # --- Zeitverlauf: Jahr x Fraktion ---
    jahre = sorted({r["datum"][:4] for r in rows if r["datum"]})
    zv = {}
    for r in rows:
        if not r["datum"]:
            continue
        key = (r["datum"][:4], r["fraktion"])
        zv[key] = zv.get(key, 0) + 1
    ws_z = wb.create_sheet("Zeitverlauf")
    ws_z.append(["Jahr"] + fraktionen)
    for j in jahre:
        ws_z.append([j] + [zv.get((j, f), 0) for f in fraktionen])
    style_header(ws_z, len(fraktionen) + 1)

    # --- Zusammenfassung ---
    ws_s = wb.create_sheet("Zusammenfassung", 0)
    daten = sorted({r["datum"] for r in rows if r["datum"]})
    info = [
        ("Bahnpolitik im Bundestag (DIP-API)", ""),
        ("Erstellt am", datetime.now().strftime("%d.%m.%Y %H:%M")),
        ("Zeitraum", f"{daten[0]} bis {daten[-1]}" if daten else "leer"),
        ("Drucksachen gesamt", len(rows)),
        ("Suchbegriffe", ", ".join(KEYWORDS)),
        ("Hinweis", "Zuordnung ueber Titelsuche, kein Anspruch auf "
                    "Vollstaendigkeit. Urheber = einbringende Fraktion "
                    "bzw. einbringendes Organ."),
    ]
    for r in info:
        ws_s.append(list(r))
    for c in ws_s["A"]:
        c.font = Font(name="Arial", bold=True)
    ws_s["A1"].font = Font(name="Arial", bold=True, size=14)

    for ws in wb.worksheets:
        autosize(ws)
    wb.save(XLSX_PATH)


# ---------------------------------------------------------------- Ablauf

def main() -> int:
    print("=" * 50)
    print(f" Bahnpolitik-Tracker (DIP-API, Version {SCRIPT_VERSION})")
    print("=" * 50)

    key = load_api_key()
    if not key:
        print(f"\nFEHLER: Kein API-Key gefunden (DIP_API_KEY oder {CRED_PATH}).")
        print("Aktuellen oeffentlichen Key holen unter:")
        print("https://dip.bundestag.de/über-dip/hilfe/api")
        return 2
    print(f"API-Key geladen: {len(key)} Zeichen")

    client = DipClient(key)
    gefunden = {}
    for kw in KEYWORDS:
        docs = client.drucksachen(kw)
        rows = rows_from_docs(docs, kw)
        neu = 0
        for r in rows:
            if r["id"] and r["id"] not in gefunden:
                gefunden[r["id"]] = r
                neu += 1
        print(f"[{kw}] {len(rows)} Treffer, davon {neu} neu in diesem Lauf")

    n = upsert(list(gefunden.values()))
    print(f"\n{len(gefunden)} eindeutige Drucksachen, {n} gespeichert/aktualisiert.")

    rows = load_all()
    if not rows:
        print("Kein Bestand vorhanden, keine Excel erzeugt.")
        return 1
    write_excel(rows)

    counts = {}
    for r in rows:
        counts[r["fraktion"]] = counts.get(r["fraktion"], 0) + 1
    print("\nBestand je Fraktion/Organ:")
    for f, c in sorted(counts.items(), key=lambda x: -x[1]):
        print(f"  {f}: {c}")
    print(f"\nExcel: {XLSX_PATH}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"\nUNERWARTETER FEHLER: {type(exc).__name__}: {exc}")
        sys.exit(1)
