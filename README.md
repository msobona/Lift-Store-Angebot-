# WAMAS Lift & Store Angebot

Eigenständiges Angebot- und Kalkulationstool für **SSI SCHÄFER WAMAS Lift & Store**.

Dieses Projekt ist **bewusst getrennt** von `WAMAS_WebApp_4.0`.  
Eigene Versionierung, eigene Releases, eigene Laufzeit (Port **8100**).

## Module

1. **Lizenzkalkulator** – IC License Price List (Basic/Advanced, Add-ons, Clients, SLL-Rabatt)
2. **IT-Aufwand** – analog `Kalkulation WAMAS Installation.xlsm` (inkl. Reisekosten-Paket Basic 5/5 · Advanced 7/7)
3. **Angebot** – druckfertiges Dokument (HTML-Vorschau, Word/PDF-Export, Archiv)

## Voraussetzungen

- Python 3.10+ (empfohlen 3.11/3.12)
- Windows oder Linux/macOS

## Start

```bash
python3 -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
pip install -r requirements.txt
python server.py
```

Oder:

```bash
./start.sh          # Linux/macOS
start.bat           # Windows
```

→ http://127.0.0.1:8100

## Version

Siehe Datei [`VERSION`](./VERSION) (aktuell **0.1.0**).  
WebApp-Versionen und dieses Tool nicht vermischen.

## Wichtige Ordner

| Pfad | Inhalt |
|------|--------|
| `data/` | Kataloge, SSI-Kontakte, Export-Einstellungen, lokale Angebote |
| `docs/templates/` | Word-Vorlagen (Logimat, Platzhalter) |
| `static/` | UI (HTML/CSS/JS) + Assets |
| `server.py` | FastAPI-Backend |
| `docx_*.py` / `pdf_export.py` | Dokument-Export |

## Herkunft

Entwickelt zuerst unter `WAMAS_WebApp_4.0/lift-store-angebot/`, danach als eigenes Repository ausgegliedert (Git-Historie der Angebot-Dateien übernommen).

## Lizenz / Intern

Internes SSI-SCHÄFER-Werkzeug – nicht öffentlich freigeben, sofern nicht freigegeben.
