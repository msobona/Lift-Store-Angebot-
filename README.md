# WAMAS Lift & Store – Angebot Generator

Web-Anwendung zur Erstellung von Software-Angeboten für **WAMAS Lift & Store** (SSI SCHÄFER).

## Funktionen

- Konfiguration von Paketen (Starter / Professional / Enterprise)
- Lift-Lizenzen, Module und Services
- Live-Preisberechnung inkl. Rabatt und MwSt.
- Speichern und Archivieren von Angeboten
- Excel-Export und Druck-/PDF-Vorschau

> Die hinterlegten Preise sind **Richtwerte zu Demonstrationszwecken** und kein verbindliches SSI-SCHÄFER-Angebot.

## Start

```bash
cd lift-store-angebot
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python server.py
```

Anschließend im Browser öffnen: [http://127.0.0.1:8100](http://127.0.0.1:8100)

## Projektstruktur

```
lift-store-angebot/
├── server.py           # FastAPI Backend
├── requirements.txt
├── data/
│   ├── catalog.json    # Pakete, Module, Preise
│   └── offers/         # Gespeicherte Angebote (JSON)
└── static/
    ├── index.html
    ├── styles.css
    └── app.js
```

## API (Auswahl)

| Methode | Pfad | Beschreibung |
|---------|------|--------------|
| GET | `/api/catalog` | Produktkatalog |
| POST | `/api/offers/calculate` | Angebot berechnen |
| POST | `/api/offers` | Angebot speichern |
| GET | `/api/offers` | Angebotsarchiv |
| GET | `/api/offers/{id}/excel` | Excel-Export |

## Hinweise

- Gespeicherte Angebote liegen lokal unter `data/offers/`.
- Preise und Module können in `data/catalog.json` angepasst werden.
