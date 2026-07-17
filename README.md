# WAMAS Lift & Store – License Calculator

IC-Preis-Kalkulator für **WAMAS Lift & Store** Lizenzen.

Preisbasis: *IC Prices, valid for projects sold from 1st January 2024.*

## Lizenzmodell

| Bereich | Inhalt |
|---------|--------|
| Instances | Basic (€ 2.200) / Advanced (€ 3.700) – jeweils inkl. 1 Opening + 1 Admin Client |
| Add-ons | Security, Velocity, Tray Weight, RFID, Permission, Printing (€ 350); Put2Light (€ 500); External Storage / Picking Trolley (€ 1.000, Advanced) |
| Clients | Additional Opening (€ 600), Admin (€ 1.000), Mobile Terminal (€ 600, Advanced), 3rd Party VLM (€ 1.000) |
| Misc | Test Instance (€ 1.300), Upgrade Fee (€ 300/Jahr) |
| Mengenrabatt | 5–10 SLL → 5%, 11–19 SLL → 15%, ≥20 SLL → 25% |

## Start

```bash
cd lift-store-angebot
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python server.py
```

→ [http://127.0.0.1:8100](http://127.0.0.1:8100)

## API

| Methode | Pfad | Beschreibung |
|---------|------|--------------|
| GET | `/api/catalog` | IC Price List |
| POST | `/api/offers/calculate` | Kalkulation |
| POST | `/api/offers` | Speichern |
| GET | `/api/offers/{id}/excel` | Excel-Export |
