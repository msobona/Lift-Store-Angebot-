# WAMAS Lift & Store – SSI SCHÄFER Kalkulator

Module in einer App:

1. **Lizenzkalkulator** – IC License Price List (Basic/Advanced, Add-ons, Clients, SLL-Rabatt)
2. **IT-Aufwand** – analog `Kalkulation WAMAS Installation.xlsm`
3. **Angebot** – druckfertiges Dokument im Stil des Word-Anhangs „Anhang zu Software WAMAS Lift & Store“

Design angelehnt an SSI SCHÄFER Manuals / Anhang (Gelb `#FFED00`, Schwarz, Weiß, WAMAS-/SSI-Logo).

## Start

```bash
cd lift-store-angebot
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python server.py
```

→ http://127.0.0.1:8100

## Quellen

- `../docs/manuals/2.8/` Manuals, Anhang Word, Excel IT-Kalkulation
- `data/catalog.json` IC-Lizenzen
- `data/it_catalog.json` IT-Stammdaten/Aufwände
