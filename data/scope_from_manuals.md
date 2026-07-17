# WAMAS Lift & Store 2.8 – Funktionsumfang für Offerten

Quellen (Repo `docs/manuals/2.8/`):

- User Manual V01.0 enUS
- Administrator Manual V01.0 enUS
- Interface Manual WLS–ERP V01.0 enUS

> System Requirements Manual (`req.system`) lag im Repo noch nicht vor und ist hier nicht ausgewertet.

## Clients

| Client | Kernnutzen laut Manual |
|--------|------------------------|
| Touch Client | Login (auch RFID), Tray Operations, Store/Pick/Inventory/Consolidation |
| Desktop Client | Stammdaten, Aufträge, Administration, Monitoring, Schnittstellen |
| Mobile Terminal | Lagerprozesse mit Scanner/Touch/Keypad |

## Kernprozesse

- Manuell: Tray request/return, store, pick, consolidate, inventory, transfer between access openings
- Auftragsbasiert: Storage / Picking / Inventory / Replenishment Demands (create → release → activate → process → finalize)
- Administration: Split/Suspend/Cancel/Takeover von Demands

## Stammdaten & Lagerstruktur

- Items, Packaging Versions, Aliases, Images, Quantity Units
- Tray Templates, Box Types, Item-to-Box Assignments
- Stock-Identifying Attributes: Batch, BBD, Production Date
- Zones, Destination Locations, External Storage Locations, Replenishment
- Storage Reservation Strategies / ABC

## Integration

- LOGIMAT-Optionen: LogiPointer, LogiTilt, LogiDrive, LogiLight
- Scales an Access Openings
- Printer/Document Management (Labels, Delivery Notes)
- Excel Import (Items, Demands)
- MFS / SOC Communication
- System Partner Interfaces:
  - WMS → ERP
  - WMS → WCS
- ERP Transport: File / Socket / Table / JMS / REST
- ERP Formate: CSV / Fixed-Length / XML / JSON

## Lizenzierung

- Digital signierte License File je System
- Import/View über Licensing Dialogs
- Optional: WAMAS Peak Period License
- License Report für User-/Lizenzverbrauch

## Typische Offerten-Bausteine

1. Basispaket (Starter / Professional / Enterprise)
2. Anzahl Lift-/Geräte-Lizenzen
3. Optionale Module (Mobile, SIA, ERP, WCS, MFS, …)
4. Implementierung & Schulung
5. Wartung / Premium Support
