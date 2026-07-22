# Feldliste Word-Angebot (aktive Vorlage)

## Aktive Vorlage

**`Angebot Logimat DE.docx`**  
Pfad: `lift-store-angebot/docs/templates/Angebot Logimat DE.docx`

Der Word-Export befüllt **diese Vorlage** (docxtpl, XML-Autoescape).  
Die **Titelseite** enthält LOGIMAT-Bild, Logo, Kundendaten und SSI-Ansprechpartner.

Fallbacks im gleichen Ordner: `anhang_angebot_vorlage.docx`, `angebot_platzhalter.docx`.

Zum Testen: Vorlage anpassen → Server neu starten → Angebot → **Word**.

## Cover-Platzhalter (Titelseite)

| Platzhalter | Bedeutung |
|-------------|-----------|
| `{{titel}}` | Titel (z. B. Angebot WAMAS® Lift & Store) |
| `{{untertitel}}` | Gelber Banner-Text |
| `{{datum}}` | Angebotsdatum |
| `{{kunde}}` | Firmenname |
| `{{ansprechpartner}}` | Name Kunde |
| `{{strasse}}` | Strasse / Nr. |
| `{{plz_ort}}` | PLZ und Ort |
| `{{telefon}}` | Telefon Kunde |
| `{{fax}}` | Fax (optional) |
| `{{email}}` | E-Mail Kunde |
| `{{ssi_kontakt_1_*}}` / `{{ssi_kontakt_2_*}}` | SSI-Ansprechpartner (Name, Position, E-Mail, Telefon) |

Adresse aus dem Formular wird automatisch in `strasse` + `plz_ort` zerlegt, wenn sie dem Muster `…, 8000 Zürich` entspricht.

## Weitere Platzhalter (Inhalt)

`{{anrede}}`, `{{einleitung}}`, `{{konfiguration}}`, `{{leistungsumfang_text}}`, `{{optionen_text}}`,  
`{{lizenz_total_chf}}`, `{{it_total_chf}}`, `{{gesamt_chf}}`, `{{preis_hinweis}}`, `{{meta_zeile}}`,  
`{{erstellt_von}}`, Bedingungen/Signaturen wie bisher.

## Tipps

- Platzhalter als **einen** Text schreiben: `{{kunde}}`
- Cover-Bilder nicht in Word „freistellen“ und ersetzen, ohne Backup
- Backup der Originaldatei: `Angebot Logimat DE.backup.docx`

Maschinenlesbare Liste: `../../data/docx_field_map.json`
