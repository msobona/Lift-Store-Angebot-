# Feldliste Word-Angebot (aktive Vorlage)

## Aktive Vorlage

**`Angebot Logimat DE.docx`**  
Pfad: `lift-store-angebot/docs/templates/Angebot Logimat DE.docx`

Der Word-Export befüllt **nur diese eine Datei** (docxtpl). Kein separates Voranstellen mehr.

Zum Testen: Datei ersetzen/anpassen → Server neu starten → Angebot → **Word**.

## Platzhalter in der aktuellen Vorlage

| Platzhalter | Bedeutung |
|-------------|-----------|
| `{{titel}}` | Titel |
| `{{untertitel}}` | Untertitel |
| `{{kunde}}` | Firma |
| `{{ansprechpartner}}` | Ansprechpartner |
| `{{email}}` | E-Mail |
| `{{adresse}}` | Adresse |
| `{{erstellt_von}}` | Erstellt von |
| `{{konfiguration}}` | Konfiguration kurz |
| `{{einleitung}}` | Einleitung |
| `{{leistungsumfang_text}}` | Leistungsumfang |
| `{{optionen_text}}` | Optionen |

Weitere verfügbare Felder (einfach in Word einfügen):  
`{{angebotsnummer}}`, `{{datum}}`, `{{gueltig_bis}}`, `{{meta_zeile}}`, `{{projekt}}`, `{{lizenz_total_chf}}`, `{{it_total_chf}}`, `{{gesamt_chf}}`, `{{preis_hinweis}}`, `{{bedingungen_text}}`, `{{schluss_text}}`, `{{gruss}}`, `{{firma_ssi}}`, Signaturfelder …

## Tipps

- Platzhalter als **einen** Text schreiben: `{{kunde}}`  
- Nach starken Word-Formatierungen ggf. Platzhalter neu tippen (sonst zerlegt Word `{{` und `}}`)  
- Fallback-Vorlagen im gleichen Ordner: `anhang_angebot_vorlage.docx`, `angebot_platzhalter.docx`

Maschinenlesbare Liste: `../../data/docx_field_map.json`
