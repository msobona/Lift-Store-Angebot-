# Feldliste Word-Angebot (aktive Vorlage)

## Aktive Vorlage

**`anhang_angebot_vorlage.docx`**  
Pfad: `lift-store-angebot/docs/templates/anhang_angebot_vorlage.docx`

Der Word-Export befüllt **diese Vorlage** (docxtpl, mit XML-Autoescape).  
Sie enthält Kundenkopf, Leistungsumfang, Preise, kaufmännische Bedingungen und Signaturen.

`Angebot Logimat DE.docx` bleibt als Design-Referenz im Ordner, wird aber **nicht** als Default genutzt (unvollständig: keine Preise/Bedingungen, Alttexte aus Maschinenangebot).

Zum Testen: Vorlage ersetzen/anpassen → Server neu starten → Angebot → **Word**.

## Platzhalter in der aktiven Vorlage

| Platzhalter | Bedeutung |
|-------------|-----------|
| `{{dokument_label}}` | Dokumenttyp |
| `{{titel}}` | Titel |
| `{{untertitel}}` | Untertitel |
| `{{meta_zeile}}` | Nummer · Datum · Gültigkeit |
| `{{kunde}}` | Firma |
| `{{projekt}}` | Projekt |
| `{{adresse}}` | Adresse |
| `{{ansprechpartner}}` | Ansprechpartner |
| `{{email}}` | E-Mail |
| `{{telefon}}` | Telefon (falls in Vorlage eingefügt) |
| `{{erstellt_von}}` | Erstellt von |
| `{{konfiguration}}` | Konfiguration kurz |
| `{{einleitung}}` | Einleitung |
| `{{leistungsumfang_text}}` | Leistungsumfang (ohne Positionspreise) |
| `{{optionen_text}}` | Optionen |
| `{{preis_hinweis}}` | Preis-/Rabatthinweis |
| `{{lizenz_total_chf}}` | Lizenz-Total CHF |
| `{{it_total_chf}}` | IT-Total CHF |
| `{{gesamt_chf}}` | Gesamttotal CHF |
| `{{bedingungen_text}}` | Kaufmännische Bedingungen |
| `{{schluss_text}}` / `{{gruss}}` / `{{firma_ssi}}` | Abschluss |
| Signaturfelder | `{{signatur_1_name}}`, `{{signatur_1_titel}}`, `{{signatur_1_rolle}}`, … |

Weitere optionale Felder: `{{angebotsnummer}}`, `{{datum}}`, `{{gueltig_bis}}`, `{{anrede}}`, `{{terms_version}}`.

## Tipps

- Platzhalter als **einen** Text schreiben: `{{kunde}}`
- Nach starken Word-Formatierungen ggf. Platzhalter neu tippen (sonst zerlegt Word `{{` und `}}`)
- Fallback: `angebot_platzhalter.docx`

Maschinenlesbare Liste: `../../data/docx_field_map.json`
