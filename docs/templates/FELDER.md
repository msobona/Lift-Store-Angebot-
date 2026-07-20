# Feldliste Word-Angebot (Platzhalter im Software-Anhang)

## So funktioniert der Export

**Ein Dokument:** `anhang_angebot_vorlage.docx`  
= Original-Software-Anhang (SSI-Design, Bilder, Funktionen) **plus** Platzhalter für Angebotsdaten.

Die App befüllt **nur dieses eine Dokument**. Es wird **nichts** mehr vorangestellt.

| Datei | Bedeutung |
|-------|-----------|
| `anhang_angebot_vorlage.docx` | Bearbeitbare Vorlage (Platzhalter + SSI-Anhang) |
| `../../../../docs/manuals/2.8/Anhang zu Software_v2.6.X.docx` | Original-Quelle (unverändert) |

## Vorlage erzeugen / zurücksetzen

```bash
cd lift-store-angebot
python scripts/build_annex_template.py
```

Das kopiert den Original-Anhang und setzt die Platzhalter-Blöcke am Anfang wieder ein.  
**Achtung:** Eigene Layout-Änderungen in der Vorlage gehen dabei verloren — vorher sichern.

## Vorlage in Word bearbeiten

1. `docs/templates/anhang_angebot_vorlage.docx` öffnen  
2. Platzhalter wie `{{kunde}}`, `{{gesamt_chf}}` verschieben oder formatieren  
3. SSI-Anhang-Teil unten unverändert lassen (oder anpassen)  
4. Unter gleichem Namen speichern  
5. In der App: Angebot → Word exportieren  

**Wichtig:** Jeden Platzhalter als **einen** durchgehenden Text schreiben (`{{kunde}}`), nicht Buchstabe für Buchstabe formatieren.

## Einfache Felder

| Platzhalter | Bedeutung |
|-------------|-----------|
| `{{dokument_label}}` | Kennzeichnung |
| `{{titel}}` | Titel |
| `{{untertitel}}` | Untertitel |
| `{{meta_zeile}}` | Nr. / Datum / Gültig / Rev. |
| `{{angebotsnummer}}` | Angebotsnummer |
| `{{datum}}` | Datum |
| `{{gueltig_bis}}` | Gültig bis |
| `{{version_software}}` | Software-Version |
| `{{erstellt_von}}` | Erstellt von |
| `{{revision_von}}` | Revision von |
| `{{kunde}}` | Firma |
| `{{projekt}}` | Projekt |
| `{{ansprechpartner}}` | Ansprechpartner |
| `{{email}}` | E-Mail |
| `{{adresse}}` | Adresse |
| `{{konfiguration}}` | Konfiguration kurz |
| `{{einleitung}}` | Einleitung |
| `{{leistungsumfang_text}}` | Leistungsumfang (Namen + Beschreibungen, Totals) |
| `{{optionen_text}}` | Optionen |
| `{{bedingungen_text}}` | Kaufmännische Bedingungen CH |
| `{{terms_version}}` | Bedingungen-Version |
| `{{preis_hinweis}}` | Preisnote |
| `{{lizenz_total_chf}}` | Lizenzen Total |
| `{{it_total_chf}}` | IT Total |
| `{{gesamt_chf}}` | Gesamttotal |
| `{{schluss_text}}` | Schlusstext |
| `{{gruss}}` | Gruss |
| `{{firma_ssi}}` | Firma SSI |
| `{{signatur_1_name}}` / `_titel` / `_rolle` | Signatur 1 |
| `{{signatur_2_name}}` / `_titel` | Signatur 2 |

Maschinenlesbare Liste: `../../data/docx_field_map.json`
