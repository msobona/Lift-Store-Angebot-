# Feldliste Word-Angebot (Platzhalter)

Vorlage: `angebot_platzhalter.docx`  
Engine: `docxtpl` (Jinja2)  
Software-Anhang (unverändert angehängt): `docs/manuals/2.8/Anhang zu Software_v2.6.X.docx`

Vorlage erzeugen/aktualisieren:

```bash
cd lift-store-angebot
.venv/bin/python scripts/build_docx_template.py
```

## Einfache Felder

In Word als `{{feldname}}` einfügen.

| Platzhalter | Bedeutung | Beispiel |
|-------------|-----------|----------|
| `{{dokument_label}}` | Kennzeichnung | Angebot / Preisliste |
| `{{titel}}` | Titel | Angebot WAMAS® Lift & Store |
| `{{untertitel}}` | Untertitel | SSI SCHÄFER · … |
| `{{meta_zeile}}` | Nr. / Datum / Gültig / Rev. | Angebot ANG-… · 17.07.2026 · … |
| `{{angebotsnummer}}` | Angebotsnummer | ANG-20260717-ABC123 |
| `{{datum}}` | Datum | 17.07.2026 |
| `{{gueltig_bis}}` | Gültig bis | 31.07.2026 |
| `{{version_software}}` | Software-Version | 2.8 |
| `{{erstellt_von}}` | Erstellt von | Florian Brunner |
| `{{revision_von}}` | Revision von | ANG-… (leer wenn neu) |
| `{{kunde}}` | Firma | Muster AG |
| `{{projekt}}` | Projekt | LOGIMAT Halle 3 |
| `{{ansprechpartner}}` | Ansprechpartner | Max Beispiel |
| `{{email}}` | E-Mail | max@muster.ch |
| `{{adresse}}` | Adresse | Musterweg 1… |
| `{{konfiguration}}` | Konfiguration kurz | Advanced (1×) · 1 Gerät |
| `{{einleitung}}` | Einleitung (vorgerendert) | Die Softwarelösung… |
| `{{optionen_text}}` | Optionen (vorgerendert) | • RFID Login … |
| `{{bedingungen_text}}` | Bedingungen CH (vorgerendert) | 1.1 Geltung … |
| `{{terms_version}}` | Bedingungen-Version | 09.2022 |
| `{{preis_hinweis}}` | Preisnote | IC + 28% · Kurs 0.93 |
| `{{lizenz_total_chf}}` | Lizenzen CHF | CHF 9'951.74 |
| `{{lizenz_ic_eur}}` | IC EUR | EUR 8'360.00 |
| `{{marge_percent}}` | Marge % | 28 |
| `{{kurs_eur_chf}}` | Kurs | 0.93 |
| `{{it_total_chf}}` | IT Total CHF | CHF 23'215.00 |
| `{{it_work_chf}}` | IT Arbeit | CHF 20'580.00 |
| `{{it_travel_chf}}` | Reise | CHF 2'635.00 |
| `{{gesamt_chf}}` | Gesamttotal | CHF 33'166.74 |
| `{{schluss_text}}` | Schlusstext | Ihrem Auftrag… |
| `{{gruss}}` | Gruss | Freundliche Grüsse |
| `{{firma_ssi}}` | Firma | SSI SCHÄFER AG |
| `{{signatur_1_name}}` | Signatur 1 | Florian Brunner |
| `{{signatur_1_titel}}` | Titel 1 | Bereichsleiter… |
| `{{signatur_1_rolle}}` | Rolle 1 | Mitglied GL |
| `{{signatur_2_name}}` | Signatur 2 | Andrej Pulfer |
| `{{signatur_2_titel}}` | Titel 2 | Leiter Vertrieb… |

## Preistabelle `lines`

In einer Word-Tabelle drei Zeilen unter dem Header:

1. `{%tr for line in lines %}` (eigene Zeile)
2. Datenzeile mit `{{ line.pos }}`, `{{ line.name }}`, …
3. `{%tr endfor %}` (eigene Zeile)

```
{%tr for line in lines %}
{{ line.pos }} | {{ line.name }} | {{ line.qty }} | {{ line.unit_price }} | {{ line.hours }} | {{ line.amount }} {{ line.currency }}
{%tr endfor %}
```

Weitere Zeilenfelder: `section`, `sku`, `description`

## Vorgerenderte Blöcke

`optionen_text` und `bedingungen_text` werden in `docx_export.py` aus den Listen gebaut.  
So entfallen verschachtelte `{% for %}`-Schleifen in Word (stabiler bei Word-XML).

## So passt du die Vorlage an

1. `angebot_platzhalter.docx` in Word öffnen  
2. Layout ändern, Platzhalter `{{…}}` beibehalten oder verschieben  
3. Speichern unter dem gleichen Dateinamen  
4. Angebot erneut als Word exportieren  

Maschinenlesbare Liste: `../../data/docx_field_map.json`
