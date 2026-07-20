#!/usr/bin/env python3
"""Erzeugt die Anhang-Platzhalter-Vorlage aus dem originalen Software-Anhang.

Ablauf:
1. Kopie von docs/manuals/2.8/Anhang zu Software_v2.6.X.docx
2. Am Anfang des gleichen Dokuments Platzhalter für Angebot einfügen
3. Speichern als docs/templates/anhang_angebot_vorlage.docx

Die Vorlage kannst du in Word bearbeiten (Platzhalter verschieben/Layout SSI).
Beim Word-Export wird NUR dieses Dokument befüllt — kein Voranstellen
einer zweiten Datei mehr.
"""
from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = ROOT.parent

ANNEX_CANDIDATES = [
    REPO_ROOT / "docs" / "manuals" / "2.8" / "Anhang zu Software_v2.6.X.docx",
    ROOT / "docs" / "manuals" / "2.8" / "Anhang zu Software_v2.6.X.docx",
]
OUT = ROOT / "docs" / "templates" / "anhang_angebot_vorlage.docx"

NAVY = "0B1F3A"
TEAL = "007A8A"
GRAY = "4A5568"


def find_annex() -> Path:
    for path in ANNEX_CANDIDATES:
        if path.exists():
            return path
    raise FileNotFoundError(
        "Original-Anhang nicht gefunden. Erwartet: docs/manuals/2.8/Anhang zu Software_v2.6.X.docx"
    )


def _set_run_props(r_pr, *, size_pt: float = 11, bold: bool = False, color: str | None = None):
    if bold:
        b = OxmlElement("w:b")
        r_pr.append(b)
    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        r_pr.append(c)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    r_pr.append(sz)
    sz_cs = OxmlElement("w:szCs")
    sz_cs.set(qn("w:val"), str(int(size_pt * 2)))
    r_pr.append(sz_cs)


def make_paragraph(text: str, *, size_pt: float = 11, bold: bool = False, color: str | None = None) -> OxmlElement:
    """Ein Absatz mit genau einem Run — wichtig für stabile docxtpl-Platzhalter."""
    p = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:after"), "80")
    spacing.set(qn("w:before"), "0")
    p_pr.append(spacing)
    p.append(p_pr)

    r = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    _set_run_props(r_pr, size_pt=size_pt, bold=bold, color=color)
    r.append(r_pr)
    t = OxmlElement("w:t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def make_page_break() -> OxmlElement:
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def placeholder_blocks() -> list[OxmlElement]:
    """Angebotsblock mit Platzhaltern — im Anhang-Dokument, vor dem Originalinhalt."""
    blocks: list[OxmlElement] = []
    blocks.append(make_paragraph("{{dokument_label}}", size_pt=10, bold=True, color=TEAL))
    blocks.append(make_paragraph("{{titel}}", size_pt=18, bold=True, color=NAVY))
    blocks.append(make_paragraph("{{untertitel}}", size_pt=11, color=TEAL))
    blocks.append(make_paragraph("{{meta_zeile}}", size_pt=10, color=GRAY))
    blocks.append(make_paragraph(""))

    blocks.append(make_paragraph("Projekt & Kunde", size_pt=14, bold=True, color=NAVY))
    blocks.append(make_paragraph("Kunde: {{kunde}}", size_pt=10))
    blocks.append(make_paragraph("Projekt: {{projekt}}", size_pt=10))
    blocks.append(make_paragraph("Adresse: {{adresse}}", size_pt=10))
    blocks.append(make_paragraph("Ansprechpartner: {{ansprechpartner}}", size_pt=10))
    blocks.append(make_paragraph("E-Mail: {{email}}", size_pt=10))
    blocks.append(make_paragraph("Erstellt von: {{erstellt_von}}", size_pt=10))
    blocks.append(make_paragraph("Konfiguration: {{konfiguration}}", size_pt=10))
    blocks.append(make_paragraph(""))

    blocks.append(make_paragraph("Einleitung", size_pt=14, bold=True, color=NAVY))
    blocks.append(make_paragraph("{{einleitung}}", size_pt=10))
    blocks.append(make_paragraph(""))

    blocks.append(make_paragraph("Leistungsumfang", size_pt=14, bold=True, color=NAVY))
    blocks.append(make_paragraph("{{leistungsumfang_text}}", size_pt=10))
    blocks.append(make_paragraph(""))

    blocks.append(make_paragraph("Optionale Leistungen", size_pt=14, bold=True, color=NAVY))
    blocks.append(make_paragraph("{{optionen_text}}", size_pt=10))
    blocks.append(make_paragraph(""))

    blocks.append(make_paragraph("Preise", size_pt=14, bold=True, color=NAVY))
    blocks.append(make_paragraph("{{preis_hinweis}}", size_pt=9, color=GRAY))
    blocks.append(make_paragraph("Softwarelizenzen: {{lizenz_total_chf}}", size_pt=11, bold=True, color=NAVY))
    blocks.append(make_paragraph("IT-Aufwand / Installation: {{it_total_chf}}", size_pt=11, bold=True, color=NAVY))
    blocks.append(make_paragraph("Gesamttotal (exkl. MwSt.): {{gesamt_chf}}", size_pt=12, bold=True, color=NAVY))
    blocks.append(make_paragraph(""))

    blocks.append(make_paragraph("Kaufmännische Bedingungen (CH {{terms_version}})", size_pt=14, bold=True, color=NAVY))
    blocks.append(make_paragraph("{{bedingungen_text}}", size_pt=9))
    blocks.append(make_paragraph(""))
    blocks.append(make_paragraph("{{schluss_text}}", size_pt=10))
    blocks.append(make_paragraph("{{gruss}}", size_pt=10))
    blocks.append(make_paragraph("{{firma_ssi}}", size_pt=10, bold=True, color=NAVY))
    blocks.append(make_paragraph(""))
    blocks.append(make_paragraph("{{signatur_1_name}}  ·  {{signatur_1_titel}}  ·  {{signatur_1_rolle}}", size_pt=9))
    blocks.append(make_paragraph("{{signatur_2_name}}  ·  {{signatur_2_titel}}", size_pt=9))
    blocks.append(make_paragraph(""))
    blocks.append(
        make_paragraph(
            "— Ab hier folgt der Originalinhalt des Software-Anhangs —",
            size_pt=9,
            color=GRAY,
        )
    )
    # Seitenumbruch vor dem bisherigen Anhang-Inhalt (SSI-Layout bleibt unverändert)
    blocks.append(make_page_break())
    return blocks


def main() -> None:
    annex = find_annex()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(annex, OUT)

    doc = Document(str(OUT))
    body = doc.element.body
    # sectPr bleibt am Ende — Platzhalter davor, vor dem bestehenden Inhalt
    insert_at = 0
    for child in list(body):
        if child.tag.endswith("sectPr"):
            break
        # nach dem ersten sichtbaren Inhalt? Wir wollen VOR allem Inhalt
        break

    for idx, block in enumerate(placeholder_blocks()):
        body.insert(insert_at + idx, block)

    doc.save(OUT)
    print(f"Wrote {OUT}")
    print(f"Source annex: {annex}")
    print("Edit this file in Word: move/keep {{placeholders}}, keep SSI annex layout below.")


if __name__ == "__main__":
    main()
