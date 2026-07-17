"""Word-Export: Angebotsseiten + originaler Software-Anhang (DOCX-Vorlage)."""

from __future__ import annotations

import io
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE_DIR = Path(__file__).resolve().parent
REPO_ROOT = BASE_DIR.parent
ANNEX_CANDIDATES = [
    REPO_ROOT / "docs" / "manuals" / "2.8" / "Anhang zu Software_v2.6.X.docx",
    BASE_DIR / "docs" / "manuals" / "2.8" / "Anhang zu Software_v2.6.X.docx",
    BASE_DIR / "docs" / "templates" / "Anhang zu Software_v2.6.X.docx",
]

SSI_YELLOW = "FFED00"
SSI_BLACK = "111111"


def find_annex_template() -> Path:
    for path in ANNEX_CANDIDATES:
        if path.exists():
            return path
    raise FileNotFoundError(
        "Word-Anhang nicht gefunden. Erwartet unter docs/manuals/2.8/Anhang zu Software_v2.6.X.docx"
    )


def _set_run_font(run, *, bold: bool = False, size_pt: float = 10, color: Optional[str] = None) -> None:
    run.bold = bold
    run.font.size = Pt(size_pt)
    run.font.name = "Calibri"
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), "Calibri")
    rFonts.set(qn("w:hAnsi"), "Calibri")
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def _shade_paragraph(paragraph, fill: str) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def _add_heading_band(doc: Document, text: str, size_pt: float = 16) -> None:
    p = doc.add_paragraph()
    _shade_paragraph(p, SSI_YELLOW)
    run = p.add_run(text)
    _set_run_font(run, bold=True, size_pt=size_pt, color=SSI_BLACK)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)


def _add_label_value(doc: Document, label: str, value: str) -> None:
    p = doc.add_paragraph()
    r1 = p.add_run(f"{label}: ")
    _set_run_font(r1, bold=True, size_pt=10)
    r2 = p.add_run(value or "—")
    _set_run_font(r2, size_pt=10)
    p.paragraph_format.space_after = Pt(2)


def _money(value: Any, currency: str = "CHF") -> str:
    try:
        amount = float(value or 0)
    except (TypeError, ValueError):
        amount = 0.0
    formatted = f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", "'")
    return f"{currency} {formatted}"


def _add_table(doc: Document, headers: List[str], rows: List[List[Any]]) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(str(header))
        _set_run_font(run, bold=True, size_pt=8)
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), SSI_YELLOW)
        shading.set(qn("w:val"), "clear")
        cell._tc.get_or_add_tcPr().append(shading)
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run("" if value is None else str(value))
            _set_run_font(run, size_pt=8)


def build_offer_front_matter(offer: Dict[str, Any]) -> Document:
    """Erzeugt die kundenspezifischen Angebotsseiten (Preise + kaufm. Bedingungen)."""
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    meta = offer.get("meta") or {}
    customer = offer.get("customer") or {}
    content = offer.get("content") or {}
    cfg = content.get("configurationSummary") or {}
    summary = offer.get("priceSummary") or {}
    lic = summary.get("license") or {}
    it = summary.get("it") or {}

    title = content.get("title") or "Angebot WAMAS® Lift & Store"
    _add_heading_band(doc, "SSI SCHÄFER AG · Angebot / Preisliste", 12)
    h = doc.add_paragraph()
    run = h.add_run(title)
    _set_run_font(run, bold=True, size_pt=18)
    h.paragraph_format.space_after = Pt(4)

    sub = doc.add_paragraph()
    run = sub.add_run(content.get("subtitle") or "WAMAS Lift & Store")
    _set_run_font(run, size_pt=11)
    sub.paragraph_format.space_after = Pt(10)

    _add_label_value(doc, "Angebotsnummer", meta.get("offerNumber", ""))
    _add_label_value(doc, "Datum", meta.get("documentDate", ""))
    _add_label_value(doc, "Gültig bis", meta.get("validUntil", ""))
    if meta.get("revisionOf"):
        _add_label_value(doc, "Revision von", meta.get("revisionOf", ""))
    _add_label_value(doc, "Kunde", customer.get("company", ""))
    _add_label_value(doc, "Projekt", customer.get("projectName", ""))
    _add_label_value(doc, "Ansprechpartner", customer.get("contact", ""))
    _add_label_value(doc, "Erstellt von", meta.get("preparedBy", ""))
    cfg_bits = " · ".join(
        x
        for x in [
            cfg.get("instanceName"),
            f"{cfg['instanceCount']}×" if cfg.get("instanceCount") else None,
            f"{cfg.get('deviceCount')} Geräte" if cfg.get("deviceCount") else None,
            f"{cfg.get('openingCount')} Öffnungen" if cfg.get("openingCount") else None,
        ]
        if x
    )
    _add_label_value(doc, "Konfiguration", cfg_bits)

    doc.add_paragraph()
    _add_heading_band(doc, "1. Preisliste – alle Positionen", 13)

    note = doc.add_paragraph()
    run = note.add_run(
        summary.get("note")
        or "Lizenzen: IC + 28% Marge, EUR→CHF 0.93. IT in CHF. Alle Beträge exkl. MwSt."
    )
    _set_run_font(run, size_pt=9)

    rows = []
    for line in offer.get("commercialLines") or []:
        rows.append(
            [
                line.get("pos"),
                (line.get("section") or "")[:28],
                line.get("name"),
                line.get("qty"),
                line.get("unitPrice"),
                line.get("hours") if line.get("hours") not in (None, 0, 0.0) else "",
                line.get("amount"),
                line.get("currency"),
            ]
        )
    _add_table(
        doc,
        ["Pos", "Bereich", "Bezeichnung", "Menge", "EP", "Std.", "Betrag", "Währ."],
        rows,
    )

    doc.add_paragraph()
    if lic.get("total") is not None:
        _add_label_value(doc, "Softwarelizenzen Verkauf", _money(lic.get("total"), lic.get("currency") or "CHF"))
        _add_label_value(
            doc,
            "davon IC / Marge / Kurs",
            f"IC {_money(lic.get('icNet'), 'EUR')} + {lic.get('marginPercent', 28)}% · Kurs {lic.get('eurToChfRate', 0.93)}",
        )
    if it.get("total") is not None:
        _add_label_value(doc, "IT-Aufwand inkl. Reise", _money(it.get("total"), it.get("currency") or "CHF"))
    if summary.get("grandTotalChf") is not None:
        p = doc.add_paragraph()
        _shade_paragraph(p, SSI_BLACK)
        run = p.add_run(f"Gesamttotal exkl. MwSt: {_money(summary.get('grandTotalChf'), 'CHF')}")
        _set_run_font(run, bold=True, size_pt=12, color="FFFFFF")

    # Gewählte Optionen kurz
    opts = content.get("selectedOptions") or []
    if opts:
        doc.add_paragraph()
        _add_heading_band(doc, "Gewählte Software-Optionen", 12)
        for opt in opts:
            p = doc.add_paragraph(style="List Bullet")
            run = p.add_run(f"{opt.get('title', '')}: {opt.get('text', '')}")
            _set_run_font(run, size_pt=9)

    # Kaufmännische Bedingungen
    terms = content.get("commercialTerms") or {}
    if terms.get("sections"):
        doc.add_page_break()
        _add_heading_band(doc, terms.get("title") or "Kaufmännische Bedingungen", 13)
        ver = doc.add_paragraph()
        run = ver.add_run(f"Version {terms.get('version', '')} · Angebotsgültigkeit {terms.get('validityDays', 14)} Tage")
        _set_run_font(run, size_pt=9)

        for sec in terms["sections"]:
            h = doc.add_paragraph()
            run = h.add_run(f"{sec.get('id', '')} {sec.get('title', '')}")
            _set_run_font(run, bold=True, size_pt=11)
            h.paragraph_format.space_before = Pt(8)
            for para in sec.get("paragraphs") or []:
                p = doc.add_paragraph()
                run = p.add_run(para)
                _set_run_font(run, size_pt=8)
                p.paragraph_format.space_after = Pt(3)
            for bullet in sec.get("bullets") or []:
                p = doc.add_paragraph(style="List Bullet")
                run = p.add_run(bullet)
                _set_run_font(run, size_pt=8)
            for para in sec.get("paragraphsAfter") or []:
                p = doc.add_paragraph()
                run = p.add_run(para)
                _set_run_font(run, size_pt=8)
            for sub in sec.get("subsections") or []:
                sh = doc.add_paragraph()
                run = sh.add_run(sub.get("title", ""))
                _set_run_font(run, bold=True, size_pt=9)
                for para in sub.get("paragraphs") or []:
                    p = doc.add_paragraph()
                    run = p.add_run(para)
                    _set_run_font(run, size_pt=8)
            if sec.get("table"):
                _add_table(doc, sec["table"]["headers"], sec["table"]["rows"])

        closing = terms.get("closing") or {}
        doc.add_paragraph()
        p = doc.add_paragraph()
        run = p.add_run(closing.get("text") or "")
        _set_run_font(run, size_pt=10)
        p = doc.add_paragraph()
        run = p.add_run(closing.get("greeting") or "Freundliche Grüsse")
        _set_run_font(run, bold=True, size_pt=10)
        p = doc.add_paragraph()
        run = p.add_run(closing.get("company") or "SSI SCHÄFER AG")
        _set_run_font(run, bold=True, size_pt=10)
        for sig in closing.get("signatories") or []:
            p = doc.add_paragraph()
            run = p.add_run(sig.get("name", ""))
            _set_run_font(run, bold=True, size_pt=9)
            if sig.get("title"):
                p2 = doc.add_paragraph()
                run = p2.add_run(sig["title"])
                _set_run_font(run, size_pt=8)
            if sig.get("role"):
                p3 = doc.add_paragraph()
                run = p3.add_run(sig["role"])
                _set_run_font(run, size_pt=8)

    # Trenner vor Software-Anhang
    doc.add_page_break()
    band = doc.add_paragraph()
    _shade_paragraph(band, SSI_YELLOW)
    run = band.add_run("Anhang zur Software · WAMAS® Lift & Store")
    _set_run_font(run, bold=True, size_pt=14)
    note = doc.add_paragraph()
    run = note.add_run(
        "Nachfolgend der Software-Anhang in der Originalvorlage "
        "„Anhang zu Software_v2.6.X.docx“ (Format, Inhalte und Abbildungen)."
    )
    _set_run_font(run, size_pt=9)
    return doc


def _insert_front_matter_into_annex(annex_doc: Document, front_doc: Document) -> None:
    """Fügt Angebotsseiten am Anfang der originalen Anhang-DOCX ein (Medien bleiben erhalten)."""
    body = annex_doc.element.body
    # Seitenumbruch vor dem bisherigen Anhang-Inhalt
    break_p = OxmlElement("w:p")
    break_r = OxmlElement("w:r")
    break_br = OxmlElement("w:br")
    break_br.set(qn("w:type"), "page")
    break_r.append(break_br)
    break_p.append(break_r)

    front_children = [
        child
        for child in front_doc.element.body
        if not child.tag.endswith("sectPr")
    ]
    # In umgekehrter Reihenfolge am Anfang einfügen
    body.insert(0, break_p)
    for child in reversed(front_children):
        body.insert(0, deepcopy(child))


def build_offer_docx(offer: Dict[str, Any]) -> bytes:
    """Baut ein DOCX: Angebots-/Preisseiten + originaler Word-Software-Anhang."""
    if offer.get("kind") != "offer_document":
        raise ValueError("Word-Export ist nur für Gesamtangebote (offer_document) verfügbar.")

    annex_path = find_annex_template()
    annex_doc = Document(str(annex_path))
    front_doc = build_offer_front_matter(offer)
    _insert_front_matter_into_annex(annex_doc, front_doc)

    stream = io.BytesIO()
    annex_doc.save(stream)
    return stream.getvalue()
