"""Word-Export über Platzhalter-Vorlage (docxtpl) + originaler Software-Anhang."""

from __future__ import annotations

import io
import json
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxtpl import DocxTemplate

BASE_DIR = Path(__file__).resolve().parent
REPO_ROOT = BASE_DIR.parent
FIELD_MAP_FILE = BASE_DIR / "data" / "docx_field_map.json"

PLACEHOLDER_CANDIDATES = [
    BASE_DIR / "docs" / "templates" / "angebot_platzhalter.docx",
    REPO_ROOT / "lift-store-angebot" / "docs" / "templates" / "angebot_platzhalter.docx",
]
ANNEX_CANDIDATES = [
    REPO_ROOT / "docs" / "manuals" / "2.8" / "Anhang zu Software_v2.6.X.docx",
    BASE_DIR / "docs" / "manuals" / "2.8" / "Anhang zu Software_v2.6.X.docx",
    BASE_DIR / "docs" / "templates" / "Anhang zu Software_v2.6.X.docx",
]


def find_placeholder_template() -> Path:
    for path in PLACEHOLDER_CANDIDATES:
        if path.exists():
            return path
    raise FileNotFoundError(
        "Platzhalter-Vorlage nicht gefunden. Erwartet: docs/templates/angebot_platzhalter.docx "
        "(mit scripts/build_docx_template.py erzeugen)."
    )


def find_annex_template() -> Path:
    for path in ANNEX_CANDIDATES:
        if path.exists():
            return path
    raise FileNotFoundError(
        "Word-Anhang nicht gefunden. Erwartet unter docs/manuals/2.8/Anhang zu Software_v2.6.X.docx"
    )


def load_field_map() -> Dict[str, Any]:
    if FIELD_MAP_FILE.exists():
        return json.loads(FIELD_MAP_FILE.read_text(encoding="utf-8"))
    return {}


def _money(value: Any, currency: str = "CHF") -> str:
    try:
        amount = float(value or 0)
    except (TypeError, ValueError):
        amount = 0.0
    formatted = f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", "'")
    return f"{currency} {formatted}"


def _fmt_num(value: Any) -> str:
    if value is None or value == "":
        return ""
    try:
        num = float(value)
    except (TypeError, ValueError):
        return str(value)
    if abs(num - int(num)) < 1e-9:
        return str(int(num))
    return f"{num:.2f}".replace(".", ",")


def _join_paragraphs(parts: List[str]) -> str:
    return "\n\n".join(p.strip() for p in parts if p and str(p).strip())


def _render_optionen_text(options: List[Dict[str, Any]]) -> str:
    if not options:
        return "Keine optionalen Softwaremodule gewählt."
    blocks = []
    for opt in options:
        title = (opt.get("title") or "").strip()
        text = (opt.get("text") or "").strip()
        if title and text:
            blocks.append(f"• {title}\n{text}")
        elif title:
            blocks.append(f"• {title}")
        elif text:
            blocks.append(f"• {text}")
    return _join_paragraphs(blocks)


def _render_bedingungen_text(terms: Dict[str, Any]) -> str:
    blocks: List[str] = []
    for sec in terms.get("sections") or []:
        sid = (sec.get("id") or "").strip()
        title = (sec.get("title") or "").strip()
        head = f"{sid} {title}".strip()
        if head:
            blocks.append(head)
        for p in sec.get("paragraphs") or []:
            if p:
                blocks.append(str(p).strip())
        for b in sec.get("bullets") or []:
            if b:
                blocks.append(f"• {str(b).strip()}")
        for p in sec.get("paragraphsAfter") or []:
            if p:
                blocks.append(str(p).strip())
    return _join_paragraphs(blocks) if blocks else "—"


def build_template_context(offer: Dict[str, Any]) -> Dict[str, Any]:
    """Mappt Angebots-JSON auf die Platzhalter-Felder der Word-Vorlage."""
    meta = offer.get("meta") or {}
    customer = offer.get("customer") or {}
    content = offer.get("content") or {}
    cfg = content.get("configurationSummary") or {}
    summary = offer.get("priceSummary") or {}
    lic = summary.get("license") or {}
    it = summary.get("it") or {}
    terms = content.get("commercialTerms") or {}
    closing = terms.get("closing") or {}
    signatories = closing.get("signatories") or [{}, {}]
    sig1 = signatories[0] if len(signatories) > 0 else {}
    sig2 = signatories[1] if len(signatories) > 1 else {}

    cfg_bits = " · ".join(
        x
        for x in [
            cfg.get("instanceName"),
            f"{cfg['instanceCount']}×" if cfg.get("instanceCount") else None,
            f"{cfg.get('deviceCount')} Geräte" if cfg.get("deviceCount") else None,
            f"{cfg.get('zoneCount')} Zone(n)" if cfg.get("zoneCount") else None,
            f"{cfg.get('openingCount')} Öffnung(en)" if cfg.get("openingCount") else None,
        ]
        if x
    )

    lines: List[Dict[str, Any]] = []
    for line in offer.get("commercialLines") or []:
        lines.append(
            {
                "pos": line.get("pos", ""),
                "section": line.get("section", ""),
                "sku": line.get("sku", ""),
                "name": line.get("name", ""),
                "description": line.get("description", ""),
                "qty": _fmt_num(line.get("qty")),
                "unit_price": _fmt_num(line.get("unitPrice")),
                "hours": _fmt_num(line.get("hours")) if line.get("hours") not in (None, 0, 0.0) else "",
                "amount": _fmt_num(line.get("amount")),
                "currency": line.get("currency", ""),
            }
        )

    einleitung = _join_paragraphs(
        [
            content.get("intro") or "",
            content.get("introVariant") or "",
            content.get("recommendation") or "",
        ]
    )

    angebotsnummer = meta.get("offerNumber", "")
    datum = meta.get("documentDate", "")
    gueltig_bis = meta.get("validUntil", "")
    revision_von = (
        meta.get("revisionOf")
        or (offer.get("sources") or {}).get("basedOnOfferNumber")
        or ""
    )
    meta_parts = [
        f"Angebot {angebotsnummer}" if angebotsnummer else "Angebot",
        datum,
        f"Gültig bis {gueltig_bis}" if gueltig_bis else "",
        f"Rev. von {revision_von}" if revision_von else "",
    ]
    meta_zeile = "  ·  ".join(p for p in meta_parts if p)

    return {
        "dokument_label": content.get("documentLabel") or "Angebot / Preisliste",
        "titel": content.get("title") or "Angebot WAMAS® Lift & Store",
        "untertitel": content.get("subtitle") or "",
        "angebotsnummer": angebotsnummer,
        "datum": datum,
        "gueltig_bis": gueltig_bis,
        "meta_zeile": meta_zeile,
        "version_software": (offer.get("branding") or {}).get("version", "2.8"),
        "erstellt_von": meta.get("preparedBy", ""),
        "revision_von": revision_von,
        "kunde": customer.get("company", ""),
        "projekt": customer.get("projectName", ""),
        "ansprechpartner": customer.get("contact", ""),
        "email": customer.get("email", ""),
        "adresse": customer.get("address", ""),
        "konfiguration": cfg_bits,
        "instance_name": cfg.get("instanceName") or "",
        "instance_count": cfg.get("instanceCount") or "",
        "device_count": cfg.get("deviceCount") or "",
        "zone_count": cfg.get("zoneCount") or "",
        "opening_count": cfg.get("openingCount") or "",
        "order_handling": "Ja" if cfg.get("hasOrderHandling") else "Nein",
        "einleitung": einleitung,
        "optionen_text": _render_optionen_text(content.get("selectedOptions") or []),
        "bedingungen_text": _render_bedingungen_text(terms),
        "preis_hinweis": summary.get("note") or "",
        "lizenz_total_chf": _money(lic.get("total"), lic.get("currency") or "CHF") if lic else "—",
        "lizenz_ic_eur": _money(lic.get("icNet"), "EUR") if lic and lic.get("icNet") is not None else "—",
        "marge_percent": lic.get("marginPercent", 28) if lic else 28,
        "kurs_eur_chf": lic.get("eurToChfRate", 0.93) if lic else 0.93,
        "it_total_chf": _money(it.get("total"), it.get("currency") or "CHF") if it else "—",
        "it_work_chf": _money(it.get("workAmount"), "CHF") if it else "—",
        "it_travel_chf": _money(it.get("travelAmount"), "CHF") if it else "—",
        "gesamt_chf": _money(summary.get("grandTotalChf"), "CHF")
        if summary.get("grandTotalChf") is not None
        else "—",
        "lines": lines,
        "terms_version": terms.get("version", "09.2022"),
        "terms_validity_days": terms.get("validityDays", 14),
        "schluss_text": closing.get("text")
        or "Ihrem Auftrag, dem wir unsere volle Aufmerksamkeit schenken, sehen wir gerne entgegen.",
        "gruss": closing.get("greeting") or "Freundliche Grüsse",
        "firma_ssi": closing.get("company") or "SSI SCHÄFER AG",
        "signatur_1_name": sig1.get("name", ""),
        "signatur_1_titel": sig1.get("title", ""),
        "signatur_1_rolle": sig1.get("role", ""),
        "signatur_2_name": sig2.get("name", ""),
        "signatur_2_titel": sig2.get("title", ""),
    }


def _insert_filled_into_annex(annex_doc: Document, filled_doc: Document) -> None:
    body = annex_doc.element.body
    break_p = OxmlElement("w:p")
    break_r = OxmlElement("w:r")
    break_br = OxmlElement("w:br")
    break_br.set(qn("w:type"), "page")
    break_r.append(break_br)
    break_p.append(break_r)

    front_children = [c for c in filled_doc.element.body if not c.tag.endswith("sectPr")]
    body.insert(0, break_p)
    for child in reversed(front_children):
        body.insert(0, deepcopy(child))


def build_offer_docx(offer: Dict[str, Any]) -> bytes:
    """Befüllt Platzhalter-Vorlage und kombiniert sie mit dem Software-Anhang."""
    if offer.get("kind") != "offer_document":
        raise ValueError("Word-Export ist nur für Gesamtangebote (offer_document) verfügbar.")

    template_path = find_placeholder_template()
    annex_path = find_annex_template()
    context = build_template_context(offer)

    tpl = DocxTemplate(str(template_path))
    tpl.render(context)

    filled_buf = io.BytesIO()
    tpl.save(filled_buf)
    filled_buf.seek(0)
    filled_doc = Document(filled_buf)

    annex_doc = Document(str(annex_path))
    _insert_filled_into_annex(annex_doc, filled_doc)

    out = io.BytesIO()
    annex_doc.save(out)
    return out.getvalue()
