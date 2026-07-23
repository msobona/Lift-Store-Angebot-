"""Word-Export: Software-Anhang mit Platzhaltern befüllen (ein Dokument)."""

from __future__ import annotations

import io
import json
import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional
from xml.sax.saxutils import escape as xml_escape

from docxtpl import DocxTemplate

BASE_DIR = Path(__file__).resolve().parent
FIELD_MAP_FILE = BASE_DIR / "data" / "docx_field_map.json"
SSI_CONTACTS_FILE = BASE_DIR / "data" / "ssi_contacts.json"

# Word-Inhaltssteuerelemente (ComboBox / Text) in Angebot Logimat DE.docx
_SDT_TEXT_MAP = {
    "Name (Innendienst)": "ssi_kontakt_1_name",
    "Text Innendienst": "ssi_kontakt_1_details",
    "Name (Aussendienst)": "ssi_kontakt_2_name",
    "Text Aussendienst": "ssi_kontakt_2_details",
    "Unterschrift links": "signatur_1_name",
    "Funktion links": "signatur_1_details",
    "Unterschrift rechts": "signatur_2_name",
    "Funktion rechts": "signatur_2_details",
}

# Primär: Logimat-Vorlage (Titelseite mit Bild + Platzhalter).
# Fallbacks: vollständige Anhang-Vorlage / schlanke Platzhalter-Vorlage.
PLACEHOLDER_CANDIDATES = [
    BASE_DIR / "docs" / "templates" / "Angebot Logimat DE.docx",
    BASE_DIR / "docs" / "templates" / "anhang_angebot_vorlage.docx",
    BASE_DIR / "docs" / "templates" / "angebot_platzhalter.docx",
]

INSTANCE_DISPLAY_NAMES = {
    "basic": "Basis-Instanz",
    "advanced": "Advanced-Instanz",
    "Basic Instance": "Basis-Instanz",
    "Advanced Instance": "Advanced-Instanz",
}


def find_placeholder_template() -> Path:
    for path in PLACEHOLDER_CANDIDATES:
        if path.exists():
            return path
    raise FileNotFoundError(
        "Word-Vorlage nicht gefunden. Erwartet z.B.:\n"
        "  docs/templates/anhang_angebot_vorlage.docx"
    )


def ensure_annex_template() -> Path:
    """Liefert die aktive Word-Vorlage (Anhang / Fallback / Logimat)."""
    return find_placeholder_template()


def load_field_map() -> Dict[str, Any]:
    if FIELD_MAP_FILE.exists():
        return json.loads(FIELD_MAP_FILE.read_text(encoding="utf-8"))
    return {}


def _load_ssi_contacts_catalog() -> List[Dict[str, Any]]:
    if not SSI_CONTACTS_FILE.exists():
        return []
    try:
        data = json.loads(SSI_CONTACTS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return list(data.get("contacts") or [])


def _lookup_ssi_contact(
    *, contact_id: str = "", name: str = ""
) -> Dict[str, Any]:
    cid = (contact_id or "").strip()
    needle = (name or "").strip().lower()
    for row in _load_ssi_contacts_catalog():
        if cid and row.get("id") == cid:
            return dict(row)
    if needle:
        for row in _load_ssi_contacts_catalog():
            if str(row.get("name") or "").strip().lower() == needle:
                return dict(row)
    return {}


def _money(value: Any, currency: str = "CHF") -> str:
    try:
        amount = float(value or 0)
    except (TypeError, ValueError):
        amount = 0.0
    if str(currency).upper() == "CHF":
        amount = float(round(amount))
        formatted = f"{amount:,.0f}".replace(",", "'")
    else:
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


def _pluralize(count: Any, singular: str, plural: str) -> Optional[str]:
    try:
        n = int(count)
    except (TypeError, ValueError):
        return None
    if n <= 0:
        return None
    return f"{n} {singular if n == 1 else plural}"


def _display_instance_name(name: Any, instance_id: Any = None) -> str:
    if instance_id and str(instance_id) in INSTANCE_DISPLAY_NAMES:
        return INSTANCE_DISPLAY_NAMES[str(instance_id)]
    raw = str(name or "").strip()
    return INSTANCE_DISPLAY_NAMES.get(raw, raw)


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


def _render_section_body(sec: Dict[str, Any], blocks: List[str], heading_prefix: str = "") -> None:
    sid = (sec.get("id") or "").strip()
    title = (sec.get("title") or "").strip()
    head = f"{heading_prefix}{sid} {title}".strip()
    if head:
        blocks.append(head)
    for p in sec.get("paragraphs") or []:
        if p:
            blocks.append(str(p).strip())
    for b in sec.get("bullets") or []:
        if b:
            blocks.append(f"• {str(b).strip()}")
    table = sec.get("table")
    if isinstance(table, dict):
        headers = [str(h).strip() for h in (table.get("headers") or []) if str(h).strip()]
        if headers:
            blocks.append(" | ".join(headers))
        for row in table.get("rows") or []:
            cells = [str(c).strip() for c in row]
            if any(cells):
                blocks.append(" | ".join(cells))
    for p in sec.get("paragraphsAfter") or []:
        if p:
            blocks.append(str(p).strip())
    for sub in sec.get("subsections") or []:
        if isinstance(sub, dict):
            _render_section_body(sub, blocks)


def _render_bedingungen_text(terms: Dict[str, Any]) -> str:
    blocks: List[str] = []
    for sec in terms.get("sections") or []:
        if isinstance(sec, dict):
            _render_section_body(sec, blocks)
    return _join_paragraphs(blocks) if blocks else "—"


def _greeting_line(contact: str) -> str:
    name = (contact or "").strip()
    if not name:
        return "Sehr geehrte Damen und Herren,"
    return f"Sehr geehrte Damen und Herren / Sehr geehrte/r {name},"


def _split_swiss_address(address: str) -> Dict[str, str]:
    """Split 'Strasse Nr, PLZ Ort' into cover fields."""
    raw = re.sub(r"\s+", " ", (address or "").strip())
    if not raw:
        return {"strasse": "", "plz_ort": "", "adresse": ""}
    match = re.match(r"^(.*?),\s*(\d{4}\s+.+)$", raw)
    if match:
        return {
            "strasse": match.group(1).strip(),
            "plz_ort": match.group(2).strip(),
            "adresse": raw,
        }
    # Fallback: last token group with PLZ
    match = re.match(r"^(.*?)(\d{4}\s+.+)$", raw)
    if match and match.group(1).strip():
        return {
            "strasse": match.group(1).strip(" ,"),
            "plz_ort": match.group(2).strip(),
            "adresse": raw,
        }
    return {"strasse": raw, "plz_ort": "", "adresse": raw}


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

    instance_name = _display_instance_name(cfg.get("instanceName"), cfg.get("instanceId"))
    cfg_bits = " · ".join(
        x
        for x in [
            instance_name or None,
            f"{cfg['instanceCount']}×" if cfg.get("instanceCount") else None,
            _pluralize(cfg.get("deviceCount"), "Gerät", "Geräte"),
            _pluralize(cfg.get("zoneCount"), "Zone", "Zonen"),
            _pluralize(cfg.get("openingCount"), "Öffnung", "Öffnungen"),
        ]
        if x
    )

    # Kunden-Word: Leistungsumfang ohne Positionspreise/Stunden
    lines: List[Dict[str, Any]] = []
    scope_groups = offer.get("scopeGroups") or content.get("scopeGroups") or []
    if scope_groups:
        pos = 0
        for group in scope_groups:
            for item in group.get("items") or []:
                pos += 1
                lines.append(
                    {
                        "pos": pos,
                        "section": group.get("title") or "",
                        "sku": "",
                        "name": item.get("name") or "",
                        "description": item.get("description") or "",
                        "qty": "",
                        "unit_price": "",
                        "hours": "",
                        "amount": "",
                        "currency": "",
                    }
                )
    else:
        for line in offer.get("commercialLines") or []:
            desc = str(line.get("description") or "")
            desc = re.sub(r"\s*·\s*IC\s+[−\-]?\s*[\d.'\s,]+\s*EUR", "", desc, flags=re.IGNORECASE)
            desc = re.sub(r"\s*IC\s+[−\-]?\s*[\d.'\s,]+\s*EUR", "", desc, flags=re.IGNORECASE).strip()
            lines.append(
                {
                    "pos": line.get("pos", ""),
                    "section": line.get("section", ""),
                    "sku": line.get("sku", ""),
                    "name": line.get("name", ""),
                    "description": desc,
                    "qty": "",
                    "unit_price": "",
                    "hours": "",
                    "amount": "",
                    "currency": "",
                }
            )

    # Keine Bereichs-Totals hier: Projektrabatt gilt nur für Gesamttotal.
    scope_text_blocks = []
    for group in scope_groups:
        scope_text_blocks.append(group.get("title") or "")
        for item in group.get("items") or []:
            name = (item.get("name") or "").strip()
            desc = (item.get("description") or "").strip()
            if name and desc:
                scope_text_blocks.append(f"• {name}\n{desc}")
            elif name:
                scope_text_blocks.append(f"• {name}")
    scope_text = _join_paragraphs(scope_text_blocks)

    einleitung = _join_paragraphs(
        [
            content.get("intro") or "",
            content.get("introVariant") or "",
            content.get("recommendation") or "",
        ]
    )

    angebotsnummer = meta.get("archiveTitle") or meta.get("offerNumber", "")
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

    note = summary.get("note")
    if note and "IC-Preise" in str(note):
        note = "Alle Preise in CHF, exkl. MwSt."
    if not note:
        note = "Alle Preise in CHF, exkl. MwSt."

    addr = _split_swiss_address(customer.get("address", ""))
    prepared_by = (meta.get("preparedBy") or "").strip()
    # Cover: SSI-Ansprechpartner | Letzte Seite: Unterschriften (getrennt)
    ssi_contacts = offer.get("ssiContacts") or meta.get("ssiContacts") or []
    ssi1 = dict(ssi_contacts[0] if len(ssi_contacts) > 0 else {})
    ssi2 = dict(ssi_contacts[1] if len(ssi_contacts) > 1 else {})
    # Signaturen aus Bedingungen / Angebot — nicht dieselben wie Cover
    offer_sigs = offer.get("signatories") or []
    raw_sig1 = dict(offer_sigs[0] if len(offer_sigs) > 0 else sig1 or {})
    raw_sig2 = dict(offer_sigs[1] if len(offer_sigs) > 1 else sig2 or {})

    def _enrich(contact: Dict[str, Any], *, name_hint: str = "") -> Dict[str, Any]:
        if contact.get("email") and contact.get("name"):
            return contact
        found = _lookup_ssi_contact(
            contact_id=str(contact.get("id") or ""),
            name=str(contact.get("name") or name_hint or ""),
        )
        if found:
            merged = dict(found)
            merged.update({k: v for k, v in contact.items() if v})
            return merged
        return contact

    ssi1 = _enrich(ssi1, name_hint=prepared_by)
    ssi2 = _enrich(ssi2)
    raw_sig1 = _enrich(raw_sig1)
    raw_sig2 = _enrich(raw_sig2)

    def _fields(primary: Dict[str, Any], *, name_fallback: str = "") -> Dict[str, str]:
        """Nur Felder der gewählten Person — kein Fallback auf andere Signatur."""
        name = (primary.get("name") or name_fallback or "").strip()
        if not name and not primary.get("id"):
            return {"name": "", "title": "", "role": "", "email": "", "phone": ""}
        return {
            "name": name,
            "title": (primary.get("title") or "").strip(),
            "role": (primary.get("role") or "").strip(),
            "email": (primary.get("email") or "").strip(),
            "phone": (primary.get("phone") or "").strip(),
        }

    f1 = _fields(ssi1, name_fallback=prepared_by)
    f2 = _fields(ssi2)
    sf1 = _fields(raw_sig1)
    sf2 = _fields(raw_sig2)
    # Cover-Fallback nur wenn leer
    if not f1["name"] and not f2["name"]:
        f1 = _fields(sig1)
        f2 = _fields(sig2)
    # Signatur-Fallback auf Cover nur wenn Signaturen komplett leer
    if not sf1["name"] and not sf2["name"]:
        sf1, sf2 = f1, f2

    ssi1_name, ssi1_pos, ssi1_mail, ssi1_tel = f1["name"], f1["title"], f1["email"], f1["phone"]
    ssi2_name, ssi2_pos, ssi2_mail, ssi2_tel = f2["name"], f2["title"], f2["email"], f2["phone"]
    sig1_name, sig1_title, sig1_role = sf1["name"], sf1["title"], sf1["role"]
    sig2_name, sig2_title, sig2_role = sf2["name"], sf2["title"], sf2["role"]
    sig1_mail, sig2_mail = sf1["email"], sf2["email"]
    if prepared_by == "" and ssi1_name:
        prepared_by = ssi1_name

    def _details(pos: str, email: str, phone: str) -> str:
        parts = [p for p in [pos, email, phone] if p]
        return "\n".join(parts)

    def _sig_details(title: str, role: str, email: str = "") -> str:
        parts = [p for p in [title, role, email] if p]
        return "\n".join(parts)

    return {
        "dokument_label": content.get("documentLabel") or "Angebot / Preisliste",
        "titel": content.get("title") or "Angebot WAMAS® Lift & Store",
        "untertitel": content.get("subtitle")
        or "SSI SCHÄFER · Softwarelösung für Vertical Lift Modules (SSI LOGIMAT®)",
        "angebotsnummer": angebotsnummer,
        "datum": datum,
        "gueltig_bis": gueltig_bis,
        "meta_zeile": meta_zeile,
        "version_software": (offer.get("branding") or {}).get("version", "2.8"),
        "erstellt_von": prepared_by,
        "revision_von": revision_von,
        "kunde": customer.get("company", ""),
        "projekt": customer.get("projectName", ""),
        "ansprechpartner": customer.get("contact", ""),
        "anrede": _greeting_line(customer.get("contact", "")),
        "email": customer.get("email", ""),
        "telefon": customer.get("phone", ""),
        "fax": customer.get("fax", "") or "",
        "adresse": addr["adresse"],
        "strasse": addr["strasse"],
        "plz_ort": addr["plz_ort"],
        "ssi_kontakt_1_name": ssi1_name,
        "ssi_kontakt_1_position": ssi1_pos,
        "ssi_kontakt_1_email": ssi1_mail,
        "ssi_kontakt_1_telefon": ssi1_tel,
        "ssi_kontakt_1_details": _details(ssi1_pos, ssi1_mail, ssi1_tel),
        "ssi_kontakt_2_name": ssi2_name,
        "ssi_kontakt_2_position": ssi2_pos,
        "ssi_kontakt_2_email": ssi2_mail,
        "ssi_kontakt_2_telefon": ssi2_tel,
        "ssi_kontakt_2_details": _details(ssi2_pos, ssi2_mail, ssi2_tel),
        "konfiguration": cfg_bits,
        "instance_name": instance_name,
        "instance_count": cfg.get("instanceCount") or "",
        "device_count": cfg.get("deviceCount") or "",
        "zone_count": cfg.get("zoneCount") or "",
        "opening_count": cfg.get("openingCount") or "",
        "order_handling": "Ja" if cfg.get("hasOrderHandling") else "Nein",
        "einleitung": einleitung,
        "optionen_text": _render_optionen_text(content.get("selectedOptions") or []),
        "bedingungen_text": _render_bedingungen_text(terms),
        "leistungsumfang_text": scope_text,
        "preis_hinweis": note,
        "lizenz_total_chf": _money(lic.get("total"), lic.get("currency") or "CHF") if lic else "—",
        "lizenz_ic_eur": "",
        "marge_percent": "",
        "kurs_eur_chf": "",
        "it_total_chf": _money(it.get("total"), it.get("currency") or "CHF") if it else "—",
        "it_work_chf": "",
        "it_travel_chf": "",
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
        "signatur_1_name": sig1_name,
        "signatur_1_titel": sig1_title,
        "signatur_1_rolle": sig1_role,
        "signatur_1_details": _sig_details(sig1_title, sig1_role, sig1_mail),
        "signatur_2_name": sig2_name,
        "signatur_2_titel": sig2_title,
        "signatur_2_rolle": sig2_role,
        "signatur_2_details": _sig_details(sig2_title, sig2_role, sig2_mail),
    }


def _set_sdt_plain_text(sdt_xml: str, text: str) -> str:
    """Ersetzt den sichtbaren Inhalt eines w:sdt und entfernt Platzhalter-Flag."""
    safe = xml_escape(text or "")
    lines = safe.split("\n")
    if len(lines) <= 1:
        run_inner = f"<w:t>{safe}</w:t>"
    else:
        parts = [f'<w:t xml:space="preserve">{lines[0]}</w:t>']
        for line in lines[1:]:
            parts.append("<w:br/>")
            parts.append(f'<w:t xml:space="preserve">{line}</w:t>')
        run_inner = "".join(parts)

    def repl_content(match: re.Match) -> str:
        block = match.group(0)
        replaced = False

        def repl_run(rm: re.Match) -> str:
            nonlocal replaced
            if replaced:
                return ""
            replaced = True
            # Keep run props if present
            open_tag = rm.group(1)
            return f"{open_tag}{run_inner}</w:r>"

        block = re.sub(r"(<w:r\b[^>]*>).*?</w:r>", repl_run, block, flags=re.DOTALL)
        if not replaced:
            block = re.sub(
                r"(<w:sdtContent\b[^>]*>)(.*?)(</w:sdtContent>)",
                rf"\1<w:p><w:r>{run_inner}</w:r></w:p>\3",
                block,
                count=1,
                flags=re.DOTALL,
            )
        return block

    out = re.sub(
        r"<w:sdtContent\b[^>]*>.*?</w:sdtContent>",
        repl_content,
        sdt_xml,
        count=1,
        flags=re.DOTALL,
    )
    out = re.sub(r"<w:showingPlcHdr\s*/>", "", out)
    out = re.sub(r'<w:rStyle\s+w:val="Platzhaltertext"\s*/>', "", out)
    return out


def _iter_top_level_sdts(xml: str):
    """Yield (start, end, chunk) for each top-level w:sdt element."""
    i = 0
    while True:
        start = xml.find("<w:sdt>", i)
        if start < 0:
            return
        depth = 0
        j = start
        while j < len(xml):
            if xml.startswith("<w:sdt>", j):
                depth += 1
                j += 7
                continue
            if xml.startswith("</w:sdt>", j):
                depth -= 1
                j += 8
                if depth == 0:
                    yield start, j, xml[start:j]
                    i = j
                    break
                continue
            j += 1
        else:
            return


def apply_sdt_contact_fields(docx_bytes: bytes, context: Dict[str, Any]) -> bytes:
    """Befüllt Word-Dropdowns/Textfelder (Cover + Unterschriften) nach docxtpl-Render."""
    zin = zipfile.ZipFile(io.BytesIO(docx_bytes), "r")
    try:
        xml = zin.read("word/document.xml").decode("utf-8")
    except KeyError:
        zin.close()
        return docx_bytes

    replacements = 0
    pieces: List[str] = []
    cursor = 0
    for start, end, chunk in _iter_top_level_sdts(xml):
        pieces.append(xml[cursor:start])
        alias_m = re.search(r'<w:alias\b[^>]*w:val="([^"]+)"', chunk)
        tag_m = re.search(r'<w:tag\b[^>]*w:val="([^"]+)"', chunk)
        key = (alias_m.group(1) if alias_m else "") or (tag_m.group(1) if tag_m else "")
        ctx_key = _SDT_TEXT_MAP.get(key)
        if ctx_key and context.get(ctx_key) is not None:
            pieces.append(_set_sdt_plain_text(chunk, str(context.get(ctx_key) or "")))
            replacements += 1
        else:
            pieces.append(chunk)
        cursor = end
    pieces.append(xml[cursor:])
    new_xml = "".join(pieces)

    if replacements == 0:
        zin.close()
        return docx_bytes

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml.encode("utf-8")
            zout.writestr(item, data)
    zin.close()
    return out.getvalue()


def build_offer_docx(offer: Dict[str, Any]) -> bytes:
    """Befüllt die Anhang-Vorlage (ein Dokument, SSI-Anhang inkl. Platzhalter)."""
    if offer.get("kind") != "offer_document":
        raise ValueError("Word-Export ist nur für Gesamtangebote (offer_document) verfügbar.")

    template_path = ensure_annex_template()
    context = build_template_context(offer)

    tpl = DocxTemplate(str(template_path))
    # autoescape=True: & und ähnliche Zeichen korrekt als XML escapen
    tpl.render(context, autoescape=True)

    out = io.BytesIO()
    tpl.save(out)
    return apply_sdt_contact_fields(out.getvalue(), context)
