"""WAMAS Lift & Store – License & IT Effort Calculator (SSI SCHÄFER)"""

from __future__ import annotations

import io
import json
import math
import re
import uuid
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from pydantic import BaseModel, Field

from docx_export import build_offer_docx

GEOADMIN_SEARCH_URL = "https://api3.geo.admin.ch/rest/services/api/SearchServer"

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
DATA_DIR = BASE_DIR / "data"
OFFERS_DIR = DATA_DIR / "offers"
CATALOG_FILE = DATA_DIR / "catalog.json"
IT_CATALOG_FILE = DATA_DIR / "it_catalog.json"
OFFER_TEMPLATE_FILE = DATA_DIR / "offer_template.json"
COMMERCIAL_TERMS_FILE = DATA_DIR / "commercial_terms_ch.json"

OFFERS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="WAMAS Lift & Store Calculator", version="3.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def load_catalog() -> Dict[str, Any]:
    return json.loads(CATALOG_FILE.read_text(encoding="utf-8"))


def load_it_catalog() -> Dict[str, Any]:
    return json.loads(IT_CATALOG_FILE.read_text(encoding="utf-8"))


def load_offer_template() -> Dict[str, Any]:
    return json.loads(OFFER_TEMPLATE_FILE.read_text(encoding="utf-8"))


def load_commercial_terms() -> Dict[str, Any]:
    return json.loads(COMMERCIAL_TERMS_FILE.read_text(encoding="utf-8"))


def offer_path(offer_id: str) -> Path:
    safe = "".join(c for c in offer_id if c.isalnum() or c in "-_")
    return OFFERS_DIR / f"{safe}.json"


# ---------------------------------------------------------------------------
# License calculator (IC)
# ---------------------------------------------------------------------------


class CustomerInfo(BaseModel):
    company: str = Field(..., min_length=1)
    contact: str = ""
    email: str = ""
    phone: str = ""
    address: str = ""
    projectName: str = ""


class OfferRequest(BaseModel):
    customer: CustomerInfo
    instanceId: str = "basic"
    instanceCount: int = Field(1, ge=1, le=50)
    selectedAddons: List[str] = Field(default_factory=list)
    extraOpeningClients: int = Field(0, ge=0, le=200)
    extraAdminClients: int = Field(0, ge=0, le=200)
    mobileTerminalClients: int = Field(0, ge=0, le=200)
    thirdPartyVlmTypes: int = Field(0, ge=0, le=50)
    testInstances: int = Field(0, ge=0, le=20)
    upgradeYears: int = Field(0, ge=0, le=20)
    notes: str = ""
    preparedBy: str = ""
    # Optional: überschreibt Katalog-Defaults (interne Kalkulation)
    licenseMarginPercent: Optional[float] = Field(None, ge=0, le=500)
    eurToChfRate: Optional[float] = Field(None, gt=0, le=10)


def resolve_discount(catalog: Dict[str, Any], sll: int) -> Dict[str, Any]:
    chosen = catalog["quantityDiscounts"][0]
    for row in catalog["quantityDiscounts"]:
        min_s = int(row["minSll"])
        max_s = row["maxSll"]
        if sll >= min_s and (max_s is None or sll <= int(max_s)):
            chosen = row
    return chosen


def calculate_offer(payload: OfferRequest) -> Dict[str, Any]:
    catalog = load_catalog()
    instances = {i["id"]: i for i in catalog["instances"]}
    addons = {a["id"]: a for a in catalog["addons"]}
    clients = {c["id"]: c for c in catalog["clientLicenses"]}
    misc = {m["id"]: m for m in catalog["misc"]}
    function_catalog = catalog.get("functionCatalog", {})

    if payload.instanceId not in instances:
        raise HTTPException(status_code=400, detail="Unbekannte Instance")

    instance = instances[payload.instanceId]
    included_functions = [
        {
            "id": fid,
            "name": function_catalog.get(fid, {}).get("name", fid),
            "description": function_catalog.get(fid, {}).get("description", ""),
            "manualRefs": function_catalog.get(fid, {}).get("manualRefs", []),
        }
        for fid in instance.get("includedFunctionIds", [])
    ]
    lines: List[Dict[str, Any]] = []
    sll = 0

    lines.append(
        {
            "sku": f"INST-{instance['id'].upper()}",
            "name": instance["name"],
            "description": instance.get("functionalSummary") or instance["description"],
            "qty": payload.instanceCount,
            "unitPrice": instance["price"],
            "total": instance["price"] * payload.instanceCount,
            "category": "instance",
            "sllUnits": payload.instanceCount,
        }
    )
    sll += payload.instanceCount

    for addon_id in payload.selectedAddons:
        addon = addons.get(addon_id)
        if not addon:
            raise HTTPException(status_code=400, detail=f"Unbekanntes Add-on: {addon_id}")
        if payload.instanceId not in addon["availableFor"]:
            raise HTTPException(
                status_code=400,
                detail=f"Add-on '{addon['name']}' ist für {instance['name']} nicht verfügbar",
            )
        qty = payload.instanceCount
        units = int(addon.get("sllUnits", 1)) * qty
        lines.append(
            {
                "sku": f"ADD-{addon_id.upper()}",
                "name": addon["name"],
                "description": addon.get("functionalDescription") or addon.get("description", ""),
                "qty": qty,
                "unitPrice": addon["price"],
                "total": addon["price"] * qty,
                "category": "addon",
                "sllUnits": units,
                "manualRefs": addon.get("manualRefs", []),
            }
        )
        sll += units

    client_qty = {
        "extra_opening": payload.extraOpeningClients,
        "extra_admin": payload.extraAdminClients,
        "mobile_terminal": payload.mobileTerminalClients,
        "third_party_vlm": payload.thirdPartyVlmTypes,
    }
    for cid, qty in client_qty.items():
        if qty <= 0:
            continue
        client = clients[cid]
        if payload.instanceId not in client["availableFor"]:
            raise HTTPException(
                status_code=400,
                detail=f"'{client['name']}' ist für {instance['name']} nicht verfügbar",
            )
        units = int(client.get("sllUnitsPerQty", 1)) * qty
        lines.append(
            {
                "sku": f"CLI-{cid.upper()}",
                "name": client["name"],
                "description": client.get("functionalDescription") or client["description"],
                "qty": qty,
                "unitPrice": client["price"],
                "total": client["price"] * qty,
                "category": "client",
                "sllUnits": units,
                "manualRefs": client.get("manualRefs", []),
            }
        )
        sll += units

    if payload.testInstances > 0:
        test = misc["test_instance"]
        units = int(test.get("sllUnitsPerQty", 1)) * payload.testInstances
        lines.append(
            {
                "sku": "MISC-TEST",
                "name": test["name"],
                "description": test.get("functionalDescription") or test["description"],
                "qty": payload.testInstances,
                "unitPrice": test["price"],
                "total": test["price"] * payload.testInstances,
                "category": "misc",
                "sllUnits": units,
            }
        )
        sll += units

    if payload.upgradeYears > 0:
        upgrade = misc["upgrade_fee"]
        lines.append(
            {
                "sku": "MISC-UPGRADE",
                "name": upgrade["name"],
                "description": upgrade.get("functionalDescription") or upgrade["description"],
                "qty": payload.upgradeYears,
                "unitPrice": upgrade["price"],
                "total": upgrade["price"] * payload.upgradeYears,
                "category": "misc",
                "sllUnits": 0,
                "recurring": True,
            }
        )

    subtotal = round(sum(l["total"] for l in lines), 2)
    discount = resolve_discount(catalog, sll)
    discount_amount = round(subtotal * (discount["percent"] / 100), 2)
    net = round(subtotal - discount_amount, 2)
    product = catalog["product"]
    margin_percent = float(
        payload.licenseMarginPercent
        if payload.licenseMarginPercent is not None
        else (product.get("licenseMarginPercent") or 0)
    )
    eur_to_chf = float(
        payload.eurToChfRate
        if payload.eurToChfRate is not None
        else (product.get("eurToChfRate") or 1)
    )
    margin_factor = 1.0 + (margin_percent / 100.0)
    margin_amount_eur = round(net * (margin_percent / 100.0), 2)
    sell_net_eur = round(net * margin_factor, 2)
    sell_net_chf = round(sell_net_eur * eur_to_chf, 2)
    # Deckungsbeitrag intern: Verkauf − Einkauf (Einkauf CHF = IC×Kurs)
    cost_chf = round(net * eur_to_chf, 2)
    contribution_margin_eur = margin_amount_eur
    contribution_margin_chf = round(sell_net_chf - cost_chf, 2)
    contribution_margin_percent = (
        round((contribution_margin_chf / sell_net_chf) * 100, 1) if sell_net_chf else 0.0
    )
    vat_rate = float(product.get("vatRate") or 0)
    vat = round(sell_net_chf * vat_rate, 2)
    gross = round(sell_net_chf + vat, 2)

    # Verkaufspreise je Position (IC → +Marge → CHF)
    for line in lines:
        ic_unit = float(line.get("unitPrice") or 0)
        ic_total = float(line.get("total") or 0)
        line["unitPriceIcEur"] = ic_unit
        line["totalIcEur"] = ic_total
        line["unitPrice"] = round(ic_unit * margin_factor * eur_to_chf, 2)
        line["total"] = round(ic_total * margin_factor * eur_to_chf, 2)
        line["currency"] = product.get("offerCurrency", "CHF")

    created = datetime.now()
    valid_until = created + timedelta(days=int(product["validityDays"]))
    included_opening = instance["includedOpeningClients"] * payload.instanceCount
    included_admin = instance["includedAdminClients"] * payload.instanceCount

    scope = [
        f"{payload.instanceCount}× {instance['name']} (IC)",
        f"Inklusive Clients: {included_opening} Opening, {included_admin} Admin",
    ]
    for fn in included_functions:
        scope.append(f"{fn['name']}: {fn['description']}")
    for addon_id in payload.selectedAddons:
        addon = addons.get(addon_id)
        if addon:
            scope.append(f"{addon['name']}: {addon.get('functionalDescription') or addon.get('description', '')}")
    scope.append(f"SLL-Einheiten: {sll} → Rabatt {discount['percent']}% ({discount['label']})")
    scope.append(f"Marge {margin_percent:.0f}% auf IC · Kurs EUR→CHF {eur_to_chf}")

    return {
        "kind": "license",
        "product": product,
        "customer": payload.customer.model_dump(),
        "configuration": {
            "instanceId": payload.instanceId,
            "instanceName": instance["name"],
            "instanceCount": payload.instanceCount,
            "selectedAddons": list(payload.selectedAddons),
            "extraOpeningClients": payload.extraOpeningClients,
            "extraAdminClients": payload.extraAdminClients,
            "mobileTerminalClients": payload.mobileTerminalClients,
            "thirdPartyVlmTypes": payload.thirdPartyVlmTypes,
            "testInstances": payload.testInstances,
            "upgradeYears": payload.upgradeYears,
            "includedOpeningClients": included_opening,
            "includedAdminClients": included_admin,
            "includedFunctions": included_functions,
            "notes": payload.notes,
            "preparedBy": payload.preparedBy,
            "sllCount": sll,
            "discountPercent": discount["percent"],
            "discountLabel": discount["label"],
            "licenseMarginPercent": margin_percent,
            "eurToChfRate": eur_to_chf,
        },
        "scopeOfSupply": scope,
        "lines": lines,
        "totals": {
            "sllCount": sll,
            "subtotal": subtotal,
            "discountPercent": discount["percent"],
            "discountAmount": discount_amount,
            "net": net,
            "icCurrency": product.get("currency", "EUR"),
            "marginPercent": margin_percent,
            "marginAmountEur": margin_amount_eur,
            "sellNetEur": sell_net_eur,
            "eurToChfRate": eur_to_chf,
            "costChf": cost_chf,
            "sellNetChf": sell_net_chf,
            "contributionMarginEur": contribution_margin_eur,
            "contributionMarginChf": contribution_margin_chf,
            "contributionMarginPercent": contribution_margin_percent,
            "vatRate": vat_rate,
            "vat": vat,
            "gross": gross,
            "currency": product.get("offerCurrency", "CHF"),
            "amount": sell_net_chf,
        },
        "meta": {
            "createdAt": created.isoformat(timespec="seconds"),
            "validUntil": valid_until.date().isoformat(),
            "offerNumber": f"WLS-{created.strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}",
            "productVersion": product.get("version", ""),
            "priceBasis": product.get("priceBasis", ""),
        },
    }


# ---------------------------------------------------------------------------
# IT effort calculator (from Excel)
# ---------------------------------------------------------------------------


class ItCustomExtension(BaseModel):
    description: str = ""
    hours: float = Field(0, ge=0, le=1000)


class ItOfferRequest(BaseModel):
    customer: CustomerInfo
    realizationPeriod: str = ""
    deviceCount: int = Field(1, ge=1, le=100)
    zoneCount: int = Field(1, ge=1, le=100)
    openingCount: int = Field(1, ge=1, le=200)
    options: Dict[str, bool] = Field(default_factory=dict)
    customExtensions: List[ItCustomExtension] = Field(default_factory=list)
    trips: int = Field(0, ge=0, le=200)
    travelHoursPerTrip: float = Field(0, ge=0, le=48)
    kmPerTrip: float = Field(0, ge=0, le=5000)
    overnightCount: int = Field(0, ge=0, le=200)
    mealCount: int = Field(0, ge=0, le=500)
    notes: str = ""
    preparedBy: str = ""
    # Nur Marge (CHF intern, kein Währungskurs)
    itMarginPercent: Optional[float] = Field(None, ge=0, le=500)


def logimat_hours(devices: int, base: float, pct2: float, pct_n: float) -> float:
    """Parity with Excel VBA LOGIMAT()."""
    if devices <= 0:
        return 0.0
    total = base
    if devices >= 2:
        total += base * (pct2 / 100.0)
    if devices >= 3:
        total += (devices - 2) * base * (pct_n / 100.0)
    return float(math.ceil(total - 1e-9))


def zone_hours(zones: int, percent: float, device_hours: float) -> float:
    if zones <= 1:
        return 0.0
    return (zones - 1) * (percent / 100.0) * device_hours


def opening_hours(openings: int, percent: float, device_hours: float, devices: int) -> float:
    if openings <= devices:
        return 0.0
    return (openings - devices) * (percent / 100.0) * device_hours


def calculate_it_offer(payload: ItOfferRequest) -> Dict[str, Any]:
    cat = load_it_catalog()
    rates = cat["rates"]
    effort = cat["effort"]
    option_defs = {o["id"]: o for o in cat["options"]}
    hourly = float(rates["hourlyRate"])

    if payload.openingCount < payload.deviceCount:
        raise HTTPException(
            status_code=400,
            detail="Anzahl Öffnungen darf nicht kleiner als Anzahl Geräte sein.",
        )

    lines: List[Dict[str, Any]] = []
    device_h = logimat_hours(
        payload.deviceCount,
        float(effort["baseHours"]),
        float(effort["secondDevicePercent"]),
        float(effort["additionalDevicePercent"]),
    )
    lines.append(
        {
            "sku": "IT-DEVICES",
            "name": "Anzahl Geräte",
            "description": "Grundaufwand Installation/Parametrierung/IBN/Schulung/Dokumentation",
            "qty": payload.deviceCount,
            "hours": device_h,
            "amount": round(device_h * hourly, 2),
            "category": "logimat",
        }
    )

    zone_h = zone_hours(payload.zoneCount, float(effort["additionalZonePercent"]), device_h)
    lines.append(
        {
            "sku": "IT-ZONES",
            "name": "Anzahl Zonen",
            "description": "Zusatzaufwand je zusätzlicher Zone",
            "qty": payload.zoneCount,
            "hours": round(zone_h, 2),
            "amount": round(zone_h * hourly, 2),
            "category": "logimat",
        }
    )

    open_h = opening_hours(
        payload.openingCount,
        float(effort["additionalOpeningPercent"]),
        device_h,
        payload.deviceCount,
    )
    lines.append(
        {
            "sku": "IT-OPENINGS",
            "name": "Anzahl Öffnungen",
            "description": "Zusatzaufwand je zusätzlicher Bedienöffnung",
            "qty": payload.openingCount,
            "hours": round(open_h, 2),
            "amount": round(open_h * hourly, 2),
            "category": "logimat",
        }
    )

    selected = {oid: bool(payload.options.get(oid)) for oid in option_defs}
    extra_openings = max(0, payload.openingCount - payload.deviceCount)

    def add_option(oid: str, hours: float, note: str = "") -> None:
        opt = option_defs[oid]
        lines.append(
            {
                "sku": f"IT-{oid.upper()}",
                "name": opt["name"],
                "description": opt["description"],
                "qty": 1 if hours else 0,
                "hours": round(hours, 2),
                "amount": round(hours * hourly, 2) if isinstance(hours, (int, float)) else 0,
                "category": "option",
                "selected": selected.get(oid, False),
                "offerBullets": opt.get("offerBullets", []),
                "note": note,
            }
        )

    add_option("orderHandling", effort["orderHandlingHours"] if selected["orderHandling"] else 0)

    add_option("externalStorage", effort["externalStorageHours"] if selected["externalStorage"] else 0)
    add_option("thirdPartyPanel", effort["thirdPartyPanelHours"] if selected["thirdPartyPanel"] else 0)

    if selected["rfid"]:
        rfid_h = (
            effort["rfidBaseHours"]
            if extra_openings == 0
            else extra_openings * effort["rfidPerExtraOpeningHours"] + effort["rfidBaseHours"]
        )
    else:
        rfid_h = 0
    add_option("rfid", rfid_h)

    if selected["scanner"]:
        scanner_h = (
            effort["scannerBaseHours"]
            if extra_openings == 0
            else extra_openings * effort["scannerPerExtraOpeningHours"] + effort["scannerBaseHours"]
        )
    else:
        scanner_h = 0
    add_option("scanner", scanner_h)

    add_option(
        "itemMasterInterface",
        effort["itemMasterInterfaceHours"] if selected["itemMasterInterface"] else 0,
    )
    add_option("pickLabel", effort["pickLabelHours"] if selected["pickLabel"] else 0)

    test_note = ""
    test_h = 0.0
    if selected["testSystem"]:
        if selected["orderHandling"]:
            test_h = float(effort["testSystemHours"])
        else:
            test_note = "Kein Order Handling gewählt!"
    add_option("testSystem", test_h, test_note)

    add_option(
        "advancedSecurity",
        effort["advancedSecurityHours"] if selected["advancedSecurity"] else 0,
    )

    for idx, ext in enumerate(payload.customExtensions[:5], start=1):
        if not ext.description and not ext.hours:
            continue
        lines.append(
            {
                "sku": f"IT-EXT-{idx}",
                "name": f"Erweiterung {idx}",
                "description": ext.description or f"Projektspezifische Erweiterung {idx}",
                "qty": 1,
                "hours": round(float(ext.hours), 2),
                "amount": round(float(ext.hours) * hourly, 2),
                "category": "custom",
            }
        )

    work_hours = round(sum(float(l["hours"] or 0) for l in lines if l["category"] != "travel"), 2)
    work_amount = round(sum(float(l["amount"] or 0) for l in lines if l["category"] != "travel"), 2)

    travel_hours = float(payload.travelHoursPerTrip)
    travel_lines = [
        {
            "sku": "IT-TRAVEL-TIME",
            "name": "Fahrzeit",
            "description": f"{payload.trips} Roundtrips × {travel_hours} h",
            "qty": payload.trips,
            "hours": travel_hours,
            "amount": round(travel_hours * hourly * payload.trips, 2),
            "category": "travel",
        },
        {
            "sku": "IT-TRAVEL-KM",
            "name": "Kilometer",
            "description": f"{payload.kmPerTrip} km Roundtrip × {payload.trips} Fahrten",
            "qty": payload.trips,
            "hours": 0,
            "amount": round(payload.kmPerTrip * float(rates["kmRate"]) * payload.trips, 2),
            "category": "travel",
        },
        {
            "sku": "IT-TRAVEL-OVERNIGHT",
            "name": "Übernachtung",
            "description": "Anzahl Übernachtungen",
            "qty": payload.overnightCount,
            "hours": 0,
            "amount": round(payload.overnightCount * float(rates["overnightRate"]), 2),
            "category": "travel",
        },
        {
            "sku": "IT-TRAVEL-MEALS",
            "name": "Verpflegung",
            "description": "Anzahl Verpflegungen",
            "qty": payload.mealCount,
            "hours": 0,
            "amount": round(payload.mealCount * float(rates["mealRate"]), 2),
            "category": "travel",
        },
    ]
    lines.extend(travel_lines)
    travel_amount_cost = round(sum(l["amount"] for l in travel_lines), 2)
    work_amount_cost = round(work_amount, 2)

    # IT intern in CHF: nur Marge, kein EUR→CHF-Kurs
    margin_percent = float(
        payload.itMarginPercent
        if payload.itMarginPercent is not None
        else (rates.get("marginPercent") or 0)
    )
    margin_factor = 1.0 + (margin_percent / 100.0)
    for line in lines:
        cost = float(line.get("amount") or 0)
        line["amountCost"] = cost
        line["amount"] = round(cost * margin_factor, 2)

    work_amount = round(work_amount_cost * margin_factor, 2)
    travel_amount = round(travel_amount_cost * margin_factor, 2)
    margin_amount = round((work_amount_cost + travel_amount_cost) * (margin_percent / 100.0), 2)
    total_hours = round(work_hours + travel_hours, 2)
    total_amount_cost = round(work_amount_cost + travel_amount_cost, 2)
    total_amount = round(total_amount_cost * margin_factor, 2)
    contribution_margin_percent = (
        round((margin_amount / total_amount) * 100, 1) if total_amount else 0.0
    )
    hourly_sell = round(hourly * margin_factor, 2)

    created = datetime.now()
    offer_sections = [
        {
            "title": f"Software WAMAS Lift & Store für {payload.deviceCount} Geräte",
            "amount": round(sum(l["amount"] for l in lines if l["category"] == "logimat"), 2),
            "bullets": cat["baseScopeBullets"],
        }
    ]
    for oid, enabled in selected.items():
        if not enabled:
            continue
        opt = option_defs[oid]
        amount = next((l["amount"] for l in lines if l.get("sku") == f"IT-{oid.upper()}"), 0)
        if oid == "testSystem" and not selected["orderHandling"]:
            continue
        offer_sections.append(
            {
                "title": opt["name"],
                "amount": amount,
                "bullets": opt.get("offerBullets", []),
                "description": opt["description"],
            }
        )
    customs = [l for l in lines if l["category"] == "custom" and l["amount"]]
    if customs:
        offer_sections.append(
            {
                "title": "Projektspezifische Erweiterungen",
                "amount": round(sum(c["amount"] for c in customs), 2),
                "bullets": [c["description"] for c in customs],
            }
        )
    offer_sections.append(
        {
            "title": "Reisekosten",
            "amount": travel_amount,
            "bullets": [
                f"Fahrzeiten: {travel_lines[0]['amount']} {cat['meta']['currency']}",
                f"Kilometerentschädigung: {travel_lines[1]['amount']} {cat['meta']['currency']}",
                f"Übernachtungskosten: {travel_lines[2]['amount']} {cat['meta']['currency']}",
                f"Verpflegungskosten: {travel_lines[3]['amount']} {cat['meta']['currency']}",
            ],
        }
    )

    return {
        "kind": "it",
        "product": {
            "name": cat["meta"]["title"],
            "version": cat["meta"]["version"],
            "currency": cat["meta"]["currency"],
            "disclaimer": cat["meta"]["disclaimer"],
            "vendor": "SSI SCHÄFER",
        },
        "customer": payload.customer.model_dump(),
        "configuration": {
            "realizationPeriod": payload.realizationPeriod,
            "deviceCount": payload.deviceCount,
            "zoneCount": payload.zoneCount,
            "openingCount": payload.openingCount,
            "options": selected,
            "customExtensions": [e.model_dump() for e in payload.customExtensions],
            "trips": payload.trips,
            "travelHoursPerTrip": payload.travelHoursPerTrip,
            "kmPerTrip": payload.kmPerTrip,
            "overnightCount": payload.overnightCount,
            "mealCount": payload.mealCount,
            "hourlyRate": hourly,
            "hourlyRateSell": hourly_sell,
            "itMarginPercent": margin_percent,
            "notes": payload.notes,
            "preparedBy": payload.preparedBy,
        },
        "lines": lines,
        "offerSections": offer_sections,
        "totals": {
            "workHours": work_hours,
            "workAmountCost": work_amount_cost,
            "workAmount": work_amount,
            "travelHours": travel_hours,
            "travelAmountCost": travel_amount_cost,
            "travelAmount": travel_amount,
            "totalHours": total_hours,
            "totalAmountCost": total_amount_cost,
            "marginPercent": margin_percent,
            "marginAmount": margin_amount,
            "contributionMarginChf": margin_amount,
            "contributionMarginPercent": contribution_margin_percent,
            "totalAmount": total_amount,
            "currency": cat["meta"]["currency"],
        },
        "meta": {
            "createdAt": created.isoformat(timespec="seconds"),
            "offerNumber": f"IT-{created.strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}",
            "source": cat["meta"]["source"],
        },
    }


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------


@app.get("/api/health")
def health():
    return {"ok": True, "service": "WAMAS Lift & Store Calculator", "modules": ["license", "it"]}


def _strip_html(text: str) -> str:
    cleaned = re.sub(r"<[^>]+>", "", text or "")
    return (
        cleaned.replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .strip()
    )


def format_swiss_address(attrs: Dict[str, Any]) -> str:
    """Normalize GeoAdmin label to 'Strasse Nr, PLZ Ort'."""
    label = _strip_html(str(attrs.get("label") or ""))
    label = re.sub(r"\s+#\s*", " ", label)
    label = re.sub(r"\s+", " ", label).strip()
    match = re.match(r"^(.+?)\s+(\d{4}\s+.+)$", label)
    if match:
        return f"{match.group(1).strip()}, {match.group(2).strip()}"
    return label


@app.get("/api/geo/address-suggest")
def address_suggest(
    q: str = Query("", min_length=0, max_length=200),
    limit: int = Query(8, ge=1, le=15),
) -> Dict[str, Any]:
    """Swiss address autocomplete via geo.admin.ch (official building addresses)."""
    query = (q or "").strip()
    if len(query) < 2:
        return {"suggestions": [], "source": "geo.admin.ch"}

    params = urllib.parse.urlencode(
        {
            "searchText": query,
            "type": "locations",
            "origins": "address",
            "limit": str(limit),
            "sr": "2056",
            "lang": "de",
        }
    )
    url = f"{GEOADMIN_SEARCH_URL}?{params}"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "WAMAS-LiftStore-Angebot/1.0 (SSI Schweiz)"},
    )
    try:
        with urllib.request.urlopen(req, timeout=6) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise HTTPException(
            status_code=502,
            detail=f"GeoAdmin-Fehler ({exc.code})",
        ) from exc
    except Exception as exc:
        raise HTTPException(
            status_code=502,
            detail="GeoAdmin Adresssuche nicht erreichbar",
        ) from exc

    suggestions: List[Dict[str, Any]] = []
    for row in payload.get("results") or []:
        attrs = row.get("attrs") or {}
        if attrs.get("origin") not in (None, "address"):
            continue
        label = format_swiss_address(attrs)
        if not label:
            continue
        suggestions.append(
            {
                "label": label,
                "lat": attrs.get("lat"),
                "lon": attrs.get("lon"),
                "featureId": attrs.get("featureId"),
            }
        )
    return {"suggestions": suggestions, "source": "geo.admin.ch"}


@app.get("/api/catalog")
def get_catalog():
    return load_catalog()


@app.get("/api/it/catalog")
def get_it_catalog():
    return load_it_catalog()


@app.get("/api/offer/template")
def get_offer_template():
    return load_offer_template()


@app.get("/api/offer/commercial-terms")
def get_commercial_terms():
    return load_commercial_terms()


class ComposeOfferRequest(BaseModel):
    licenseOfferId: Optional[str] = None
    itOfferId: Optional[str] = None
    license: Optional[Dict[str, Any]] = None
    it: Optional[Dict[str, Any]] = None
    basedOnOfferNumber: Optional[str] = None
    # Kommerzielle Anpassung vor Angebotserzeugung (Kundenangebot)
    discountPercent: Optional[float] = Field(None, ge=0, le=100)
    discountAmountChf: Optional[float] = Field(None, ge=0)
    roundTo: Optional[int] = Field(None, ge=0)  # 0/None = aus, sonst z.B. 10/50/100


def _load_saved_offer(offer_id: Optional[str]) -> Optional[Dict[str, Any]]:
    if not offer_id:
        return None
    path = offer_path(offer_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail=f"Angebot nicht gefunden: {offer_id}")
    return json.loads(path.read_text(encoding="utf-8"))


@app.post("/api/offer/compose")
def compose_offer(payload: ComposeOfferRequest):
    """Baut ein Angebotsdokument aus Lizenz- und/oder IT-Kalkulation im Stil des Word-Anhangs."""
    template = load_offer_template()
    license_offer = payload.license or _load_saved_offer(payload.licenseOfferId)
    it_offer = payload.it or _load_saved_offer(payload.itOfferId)
    if not license_offer and not it_offer:
        raise HTTPException(status_code=400, detail="Lizenz- und/oder IT-Kalkulation erforderlich")

    customer = (license_offer or it_offer).get("customer", {})
    prepared_by = (
        (license_offer or {}).get("configuration", {}).get("preparedBy")
        or (it_offer or {}).get("configuration", {}).get("preparedBy")
        or ""
    )

    selected_addons = set((license_offer or {}).get("configuration", {}).get("selectedAddons") or [])
    instance_name = (license_offer or {}).get("configuration", {}).get("instanceName", "")
    has_order_handling = (
        (license_offer or {}).get("configuration", {}).get("instanceId") == "advanced"
        or bool((it_offer or {}).get("configuration", {}).get("options", {}).get("orderHandling"))
    )

    standard = []
    for fn in template["standardFunctions"]:
        if fn.get("requiresOrderHandling") and not has_order_handling:
            continue
        standard.append(fn)

    options = []
    option_texts = template.get("optionTexts", {})
    for addon_id in sorted(selected_addons):
        if addon_id in option_texts:
            options.append({"id": addon_id, **option_texts[addon_id]})

    # also include IT-selected options not already covered
    it_opts = (it_offer or {}).get("configuration", {}).get("options") or {}
    it_to_licenseish = {
        "externalStorage": "external_storage",
        "rfid": "rfid_login",
        "pickLabel": "printing_support",
        "advancedSecurity": "advanced_security",
    }
    for it_id, lic_id in it_to_licenseish.items():
        if it_opts.get(it_id) and lic_id in option_texts and lic_id not in selected_addons:
            options.append({"id": lic_id, **option_texts[lic_id]})

    commercial: List[Dict[str, Any]] = []
    pos = 0

    def add_commercial(section: str, line: Dict[str, Any], *, amount_key: str = "total") -> None:
        nonlocal pos
        pos += 1
        amount = line.get(amount_key, line.get("amount", line.get("total")))
        commercial.append(
            {
                "pos": pos,
                "section": section,
                "sku": line.get("sku", ""),
                "name": line.get("name"),
                "description": line.get("description") or line.get("note") or "",
                "qty": line.get("qty"),
                "unitPrice": line.get("unitPrice"),
                "hours": line.get("hours"),
                "amount": amount,
                "currency": line.get("currency"),
                "category": line.get("category", ""),
            }
        )

    if license_offer:
        lic_totals = license_offer.get("totals") or {}
        lic_currency = lic_totals.get("currency", "CHF")
        margin_pct = float(lic_totals.get("marginPercent") or 0)
        fx = float(lic_totals.get("eurToChfRate") or 1)
        margin_factor = 1.0 + (margin_pct / 100.0)
        for line in license_offer.get("lines", []):
            # Kundenangebot: nur Verkaufspreise, keine IC-/Einkaufspreise in der Beschreibung
            add_commercial(
                "A · Softwarelizenzen (Verkauf CHF)",
                {
                    **line,
                    "currency": lic_currency,
                    "description": line.get("description") or "",
                },
                amount_key="total",
            )
        if lic_totals.get("discountAmount"):
            disc_chf = round(float(lic_totals["discountAmount"]) * margin_factor * fx, 2)
            pos += 1
            commercial.append(
                {
                    "pos": pos,
                    "section": "A · Softwarelizenzen (Verkauf CHF)",
                    "sku": "DISC-SLL",
                    "name": f"Mengenrabatt SLL ({lic_totals.get('discountPercent', 0)}%)",
                    "description": license_offer.get("configuration", {}).get("discountLabel", "") or "",
                    "qty": 1,
                    "unitPrice": -disc_chf,
                    "hours": None,
                    "amount": -disc_chf,
                    "currency": lic_currency,
                    "category": "discount",
                }
            )

    if it_offer:
        it_currency = it_offer.get("totals", {}).get("currency", "CHF")
        hourly = (it_offer.get("configuration") or {}).get("hourlyRateSell") or (
            it_offer.get("configuration") or {}
        ).get("hourlyRate")
        for line in it_offer.get("lines", []):
            # nur relevante / gewählte Positionen mit Aufwand oder Betrag
            if line.get("category") == "option" and not line.get("selected") and not line.get("amount"):
                continue
            if not (line.get("amount") or line.get("hours") or line.get("qty")):
                continue
            if line.get("category") == "travel" and not line.get("amount") and not line.get("qty"):
                continue
            section = (
                "C · Reisekosten"
                if line.get("category") == "travel"
                else "B · IT-Aufwand / Services"
            )
            unit = None
            if line.get("category") != "travel" and hourly and line.get("hours"):
                unit = hourly
            elif line.get("qty") and line.get("amount") and line.get("category") == "travel":
                qty = float(line.get("qty") or 0)
                unit = round(float(line.get("amount") or 0) / qty, 2) if qty else None
            add_commercial(
                section,
                {
                    **line,
                    "unitPrice": unit,
                    "currency": it_currency,
                },
                amount_key="amount",
            )

    license_totals = (license_offer or {}).get("totals")
    it_totals = (it_offer or {}).get("totals")
    lic_sell_chf = (license_totals or {}).get("sellNetChf")
    it_total_chf = (it_totals or {}).get("totalAmount")
    subtotal_chf = None
    if lic_sell_chf is not None or it_total_chf is not None:
        subtotal_chf = round(float(lic_sell_chf or 0) + float(it_total_chf or 0), 2)

    # Kunden-Leistungsumfang ohne Positionspreise / Stunden
    scope_groups: List[Dict[str, Any]] = []
    if license_offer:
        lic_items = []
        for line in license_offer.get("lines") or []:
            lic_items.append(
                {
                    "name": line.get("name") or "",
                    "description": line.get("description") or "",
                }
            )
        if (license_totals or {}).get("discountPercent"):
            lic_items.append(
                {
                    "name": f"Mengenrabatt SLL ({license_totals.get('discountPercent')}%)",
                    "description": (license_offer.get("configuration") or {}).get("discountLabel") or "",
                }
            )
        scope_groups.append(
            {
                "id": "license",
                "title": "A · Softwarelizenzen",
                "items": lic_items,
                "total": lic_sell_chf,
                "currency": (license_totals or {}).get("currency", "CHF"),
            }
        )

    if it_offer:
        it_cfg = it_offer.get("configuration") or {}
        it_items = [
            {
                "name": f"Anzahl Geräte: {it_cfg.get('deviceCount', 0)}",
                "description": "Geräte, die installiert und in Betrieb genommen werden.",
            },
            {
                "name": f"Anzahl Zonen: {it_cfg.get('zoneCount', 0)}",
                "description": "Zonen mit LOGIMAT / WAMAS Lift & Store.",
            },
            {
                "name": f"Anzahl Öffnungen: {it_cfg.get('openingCount', 0)}",
                "description": "Gesamtanzahl Bedienöffnungen.",
            },
        ]
        # Gewählte IT-Optionen mit Beschreibung (ohne Stunden/Preise)
        it_cat = load_it_catalog()
        option_defs = {o["id"]: o for o in it_cat.get("options") or []}
        selected = it_cfg.get("options") or {}
        for oid, enabled in selected.items():
            if not enabled:
                continue
            opt = option_defs.get(oid)
            if not opt:
                continue
            if oid == "testSystem" and not selected.get("orderHandling"):
                continue
            it_items.append(
                {
                    "name": opt.get("name") or oid,
                    "description": opt.get("description") or "",
                }
            )
        for ext in it_cfg.get("customExtensions") or []:
            if float(ext.get("hours") or 0) <= 0 and not (ext.get("description") or "").strip():
                continue
            desc = (ext.get("description") or "").strip() or "Projektspezifische Erweiterung"
            it_items.append({"name": desc, "description": "Projektspezifische Erweiterung."})
        if float((it_totals or {}).get("travelAmount") or 0) > 0:
            it_items.append(
                {
                    "name": "Reisekosten",
                    "description": "Fahrten, Kilometer, Übernachtung und Verpflegung gemäss Projektplanung.",
                }
            )
        scope_groups.append(
            {
                "id": "it",
                "title": "B · IT-Aufwand / Installation",
                "items": it_items,
                "total": it_total_chf,
                "currency": (it_totals or {}).get("currency", "CHF"),
            }
        )

    # Rabatt / Abrundung auf Gesamttotal
    adj_notes: List[str] = []
    discount_percent = float(payload.discountPercent or 0)
    discount_amount_chf = float(payload.discountAmountChf or 0)
    round_to = int(payload.roundTo or 0)
    commercial_discount = 0.0
    grand_before_round = subtotal_chf
    grand_chf = subtotal_chf
    if grand_chf is not None:
        if discount_percent > 0:
            pct_disc = round(grand_chf * (discount_percent / 100.0), 2)
            commercial_discount += pct_disc
            adj_notes.append(f"Projektrabatt {discount_percent:g}%")
        if discount_amount_chf > 0:
            commercial_discount += discount_amount_chf
            adj_notes.append(f"Rabatt CHF {discount_amount_chf:,.2f}".replace(",", "'"))
        commercial_discount = round(min(commercial_discount, grand_chf), 2)
        grand_before_round = round(grand_chf - commercial_discount, 2)
        grand_chf = grand_before_round
        if round_to and round_to > 0:
            # Abrunden auf Vielfaches (kundenseitig „glatter“ Preis)
            grand_chf = float(int(grand_chf // round_to) * round_to)
            adj_notes.append(f"Abgerundet auf {round_to} CHF")

    rounding_amount = None
    if grand_before_round is not None and grand_chf is not None:
        rounding_amount = round(grand_before_round - grand_chf, 2)

    price_summary = {
        "license": {
            "label": "Softwarelizenzen Verkauf",
            "currency": (license_totals or {}).get("currency", "CHF"),
            "icCurrency": (license_totals or {}).get("icCurrency", "EUR"),
            "subtotal": (license_totals or {}).get("subtotal"),
            "discountPercent": (license_totals or {}).get("discountPercent"),
            "discountAmount": (license_totals or {}).get("discountAmount"),
            "icNet": (license_totals or {}).get("net"),
            "marginPercent": (license_totals or {}).get("marginPercent"),
            "marginAmountEur": (license_totals or {}).get("marginAmountEur"),
            "sellNetEur": (license_totals or {}).get("sellNetEur"),
            "eurToChfRate": (license_totals or {}).get("eurToChfRate"),
            "total": lic_sell_chf,
            "sllCount": (license_totals or {}).get("sllCount"),
        }
        if license_totals
        else None,
        "it": {
            "label": "IT-Aufwand / Installation",
            "currency": (it_totals or {}).get("currency", "CHF"),
            "workHours": (it_totals or {}).get("workHours"),
            "workAmount": (it_totals or {}).get("workAmount"),
            "travelAmount": (it_totals or {}).get("travelAmount"),
            "marginPercent": (it_totals or {}).get("marginPercent"),
            "total": it_total_chf,
        }
        if it_totals
        else None,
        "subtotalChf": subtotal_chf,
        "commercialDiscountChf": commercial_discount if commercial_discount else None,
        "discountPercent": discount_percent if discount_percent else None,
        "discountAmountChf": discount_amount_chf if discount_amount_chf else None,
        "roundTo": round_to if round_to else None,
        "roundingAmountChf": rounding_amount if rounding_amount else None,
        "adjustmentNotes": adj_notes,
        "grandTotalChf": grand_chf,
        "note": (
            "Alle Preise in CHF, exkl. MwSt. "
            "Leistungsumfang ohne Einzelpreise; Gesamttotal gemäss Aufstellung."
            + ((" · " + " · ".join(adj_notes)) if adj_notes else "")
        ),
    }

    created = datetime.now()
    commercial_terms = load_commercial_terms()
    validity_days = int(commercial_terms.get("validityDays") or 14)
    intro_variant = (
        template.get("introOrderHandling")
        if has_order_handling
        else template.get("introStandalone")
    )
    doc = {
        "kind": "offer_document",
        "meta": {
            "offerNumber": f"ANG-{created.strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}",
            "createdAt": created.isoformat(timespec="seconds"),
            "validUntil": (created + timedelta(days=validity_days)).strftime("%d.%m.%Y"),
            "validityDays": validity_days,
            "preparedBy": prepared_by,
            "templateSource": "Anhang zu Software_v2.6.X.docx · Kaufmännische Bedingungen CH 09.2022",
            "documentDate": created.strftime("%d.%m.%Y"),
        },
        "customer": customer,
        "branding": {
            "product": "WAMAS Lift & Store",
            "vendor": "SSI SCHÄFER",
            "version": "2.8",
            "softwareVersionLabel": template.get(
                "softwareVersionLabel", "WAMAS® Lift & Store Software Version 2.8"
            ),
        },
        "content": {
            "documentLabel": "Angebot / Preisliste",
            "title": template["title"],
            "subtitle": template["subtitle"],
            "intro": template["intro"],
            "introVariant": intro_variant or "",
            "standardLead": template.get("standardLead", ""),
            "recommendation": template["recommendation"],
            "footnotes": template.get("footnotes", []),
            "annexLabel": template.get("documentLabel", "Anhang zur Software"),
            "configurationSummary": {
                "instanceName": instance_name,
                "instanceCount": (license_offer or {}).get("configuration", {}).get("instanceCount"),
                "deviceCount": (it_offer or {}).get("configuration", {}).get("deviceCount"),
                "zoneCount": (it_offer or {}).get("configuration", {}).get("zoneCount"),
                "openingCount": (it_offer or {}).get("configuration", {}).get("openingCount"),
                "hasOrderHandling": has_order_handling,
            },
            "standardFunctions": standard,
            "machineOptionsLead": template.get("machineOptionsLead", ""),
            "selectedOptions": options,
            "hardwareOptions": template.get("hardwareOptions", []),
            "clients": template["clients"],
            "architecture": template.get("architecture", {}),
            "requirements": template.get("requirements", {}),
            "acceptance": template.get("acceptance", ""),
            "responsibilities": template["responsibilities"],
            "documentsLead": template.get("documentsLead", ""),
            "documents": template["documents"],
            "closing": template["closing"],
            "itOfferSections": (it_offer or {}).get("offerSections") or [],
            "commercialTerms": commercial_terms,
            "scopeGroups": scope_groups,
        },
        "commercialLines": commercial,
        "scopeGroups": scope_groups,
        "priceSummary": price_summary,
        "totals": {
            "license": license_totals,
            "it": it_totals,
        },
        "sources": {
            "licenseOfferNumber": (license_offer or {}).get("meta", {}).get("offerNumber"),
            "itOfferNumber": (it_offer or {}).get("meta", {}).get("offerNumber"),
            "basedOnOfferNumber": payload.basedOnOfferNumber,
        },
        # Vollständige Quellkalkulationen für späteres Wiederöffnen / Bearbeiten
        "editable": {
            "license": license_offer,
            "it": it_offer,
        },
    }
    if payload.basedOnOfferNumber:
        doc["meta"]["revisionOf"] = payload.basedOnOfferNumber
    return doc


@app.post("/api/offer/compose/save")
def save_composed_offer(payload: ComposeOfferRequest):
    """Speichert das zusammengesetzte Angebotsdokument (alle Preise) im Archiv."""
    doc = compose_offer(payload)
    offer_id = doc["meta"]["offerNumber"]
    doc["id"] = offer_id
    if payload.basedOnOfferNumber:
        doc.setdefault("sources", {})["basedOnOfferNumber"] = payload.basedOnOfferNumber
        doc.setdefault("meta", {})["revisionOf"] = payload.basedOnOfferNumber
    offer_path(offer_id).write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    return doc


@app.post("/api/offers/calculate")
def api_calculate(payload: OfferRequest):
    return calculate_offer(payload)


@app.post("/api/offers")
def api_save_offer(payload: OfferRequest):
    offer = calculate_offer(payload)
    offer_id = offer["meta"]["offerNumber"]
    offer["id"] = offer_id
    offer_path(offer_id).write_text(json.dumps(offer, ensure_ascii=False, indent=2), encoding="utf-8")
    return offer


@app.post("/api/it/calculate")
def api_it_calculate(payload: ItOfferRequest):
    return calculate_it_offer(payload)


@app.post("/api/it/offers")
def api_it_save(payload: ItOfferRequest):
    offer = calculate_it_offer(payload)
    offer_id = offer["meta"]["offerNumber"]
    offer["id"] = offer_id
    offer_path(offer_id).write_text(json.dumps(offer, ensure_ascii=False, indent=2), encoding="utf-8")
    return offer


@app.get("/api/offers")
def api_list_offers():
    items = []
    for path in sorted(OFFERS_DIR.glob("*.json"), reverse=True):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            kind = data.get("kind") or ("it" if str(data.get("id", "")).startswith("IT-") else "license")
            totals = data.get("totals", {})
            cfg = data.get("configuration") or data.get("content", {}).get("configurationSummary") or {}
            summary = cfg.get("instanceName") or (
                f"{cfg.get('deviceCount', '')} Geräte" if cfg.get("deviceCount") else ""
            )
            if kind == "offer_document":
                lic = (totals.get("license") or {}).get("sellNetChf")
                if lic is None:
                    lic = (totals.get("license") or {}).get("net")
                it_amt = (totals.get("it") or {}).get("totalAmount")
                amount_parts = []
                if lic is not None:
                    amount_parts.append(f"{lic} CHF")
                if it_amt is not None:
                    amount_parts.append(f"{it_amt} CHF")
                grand = data.get("priceSummary", {}).get("grandTotalChf")
                amount_display = (
                    f"{grand} CHF"
                    if grand is not None
                    else (" + ".join(amount_parts) if amount_parts else None)
                )
                items.append(
                    {
                        "id": data.get("id") or path.stem,
                        "kind": kind,
                        "offerNumber": data.get("meta", {}).get("offerNumber", path.stem),
                        "company": data.get("customer", {}).get("company", ""),
                        "projectName": data.get("customer", {}).get("projectName", ""),
                        "summary": summary or "Gesamtangebot",
                        "amount": amount_display,
                        "currency": "",
                        "createdAt": data.get("meta", {}).get("createdAt"),
                    }
                )
            else:
                items.append(
                    {
                        "id": data.get("id") or path.stem,
                        "kind": kind,
                        "offerNumber": data.get("meta", {}).get("offerNumber", path.stem),
                        "company": data.get("customer", {}).get("company", ""),
                        "projectName": data.get("customer", {}).get("projectName", ""),
                        "summary": summary,
                        "amount": (
                            totals.get("sellNetChf")
                            if kind == "license"
                            else totals.get("totalAmount")
                        ),
                        "currency": totals.get("currency", "CHF" if kind == "license" else "EUR"),
                        "createdAt": data.get("meta", {}).get("createdAt"),
                    }
                )
        except Exception:
            continue
    return {"offers": items}


@app.get("/api/offers/{offer_id}")
def api_get_offer(offer_id: str):
    path = offer_path(offer_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Angebot nicht gefunden")
    return json.loads(path.read_text(encoding="utf-8"))


@app.delete("/api/offers/{offer_id}")
def api_delete_offer(offer_id: str):
    path = offer_path(offer_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Angebot nicht gefunden")
    path.unlink()
    return {"ok": True}


@app.get("/api/offers/{offer_id}/excel")
def api_export_excel(offer_id: str):
    path = offer_path(offer_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Angebot nicht gefunden")
    offer = json.loads(path.read_text(encoding="utf-8"))
    wb = Workbook()
    ws = wb.active
    ws.title = "Preise"
    kind = offer.get("kind", "license")
    meta = offer["meta"]
    customer = offer["customer"]
    totals = offer.get("totals") or {}
    cfg = offer.get("configuration") or offer.get("content", {}).get("configurationSummary") or {}

    if kind == "offer_document":
        summary = offer.get("priceSummary") or {}
        rows = [
            ["WAMAS Lift & Store – Gesamtangebot / Preisliste"],
            ["Nummer", meta.get("offerNumber")],
            ["Datum", meta.get("documentDate")],
            ["Kunde", customer.get("company")],
            ["Projekt", customer.get("projectName")],
            ["Konfiguration", cfg.get("instanceName")],
            ["Geräte / Zonen / Öffnungen", f'{cfg.get("deviceCount")} / {cfg.get("zoneCount")} / {cfg.get("openingCount")}'],
            [],
            ["Pos", "Bereich", "SKU", "Bezeichnung", "Beschreibung", "Menge", "Einzelpreis", "Stunden", "Betrag", "Währung"],
        ]
        for line in offer.get("commercialLines") or []:
            rows.append(
                [
                    line.get("pos"),
                    line.get("section"),
                    line.get("sku"),
                    line.get("name"),
                    line.get("description"),
                    line.get("qty"),
                    line.get("unitPrice"),
                    line.get("hours"),
                    line.get("amount"),
                    line.get("currency"),
                ]
            )
        lic = summary.get("license") or {}
        it = summary.get("it") or {}
        rows.extend(
            [
                [],
                ["Zusammenfassung"],
                ["Softwarelizenzen Verkauf CHF", lic.get("total"), lic.get("currency")],
                ["IT-Aufwand Total", it.get("total"), it.get("currency")],
                ["Gesamttotal CHF", summary.get("grandTotalChf"), "CHF"],
                ["Hinweis", summary.get("note")],
            ]
        )
    elif kind == "it":
        rows = [
            ["IT-Kalkulation WAMAS Lift & Store"],
            ["Nummer", meta.get("offerNumber")],
            ["Kunde", customer.get("company")],
            ["Projekt", customer.get("projectName")],
            ["Realisierungszeitraum", cfg.get("realizationPeriod")],
            ["Geräte / Zonen / Öffnungen", f'{cfg.get("deviceCount")} / {cfg.get("zoneCount")} / {cfg.get("openingCount")}'],
            ["Stundensatz", cfg.get("hourlyRate")],
            [],
            ["SKU", "Bezeichnung", "Beschreibung", "Menge", "Stunden", "Betrag"],
        ]
        for line in offer["lines"]:
            rows.append(
                [
                    line.get("sku"),
                    line.get("name"),
                    line.get("description"),
                    line.get("qty"),
                    line.get("hours"),
                    line.get("amount"),
                ]
            )
        rows.extend(
            [
                [],
                ["IT-Aufwand Stunden", totals.get("workHours")],
                ["IT-Aufwand Einkauf CHF", totals.get("workAmountCost")],
                ["Reisekosten Einkauf CHF", totals.get("travelAmountCost")],
                ["Marge %", totals.get("marginPercent")],
                ["Marge CHF", totals.get("marginAmount")],
                ["Deckungsbeitrag CHF / DB", totals.get("contributionMarginChf")],
                ["DB % vom Verkauf", totals.get("contributionMarginPercent")],
                ["Verkauf Total CHF", totals.get("totalAmount")],
            ]
        )
    else:
        rows = [
            ["WAMAS Lift & Store – License Calculator (IC)"],
            ["Nummer", meta.get("offerNumber")],
            ["Kunde", customer.get("company")],
            ["Instance", cfg.get("instanceName")],
            ["SLL", totals.get("sllCount")],
            ["Marge %", totals.get("marginPercent")],
            ["Kurs EUR→CHF", totals.get("eurToChfRate")],
            [],
            ["SKU", "Bezeichnung", "Menge", "IC EUR", "Verkauf CHF", "SLL"],
        ]
        for line in offer["lines"]:
            rows.append(
                [
                    line.get("sku"),
                    line.get("name"),
                    line.get("qty"),
                    line.get("totalIcEur"),
                    line.get("total"),
                    line.get("sllUnits", 0),
                ]
            )
        rows.extend(
            [
                [],
                ["IC Zwischensumme EUR", totals.get("subtotal")],
                ["IC Rabatt EUR", totals.get("discountAmount")],
                ["IC Total EUR", totals.get("net")],
                ["Marge EUR", totals.get("marginAmountEur")],
                ["Verkauf EUR", totals.get("sellNetEur")],
                ["Einkauf CHF (IC×Kurs)", totals.get("costChf")],
                ["Verkauf CHF", totals.get("sellNetChf")],
                ["Deckungsbeitrag EUR", totals.get("contributionMarginEur")],
                ["Deckungsbeitrag CHF / DB", totals.get("contributionMarginChf")],
                ["DB % vom Verkauf", totals.get("contributionMarginPercent")],
            ]
        )

    for row in rows:
        ws.append(row)
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    filename = f"{meta.get('offerNumber', offer_id)}.xlsx"
    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/api/offers/{offer_id}/docx")
def api_export_docx(offer_id: str):
    """Word-Export: Angebotsseiten + originaler Software-Anhang (DOCX-Vorlage)."""
    path = offer_path(offer_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Angebot nicht gefunden")
    offer = json.loads(path.read_text(encoding="utf-8"))
    if offer.get("kind") != "offer_document":
        raise HTTPException(
            status_code=400,
            detail="Word-Export nur für Gesamtangebote. Bitte zuerst „Angebot erzeugen“.",
        )
    try:
        data = build_offer_docx(offer)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Word-Export fehlgeschlagen: {exc}") from exc

    filename = f"{offer.get('meta', {}).get('offerNumber', offer_id)}.docx"
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/")
def root():
    return FileResponse(STATIC_DIR / "index.html")


app.mount("/static", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("server:app", host="0.0.0.0", port=8100, reload=False)
