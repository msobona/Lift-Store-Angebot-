"""WAMAS Lift & Store – License Calculator (IC Price List)"""

from __future__ import annotations

import io
import json
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from pydantic import BaseModel, Field

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
DATA_DIR = BASE_DIR / "data"
OFFERS_DIR = DATA_DIR / "offers"
CATALOG_FILE = DATA_DIR / "catalog.json"

OFFERS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="WAMAS Lift & Store License Calculator", version="2.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def load_catalog() -> Dict[str, Any]:
    return json.loads(CATALOG_FILE.read_text(encoding="utf-8"))


def offer_path(offer_id: str) -> Path:
    safe = "".join(c for c in offer_id if c.isalnum() or c in "-_")
    return OFFERS_DIR / f"{safe}.json"


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

    if payload.instanceId not in instances:
        raise HTTPException(status_code=400, detail="Unbekannte Instance")

    instance = instances[payload.instanceId]
    lines: List[Dict[str, Any]] = []
    sll = 0

    instance_total = instance["price"] * payload.instanceCount
    lines.append(
        {
            "sku": f"INST-{instance['id'].upper()}",
            "name": instance["name"],
            "description": instance["description"],
            "qty": payload.instanceCount,
            "unitPrice": instance["price"],
            "total": instance_total,
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
        total = addon["price"] * qty
        units = int(addon.get("sllUnits", 1)) * qty
        lines.append(
            {
                "sku": f"ADD-{addon_id.upper()}",
                "name": addon["name"],
                "description": f"Add-on License (one-time, je Instance)",
                "qty": qty,
                "unitPrice": addon["price"],
                "total": total,
                "category": "addon",
                "sllUnits": units,
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
        total = client["price"] * qty
        units = int(client.get("sllUnitsPerQty", 1)) * qty
        lines.append(
            {
                "sku": f"CLI-{cid.upper()}",
                "name": client["name"],
                "description": client["description"],
                "qty": qty,
                "unitPrice": client["price"],
                "total": total,
                "category": "client",
                "sllUnits": units,
            }
        )
        sll += units

    if payload.testInstances > 0:
        test = misc["test_instance"]
        total = test["price"] * payload.testInstances
        units = int(test.get("sllUnitsPerQty", 1)) * payload.testInstances
        lines.append(
            {
                "sku": "MISC-TEST",
                "name": test["name"],
                "description": test["description"],
                "qty": payload.testInstances,
                "unitPrice": test["price"],
                "total": total,
                "category": "misc",
                "sllUnits": units,
            }
        )
        sll += units

    if payload.upgradeYears > 0:
        upgrade = misc["upgrade_fee"]
        total = upgrade["price"] * payload.upgradeYears
        lines.append(
            {
                "sku": "MISC-UPGRADE",
                "name": upgrade["name"],
                "description": upgrade["description"],
                "qty": payload.upgradeYears,
                "unitPrice": upgrade["price"],
                "total": total,
                "category": "misc",
                "sllUnits": 0,
                "recurring": True,
            }
        )

    subtotal = round(sum(l["total"] for l in lines), 2)
    discount = resolve_discount(catalog, sll)
    discount_amount = round(subtotal * (discount["percent"] / 100), 2)
    net = round(subtotal - discount_amount, 2)
    vat_rate = float(catalog["product"].get("vatRate") or 0)
    vat = round(net * vat_rate, 2)
    gross = round(net + vat, 2)

    created = datetime.now()
    valid_until = created + timedelta(days=int(catalog["product"]["validityDays"]))

    included_opening = instance["includedOpeningClients"] * payload.instanceCount
    included_admin = instance["includedAdminClients"] * payload.instanceCount

    scope = [
        f"{payload.instanceCount}× {instance['name']} (IC)",
        f"Inklusive Clients: {included_opening} Opening, {included_admin} Admin",
        "Enthaltene Funktionen: " + ", ".join(instance["includedFunctions"]),
    ]
    if payload.selectedAddons:
        names = [addons[a]["name"] for a in payload.selectedAddons if a in addons]
        scope.append("Add-ons: " + ", ".join(names))
    if payload.extraOpeningClients:
        scope.append(f"Zusätzliche Opening Clients: {payload.extraOpeningClients}")
    if payload.extraAdminClients:
        scope.append(f"Zusätzliche Admin Clients: {payload.extraAdminClients}")
    if payload.mobileTerminalClients:
        scope.append(f"Mobile Terminal Clients: {payload.mobileTerminalClients}")
    if payload.thirdPartyVlmTypes:
        scope.append(f"3rd Party VLM Types: {payload.thirdPartyVlmTypes}")
    if payload.testInstances:
        scope.append(f"Test Instances: {payload.testInstances}")
    if payload.upgradeYears:
        scope.append(f"Upgrade Fee: {payload.upgradeYears} Jahr(e)")
    scope.append(f"SLL-Einheiten: {sll} → Rabatt {discount['percent']}% ({discount['label']})")

    return {
        "product": catalog["product"],
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
            "includedFunctions": instance["includedFunctions"],
            "notes": payload.notes,
            "preparedBy": payload.preparedBy,
            "sllCount": sll,
            "discountPercent": discount["percent"],
            "discountLabel": discount["label"],
        },
        "scopeOfSupply": scope,
        "lines": lines,
        "totals": {
            "sllCount": sll,
            "subtotal": subtotal,
            "discountPercent": discount["percent"],
            "discountAmount": discount_amount,
            "net": net,
            "vatRate": vat_rate,
            "vat": vat,
            "gross": gross,
            "currency": catalog["product"]["currency"],
        },
        "meta": {
            "createdAt": created.isoformat(timespec="seconds"),
            "validUntil": valid_until.date().isoformat(),
            "offerNumber": f"WLS-{created.strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}",
            "productVersion": catalog["product"].get("version", ""),
            "priceBasis": catalog["product"].get("priceBasis", ""),
        },
    }


@app.get("/api/health")
def health():
    return {"ok": True, "service": "WAMAS Lift & Store License Calculator"}


@app.get("/api/catalog")
def get_catalog():
    return load_catalog()


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


@app.get("/api/offers")
def api_list_offers():
    items = []
    for path in sorted(OFFERS_DIR.glob("*.json"), reverse=True):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            items.append(
                {
                    "id": data.get("id") or path.stem,
                    "offerNumber": data.get("meta", {}).get("offerNumber", path.stem),
                    "company": data.get("customer", {}).get("company", ""),
                    "projectName": data.get("customer", {}).get("projectName", ""),
                    "instanceName": data.get("configuration", {}).get("instanceName", ""),
                    "sllCount": data.get("totals", {}).get("sllCount"),
                    "gross": data.get("totals", {}).get("gross"),
                    "currency": data.get("totals", {}).get("currency", "EUR"),
                    "createdAt": data.get("meta", {}).get("createdAt"),
                    "validUntil": data.get("meta", {}).get("validUntil"),
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
    ws.title = "License Calc"

    meta = offer["meta"]
    customer = offer["customer"]
    totals = offer["totals"]
    cfg = offer["configuration"]

    rows = [
        ["WAMAS Lift & Store – License Calculator (IC)"],
        ["Angebotsnummer", meta.get("offerNumber")],
        ["Preisbasis", meta.get("priceBasis")],
        ["Erstellt am", meta.get("createdAt")],
        ["Gültig bis", meta.get("validUntil")],
        ["Erstellt von", cfg.get("preparedBy") or ""],
        [],
        ["Kunde", customer.get("company")],
        ["Projekt", customer.get("projectName")],
        ["Instance", cfg.get("instanceName")],
        ["Instance Count", cfg.get("instanceCount")],
        ["SLL", totals.get("sllCount")],
        ["Rabatt", f'{cfg.get("discountLabel")}'],
        [],
        ["Pos.", "SKU", "Bezeichnung", "Beschreibung", "Menge", "Einzelpreis", "Summe", "SLL"],
    ]
    for i, line in enumerate(offer["lines"], start=1):
        rows.append(
            [
                i,
                line.get("sku"),
                line.get("name"),
                line.get("description"),
                line.get("qty"),
                line.get("unitPrice"),
                line.get("total"),
                line.get("sllUnits", 0),
            ]
        )
    rows.extend(
        [
            [],
            ["Zwischensumme", totals.get("subtotal")],
            ["Rabatt", totals.get("discountAmount")],
            ["Netto / IC Total", totals.get("net")],
            [],
            ["Notizen", cfg.get("notes") or ""],
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


@app.get("/")
def root():
    return FileResponse(STATIC_DIR / "index.html")


app.mount("/static", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("server:app", host="0.0.0.0", port=8100, reload=False)
