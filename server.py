"""WAMAS Lift & Store – Angebot Generator"""

from __future__ import annotations

import json
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from pydantic import BaseModel, Field
import io

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
DATA_DIR = BASE_DIR / "data"
OFFERS_DIR = DATA_DIR / "offers"
CATALOG_FILE = DATA_DIR / "catalog.json"

OFFERS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="WAMAS Lift & Store Angebot Generator", version="1.0.0")
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
    packageId: str
    liftCount: int = Field(1, ge=1, le=50)
    selectedFeatures: List[str] = Field(default_factory=list)
    trainingDays: int = Field(0, ge=0, le=30)
    includeImplementation: bool = True
    includeMaintenance: bool = True
    discountId: str = "none"
    notes: str = ""
    preparedBy: str = ""


def calculate_offer(payload: OfferRequest) -> Dict[str, Any]:
    catalog = load_catalog()
    packages = {p["id"]: p for p in catalog["basePackages"]}
    features = {f["id"]: f for f in catalog["features"]}
    addons = {a["id"]: a for a in catalog["addons"]}
    discounts = {d["id"]: d for d in catalog["discounts"]}

    if payload.packageId not in packages:
        raise HTTPException(status_code=400, detail="Unbekanntes Paket")
    if payload.discountId not in discounts:
        raise HTTPException(status_code=400, detail="Unbekannter Rabatt")

    package = packages[payload.packageId]
    included = set(package.get("features", []))
    lines: List[Dict[str, Any]] = []

    lines.append(
        {
            "sku": f"PKG-{package['id'].upper()}",
            "name": f"WAMAS Lift & Store – Paket {package['name']}",
            "description": package["description"],
            "qty": 1,
            "unitPrice": package["basePrice"],
            "total": package["basePrice"],
            "category": "package",
        }
    )

    extra_lifts = max(0, payload.liftCount - int(package["includedLiftLicenses"]))
    if extra_lifts:
        unit = addons["extra_lift_license"]["unitPrice"]
        lines.append(
            {
                "sku": "ADD-LIFT",
                "name": addons["extra_lift_license"]["name"],
                "description": addons["extra_lift_license"]["description"],
                "qty": extra_lifts,
                "unitPrice": unit,
                "total": extra_lifts * unit,
                "category": "addon",
            }
        )

    # Always include package features; allow extra selected features
    chosen = set(payload.selectedFeatures) | included
    for fid in sorted(chosen):
        feat = features.get(fid)
        if not feat:
            continue
        # Package-included features shown at 0 if they were in package; otherwise list price
        price = 0 if fid in included else feat["price"]
        if price == 0 and fid not in included:
            continue
        lines.append(
            {
                "sku": f"FEAT-{fid.upper()}",
                "name": feat["name"],
                "description": feat["description"],
                "qty": 1,
                "unitPrice": price,
                "total": price,
                "category": "feature",
                "includedInPackage": fid in included,
            }
        )

    if payload.includeImplementation:
        impl = addons["implementation"]
        lines.append(
            {
                "sku": "SVC-IMPL",
                "name": impl["name"],
                "description": impl["description"],
                "qty": 1,
                "unitPrice": impl["unitPrice"],
                "total": impl["unitPrice"],
                "category": "service",
            }
        )

    if payload.trainingDays > 0:
        train = addons["training"]
        total = payload.trainingDays * train["unitPrice"]
        lines.append(
            {
                "sku": "SVC-TRAIN",
                "name": train["name"],
                "description": train["description"],
                "qty": payload.trainingDays,
                "unitPrice": train["unitPrice"],
                "total": total,
                "category": "service",
            }
        )

    software_subtotal = sum(
        l["total"] for l in lines if l["category"] in {"package", "addon", "feature"}
    )

    if payload.includeMaintenance:
        maint = addons["maintenance"]
        yearly = round(software_subtotal * (maint["unitPrice"] / 100), 2)
        lines.append(
            {
                "sku": "SVC-MAINT",
                "name": maint["name"],
                "description": f"{maint['description']} ({maint['unitPrice']}% vom Softwarewert)",
                "qty": 1,
                "unitPrice": yearly,
                "total": yearly,
                "category": "service",
                "recurring": True,
            }
        )

    subtotal = round(sum(l["total"] for l in lines), 2)
    discount = discounts[payload.discountId]
    discount_amount = round(subtotal * (discount["percent"] / 100), 2)
    net = round(subtotal - discount_amount, 2)
    vat_rate = catalog["product"]["vatRate"]
    vat = round(net * vat_rate, 2)
    gross = round(net + vat, 2)

    created = datetime.now()
    valid_until = created + timedelta(days=int(catalog["product"]["validityDays"]))

    return {
        "product": catalog["product"],
        "customer": payload.customer.model_dump(),
        "configuration": {
            "packageId": payload.packageId,
            "packageName": package["name"],
            "liftCount": payload.liftCount,
            "includedLiftLicenses": package["includedLiftLicenses"],
            "selectedFeatures": sorted(chosen),
            "trainingDays": payload.trainingDays,
            "includeImplementation": payload.includeImplementation,
            "includeMaintenance": payload.includeMaintenance,
            "discountId": payload.discountId,
            "discountName": discount["name"],
            "discountPercent": discount["percent"],
            "notes": payload.notes,
            "preparedBy": payload.preparedBy,
        },
        "lines": lines,
        "totals": {
            "softwareSubtotal": software_subtotal,
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
        },
    }


@app.get("/api/health")
def health():
    return {"ok": True, "service": "WAMAS Lift & Store Angebot Generator"}


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
    path = offer_path(offer_id)
    path.write_text(json.dumps(offer, ensure_ascii=False, indent=2), encoding="utf-8")
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
                    "packageName": data.get("configuration", {}).get("packageName", ""),
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
    ws.title = "Angebot"

    meta = offer["meta"]
    customer = offer["customer"]
    totals = offer["totals"]
    cfg = offer["configuration"]

    rows = [
        ["WAMAS Lift & Store – Angebot"],
        ["Angebotsnummer", meta.get("offerNumber")],
        ["Erstellt am", meta.get("createdAt")],
        ["Gültig bis", meta.get("validUntil")],
        ["Erstellt von", cfg.get("preparedBy") or ""],
        [],
        ["Kunde", customer.get("company")],
        ["Ansprechpartner", customer.get("contact")],
        ["E-Mail", customer.get("email")],
        ["Telefon", customer.get("phone")],
        ["Adresse", customer.get("address")],
        ["Projekt", customer.get("projectName")],
        [],
        ["Paket", cfg.get("packageName")],
        ["Anzahl Lifte", cfg.get("liftCount")],
        ["Rabatt", f'{cfg.get("discountName")} ({cfg.get("discountPercent")}%)'],
        [],
        ["Pos.", "SKU", "Bezeichnung", "Beschreibung", "Menge", "Einzelpreis", "Summe"],
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
            ]
        )

    rows.extend(
        [
            [],
            ["Zwischensumme", totals.get("subtotal")],
            ["Rabatt", totals.get("discountAmount")],
            ["Netto", totals.get("net")],
            [f'MwSt. ({int(totals.get("vatRate", 0) * 100)}%)', totals.get("vat")],
            ["Brutto", totals.get("gross")],
            [],
            ["Hinweis", offer["product"].get("disclaimer")],
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
