"""Read-only Cynteka resolver contract and validation helpers.

Network transport stays in the VPS adapter; this module never reads workbook
cells or performs writes.
"""
from __future__ import annotations

from urllib.parse import urlparse

ALLOWED_HOSTS = {"reformenginiring.cynteka.ru", "partner.cynteka.ru"}
SCOPES = {"cynteka.reformenginiring.read", "cynteka.partner.read"}
STATUSES = {"verified", "needs_review", "failed"}


def normalize_request(payload: dict) -> dict:
    if not isinstance(payload, dict) or not isinstance(payload.get("material"), dict):
        raise ValueError("material is required")
    material = payload["material"]
    name = str(material.get("name", "")).strip()
    if not name or len(name) > 240:
        raise ValueError("material.name is required")
    url = str(payload.get("url", ""))
    parsed = urlparse(url)
    if parsed.scheme != "https" or parsed.hostname not in ALLOWED_HOSTS or parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise ValueError("url must be an allowlisted credential-free HTTPS URL")
    source_cell = str(payload.get("source_cell", ""))
    scope = str(payload.get("credential_scope", ""))
    if not source_cell or scope not in SCOPES:
        raise ValueError("source_cell and a valid credential_scope are required")
    return {"material": {key: str(material[key]).strip() for key in ("name", "code", "article", "brand") if material.get(key) is not None}, "url": url, "source_cell": source_cell, "credential_scope": scope}


def normalize_result(payload: dict) -> dict:
    if not isinstance(payload, dict) or payload.get("status") not in STATUSES:
        raise ValueError("invalid resolver status")
    confidence = payload.get("confidence", 0.0)
    if not isinstance(confidence, (int, float)) or not 0 <= confidence <= 1:
        raise ValueError("confidence must be between 0 and 1")
    result = {"status": payload["status"], "source_url": str(payload.get("source_url", "")), "source_cell": str(payload.get("source_cell", "")), "matched_material": payload.get("matched_material") or {}, "price": payload.get("price"), "identifiers": payload.get("identifiers") or [], "evidence": payload.get("evidence") or [], "retrieved_at": str(payload.get("retrieved_at", "")), "confidence": float(confidence), "error": payload.get("error")}
    if result["status"] != "verified": result["price"] = None
    return result
