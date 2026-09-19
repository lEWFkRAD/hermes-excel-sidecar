"""Read-only Cynteka resolver contract and validation helpers.

Network transport stays in the VPS adapter; this module never reads workbook
cells or performs writes.
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation
from urllib.parse import quote, urlencode, urlparse

ALLOWED_HOSTS = {"reformenginiring.cynteka.ru", "partner.cynteka.ru"}
SCOPES = {"cynteka.reformenginiring.read", "cynteka.partner.read"}
STATUSES = {"verified", "needs_review", "failed"}
QUERY_TYPES = {"company", "project", "request", "offer", "invoice", "payment", "unpaid_invoice", "material"}
TENANTS = {"reformenginiring", "partner"}
QUERY_FIELDS = {
    "company": {"id", "name", "shortName", "inn", "kpp", "ogrn", "email", "page", "pageSize"},
    "project": {"ids", "externalId", "page", "pageSize", "include", "excludedIds"},
    "request": {"state", "project", "creator", "responsible", "creatorCompany", "category", "search", "searchByWhat", "page", "batchSize", "totalCount", "creationDateFrom", "creationDateTo", "finishDateFrom", "finishDateTo", "hasWaiting", "hasWaitingWithoutOffer", "hasNotAccepted", "offerHasNotFullDelivery", "offerHasNotFullPayment"},
    "offer": {"ids", "search", "state", "supplier", "payer", "project", "orderId", "category", "page", "pageSize", "totalCount", "hasApprovedPayment", "hasUnapprovedPayment", "paymentAmountFrom", "paymentAmountTo", "acceptDateFrom", "acceptDateTo"},
    "invoice": {"ids", "search", "state", "supplier", "payer", "project", "orderId", "category", "page", "pageSize", "totalCount", "hasApprovedPayment", "hasUnapprovedPayment", "paymentAmountFrom", "paymentAmountTo", "acceptDateFrom", "acceptDateTo"},
    "payment": {"offerId", "orderId", "accepted", "page", "pageSize", "totalCount", "paymentDateFrom", "paymentDateTo", "planPaymentDateFrom", "planPaymentDateTo", "currency"},
    "unpaid_invoice": {"supplier", "payer", "project", "orderId", "state", "page", "pageSize", "include_partial"},
    "material": {"search", "page", "pageSize"},
}
PATHS = {
    "company": "/api/v1/company",
    "project": "/api/v1/projects",
    "request": "/api/v1/orders",
    "offer": "/api/v1/offers",
    "invoice": "/api/v1/offers",
    "payment": "/api/v1/payments",
    "material": "/api/v1/offers",
}


def normalize_query(payload: dict) -> dict:
    if not isinstance(payload, dict):
        raise ValueError("query payload is required")
    query_type = str(payload.get("query_type", "")).strip()
    tenant = str(payload.get("tenant", "")).strip()
    if query_type not in QUERY_TYPES:
        raise ValueError("unsupported Cynteka query_type")
    if tenant not in TENANTS:
        raise ValueError("unsupported Cynteka tenant")
    filters = payload.get("filters", {})
    if not isinstance(filters, dict):
        raise ValueError("filters must be an object")
    unknown = set(filters) - QUERY_FIELDS[query_type]
    if unknown:
        raise ValueError(f"unsupported filters: {', '.join(sorted(unknown))}")
    normalized = {}
    for key, value in filters.items():
        if value is None or value == "":
            continue
        if isinstance(value, (dict, list)) and key not in {"ids"}:
            raise ValueError(f"filter {key} must be scalar")
        if isinstance(value, str) and len(value) > 240:
            raise ValueError(f"filter {key} is too long")
        normalized[key] = value
    return {"query_type": query_type, "tenant": tenant, "path": PATHS.get(query_type), "filters": normalized, "credential_scope": str(payload.get("credential_scope", f"cynteka.{tenant}.read"))}


def build_query_url(base_url: str, query: dict) -> str:
    normalized = normalize_query(query)
    parsed = urlparse(str(base_url).rstrip("/"))
    if parsed.scheme != "https" or parsed.hostname not in ALLOWED_HOSTS or parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise ValueError("base_url must be an allowlisted HTTPS Cynteka URL")
    if normalized["query_type"] == "unpaid_invoice":
        raise ValueError("unpaid_invoice is a derived query and requires payment reconciliation")
    path = quote(normalized["path"], safe="/")
    return f"{parsed.scheme}://{parsed.netloc}{path}?{urlencode(normalized['filters'], doseq=True)}"


def payment_amounts(total_amount, payments: list[dict]) -> tuple[Decimal, Decimal]:
    try:
        total = Decimal(str(total_amount))
    except (InvalidOperation, TypeError, ValueError) as error:
        raise ValueError("total amount must be numeric") from error
    paid = Decimal("0")
    for payment in payments or []:
        if not isinstance(payment, dict) or payment.get("accepted") is False:
            continue
        try:
            paid += Decimal(str(payment.get("amount", payment.get("sum", payment.get("paymentAmount", "0")))))
        except (InvalidOperation, TypeError, ValueError) as error:
            raise ValueError("payment amount must be numeric") from error
    return total, paid


def classify_payment_state(total_amount, payments: list[dict]) -> str:
    total, paid = payment_amounts(total_amount, payments)
    if paid <= 0:
        return "UNPAID"
    if paid < total:
        return "PARTIALLY_PAID"
    return "PAID"


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
