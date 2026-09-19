import unittest

from connectors.cynteka_resolver import (
    build_query_url,
    classify_payment_state,
    normalize_query,
    normalize_request,
    normalize_result,
)


class CyntekaResolverContractTests(unittest.TestCase):
    def test_request_requires_allowlisted_https_and_scope(self):
        request = normalize_request({"material": {"name": "Насос", "article": "A-1"}, "url": "https://reformenginiring.cynteka.ru/item/1", "source_cell": "Sheet1!Z10", "credential_scope": "cynteka.reformenginiring.read"})
        self.assertEqual(request["material"]["name"], "Насос")
        with self.assertRaises(ValueError):
            normalize_request({"material": {"name": "X"}, "url": "http://reformenginiring.cynteka.ru/item/1", "source_cell": "A1", "credential_scope": "cynteka.reformenginiring.read"})
        with self.assertRaises(ValueError):
            normalize_request({"material": {"name": "X"}, "url": "https://example.com/item", "source_cell": "A1", "credential_scope": "cynteka.reformenginiring.read"})

    def test_ambiguous_result_cannot_carry_a_write_price(self):
        result = normalize_result({"status": "needs_review", "source_url": "https://reformenginiring.cynteka.ru/item/1", "source_cell": "Z10", "price": {"value": 1}, "confidence": 0.4})
        self.assertIsNone(result["price"])
        with self.assertRaises(ValueError):
            normalize_result({"status": "verified", "confidence": 2})

    def test_query_rejects_unknown_filters_and_builds_official_path(self):
        query = normalize_query({"query_type": "request", "tenant": "reformenginiring", "filters": {"project": 42, "state": "ACTIVE", "page": 2, "batchSize": 100}})
        self.assertEqual(query["path"], "/api/v1/orders")
        self.assertEqual(query["filters"]["batchSize"], 100)
        with self.assertRaises(ValueError):
            normalize_query({"query_type": "request", "tenant": "reformenginiring", "filters": {"arbitraryUrl": "https://evil.example"}})

    def test_query_url_uses_batch_size_for_requests_and_preserves_encoding(self):
        url = build_query_url("https://reformenginiring.cynteka.ru", {"query_type": "request", "tenant": "reformenginiring", "filters": {"search": "Школа 1500", "batchSize": 100}})
        self.assertIn("/api/v1/orders?", url)
        self.assertIn("search=%D0%A8%D0%BA%D0%BE%D0%BB%D0%B0+1500", url)

    def test_payment_state_distinguishes_unpaid_and_partial(self):
        self.assertEqual(classify_payment_state("100.00", []), "UNPAID")
        self.assertEqual(classify_payment_state("100.00", [{"amount": "25.00"}]), "PARTIALLY_PAID")
        self.assertEqual(classify_payment_state("100.00", [{"amount": "100.00"}]), "PAID")
        with self.assertRaises(ValueError):
            classify_payment_state("bad", [])
