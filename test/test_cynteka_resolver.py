import unittest

from connectors.cynteka_resolver import normalize_request, normalize_result


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
