import unittest

from inventory_allocation_app import license_client


class InventoryLicenseClientTest(unittest.TestCase):
    def test_product_identity_is_separate_from_product_code_formatter(self):
        self.assertEqual(license_client.PRODUCT_CODE, "inventory-allocator")
        self.assertNotEqual(license_client.PRODUCT_CODE, "product-code-formatter")
        cfg = license_client.default_config()
        self.assertEqual(cfg["product_code"], "inventory-allocator")
        self.assertIn("InventoryAllocator", str(license_client.CONFIG_PATH))
        self.assertNotIn("product_code_config.json", str(license_client.CONFIG_PATH))

    def test_metadata_requires_inventory_product_marker(self):
        self.assertTrue(license_client.metadata_matches_inventory_product({"product_code": "inventory-allocator"}))
        self.assertTrue(license_client.metadata_matches_inventory_product({"application": "InventoryAllocator"}))
        self.assertFalse(license_client.metadata_matches_inventory_product({"product_code": "product-code-formatter"}))
        self.assertFalse(license_client.metadata_matches_inventory_product({}))

    def test_license_server_accepts_plain_lan_ip(self):
        url = license_client.keygen_url("192.168.1.10:3000", "acct", "licenses/actions/validate-key")
        self.assertEqual(url, "http://192.168.1.10:3000/v1/accounts/acct/licenses/actions/validate-key")
        self.assertTrue(license_client.keygen_is_local_http_url(url))


if __name__ == "__main__":
    unittest.main()
