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


if __name__ == "__main__":
    unittest.main()
