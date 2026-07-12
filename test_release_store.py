import hashlib
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "license_server"))

from release_store import ReleaseBundleError, channel_release_status, install_release_bundle


class ReleaseStoreTests(unittest.TestCase):
    def make_bundle(self, root, version="0.4.0", bad_sha=False, traversal=False):
        package_id = "ProductCodeFormatter.App.Test"
        channel = "test"
        package_name = f"{package_id}-{version}-{channel}-full.nupkg"
        setup_name = f"{package_id}-{channel}-Setup.exe"
        package_bytes = f"package-{version}".encode("utf-8")
        sha256 = hashlib.sha256(package_bytes).hexdigest().upper()
        if bad_sha:
            sha256 = "0" * 64
        asset = {
            "PackageId": package_id,
            "Version": version,
            "Type": "Full",
            "FileName": package_name,
            "SHA1": hashlib.sha1(package_bytes).hexdigest().upper(),
            "SHA256": sha256,
            "Size": len(package_bytes),
            "NotesMarkdown": f"Release {version}",
            "NotesHTML": f"<p>Release {version}</p>",
        }
        bundle = Path(root) / f"bundle-{version}.zip"
        with zipfile.ZipFile(bundle, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("releases.test.json", json.dumps({"Assets": [asset]}))
            archive.writestr("assets.test.json", json.dumps([]))
            archive.writestr("RELEASES-test", "fixture")
            archive.writestr(package_name, package_bytes)
            archive.writestr(setup_name, b"setup")
            if traversal:
                archive.writestr("../escape.txt", b"blocked")
        return bundle

    def test_valid_bundle_is_published_and_visible(self):
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            bundle = self.make_bundle(root)
            release = install_release_bundle(bundle, root / "server")
            self.assertEqual("test", release["channel"])
            self.assertEqual("0.4.0", release["version"])
            status = channel_release_status(root / "server", "test")
            self.assertTrue(status["available"])
            self.assertTrue(status["setup_available"])
            self.assertEqual("0.4.0", status["version"])

    def test_bad_hash_does_not_replace_existing_feed(self):
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            server = root / "server"
            install_release_bundle(self.make_bundle(root, "0.4.0"), server)
            with self.assertRaisesRegex(ReleaseBundleError, "SHA-256"):
                install_release_bundle(self.make_bundle(root, "0.4.1", bad_sha=True), server)
            self.assertEqual("0.4.0", channel_release_status(server, "test")["version"])

    def test_path_traversal_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            with self.assertRaisesRegex(ReleaseBundleError, "đường dẫn không hợp lệ"):
                install_release_bundle(self.make_bundle(root, traversal=True), root / "server")
            self.assertFalse((root / "escape.txt").exists())

    def test_corrupt_zip_is_reported(self):
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            bundle = root / "broken.zip"
            bundle.write_bytes(b"not-a-zip")
            with self.assertRaisesRegex(ReleaseBundleError, "Không đọc"):
                install_release_bundle(bundle, root / "server")


if __name__ == "__main__":
    unittest.main()