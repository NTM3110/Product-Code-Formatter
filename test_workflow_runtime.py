import json
import tempfile
import time
import unittest
from pathlib import Path

import app
from product_code import config_store
from product_code.workflow_runtime import WorkflowFailure, WorkflowJobManager, WorkflowSessionStore, stable_signature


class ConfigMigrationTests(unittest.TestCase):
    def test_legacy_vietmax_profiles_migrate_into_canonical_scopes(self):
        raw = {
            "selected_profile": "vietmax",
            "profiles": {
                "vietmax_mua_vao": {
                    "prefixes": {"0101": "AA"},
                    "company_group_assignments": {"0101": "services"},
                    "form_mapping_presets": [{"id": "custom", "label": "Dịch vụ", "scope": "purchase"}],
                },
                "vietmax_ban_ra": {
                    "manual_code_overrides": {"0102|||Hàng": "HANG"},
                },
            },
        }

        normalized = app.normalize_config(raw)

        purchase = normalized["profiles"]["vietmax"]["scopes"]["purchase"]
        sales = normalized["profiles"]["vietmax"]["scopes"]["sales"]
        self.assertEqual(purchase["prefixes"], {"0101": "AA"})
        self.assertEqual(purchase["company_group_assignments"], {"0101": "services"})
        self.assertEqual(purchase["form_mapping_presets"][0]["id"], "custom")
        self.assertEqual(sales["manual_code_overrides"], {"0102|||Hàng": "HANG"})

    def test_canonical_storage_drops_only_legacy_profile_copies(self):
        normalized = app.normalize_config({
            "selected_profile": "vietmax",
            "license": {"license_key": "keep-me"},
            "profiles": {
                "vietmax": {"scopes": {"purchase": {"prefixes": {"0101": "AA"}}}},
                "vietmax_mua_vao": {"prefixes": {"legacy": "L"}},
                "son_phuong": {"manual_code_overrides": {"0101|||Hàng A": "B"}},
            },
        })

        stored = app.canonical_config_for_storage(normalized)

        self.assertEqual(stored["config_schema_version"], 3)
        self.assertNotIn("vietmax_mua_vao", stored["profiles"])
        self.assertNotIn("vietmax_ban_ra", stored["profiles"])
        self.assertEqual(stored["profiles"]["vietmax"]["scopes"]["purchase"]["prefixes"], {"0101": "AA"})
        self.assertEqual(stored["profiles"]["son_phuong"]["manual_code_overrides"], {"0101|||Hàng A": "B"})
        self.assertNotIn("license", stored)

    def test_atomic_save_writes_canonical_schema(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "product_code_config.json"
            config_store.save_config_file(
                path,
                {"profiles": {"vietmax_mua_vao": {"prefixes": {"0101": "AA"}}}},
                app.normalize_config,
                serializer=app.canonical_config_for_storage,
            )
            stored = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(stored["config_schema_version"], 3)
            self.assertNotIn("vietmax_mua_vao", stored["profiles"])
            self.assertEqual(stored["profiles"]["vietmax"]["scopes"]["purchase"]["prefixes"], {"0101": "AA"})

    def test_routine_save_preserves_custom_form_missing_from_stale_ui(self):
        from web_api import keep_form_mapping_presets

        profile_cfg = {
            "form_mapping_presets": [
                {"id": "custom_service", "label": "Dịch vụ", "group_id": "services"},
                {"id": "default_purchase", "label": "Hóa đơn mua hàng", "group_id": "materials"},
            ],
        }
        payload = {
            "form_mapping_presets": [
                {"id": "default_purchase", "label": "Hóa đơn mua hàng mới", "group_id": "materials"},
            ],
        }

        merged = keep_form_mapping_presets(payload, profile_cfg)

        self.assertEqual([item["id"] for item in merged], ["default_purchase", "custom_service"])
        self.assertEqual(merged[0]["label"], "Hóa đơn mua hàng mới")

    def test_explicit_form_mapping_save_can_delete_a_form(self):
        from web_api import keep_form_mapping_presets

        profile_cfg = {
            "form_mapping_presets": [
                {"id": "custom_service", "label": "Dịch vụ"},
                {"id": "default_purchase", "label": "Hóa đơn mua hàng"},
            ],
        }
        payload = {
            "replace_form_mapping_presets": True,
            "form_mapping_presets": [
                {"id": "default_purchase", "label": "Hóa đơn mua hàng"},
            ],
        }

        replaced = keep_form_mapping_presets(payload, profile_cfg)

        self.assertEqual([item["id"] for item in replaced], ["default_purchase"])


class WorkflowRuntimeTests(unittest.TestCase):
    def test_session_keeps_artifacts_until_close(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            uploads = root / "uploads"
            uploads.mkdir()
            (uploads / "source.xls").write_bytes(b"source")
            store = WorkflowSessionStore(root / "sessions", uploads)
            artifact = store.register_file("source.xls", kind="source:vietmax-purchase", original_name="source.xls")

            self.assertTrue((uploads / "source.xls").exists())
            self.assertEqual(store.artifact(artifact["artifact_id"])["saved_name"], "source.xls")

            store.close()
            self.assertFalse((uploads / "source.xls").exists())
            self.assertFalse(store.session_dir.exists())

    def test_new_session_cleans_crash_leftovers(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            uploads = root / "uploads"
            uploads.mkdir()
            (uploads / "old.xls").write_bytes(b"old")
            sessions = root / "sessions"
            old_session = sessions / "old"
            old_session.mkdir(parents=True)
            (old_session / "manifest.json").write_text(json.dumps({
                "artifacts": {"one": {"saved_name": "old.xls"}},
            }), encoding="utf-8")

            store = WorkflowSessionStore(sessions, uploads)

            self.assertFalse((uploads / "old.xls").exists())
            self.assertFalse(old_session.exists())
            store.close()

    def test_jobs_deduplicate_success_and_require_explicit_retry_after_failure(self):
        manager = WorkflowJobManager(max_workers=1)
        calls = []

        def successful(progress):
            calls.append("success")
            progress(1, 1, "done")
            return {"value": 1}

        first = manager.start("process", "same", successful)
        while manager.get(first["job_id"])["status"] not in {"succeeded", "failed"}:
            time.sleep(0.01)
        second = manager.start("process", "same", successful)
        self.assertEqual(first["job_id"], second["job_id"])
        self.assertEqual(calls, ["success"])

        def failing(progress):
            calls.append("failed")
            raise WorkflowFailure("BAD_CONFIG", "bad config", stage="5", field="columns")

        failed = manager.start("process", "failure", failing)
        while manager.get(failed["job_id"])["status"] not in {"succeeded", "failed"}:
            time.sleep(0.01)
        same_failure = manager.start("process", "failure", failing)
        self.assertEqual(failed["job_id"], same_failure["job_id"])
        retried = manager.start("process", "failure", successful, retry=True)
        self.assertNotEqual(failed["job_id"], retried["job_id"])
        manager.shutdown()

    def test_stable_signature_ignores_dictionary_order(self):
        self.assertEqual(stable_signature({"a": 1, "b": 2}), stable_signature({"b": 2, "a": 1}))


if __name__ == "__main__":
    unittest.main()
