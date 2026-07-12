"""Build metadata embedded in ProductCodeFormatter releases."""

import json
import os
import sys
from pathlib import Path


DEFAULT_VERSION = "0.4.0"
DEFAULT_NOTES = "Velopack installer and reliable Raspberry Pi updates."


def _embedded_metadata():
    candidates = []
    bundle_root = getattr(sys, "_MEIPASS", "")
    if bundle_root:
        candidates.append(Path(bundle_root) / "release_metadata.json")
    candidates.append(Path(__file__).resolve().with_name("release_metadata.json"))
    for path in candidates:
        try:
            if path.exists():
                value = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(value, dict):
                    return value
        except (OSError, ValueError, TypeError):
            continue
    return {}


_METADATA = _embedded_metadata()
VERSION = str(os.environ.get("PRODUCT_CODE_FORMATTER_BUILD_VERSION") or _METADATA.get("version") or DEFAULT_VERSION)
NOTES = str(os.environ.get("PRODUCT_CODE_FORMATTER_BUILD_NOTES") or _METADATA.get("notes") or DEFAULT_NOTES)
CHANNEL = str(os.environ.get("PRODUCT_CODE_FORMATTER_BUILD_CHANNEL") or _METADATA.get("channel") or "dev")