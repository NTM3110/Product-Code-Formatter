"""Small config-file IO layer.

The schema and normalization stay in app.py for now; this module owns durable
read/write behavior so future config changes have a single IO boundary.
"""

import json


def load_config_file(path, default_factory, normalizer):
    if not path.exists():
        return default_factory()
    try:
        return normalizer(json.loads(path.read_text(encoding="utf-8")))
    except Exception:
        return default_factory()


def save_config_file(path, cfg, normalizer):
    normalized = normalizer(cfg)
    path.write_text(json.dumps(normalized, ensure_ascii=False, indent=2), encoding="utf-8")
    return normalized
