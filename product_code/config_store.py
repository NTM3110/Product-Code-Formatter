"""Small config-file IO layer.

The schema and normalization stay in app.py for now; this module owns durable
read/write behavior so future config changes have a single IO boundary.
"""

import json
import os
import shutil
from datetime import datetime


def load_config_file(path, default_factory, normalizer):
    if not path.exists():
        return default_factory()
    try:
        return normalizer(json.loads(path.read_text(encoding="utf-8")))
    except Exception:
        return default_factory()


def save_config_file(path, cfg, normalizer, serializer=None):
    normalized = normalizer(cfg)
    stored = serializer(normalized) if serializer else normalized
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temporary.write_text(json.dumps(stored, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)
    return normalized


def backup_config_file(path, backup_dir):
    if not path.exists() or not path.is_file():
        return None
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    target = backup_dir / f"{path.stem}.{stamp}{path.suffix}"
    shutil.copy2(path, target)
    return target
