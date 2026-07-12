"""Publish a Windows EXE and signed-by-hash manifest to the Pi release folder."""

import argparse
import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path


parser = argparse.ArgumentParser()
parser.add_argument("exe", type=Path)
parser.add_argument("--version", required=True)
parser.add_argument("--notes", default="")
parser.add_argument("--output", type=Path, default=Path(__file__).parent / "releases")
args = parser.parse_args()

if not args.exe.exists() or args.exe.suffix.lower() != ".exe":
    raise SystemExit("EXE không tồn tại hoặc không phải .exe")
args.output.mkdir(parents=True, exist_ok=True)
filename = f"ProductCodeFormatter_v{args.version}.exe"
target = args.output / filename
shutil.copy2(args.exe, target)
for old_release in args.output.glob("ProductCodeFormatter*.exe"):
    if old_release != target:
        old_release.unlink(missing_ok=True)
digest = hashlib.sha256(target.read_bytes()).hexdigest()
manifest = {
    "version": args.version,
    "platform": "windows-x64",
    "filename": filename,
    "download_url": f"/downloads/{filename}",
    "sha256": digest,
    "size": target.stat().st_size,
    "notes": args.notes,
    "published_at": datetime.now(timezone.utc).isoformat(),
}
(args.output / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(manifest, ensure_ascii=False, indent=2))
