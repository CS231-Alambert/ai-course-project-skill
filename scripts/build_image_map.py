#!/usr/bin/env python3
"""Build an image-map.json from a directory of figures + optional labels file.

Usage:
  # From a labels.json (list of expected labels):
  python build_image_map.py labels.json figures/ image-map.json

  # Or just scan a directory and map filename→path:
  python build_image_map.py --scan figures/ image-map.json
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

IMG_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".webp"}


def scan_dir(img_dir: Path) -> dict[str, str]:
    """Map every image filename (stem) → absolute path."""
    result = {}
    if not img_dir.is_dir():
        return result
    for f in sorted(img_dir.iterdir()):
        if f.suffix.lower() in IMG_EXTS:
            result[f.stem] = str(f.resolve())
    return result


def from_labels(labels_path: Path, img_dir: Path) -> dict[str, str]:
    labels = json.loads(labels_path.read_text(encoding="utf-8"))
    label_list = labels.get("labels", []) if isinstance(labels, dict) else labels
    result = {}
    for label in label_list:
        for ext in IMG_EXTS:
            cand = img_dir / f"{label}{ext}"
            if cand.exists():
                result[label] = str(cand.resolve())
                break
    return result


def main() -> int:
    ap = argparse.ArgumentParser(description="Build image-map.json for thesis figures.")
    ap.add_argument("source", help="Labels JSON file, or --scan flag")
    ap.add_argument("img_dir", type=Path, help="Directory containing image files")
    ap.add_argument("output", type=Path, help="Output image-map.json path")
    ap.add_argument("--scan", action="store_true", help="Scan img_dir directly, ignore labels file")
    args = ap.parse_args()

    if args.scan:
        result = scan_dir(args.img_dir)
    else:
        result = from_labels(Path(args.source), args.img_dir)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✓  {args.output}  ({len(result)} images mapped)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
