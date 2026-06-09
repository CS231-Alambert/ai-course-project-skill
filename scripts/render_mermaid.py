#!/usr/bin/env python3
"""Render Mermaid diagrams from a Markdown file to PNG images.

Requires: mmdc (npm i -g @mermaid-js/mermaid-cli) or Docker.
If mmdc not found, prints instructions and exits cleanly.

Usage:
  python render_mermaid.py thesis.md figures/
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
from pathlib import Path


def extract_mermaid_blocks(path: Path) -> list[dict]:
    """Extract ```mermaid blocks with their following captions."""
    blocks = []
    lines = path.read_text(encoding="utf-8").splitlines()
    i = 0
    while i < len(lines):
        s = lines[i].strip()
        if s.lower().startswith("```mermaid"):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            i += 1  # skip closing ```
            # Look ahead for a figure caption
            cap = f"diagram_{len(blocks)+1:02d}"
            if i < len(lines):
                cm = re.match(r"^图\s*\d+[\.\-]\d*\s*(.+)", lines[i].strip())
                if cm:
                    cap = cm.group(1).strip()[:40]
            blocks.append({"code": "\n".join(code_lines), "caption": cap, "index": len(blocks) + 1})
        else:
            i += 1
    return blocks


def render_blocks(blocks: list[dict], out_dir: Path) -> dict[str, Path]:
    mmdc = shutil.which("mmdc")
    image_map = {}

    if not mmdc:
        print("⚠  mmdc not found. Install with: npm i -g @mermaid-js/mermaid-cli")
        print("   Or use Docker: docker run --rm -v $PWD:/data minlag/mermaid-cli ...")
        print("   Writing .mmd source files for manual rendering.\n")
        for b in blocks:
            mmd_path = out_dir / f"{b['caption']}.mmd"
            mmd_path.write_text(b["code"], encoding="utf-8")
        return image_map

    out_dir.mkdir(parents=True, exist_ok=True)
    for b in blocks:
        mmd = out_dir / f"_{b['index']:02d}.mmd"
        png = out_dir / f"{b['caption']}.png"
        mmd.write_text(b["code"], encoding="utf-8")
        try:
            subprocess.run([mmdc, "-i", str(mmd), "-o", str(png), "-w", "1200"],
                           check=True, capture_output=True, timeout=120)
            image_map[b["caption"]] = png
            print(f"  ✓  {png.name}")
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            print(f"  ✗  {b['caption']}: {e}")
            mmd.rename(out_dir / f"{b['caption']}.mmd")
    return image_map


def main() -> int:
    if len(sys.argv) < 3:
        print("Usage: python render_mermaid.py <thesis.md> <output-dir>")
        return 1

    md_path = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])
    blocks = extract_mermaid_blocks(md_path)
    if not blocks:
        print("No Mermaid blocks found.")
        return 0

    print(f"Found {len(blocks)} Mermaid block(s). Rendering...")
    render_blocks(blocks, out_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
