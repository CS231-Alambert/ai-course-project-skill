#!/usr/bin/env python3
"""Initialize a lightweight course project workspace — one command, done."""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

WORKSPACE_DIRS = [
    "paper-output/figures",
    "paper-output/screenshots",
    "paper-context/evidence",
    "paper-context/literature",
]


def init_workspace(target: str, overwrite: bool = False) -> Path:
    root = Path(target).resolve()
    root.mkdir(parents=True, exist_ok=True)

    for sub in WORKSPACE_DIRS:
        (root / sub).mkdir(parents=True, exist_ok=True)

    # Copy bundled templates if not present
    tmpl_src = Path(__file__).resolve().parents[1] / "templates"
    tmpl_dst = root / "templates"
    if tmpl_src.exists():
        if tmpl_dst.exists() and overwrite:
            shutil.rmtree(tmpl_dst)
        if not tmpl_dst.exists():
            shutil.copytree(tmpl_src, tmpl_dst)

    # Write .gitkeep so empty dirs survive git
    for sub in WORKSPACE_DIRS:
        (root / sub / ".gitkeep").touch(exist_ok=True)

    print(f"Workspace ready: {root}")
    print(f"  paper-output/   — your thesis .md and .docx land here")
    print(f"  paper-context/  — evidence and literature references")
    print(f"  templates/      — thesis spec and figure registry YAML")
    print()
    print("Next steps:")
    print(f"  1. Edit {root / 'templates' / 'thesis-spec.yaml'}")
    print(f"  2. Run: python scripts/build_evidence.py <your-project-dir>")
    print(f"  3. Write your thesis as Markdown in paper-output/")
    print(f"  4. Run: python scripts/generate_docx.py paper-output/thesis.md paper-output/thesis.docx")
    return root


def main() -> int:
    parser = argparse.ArgumentParser(description="Initialize a course project paper workspace.")
    parser.add_argument("target", nargs="?", default=".", help="Project directory (default: current)")
    parser.add_argument("--overwrite", action="store_true", help="Replace existing templates/")
    args = parser.parse_args()
    init_workspace(args.target, overwrite=args.overwrite)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
