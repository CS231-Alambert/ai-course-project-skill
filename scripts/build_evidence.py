#!/usr/bin/env python3
"""Scan a project directory and extract structured evidence for course report writing.

Produces paper-context/evidence/ with:
  project-evidence.json  — machine-readable full dump
  tech-stack.md          — detected languages, frameworks, dependencies
  code-structure.md      — source file inventory
  api-list.md            — API / route / entrypoint candidates
  database-schema.md     — schema / model / data files
  test-results.md        — test files and coverage hints
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path

IGNORE_DIRS = {
    ".git", ".idea", ".vscode", "__pycache__", "node_modules",
    "dist", "build", "target", ".next", ".nuxt", "coverage",
    "venv", ".venv", ".conda", "paper-output", "paper-context",
}

TECH_MARKERS: dict[str, str] = {
    "package.json": "Node.js / TypeScript",
    "requirements.txt": "Python",
    "pyproject.toml": "Python",
    "setup.py": "Python",
    "Pipfile": "Python",
    "go.mod": "Go",
    "Cargo.toml": "Rust",
    "pom.xml": "Java / Maven",
    "build.gradle": "Java / Gradle",
    "CMakeLists.txt": "C/C++ / CMake",
    "Makefile": "C/C++",
    "composer.json": "PHP",
    "pubspec.yaml": "Flutter / Dart",
    "app.json": "Mini Program / Mobile",
    "Dockerfile": "Docker",
    "docker-compose.yml": "Docker Compose",
}

SOURCE_EXTS = {".py", ".js", ".jsx", ".ts", ".tsx", ".java", ".go", ".rs",
               ".php", ".vue", ".cpp", ".c", ".h", ".hpp", ".cs",
               ".sql", ".xml", ".yml", ".yaml", ".json", ".toml",
               ".ipynb", ".R", ".m", ".swift", ".kt", ".dart"}

TEST_HINTS = ("test", "tests", "spec", "__tests__", "junit", "pytest", "coverage", "benchmark")
SCHEMA_HINTS = ("schema", "migration", "migrations", "sql", "entity", "model", "mapper", "dataset")
API_HINTS = ("controller", "route", "api", "handler", "service", "endpoint", "router", "view")


@dataclass
class Evidence:
    root: str
    tech_markers: list[dict[str, str]]
    source_files: list[str]
    possible_api_files: list[str]
    possible_schema_files: list[str]
    possible_test_files: list[str]
    notebook_files: list[str]
    doc_files: list[str]


def should_skip(path: Path, root: Path) -> bool:
    try:
        for part in path.relative_to(root).parts:
            if part in IGNORE_DIRS or part.startswith("."):
                return True
    except ValueError:
        return True
    return False


def build_evidence(project_root: Path) -> Evidence:
    files = [p for p in sorted(project_root.rglob("*")) if p.is_file() and not should_skip(p, project_root)]

    tech = []
    src = []
    api = []
    schema = []
    tests = []
    notebooks = []
    docs = []

    for path in files:
        r = path.relative_to(project_root).as_posix()
        lower = r.lower()

        if path.name in TECH_MARKERS:
            tech.append({"file": r, "technology": TECH_MARKERS[path.name]})

        ext = path.suffix.lower()
        if ext in SOURCE_EXTS:
            src.append(r)

        if ext in {".py", ".js", ".ts", ".java", ".go", ".php", ".ipynb"}:
            if any(kw in lower for kw in API_HINTS):
                api.append(r)

        if ext == ".sql" or any(kw in lower for kw in SCHEMA_HINTS):
            schema.append(r)

        if any(kw in lower for kw in TEST_HINTS):
            tests.append(r)

        if ext == ".ipynb":
            notebooks.append(r)

        if ext in {".md", ".rst", ".txt", ".pdf"}:
            docs.append(r)

    return Evidence(
        root=str(project_root),
        tech_markers=tech,
        source_files=src[:500],
        possible_api_files=api[:200],
        possible_schema_files=schema[:200],
        possible_test_files=tests[:200],
        notebook_files=notebooks[:50],
        doc_files=docs[:100],
    )


def _write_md(path: Path, title: str, items: list[str], empty_msg: str) -> None:
    lines = [f"# {title}", ""]
    if items:
        lines.extend(f"- `{item}`" for item in items)
    else:
        lines.append(empty_msg)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_outputs(ev: Evidence, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)

    (out_dir / "project-evidence.json").write_text(
        json.dumps(asdict(ev), ensure_ascii=False, indent=2), encoding="utf-8")

    _write_md(out_dir / "code-structure.md", "Code Structure", ev.source_files,
              "No source files detected.")

    lines = ["# Technology Stack", ""]
    if ev.tech_markers:
        lines.extend(f"- `{m['file']}` → {m['technology']}" for m in ev.tech_markers)
    else:
        lines.append("No known tech markers found. Check manually.")
    (out_dir / "tech-stack.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    _write_md(out_dir / "api-list.md", "API / Entry Points", ev.possible_api_files,
              "No API candidates detected.")

    _write_md(out_dir / "database-schema.md", "Data & Schema", ev.possible_schema_files,
              "No schema/model files detected.")

    _write_md(out_dir / "test-results.md", "Tests", ev.possible_test_files,
              "No test files detected.")

    if ev.notebook_files:
        _write_md(out_dir / "notebooks.md", "Jupyter Notebooks", ev.notebook_files, "")

    if ev.doc_files:
        _write_md(out_dir / "documentation.md", "Documentation Files", ev.doc_files, "")


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract evidence from a project for course report writing.")
    parser.add_argument("project", nargs="?", default=".", help="Project directory to scan")
    parser.add_argument("--out", default="paper-context/evidence", help="Output directory")
    args = parser.parse_args()

    root = Path(args.project).resolve()
    out = Path(args.out).resolve()
    ev = build_evidence(root)
    write_outputs(ev, out)

    print(f"Evidence written to {out}/")
    print(f"  Source files : {len(ev.source_files)}")
    print(f"  Tech markers : {len(ev.tech_markers)}")
    print(f"  API candidates: {len(ev.possible_api_files)}")
    print(f"  Schema files : {len(ev.possible_schema_files)}")
    print(f"  Test files   : {len(ev.possible_test_files)}")
    if ev.notebook_files:
        print(f"  Notebooks    : {len(ev.notebook_files)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
