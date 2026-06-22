#!/usr/bin/env python3
"""Scan a project directory and extract structured evidence for course report writing.

Produces paper-context/evidence/ with:
  project-evidence.json  — machine-readable full dump
  tech-stack.md          — detected languages, frameworks, dependencies
  code-structure.md      — source file inventory with function/class signatures
  api-list.md            — API routes with HTTP methods, params, handlers
  database-schema.md     — schema/models with table/column names
  config-values.md       — hyperparameters, constants, magic numbers
  test-results.md        — test files and coverage hints
  readme-summary.md      — README content digest
  git-info.md            — git history, branches, recent changes
  dependencies.md        — import graph and external dependencies
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
from collections import defaultdict
from dataclasses import asdict, dataclass, field
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
class FuncInfo:
    name: str
    file: str
    line: int
    signature: str = ""
    docstring: str = ""
    decorators: list[str] = field(default_factory=list)


@dataclass
class ApiRoute:
    file: str
    method: str = ""
    path: str = ""
    handler: str = ""
    params: list[str] = field(default_factory=list)


@dataclass
class SchemaInfo:
    file: str
    entity_type: str = ""  # "table", "model", "collection"
    name: str = ""
    columns: list[str] = field(default_factory=list)


@dataclass
class ConfigValue:
    file: str
    key: str
    value: str
    context: str = ""


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

    # New rich fields
    functions: list[FuncInfo] = field(default_factory=list)
    api_routes: list[ApiRoute] = field(default_factory=list)
    schemas: list[SchemaInfo] = field(default_factory=list)
    config_values: list[ConfigValue] = field(default_factory=list)
    imports_graph: dict[str, list[str]] = field(default_factory=dict)
    imports_external: dict[str, int] = field(default_factory=dict)
    file_sizes: dict[str, int] = field(default_factory=dict)
    readme_summary: str = ""
    git_info: dict = field(default_factory=dict)


def should_skip(path: Path, root: Path) -> bool:
    try:
        for part in path.relative_to(root).parts:
            if part in IGNORE_DIRS or part.startswith("."):
                return True
    except ValueError:
        return True
    return False


# ── Enhanced extractors ───────────────────────────────────────

def extract_python_functions(path: Path) -> list[FuncInfo]:
    """Extract function/class signatures from Python files."""
    funcs = []
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except Exception:
        return funcs

    rel = path.name  # We'll resolve relative paths later
    i = 0
    while i < len(lines):
        s = lines[i].strip()

        # Decorators
        decorators = []
        while s.startswith("@"):
            decorators.append(s)
            i += 1
            if i < len(lines):
                s = lines[i].strip()
            else:
                break

        # Function definition
        m = re.match(r"(?:async\s+)?def\s+(\w+)\s*\(([^)]*)\)", s)
        if m:
            name = m.group(1)
            params = m.group(2)
            sig = f"def {name}({params})"
            if decorators:
                sig = "\n".join(decorators) + "\n" + sig

            # Look for docstring
            doc = ""
            j = i + 1
            while j < len(lines) and (lines[j].strip().startswith('"""')
                                       or lines[j].strip().startswith("'''")
                                       or (j == i + 1 and lines[j].strip().startswith("#"))):
                doc += lines[j].strip() + " "
                if '"""' in lines[j] and lines[j].count('"""') >= 2:
                    break
                if "'''" in lines[j] and lines[j].count("'''") >= 2:
                    break
                j += 1
            doc = doc.strip()[:200]

            funcs.append(FuncInfo(
                name=name, file=rel, line=i + 1,
                signature=sig, docstring=doc,
                decorators=decorators,
            ))

        # Class definition
        cm = re.match(r"class\s+(\w+)\s*(?:\(([^)]*)\))?:", s)
        if cm:
            funcs.append(FuncInfo(
                name=cm.group(1), file=rel, line=i + 1,
                signature=f"class {cm.group(1)}({cm.group(2) or ''})",
            ))

        i += 1

    return funcs


def extract_js_ts_functions(path: Path) -> list[FuncInfo]:
    """Extract function/class signatures from JS/TS files."""
    funcs = []
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except Exception:
        return funcs

    rel = path.name
    i = 0
    while i < len(lines):
        s = lines[i].strip()

        # function / arrow function / const function
        for pat, tp in [(r"(?:export\s+)?(?:async\s+)?function\s+(\w+)\s*\(([^)]*)\)", "function"),
                        (r"(?:export\s+)?(?:const|let|var)\s+(\w+)\s*=\s*(?:async\s+)?\(([^)]*)\)", "arrow"),
                        (r"(?:export\s+)?class\s+(\w+)(?:\s+extends\s+(\w+))?\s*\{", "class")]:
            for m in [re.match(pat, s)]:
                if not m:
                    continue
                name = m.group(1)
                if tp == "class":
                    extends = m.group(2) if m.lastindex and m.lastindex >= 2 else ""
                    sig = f"class {name}" + (f" extends {extends}" if extends else "")
                    funcs.append(FuncInfo(name=name, file=rel, line=i+1, signature=sig))
                else:
                    params = (m.group(2) if m.lastindex and m.lastindex >= 2 else "") or ""
                    funcs.append(FuncInfo(name=name, file=rel, line=i+1,
                                          signature=f"{tp} {name}({params})"))

        i += 1
    return funcs


def extract_java_functions(path: Path) -> list[FuncInfo]:
    """Extract method/class signatures from Java files."""
    funcs = []
    try:
        content = path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return funcs

    rel = path.name
    for m in re.finditer(
        r'(?:public|private|protected|static|\s)+[\w<>\[\],\s]+\s+(\w+)\s*\(([^)]*)\)\s*(?:\{|throws)',
        content
    ):
        funcs.append(FuncInfo(
            name=m.group(1), file=rel,
            line=content[:m.start()].count('\n') + 1,
            signature=f"method {m.group(1)}({m.group(2)})",
        ))
    return funcs


def extract_python_imports(path: Path) -> list[str]:
    """Extract import statements from Python files."""
    imports = []
    try:
        for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
            s = line.strip()
            if s.startswith("import ") or s.startswith("from "):
                imports.append(s)
    except Exception:
        pass
    return imports


def extract_api_routes(path: Path) -> list[ApiRoute]:
    """Extract API route definitions from Python/JS/Java files."""
    routes = []
    try:
        content = path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return routes

    rel = path.name

    # Python Flask / FastAPI patterns
    for m in re.finditer(
        r'@(?:app|router|bp|api)\.(?:route|get|post|put|delete|patch)\s*\(\s*["\']([^"\']+)["\']',
        content
    ):
        method_match = re.search(r'\.(\w+)\s*\(', m.group(0))
        method = method_match.group(1) if method_match else "route"
        routes.append(ApiRoute(
            file=rel, method=method.upper(), path=m.group(1),
            handler="",  # Try to find the function name on the next line
        ))

    # Python Django URL patterns
    for m in re.finditer(r"path\s*\(\s*['\"]([^'\"]+)['\"]\s*,\s*(\w+)", content):
        routes.append(ApiRoute(file=rel, path=m.group(1), handler=m.group(2)))

    # Express.js patterns
    for m in re.finditer(
        r'(?:app|router)\.(?:get|post|put|delete|patch|use)\s*\(\s*["\']([^"\']+)["\']',
        content
    ):
        method_match = re.search(r'\.(\w+)\s*\(', m.group(0))
        method = method_match.group(1) if method_match else "get"
        routes.append(ApiRoute(
            file=rel, method=method.upper(), path=m.group(1),
        ))

    # Java Spring patterns
    for m in re.finditer(
        r'@(?:GetMapping|PostMapping|PutMapping|DeleteMapping|RequestMapping)\s*(?:\(\s*(?:value\s*=\s*)?["\']([^"\']+)["\'])?',
        content
    ):
        ann = re.search(r'@(\w+)', m.group(0))
        ann_name = ann.group(1) if ann else "RequestMapping"
        method_map = {"GetMapping": "GET", "PostMapping": "POST", "PutMapping": "PUT",
                      "DeleteMapping": "DELETE", "RequestMapping": "ALL"}
        routes.append(ApiRoute(
            file=rel, method=method_map.get(ann_name, "GET"),
            path=m.group(1) or "/",
        ))

    return routes


def extract_sql_schema(path: Path) -> list[SchemaInfo]:
    """Extract table names and columns from SQL files."""
    schemas = []
    try:
        content = path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return schemas

    for m in re.finditer(
        r'CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?[`"\[]?(\w+)[`"\]]?\s*\(([\s\S]*?)\)\s*;',
        content, re.IGNORECASE
    ):
        name = m.group(1)
        body = m.group(2)
        cols = re.findall(r'^\s*[`"\[]?(\w+)[`"\]]?\s+', body, re.MULTILINE)
        schemas.append(SchemaInfo(
            file=path.name, entity_type="table", name=name, columns=cols,
        ))

    return schemas


def extract_python_model(path: Path) -> list[SchemaInfo]:
    """Extract ORM model info from Python files."""
    schemas = []
    try:
        content = path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return schemas

    # SQLAlchemy / Django models
    for m in re.finditer(r'class\s+(\w+)\s*\([^)]*(?:Model|Base|db\.Model)[^)]*\):', content):
        name = m.group(1)
        body_start = m.end()
        body_end = content.find("\nclass ", body_start)
        if body_end == -1:
            body_end = len(content)
        body = content[body_start:body_end]

        cols = re.findall(r'^\s*(\w+)\s*=\s*', body, re.MULTILINE)
        schemas.append(SchemaInfo(
            file=path.name, entity_type="model", name=name, columns=cols,
        ))

    return schemas


def extract_config_values(path: Path) -> list[ConfigValue]:
    """Extract config values, hyperparameters, constants."""
    configs = []
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except Exception:
        return configs

    rel = path.name
    for i, line in enumerate(lines):
        s = line.strip()

        # Python constant assignments
        m = re.match(r'^([A-Z][A-Z_0-9]+)\s*=\s*(.+?)(?:\s*#.*)?$', s)
        if m:
            configs.append(ConfigValue(
                file=rel, key=m.group(1), value=m.group(2).strip()[:80],
                context=f"line {i+1}",
            ))

        # Hyperparameter patterns
        for hp_pat, hp_name in [
            (r'(?:lr|learning_rate|LEARNING_RATE)\s*=\s*([0-9.eE+-]+)', "learning_rate"),
            (r'(?:batch_size|BATCH_SIZE)\s*=\s*(\d+)', "batch_size"),
            (r'(?:epochs|EPOCHS|num_epochs)\s*=\s*(\d+)', "epochs"),
            (r'(?:hidden_size|HIDDEN_SIZE)\s*=\s*(\d+)', "hidden_size"),
            (r'(?:dropout|DROPOUT)\s*=\s*([0-9.]+)', "dropout"),
            (r'(?:optimizer|OPTIMIZER)\s*=\s*["\'](\w+)["\']', "optimizer"),
        ]:
            m = re.search(hp_pat, s, re.IGNORECASE)
            if m:
                configs.append(ConfigValue(
                    file=rel, key=hp_name, value=m.group(1),
                    context=f"line {i+1}",
                ))

    return configs


def extract_readme_summary(project_root: Path) -> str:
    """Extract a summary of the README file."""
    candidates = ["README.md", "README.txt", "readme.md", "Readme.md", "README.rst"]
    for cand in candidates:
        path = project_root / cand
        if path.exists() and path.is_file():
            try:
                text = path.read_text(encoding="utf-8", errors="replace")
                # Grab first 300 chars after any badges
                clean = re.sub(r"!\[.*?\]\(.*?\)", "", text)
                clean = re.sub(r"\[!\[.*?\].*?\n", "", clean)
                clean = re.sub(r"<[^>]+>", "", clean)
                clean = re.sub(r"---.*?---", "", clean, flags=re.DOTALL)
                clean = re.sub(r"\n{3,}", "\n\n", clean)
                return clean.strip()[:500]
            except Exception:
                return ""
    return ""


def extract_git_info(project_root: Path) -> dict:
    """Extract git repository metadata."""
    info = {}
    git_dir = project_root / ".git"
    if not git_dir.exists():
        return info

    def _run(cmd: list[str]) -> str:
        try:
            result = subprocess.run(
                cmd, cwd=str(project_root),
                capture_output=True, text=True, timeout=15,
            )
            return result.stdout.strip()
        except Exception:
            return ""

    info["branch"] = _run(["git", "rev-parse", "--abbrev-ref", "HEAD"])
    info["last_commit"] = _run(["git", "log", "-1", "--format=%s"])
    info["last_commit_date"] = _run(["git", "log", "-1", "--format=%ci"])
    info["total_commits"] = _run(["git", "rev-list", "--count", "HEAD"])
    info["contributors"] = _run(["git", "shortlog", "-sn", "HEAD"]).split("\n")[:5]
    info["recent_changes"] = _run(["git", "log", "--oneline", "-5"])
    info["remote_url"] = _run(["git", "remote", "get-url", "origin"])

    return {k: v for k, v in info.items() if v}


def calc_file_sizes(project_root: Path) -> dict[str, int]:
    """Calculate file sizes (in lines) for source files."""
    sizes = {}
    for ext in SOURCE_EXTS:
        for path in project_root.rglob(f"*{ext}"):
            if should_skip(path, project_root):
                continue
            try:
                rel = path.relative_to(project_root).as_posix()
                sizes[rel] = path.stat().st_size
            except Exception:
                pass
    return sizes


# ── Top-level builder ────────────────────────────────────────

def build_evidence(project_root: Path) -> Evidence:
    files = [p for p in sorted(project_root.rglob("*"))
             if p.is_file() and not should_skip(p, project_root)]

    tech = []
    src = []
    api_files = []
    schema_files = []
    tests = []
    notebooks = []
    docs = []

    all_funcs = []
    all_routes = []
    all_schemas = []
    all_configs = []
    imports_graph = defaultdict(list)
    imports_external = defaultdict(int)

    for path in files:
        r = path.relative_to(project_root).as_posix()
        lower = r.lower()
        ext = path.suffix.lower()

        # Tech markers
        if path.name in TECH_MARKERS:
            tech.append({"file": r, "technology": TECH_MARKERS[path.name]})

        # Source files
        if ext in SOURCE_EXTS:
            src.append(r)

        # API candidates
        if ext in {".py", ".js", ".ts", ".java", ".go", ".php"}:
            if any(kw in lower for kw in API_HINTS):
                api_files.append(r)
            # Always try to extract routes from these
            routes = extract_api_routes(path)
            if routes:
                all_routes.extend(routes)
                if r not in api_files:
                    api_files.append(r)

        # Schema
        if ext == ".sql":
            schema_files.append(r)
            tbls = extract_sql_schema(path)
            if tbls:
                all_schemas.extend(tbls)
        if ext in {".py", ".js", ".ts", ".java"}:
            if any(kw in lower for kw in SCHEMA_HINTS):
                schema_files.append(r)
            # Try to extract models
            if ext == ".py":
                models = extract_python_model(path)
                if models:
                    all_schemas.extend(models)
                    if r not in schema_files:
                        schema_files.append(r)

        # Tests
        if any(kw in lower for kw in TEST_HINTS):
            tests.append(r)

        # Notebooks
        if ext == ".ipynb":
            notebooks.append(r)

        # Docs
        if ext in {".md", ".rst", ".txt", ".pdf"}:
            docs.append(r)

        # Function signatures (limit to keep reasonable size)
        if len(all_funcs) < 200:
            if ext == ".py":
                all_funcs.extend(extract_python_functions(path))
            elif ext in {".js", ".ts", ".jsx", ".tsx"}:
                all_funcs.extend(extract_js_ts_functions(path))
            elif ext == ".java":
                all_funcs.extend(extract_java_functions(path))

        # Config values
        if ext in {".py", ".yml", ".yaml", ".json", ".toml"} and len(all_configs) < 100:
            all_configs.extend(extract_config_values(path))

        # Imports
        if ext == ".py" and len(imports_graph) < 100:
            imports_graph[r] = extract_python_imports(path)
            # Count external packages
            for imp_line in imports_graph[r]:
                m = re.match(r'(?:from|import)\s+(\w+)', imp_line)
                if m:
                    pkg = m.group(1)
                    if pkg not in {"os", "sys", "re", "json", "math", "time",
                                   "datetime", "typing", "collections", "pathlib",
                                   "dataclasses", "abc", "functools", "itertools"}:
                        imports_external[pkg] += 1

    return Evidence(
        root=str(project_root),
        tech_markers=tech,
        source_files=src[:500],
        possible_api_files=api_files[:200],
        possible_schema_files=schema_files[:200],
        possible_test_files=tests[:200],
        notebook_files=notebooks[:50],
        doc_files=docs[:100],
        functions=all_funcs[:200],
        api_routes=all_routes[:100],
        schemas=all_schemas[:50],
        config_values=all_configs[:100],
        imports_graph=dict(imports_graph),
        imports_external=dict(sorted(imports_external.items(), key=lambda x: -x[1])[:30]),
        file_sizes=calc_file_sizes(project_root),
        readme_summary=extract_readme_summary(project_root),
        git_info=extract_git_info(project_root),
    )


# ── Output writers ────────────────────────────────────────────

def _write_md(path: Path, title: str, items: list[str], empty_msg: str) -> None:
    lines = [f"# {title}", ""]
    if items:
        lines.extend(f"- `{item}`" for item in items)
    else:
        lines.append(empty_msg)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_outputs(ev: Evidence, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)

    # ── JSON dump ──
    (out_dir / "project-evidence.json").write_text(
        json.dumps(asdict(ev), ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )

    # ── Code Structure ──
    _write_md(out_dir / "code-structure.md", "Code Structure", ev.source_files,
              "No source files detected.")

    # ── Tech Stack ──
    lines = ["# Technology Stack", ""]
    if ev.tech_markers:
        lines.extend(f"- `{m['file']}` → {m['technology']}" for m in ev.tech_markers)
    else:
        lines.append("No known tech markers found.")
    # Dependency summary
    if ev.imports_external:
        lines.append("")
        lines.append("## External Dependencies")
        lines.extend(f"- `{pkg}` ({cnt} references)" for pkg, cnt in list(ev.imports_external.items())[:20])
    (out_dir / "tech-stack.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    # ── API List ──
    lines = ["# API / Entry Points", ""]
    if ev.api_routes:
        lines.append("| File | Method | Path | Handler |")
        lines.append("|------|--------|------|---------|")
        for r in ev.api_routes[:50]:
            lines.append(f"| `{r.file}` | {r.method} | `{r.path}` | `{r.handler}` |")
        lines.append("")
    if ev.possible_api_files:
        lines.append("## Candidate Files")
        lines.extend(f"- `{f}`" for f in ev.possible_api_files)
    else:
        lines.append("No API candidates detected.")
    (out_dir / "api-list.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    # ── Database Schema ──
    lines = ["# Data & Schema", ""]
    if ev.schemas:
        for s in ev.schemas[:30]:
            lines.append(f"## {s.entity_type.title()}: {s.name}")
            lines.append(f"**File:** `{s.file}`")
            if s.columns:
                lines.append(f"**Fields:** {', '.join(f'`{c}`' for c in s.columns[:20])}")
            lines.append("")
    if ev.possible_schema_files:
        lines.append("## Candidate Files")
        lines.extend(f"- `{f}`" for f in ev.possible_schema_files)
    if not ev.schemas and not ev.possible_schema_files:
        lines.append("No schema/model files detected.")
    (out_dir / "database-schema.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    # ── Config Values ──
    lines = ["# Configuration & Hyperparameters", ""]
    if ev.config_values:
        lines.append("| File | Key | Value |")
        lines.append("|------|-----|-------|")
        for c in ev.config_values[:50]:
            lines.append(f"| `{c.file}` | `{c.key}` | `{c.value}` |")
    else:
        lines.append("No config values detected.")
    (out_dir / "config-values.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    # ── Function signatures ──
    lines = ["# Function & Class Signatures", ""]
    if ev.functions:
        by_file = defaultdict(list)
        for f in ev.functions:
            by_file[f.file].append(f)
        for file, funcs in sorted(by_file.items()):
            lines.append(f"## {file}")
            for f in funcs[:15]:
                lines.append(f"```python")
                lines.append(f.signature)
                if f.docstring:
                    lines.append(f"# {f.docstring}")
                lines.append(f"```")
                lines.append("")
    else:
        lines.append("No function signatures extracted.")
    (out_dir / "functions.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    # ── Tests ──
    _write_md(out_dir / "test-results.md", "Tests", ev.possible_test_files,
              "No test files detected.")

    # ── README Summary ──
    if ev.readme_summary:
        (out_dir / "readme-summary.md").write_text(
            f"# README Summary\n\n{ev.readme_summary}\n", encoding="utf-8"
        )

    # ── Git Info ──
    if ev.git_info:
        lines = ["# Git Repository Info", ""]
        for k, v in ev.git_info.items():
            label = {
                "branch": "Branch", "last_commit": "Last Commit",
                "last_commit_date": "Date", "total_commits": "Total Commits",
                "remote_url": "Remote", "recent_changes": "Recent Changes",
                "contributors": "Contributors",
            }.get(k, k.replace("_", " ").title())
            if isinstance(v, list):
                lines.append(f"**{label}:**")
                lines.extend(f"- {item}" for item in v)
            else:
                lines.append(f"- **{label}:** {v}")
        (out_dir / "git-info.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    # ── Notebooks ──
    if ev.notebook_files:
        _write_md(out_dir / "notebooks.md", "Jupyter Notebooks", ev.notebook_files, "")

    # ── Docs ──
    if ev.doc_files:
        _write_md(out_dir / "documentation.md", "Documentation Files", ev.doc_files, "")


# ── CLI ──────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract rich evidence from a project for course report writing."
    )
    parser.add_argument("project", nargs="?", default=".", help="Project directory to scan")
    parser.add_argument("--out", default="paper-context/evidence", help="Output directory")
    args = parser.parse_args()

    root = Path(args.project).resolve()
    out = Path(args.out).resolve()
    ev = build_evidence(root)
    write_outputs(ev, out)

    print(f"Evidence written to {out}/")
    print(f"  Source files  : {len(ev.source_files)}")
    print(f"  Tech markers  : {len(ev.tech_markers)}")
    print(f"  Functions     : {len(ev.functions)}")
    print(f"  API routes    : {len(ev.api_routes)}")
    print(f"  Schema entities: {len(ev.schemas)}")
    print(f"  Config values : {len(ev.config_values)}")
    print(f"  Test files    : {len(ev.possible_test_files)}")
    if ev.notebook_files:
        print(f"  Notebooks     : {len(ev.notebook_files)}")
    if ev.readme_summary:
        print(f"  README        : found ({len(ev.readme_summary)} chars)")
    if ev.git_info:
        print(f"  Git info      : {len(ev.git_info)} fields")
    print(f"  Ext. deps     : {len(ev.imports_external)} packages")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
