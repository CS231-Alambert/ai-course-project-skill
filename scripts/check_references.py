#!/usr/bin/env python3
"""Validate and report on a Markdown reference list.

Checks: duplicate indices, missing years, language balance, DOI presence.

Usage:
  python check_references.py paper-output/thesis.md
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

REF_RE = re.compile(r"^\[(\d+)\]\s*(.+)$")


def parse_refs(path: Path) -> list[dict]:
    """Extract reference lines from a Markdown file."""
    refs = []
    in_refs = False
    for line in path.read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if s.startswith("##") and "参考" in s:
            in_refs = True
            continue
        if in_refs and s.startswith("##") and "参考" not in s:
            break
        if not in_refs:
            continue
        m = REF_RE.match(s)
        if m:
            body = m.group(2).strip()
            refs.append({
                "index": int(m.group(1)),
                "raw": s,
                "lang": "zh" if re.search(r"[一-鿿]", body) else "en",
                "year": _extract_year(body),
                "has_doi": "doi" in body.lower(),
            })
    return refs


def _extract_year(text: str) -> int | None:
    years = re.findall(r"\b(20\d{2})\b", text)
    return int(years[0]) if years else None


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python check_references.py <thesis.md>")
        return 1

    refs = parse_refs(Path(sys.argv[1]))
    if not refs:
        print("No references found (look for ## 参考文献 section).")
        return 0

    # Duplicates
    indices = [r["index"] for r in refs]
    dups = [i for i, c in Counter(indices).items() if c > 1]

    # Stats
    langs = Counter(r["lang"] for r in refs)
    bad_years = [r for r in refs if r["year"] is None or r["year"] < 2018]
    has_doi = sum(1 for r in refs if r["has_doi"])

    print(f"参考文献总数: {len(refs)}")
    print(f"  中文 (zh): {langs.get('zh', 0)}")
    print(f"  英文 (en): {langs.get('en', 0)}")
    print(f"  含 DOI  : {has_doi}")
    print()

    if dups:
        print(f"⚠  重复编号: {dups}")
    if bad_years:
        print(f"⚠  年份缺失或过旧 ({len(bad_years)} 条):")
        for r in bad_years:
            print(f"   [{r['index']}] {r['raw'][:80]}")
    if not dups and not bad_years:
        print("✓  格式检查通过。")

    # Suggest: at least 10 references for a decent course paper
    if len(refs) < 10:
        print(f"\n💡 提示：当前 {len(refs)} 条参考文献偏少，建议 ≥ 10 条。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
