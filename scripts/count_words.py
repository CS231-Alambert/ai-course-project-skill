#!/usr/bin/env python3
"""Count words per chapter in a Markdown thesis file.

Outputs a table showing chapter titles and approximate word counts
(Chinese characters + English words).
"""

from __future__ import annotations

import re
import sys
from pathlib import Path


def approx_words(text: str) -> int:
    """Chinese chars + English words ≈ total word count."""
    cn = len(re.findall(r"[一-鿿]", text))
    en = len(re.findall(r"[A-Za-z]+(?:[-'][A-Za-z]+)*", text))
    digits = len(re.findall(r"[0-9]+", text))
    return cn + en + digits


def count_chapters(path: Path) -> list[dict]:
    chapters = []
    current_title = "(前言)"
    current_text: list[str] = []

    for line in path.read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if s.startswith("## ") and not s.startswith("### "):
            if current_text:
                chapters.append({"title": current_title, "words": approx_words("\n".join(current_text))})
            current_title = s[3:].strip()
            current_text = []
        else:
            current_text.append(line)

    if current_text:
        chapters.append({"title": current_title, "words": approx_words("\n".join(current_text))})

    return chapters


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python count_words.py <thesis.md>")
        return 1

    chapters = count_chapters(Path(sys.argv[1]))
    total = 0
    print(f"{'章节':<30} {'字数':>8}")
    print("-" * 42)
    for ch in chapters:
        print(f"{ch['title']:<30} {ch['words']:>8,}")
        total += ch["words"]
    print("-" * 42)
    print(f"{'合计':<30} {total:>8,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
