#!/usr/bin/env python3
"""Parse a course requirement document and extract formatting specifications.

Reads .md / .txt / .pdf / .docx files and extracts:
  - Font names (body + heading)
  - Font sizes (in pt)
  - Line spacing
  - Page margins
  - First-line indent
  - Word count requirements
  - Required chapter structure
  - Submission/formatting rules
  - Reference count / language balance hints

Outputs a thesis-spec.yaml fragment to stdout (or --merge into existing spec).

Usage:
  python parse_requirements.py 课程要求.txt
  python parse_requirements.py 人工智能导论大作业要求.docx --out templates/thesis-spec.yaml
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

FONT_SIZE_CN = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}

FONT_NAMES = [
    "宋体", "黑体", "楷体", "仿宋", "华文中宋", "微软雅黑",
    "Times New Roman", "Consolas", "Arial", "Courier New",
]

# ── File readers ──────────────────────────────────────────────

def read_text(path: Path) -> str:
    """Read plain text from .md / .txt."""
    return path.read_text(encoding="utf-8", errors="replace")


def read_docx(path: Path) -> str:
    """Extract raw text from .docx using python-docx."""
    try:
        from docx import Document
        doc = Document(str(path))
        return "\n".join(p.text for p in doc.paragraphs)
    except ImportError:
        # Fallback: try antiword or strings
        try:
            return subprocess.check_output(
                ["antiword", str(path)], text=True, timeout=30
            )
        except Exception:
            pass
        # Last resort: read as raw and strip XML tags
        raw = path.read_bytes()
        text = re.sub(rb"<[^>]+>", b" ", raw)
        text = re.sub(rb"\s+", b" ", text)
        return text.decode("utf-8", errors="replace")


def read_pdf(path: Path) -> str:
    """Extract text from PDF. Tries pdftotext, then PyPDF2, then pdfplumber."""
    # Try pdftotext first (fastest)
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", str(path), "-"],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    # Try PyPDF2
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(str(path))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except ImportError:
        pass

    # Try pdfplumber
    try:
        import pdfplumber
        with pdfplumber.open(str(path)) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except ImportError:
        pass

    print("⚠ 需要安装 PDF 解析库：pip install PyPDF2 或 pip install pdfplumber", file=sys.stderr)
    print("  或安装系统工具：apt install poppler-utils (提供 pdftotext)", file=sys.stderr)
    return ""


# ── Parsers ──────────────────────────────────────────────────

def extract_font(text: str) -> dict:
    """Extract font specifications."""
    result = {"body": None, "heading": None, "title": None}

    for fn in FONT_NAMES:
        positions = [m.start() for m in re.finditer(re.escape(fn), text)]
        for pos in positions:
            ctx = text[max(0, pos - 30):pos + len(fn) + 40].lower()
            if any(w in ctx for w in ["正文", "段落", "内容", "文字"]):
                if not result["body"]:
                    result["body"] = fn
            elif any(w in ctx for w in ["标题", "题目", "章", "节名"]):
                if not result["heading"]:
                    result["heading"] = fn
            elif any(w in ctx for w in ["大标题", "论文题目", "封面"]):
                if not result["title"]:
                    result["title"] = fn

    # Fallback: first mention of 宋体 is usually body font
    for fn in FONT_NAMES:
        if fn in text:
            if not result["body"]:
                result["body"] = fn
            break

    return {k: v for k, v in result.items() if v}


def extract_font_sizes(text: str) -> dict:
    """Extract font size specifications in pt."""
    result = {}

    for name, pt in FONT_SIZE_CN.items():
        for m in re.finditer(re.escape(name), text):
            ctx_start = max(0, m.start() - 40)
            ctx = text[ctx_start:m.end() + 30].lower()

            if any(w in ctx for w in ["正文", "段落", "内容"]):
                if "body" not in result:
                    result["body"] = pt
            elif any(w in ctx for w in ["一级", "章标题", "大标题", "论文题目"]):
                if "title" not in result:
                    result["title"] = pt
            elif any(w in ctx for w in ["二级", "节标题"]):
                if "h2" not in result:
                    result["h2"] = pt
            elif any(w in ctx for w in ["标题", "题目"]):
                if "heading" not in result:
                    result["heading"] = pt

    # Try numeric pt specs
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*pt", text, re.IGNORECASE):
        pt = float(m.group(1))
        ctx_start = max(0, m.start() - 30)
        ctx = text[ctx_start:m.end() + 20].lower()
        if any(w in ctx for w in ["正文", "段落", "内容"]):
            if "body" not in result:
                result["body"] = pt
        elif any(w in ctx for w in ["一级", "大标题"]):
            if "title" not in result:
                result["title"] = pt
        elif any(w in ctx for w in ["标题"]):
            if "heading" not in result:
                result["heading"] = pt

    return result


def extract_line_spacing(text: str) -> dict:
    """Extract line spacing info."""
    result = {"body": None}

    for m in re.finditer(r"(\d+\.?\d*)\s*倍\s*行\s*距", text):
        val = float(m.group(1))
        ctx_start = max(0, m.start() - 15)
        ctx = text[ctx_start:m.end() + 10]
        if any(w in ctx for w in ["正文", "段落"]) or True:
            result["body"] = val

    if "单倍行距" in text and not result["body"]:
        result["body"] = 1.0
    if "双倍行距" in text and not result["body"]:
        result["body"] = 2.0

    # Footnote / ref spacing
    ref_match = re.search(r"参考文献[^。]*?(\d+\.?\d*)\s*倍", text)
    if ref_match:
        result["references"] = float(ref_match.group(1))
    if "参考文献" in text and ("单倍" in text or "single" in text.lower()):
        result["references"] = 1.0

    return result


def extract_margins(text: str) -> dict:
    """Extract page margin info."""
    result = {"top": None, "bottom": None, "left": None, "right": None}

    # 上下Xcm 左右Ycm pattern
    m = re.search(r"上\s*下\s*(\d+\.?\d*)\s*cm[，,\s]*左\s*右\s*(\d+\.?\d*)\s*cm", text)
    if m:
        result["top"] = result["bottom"] = float(m.group(1))
        result["left"] = result["right"] = float(m.group(2))

    # Individual margins
    for side, pat in [("top", r"上[边距]*\s*(\d+\.?\d*)\s*cm"),
                       ("bottom", r"下[边距]*\s*(\d+\.?\d*)\s*cm"),
                       ("left", r"左[边距]*\s*(\d+\.?\d*)\s*cm"),
                       ("right", r"右[边距]*\s*(\d+\.?\d*)\s*cm")]:
        m = re.search(pat, text)
        if m:
            result[side] = float(m.group(1))

    return {k: v for k, v in result.items() if v is not None}


def extract_indent(text: str) -> dict:
    """Extract first-line indent info."""
    result = {"body": None}

    m = re.search(r"首行缩进\s*(\d+)\s*字符", text)
    if m:
        result["body"] = f"{m.group(1)}字符"

    m = re.search(r"首行缩进\s*(\d+)\s*pt", text)
    if m:
        result["body_pt"] = int(m.group(1))

    return {k: v for k, v in result.items() if v is not None}


def extract_word_count(text: str) -> dict:
    """Extract word count requirements."""
    result = {"min": None, "max": None}

    # "不少于3000字" / "至少5000字"
    for pat in [r"不少于\s*(\d+)\s*字", r"至少\s*(\d+)\s*字",
                r"最低\s*(\d+)\s*字", r"不少于\s*(\d+)\s*(?:字|词)"]:
        m = re.search(pat, text)
        if m:
            result["min"] = int(m.group(1))
            break

    # "3000-5000字" / "3000～5000字"
    m = re.search(r"(\d+)\s*[-~～]\s*(\d+)\s*字", text)
    if m:
        result["min"] = int(m.group(1))
        result["max"] = int(m.group(2))

    # "不超过5000字"
    m = re.search(r"不超过\s*(\d+)\s*字", text)
    if m:
        result["max"] = int(m.group(1))

    # "字数要求：3000字左右"
    m = re.search(r"字数[：:要求]*\s*(\d+)\s*字", text)
    if m:
        result["min"] = int(m.group(1))

    return {k: v for k, v in result.items() if v is not None}


def extract_chapters(text: str) -> list[str]:
    """Extract required chapter structure."""
    chapters = []
    seen = set()

    # "第X章 XXX" pattern
    for m in re.finditer(r"第[一二三四五六七八九十\d]+章\s*[^\n，。；;]{1,30}", text):
        ch = m.group(0).strip()
        if ch not in seen:
            chapters.append(ch)
            seen.add(ch)

    # Common required sections mentioned without 章 prefix
    section_keywords = ["摘要", "关键词", "引言", "相关工作", "方法", "实验",
                        "总结", "展望", "参考文献", "致谢", "附录"]
    for kw in section_keywords:
        # Look for "必须包含xxx" / "应包括xxx" / "要有xxx"
        for pat in [rf"(?:必须|应|需|要)\s*(?:包含|有|包括|写)\s*{kw}",
                    rf"{kw}\s*(?:章节|部分|内容)"]:
            if re.search(pat, text) and kw not in seen:
                chapters.append(kw)
                seen.add(kw)
                break

    return chapters[:15]


def extract_reference_rules(text: str) -> dict:
    """Extract reference formatting rules."""
    result = {}

    m = re.search(r"(?:参考文献|引用)\s*(?:不少于|至少|≥)\s*(\d+)\s*(?:篇|条|个)", text)
    if m:
        result["min_count"] = int(m.group(1))

    # Language balance
    if re.search(r"中英文", text) or re.search(r"中.*英.*文献", text):
        result["language"] = "中英文混合"
    elif re.search(r"英文\s*(?:文献|参考)", text):
        result["language"] = "英文为主"
    elif re.search(r"中文\s*(?:文献|参考)", text):
        result["language"] = "中文为主"

    # Year requirements
    m = re.search(r"近\s*(\d+)\s*年", text)
    if m:
        result["year_range"] = f"近{m.group(1)}年"

    # Format style
    if "GB/T" in text or "GB7714" in text:
        result["style"] = "GB/T 7714"
    elif "APA" in text:
        result["style"] = "APA"
    elif "MLA" in text:
        result["style"] = "MLA"

    return result


def extract_other_rules(text: str) -> list[str]:
    """Extract other notable formatting rules."""
    rules = []

    if re.search(r"(?:不得|禁止|不要|不能).*抄袭", text):
        rules.append("禁止抄袭")
    if re.search(r"AIGC|AI\s*(?:生成|辅助|工具)|生成式\s*AI", text):
        rules.append("需声明 AIGC 使用情况")
    if re.search(r"查重", text):
        m = re.search(r"查重[率比].*?(\d+)%", text)
        if m:
            rules.append(f"查重率 < {m.group(1)}%")
        else:
            rules.append("需查重")
    if re.search(r"页眉|页脚|页码", text):
        rules.append("需添加页眉/页脚/页码")
    if re.search(r"目录", text):
        rules.append("需生成目录")

    return rules


# ── Main parser ───────────────────────────────────────────────

def parse_requirements(text: str, source_name: str = "") -> dict:
    """Parse a course requirement document and return structured format config."""
    font = extract_font(text)
    sizes = extract_font_sizes(text)
    spacing = extract_line_spacing(text)
    margins = extract_margins(text)
    indent = extract_indent(text)
    word_count = extract_word_count(text)
    chapters = extract_chapters(text)
    ref_rules = extract_reference_rules(text)
    other = extract_other_rules(text)

    return {
        "_source": source_name,
        "font": font,
        "font_sizes_pt": sizes,
        "line_spacing": spacing,
        "margins_cm": margins,
        "indent": indent,
        "word_count_requirement": word_count,
        "required_chapters": chapters,
        "reference_rules": ref_rules,
        "other_rules": other,
    }


# ── Output formatters ─────────────────────────────────────────

def format_for_claude(spec: dict) -> str:
    """Format the parsed spec as a human-readable summary for Claude."""
    lines = ["## 从课程要求文档提取的格式规格", ""]

    if spec.get("font"):
        lines.append("### 字体要求")
        for k, v in spec["font"].items():
            label = {"body": "正文", "heading": "标题", "title": "大标题"}.get(k, k)
            lines.append(f"- {label}：{v}")
        lines.append("")

    if spec.get("font_sizes_pt"):
        lines.append("### 字号要求")
        for k, v in spec["font_sizes_pt"].items():
            label = {"body": "正文", "heading": "标题", "title": "大标题", "h2": "二级标题"}.get(k, k)
            lines.append(f"- {label}：{v}pt")
        lines.append("")

    if spec.get("line_spacing"):
        lines.append("### 行距要求")
        for k, v in spec["line_spacing"].items():
            label = {"body": "正文", "references": "参考文献"}.get(k, k)
            lines.append(f"- {label}：{v}倍行距")
        lines.append("")

    if spec.get("margins_cm"):
        lines.append("### 页边距")
        margins = spec["margins_cm"]
        lines.append(f"- 上 {margins.get('top', '?')}cm 下 {margins.get('bottom', '?')}cm")
        lines.append(f"- 左 {margins.get('left', '?')}cm 右 {margins.get('right', '?')}cm")
        lines.append("")

    if spec.get("indent"):
        lines.append("### 缩进")
        for k, v in spec["indent"].items():
            if k.endswith("_pt"):
                continue
            lines.append(f"- {k}：{v}")
        lines.append("")

    if spec.get("word_count_requirement"):
        wc = spec["word_count_requirement"]
        lines.append("### 字数要求")
        if wc.get("min") and wc.get("max"):
            lines.append(f"- {wc['min']}-{wc['max']} 字")
        elif wc.get("min"):
            lines.append(f"- 不少于 {wc['min']} 字")
        elif wc.get("max"):
            lines.append(f"- 不超过 {wc['max']} 字")
        lines.append("")

    if spec.get("required_chapters"):
        lines.append("### 要求章节结构")
        for i, ch in enumerate(spec["required_chapters"], 1):
            lines.append(f"{i}. {ch}")
        lines.append("")

    if spec.get("reference_rules"):
        lines.append("### 参考文献要求")
        for k, v in spec["reference_rules"].items():
            label = {"min_count": "最少篇数", "language": "语言要求",
                     "year_range": "年份范围", "style": "引用格式"}.get(k, k)
            lines.append(f"- {label}：{v}")
        lines.append("")

    if spec.get("other_rules"):
        lines.append("### 其他要求")
        for r in spec["other_rules"]:
            lines.append(f"- {r}")
        lines.append("")

    return "\n".join(lines)


# ── CLI ──────────────────────────────────────────────────────

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Parse a course requirement document and extract formatting specs."
    )
    ap.add_argument("file", type=Path, help="Requirement document (.md/.txt/.pdf/.docx)")
    ap.add_argument("--out", type=Path, default=None,
                    help="Output path (prints to stdout if omitted)")
    ap.add_argument("--format", choices=["yaml", "text", "summary"],
                    default="summary", help="Output format")
    args = ap.parse_args()

    path: Path = args.file.resolve()
    if not path.exists():
        print(f"Error: {path} not found", file=sys.stderr)
        return 1

    ext = path.suffix.lower()
    if ext in (".md", ".txt", ".markdown"):
        text = read_text(path)
    elif ext in (".docx", ".doc"):
        text = read_docx(path)
    elif ext == ".pdf":
        text = read_pdf(path)
    else:
        print(f"Error: unsupported format {ext}", file=sys.stderr)
        return 1

    if not text.strip():
        print(f"Error: could not extract text from {path}", file=sys.stderr)
        return 1

    spec = parse_requirements(text, path.name)

    if args.format == "summary":
        output = format_for_claude(spec)
    elif args.format == "yaml":
        import yaml
        # Remove internal key
        clean = {k: v for k, v in spec.items() if not k.startswith("_")}
        output = yaml.dump(clean, allow_unicode=True, default_flow_style=False, sort_keys=False)
    else:
        output = format_for_claude(spec)

    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(output, encoding="utf-8")
        print(f"✓ 格式规格已写入 {args.out}")
    else:
        print(output)

    # Summary to stderr
    detected = []
    if spec.get("font"):
        detected.append(f"字体：{spec['font']}")
    if spec.get("font_sizes_pt"):
        detected.append(f"字号：{spec['font_sizes_pt']}")
    if spec.get("line_spacing"):
        detected.append(f"行距：{spec['line_spacing']}")
    if spec.get("word_count_requirement"):
        detected.append(f"字数：{spec['word_count_requirement']}")
    if spec.get("required_chapters"):
        detected.append(f"章节：{len(spec['required_chapters'])} 章")
    if detected:
        print(f"\n📋 从「{path.name}」检测到: {' | '.join(detected)}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
