# ai-course-project-skill

> AI Course Thesis Workbench: Source Code → Evidence → Markdown → DOCX
> 轻量级 AI 课程论文工作台：从项目源码到格式规范的课程论文 DOCX，全自动管道

[![Python](https://img.shields.io/badge/Python-3.8+-blue)](scripts/)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
[![Lines](https://img.shields.io/badge/Total%20Lines-~800-green)](scripts/)

[English](#english) | [中文](#中文)

---

## English

### Overview

A Claude Code skill that automates AI/CS course project report generation. It scans your project source code, extracts technical evidence (tech stack, API endpoints, database schemas, test coverage), and produces a professionally formatted DOCX file — all driven by Markdown and YAML configuration.

### Pipeline

```
Source Code  →  build_evidence.py  →  Evidence (JSON + MD)
     │                                      │
     │                               You write thesis.md
     │                               (guided by chapter-patterns.md)
     │                                      │
     └──────────────┬───────────────────────┘
                    ▼
           generate_docx.py
                    │
                    ▼
              thesis.docx
    (Chinese academic formatting:
     Song Ti body, Hei Ti headings,
     1.5x line spacing, LaTeX formulas,
     Mermaid diagrams, GFM tables)
```

### Key Features

- **Evidence Extraction**: Automatically detects framework, API routes, DB schemas, test coverage from source
- **Markdown → DOCX**: Regex-based Markdown parser with GFM table support, LaTeX formulas, Mermaid diagrams
- **Chinese Academic Formatting**: Song Ti / Hei Ti fonts, first-line indent, 1.5x line spacing, page break management
- **Modular Scripts**: 7 independent scripts, each with argparse CLI
- **Zero Node.js**: Python-only (optional Mermaid CLI for diagrams)

### Tech Stack

| Script | Lines | Purpose |
|--------|-------|---------|
| `build_evidence.py` | 193 | Source code scanning + evidence extraction |
| `generate_docx.py` | 286 | Markdown to DOCX core converter |
| `init_project.py` | ~60 | One-click workspace initialization |
| `build_image_map.py` | ~50 | Image directory to JSON mapping |
| `render_mermaid.py` | ~70 | Mermaid to PNG rendering |
| `count_words.py` | ~50 | Chapter word count |
| `check_references.py` | ~80 | Reference validation |

---

## 中文

### 概述

从项目源码、实验数据、图表和参考文献，结构化地生成格式规范的课程论文 DOCX。

### 与完整版的关系

本 skill 是从 [chinese-thesis-workbench-skill](https://github.com/ZyhSechub/chinese-thesis-workbench-skill) 裁剪而来，保留了核心的 **证据提取 → Markdown 写作 → DOCX 生成** 链路，去掉了毕业论文学位论文才需要的：

- M2/M3/M4 DOCX 路径（模板填充、样文贴近、已有稿编辑）
- 14-phase 状态机和阻断报告机制
- AIGC 风格治理
- Word 批注修订
- Playwright 自动截图
- PDF 参考文献抽取

**保留 7 个核心脚本 → ~800 行 Python**，零 Node.js 依赖（可选 Mermaid CLI 除外）。

### 工程设计亮点

- **数据类设计**: 使用 `@dataclass` 实现结构化证据提取，类型安全
- **智能文件扫描**: 自动忽略 `node_modules/`、`target/`、`__pycache__` 等构建目录
- **框架自动检测**: 通过 `package.json`、`pyproject.toml`、`go.mod` 等标记文件识别技术栈
- **正则解析器**: 完整的 GFM Markdown 解析器，支持表格、代码块、LaTeX 公式
- **OOXML 深度操作**: 通过 python-docx 直接操作 XML 元素实现中文字体设置

### 快速开始

```bash
# 1. 安装依赖
pip install python-docx PyYAML

# 2. 在你的项目根目录初始化
python scripts/init_project.py /path/to/your-project

# 3. 编辑论文元信息
vim /path/to/your-project/templates/thesis-spec.yaml

# 4. 提取项目证据
python scripts/build_evidence.py /path/to/your-project

# 5. 写 Markdown 论文（参考 references/chapter-patterns.md）

# 6. 生成 DOCX
python scripts/generate_docx.py \
  paper-output/thesis.md \
  paper-output/thesis.docx \
  --image-map paper-output/image-map.json
```

### 目录结构

```
ai-course-project-skill/
├── SKILL.md                    ← Claude Code Skill 入口
├── README.md                   ← 本文件
├── requirements.txt
│
├── scripts/
│   ├── init_project.py         ← 一键初始化工作区
│   ├── build_evidence.py       ← 扫描源码，提取证据
│   ├── generate_docx.py        ← Markdown → DOCX（核心）
│   ├── build_image_map.py      ← 图片目录 → JSON 映射
│   ├── render_mermaid.py       ← Mermaid → PNG
│   ├── count_words.py          ← 按章统计字数
│   └── check_references.py     ← 参考文献检查
│
├── templates/
│   ├── thesis-spec.yaml        ← 论文元信息模板
│   └── figure-registry.yaml    ← 图表注册表模板
│
└── references/
    └── chapter-patterns.md      ← AI 课程论文章节写作指南
```

### 你的项目结构（初始化后）

```
your-project/
├── paper-output/               ← 交付物
│   ├── thesis.md               ← 你写的 Markdown 论文
│   ├── thesis.docx             ← 生成的 DOCX
│   ├── image-map.json          ← 图片映射
│   ├── figures/                ← 图表 PNG
│   └── screenshots/            ← 截图
│
├── paper-context/              ← 证据材料
│   ├── evidence/               ← 从源码提取的证据
│   │   ├── project-evidence.json
│   │   ├── tech-stack.md
│   │   ├── code-structure.md
│   │   └── api-list.md
│   └── literature/             ← 参考文献
│
└── templates/                  ← 你的论文配置
    ├── thesis-spec.yaml
    └── figure-registry.yaml
```

## License

MIT License — 基于 [chinese-thesis-workbench-skill](https://github.com/ZyhSechub/chinese-thesis-workbench-skill) 裁剪。

## Author

**Lambert Liu** (CS231-Alambert)
- GitHub: [@CS231-Alambert](https://github.com/CS231-Alambert)
- 计算机科学与技术（中外合作办学）
