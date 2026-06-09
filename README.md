# ai-course-project-skill

> 轻量级 AI 课程论文工作台：源码 → 证据 → Markdown → DOCX

从项目源码、实验数据、图表和参考文献，结构化地生成格式规范的课程论文 DOCX。

## 与完整版的关系

本 skill 是从 [chinese-thesis-workbench-skill](https://github.com/ZyhSechub/chinese-thesis-workbench-skill) 裁剪而来，保留了核心的 **证据提取 → Markdown 写作 → DOCX 生成** 链路，去掉了毕业论文学位论文才需要的：

- M2/M3/M4 DOCX 路径（模板填充、样文贴近、已有稿编辑）
- 14-phase 状态机和阻断报告机制
- AIGC 风格治理
- Word 批注修订
- Playwright 自动截图
- PDF 参考文献抽取

**保留 7 个核心脚本 → ~800 行 Python**，零 Node.js 依赖（可选 Mermaid CLI 除外）。

## 快速开始

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

## 目录结构

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

## 你的项目结构（初始化后）

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

## 许可

MIT License — 基于 [chinese-thesis-workbench-skill](https://github.com/ZyhSechub/chinese-thesis-workbench-skill) 裁剪。
