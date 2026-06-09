---
name: ai-course-project
description: Write formatted Chinese course reports for AI/CS projects. Convert source code, experiments, figures, and references into a well-structured Markdown→DOCX paper. Use when the user asks to write an AI course project report, generate a thesis, format a paper, or produce a project write-up with evidence from their code and data.
---

# AI Course Project Workbench

轻量级 AI 课程论文工作台——从项目源码、实验数据、图表截图和参考文献，结构化地写成规范的课程论文 DOCX。

## 适用场景

- 人工智能导论、机器学习、深度学习等课程大作业
- 有真实源码和实验数据的项目报告
- 需要按学术格式输出 DOCX 的课程论文

## 工作流（3 步）

### Step 1: 初始化 + 提取证据

```bash
# 在项目根目录初始化工作区
python scripts/init_project.py .

# 编辑论文元信息
vim templates/thesis-spec.yaml

# 从源码中提取证据
python scripts/build_evidence.py . --out paper-context/evidence
```

### Step 2: 写 Markdown 论文

在 `paper-output/<论文标题>.md` 中按推荐结构写作：

```markdown
# 论文标题

## 摘要
摘要正文...

## 关键词
关键词：xxx；xxx；xxx

## 第1章 引言
...

## 第2章 相关工作
...

## 第3章 方法
...

## 第4章 实验
...

## 第5章 总结与展望
...

## 参考文献
[1] ...
[2] ...
```

图表使用：
```markdown
![模型架构图](figures/model-arch.png)

| 方法 | 准确率 | F1 |
|------|--------|----|
| Baseline | 0.82 | 0.79 |
| Ours | 0.89 | 0.86 |

表4-1 实验结果对比
```

LaTeX 公式：
```markdown
$$E = mc^2$$

$$
L = -\frac{1}{N}\sum_{i=1}^N y_i \log(\hat{y}_i)
$$
```

分步写作可参考 `references/chapter-patterns.md` 中每章的详细指南。

### Step 3: 生成 DOCX

```bash
# 管理图片映射
python scripts/build_image_map.py --scan paper-output/figures/ paper-output/image-map.json

# 生成 DOCX
python scripts/generate_docx.py paper-output/thesis.md paper-output/thesis.docx \
  --image-map paper-output/image-map.json

# 检查字数
python scripts/count_words.py paper-output/thesis.md

# 检查参考文献
python scripts/check_references.py paper-output/thesis.md
```

### Step 4: 生成待办清单（输出后必做）

DOCX 生成后，**必须**执行此步骤——解析论文中的截图占位符，生成一份待办清单 DOCX 到用户桌面：

```bash
python scripts/generate_task_checklist.py paper-output/thesis.md
```

**输出内容**：
- 在桌面生成 `<论文标题>-待办清单.docx`
- 列出所有 `【待补：xxx】` 截图任务及目标路径
- 给出 image-map 更新命令
- 最终交付检查清单（带勾选框）
- 关键文件路径速查表

**然后告诉用户**：
1. 需要截图的具体内容、来源文件、存放路径
2. 完成后将文件路径提供给你，你重新生成完整 DOCX

## 脚本速览

| 脚本 | 用途 |
|------|------|
| `init_project.py` | 一键初始化论文工作区 |
| `build_evidence.py` | 扫描项目源码，提取技术栈/API/Schema/测试证据 |
| `generate_docx.py` | Markdown → 格式化 DOCX |
| `build_image_map.py` | 将图片目录映射为 JSON，供 DOCX 生成器使用 |
| `render_mermaid.py` | 渲染 Markdown 中的 Mermaid 图表为 PNG |
| `count_words.py` | 按章统计字数 |
| `check_references.py` | 检查参考文献编号、年份、语言分布 |
| `generate_task_checklist.py` | **新增** — 解析缺失素材，生成桌面待办清单 DOCX |

## 关键原则

1. **不编造事实**：正文只能引用 `paper-context/evidence/` 中有据可查的内容
2. **图表可追溯**：每张图/表在 `templates/figure-registry.yaml` 中记录来源和状态
3. **代码绑定**：方法章节必须绑定真实模块、截图、核心逻辑
4. **可复现**：实验设置足够详细，让读者能复现你的结果

## 格式默认值

格式已按学校要求配置（宋体标题 + 1.5 倍行距）：

| 元素 | 字体 | 字号 | 备注 |
|------|------|------|------|
| 大标题 | 宋体 + Times New Roman | 22pt (二号) | 加粗居中 |
| 一级标题 | 宋体 | 16pt (三号) | 加粗 |
| 二级标题 | 宋体 | 14pt (四号) | 加粗 |
| 正文 | 宋体 + Times New Roman | 12pt (小四) | **1.5 倍行距**，首行缩进 24pt |
| 图表标题 | 宋体 | 10.5pt (五号) | 单倍行距居中 |
| 参考文献 | 宋体 | 10.5pt | 悬挂缩进 |
| 代码 | Consolas | 9pt | |
| 公式 | Times New Roman | 12pt | 居中 |

> 如需调整，编辑 `scripts/generate_docx.py` 中 `build_style_profile()` 函数。

## 依赖

```bash
pip install python-docx PyYAML
```

可选（Mermaid 图表渲染）：
```bash
npm i -g @mermaid-js/mermaid-cli
```
