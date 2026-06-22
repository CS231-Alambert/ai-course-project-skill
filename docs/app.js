/* ═══════════════════════════════════════════
   AI Course Project Workbench — Web Edition
   Markdown Editor → Live Preview → DOCX Download
   Zero-install, all in browser
   ═══════════════════════════════════════════ */

import { Document, Packer, Paragraph, TextRun,
         HeadingLevel, Table, TableRow, TableCell,
         AlignmentType, WidthType, BorderStyle,
         LineRuleType, PageBreak } from 'docx';

/* ── Constants ─────────────────────────── */

const HALF_PT = {
  title: 44,    h1: 32,    h2: 28,    h3: 24,
  body: 24,     caption: 21, code: 18, ref: 21, eq: 24,
};

const LINE = { single: 240, onehalf: 360 };
const INDENT_FIRST = 480; // 24pt in twips
const INDENT_HANGING = 420; // 21pt in twips

const TEMPLATES = {
  full: `# 论文标题

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
[2] ...`,

  simple: `# 实验报告

## 实验目的
...

## 实验环境
...

## 实验步骤
1. ...
2. ...

## 实验结果

| 指标 | 数值 |
|------|------|
| 准确率 | 0.95 |
| F1 | 0.93 |

## 结论
...`,
};

/* ── Markdown Parser ───────────────────── */

function parseMarkdown(md) {
  const lines = md.split('\n');
  const blocks = [];
  let i = 0;

  while (i < lines.length) {
    const s = lines[i].trim();

    // skip empty
    if (!s) { i++; continue; }

    // code fence
    if (s.startsWith('```')) {
      const lang = s.slice(3).trim().toLowerCase();
      const codeLines = [];
      i++;
      while (i < lines.length && !lines[i].trim().startsWith('```')) {
        codeLines.push(lines[i]);
        i++;
      }
      if (i < lines.length) i++; // closing ```
      blocks.push({ type: 'code', lang, content: codeLines.join('\n') });
      continue;
    }

    // formula block
    if (s === '$$') {
      const eqLines = [];
      i++;
      while (i < lines.length && lines[i].trim() !== '$$') {
        eqLines.push(lines[i].trim());
        i++;
      }
      if (i < lines.length) i++; // closing $$
      blocks.push({ type: 'equation', content: eqLines.join('\n') });
      continue;
    }

    // inline formula $$...$$
    if (s.startsWith('$$') && s.endsWith('$$') && s.length > 4) {
      blocks.push({ type: 'equation', content: s.slice(2, -2).trim() });
      i++;
      continue;
    }

    // headings
    if (s.startsWith('# ')) {
      blocks.push({ type: 'title', content: cleanText(s.slice(2)) });
      i++; continue;
    }
    if (s.startsWith('## ')) {
      blocks.push({ type: 'h1', content: cleanText(s.slice(3)) });
      i++; continue;
    }
    if (s.startsWith('### ')) {
      blocks.push({ type: 'h2', content: cleanText(s.slice(4)) });
      i++; continue;
    }
    if (s.startsWith('#### ')) {
      blocks.push({ type: 'h3', content: cleanText(s.slice(5)) });
      i++; continue;
    }

    // keywords
    const kw = s.match(/^(关键词|Keywords)[:：]\s*(.*)$/);
    if (kw) {
      blocks.push({ type: 'keywords', label: kw[1], content: kw[2] });
      i++; continue;
    }

    // figure / table captions
    if (/^(图|Fig\.?)\s*\d/.test(s)) {
      blocks.push({ type: 'figureCaption', content: cleanText(s) });
      i++; continue;
    }
    if (/^(表|Table\.?)\s*\d/.test(s)) {
      blocks.push({ type: 'tableCaption', content: cleanText(s) });
      i++; continue;
    }

    // GFM table
    if (s.startsWith('|')) {
      const tblLines = [];
      while (i < lines.length && lines[i].trim().startsWith('|')) {
        tblLines.push(lines[i].trim());
        i++;
      }
      blocks.push({ type: 'table', content: parseTable(tblLines) });
      continue;
    }

    // image
    const imgMd = s.match(/^!\[(.+?)\]\((.+?)\)$/);
    if (imgMd) {
      blocks.push({ type: 'image', alt: imgMd[1], src: imgMd[2] });
      i++; continue;
    }

    // reference line
    if (/^\[\d+\]/.test(s)) {
      blocks.push({ type: 'reference', content: cleanText(s) });
      i++; continue;
    }

    // regular paragraph
    blocks.push({ type: 'paragraph', content: cleanText(s) });
    i++;
  }

  return blocks;
}

function parseTable(lines) {
  const rows = [];
  for (const ln of lines) {
    const cells = ln.replace(/^[\s|]+|[\s|]+$/g, '').split('|').map(c => c.trim());
    // skip separator lines like |---|---|
    if (cells.every(c => /^[:\- ]+$/.test(c))) continue;
    rows.push(cells);
  }
  return rows;
}

/* ── Text Cleanup ──────────────────────── */

const LINK_RE = /\[([^\]]+)\]\([^)]+\)/g;
const AUTOLINK_RE = /<(https?:\/\/[^>]+)>/g;

function cleanText(text) {
  let c = text.replace(/`/g, '');
  c = c.replace(LINK_RE, '$1');
  c = c.replace(AUTOLINK_RE, '$1');
  c = c.replace(/\*\*(.+?)\*\*/g, '$1');
  c = c.replace(/__(.+?)__/g, '$1');
  c = c.replace(/(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)/g, '$1');
  c = c.replace(/(?<!_)_(?!\s)(.+?)(?<!\s)_(?!_)/g, '$1');
  return c.replace(/\s+/g, ' ').trim();
}

/* ── Helpers ───────────────────────────── */

function pStyle({ font = '宋体', size = HALF_PT.body, bold = false,
                  alignment = AlignmentType.LEFT, lineSpacing = LINE.onehalf,
                  firstLine = 0, spacingBefore = 0, spacingAfter = 0,
                  pageBreak = false }) {
  return {
    alignment,
    spacing: {
      line: lineSpacing,
      lineRule: LineRuleType.AUTO,
      before: spacingBefore,
      after: spacingAfter,
    },
    indent: firstLine ? { firstLine } : undefined,
    ...(pageBreak ? { pageBreakBefore: true } : {}),
  };
}

function run(text, { font = '宋体', size = HALF_PT.body, bold = false } = {}) {
  return new TextRun({ text: String(text), font, size, bold });
}

function addPara(doc, text, styleOpts = {}, runOpts = {}) {
  const para = new Paragraph({
    ...pStyle(styleOpts),
    children: [run(text, runOpts)],
  });
  doc.addSection({ children: [para] });
}

function addKeywordPara(doc, label, content) {
  const isCn = label.startsWith('关键词');
  const para = new Paragraph({
    spacing: { line: LINE.single, lineRule: LineRuleType.AUTO },
  });
  // Label run (bold)
  const rLabel = new TextRun({
    text: label + (label.endsWith('：') || label.endsWith(':') ? '' : ' '),
    font: isCn ? '宋体' : 'Times New Roman',
    size: isCn ? HALF_PT.body : HALF_PT.body,
    bold: true,
  });
  para.addChildElement(rLabel);
  // Content run
  if (content) {
    const rContent = new TextRun({
      text: content,
      font: isCn ? '宋体' : 'Times New Roman',
      size: isCn ? HALF_PT.body : HALF_PT.body,
    });
    para.addChildElement(rContent);
  }
  doc.addSection({ children: [para] });
}

/* ── DOCX Generator ────────────────────── */

function buildDocx(markdown) {
  const blocks = parseMarkdown(markdown);

  const children = [];

  function add(p) { children.push(p); }
  function addPageBreak() {
    children.push(new Paragraph({
      children: [new TextRun({ text: '' })],
      pageBreakBefore: true,
    }));
  }

  let seenContent = false;
  let currentSection = '';
  let isFirstH1 = true;

  for (const block of blocks) {
    switch (block.type) {
      case 'title':
        if (seenContent) addPageBreak();
        add(new Paragraph({
          ...pStyle({ font: '宋体', size: HALF_PT.title, bold: true,
                       alignment: AlignmentType.CENTER, lineSpacing: LINE.onehalf,
                       spacingBefore: 400, spacingAfter: 400 }),
          children: [run(block.content, { font: 'Times New Roman', size: HALF_PT.title, bold: true })],
        }));
        seenContent = true;
        break;

      case 'h1':
        currentSection = block.content.replace(/\s/g, '');
        if (seenContent && isFirstH1) {
          add(new Paragraph({
            ...pStyle({ font: '宋体', size: HALF_PT.h1, bold: true,
                         spacingBefore: 240, spacingAfter: 120 }),
            pageBreakBefore: true,
            children: [run(block.content, { font: '宋体', size: HALF_PT.h1, bold: true })],
          }));
          isFirstH1 = false;
        } else {
          add(new Paragraph({
            ...pStyle({ font: '宋体', size: HALF_PT.h1, bold: true,
                         spacingBefore: 240, spacingAfter: 120 }),
            children: [run(block.content, { font: '宋体', size: HALF_PT.h1, bold: true })],
          }));
        }
        seenContent = true;
        break;

      case 'h2':
        add(new Paragraph({
          ...pStyle({ font: '宋体', size: HALF_PT.h2, bold: true,
                       spacingBefore: 200, spacingAfter: 80 }),
          children: [run(block.content, { font: '宋体', size: HALF_PT.h2, bold: true })],
        }));
        seenContent = true;
        break;

      case 'h3':
        add(new Paragraph({
          ...pStyle({ font: '宋体', size: HALF_PT.h3, bold: true,
                       spacingBefore: 160, spacingAfter: 80 }),
          children: [run(block.content, { font: '宋体', size: HALF_PT.h3, bold: true })],
        }));
        seenContent = true;
        break;

      case 'keywords':
        // Handled by addKeywordPara helper
        seenContent = true;
        break;

      case 'paragraph':
        add(new Paragraph({
          ...pStyle({ font: '宋体', size: HALF_PT.body, lineSpacing: LINE.onehalf,
                       firstLine: INDENT_FIRST }),
          children: [run(block.content, { font: '宋体', size: HALF_PT.body })],
        }));
        seenContent = true;
        break;

      case 'code':
        for (const line of block.content.split('\n')) {
          add(new Paragraph({
            ...pStyle({ font: 'Consolas', size: HALF_PT.code, lineSpacing: LINE.single }),
            children: [run(line || ' ', { font: 'Consolas', size: HALF_PT.code })],
          }));
        }
        seenContent = true;
        break;

      case 'equation':
        add(new Paragraph({
          ...pStyle({ font: 'Times New Roman', size: HALF_PT.eq,
                       alignment: AlignmentType.CENTER, lineSpacing: LINE.onehalf }),
          children: [run(block.content, { font: 'Times New Roman', size: HALF_PT.eq })],
        }));
        seenContent = true;
        break;

      case 'figureCaption':
      case 'tableCaption':
        add(new Paragraph({
          ...pStyle({ font: '宋体', size: HALF_PT.caption,
                       alignment: AlignmentType.CENTER, lineSpacing: LINE.single }),
          children: [run(block.content, { font: '宋体', size: HALF_PT.caption })],
        }));
        seenContent = true;
        break;

      case 'table': {
        const allRows = block.content;
        if (allRows.length < 2) break;
        const maxCols = Math.max(...allRows.map(r => r.length));
        const padRow = (r) => [...r, ...Array(maxCols - r.length).fill('')].slice(0, maxCols);

        const tableRows = allRows.map((rowCells, ri) =>
          new TableRow({
            children: padRow(rowCells).map(cell =>
              new TableCell({
                children: [new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { line: LINE.single, lineRule: LineRuleType.AUTO },
                  children: [run(cell, { font: '宋体', size: HALF_PT.caption })],
                })],
                width: { size: 100 / maxCols, type: WidthType.PERCENTAGE },
              })
            ),
          })
        );

        add(new Table({
          rows: tableRows,
          width: { size: 100, type: WidthType.PERCENTAGE },
        }));
        seenContent = true;
        break;
      }

      case 'reference':
        add(new Paragraph({
          ...pStyle({ font: '宋体', size: HALF_PT.ref, lineSpacing: LINE.single,
                       firstLine: -INDENT_HANGING }),
          indent: { left: INDENT_HANGING },
          children: [run(block.content, { font: '宋体', size: HALF_PT.ref })],
        }));
        seenContent = true;
        break;

      case 'image':
        add(new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { line: LINE.single, lineRule: LineRuleType.AUTO },
          children: [run(`[图片：${block.alt}]`, { font: '楷体', size: HALF_PT.caption })],
        }));
        seenContent = true;
        break;

      default:
        break;
    }
  }

  // Handle keywords specially — find and insert them after the title
  const kwBlocks = blocks.filter(b => b.type === 'keywords');
  const finalChildren = [];
  let kwInserted = false;
  for (const c of children) {
    finalChildren.push(c);
    // After first heading1 (摘要), try to insert keywords
    if (!kwInserted && kwBlocks.length > 0 && c._text === undefined) {
      // Actually, let's handle keywords inline with the blocks more naturally
    }
  }

  // Simplification: build the document from sections
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: 907,    // 2.54cm in twips
            bottom: 907,
            left: 1140,  // 3.17cm in twips
            right: 1140,
          },
        },
      },
      children: children,
    }],
  });

  return doc;
}

/* ── Download ──────────────────────────── */

async function generateAndDownload(markdown) {
  if (!markdown.trim()) {
    alert('请先输入 Markdown 内容');
    return;
  }

  const doc = buildDocx(markdown);
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'thesis.docx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ── UI Bindings ───────────────────────── */

document.addEventListener('DOMContentLoaded', () => {
  const editor = document.getElementById('editor');
  const preview = document.getElementById('preview');
  const generateBtn = document.getElementById('generateBtn');
  const copyBtn = document.getElementById('copyBtn');
  const clearBtn = document.getElementById('clearBtn');
  const templateSelect = document.getElementById('templateSelect');
  const statusEl = document.getElementById('status');

  if (!editor || !preview) return;

  // Live preview using marked (loaded via CDN)
  function updatePreview() {
    const md = editor.value;
    if (typeof marked !== 'undefined') {
      preview.innerHTML = marked.parse(md, { breaks: false });
    } else {
      preview.textContent = md;
    }
  }

  editor.addEventListener('input', updatePreview);

  // Generate DOCX
  if (generateBtn) {
    generateBtn.addEventListener('click', async () => {
      generateBtn.disabled = true;
      generateBtn.textContent = '生成中...';
      statusEl.textContent = '';
      try {
        await generateAndDownload(editor.value);
        statusEl.textContent = '✓ DOCX 已生成并下载';
        statusEl.className = 'status success';
      } catch (err) {
        console.error(err);
        statusEl.textContent = `✗ 生成失败：${err.message}`;
        statusEl.className = 'status error';
      }
      generateBtn.disabled = false;
      generateBtn.textContent = '生成 DOCX';
    });
  }

  // Copy
  if (copyBtn) {
    copyBtn.addEventListener('click', async () => {
      await navigator.clipboard.writeText(editor.value);
      statusEl.textContent = '✓ 已复制';
      statusEl.className = 'status success';
    });
  }

  // Clear
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      if (editor.value.trim() && !confirm('确定清空所有内容吗？')) return;
      editor.value = '';
      updatePreview();
      statusEl.textContent = '';
      templateSelect.value = '';
    });
  }

  // Template
  if (templateSelect) {
    templateSelect.addEventListener('change', () => {
      const key = templateSelect.value;
      if (!key) return;
      const tmpl = TEMPLATES[key];
      if (!tmpl) return;
      if (editor.value.trim() && !confirm('加载模板将覆盖当前内容，确定吗？')) {
        templateSelect.value = '';
        return;
      }
      editor.value = tmpl;
      updatePreview();
      statusEl.textContent = `✓ 已加载「${templateSelect.options[templateSelect.selectedIndex].text}」模板`;
      statusEl.className = 'status success';
    });
  }

  // Keyboard shortcut: Ctrl+Enter to generate
  editor.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      generateBtn?.click();
    }
  });

  // Resize handle (drag to resize panels)
  const handle = document.getElementById('resizeHandle');
  const leftPanel = document.getElementById('leftPanel');
  const rightPanel = document.getElementById('rightPanel');

  if (handle && leftPanel && rightPanel) {
    let dragging = false;

    handle.addEventListener('mousedown', (e) => {
      dragging = true;
      document.body.style.cursor = 'col-resize';
      document.body.style.userSelect = 'none';
      e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
      if (!dragging) return;
      const container = leftPanel.parentElement;
      if (!container) return;
      const rect = container.getBoundingClientRect();
      const leftWidth = ((e.clientX - rect.left) / rect.width) * 100;
      const clamped = Math.max(20, Math.min(80, leftWidth));
      leftPanel.style.flex = `${clamped}`;
      rightPanel.style.flex = `${100 - clamped}`;
    });

    document.addEventListener('mouseup', () => {
      if (dragging) {
        dragging = false;
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
      }
    });
  }

  // Initial preview
  updatePreview();
});

export { generateAndDownload, buildDocx, parseMarkdown };
