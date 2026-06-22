/* ═══════════════════════════════════════════
   AI Course Project Workbench — Web Edition
   Drag & Drop → Markdown Editor → Live Preview → DOCX
   Zero-install, all in browser
   ═══════════════════════════════════════════ */

import { Document, Packer, Paragraph, TextRun,
         Table, TableRow, TableCell,
         AlignmentType, WidthType,
         LineRuleType } from 'docx';

/* ── Constants ─────────────────────────── */

const HALF_PT = {
  title: 44,    h1: 32,    h2: 28,    h3: 24,
  body: 24,     caption: 21, code: 18, ref: 21, eq: 24,
};
const LINE = { single: 240, onehalf: 360 };
const INDENT_FIRST = 480;
const INDENT_HANGING = 420;

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
      if (i < lines.length) i++;
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
      if (i < lines.length) i++;
      blocks.push({ type: 'equation', content: eqLines.join('\n') });
      continue;
    }

    // inline formula
    if (s.startsWith('$$') && s.endsWith('$$') && s.length > 4) {
      blocks.push({ type: 'equation', content: s.slice(2, -2).trim() });
      i++; continue;
    }

    // headings
    if (s.startsWith('# ')) { blocks.push({ type: 'title', content: cleanText(s.slice(2)) }); i++; continue; }
    if (s.startsWith('## ')) { blocks.push({ type: 'h1', content: cleanText(s.slice(3)) }); i++; continue; }
    if (s.startsWith('### ')) { blocks.push({ type: 'h2', content: cleanText(s.slice(4)) }); i++; continue; }
    if (s.startsWith('#### ')) { blocks.push({ type: 'h3', content: cleanText(s.slice(5)) }); i++; continue; }

    // keywords
    const kw = s.match(/^(关键词|Keywords)[:：]\s*(.*)$/);
    if (kw) { blocks.push({ type: 'keywords', label: kw[1], content: kw[2] }); i++; continue; }

    // figure / table captions
    if (/^(图|Fig\.?)\s*\d/.test(s)) { blocks.push({ type: 'figureCaption', content: cleanText(s) }); i++; continue; }
    if (/^(表|Table\.?)\s*\d/.test(s)) { blocks.push({ type: 'tableCaption', content: cleanText(s) }); i++; continue; }

    // GFM table
    if (s.startsWith('|')) {
      const tblLines = [];
      while (i < lines.length && lines[i].trim().startsWith('|')) { tblLines.push(lines[i].trim()); i++; }
      blocks.push({ type: 'table', content: parseTable(tblLines) });
      continue;
    }

    // image
    const imgMd = s.match(/^!\[(.+?)\]\((.+?)\)$/);
    if (imgMd) { blocks.push({ type: 'image', alt: imgMd[1], src: imgMd[2] }); i++; continue; }

    // reference
    if (/^\[\d+\]/.test(s)) { blocks.push({ type: 'reference', content: cleanText(s) }); i++; continue; }

    // paragraph
    blocks.push({ type: 'paragraph', content: cleanText(s) });
    i++;
  }
  return blocks;
}

function parseTable(lines) {
  const rows = [];
  for (const ln of lines) {
    const cells = ln.replace(/^[\s|]+|[\s|]+$/g, '').split('|').map(c => c.trim());
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

function run(text, { font = '宋体', size = HALF_PT.body, bold = false } = {}) {
  return new TextRun({ text: String(text), font, size, bold });
}

function pStyle(opts = {}) {
  const { font = '宋体', size = HALF_PT.body, bold = false,
          alignment = AlignmentType.LEFT, lineSpacing = LINE.onehalf,
          firstLine = 0, before = 0, after = 0, pageBreak = false } = opts;
  return {
    alignment,
    spacing: { line: lineSpacing, lineRule: LineRuleType.AUTO, before, after },
    indent: firstLine ? { firstLine } : undefined,
    ...(pageBreak ? { pageBreakBefore: true } : {}),
  };
}

function makePara(text, styleOpts = {}, runOpts = {}) {
  return new Paragraph({
    ...pStyle(styleOpts),
    children: [run(text, { font: styleOpts.font || '宋体', size: styleOpts.size || HALF_PT.body, bold: styleOpts.bold || false, ...runOpts })],
  });
}

/* ── DOCX Generator ────────────────────── */

function buildDocx(markdown) {
  const blocks = parseMarkdown(markdown);
  const children = [];
  let seenContent = false;

  for (const block of blocks) {
    switch (block.type) {

      case 'title':
        if (seenContent) {
          children.push(new Paragraph({
            ...pStyle(), pageBreakBefore: true,
            children: [new TextRun({ text: '' })],
          }));
        }
        children.push(makePara(block.content, {
          font: '宋体', size: HALF_PT.title, bold: true,
          alignment: AlignmentType.CENTER, lineSpacing: LINE.onehalf,
          before: 400, after: 400,
        }));
        seenContent = true;
        break;

      case 'h1':
        if (seenContent && children.length > 2) {
          children.push(new Paragraph({
            ...pStyle(), pageBreakBefore: true,
            children: [new TextRun({ text: '' })],
          }));
        }
        children.push(makePara(block.content, {
          font: '宋体', size: HALF_PT.h1, bold: true,
          before: 240, after: 120,
        }));
        seenContent = true;
        break;

      case 'h2':
        children.push(makePara(block.content, {
          font: '宋体', size: HALF_PT.h2, bold: true,
          before: 200, after: 80,
        }));
        seenContent = true;
        break;

      case 'h3':
        children.push(makePara(block.content, {
          font: '宋体', size: HALF_PT.h3, bold: true,
          before: 160, after: 80,
        }));
        seenContent = true;
        break;

      case 'paragraph':
        children.push(makePara(block.content, {
          font: '宋体', size: HALF_PT.body, lineSpacing: LINE.onehalf,
          firstLine: INDENT_FIRST,
        }));
        seenContent = true;
        break;

      case 'code':
        for (const line of block.content.split('\n')) {
          children.push(makePara(line || ' ', {
            font: 'Consolas', size: HALF_PT.code, lineSpacing: LINE.single,
          }));
        }
        seenContent = true;
        break;

      case 'equation':
        children.push(makePara(block.content, {
          font: 'Times New Roman', size: HALF_PT.eq,
          alignment: AlignmentType.CENTER, lineSpacing: LINE.onehalf,
        }));
        seenContent = true;
        break;

      case 'figureCaption':
      case 'tableCaption':
        children.push(makePara(block.content, {
          font: '宋体', size: HALF_PT.caption,
          alignment: AlignmentType.CENTER, lineSpacing: LINE.single,
        }));
        seenContent = true;
        break;

      case 'table': {
        const rows = block.content;
        if (rows.length < 2) break;
        const maxCols = Math.max(...rows.map(r => r.length));
        const pad = (r) => [...r, ...Array(maxCols - r.length).fill('')].slice(0, maxCols);
        children.push(new Table({
          rows: rows.map(rowCells =>
            new TableRow({
              children: pad(rowCells).map(cell =>
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
          ),
          width: { size: 100, type: WidthType.PERCENTAGE },
        }));
        seenContent = true;
        break;
      }

      case 'reference':
        children.push(new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { line: LINE.single, lineRule: LineRuleType.AUTO },
          indent: { left: INDENT_HANGING, firstLine: -INDENT_HANGING },
          children: [run(block.content, { font: '宋体', size: HALF_PT.ref })],
        }));
        seenContent = true;
        break;

      case 'image':
        children.push(makePara(`[图片：${block.alt}]`, {
          font: '楷体', size: HALF_PT.caption,
          alignment: AlignmentType.CENTER, lineSpacing: LINE.single,
        }));
        seenContent = true;
        break;

      case 'keywords': {
        const isCn = block.label.startsWith('关键词');
        const p = new Paragraph({
          spacing: { line: LINE.single, lineRule: LineRuleType.AUTO },
        });
        p.addChildElement(run(block.label, { font: isCn ? '宋体' : 'Times New Roman', size: HALF_PT.body, bold: true }));
        if (block.content) {
          p.addChildElement(run(block.content, { font: isCn ? '宋体' : 'Times New Roman', size: HALF_PT.body }));
        }
        children.push(p);
        seenContent = true;
        break;
      }
    }
  }

  return new Document({
    sections: [{
      properties: {
        page: { margin: { top: 907, bottom: 907, left: 1140, right: 1140 } },
      },
      children,
    }],
  });
}

/* ── Download ──────────────────────────── */

async function generateAndDownload(markdown, filename = 'thesis.docx') {
  const doc = buildDocx(markdown);
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ── Word Count ────────────────────────── */

function countChineseWords(text) {
  const chinese = (text.match(/[一-鿿]/g) || []).length;
  const english = (text.match(/[a-zA-Z]+/g) || []).length;
  return { chinese, english, total: chinese + english };
}

/* ── Status Helpers ────────────────────── */

const STATUS_MESSAGES = {
  generated: '✅ DOCX 已下载 — 用 Word/WPS 打开即可看到格式',
  copied: '✅ 内容已复制到剪贴板',
  cleared: '🗑 编辑器已清空',
  fileLoaded: (name) => `📂 已加载：${name}`,
  templateLoaded: (name) => `📋 已加载「${name}」模板 — 修改内容后点"生成 DOCX"`,
  error: (msg) => `❌ ${msg}`,
};

/* ── UI Initialization ─────────────────── */

document.addEventListener('DOMContentLoaded', () => {
  const editor = document.getElementById('editor');
  const preview = document.getElementById('preview');
  const generateBtn = document.getElementById('generateBtn');
  const copyBtn = document.getElementById('copyBtn');
  const clearBtn = document.getElementById('clearBtn');
  const templateSelect = document.getElementById('templateSelect');
  const statusEl = document.getElementById('status');
  const wordCountEl = document.getElementById('wordCount');
  const dropOverlay = document.getElementById('dropOverlay');
  const fileInput = document.getElementById('fileInput');
  const onboarding = document.getElementById('onboarding');
  const onClose = document.getElementById('onClose');
  const previewEmpty = preview?.querySelector('.preview-empty');

  if (!editor || !preview) return;

  /* ── Onboarding ────────────────── */
  if (onboarding && onClose) {
    const dismissed = localStorage.getItem('onboarding-dismissed');
    if (dismissed) onboarding.style.display = 'none';
    onClose.addEventListener('click', () => {
      onboarding.style.display = 'none';
      localStorage.setItem('onboarding-dismissed', '1');
    });
  }

  /* ── Live Preview ──────────────── */
  function updatePreview() {
    const md = editor.value;
    if (md.trim() && typeof marked !== 'undefined') {
      preview.innerHTML = marked.parse(md, { breaks: false });
    } else if (!md.trim() && previewEmpty) {
      preview.innerHTML = '';
      preview.appendChild(previewEmpty);
    } else {
      preview.textContent = md;
    }
  }

  function updateWordCount() {
    const { total } = countChineseWords(editor.value);
    if (wordCountEl) wordCountEl.textContent = `${total.toLocaleString()} 字`;
  }

  function setStatus(msg, type = '') {
    if (statusEl) {
      statusEl.textContent = msg;
      statusEl.className = 'status ' + type;
    }
  }

  editor.addEventListener('input', () => {
    updatePreview();
    updateWordCount();
    setStatus('');
  });

  /* ── Generate DOCX ─────────────── */
  if (generateBtn) {
    generateBtn.addEventListener('click', async () => {
      if (!editor.value.trim()) {
        setStatus('⚠️ 请先输入 Markdown 内容，或拖入 .md 文件，或加载模板', 'warn');
        return;
      }
      generateBtn.disabled = true;
      generateBtn.textContent = '生成中…';
      setStatus('');
      try {
        await generateAndDownload(editor.value);
        setStatus(STATUS_MESSAGES.generated, 'success');
      } catch (err) {
        console.error(err);
        setStatus(STATUS_MESSAGES.error(err.message), 'error');
      }
      generateBtn.disabled = false;
      generateBtn.textContent = '生成 DOCX';
    });
  }

  /* ── Copy ───────────────────────── */
  if (copyBtn) {
    copyBtn.addEventListener('click', async () => {
      if (!editor.value.trim()) return;
      await navigator.clipboard.writeText(editor.value);
      setStatus(STATUS_MESSAGES.copied, 'success');
    });
  }

  /* ── Clear ──────────────────────── */
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      if (editor.value.trim() && !confirm('确定清空编辑器中的所有内容吗？')) return;
      editor.value = '';
      updatePreview();
      updateWordCount();
      if (templateSelect) templateSelect.value = '';
      setStatus(STATUS_MESSAGES.cleared, '');
    });
  }

  /* ── Template ───────────────────── */
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
      updateWordCount();
      const name = templateSelect.options[templateSelect.selectedIndex].text;
      setStatus(STATUS_MESSAGES.templateLoaded(name), 'success');
    });
  }

  /* ── Keyboard Shortcut ──────────── */
  editor.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      generateBtn?.click();
    }
  });

  /* ── Drag & Drop ────────────────── */
  let dragCounter = 0;

  function showDropOverlay() {
    if (dropOverlay) dropOverlay.classList.add('active');
  }

  function hideDropOverlay() {
    if (dropOverlay) dropOverlay.classList.remove('active');
    dragCounter = 0;
  }

  document.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragCounter++;
    showDropOverlay();
  });

  document.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dragCounter--;
    if (dragCounter <= 0) hideDropOverlay();
  });

  document.addEventListener('dragover', (e) => {
    e.preventDefault();
  });

  document.addEventListener('drop', (e) => {
    e.preventDefault();
    hideDropOverlay();
    const files = e.dataTransfer?.files;
    if (files && files.length > 0) {
      loadFile(files[0]);
    }
  });

  /* ── File Input ─────────────────── */
  if (fileInput) {
    fileInput.addEventListener('change', () => {
      const file = fileInput.files?.[0];
      if (file) loadFile(file);
      fileInput.value = '';
    });
  }

  function loadFile(file) {
    const validExts = ['.md', '.txt', '.markdown'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!validExts.includes(ext)) {
      setStatus(STATUS_MESSAGES.error(`不支持的文件类型（${ext}），请使用 .md / .txt / .markdown`), 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      const content = e.target?.result;
      if (typeof content !== 'string') return;
      if (editor.value.trim() && !confirm(`加载「${file.name}」将覆盖当前内容，确定吗？`)) return;
      editor.value = content;
      updatePreview();
      updateWordCount();
      if (templateSelect) templateSelect.value = '';
      setStatus(STATUS_MESSAGES.fileLoaded(file.name), 'success');
    };
    reader.onerror = () => {
      setStatus(STATUS_MESSAGES.error('文件读取失败，请重试'), 'error');
    };
    reader.readAsText(file);
  }

  /* ── Resize Handle ──────────────── */
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
      const pct = Math.max(20, Math.min(80, ((e.clientX - rect.left) / rect.width) * 100));
      leftPanel.style.flex = `${pct}`;
      rightPanel.style.flex = `${100 - pct}`;
    });
    document.addEventListener('mouseup', () => {
      if (dragging) {
        dragging = false;
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
      }
    });
  }

  /* ── Init ────────────────────────── */
  updatePreview();
  updateWordCount();
});

export { generateAndDownload, buildDocx, parseMarkdown };
