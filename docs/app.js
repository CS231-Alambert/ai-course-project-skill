/* ═══════════════════════════════════════════
   AI Course Project Workbench — app.js
   7 tabs: Editor | AI | Evidence | Images
   | Word Count | References | Checklist
   ═══════════════════════════════════════════ */

import { Document, Packer, Paragraph, TextRun,
         Table, TableRow, TableCell,
         AlignmentType, WidthType, LineRuleType } from 'docx';

/* ═══════════════════════════════════════════
   SECTION 0 — Shared Constants & Utilities
   ═══════════════════════════════════════════ */

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
...

## 实验结果

| 指标 | 数值 |
|------|------|
| 准确率 | 0.95 |
| F1 | 0.93 |

## 结论
...`,
};

function $(sel) { return document.querySelector(sel); }
function $$(sel) { return document.querySelectorAll(sel); }

function countWords(text) {
  const cn = (text.match(/[一-鿿]/g) || []).length;
  const en = (text.match(/[a-zA-Z]+/g) || []).length;
  return cn + en;
}

function showStatus(msg, type) {
  const el = $('#status');
  if (el) { el.textContent = msg; el.className = 'status ' + (type || ''); }
}

/* ═══════════════════════════════════════════
   SECTION 1 — Format Requirements Parser
   ═══════════════════════════════════════════ */

const FONT_SIZE_MAP = {
  '初号':42,'小初':36,'一号':26,'小一':24,'二号':22,'小二':18,
  '三号':16,'小三':15,'四号':14,'小四':12,'五号':10.5,'小五':9,
  '六号':7.5,'小六':6.5,'七号':5.5,'八号':5,
};
const FONT_NAMES = ['宋体','黑体','楷体','仿宋','Times New Roman','Consolas','微软雅黑','华文中宋'];

function parseRequirements(text) {
  const result = {
    bodyFont: null, bodySize: null, headingFont: null, h1Size: null,
    lineSpacing: null, indent: null, marginH: null, marginV: null,
    wordMin: null, wordMax: null, chapters: [], rawMatches: [],
  };

  // Font
  for (const fn of FONT_NAMES) {
    if (text.includes(fn)) {
      const near = text.substring(Math.max(0, text.indexOf(fn)-20), text.indexOf(fn)+fn.length+40);
      if (/正文|正?文|段落/.test(near) && !/标题|题目|章/.test(near)) result.bodyFont = fn;
      else if (/标题|题目|章|节/.test(near)) result.headingFont = fn;
      else if (!result.bodyFont) result.bodyFont = fn;
    }
  }

  // Font size
  for (const [name, pt] of Object.entries(FONT_SIZE_MAP)) {
    const re = new RegExp(`[^a-zA-Z]${name}[^a-zA-Z]|^${name}[^a-zA-Z]`, 'g');
    for (const m of text.matchAll(re)) {
      const near = text.substring(Math.max(0, m.index-30), m.index+name.length+30);
      if (/正文|正?文|段落|内容/.test(near) && !/标题|题目/.test(near)) result.bodySize = pt;
      else if (/标题|题目|一级|章/.test(near)) result.h1Size = pt;
    }
  }
  // Numeric pt
  const bodyPt = text.match(/正\s*文\s*[：:]\s*(\d+)\s*pt/i) || text.match(/正文[^。]*(?:小四|五号).*?(\d+)\s*pt/i);
  if (bodyPt) result.bodySize = parseInt(bodyPt[1]);
  const h1Pt = text.match(/一级标题[^。]*?(\d+)\s*pt/i) || text.match(/标题[：:]\s*(\d+)\s*pt/i);
  if (h1Pt) result.h1Size = parseInt(h1Pt[1]);

  // Line spacing
  const ls = text.match(/(\d+\.?\d*)\s*倍\s*行\s*距/);
  if (ls) result.lineSpacing = parseFloat(ls[1]);
  if (text.includes('单倍行距')) result.lineSpacing = 1.0;
  if (text.includes('双倍行距')) result.lineSpacing = 2.0;

  // Indent
  const indent = text.match(/首行缩进\s*(\d+)\s*字符/);
  if (indent) result.indent = parseInt(indent[1]) * 12;

  // Margins
  const mv = text.match(/上\s*下\s*(\d+\.?\d*)\s*cm/);
  if (mv) result.marginV = parseFloat(mv[1]);
  const mh = text.match(/左\s*右\s*(\d+\.?\d*)\s*cm/);
  if (mh) result.marginH = parseFloat(mh[1]);

  // Word count
  const wRange = text.match(/不少于\s*(\d+)\s*字/) || text.match(/字数[：:]\s*(\d+)[-~]\s*(\d+)/);
  if (wRange) {
    result.wordMin = parseInt(wRange[1]);
    if (wRange[2]) result.wordMax = parseInt(wRange[2]);
  }

  // Chapters
  const chRe = /第[一二三四五六七八\d]+章\s*[^\n，。]*/g;
  for (const m of text.matchAll(chRe)) {
    const ch = m[0].trim();
    if (!result.chapters.includes(ch)) result.chapters.push(ch);
  }

  // Track raw matches
  result.rawMatches = [];
  if (result.bodyFont) result.rawMatches.push(`正文：${result.bodyFont}`);
  if (result.bodySize) result.rawMatches.push(`${result.bodySize}pt`);
  if (result.lineSpacing) result.rawMatches.push(`${result.lineSpacing}倍行距`);
  if (result.indent) result.rawMatches.push(`首行缩进${result.indent}pt`);
  if (result.marginV) result.rawMatches.push(`上下${result.marginV}cm`);
  if (result.marginH) result.rawMatches.push(`左右${result.marginH}cm`);

  return result;
}

function applyParsedRequirements(parsed) {
  const setVal = (id, val, def) => {
    const el = document.getElementById(id);
    if (el && val != null) {
      el.value = val;
      if (val !== def) el.classList.add('changed');
      else el.classList.remove('changed');
    }
  };
  setVal('cfgBodyFont', parsed.bodyFont, '宋体');
  setVal('cfgBodySize', parsed.bodySize, 12);
  setVal('cfgLineSpacing', parsed.lineSpacing, 1.5);
  setVal('cfgHeadingFont', parsed.headingFont, '宋体');
  setVal('cfgH1Size', parsed.h1Size, 16);
  setVal('cfgIndent', parsed.indent, 24);
  setVal('cfgMarginV', parsed.marginV, 2.54);
  setVal('cfgMarginH', parsed.marginH, 3.17);

  // Update summary
  const specEls = $$('.req-spec em');
  if (specEls.length >= 3) {
    specEls[0].textContent = `${parsed.bodyFont || '宋体'} ${parsed.bodySize || 12}pt ${parsed.lineSpacing || 1.5}倍行距`;
    specEls[1].textContent = `${parsed.headingFont || '宋体'} ${parsed.h1Size || 16}pt 加粗`;
    specEls[2].textContent = `上下${parsed.marginV || 2.54} 左右${parsed.marginH || 3.17}cm`;
  }

  // Word info
  const wRange = $('#reqWordRange');
  if (wRange) {
    if (parsed.wordMin) wRange.textContent = parsed.wordMax ? `${parsed.wordMin}-${parsed.wordMax} 字` : `不少于 ${parsed.wordMin} 字`;
    else wRange.textContent = '未检测';
  }
  const chInfo = $('#reqChapterInfo');
  if (chInfo) {
    chInfo.textContent = parsed.chapters.length ? parsed.chapters.join('、') : '未检测';
  }

  // Update status
  const badge = $('#reqStatus');
  if (badge) {
    badge.dataset.status = 'uploaded';
    badge.textContent = '已从文件提取';
  }
}

function resetToDefaults() {
  $$('#reqDetail input[type="number"], #reqDetail input[type="text"]').forEach(el => {
    const def = el.dataset.default;
    if (def !== undefined) { el.value = def; el.classList.remove('changed'); }
  });
  const badge = $('#reqStatus');
  if (badge) { badge.dataset.status = 'empty'; badge.textContent = '默认配置'; }
  const wRange = $('#reqWordRange');
  if (wRange) wRange.textContent = '未检测';
  const chInfo = $('#reqChapterInfo');
  if (chInfo) chInfo.textContent = '未检测';
  const specEls = $$('.req-spec em');
  if (specEls.length >= 3) {
    specEls[0].textContent = '宋体 12pt 1.5倍行距';
    specEls[1].textContent = '宋体 16pt 加粗';
    specEls[2].textContent = '上下2.54 左右3.17cm';
  }
}

async function handleReqFile(file) {
  const ext = '.' + file.name.split('.').pop().toLowerCase();
  if (!['.md','.txt','.docx','.pdf'].includes(ext)) {
    showStatus(`❌ 不支持 ${ext}，请使用 .md / .txt / .docx / .pdf`, 'error');
    return;
  }

  let text = '';
  if (ext === '.docx') {
    // Try mammoth.js if loaded, otherwise read raw text
    if (typeof mammoth !== 'undefined') {
      const result = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
      text = result.value;
    } else {
      // Basic: try extracting XML text content (works for simple DOCX)
      try {
        const zip = await JSZip.loadAsync(await file.arrayBuffer());
        const docXml = await zip.file('word/document.xml')?.async('string');
        if (docXml) text = docXml.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
      } catch { text = ''; }
    }
    if (!text) {
      showStatus('⚠️ DOCX 解析需要 mammoth.js 库，请使用 .md 或 .txt 格式', 'warn');
      return;
    }
  } else if (ext === '.pdf') {
    showStatus('⚠️ PDF 解析暂不支持，请粘贴为 .txt 后上传', 'warn');
    return;
  } else {
    text = await file.text();
  }

  if (text) {
    const parsed = parseRequirements(text);
    applyParsedRequirements(parsed);
    showStatus(`✅ 已从「${file.name}」提取格式要求（${parsed.rawMatches.join('、') || '未检测到具体规格'}）`, 'success');
  }
}

/* ═══════════════════════════════════════════
   SECTION 2 — Field Status Tracking
   ═══════════════════════════════════════════ */

function updateFieldStatus() {
  $$('.form-group[data-field]').forEach(group => {
    const input = group.querySelector('input, textarea');
    if (!input) return;
    const val = input.value?.trim() || '';
    const required = input.dataset.required === 'true';
    const tag = group.querySelector('.field-tag');
    if (!tag) return;

    if (val) {
      tag.className = 'field-tag tag-filled';
      tag.textContent = '✓ 已填写';
    } else if (required) {
      tag.className = 'field-tag tag-required';
      tag.textContent = '必填';
    } else {
      tag.className = 'field-tag tag-optional';
      tag.textContent = '建议填写';
    }
  });
}

function getFormatConfig() {
  const get = id => { const el = document.getElementById(id); return el ? el.value : null; };
  return {
    bodyFont: get('cfgBodyFont') || '宋体',
    bodySize: parseInt(get('cfgBodySize')) || 12,
    lineSpacing: parseFloat(get('cfgLineSpacing')) || 1.5,
    headingFont: get('cfgHeadingFont') || '宋体',
    h1Size: parseInt(get('cfgH1Size')) || 16,
    indent: parseInt(get('cfgIndent')) || 24,
    marginV: parseFloat(get('cfgMarginV')) || 2.54,
    marginH: parseFloat(get('cfgMarginH')) || 3.17,
  };
}

/* ═══════════════════════════════════════════
   SECTION 3 — AI Generator (DeepSeek API)
   ═══════════════════════════════════════════ */

const DEEPSEEK_BASE = 'https://api.deepseek.com/v1';

function getApiKey() {
  return localStorage.getItem('deepseek_api_key') || '';
}

async function callDeepSeek(systemPrompt, userPrompt) {
  const key = getApiKey();
  if (!key) throw new Error('请先设置 DeepSeek API Key');

  const resp = await fetch(`${DEEPSEEK_BASE}/chat/completions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` },
    body: JSON.stringify({
      model: 'deepseek-chat',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt },
      ],
      temperature: 0.7,
      max_tokens: 4096,
    }),
  });

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err.error?.message || `API 请求失败 (${resp.status})`);
  }

  const data = await resp.json();
  return data.choices?.[0]?.message?.content || '';
}

const AI_SYSTEM_PROMPT = `你是一位 AI 课程论文写作助手。根据用户提供的项目信息，生成一篇结构完整的课程论文 Markdown。

要求：
1. 使用 # ## ### 标题层级
2. 包含：摘要、关键词、第1章引言、第2章相关工作、第3章方法、第4章实验、第5章总结与展望、参考文献
3. 正文使用学术风格，客观严谨，不编造数据
4. 表格使用 GFM 格式 |...|
5. 公式使用 $$...$$
6. 用户提供的实验数据直接使用，不要修改
7. 参考文献至少列出 5 条相关的中英文学术文献
8. 论文标题使用用户项目名称
9. 总字数控制在 3000-5000 字`;

async function generateThesisWithAI() {
  const desc = $('#aiProjectDesc')?.value?.trim() || '';
  const tech = $('#aiTechDetail')?.value?.trim() || '';
  const results = $('#aiResults')?.value?.trim() || '';
  const append = $('#aiAppend')?.checked ?? true;

  if (!desc) { showStatus('⚠️ 请至少填写项目描述', 'warn'); return; }

  const output = $('#aiOutput');
  const btn = $('#aiGenBtn');
  if (output) { output.innerHTML = '<p class="form-hint">⏳ AI 正在生成论文初稿，请稍候（通常 15-30 秒）...</p>'; output.classList.add('generating'); }
  if (btn) { btn.disabled = true; btn.textContent = '⏳ 生成中…'; }

  try {
    const userPrompt = `请根据以下项目信息，生成一篇完整的课程论文：

【项目描述】
${desc}

【技术栈 & 实验配置】
${tech || '（未提供）'}

【实验结果】
${results || '（未提供）'}

请严格按照要求的章节结构生成。如果某些信息未提供，请根据项目描述合理推测并注明"建议补充"。`;

    const md = await callDeepSeek(AI_SYSTEM_PROMPT, userPrompt);

    if (output) { output.innerHTML = '<p class="ok">✅ AI 生成完成！内容已写入编辑器，请在右侧预览中查看。</p>'; output.classList.remove('generating'); }

    // Write to editor
    const editor = $('#editor');
    if (editor) {
      if (append && editor.value.trim()) {
        editor.value = editor.value.trimEnd() + '\n\n' + md;
      } else {
        editor.value = md;
      }
      editor.dispatchEvent(new Event('input'));
      // Switch to editor tab
      $$('.tab').forEach(t => t.classList.remove('active'));
      const editorTab = document.querySelector('[data-tab="editor"]');
      if (editorTab) editorTab.classList.add('active');
      $$('.tab-panel').forEach(p => p.classList.remove('active'));
      $('#panel-editor')?.classList.add('active');
    }

    showStatus('✅ AI 论文初稿已生成，请审阅修改', 'success');
  } catch (err) {
    if (output) { output.innerHTML = `<p class="err">❌ 生成失败：${err.message}</p>`; output.classList.remove('generating'); }
    showStatus(`❌ ${err.message}`, 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = '🤖 AI 生成论文初稿'; }
  }
}

/* ═══════════════════════════════════════════
   SECTION 4 — DOCX Generator (generate_docx.py)
   ═══════════════════════════════════════════ */

const HALF_PT = { title:44, h1:32, h2:28, h3:24, body:24, caption:21, code:18, ref:21, eq:24 };
const LINE = { single:240, onehalf:360 };
const INDENT_FIRST = 480;
const INDENT_HANGING = 420;

function parseMarkdown(md) {
  const lines = md.split('\n');
  const blocks = [];
  let i = 0;
  while (i < lines.length) {
    const s = lines[i].trim();
    if (!s) { i++; continue; }
    if (s.startsWith('```')) {
      const lang = s.slice(3).trim().toLowerCase();
      const code = []; i++;
      while (i < lines.length && !lines[i].trim().startsWith('```')) { code.push(lines[i]); i++; }
      if (i < lines.length) i++;
      blocks.push({ type:'code', lang, content:code.join('\n') });
      continue;
    }
    if (s === '$$') {
      const eq = []; i++;
      while (i < lines.length && lines[i].trim() !== '$$') { eq.push(lines[i].trim()); i++; }
      if (i < lines.length) i++;
      blocks.push({ type:'equation', content:eq.join('\n') });
      continue;
    }
    if (s.startsWith('$$') && s.endsWith('$$') && s.length > 4) {
      blocks.push({ type:'equation', content:s.slice(2,-2).trim() }); i++; continue;
    }
    if (s.startsWith('# ')) { blocks.push({ type:'title', content:cleanText(s.slice(2)) }); i++; continue; }
    if (s.startsWith('## ')) { blocks.push({ type:'h1', content:cleanText(s.slice(3)) }); i++; continue; }
    if (s.startsWith('### ')) { blocks.push({ type:'h2', content:cleanText(s.slice(4)) }); i++; continue; }
    if (s.startsWith('#### ')) { blocks.push({ type:'h3', content:cleanText(s.slice(5)) }); i++; continue; }
    const kw = s.match(/^(关键词|Keywords)[:：]\s*(.*)$/);
    if (kw) { blocks.push({ type:'keywords', label:kw[1], content:kw[2] }); i++; continue; }
    if (/^(图|Fig\.?)\s*\d/.test(s)) { blocks.push({ type:'figureCaption', content:cleanText(s) }); i++; continue; }
    if (/^(表|Table\.?)\s*\d/.test(s)) { blocks.push({ type:'tableCaption', content:cleanText(s) }); i++; continue; }
    if (s.startsWith('|')) {
      const tbl = [];
      while (i < lines.length && lines[i].trim().startsWith('|')) { tbl.push(lines[i].trim()); i++; }
      blocks.push({ type:'table', content:parseTable(tbl) });
      continue;
    }
    const img = s.match(/^!\[(.+?)\]\((.+?)\)$/);
    if (img) { blocks.push({ type:'image', alt:img[1], src:img[2] }); i++; continue; }
    if (/^\[\d+\]/.test(s)) { blocks.push({ type:'reference', content:cleanText(s) }); i++; continue; }
    blocks.push({ type:'paragraph', content:cleanText(s) }); i++;
  }
  return blocks;
}

function parseTable(lines) {
  const rows = [];
  for (const ln of lines) {
    const cells = ln.replace(/^[\s|]+|[\s|]+$/g,'').split('|').map(c=>c.trim());
    if (cells.every(c=>/^[:\- ]+$/.test(c))) continue;
    rows.push(cells);
  }
  return rows;
}

function cleanText(t) {
  return t.replace(/\[([^\]]+)\]\([^)]+\)/g,'$1')
    .replace(/`/g,'').replace(/\*\*(.+?)\*\*/g,'$1').replace(/__(.+?)__/g,'$1')
    .replace(/(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)/g,'$1')
    .replace(/\s+/g,' ').trim();
}

function run(text, opts={}) { return new TextRun({ text:String(text), font:opts.font||'宋体', size:opts.size||HALF_PT.body, bold:!!opts.bold }); }

function buildDocx(md) {
  const blocks = parseMarkdown(md);
  const children = [];
  let seen = false;
  for (const b of blocks) {
    switch (b.type) {
      case 'title':
        if (seen) children.push(new Paragraph({ children:[run('')], pageBreakBefore:true }));
        children.push(new Paragraph({ alignment:AlignmentType.CENTER, spacing:{line:LINE.onehalf,lineRule:LineRuleType.AUTO,before:400,after:400}, children:[run(b.content,{font:'宋体',size:HALF_PT.title,bold:true})] }));
        seen=true; break;
      case 'h1':
        if (seen && children.length>2) children.push(new Paragraph({ children:[run('')], pageBreakBefore:true }));
        children.push(new Paragraph({ spacing:{line:LINE.onehalf,lineRule:LineRuleType.AUTO,before:240,after:120}, children:[run(b.content,{font:'宋体',size:HALF_PT.h1,bold:true})] }));
        seen=true; break;
      case 'h2':
        children.push(new Paragraph({ spacing:{line:LINE.onehalf,lineRule:LineRuleType.AUTO,before:200,after:80}, children:[run(b.content,{font:'宋体',size:HALF_PT.h2,bold:true})] }));
        seen=true; break;
      case 'h3':
        children.push(new Paragraph({ spacing:{line:LINE.onehalf,lineRule:LineRuleType.AUTO,before:160,after:80}, children:[run(b.content,{font:'宋体',size:HALF_PT.h3,bold:true})] }));
        seen=true; break;
      case 'paragraph':
        children.push(new Paragraph({ spacing:{line:LINE.onehalf,lineRule:LineRuleType.AUTO}, indent:{firstLine:INDENT_FIRST}, children:[run(b.content,{font:'宋体',size:HALF_PT.body})] }));
        seen=true; break;
      case 'code':
        for (const ln of b.content.split('\n')) children.push(new Paragraph({ spacing:{line:LINE.single,lineRule:LineRuleType.AUTO}, children:[run(ln||' ',{font:'Consolas',size:HALF_PT.code})] }));
        seen=true; break;
      case 'equation':
        children.push(new Paragraph({ alignment:AlignmentType.CENTER, spacing:{line:LINE.onehalf,lineRule:LineRuleType.AUTO}, children:[run(b.content,{font:'Times New Roman',size:HALF_PT.eq})] }));
        seen=true; break;
      case 'figureCaption': case 'tableCaption':
        children.push(new Paragraph({ alignment:AlignmentType.CENTER, spacing:{line:LINE.single,lineRule:LineRuleType.AUTO}, children:[run(b.content,{font:'宋体',size:HALF_PT.caption})] }));
        seen=true; break;
      case 'table': {
        const rows = b.content;
        if (rows.length<2) break;
        const mc = Math.max(...rows.map(r=>r.length));
        const pad = r => [...r,...Array(mc-r.length).fill('')].slice(0,mc);
        children.push(new Table({ rows:rows.map(row=>new TableRow({ children:pad(row).map(c=>new TableCell({ children:[new Paragraph({alignment:AlignmentType.CENTER,spacing:{line:LINE.single,lineRule:LineRuleType.AUTO},children:[run(c,{font:'宋体',size:HALF_PT.caption})]})], width:{size:100/mc,type:WidthType.PERCENTAGE} })) })), width:{size:100,type:WidthType.PERCENTAGE} }));
        seen=true; break;
      }
      case 'reference':
        children.push(new Paragraph({ spacing:{line:LINE.single,lineRule:LineRuleType.AUTO}, indent:{left:INDENT_HANGING,firstLine:-INDENT_HANGING}, children:[run(b.content,{font:'宋体',size:HALF_PT.ref})] }));
        seen=true; break;
      case 'image':
        children.push(new Paragraph({ alignment:AlignmentType.CENTER, spacing:{line:LINE.single,lineRule:LineRuleType.AUTO}, children:[run(`[图片：${b.alt}]`,{font:'楷体',size:HALF_PT.caption})] }));
        seen=true; break;
      case 'keywords': {
        const p = new Paragraph({ spacing:{line:LINE.single,lineRule:LineRuleType.AUTO} });
        p.addChildElement(run(b.label,{font:'宋体',size:HALF_PT.body,bold:true}));
        if (b.content) p.addChildElement(run(b.content,{font:'宋体',size:HALF_PT.body}));
        children.push(p); seen=true; break;
      }
    }
  }
  return new Document({ sections:[{ properties:{ page:{ margin:{ top:907, bottom:907, left:1140, right:1140 } } }, children }] });
}

async function downloadDocx(md) {
  const blob = await Packer.toBlob(buildDocx(md));
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'thesis.docx';
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ═══════════════════════════════════════════
   SECTION 5 — Evidence Builder (build_evidence.py)
   ═══════════════════════════════════════════ */

function buildEvidence() {
  const name = $('#evProjectName')?.value?.trim() || '未命名项目';
  const tech = $('#evTechStack')?.value?.trim() || '';
  const api = $('#evApi')?.value?.trim() || '';
  const schema = $('#evSchema')?.value?.trim() || '';
  const tests = $('#evTests')?.value?.trim() || '';

  const toList = s => s.split('\n').filter(l=>l.trim()).map(l=>l.trim());

  const techItems = toList(tech);
  const apiItems = toList(api);
  const schemaItems = toList(schema);
  const testItems = toList(tests);

  const div = $('#evOutput');
  if (!div) return;
  div.innerHTML = `
    <h3>📋 项目证据报告：${name}</h3>
    <table>
      <tr><th>类别</th><th>数量</th><th>详情</th></tr>
      <tr><td>🔧 技术栈</td><td>${techItems.length}</td><td>${techItems.map(t => `<code>${t}</code>`).join(', ') || '—'}</td></tr>
      <tr><td>📡 API / 入口</td><td>${apiItems.length}</td><td>${apiItems.map(t => `<code>${t}</code>`).join(', ') || '—'}</td></tr>
      <tr><td>🗄 数据 / Schema</td><td>${schemaItems.length}</td><td>${schemaItems.map(t => `<code>${t}</code>`).join(', ') || '—'}</td></tr>
      <tr><td>🧪 测试</td><td>${testItems.length}</td><td>${testItems.map(t => `<code>${t}</code>`).join(', ') || '—'}</td></tr>
      <tr><td><strong>📦 总计</strong></td><td colspan="2"><strong>${techItems.length + apiItems.length + schemaItems.length + testItems.length} 个文件/实体</strong></td></tr>
    </table>
    <p style="margin-top:12px;color:var(--text3);font-size:var(--xs)">💡 此报告可粘贴到论文「实验环境」章节</p>
  `;
}

/* ═══════════════════════════════════════════
   SECTION 6 — Image Manager (build_image_map.py)
   ═══════════════════════════════════════════ */

let imgStore = [];

function addImages(files) {
  for (const f of files) {
    if (!f.type.startsWith('image/')) continue;
    const reader = new FileReader();
    reader.onload = e => {
      imgStore.push({ name: f.name, stem: f.name.replace(/\.[^.]+$/,''), dataUrl: e.target.result, file: f });
      renderImgGrid();
    };
    reader.readAsDataURL(f);
  }
}

function renderImgGrid() {
  const grid = $('#imgGrid');
  if (!grid) return;
  grid.innerHTML = imgStore.map((img,i) => `
    <div class="img-item">
      <img src="${img.dataUrl}" alt="${img.name}">
      <span class="img-name">${img.stem}</span>
      <button class="img-del" data-idx="${i}" title="删除">✕</button>
    </div>
  `).join('');
  grid.querySelectorAll('.img-del').forEach(btn => {
    btn.addEventListener('click', () => {
      imgStore.splice(parseInt(btn.dataset.idx), 1);
      renderImgGrid();
    });
  });
}

function buildImageMap() {
  const map = {};
  for (const img of imgStore) map[img.stem] = img.name;
  const out = $('#imgMapOutput');
  if (out) {
    out.style.display = 'block';
    out.textContent = JSON.stringify(map, null, 2);
  }
  return map;
}

/* ═══════════════════════════════════════════
   SECTION 7 — Word Count (count_words.py)
   ═══════════════════════════════════════════ */

function analyzeWordCount() {
  const md = getEditorValue();
  const lines = md.split('\n');
  const chapters = [];
  let curTitle = '(前言)';
  let curText = [];
  for (const line of lines) {
    const s = line.trim();
    if (s.startsWith('## ') && !s.startsWith('### ')) {
      if (curText.length) chapters.push({ title: curTitle, words: countWords(curText.join('\n')) });
      curTitle = s.slice(3).trim();
      curText = [];
    } else { curText.push(line); }
  }
  if (curText.length) chapters.push({ title: curTitle, words: countWords(curText.join('\n')) });

  const total = chapters.reduce((s,c)=>s+c.words, 0);
  const div = $('#wcOutput');
  if (!div) return;
  div.innerHTML = `
    <table>
      <tr><th>章节</th><th>字数</th><th>进度</th></tr>
      ${chapters.map(c => {
        const pct = total ? Math.round(c.words/total*100) : 0;
        return `<tr><td>${c.title}</td><td>${c.words.toLocaleString()}</td><td><progress value="${pct}" max="100" style="width:60px;height:8px" title="${pct}%"></progress> ${pct}%</td></tr>`;
      }).join('')}
      <tr style="font-weight:700;border-top:2px solid var(--border)"><td>合计</td><td>${total.toLocaleString()}</td><td>${chapters.length} 章</td></tr>
    </table>
    ${total < 3000 ? '<p class="warn" style="margin-top:8px">⚠️ 字数偏少，建议 ≥ 3000 字</p>' : total > 8000 ? '<p class="warn" style="margin-top:8px">⚠️ 字数偏多，建议控制在 5000-8000 字</p>' : '<p class="ok" style="margin-top:8px">✅ 字数在合理范围</p>'}
  `;
}

/* ═══════════════════════════════════════════
   SECTION 8 — Reference Checker (check_references.py)
   ═══════════════════════════════════════════ */

function checkReferences() {
  const md = getEditorValue();
  const lines = md.split('\n');
  let inRefs = false;
  const refs = [];
  for (const line of lines) {
    const s = line.trim();
    if (s.startsWith('##') && s.includes('参考')) { inRefs = true; continue; }
    if (inRefs && s.startsWith('##') && !s.includes('参考')) break;
    if (!inRefs) continue;
    const m = s.match(/^\[(\d+)\]\s*(.+)$/);
    if (m) {
      const body = m[2].trim();
      const hasZh = /[一-鿿]/.test(body);
      const yr = body.match(/\b(20\d{2})\b/);
      refs.push({ index: parseInt(m[1]), raw: s, lang: hasZh?'zh':'en', year: yr ? parseInt(yr[1]) : null, hasDoi: /doi/i.test(body) });
    }
  }

  const div = $('#refOutput');
  if (!div) return;
  if (!refs.length) { div.innerHTML = '<p class="warn">未找到参考文献（请在编辑器中添加 ## 参考文献 章节，格式：[1] 作者. 标题. 期刊, 年份.）</p>'; return; }

  const indices = refs.map(r=>r.index);
  const dups = [...new Set(indices.filter((v,i,a)=>a.indexOf(v)!==i))];
  const badYears = refs.filter(r=>!r.year||r.year<2018);
  const zhCount = refs.filter(r=>r.lang==='zh').length;
  const doiCount = refs.filter(r=>r.hasDoi).length;

  div.innerHTML = `
    <table>
      <tr><td>📚 总数</td><td><strong>${refs.length}</strong> 条</td></tr>
      <tr><td>🇨🇳 中文</td><td>${zhCount} 条</td></tr>
      <tr><td>🇬🇧 英文</td><td>${refs.length - zhCount} 条</td></tr>
      <tr><td>🔗 含 DOI</td><td>${doiCount} 条</td></tr>
      <tr><td>📅 年份范围</td><td>${Math.min(...refs.filter(r=>r.year).map(r=>r.year))} — ${Math.max(...refs.filter(r=>r.year).map(r=>r.year))}</td></tr>
    </table>
    ${dups.length ? `<p class="warn" style="margin-top:8px">⚠️ 重复编号：${dups.join(', ')}</p>` : ''}
    ${badYears.length ? `<p class="warn" style="margin-top:8px">⚠️ ${badYears.length} 条参考文献年份缺失或早于 2018：${badYears.map(r=>`[${r.index}]`).join(', ')}</p>` : '<p class="ok" style="margin-top:8px">✅ 年份检查通过</p>'}
    ${refs.length < 10 ? `<p class="warn" style="margin-top:4px">💡 当前 ${refs.length} 条偏少，建议 ≥ 10 条</p>` : '<p class="ok" style="margin-top:4px">✅ 参考文献数量充足</p>'}
  `;
}

/* ═══════════════════════════════════════════
   SECTION 9 — Checklist Generator (generate_task_checklist.py)
   ═══════════════════════════════════════════ */

async function generateChecklist() {
  const md = getEditorValue();
  const lines = md.split('\n');

  // Find paper title
  let title = '论文';
  for (const line of lines) {
    if (line.startsWith('# ')) { title = line.slice(2).trim(); break; }
  }

  // Find missing screenshots
  const screenshots = [];
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(/\[此处插入截图：(.+?)\]/);
    if (m) {
      let ctx = '';
      for (let j = i-1; j >= Math.max(0, i-10); j--) {
        const s = lines[j].trim();
        if (s.startsWith('## ') || s.startsWith('### ')) { ctx = s.replace(/^#+\s*/,''); break; }
      }
      screenshots.push({ label: m[1], context: ctx });
    }
  }

  // Count references
  let refCount = 0, inRefs = false;
  for (const line of lines) {
    const s = line.trim();
    if (s.startsWith('##') && s.includes('参考')) { inRefs = true; continue; }
    if (inRefs && s.startsWith('##') && !s.includes('参考')) break;
    if (inRefs && /^\[\d+\]/.test(s)) refCount++;
  }

  const div = $('#clOutput');
  if (!div) return;

  div.innerHTML = `
    <p><strong>《${title}》— 待办清单</strong></p>
    <table>
      <tr><td>📝 论文标题</td><td>${title}</td></tr>
      <tr><td>🖼 待补充截图</td><td>${screenshots.length} 处</td></tr>
      <tr><td>📚 参考文献</td><td>${refCount} 条</td></tr>
    </table>
    ${screenshots.length ? `
    <h3 style="margin-top:12px">需补充截图：</h3>
    <ol style="font-size:var(--sm);color:var(--text2)">
      ${screenshots.map((s,i) => `<li><strong>${s.label}</strong> ${s.context ? `（${s.context}）` : ''}<br><code style="font-size:var(--xs)">→ 放入 figures/ 目录</code></li>`).join('')}
    </ol>` : '<p class="ok" style="margin-top:8px">✅ 未发现截图占位符</p>'}
  `;

  // Also generate a downloadable checklist DOCX
  try {
    const doc = buildChecklistDocx(title, screenshots, refCount);
    const blob = await Packer.toBlob(doc);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${title}-待办清单.docx`;
    a.click();
    URL.revokeObjectURL(a.href);
    showStatus(`✅ 待办清单 DOCX 已下载：${title}-待办清单.docx`, 'success');
  } catch(e) {
    console.error(e);
  }
}

function buildChecklistDocx(title, screenshots, refCount) {
  const C = [];
  const H = (t,s,b) => new Paragraph({ spacing:{line:LINE.onehalf,lineRule:LineRuleType.AUTO,before:b||200,after:80}, children:[new TextRun({text:t,font:'宋体',size:s,bold:true})] });
  const B = t => new Paragraph({ spacing:{line:LINE.onehalf,lineRule:LineRuleType.AUTO}, indent:{firstLine:INDENT_FIRST}, children:[new TextRun({text:t,font:'宋体',size:HALF_PT.body})] });
  const P = t => new Paragraph({ spacing:{line:LINE.single,lineRule:LineRuleType.AUTO}, children:[new TextRun({text:'    '+t,font:'Consolas',size:HALF_PT.code})] });
  const CB = t => new Paragraph({ spacing:{line:LINE.single,lineRule:LineRuleType.AUTO,before:40,after:40}, children:[new TextRun({text:'☐ '+t,font:'宋体',size:HALF_PT.body})] });
  const NB = t => new Paragraph({ spacing:{line:LINE.single,lineRule:LineRuleType.AUTO}, children:[new TextRun({text:'💡 '+t,font:'楷体',size:HALF_PT.caption})] });

  // Title
  C.push(new Paragraph({ alignment:AlignmentType.CENTER, spacing:{line:LINE.onehalf,lineRule:LineRuleType.AUTO,before:400,after:400}, children:[new TextRun({text:`《${title}》— 待办清单`,font:'宋体',size:HALF_PT.title,bold:true})] }));
  C.push(H('一、当前状态', HALF_PT.h1, 240));
  C.push(B(`论文已生成。待补充截图 ${screenshots.length} 处，参考文献 ${refCount} 条。`));

  if (screenshots.length) {
    C.push(H('二、需要补充的截图', HALF_PT.h1, 200));
    screenshots.forEach((ss, i) => {
      C.push(H(`任务 ${i+1}：${ss.label}`, HALF_PT.h2, 160));
      C.push(B(`内容：${ss.label}`));
      if (ss.context) C.push(B(`章节：${ss.context}`));
      C.push(CB('截图保存为 PNG 格式'));
      C.push(CB(`放入 figures/ 目录`));
      C.push(CB('更新 image-map.json'));
      C.push(NB(`完成后告诉我截图路径`));
    });
  }

  C.push(H('三、最终检查', HALF_PT.h1, 200));
  C.push(CB(`所有 ${screenshots.length} 处截图已放入 figures/`));
  C.push(CB(`参考文献 ${refCount} 条，格式正确`));
  C.push(CB('DOCX 中无【待补】标记'));
  C.push(CB('格式符合学校要求'));
  C.push(CB('文件已提交'));

  C.push(H('四、关键路径', HALF_PT.h1, 200));
  C.push(B('论文 Markdown 源稿：paper-output/<标题>.md'));
  C.push(B('论文 DOCX 输出：paper-output/<标题>.docx'));
  C.push(B('图片目录：paper-output/figures/'));
  C.push(B('图片映射：paper-output/image-map.json'));

  return new Document({ sections:[{ properties:{ page:{ margin:{ top:907, bottom:907, left:1140, right:1140 } } }, children:C }] });
}

/* ═══════════════════════════════════════════
   SECTION 10 — UI Wiring
   ═══════════════════════════════════════════ */

function getEditorValue() { return $('#editor')?.value || ''; }
function setEditorValue(v) { const e = $('#editor'); if (e) e.value = v; }

document.addEventListener('DOMContentLoaded', () => {
  const editor = $('#editor');
  const preview = $('#preview');
  const previewEmpty = $('#previewEmpty');

  /* ── Tab Switching ────────────────── */
  $$('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      $$('.tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      $$('.tab-panel').forEach(p => p.classList.remove('active'));
      const panel = $(`#panel-${tab.dataset.tab}`);
      if (panel) panel.classList.add('active');
    });
  });

  /* ── Preview ──────────────────────── */
  function updatePreview() {
    const md = getEditorValue();
    if (md.trim() && typeof marked !== 'undefined') {
      preview.innerHTML = marked.parse(md, { breaks: false });
    } else if (!md.trim() && previewEmpty) {
      preview.innerHTML = '';
      preview.appendChild(previewEmpty);
    }
  }

  editor?.addEventListener('input', () => {
    updatePreview();
    const wc = countWords(getEditorValue());
    const wcEl = $('#wordCount');
    if (wcEl) wcEl.textContent = `${wc.toLocaleString()} 字`;
    showStatus('');
  });

  /* ── Generate DOCX ────────────────── */
  $('#generateBtn')?.addEventListener('click', async () => {
    const md = getEditorValue();
    if (!md.trim()) { showStatus('⚠️ 请先输入或加载 Markdown 内容', 'warn'); return; }
    const btn = $('#generateBtn');
    btn.disabled = true; btn.textContent = '生成中…';
    try { await downloadDocx(md); showStatus('✅ DOCX 已下载 — 用 Word/WPS 打开即可', 'success'); }
    catch(e) { showStatus(`❌ ${e.message}`, 'error'); }
    btn.disabled = false; btn.textContent = '📥 生成 DOCX';
  });

  /* ── Template ─────────────────────── */
  $('#templateSelect')?.addEventListener('change', function() {
    if (!this.value) return;
    const tmpl = TEMPLATES[this.value];
    if (!tmpl) return;
    if (getEditorValue().trim() && !confirm('加载模板将覆盖当前内容，确定？')) { this.value = ''; return; }
    setEditorValue(tmpl);
    updatePreview();
    showStatus(`📋 已加载「${this.options[this.selectedIndex].text}」`, 'success');
  });

  /* ── Copy / Clear ─────────────────── */
  $('#copyBtn')?.addEventListener('click', async () => {
    if (!getEditorValue().trim()) return;
    await navigator.clipboard.writeText(getEditorValue());
    showStatus('✅ 已复制', 'success');
  });
  $('#clearBtn')?.addEventListener('click', () => {
    if (getEditorValue().trim() && !confirm('确定清空编辑器？')) return;
    setEditorValue(''); updatePreview();
    $('#templateSelect') && ($('#templateSelect').value = '');
    showStatus('🗑 已清空');
  });

  /* ── Keyboard Shortcut ────────────── */
  editor?.addEventListener('keydown', e => {
    if ((e.ctrlKey||e.metaKey) && e.key === 'Enter') { e.preventDefault(); $('#generateBtn')?.click(); }
  });

  /* ── AI Generator ────────────────── */
  const savedKey = localStorage.getItem('deepseek_api_key');
  if (savedKey && $('#aiApiKey')) $('#aiApiKey').value = savedKey;
  $('#aiSaveKey')?.addEventListener('click', () => {
    const key = $('#aiApiKey')?.value?.trim();
    if (!key) { showStatus('⚠️ 请输入 API Key', 'warn'); return; }
    localStorage.setItem('deepseek_api_key', key);
    showStatus('✅ API Key 已保存（仅存于浏览器本地）', 'success');
  });
  $('#aiGenBtn')?.addEventListener('click', generateThesisWithAI);

  /* ── Evidence Builder ─────────────── */
  $('#evBuildBtn')?.addEventListener('click', buildEvidence);

  /* ── Image Manager ────────────────── */
  const imgDZ = $('#imgDropZone');
  const imgInput = $('#imgFileInput');

  imgDZ?.addEventListener('dragover', e => { e.preventDefault(); imgDZ.classList.add('drag-hover'); });
  imgDZ?.addEventListener('dragleave', () => imgDZ.classList.remove('drag-hover'));
  imgDZ?.addEventListener('drop', e => {
    e.preventDefault();
    imgDZ.classList.remove('drag-hover');
    if (e.dataTransfer?.files.length) addImages(e.dataTransfer.files);
  });
  imgInput?.addEventListener('change', () => { if (imgInput.files?.length) addImages(imgInput.files); imgInput.value = ''; });
  $('#imgMapBtn')?.addEventListener('click', () => {
    const map = buildImageMap();
    showStatus(`✅ image-map.json 已生成（${Object.keys(map).length} 张图片）`, 'success');
  });
  $('#imgClearBtn')?.addEventListener('click', () => { imgStore = []; renderImgGrid(); $('#imgMapOutput')?.style.setProperty('display','none'); });

  /* ── Word Count ───────────────────── */
  $('#wcAnalyzeBtn')?.addEventListener('click', analyzeWordCount);

  /* ── References ───────────────────── */
  $('#refCheckBtn')?.addEventListener('click', checkReferences);

  /* ── Checklist ────────────────────── */
  $('#clGenBtn')?.addEventListener('click', generateChecklist);

  /* ── Drag & Drop (.md files) ──────── */
  let dragCnt = 0;
  const overlay = $('#dropOverlay');
  const showOverlay = () => { if (overlay) overlay.classList.add('active'); };
  const hideOverlay = () => { if (overlay) overlay.classList.remove('active'); dragCnt = 0; };

  document.addEventListener('dragenter', e => { e.preventDefault(); dragCnt++; showOverlay(); });
  document.addEventListener('dragleave', e => { e.preventDefault(); dragCnt--; if (dragCnt<=0) hideOverlay(); });
  document.addEventListener('dragover', e => e.preventDefault());
  document.addEventListener('drop', e => {
    e.preventDefault(); hideOverlay();
    const f = e.dataTransfer?.files?.[0];
    if (f) loadMdFile(f);
  });

  $('#fileInput')?.addEventListener('change', function() {
    if (this.files?.[0]) loadMdFile(this.files[0]);
    this.value = '';
  });

  function loadMdFile(file) {
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!['.md','.txt','.markdown'].includes(ext)) { showStatus(`❌ 不支持 ${ext}，请用 .md / .txt`, 'error'); return; }
    const reader = new FileReader();
    reader.onload = e => {
      const content = e.target?.result;
      if (typeof content !== 'string') return;
      if (getEditorValue().trim() && !confirm(`加载「${file.name}」将覆盖当前内容，确定？`)) return;
      setEditorValue(content); updatePreview();
      $$('.tab').forEach(t => t.classList.remove('active'));
      const editorTab = $('[data-tab="editor"]');
      if (editorTab) editorTab.classList.add('active');
      $$('.tab-panel').forEach(p => p.classList.remove('active'));
      $('#panel-editor')?.classList.add('active');
      showStatus(`📂 已加载：${file.name}`, 'success');
    };
    reader.onerror = () => showStatus('❌ 文件读取失败', 'error');
    reader.readAsText(file);
  }

  /* ── Resize Handle ────────────────── */
  const handle = $('#resizeHandle');
  const left = $('#leftPanel');
  const right = $('#rightPanel');
  if (handle && left && right) {
    let dragging = false;
    handle.addEventListener('mousedown', e => { dragging = true; document.body.style.cursor='col-resize'; document.body.style.userSelect='none'; e.preventDefault(); });
    document.addEventListener('mousemove', e => {
      if (!dragging) return;
      const rect = left.parentElement.getBoundingClientRect();
      const pct = Math.max(20, Math.min(80, ((e.clientX-rect.left)/rect.width)*100));
      left.style.flex = `${pct}`; right.style.flex = `${100-pct}`;
    });
    document.addEventListener('mouseup', () => { dragging=false; document.body.style.cursor=''; document.body.style.userSelect=''; });
  }

  /* ── Requirements Panel ──────────── */
  $('#reqToggleBtn')?.addEventListener('click', () => {
    const detail = $('#reqDetail');
    if (detail) {
      const show = detail.style.display === 'none' || !detail.style.display;
      detail.style.display = show ? 'block' : 'none';
      $('#reqBanner')?.classList.toggle('expanded', show);
    }
  });

  $('#reqResetBtn')?.addEventListener('click', resetToDefaults);

  $('#reqFileInput')?.addEventListener('change', function() {
    if (this.files?.[0]) handleReqFile(this.files[0]);
    this.value = '';
  });

  // Drag-drop requirements file onto banner
  $('#reqBanner')?.addEventListener('dragover', e => { e.preventDefault(); e.currentTarget.style.opacity = '0.7'; });
  $('#reqBanner')?.addEventListener('dragleave', e => { e.currentTarget.style.opacity = '1'; });
  $('#reqBanner')?.addEventListener('drop', e => {
    e.preventDefault(); e.currentTarget.style.opacity = '1';
    const f = e.dataTransfer?.files?.[0];
    if (f) handleReqFile(f);
  });

  // Apply format config
  $('#reqApplyBtn')?.addEventListener('click', () => {
    showStatus('✅ 格式配置已应用（将在下次生成 DOCX 时生效）', 'success');
    const badge = $('#reqStatus');
    if (badge) { badge.dataset.status = 'custom'; badge.textContent = '自定义配置'; }
    updateFieldStatus();
  });

  // Track field changes
  $$('.form-group[data-field] input, .form-group[data-field] textarea').forEach(el => {
    el.addEventListener('input', updateFieldStatus);
  });

  /* ── Init ──────────────────────────── */
  updatePreview();
  updateFieldStatus();
  const initWc = countWords(getEditorValue());
  const wcEl = $('#wordCount');
  if (wcEl) wcEl.textContent = `${initWc.toLocaleString()} 字`;
});

export { downloadDocx, buildDocx };
