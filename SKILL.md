---
name: docx-master
description: >
  Zero-setup .docx generator with template library and document scanner.
  Describe what you want — or scan an existing .docx — then generate.
  Auto-installs python-docx or JSZip. No manual setup needed.
---

# Docx Master

> Describe → scan → template-match → generate.  Zero config.

**Full pipeline**: `Auto-install Engine → Match Template / Scan → Fill Content → Generate .docx → Verify`

---

## Quick Start (for Claude users)

1. Put this folder in `~/.claude/skills/docx-master/`
2. Ask Claude: "生成一份实验报告" / "create a letter" / "把这张表做成 Word"
3. Done. No `pip install`, no `npm install`.

---

## Pipeline Decision Tree

```
User asks for a document
  │
  ├─ User provides an existing .docx as reference?
  │     └─ YES → SCAN it:  python scripts/scan_docx.py input.docx template.json
  │              → Review extracted JSON, fill content, generate
  │
  ├─ Document type matches a known template?
  │     └─ YES → LOAD templates/<type>.json
  │              → Replace {{PLACEHOLDERS}} with user's content
  │              → Generate
  │
  └─ No template fits?
        └─ BUILD JSON from scratch
           → Ask 2-3 clarifying questions if vague
           → Generate
```

---

## Engine Bootstrap (auto-install, once per session)

```bash
# Ensure python-docx
python -c "from docx import Document" 2>/dev/null \
  || python -m pip install python-docx -q 2>/dev/null \
  || python3 -c "from docx import Document" 2>/dev/null \
  || python3 -m pip install python-docx -q 2>/dev/null \
  || true

# Ensure Node + JSZip
node -e "require('jszip')" 2>/dev/null \
  || (cd "${SKILL_DIR}/scripts" && npm install jszip --no-save 2>/dev/null) \
  || true
```

Detect engine:

```bash
python -c "from docx import Document; print('python-docx')" 2>/dev/null \
  || python3 -c "from docx import Document; print('python-docx')" 2>/dev/null \
  || node -e "require('jszip'); console.log('node-jszip')" 2>/dev/null \
  || echo "none"
```

| Result | Engine | Priority |
|--------|--------|----------|
| `python-docx` | Python | Preferred — best formatting |
| `node-jszip` | Node.js | Fallback — no Python needed |
| `none` | — | Tell user to install Python or Node |

---

## Tool 1: Document Scanner

Reverse-engineer a `.docx` into a JSON template.

```bash
python "${SKILL_DIR}/scripts/scan_docx.py" <input.docx> [output.json]
```

What it extracts:

| Element | Detection method |
|---------|-----------------|
| Title | Title style, or large centered bold text |
| Headings | Heading 1-4 styles (supports WPS variants) |
| Paragraphs | Normal style — preserves bold, italic, alignment, font size, color |
| Bullets | List Bullet / List Paragraph styles |
| Tables | All tables — first row → headers, remaining → data rows |

Use cases:
- "把这个合同模板改成新的" → scan old → fill new fields → generate
- "参考这份报告的格式，写一个新的" → scan → keep structure → replace content
- "帮我看看这个 docx 的结构" → scan → print JSON

After scanning, the LLM reviews the JSON, replaces placeholder text with user content, then feeds it to `generate_docx.py`.

---

## Tool 2: Template Library

Six pre-built templates in `templates/`. Each uses `{{PLACEHOLDER}}` syntax.

| Template | File | When to use |
|----------|------|-------------|
| 通用报告 | `report.json` | 项目报告、实验报告、工作总结 |
| 正式信函 | `letter.json` | 商务信函、申请书、通知函 |
| 备忘录 | `memo.json` | 内部备忘、工作记录、待办事项 |
| 表格报告 | `table_report.json` | 数据报告、统计表、财务报表 |
| 简历 | `resume.json` | 个人简历、CV |
| 会议纪要 | `meeting_minutes.json` | 会议记录、决议跟踪 |

### Template matching logic

When the user describes what they want, run this mental checklist:

1. **Contains "报告" / "report" / "实验" / "总结"** → `report.json`
2. **Contains "信" / "letter" / "申请" / "通知" / "函"** → `letter.json`
3. **Contains "备忘" / "memo" / "记录" / "待办"** → `memo.json`
4. **Contains "表" / "数据" / "统计" / "table" / "报表"** → `table_report.json`
5. **Contains "简历" / "resume" / "CV" / "求职"** → `resume.json`
6. **Contains "会议" / "meeting" / "纪要" / "讨论"** → `meeting_minutes.json`
7. **Otherwise** → Build JSON from scratch

### Template workflow

```
1. Read the matched template file
2. For each {{PLACEHOLDER}}:
   - If user provided the value → replace it
   - If not → ask the user (or infer from context if obvious)
3. Remove any unused optional rows/items
4. Write filled JSON → generate .docx
```

Templates are just JSON files. The LLM can add/remove sections, tables, or bullets freely — the template is a starting point, not a straitjacket.

---

## Tool 3: Generate

```bash
# Python (preferred)
python "${SKILL_DIR}/scripts/generate_docx.py" <output.docx> <content.json>

# Node.js fallback
node "${SKILL_DIR}/scripts/generate_docx.js" <output.docx> <content.json>
```

Always absolute paths. Temp JSON is cleaned up after generation.

---

## JSON Content Format

```json
{
  "title": "Document Title",
  "sections": [
    {
      "heading": "Section Heading (optional)",
      "level": 1,
      "content": [
        { "type": "paragraph", "text": "Normal text..." },
        { "type": "paragraph", "text": "Bold text", "bold": true },
        { "type": "paragraph", "text": "Centered", "align": "center" },
        { "type": "bullet", "text": "List item" },
        { "type": "heading", "text": "Sub-heading", "level": 2 },
        { "type": "table", "headers": ["A","B"], "rows": [["1","2"]] },
        { "type": "image", "path": "/abs/path/img.png", "caption": "Fig 1" }
      ]
    }
  ]
}
```

### Content types

| `type` | Required | Optional |
|--------|----------|----------|
| `paragraph` | `text` | `bold`, `italic`, `align` (left/center/right), `sz` (half-pts: 24=12pt), `color` (hex) |
| `heading` | `text`, `level` (1-4) | — |
| `bullet` | `text` | — |
| `table` | `headers`, `rows` | `colWidths` (EMU array) |
| `image` | `path` (abs path) | `width`, `height` (EMU, default 5400000×3000000), `caption` |

### EMU reference

| Size | EMU |
|------|-----|
| A4 full width (17 cm) | 6120000 |
| Half width | 3060000 |
| 1 inch | 914400 |
| 1 cm | 360000 |

---

## Engines Compared

| Feature | python-docx | Node + JSZip |
|---------|-------------|--------------|
| Formatting | Pt/Emu/RGB objects, styles | Raw OOXML XML strings |
| Fonts | East-Asian + Latin dual config | Hardcoded 宋体 + Arial |
| Tables | Table Grid style + per-cell width | Hand-rolled OOXML |
| Images | `add_picture()` native API | Manual zip + rels entries |
| Output quality | Production-grade | Good, minor edge cases |
| Dependency | `pip install python-docx` | `npm install jszip` |

Both produce valid `.docx`. Python engine produces better-styled output.
When both are available, **always prefer Python**.

> Note: The scanner (`scan_docx.py`) requires Python. It will auto-install python-docx.

---

## File Output

- Default: user's working directory
- Name: descriptive, based on content
- After generation: report path + size

---

## Post-Generation

1. Confirm file path and size
2. Optional: `unzip -q <file.docx>` checks zip integrity
3. User opens with Word / WPS / LibreOffice
