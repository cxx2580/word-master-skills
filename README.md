# Docx Master

Generate `.docx` Word documents from JSON — with template library and document scanner.
Zero manual setup. Auto-installs dependencies on first run.

```
Describe → Match template / Scan existing → Fill → Generate .docx
```

## What It Does

- **Generate** `.docx` from JSON (headings, paragraphs, tables, bullets, images)
- **Scan** existing `.docx` to extract structure as reusable JSON
- **Template library** — 6 built-in templates (report, letter, memo, table report, resume, meeting minutes)
- **Auto-install** — detects Python/Node, installs python-docx or JSZip as needed
- **Dual engine** — Python (python-docx, best quality) or Node.js (JSZip, universal)

## Quick Start

### As a Claude Code skill

```bash
git clone https://github.com/YOU/word-master-skills.git ~/.claude/skills/docx-master
```

Then ask Claude: "生成一份实验报告". The skill handles everything.

### Standalone CLI

```bash
git clone https://github.com/YOU/word-master-skills.git
cd word-master-skills

# Scan an existing docx to JSON
python scripts/scan_docx.py input.docx template.json

# Edit template.json, then generate
python scripts/generate_docx.py output.docx template.json

# Or use Node.js (no Python needed)
node scripts/generate_docx.js output.docx content.json
```

## Requirements

One of:

| Runtime | Auto-installed dep |
|---------|--------------------|
| Python ≥ 3.8 | `python-docx` |
| Node.js ≥ 12 | `jszip` |

Scanner requires Python. Generator works with either.

## JSON Format

```json
{
  "title": "Monthly Report",
  "sections": [
    {
      "heading": "Summary",
      "level": 1,
      "content": [
        { "type": "paragraph", "text": "Revenue up 15%", "bold": true },
        { "type": "bullet", "text": "New customers: 240" },
        { "type": "table",
          "headers": ["Month", "Revenue"],
          "rows": [["Jan", "$50K"], ["Feb", "$55K"]]
        }
      ]
    }
  ]
}
```

## Content Types

| Type | Description | Key options |
|------|-------------|-------------|
| `paragraph` | Body text | `bold`, `italic`, `align`, `sz`, `color` |
| `heading` | Section heading (1-4) | `level` |
| `bullet` | Bullet list item | — |
| `table` | Grid with headers + rows | `colWidths` |
| `image` | Embedded picture | `width`, `height`, `caption` |

## Templates

Six pre-built templates with `{{PLACEHOLDER}}` slots:

| Template | Use case |
|----------|----------|
| `report.json` | Project reports, lab reports, work summaries |
| `letter.json` | Business letters, applications, notifications |
| `memo.json` | Internal memos, work logs, to-do lists |
| `table_report.json` | Data reports, statistics, financial tables |
| `resume.json` | CVs / resumes (Chinese style) |
| `meeting_minutes.json` | Meeting records, decisions, action items |

## Scanner

Reverse-engineer any `.docx` into a JSON template:

```bash
python scripts/scan_docx.py existing.docx structure.json
```

Extracts: title, headings (1-4), paragraphs (with bold/italic/align/color), bullet lists, tables (headers + data). Works with Word, WPS, and LibreOffice output.

## Engine Choice

| Engine | Quality | Startup | Best for |
|--------|---------|---------|----------|
| **python-docx** | Production | ~0.3s | General use, Chinese docs |
| **Node + JSZip** | Good | ~0.1s | No-Python environments, CI/CD |

Both produce valid `.docx`. Python is preferred when available.

## File Structure

```
word-master-skills/
├── SKILL.md                    # Claude Code skill definition
├── README.md                   # This file
├── .gitignore
├── scripts/
│   ├── generate_docx.py        # Python generator
│   ├── generate_docx.js        # Node.js generator
│   └── scan_docx.py            # Document scanner
└── templates/
    ├── report.json
    ├── letter.json
    ├── memo.json
    ├── table_report.json
    ├── resume.json
    └── meeting_minutes.json
```

## License

MIT
