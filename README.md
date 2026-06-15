# submission-format

A [Claude Code](https://claude.com/claude-code) skill that converts a manuscript
(Markdown / LaTeX / Word / HTML — anything pandoc can read) into a **clean
journal-submission Word (.docx) format**.

"Clean" means flat `Normal`-only paragraphs carrying **direct formatting**
(not Word semantic styles like Heading 1 / Body Text). The output reproduces a
fixed submission style:

- Times New Roman 11pt, A4, justified body, `docGrid` line pitch
- 14pt bold section headings / 11pt bold subsection headings
- structured Summary abstract (independent bold sub-labels)
- explicit page breaks before each major section (declaration blocks kept together)
- structured title page: authors in one line with affiliation superscripts,
  `*`/`**` corresponding marks, auto-lettered affiliations, corresponding-author lines
- full-grid 9pt tables, **wide tables auto-rotated to landscape**, bold headers,
  mid-dot decimals
- "Figure N." / "Table N." captions (figure caption below, table caption above)
- left-aligned, auto-numbered references

It is a **general engine**: content is extracted from the pandoc AST and rendered
against a fixed format spec. Nothing about any specific manuscript is hard-coded.

## Usage

```bash
python3 scripts/build_submission.py INPUT -o OUT.docx \
    [--resource-path DIR]...      # folders to search for figures (also for LaTeX \input)
    [--bibliography FILE]         # standalone references source (thebibliography / pandoc-readable)
    [--styles STYLES.docx]        # style asset pack (default assets/reference_styles.docx)
    [--landscape-fit F]           # natural total width > portrait width × F → landscape (default 1.6)
    [--landscape-mincols N]       # optional override: >= N columns forces landscape (default 99 = off)
    [--body-only]                 # fragment mode: skip title page / Summary, render body only
```

`INPUT` can be `.md` / `.tex` / `.docx` / `.html` / … (for PDF, extract text first).
Requires **pandoc**; LibreOffice + `pdftoppm` optional, only for render-checking.

The engine only renders content that **actually exists in the input** — figure files
must be reachable via `--resource-path`, references via the source or `--bibliography`,
and multi-file LaTeX needs its `\input` paths resolvable (run from the project root).

## Layout

| Path | Purpose |
|------|---------|
| `SKILL.md` | skill instructions (loaded by Claude Code) |
| `scripts/build_submission.py` | the engine (Python + pandoc) |
| `assets/reference_styles.docx` | style asset pack — styles + numbering, no content |
| `references/target-format-spec.md` | the target format specification (numeric) |

## Install

Clone into your Claude Code skills directory:

```bash
git clone https://github.com/buergerSUN-CN/submission-format \
    ~/.claude/skills/submission-format
```
