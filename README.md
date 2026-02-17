# Crowe Document Formatter — Audit & Convert Tool

## Overview

This Python tool takes any `.docx` document and either:
- **AUDIT** — checks it against the Crowe template formatting rules, highlighting violations in yellow
- **CONVERT** — rewrites the content into the branded Crowe template (with logo, amber background, and correct styles)

## What's Preserved Automatically

| Element | How It's Handled |
|---------|-----------------|
| **Crowe logo** (`crowe_logo.png`, `crowe_logo.svg`) | Embedded in the template header; stays on every page |
| **Amber/gold gradient** (`background.jpg`) | Full-page background image anchored in headers |
| **Cover page** (Section 1) | Kept exactly as-is; only the title text is changeable |
| **Back page** (Section 3) | Kept exactly as-is (Contact Info, About Us, legal disclaimer) |
| **Footer tagline** | "Smart decisions. Lasting value." stays on all pages |
| **Styles & numbering** | All 6 custom styles + numbering definitions from the template |

## Setup

```bash
# Install dependencies
pip install lxml python-docx

# Ensure the template_assets folder contains:
#   template.docx      — the original Crowe template
#   background.jpg     — the amber gradient image (extracted)
#   crowe_logo.png     — the Crowe logo (extracted)
#   crowe_logo.svg     — the Crowe logo SVG (extracted)
```

## Usage

### Audit Mode
Checks a document for formatting violations and highlights problems in yellow:

```bash
python crowe_formatter.py audit  input.docx  audited_output.docx
```

### Convert Mode
Converts any user document into the Crowe branded format:

```bash
python crowe_formatter.py convert  input.docx  converted_output.docx --title "My Report Title"
```

If `--title` is omitted, the original template title is kept.

## Audit Rules Checked

| # | Rule | What Gets Flagged |
|---|------|-------------------|
| 1 | Font must be Arial | Any non-Arial font in text |
| 2 | Body text = 10pt | Body text not using 10pt (sz=20) |
| 3 | H1 headings = 18pt bold | HeadingStyle1 with wrong size |
| 4 | No unicode bullets | Inline bullet chars (•, ▪, etc.) instead of Word numbering |
| 5 | Line spacing standard | Non-standard line spacing values |
| 6 | Indent values standard | Non-standard indent for bulleted content |

## Content Hierarchy Mapping

When converting, user content is mapped to the Crowe format hierarchy:

| User Content | Crowe Format |
|-------------|-------------|
| H1 / Title headings | `HeadingStyle1-18pt` — 18pt bold (section titles) |
| H2 / Sub-headings | `HeadingStyle2-14pt` — 12pt bold (subtopic headings) |
| H3 / Bold sub-sub-headings | `BodyCopy-Arial10pt` bold (inline subheads) |
| Bold bullet + description | Bold bullet heading → indented body paragraph |
| Regular bullets | `BodyCopy-Arial10pt` with Word numbering |
| Body text | `BodyCopy-Arial10pt` — 10pt normal |

## File Structure

```
crowe_tool/
├── crowe_formatter.py          # Main tool (audit + convert)
├── template_assets/
│   ├── template.docx           # Original Crowe template (DO NOT MODIFY)
│   ├── background.jpg          # Amber gradient (199 KB, full-page)
│   ├── crowe_logo.png          # Crowe logo raster (25 KB)
│   └── crowe_logo.svg          # Crowe logo vector (2 KB)
└── README.md                   # This file
```

## Notes

- The template is A4 sized (11906 × 16838 DXA)
- All images are embedded in the template's headers via anchored drawings
- The cover page subtitle in the footer may need manual updating for different reports
- Numbering IDs 55 and 56 from the template are used for converted bullets
