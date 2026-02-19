#!/usr/bin/env python3
"""
Crowe Document Formatter - Audit & Convert Tool
================================================

This tool takes a user-uploaded .docx document and either:
  1. AUDIT  – flags formatting violations by highlighting lines in yellow
  2. CONVERT – rewrites the content into the Crowe branded template

The Crowe template structure:
  Section 1 (Cover Page)  – paragraphs 0-6   → KEPT AS-IS (only title is editable)
  Section 2 (Content)     – paragraphs 7-179 → REPLACED with user content
  Section 3 (Back Page)   – paragraphs 180+  → KEPT AS-IS

Usage:
    python crowe_formatter.py audit   input.docx  output_audit.docx
    python crowe_formatter.py convert input.docx  output_converted.docx --title "My Report Title"
"""

import argparse
import copy
import logging
import os
import re
import shutil
import sys
import tempfile
import zipfile
from lxml import etree
from pathlib import Path

logger = logging.getLogger(__name__)

# ─── Namespace map ──────────────────────────────────────────────────────────
NSMAP = {
    'w':    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':   'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':    'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':  'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'mc':   'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w14':  'http://schemas.microsoft.com/office/word/2010/wordml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
}

W = NSMAP['w']

def wn(tag):
    """Create a tag in the w: namespace."""
    return f'{{{W}}}{tag}'

def wattr(attr):
    """Create an attribute name in the w: namespace."""
    return f'{{{W}}}{attr}'

# ─── TEMPLATE CONSTANTS ─────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).parent
TEMPLATE_PATH = SCRIPT_DIR / 'template_assets' / 'template.docx'

# Section boundary paragraph indices (0-based within all <w:p> elements)
SECTION1_END = 6      # Cover page: P0–P6 (inclusive)
SECTION3_START = 180   # Back page: P180 onwards

# Style IDs
STYLE_H1          = 'HeadingStyle1-18pt'
STYLE_H2          = 'HeadingStyle2-14pt'
STYLE_BODY        = 'BodyCopy-Arial10pt'
STYLE_COVER_TITLE = 'CoverText-Aprial18pt'

# Numbering abstract IDs used in the template
BULLET_ABSTRACT_ID = '59'  # abstractNumId for standard bullets

# Table shading colors (matching Crowe template branding)
TABLE_HEADER_FILL = 'F5A800'    # Amber/gold for header row
TABLE_BODY_FILL   = 'FDF1E7'    # Light cream for data rows
TABLE_BORDER_COLOR = '011E41'   # Dark navy for borders


# ═══════════════════════════════════════════════════════════════════════════
# PARSER — Read user-uploaded .docx into structured content
# ═══════════════════════════════════════════════════════════════════════════

class ContentBlock:
    """Represents one parsed content element from the user document."""
    def __init__(self, level, text, bold=False, children=None):
        self.level = level        # 'h1', 'h2', 'h3', 'bullet', 'bullet_bold', 'body', 'body_indent', 'table'
        self.text = text
        self.bold = bold
        self.children = children or []  # For bullet_bold + description pairs

    def __repr__(self):
        return f"ContentBlock({self.level}, '{self.text[:50]}...', bold={self.bold})"


class TableBlock:
    """Represents a table parsed from the user document."""
    def __init__(self, headers, rows, col_widths=None):
        self.level = 'table'
        self.headers = headers      # List of cell texts for header row
        self.rows = rows            # List of lists of cell texts for data rows
        self.col_widths = col_widths  # List of column widths in twips (optional)
        self.text = ''              # For compatibility with merge logic
        self.bold = False
        self.children = []

    def __repr__(self):
        return f"TableBlock({len(self.headers)} cols, {len(self.rows)} rows)"


def parse_user_document(docx_path):
    """
    Parse a user-uploaded .docx into a list of ContentBlock objects.
    Detects headings, bullets, bold text, and body paragraphs.
    """
    tmp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            z.extractall(tmp_dir)

        doc_xml_path = os.path.join(tmp_dir, 'word', 'document.xml')
        tree = etree.parse(doc_xml_path)
        root = tree.getroot()
        body = root.find(f'.//{wn("body")}')

        blocks = []
        for child in body:
            # Handle tables
            if child.tag == wn('tbl'):
                table_block = _parse_table(child)
                if table_block:
                    blocks.append(table_block)
                continue

            # Handle paragraphs
            if child.tag != wn('p'):
                continue

            para = child
            text = _get_para_text(para)
            if not text.strip():
                continue

            ppr = para.find(wn('pPr'))
            style = _get_style(ppr)
            has_numbering = _has_numbering(ppr)
            is_bold = _is_para_bold(para)
            heading_level = _detect_heading_level(style, ppr, para)

            if heading_level == 1:
                blocks.append(ContentBlock('h1', text, bold=True))
            elif heading_level == 2:
                blocks.append(ContentBlock('h2', text, bold=True))
            elif heading_level == 3:
                blocks.append(ContentBlock('h3', text, bold=True))
            elif has_numbering and is_bold:
                blocks.append(ContentBlock('bullet_bold', text, bold=True))
            elif has_numbering:
                blocks.append(ContentBlock('bullet', text))
            elif is_bold:
                blocks.append(ContentBlock('h3', text, bold=True))
            else:
                blocks.append(ContentBlock('body', text))

        # Post-process: merge bullet_bold + following body into pairs
        merged = _merge_bullet_descriptions(blocks)
        return merged

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _get_para_text(para):
    """Extract all text from a paragraph."""
    texts = []
    for t in para.iter(wn('t')):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)


def _parse_table(tbl):
    """Parse a w:tbl element into a TableBlock."""
    # Get grid columns for widths
    grid = tbl.find(wn('tblGrid'))
    col_widths = []
    if grid is not None:
        for gc in grid.findall(wn('gridCol')):
            w = gc.get(wattr('w'))
            if w:
                try:
                    col_widths.append(int(w))
                except ValueError:
                    pass

    rows_data = []
    for tr in tbl.findall(wn('tr')):
        row_cells = []
        for tc in tr.findall(wn('tc')):
            cell_text = _get_cell_text(tc)
            row_cells.append(cell_text)
        if row_cells:
            rows_data.append(row_cells)

    if not rows_data:
        return None

    # First row is treated as header
    headers = rows_data[0]
    data_rows = rows_data[1:]

    return TableBlock(headers, data_rows, col_widths if col_widths else None)


def _get_cell_text(tc):
    """Extract all text from a table cell, joining paragraphs with newlines."""
    texts = []
    for p in tc.findall(wn('p')):
        t = _get_para_text(p)
        if t.strip():
            texts.append(t.strip())
    return '\n'.join(texts)


def _get_style(ppr):
    """Get the style ID from paragraph properties."""
    if ppr is None:
        return None
    pstyle = ppr.find(wn('pStyle'))
    if pstyle is not None:
        return pstyle.get(wattr('val'))
    return None


def _has_numbering(ppr):
    """Check if paragraph has numbering (bullets/lists)."""
    if ppr is None:
        return False
    numpr = ppr.find(wn('numPr'))
    if numpr is not None:
        numid = numpr.find(wn('numId'))
        if numid is not None:
            val = numid.get(wattr('val'))
            return val is not None and val != '0'
    return False


def _is_para_bold(para):
    """Check if the first run of a paragraph is bold."""
    runs = para.findall(wn('r'))
    for r in runs:
        rpr = r.find(wn('rPr'))
        if rpr is not None:
            b = rpr.find(wn('b'))
            if b is not None:
                val = b.get(wattr('val'), 'true')
                if val != '0' and val != 'false':
                    return True
        # Also check if style itself is a heading (inherently bold)
        break
    return False


def _detect_heading_level(style, ppr, para):
    """
    Detect the heading level from style name and formatting.
    Returns 1, 2, 3, or 0 (not a heading).
    """
    if style is None:
        style = ''
    style_lower = style.lower()

    # Built-in heading styles
    if style_lower in ('heading1', 'heading 1') or style == STYLE_H1:
        return 1
    if style_lower in ('heading2', 'heading 2'):
        return 2
    if style_lower in ('heading3', 'heading 3'):
        return 3

    # Detect by outline level
    if ppr is not None:
        outline = ppr.find(wn('outlineLvl'))
        if outline is not None:
            lvl = int(outline.get(wattr('val'), '9'))
            if lvl == 0:
                return 1
            if lvl == 1:
                return 2
            if lvl == 2:
                return 3

    # Detect by font size in runs (large text = heading)
    max_size = 0
    for r in para.findall(wn('r')):
        rpr = r.find(wn('rPr'))
        if rpr is not None:
            sz = rpr.find(wn('sz'))
            if sz is not None:
                try:
                    max_size = max(max_size, int(sz.get(wattr('val'), '0')))
                except ValueError:
                    pass

    if max_size >= 36:  # 18pt+
        return 1
    if max_size >= 28:  # 14pt+
        return 2
    if max_size >= 24 and _is_para_bold(para):  # 12pt bold
        return 3

    return 0


def _merge_bullet_descriptions(blocks):
    """
    Merge consecutive bullet_bold + body into paired structures.
    Also group body paragraphs that follow bullet_bold as descriptions.
    """
    merged = []
    i = 0
    while i < len(blocks):
        block = blocks[i]
        if block.level == 'bullet_bold':
            # Look ahead for body/description lines
            descriptions = []
            j = i + 1
            while j < len(blocks) and blocks[j].level == 'body':
                descriptions.append(blocks[j].text)
                j += 1
            block.children = descriptions
            merged.append(block)
            i = j
        else:
            merged.append(block)
            i += 1
    return merged


# ═══════════════════════════════════════════════════════════════════════════
# AUDIT — Check formatting and flag violations in yellow
# ═══════════════════════════════════════════════════════════════════════════

# Yellow highlight color code
YELLOW_HIGHLIGHT = '7'  # OOXML highlight yellow index

class AuditResult:
    """Holds audit findings."""
    def __init__(self):
        self.issues = []  # List of (para_index, description)

    def add(self, idx, desc):
        self.issues.append((idx, desc))

    def summary(self):
        if not self.issues:
            return "✅ No formatting issues found."
        lines = [f"⚠️  Found {len(self.issues)} formatting issue(s):\n"]
        for idx, desc in self.issues:
            lines.append(f"  P{idx}: {desc}")
        return '\n'.join(lines)


def audit_document(input_path, output_path):
    """
    Audit a user document for Crowe template formatting violations.
    Creates a copy with yellow-highlighted problem lines.
    Returns an AuditResult.
    """
    result = AuditResult()

    # Work on a copy
    shutil.copy2(input_path, output_path)

    tmp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(output_path, 'r') as z:
            z.extractall(tmp_dir)

        doc_xml_path = os.path.join(tmp_dir, 'word', 'document.xml')
        tree = etree.parse(doc_xml_path)
        root = tree.getroot()
        body = root.find(f'.//{wn("body")}')

        paragraphs = body.findall(wn('p'))

        for i, para in enumerate(paragraphs):
            issues = _audit_paragraph(i, para)
            for issue_desc in issues:
                result.add(i, issue_desc)
                _highlight_paragraph_yellow(para)

        # Write back
        tree.write(doc_xml_path, xml_declaration=True, encoding='UTF-8',
                   standalone=True)

        # Repack
        _repack_docx(tmp_dir, output_path)

        return result

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _audit_paragraph(idx, para):
    """Check a single paragraph for formatting violations. Returns list of issue strings."""
    issues = []
    text = _get_para_text(para)
    if not text.strip():
        return issues

    ppr = para.find(wn('pPr'))
    style = _get_style(ppr)

    # Rule 1: Font must be Arial
    for r in para.findall(wn('r')):
        rpr = r.find(wn('rPr'))
        if rpr is not None:
            rfonts = rpr.find(wn('rFonts'))
            if rfonts is not None:
                for attr_name in ['ascii', 'hAnsi', 'cs']:
                    font = rfonts.get(wattr(attr_name))
                    if font and font.lower() != 'arial' and font.lower() != 'symbol':
                        issues.append(f"Font '{font}' used — must be Arial")
                        break

    # Rule 2: Body text size must be 10pt (sz=20)
    valid_styles = [STYLE_BODY]
    if style in valid_styles:
        for r in para.findall(wn('r')):
            rpr = r.find(wn('rPr'))
            if rpr is not None:
                sz = rpr.find(wn('sz'))
                if sz is not None:
                    val = sz.get(wattr('val'))
                    if val and val not in ('20', '24'):  # 10pt or 12pt (subtopic override)
                        issues.append(f"Body text size sz={val} — expected 20 (10pt)")

    # Rule 3: HeadingStyle1 must be 18pt bold
    if style == STYLE_H1:
        for r in para.findall(wn('r')):
            rt = _get_run_text(r)
            if not rt.strip():
                continue
            rpr = r.find(wn('rPr'))
            if rpr is not None:
                sz = rpr.find(wn('sz'))
                if sz is not None:
                    val = sz.get(wattr('val'))
                    if val and val != '36':
                        issues.append(f"H1 heading size sz={val} — expected 36 (18pt)")

    # Rule 4: Check for unicode bullet characters in text (should use numbering)
    if '•' in text or '▪' in text or '‣' in text or '►' in text:
        has_num = _has_numbering(ppr)
        if not has_num:
            issues.append("Unicode bullet character in text — use Word numbering instead")

    # Rule 5: Line spacing checks
    if ppr is not None and style in [STYLE_BODY, STYLE_H2]:
        spacing = ppr.find(wn('spacing'))
        if spacing is not None:
            line = spacing.get(wattr('line'))
            if line:
                try:
                    line_val = int(line)
                    if style == STYLE_BODY and line_val not in (240, 276, 288):
                        issues.append(f"Body line spacing {line_val} — expected 240/276/288")
                except ValueError:
                    pass

    # Rule 6: Indentation checks for bulleted content
    if _has_numbering(ppr) and style == STYLE_BODY:
        if ppr is not None:
            ind = ppr.find(wn('ind'))
            if ind is not None:
                left = ind.get(wattr('left'), '0')
                try:
                    left_val = int(left)
                    if left_val not in (0, 142, 284, 426, 720):
                        issues.append(f"Bullet indent left={left_val} — non-standard value")
                except ValueError:
                    pass

    return issues


def _get_run_text(run):
    """Get text from a single run."""
    texts = []
    for t in run.findall(wn('t')):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)


def _highlight_paragraph_yellow(para):
    """Add yellow highlight to all runs in a paragraph."""
    for r in para.findall(wn('r')):
        rpr = r.find(wn('rPr'))
        if rpr is None:
            rpr = etree.SubElement(r, wn('rPr'))
            # Insert rPr as first child of run
            r.insert(0, rpr)

        # Add highlight
        highlight = rpr.find(wn('highlight'))
        if highlight is None:
            highlight = etree.SubElement(rpr, wn('highlight'))
        highlight.set(wattr('val'), 'yellow')

    # Also add shading to paragraph for empty-ish lines
    ppr = para.find(wn('pPr'))
    if ppr is None:
        ppr = etree.SubElement(para, wn('pPr'))
        para.insert(0, ppr)

    shd = ppr.find(wn('shd'))
    if shd is None:
        shd = etree.SubElement(ppr, wn('shd'))
    shd.set(wattr('val'), 'clear')
    shd.set(wattr('color'), 'auto')
    shd.set(wattr('fill'), 'FFFF00')


# ═══════════════════════════════════════════════════════════════════════════
# CONVERTER — Rebuild content into the Crowe template
# ═══════════════════════════════════════════════════════════════════════════

def convert_document(input_path, output_path, title=None):
    """
    Convert user content into the Crowe branded template.

    Steps:
      1. Copy the original Crowe template
      2. Parse the user document for content
      3. Clear Section 2 paragraphs from the template
      4. Insert new formatted paragraphs
      5. Update cover page title if provided
      6. Save
    """
    # Parse user content first
    content_blocks = parse_user_document(input_path)

    if not content_blocks:
        logger.warning("No content found in the input document.")
        shutil.copy2(str(TEMPLATE_PATH), output_path)
        return

    # Work on the template
    tmp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(str(TEMPLATE_PATH), 'r') as z:
            z.extractall(tmp_dir)

        doc_xml_path = os.path.join(tmp_dir, 'word', 'document.xml')
        tree = etree.parse(doc_xml_path)
        root = tree.getroot()
        body = root.find(f'.//{wn("body")}')

        all_children = list(body)
        paragraphs = [c for c in all_children if c.tag == wn('p')]
        non_para_elements = [(all_children.index(c), c) for c in all_children if c.tag != wn('p')]

        # ── Identify section boundaries ──
        # Section 1 ends at the first <w:sectPr> inside a paragraph (P6)
        # Section 2 ends at the second <w:sectPr> inside a paragraph (P179)
        # Section 3 is everything after the second sectPr paragraph
        sec_breaks = []
        for i, para in enumerate(paragraphs):
            ppr = para.find(wn('pPr'))
            if ppr is not None:
                sectpr = ppr.find(wn('sectPr'))
                if sectpr is not None:
                    sec_breaks.append(i)

        if len(sec_breaks) >= 2:
            sec1_end_idx = sec_breaks[0]
            sec2_break_idx = sec_breaks[1]
        else:
            sec1_end_idx = SECTION1_END
            sec2_break_idx = SECTION3_START - 1

        # ── Keep Section 1 and Section 3 paragraphs ──
        section1_paras = paragraphs[:sec1_end_idx + 1]
        # The section 2 break paragraph contains <w:sectPr> that defines content section layout.
        # We must preserve it (with its sectPr) but can replace its text content.
        sec2_break_para = paragraphs[sec2_break_idx]
        section3_paras = paragraphs[sec2_break_idx + 1:]

        # ── Update cover title if provided ──
        if title:
            _update_cover_title(section1_paras, title)

        # ── Build new Section 2 content ──
        new_section2_paras = _build_section2_xml(content_blocks, tree)

        # ── Rebuild body ──
        # Clear all children from body
        for child in list(body):
            body.remove(child)

        # Re-add: Section 1
        for p in section1_paras:
            body.append(p)

        # Re-add: New Section 2 content
        for p in new_section2_paras:
            body.append(p)

        # Re-add: Section 2 break paragraph (preserves sectPr for content layout)
        # Clear its text runs but keep pPr with sectPr intact
        for r in list(sec2_break_para.findall(wn('r'))):
            sec2_break_para.remove(r)
        body.append(sec2_break_para)

        # Re-add: Section 3
        for p in section3_paras:
            body.append(p)

        # Re-add non-paragraph elements (like final sectPr)
        for orig_idx, elem in non_para_elements:
            if elem.tag == wn('sectPr'):
                # Final section properties go at the end of body
                body.append(elem)

        # ── Write back ──
        tree.write(doc_xml_path, xml_declaration=True, encoding='UTF-8',
                   standalone=True)

        _repack_docx(tmp_dir, output_path)
        logger.info("Converted document saved to: %s", output_path)

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _update_cover_title(section1_paras, title):
    """Update the title text on the cover page (P1)."""
    for para in section1_paras:
        ppr = para.find(wn('pPr'))
        style = _get_style(ppr)
        if style == STYLE_COVER_TITLE:
            text = _get_para_text(para)
            if text.strip():
                # Found the title paragraph
                # The template splits the title across multiple runs.
                # Keep the first run (preserves formatting), remove the rest.
                runs = para.findall(wn('r'))
                if runs:
                    # Set first run text to full new title
                    first_t = runs[0].find(wn('t'))
                    if first_t is not None:
                        first_t.text = title
                        first_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    # Remove all subsequent runs to avoid text duplication
                    for r in runs[1:]:
                        para.remove(r)
                return


def _build_section2_xml(content_blocks, tree):
    """
    Build the XML paragraphs for Section 2 from parsed content blocks.
    Returns a list of lxml elements.
    """
    paragraphs = []

    # Start with some empty spacing paragraphs (like original has P7-P14)
    for _ in range(4):
        paragraphs.append(_make_empty_para(STYLE_H1, spacing_before='0', spacing_after='240'))

    for block in content_blocks:
        if block.level == 'h1':
            # Add spacing before H1
            paragraphs.append(_make_empty_para(STYLE_H1, spacing_before='0', spacing_after='240'))
            paragraphs.append(_make_h1_para(block.text))
            paragraphs.append(_make_empty_para(STYLE_H2))

        elif block.level == 'h2':
            paragraphs.append(_make_h2_subtopic_para(block.text))

        elif block.level == 'h3':
            paragraphs.append(_make_h3_bold_subhead(block.text))

        elif block.level == 'bullet_bold':
            paragraphs.append(_make_bullet_bold_para(block.text))
            # Add description children as indented body
            for desc in block.children:
                paragraphs.append(_make_indented_body_para(desc))

        elif block.level == 'bullet':
            paragraphs.append(_make_bullet_body_para(block.text))

        elif block.level == 'table':
            # Add spacing before table
            paragraphs.append(_make_empty_para(STYLE_BODY))
            paragraphs.append(_make_table_xml(block))
            # Add spacing after table
            paragraphs.append(_make_empty_para(STYLE_BODY))

        elif block.level == 'body':
            paragraphs.append(_make_body_para(block.text))

        elif block.level == 'body_indent':
            paragraphs.append(_make_indented_body_para(block.text))

    # Add trailing empty paragraphs
    for _ in range(2):
        paragraphs.append(_make_empty_para(STYLE_H2))

    return paragraphs


# ─── Paragraph builders ─────────────────────────────────────────────────────

def _make_empty_para(style_id, spacing_before=None, spacing_after=None):
    """Create an empty spacing paragraph."""
    para = etree.Element(wn('p'))
    ppr = etree.SubElement(para, wn('pPr'))
    pstyle = etree.SubElement(ppr, wn('pStyle'))
    pstyle.set(wattr('val'), style_id)

    if spacing_before is not None or spacing_after is not None:
        sp = etree.SubElement(ppr, wn('spacing'))
        if spacing_before is not None:
            sp.set(wattr('before'), spacing_before)
        if spacing_after is not None:
            sp.set(wattr('after'), spacing_after)

    return para


def _make_h1_para(text):
    """
    HeadingStyle1-18pt — Major section heading.
    Style: HeadingStyle1-18pt (inherently 18pt bold from style definition)
    """
    para = etree.Element(wn('p'))
    ppr = etree.SubElement(para, wn('pPr'))
    pstyle = etree.SubElement(ppr, wn('pStyle'))
    pstyle.set(wattr('val'), STYLE_H1)

    sp = etree.SubElement(ppr, wn('spacing'))
    sp.set(wattr('before'), '0')
    sp.set(wattr('after'), '240')

    run = etree.SubElement(para, wn('r'))
    t = etree.SubElement(run, wn('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return para


def _make_h2_subtopic_para(text):
    """
    HeadingStyle2-14pt with run override to 12pt (sz=24).
    Used for subtopic headings like "Changes in ITR filing provisions".
    """
    para = etree.Element(wn('p'))
    ppr = etree.SubElement(para, wn('pPr'))
    pstyle = etree.SubElement(ppr, wn('pStyle'))
    pstyle.set(wattr('val'), STYLE_H2)

    sp = etree.SubElement(ppr, wn('spacing'))
    sp.set(wattr('before'), '0')
    sp.set(wattr('after'), '')
    sp.set(wattr('line'), '480')

    run = etree.SubElement(para, wn('r'))
    rpr = etree.SubElement(run, wn('rPr'))
    sz = etree.SubElement(rpr, wn('sz'))
    sz.set(wattr('val'), '24')
    szCs = etree.SubElement(rpr, wn('szCs'))
    szCs.set(wattr('val'), '24')

    t = etree.SubElement(run, wn('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return para


def _make_h3_bold_subhead(text):
    """
    BodyCopy-Arial10pt with bold run.
    Used for sub-sub-headings like "Background & Objective".
    """
    para = etree.Element(wn('p'))
    ppr = etree.SubElement(para, wn('pPr'))
    pstyle = etree.SubElement(ppr, wn('pStyle'))
    pstyle.set(wattr('val'), STYLE_BODY)

    run = etree.SubElement(para, wn('r'))
    rpr = etree.SubElement(run, wn('rPr'))
    b = etree.SubElement(rpr, wn('b'))
    bCs = etree.SubElement(rpr, wn('bCs'))
    bCs.set(wattr('val'), '0')

    t = etree.SubElement(run, wn('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return para


def _make_bullet_bold_para(text):
    """
    BodyCopy-Arial10pt + bullet numbering + bold.
    Used for bold bullet headings like "• Due date for filing of Return".
    """
    # Strip leading bullet characters if user had them
    text = _strip_bullet_chars(text)

    para = etree.Element(wn('p'))
    ppr = etree.SubElement(para, wn('pPr'))
    pstyle = etree.SubElement(ppr, wn('pStyle'))
    pstyle.set(wattr('val'), STYLE_BODY)

    # Numbering
    numpr = etree.SubElement(ppr, wn('numPr'))
    ilvl = etree.SubElement(numpr, wn('ilvl'))
    ilvl.set(wattr('val'), '0')
    numid = etree.SubElement(numpr, wn('numId'))
    numid.set(wattr('val'), '56')  # Using numId 56 from template

    # Spacing
    sp = etree.SubElement(ppr, wn('spacing'))
    sp.set(wattr('line'), '276')

    # Indent
    ind = etree.SubElement(ppr, wn('ind'))
    ind.set(wattr('left'), '426')
    ind.set(wattr('hanging'), '426')

    # Bold run
    run = etree.SubElement(para, wn('r'))
    rpr = etree.SubElement(run, wn('rPr'))
    b = etree.SubElement(rpr, wn('b'))

    t = etree.SubElement(run, wn('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return para


def _make_bullet_body_para(text):
    """
    BodyCopy-Arial10pt + bullet numbering, normal weight.
    Used for regular bullet list items.
    """
    text = _strip_bullet_chars(text)

    para = etree.Element(wn('p'))
    ppr = etree.SubElement(para, wn('pPr'))
    pstyle = etree.SubElement(ppr, wn('pStyle'))
    pstyle.set(wattr('val'), STYLE_BODY)

    # Numbering
    numpr = etree.SubElement(ppr, wn('numPr'))
    ilvl = etree.SubElement(numpr, wn('ilvl'))
    ilvl.set(wattr('val'), '0')
    numid = etree.SubElement(numpr, wn('numId'))
    numid.set(wattr('val'), '55')  # Using numId 55 from template

    # Indent
    ind = etree.SubElement(ppr, wn('ind'))
    ind.set(wattr('left'), '426')
    ind.set(wattr('hanging'), '426')

    # Spacing
    sp = etree.SubElement(ppr, wn('spacing'))
    sp.set(wattr('line'), '276')

    run = etree.SubElement(para, wn('r'))
    t = etree.SubElement(run, wn('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return para


def _make_indented_body_para(text):
    """
    BodyCopy-Arial10pt with left indent (no numbering).
    Used for description text under bold bullet headings.
    """
    para = etree.Element(wn('p'))
    ppr = etree.SubElement(para, wn('pPr'))
    pstyle = etree.SubElement(ppr, wn('pStyle'))
    pstyle.set(wattr('val'), STYLE_BODY)

    # Indent (matches template: left=426 for body under bullets, or 284)
    ind = etree.SubElement(ppr, wn('ind'))
    ind.set(wattr('left'), '426')

    # Spacing
    sp = etree.SubElement(ppr, wn('spacing'))
    sp.set(wattr('line'), '276')

    run = etree.SubElement(para, wn('r'))
    t = etree.SubElement(run, wn('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return para


def _make_body_para(text):
    """
    BodyCopy-Arial10pt, no indent, normal weight.
    Standard body paragraph.
    """
    para = etree.Element(wn('p'))
    ppr = etree.SubElement(para, wn('pPr'))
    pstyle = etree.SubElement(ppr, wn('pStyle'))
    pstyle.set(wattr('val'), STYLE_BODY)

    run = etree.SubElement(para, wn('r'))
    t = etree.SubElement(run, wn('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return para


def _make_table_xml(table_block):
    """
    Build a w:tbl element from a TableBlock with Crowe-branded shading.
    Header row gets amber fill, data rows get light cream fill, navy borders.
    """
    tbl = etree.Element(wn('tbl'))

    # ── Table properties ──
    tblPr = etree.SubElement(tbl, wn('tblPr'))
    tblW = etree.SubElement(tblPr, wn('tblW'))
    tblW.set(wattr('w'), '5000')
    tblW.set(wattr('type'), 'pct')
    jc = etree.SubElement(tblPr, wn('jc'))
    jc.set(wattr('val'), 'center')

    # Table-level borders
    tblBorders = etree.SubElement(tblPr, wn('tblBorders'))
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = etree.SubElement(tblBorders, wn(side))
        border.set(wattr('val'), 'single')
        border.set(wattr('sz'), '8')
        border.set(wattr('space'), '0')
        border.set(wattr('color'), TABLE_BORDER_COLOR)

    # Cell margins
    tblCellMar = etree.SubElement(tblPr, wn('tblCellMar'))
    cm_left = etree.SubElement(tblCellMar, wn('left'))
    cm_left.set(wattr('w'), '0')
    cm_left.set(wattr('type'), 'dxa')
    cm_right = etree.SubElement(tblCellMar, wn('right'))
    cm_right.set(wattr('w'), '0')
    cm_right.set(wattr('type'), 'dxa')

    # ── Grid columns ──
    num_cols = len(table_block.headers)
    if table_block.rows:
        num_cols = max(num_cols, max(len(r) for r in table_block.rows))

    tblGrid = etree.SubElement(tbl, wn('tblGrid'))
    if table_block.col_widths and len(table_block.col_widths) >= num_cols:
        for i in range(num_cols):
            gc = etree.SubElement(tblGrid, wn('gridCol'))
            gc.set(wattr('w'), str(table_block.col_widths[i]))
    else:
        # Distribute evenly across typical A4 content width
        total_w = 9164
        col_w = total_w // num_cols
        for _ in range(num_cols):
            gc = etree.SubElement(tblGrid, wn('gridCol'))
            gc.set(wattr('w'), str(col_w))

    # ── Header row ──
    _add_table_row(tbl, table_block.headers, num_cols, is_header=True)

    # ── Data rows ──
    for row_data in table_block.rows:
        _add_table_row(tbl, row_data, num_cols, is_header=False)

    return tbl


def _add_table_row(tbl, cells, num_cols, is_header=False):
    """Add a table row (w:tr) with properly styled cells."""
    tr = etree.SubElement(tbl, wn('tr'))

    for col_idx in range(num_cols):
        cell_text = cells[col_idx] if col_idx < len(cells) else ''
        tc = _make_table_cell(cell_text, is_header)
        tr.append(tc)


def _make_table_cell(text, is_header):
    """Create a single table cell (w:tc) with Crowe shading."""
    tc = etree.Element(wn('tc'))

    # ── Cell properties ──
    tcPr = etree.SubElement(tc, wn('tcPr'))

    # Cell borders
    tcBorders = etree.SubElement(tcPr, wn('tcBorders'))
    for side in ('top', 'left', 'bottom', 'right'):
        border = etree.SubElement(tcBorders, wn(side))
        border.set(wattr('val'), 'single')
        border.set(wattr('sz'), '8')
        border.set(wattr('space'), '0')
        border.set(wattr('color'), TABLE_BORDER_COLOR)

    # Cell shading — amber for header, cream for data
    shd = etree.SubElement(tcPr, wn('shd'))
    shd.set(wattr('val'), 'clear')
    shd.set(wattr('color'), 'auto')
    shd.set(wattr('fill'), TABLE_HEADER_FILL if is_header else TABLE_BODY_FILL)

    # Cell margins
    tcMar = etree.SubElement(tcPr, wn('tcMar'))
    mar_top = etree.SubElement(tcMar, wn('top'))
    mar_top.set(wattr('w'), '15')
    mar_top.set(wattr('type'), 'dxa')
    mar_left = etree.SubElement(tcMar, wn('left'))
    mar_left.set(wattr('w'), '97')
    mar_left.set(wattr('type'), 'dxa')
    mar_bottom = etree.SubElement(tcMar, wn('bottom'))
    mar_bottom.set(wattr('w'), '0')
    mar_bottom.set(wattr('type'), 'dxa')
    mar_right = etree.SubElement(tcMar, wn('right'))
    mar_right.set(wattr('w'), '97')
    mar_right.set(wattr('type'), 'dxa')

    # Vertical alignment
    vAlign = etree.SubElement(tcPr, wn('vAlign'))
    vAlign.set(wattr('val'), 'bottom' if is_header else 'center')

    # ── Cell paragraph ──
    # If text has newlines (multiple paragraphs in source cell), create multiple <w:p>
    lines = text.split('\n') if text else ['']
    for line in lines:
        para = etree.SubElement(tc, wn('p'))
        ppr = etree.SubElement(para, wn('pPr'))
        pstyle = etree.SubElement(ppr, wn('pStyle'))
        pstyle.set(wattr('val'), STYLE_H2)
        sp = etree.SubElement(ppr, wn('spacing'))
        sp.set(wattr('line'), '276')
        sp.set(wattr('lineRule'), 'auto')

        run = etree.SubElement(para, wn('r'))
        rpr = etree.SubElement(run, wn('rPr'))

        if is_header:
            b = etree.SubElement(rpr, wn('b'))
        else:
            b = etree.SubElement(rpr, wn('b'))
            b.set(wattr('val'), '0')

        sz = etree.SubElement(rpr, wn('sz'))
        sz.set(wattr('val'), '20')
        szCs = etree.SubElement(rpr, wn('szCs'))
        szCs.set(wattr('val'), '20')

        t = etree.SubElement(run, wn('t'))
        t.text = line
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return tc


def _strip_bullet_chars(text):
    """Remove leading bullet characters and whitespace."""
    text = text.lstrip()
    bullet_chars = ['•', '●', '▪', '▸', '►', '‣', '-', '–', '—', '◦', '○']
    for ch in bullet_chars:
        if text.startswith(ch):
            text = text[len(ch):].lstrip()
            break
    return text


# ═══════════════════════════════════════════════════════════════════════════
# UTILITIES
# ═══════════════════════════════════════════════════════════════════════════

def _repack_docx(unpacked_dir, output_path):
    """Repack an unpacked directory into a .docx file."""
    # Remove existing output if it exists
    if os.path.exists(output_path):
        os.remove(output_path)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root_dir, dirs, files in os.walk(unpacked_dir):
            for f in files:
                file_path = os.path.join(root_dir, f)
                arcname = os.path.relpath(file_path, unpacked_dir)
                zf.write(file_path, arcname)


# ═══════════════════════════════════════════════════════════════════════════
# CLI ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description='Crowe Document Formatter — Audit & Convert Tool',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python crowe_formatter.py audit input.docx audited_output.docx
  python crowe_formatter.py convert input.docx converted_output.docx --title "My Report"
        """
    )
    parser.add_argument('mode', choices=['audit', 'convert'],
                        help='Operation mode: audit or convert')
    parser.add_argument('input', help='Path to input .docx file')
    parser.add_argument('output', help='Path to output .docx file')
    parser.add_argument('--title', default=None,
                        help='(convert only) Title for the cover page')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}")
        sys.exit(1)

    if not TEMPLATE_PATH.exists():
        print(f"Error: Template not found at {TEMPLATE_PATH}")
        print("  Place the Crowe template at: template_assets/template.docx")
        sys.exit(1)

    if args.mode == 'audit':
        print(f"🔍 Auditing: {args.input}")
        result = audit_document(args.input, args.output)
        print(result.summary())
        print(f"📄 Audited document saved to: {args.output}")

    elif args.mode == 'convert':
        print(f"🔄 Converting: {args.input}")
        convert_document(args.input, args.output, title=args.title)


if __name__ == '__main__':
    main()
