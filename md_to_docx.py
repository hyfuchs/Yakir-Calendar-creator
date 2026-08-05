#!/usr/bin/env python3
"""
md_to_docx.py — convert the modiin_weeks column guide (Markdown) to a Word .docx
for email distribution. Purpose-built for this guide's structure (H1, blockquote,
H2 language sections, bold paragraphs, bullet lists, GFM tables, --- rules).

Handles Hebrew RTL: paragraphs/cells that contain Hebrew get bidi + right-align,
and every run sets BOTH w:sz and w:szCs (else Word renders Hebrew at 11pt).

Usage:
    python3 md_to_docx.py --in modiin_weeks_column_guide.md          # writes .docx next to input
    python3 md_to_docx.py --in guide.md --out guide.docx
"""
import argparse
import os
import re
import sys

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    sys.exit("python-docx not installed. Run: pip install python-docx")

FONT = "Arial"
HEB = re.compile(r"[֐-׿]")


def has_heb(s):
    return bool(HEB.search(s or ""))


def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def _bidi(para):
    """Mark a paragraph RTL (bidi before jc/spacing in schema order)."""
    pPr = para._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is not None:
        return
    bidi = _el("w:bidi", **{"w:val": "1"})
    anchor = None
    for child in pPr:
        if child.tag in (qn("w:spacing"), qn("w:ind"), qn("w:jc")):
            anchor = child
            break
    (anchor.addprevious(bidi) if anchor is not None else pPr.insert(0, bidi))


# Inline **bold** parser → list of (text, bold) segments.
def segments(text):
    out, i = [], 0
    for m in re.finditer(r"\*\*(.+?)\*\*", text):
        if m.start() > i:
            out.append((text[i:m.start()], False))
        out.append((m.group(1), True))
        i = m.end()
    if i < len(text):
        out.append((text[i:], False))
    return out or [(text, False)]


def add_run(para, text, size, bold=False, color=None):
    r = para.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    rpr = r._r.get_or_add_rPr()
    rpr.append(_el("w:rFonts", **{"w:ascii": FONT, "w:hAnsi": FONT, "w:cs": FONT}))
    rpr.append(_el("w:sz", **{"w:val": str(int(size * 2))}))
    rpr.append(_el("w:szCs", **{"w:val": str(int(size * 2))}))   # complex-script size
    if bold:
        rpr.append(_el("w:bCs", **{"w:val": "1"}))
    if color:
        rpr.append(_el("w:color", **{"w:val": color}))
    return r


def write_para(doc, text, size=11, bold_all=False, rtl=None, space_after=4, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    rtl = has_heb(text) if rtl is None else rtl
    if rtl:
        _bidi(p)
        p.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.LEFT
    for seg, b in segments(text):
        add_run(p, seg, size, bold=b or bold_all)
    return p


def add_table(doc, rows, rtl):
    ncols = len(rows[0])
    t = doc.add_table(rows=len(rows), cols=ncols)
    t.style = "Table Grid"
    if rtl:
        t.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # bidiVisual so column order mirrors for RTL
        t._tbl.tblPr.append(_el("w:bidiVisual"))
    for ri, row in enumerate(rows):
        for ci, cell in enumerate(row):
            c = t.rows[ri].cells[ci]
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            cell_rtl = has_heb(cell)
            if cell_rtl:
                _bidi(p)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for seg, b in segments(cell):
                add_run(p, seg, 10, bold=(b or ri == 0))
    return t


def parse_and_build(md_text, doc):
    lines = md_text.splitlines()
    i = 0
    pending_table = []

    def flush_table():
        nonlocal pending_table
        if pending_table:
            rtl = any(has_heb(c) for c in pending_table[0])
            # rtl decided per-table by its content
            rtl = any(has_heb(c) for row in pending_table for c in row)
            add_table(doc, pending_table, rtl)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
            pending_table = []

    for raw in lines:
        line = raw.rstrip("\n")
        s = line.strip()

        # table row?
        if s.startswith("|") and s.endswith("|"):
            cells = [c.strip() for c in s.strip("|").split("|")]
            if all(set(c) <= set("-: ") for c in cells):   # separator row
                continue
            pending_table.append(cells)
            continue
        else:
            flush_table()

        if not s:
            continue
        if s.startswith("# "):
            write_para(doc, s[2:], size=18, bold_all=True, space_after=8)
        elif s.startswith("## "):
            write_para(doc, s[3:], size=15, bold_all=True, space_after=6)
        elif s.startswith("### "):
            write_para(doc, s[4:], size=13, bold_all=True, space_after=4)
        elif s == "---":
            hr = doc.add_paragraph()
            pPr = hr._p.get_or_add_pPr()
            pbdr = _el("w:pBdr")
            pbdr.append(_el("w:bottom", **{"w:val": "single", "w:sz": "6",
                                           "w:space": "1", "w:color": "999999"}))
            pPr.append(pbdr)
        elif s.startswith(">"):
            write_para(doc, s.lstrip("> ").strip(), size=11, space_after=2)
        elif s.startswith("- "):
            p = write_para(doc, "•  " + s[2:], size=11, space_after=2)
        else:
            write_para(doc, s, size=11)

    flush_table()


def main():
    ap = argparse.ArgumentParser(description="Convert the column-guide Markdown to .docx")
    ap.add_argument("--in", dest="inp", required=True, help="input .md")
    ap.add_argument("--out", default=None, help="output .docx (default: next to input)")
    args = ap.parse_args()

    if not os.path.exists(args.inp):
        sys.exit(f"input not found: {args.inp}")
    out = args.out or os.path.splitext(os.path.abspath(args.inp))[0] + ".docx"

    with open(args.inp, encoding="utf-8") as f:
        md = f.read()

    doc = Document()
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Pt(54)   # 0.75"
        sec.top_margin = sec.bottom_margin = Pt(54)
    parse_and_build(md, doc)
    doc.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
