"""Unified clause chunker for .docx and .pdf.

DOCX: headings are bold "N. Heading" paragraphs, or paragraphs carrying Word auto-numbering
(w:numPr or a "List Number" style), whose number is NOT part of the paragraph text.
PDF: strips lines that repeat at the top or bottom of most pages (headers, footers, page numbers),
then merges body lines into clauses, so a clause split across a page break comes back whole.
Every clause keeps spans back to the source (paragraph offsets or PDF line boxes) for the writers.
"""
from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

import pymupdf
from docx import Document
from docx.oxml.ns import qn

from redline.core.models import Clause, Span
from redline.core.ooxml import paragraph_text

HEADING_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+([A-Z][^.]{1,80})$")


def is_auto_numbered(par) -> bool:
    ppr = par._p.pPr
    if ppr is not None and ppr.find(qn("w:numPr")) is not None:
        return True
    return (par.style is not None and par.style.name.startswith("List Number"))


def is_heading(par) -> bool:
    text = paragraph_text(par._p).strip()
    if not text:
        return False
    runs = [r for r in par.runs if r.text.strip()]
    bold = bool(runs) and all(r.bold for r in runs)
    if is_auto_numbered(par):
        return bold and len(text) < 90
    return bold and bool(HEADING_RE.match(text))


def chunk_docx(path: str | Path) -> list[Clause]:
    doc = Document(path)
    clauses: list[Clause] = []
    current: Clause | None = None
    auto_n = 0
    for pi, par in enumerate(doc.paragraphs):
        text = paragraph_text(par._p)
        if is_heading(par):
            if is_auto_numbered(par):
                auto_n += 1
                cid, head = str(auto_n), text.strip()
            else:
                m = HEADING_RE.match(text.strip())
                cid, head = m.group(1), m.group(2)
            current = Clause(id=cid, heading=head, text="")
            clauses.append(current)
            continue
        if current is None or not text.strip():
            continue
        sep = "\n" if current.text else ""
        start = len(current.text) + len(sep)
        current.text += sep + text
        current.spans.append(Span(paragraph=pi, start=start, end=start + len(text)))
    return clauses


def _norm(s: str) -> str:
    return re.sub(r"\d+", "#", s.strip().lower())


def _lines(doc) -> list[tuple[int, tuple, str, bool]]:
    out = []
    for page in doc:
        for b in page.get_text("dict", sort=True)["blocks"]:
            for line in b.get("lines", []):
                text = "".join(s["text"] for s in line["spans"])
                if not text.strip():
                    continue
                bold = all(s["flags"] & 16 for s in line["spans"] if s["text"].strip())
                out.append((page.number, tuple(line["bbox"]), text, bold))
    return out


def repeating_lines(doc, lines, band: float = 0.1, share: float = 0.5) -> set[tuple[int, tuple]]:
    """Lines in the top/bottom band whose text (digits ignored) repeats on most pages."""
    if doc.page_count < 2:
        return set()
    h = doc[0].rect.height
    edge = [(pg, bb, t) for (pg, bb, t, _) in lines if bb[1] < h * band or bb[3] > h * (1 - band)]
    counts = Counter(norm for (_pg, norm) in {(pg, _norm(t)) for (pg, bb, t) in edge})
    repeated = {k for k, c in counts.items() if c >= max(2, share * doc.page_count)}
    return {(pg, bb) for (pg, bb, t) in edge if _norm(t) in repeated}


def chunk_pdf(path: str | Path) -> tuple[list[Clause], int]:
    """Returns clauses and the number of header/footer lines removed."""
    doc = pymupdf.open(path)
    lines = _lines(doc)
    drop = repeating_lines(doc, lines)
    clauses: list[Clause] = []
    current: Clause | None = None
    for (pg, bb, text, bold) in lines:
        if (pg, bb) in drop:
            continue
        m = HEADING_RE.match(text.strip())
        if m and bold:
            current = Clause(id=m.group(1), heading=m.group(2).strip(), text="")
            clauses.append(current)
            continue
        if current is None:
            continue
        t = text.strip()
        sep = "" if (not current.text or current.text.endswith("-")) else " "
        start = len(current.text) + len(sep)
        current.text += sep + t
        current.spans.append(Span(page=pg, start=start, end=start + len(t), bbox=bb))
    return clauses, len(drop)


def chunk(path: str | Path) -> list[Clause]:
    return chunk_pdf(path)[0] if str(path).lower().endswith(".pdf") else chunk_docx(path)
