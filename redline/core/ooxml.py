"""Low-level WordprocessingML helpers: run text, run splitting, accept/reject views.

Paragraph text is read from the direct w:r children of w:p only. Runs inside w:del and w:ins are
not direct children, so the "original text" offsets of a paragraph stay stable while tracked
changes are being added to it.
"""
from __future__ import annotations

import copy

from docx.oxml.ns import qn

W_R, W_T, W_TAB, W_BR, W_CR = qn("w:r"), qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")
W_RPR, W_DEL, W_INS, W_DELTEXT = qn("w:rPr"), qn("w:del"), qn("w:ins"), qn("w:delText")
W_HYPERLINK = qn("w:hyperlink")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_CONTENT = {W_T: None, W_TAB: "\t", W_BR: "\n", W_CR: "\n"}


def run_text(r) -> str:
    out = []
    for child in r:
        if child.tag == W_T:
            out.append(child.text or "")
        elif child.tag in (W_TAB, W_BR, W_CR):
            out.append(_CONTENT[child.tag])
    return "".join(out)


def direct_runs(p) -> list:
    """Runs whose text counts as the paragraph's current, untracked text."""
    runs = []
    for child in p:
        if child.tag == W_R:
            runs.append(child)
        elif child.tag == W_HYPERLINK:
            runs.extend(c for c in child if c.tag == W_R)
    return runs


def paragraph_text(p) -> str:
    return "".join(run_text(r) for r in direct_runs(p))


def _content_children(r):
    return [c for c in r if c.tag in _CONTENT]


def explode_run(r) -> list:
    """Split a run holding several content children (text, tab, break) into one run per child."""
    content = _content_children(r)
    if len(content) <= 1:
        return [r]
    parent = r.getparent()
    idx = parent.index(r)
    new_runs = []
    for c in content:
        nr = copy.deepcopy(r)
        for cc in list(nr):
            if cc.tag != W_RPR:
                nr.remove(cc)
        nr.append(copy.deepcopy(c))
        new_runs.append(nr)
    parent.remove(r)
    for i, nr in enumerate(new_runs):
        parent.insert(idx + i, nr)
    return new_runs


def split_run(r, k: int):
    """Split a single-w:t run at character k. Returns (left, right); either may be r itself."""
    t = r.find(W_T)
    text = t.text or ""
    if k <= 0 or k >= len(text):
        raise ValueError("split point must be inside the run")
    right = copy.deepcopy(r)
    t.text = text[:k]
    t.set(XML_SPACE, "preserve")
    rt = right.find(W_T)
    rt.text = text[k:]
    rt.set(XML_SPACE, "preserve")
    r.addnext(right)
    return r, right


def runs_covering(p, start: int, end: int) -> list:
    """Return the direct runs that exactly cover [start, end), splitting runs at the edges."""
    # explode multi-content runs that touch the range, so every run has one content child
    pos = 0
    for r in list(direct_runs(p)):
        txt = run_text(r)
        if pos < end and pos + len(txt) > start and len(_content_children(r)) > 1:
            explode_run(r)
        pos += len(txt)

    for boundary in (start, end):
        pos = 0
        for r in direct_runs(p):
            txt = run_text(r)
            if pos < boundary < pos + len(txt):
                split_run(r, boundary - pos)
                break
            pos += len(txt)

    covered, pos = [], 0
    for r in direct_runs(p):
        txt = run_text(r)
        if pos >= start and pos + len(txt) <= end and txt:
            covered.append(r)
        pos += len(txt)
    if "".join(run_text(r) for r in covered) != paragraph_text(p)[start:end]:
        raise ValueError("runs do not cover the requested range")
    return covered


def view_text(p, mode: str) -> str:
    """Text of a paragraph after accepting ('accept') or rejecting ('reject') all tracked changes."""
    out = []

    def walk(el):
        for child in el:
            if child.tag == W_R:
                for c in child:
                    if c.tag == W_T:
                        out.append(c.text or "")
                    elif c.tag == W_DELTEXT:
                        out.append(c.text or "")
                    elif c.tag in (W_TAB, W_BR, W_CR):
                        out.append(_CONTENT[c.tag])
            elif child.tag == W_DEL:
                if mode == "reject":
                    walk(child)
            elif child.tag == W_INS:
                if mode == "accept":
                    walk(child)
            elif child.tag == W_HYPERLINK:
                walk(child)

    walk(p)
    return "".join(out)
