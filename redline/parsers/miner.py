"""Historical playbook miner: pull past markups out of .docx and .pdf contracts.

DOCX: reads w:del / w:ins / comment ranges straight from the XML. python-docx's paragraph.text
skips deleted text, so a miner built on it would silently lose every deletion.

PDF: two kinds of marked-up PDF turn up in a contract archive.
  1. Annotated PDFs (Acrobat, Bluebeam): StrikeOut / Caret / Text annotations with popup notes.
  2. Printed redlines (a Word compare or tracked changes exported to PDF): no annotations at all,
     only coloured text with a line through it (deleted) or under it (inserted).
Both are read here.
"""
from __future__ import annotations

import re
from pathlib import Path

import pymupdf
from docx import Document
from docx.oxml.ns import qn

from redline.core.models import HistoricalEdit
from redline.core.ooxml import W_DEL, W_DELTEXT, W_INS, W_T, view_text
from redline.parsers.chunker import HEADING_RE, is_auto_numbered, is_heading

_C_START = qn("w:commentRangeStart")


def _text_of(el, tag) -> str:
    return "".join(t.text or "" for t in el.iter(tag))


def _comments(doc) -> dict[str, str]:
    try:
        return {str(c.comment_id): c.text for c in doc.comments}
    except Exception:
        return {}


def mine_docx(path: str | Path) -> list[HistoricalEdit]:
    path = Path(path)
    doc = Document(path)
    comments = _comments(doc)
    out: list[HistoricalEdit] = []
    clause_id, heading, auto_n = "", "", 0
    for par in doc.paragraphs:
        p = par._p
        original = view_text(p, "reject")
        if is_heading(par):
            if is_auto_numbered(par):
                auto_n += 1
                clause_id, heading = str(auto_n), original.strip()
            else:
                m = HEADING_RE.match(original.strip())
                clause_id, heading = (m.group(1), m.group(2)) if m else ("", original.strip())
            continue
        children = list(p)
        open_comment = ""
        i = 0
        while i < len(children):
            el = children[i]
            if el.tag == _C_START:
                open_comment = comments.get(el.get(qn("w:id")), "")
            elif el.tag in (W_DEL, W_INS):
                deleted = _text_of(el, W_DELTEXT) if el.tag == W_DEL else ""
                inserted = _text_of(el, W_T) if el.tag == W_INS else ""
                author = el.get(qn("w:author"), "")
                if el.tag == W_DEL and i + 1 < len(children) and children[i + 1].tag == W_INS:
                    inserted = _text_of(children[i + 1], W_T)
                    i += 1
                kind = "replace" if deleted and inserted else ("delete" if deleted else "insert")
                out.append(HistoricalEdit(source=path.name, clause_id=clause_id, heading=heading, kind=kind,
                                          deleted=deleted, inserted=inserted, comment=open_comment,
                                          context=original, author=author))
                open_comment = ""
            i += 1
    return out


# ------------------------------------------------------------------------------------------ PDF

def _page_heading_index(doc) -> list[tuple[int, float, str, str]]:
    """(page, y, clause_id, heading) for every numbered heading line, in reading order."""
    idx = []
    for page in doc:
        for b in page.get_text("dict")["blocks"]:
            for line in b.get("lines", []):
                text = "".join(s["text"] for s in line["spans"]).strip()
                bold = all(s["flags"] & 16 for s in line["spans"] if s["text"].strip())
                m = HEADING_RE.match(text)
                if m and bold:
                    idx.append((page.number, line["bbox"][1], m.group(1), m.group(2)))
    return idx


def _clause_at(index, page_no: int, y: float) -> tuple[str, str]:
    best = ("", "")
    for (pg, hy, cid, head) in index:
        if (pg, hy) <= (page_no, y):
            best = (cid, head)
    return best


def _annotation_edits(doc, path: Path, index) -> list[HistoricalEdit]:
    out = []
    for page in doc:
        for a in page.annots() or []:
            kind = a.type[1]
            if kind not in ("StrikeOut", "Caret", "Text", "FreeText"):
                continue
            content = (a.info.get("content") or "").strip()
            deleted = ""
            if kind == "StrikeOut":
                verts = a.vertices or []
                quads = [pymupdf.Quad(verts[i:i + 4]) for i in range(0, len(verts), 4)]
                deleted = " ".join(page.get_textbox(q.rect).strip() for q in quads)
            inserted, comment = "", content
            m = re.match(r"(?is)replace with:\s*(.+?)(?:\n(.*))?$", content)
            if m:
                inserted, comment = m.group(1).strip(), (m.group(2) or "").strip()
            cid, head = _clause_at(index, page.number, a.rect.y0)
            out.append(HistoricalEdit(source=path.name, clause_id=cid, heading=head,
                                      kind="replace" if deleted and inserted else ("delete" if deleted else "comment"),
                                      deleted=deleted, inserted=inserted, comment=comment,
                                      author=a.info.get("title", "")))
    return out


def _printed_redline_edits(doc, path: Path, index) -> list[HistoricalEdit]:
    """Coloured characters with a line through the middle are deleted; with a line under them, inserted."""
    out: list[HistoricalEdit] = []
    marks: list[tuple[int, float, str, str]] = []     # (page, y, 'del'|'ins', char)
    for page in doc:
        lines = []
        for d in page.get_drawings():
            for it in d["items"]:
                if it[0] == "l" and abs(it[1].y - it[2].y) < 0.6 and (d.get("color") or (0, 0, 0)) != (0, 0, 0):
                    lines.append((min(it[1].x, it[2].x), max(it[1].x, it[2].x), it[1].y))
        for b in page.get_text("rawdict")["blocks"]:
            for line in b.get("lines", []):
                for s in line["spans"]:
                    if s["color"] == 0:
                        continue
                    for ch in s["chars"]:
                        x0, y0, x1, y1 = ch["bbox"]
                        cx, h = (x0 + x1) / 2, (y1 - y0) or 1
                        kind = None
                        for (lx0, lx1, ly) in lines:
                            if lx0 - 0.5 <= cx <= lx1 + 0.5 and y0 <= ly <= y1 + 2:
                                kind = "del" if (ly - y0) / h < 0.72 else "ins"
                                break
                        if kind:
                            marks.append((page.number, y0, kind, ch["c"]))
    # group consecutive characters of the same kind; a del run followed by an ins run is a replace
    runs: list[list] = []          # [page, y, kind, text, clause]
    for (pg, y, kind, ch) in marks:
        clause = _clause_at(index, pg, y)
        if runs and runs[-1][2] == kind and runs[-1][4] == clause:
            if (pg, round(y)) != runs[-1][5] and not runs[-1][3].endswith((" ", "-")):
                runs[-1][3] += " "      # the text wrapped to a new line
            runs[-1][3] += ch
            runs[-1][5] = (pg, round(y))
        else:
            runs.append([pg, y, kind, ch, clause, (pg, round(y))])
    i = 0
    while i < len(runs):
        pg, y, kind, text, _clause, _ = runs[i]
        deleted, inserted = (text, "") if kind == "del" else ("", text)
        if kind == "del" and i + 1 < len(runs) and runs[i + 1][2] == "ins" and runs[i + 1][4] == _clause:
            inserted = runs[i + 1][3]
            i += 1
        cid, head = _clause_at(index, pg, y)
        out.append(HistoricalEdit(source=path.name, clause_id=cid, heading=head,
                                  kind="replace" if deleted and inserted else ("delete" if deleted else "insert"),
                                  deleted=deleted.strip(), inserted=inserted.strip()))
        i += 1
    return out


def mine_pdf(path: str | Path) -> list[HistoricalEdit]:
    path = Path(path)
    doc = pymupdf.open(path)
    index = _page_heading_index(doc)
    edits = _annotation_edits(doc, path, index)
    return edits if edits else _printed_redline_edits(doc, path, index)


def mine(path: str | Path) -> list[HistoricalEdit]:
    return mine_pdf(path) if str(path).lower().endswith(".pdf") else mine_docx(path)
