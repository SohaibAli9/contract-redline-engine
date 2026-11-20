"""Write strikethroughs and popup notes onto a .pdf, placed from character boxes.

Each edit is mapped back through the chunker's spans to the exact PDF lines it came from, and the
strikeout quads are built from those lines' own character boxes, so a phrase that wraps across
lines or across a page break gets one quad per line on the right pages. After writing, the text
under every strikeout is read back and compared with the text the edit meant to strike.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

import pymupdf

from redline.core.models import Clause, VerifiedEdit


@dataclass
class PdfResult:
    written: int
    readback_ok: int
    log: list[str]


def _line_chars(page, bbox) -> list[dict]:
    for b in page.get_text("rawdict", sort=True)["blocks"]:
        for line in b.get("lines", []):
            if all(abs(a - c) < 0.01 for a, c in zip(line["bbox"], bbox)):
                chars = [ch for s in line["spans"] for ch in s["chars"]]
                # the chunker stored the stripped line text
                while chars and chars[0]["c"].isspace():
                    chars.pop(0)
                while chars and chars[-1]["c"].isspace():
                    chars.pop()
                return chars
    return []


def _text_under(page, rect) -> str:
    """Characters whose centre lies inside rect (stricter than get_textbox, which takes any overlap)."""
    out = []
    for b in page.get_text("rawdict", sort=True)["blocks"]:
        for line in b.get("lines", []):
            for s in line["spans"]:
                for ch in s["chars"]:
                    x0, y0, x1, y1 = ch["bbox"]
                    if rect.contains(pymupdf.Point((x0 + x1) / 2, (y0 + y1) / 2)):
                        out.append(ch["c"])
    return "".join(out)


def _ws(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def apply_edits(src: str, dst: str, clauses: list[Clause], edits: list[VerifiedEdit],
                author: str = "Redline Engine") -> PdfResult:
    doc = pymupdf.open(src)
    by_id = {c.id: c for c in clauses}
    log: list[str] = []
    written = readback_ok = 0
    for e in edits:
        c = by_id[e.clause_id]
        per_page: dict[int, list] = {}
        insert_point = None
        for sp in c.spans:
            if e.start == e.end:
                if sp.start <= e.start <= sp.end:
                    chars = _line_chars(doc[sp.page], sp.bbox)
                    k = min(e.start - sp.start, len(chars)) - 1
                    x = chars[k]["bbox"][2] if chars and k >= 0 else sp.bbox[2]
                    insert_point = (sp.page, pymupdf.Point(x, sp.bbox[1]))
                continue
            a, b = max(e.start, sp.start), min(e.end, sp.end)
            if a >= b:
                continue
            chars = _line_chars(doc[sp.page], sp.bbox)
            seg = chars[a - sp.start:b - sp.start]
            if not seg:
                continue
            r = pymupdf.Rect(seg[0]["bbox"])
            for ch in seg[1:]:
                r |= pymupdf.Rect(ch["bbox"])
            per_page.setdefault(sp.page, []).append(r.quad)

        note = (f"Replace with: {e.replacement_text}" if e.end > e.start else f"Insert: {e.replacement_text}")
        if e.comment:
            note += f"\n{e.comment}"
        if insert_point is not None:
            pg, pt = insert_point
            page = doc[pg]                       # keep the page alive while the annot is edited
            a = page.add_caret_annot(pt)
            a.set_colors(stroke=(0.85, 0.0, 0.0))
            a.set_info(content=note, title=author)
            a.update()
            written += 1
            readback_ok += 1
            continue
        if not per_page:
            log.append(f"clause {e.clause_id}: could not place '{e.original_text[:40]}'")
            continue
        got = []
        for i, (pg, quads) in enumerate(sorted(per_page.items())):
            page = doc[pg]
            a = page.add_strikeout_annot(quads)
            a.set_colors(stroke=(0.85, 0.0, 0.0))
            a.set_info(content=note if i == 0 else "(continued from previous page)", title=author)
            a.update()
            got += [_text_under(page, q.rect) for q in quads]
        written += 1
        if _ws(" ".join(got)) == _ws(e.original_text):
            readback_ok += 1
        else:
            log.append(f"clause {e.clause_id}: readback '{_ws(' '.join(got))[:60]}' != '{_ws(e.original_text)[:60]}'")
    doc.save(dst, garbage=3, deflate=True)
    return PdfResult(written=written, readback_ok=readback_ok, log=log)
