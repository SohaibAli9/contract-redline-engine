"""Write native Word Track Changes (w:del / w:ins) and anchored margin comments into a .docx."""
from __future__ import annotations

import copy
from dataclasses import dataclass
from datetime import datetime, timezone

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from redline.core.ooxml import (
    W_DELTEXT, W_RPR, W_T, XML_SPACE, direct_runs, paragraph_text, runs_covering,
)


@dataclass
class DocxEdit:
    paragraph: int          # index into document.paragraphs (body paragraphs)
    start: int              # offsets into the paragraph's current text
    end: int
    replacement: str
    comment: str = ""


def _next_revision_id(body) -> int:
    ids = [int(el.get(qn("w:id"))) for el in body.iter(qn("w:ins"), qn("w:del"))
           if (el.get(qn("w:id")) or "").isdigit()]
    return max(ids, default=0) + 1


def _rev(tag: str, rid: int, author: str, date: str):
    el = OxmlElement(tag)
    el.set(qn("w:id"), str(rid))
    el.set(qn("w:author"), author)
    el.set(qn("w:date"), date)
    return el


def _text_run(text: str, rpr_source=None):
    r = OxmlElement("w:r")
    if rpr_source is not None and rpr_source.find(W_RPR) is not None:
        r.append(copy.deepcopy(rpr_source.find(W_RPR)))
    t = OxmlElement("w:t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    return r


def _to_deleted(r):
    for t in r.findall(W_T):
        t.tag = W_DELTEXT
        t.set(XML_SPACE, "preserve")
    for it in r.findall(qn("w:instrText")):
        it.tag = qn("w:delInstrText")


def apply_edits(src: str, dst: str, edits: list[DocxEdit], author: str = "Redline Engine",
                date: str | None = None) -> list[str]:
    """Apply edits and save to dst. Returns a log line for every edit that was skipped."""
    date = date or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    doc = Document(src)
    body = doc.element.body
    paragraphs = doc.paragraphs
    rid = _next_revision_id(body)
    log: list[str] = []

    by_para: dict[int, list[DocxEdit]] = {}
    for e in edits:
        by_para.setdefault(e.paragraph, []).append(e)

    for pi, plist in by_para.items():
        p = paragraphs[pi]._p
        original = paragraph_text(p)
        plist.sort(key=lambda e: e.start, reverse=True)   # right to left keeps offsets valid
        last_start = len(original) + 1
        for e in plist:
            if not (0 <= e.start <= e.end <= len(original)) or e.end > last_start:
                log.append(f"p{pi} [{e.start}:{e.end}] skipped: out of range or overlaps another edit")
                continue
            last_start = e.start
            try:
                targets = runs_covering(p, e.start, e.end) if e.end > e.start else []
            except ValueError as exc:
                log.append(f"p{pi} [{e.start}:{e.end}] skipped: {exc}")
                continue
            if targets and len({t.getparent() for t in targets}) > 1:
                log.append(f"p{pi} [{e.start}:{e.end}] skipped: range crosses a hyperlink boundary")
                continue

            # anchor: where the change goes, and which run lends its formatting
            if targets:
                anchor_before, style_src = targets[0], targets[0]
            else:
                pos, anchor_before, style_src = 0, None, None
                for r in direct_runs(p):
                    if pos >= e.start:
                        anchor_before = r
                        break
                    style_src = r
                    pos += len("".join(t.text or "" for t in r.findall(W_T)))

            new_nodes = []
            if targets:
                d = _rev("w:del", rid, author, date); rid += 1
                targets[0].addprevious(d)
                for r in targets:
                    d.append(r)          # moves the run inside w:del
                    _to_deleted(r)
                new_nodes.append(d)
            if e.replacement:
                ins = _rev("w:ins", rid, author, date); rid += 1
                ins.append(_text_run(e.replacement, style_src))
                new_nodes.append(ins)

            # place nodes: deletion already sits where the targets were; insertion follows it
            if targets:
                cursor = new_nodes[0]
                for n in new_nodes[1:]:
                    cursor.addnext(n); cursor = n
            else:
                if anchor_before is not None:
                    for n in new_nodes:
                        anchor_before.addprevious(n)
                else:
                    for n in new_nodes:
                        p.append(n)

            if e.comment and new_nodes:
                c = doc.comments.add_comment(text=e.comment, author=author, initials="RE")
                cid = str(c.comment_id)
                start = OxmlElement("w:commentRangeStart"); start.set(qn("w:id"), cid)
                end = OxmlElement("w:commentRangeEnd"); end.set(qn("w:id"), cid)
                ref_run = OxmlElement("w:r")
                rpr = OxmlElement("w:rPr"); rs = OxmlElement("w:rStyle"); rs.set(qn("w:val"), "CommentReference")
                rpr.append(rs); ref_run.append(rpr)
                ref = OxmlElement("w:commentReference"); ref.set(qn("w:id"), cid); ref_run.append(ref)
                new_nodes[0].addprevious(start)
                new_nodes[-1].addnext(end)
                end.addnext(ref_run)
    doc.save(dst)
    return log
