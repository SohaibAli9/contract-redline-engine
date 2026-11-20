"""Unit tests for every check the engine relies on. No model calls."""
from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pymupdf
import pytest
from docx import Document
from docx.shared import Pt

from redline.core.models import Clause, ProposedEdit, VerifiedEdit
from redline.core.ooxml import paragraph_text, view_text
from redline.core.verify import locate, trim_to_change, verify_edit
from redline.llm.client import _repair, strict_schema
from redline.core.models import ClauseReview
from redline.parsers.chunker import chunk_docx, chunk_pdf
from redline.parsers.miner import mine_docx, mine_pdf
from redline.writers.docx_writer import DocxEdit, apply_edits
from redline.writers import pdf_writer


def _doc(path: Path, pieces: list[tuple[str, bool]], auto=False):
    d = Document()
    if auto:
        h = d.add_paragraph(style="List Number"); h.add_run("Retainage").bold = True
    else:
        h = d.add_paragraph(); h.add_run("5. Retainage").bold = True
    p = d.add_paragraph()
    for text, bold in pieces:
        r = p.add_run(text); r.bold = bold
    d.save(path)
    return path


PIECES = [("Contractor shall withhold retainage of ", False), ("ten percent", True),
          (" (10%) from each payment.", False)]


def test_tracked_change_across_formatted_runs(tmp_path):
    src = _doc(tmp_path / "a.docx", PIECES)
    text = paragraph_text(Document(src).paragraphs[1]._p)
    s = text.index("ten percent (10%)")
    log = apply_edits(str(src), str(tmp_path / "b.docx"), [DocxEdit(1, s, s + 17, "five percent (5%)", "Cap at 5%.")])
    assert log == []
    p = Document(tmp_path / "b.docx").paragraphs[1]._p
    assert view_text(p, "reject") == text                                   # reject all = original
    assert view_text(p, "accept") == text.replace("ten percent (10%)", "five percent (5%)")
    xml = p.xml
    assert "<w:delText" in xml and "<w:ins " in xml and "commentRangeStart" in xml
    # bold run that was struck keeps its formatting inside w:del
    d = p.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del")
    assert d.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b") is not None


def test_two_edits_same_paragraph_and_insert(tmp_path):
    src = _doc(tmp_path / "a.docx", PIECES)
    text = paragraph_text(Document(src).paragraphs[1]._p)
    edits = [DocxEdit(1, text.index("ten"), text.index("ten") + 3, "five"),
             DocxEdit(1, len(text), len(text), " Released at Substantial Completion.")]
    assert apply_edits(str(src), str(tmp_path / "b.docx"), edits) == []
    p = Document(tmp_path / "b.docx").paragraphs[1]._p
    assert view_text(p, "reject") == text
    assert view_text(p, "accept") == text.replace("ten", "five", 1) + " Released at Substantial Completion."


def test_overlapping_edit_is_skipped_and_logged(tmp_path):
    src = _doc(tmp_path / "a.docx", PIECES)
    log = apply_edits(str(src), str(tmp_path / "b.docx"), [DocxEdit(1, 5, 20, "x"), DocxEdit(1, 10, 25, "y")])
    assert len(log) == 1 and "overlaps" in log[0]


def test_output_opens_in_libreoffice(tmp_path):
    src = _doc(tmp_path / "a.docx", PIECES)
    apply_edits(str(src), str(tmp_path / "b.docx"), [DocxEdit(1, 0, 10, "The GC", "c")])
    r = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp_path),
                        str(tmp_path / "b.docx")], capture_output=True, timeout=120)
    assert r.returncode == 0 and (tmp_path / "b.pdf").exists()


def test_miner_reads_deletions_python_docx_text_skips(tmp_path):
    src = _doc(tmp_path / "a.docx", PIECES)
    text = paragraph_text(Document(src).paragraphs[1]._p)
    s = text.index("ten percent (10%)")
    apply_edits(str(src), str(tmp_path / "b.docx"), [DocxEdit(1, s, s + 17, "five percent (5%)", "Cap at 5%.")],
                author="M. Reyes")
    assert "ten percent" not in Document(tmp_path / "b.docx").paragraphs[1].text   # python-docx drops it
    ev = mine_docx(tmp_path / "b.docx")
    assert len(ev) == 1
    e = ev[0]
    assert (e.clause_id, e.kind, e.deleted, e.inserted, e.comment, e.author) == \
           ("5", "replace", "ten percent (10%)", "five percent (5%)", "Cap at 5%.", "M. Reyes")


def test_chunker_auto_numbered_headings(tmp_path):
    src = _doc(tmp_path / "a.docx", PIECES, auto=True)
    cl = chunk_docx(src)
    assert [(c.id, c.heading) for c in cl] == [("1", "Retainage")]


def _pdf_with_split_clause(path: Path):
    doc = pymupdf.open()
    for i in range(3):
        page = doc.new_page()
        page.insert_text((72, 40), f"Subcontract No. 9-{100 + i}  |  Example GC", fontsize=9)
        page.insert_text((280, 800), f"Page {i + 1}", fontsize=9)
    doc[0].insert_text((72, 700), "6. Liquidated Damages", fontname="hebo", fontsize=11)
    doc[0].insert_text((72, 730), "Subcontractor shall pay liquidated damages of $750 per day for each day", fontsize=10)
    doc[1].insert_text((72, 90), "of delay, without any cap or limit.", fontsize=10)
    doc[1].insert_text((72, 130), "7. Indemnification", fontname="hebo", fontsize=11)
    doc[1].insert_text((72, 160), "Subcontractor shall indemnify Contractor.", fontsize=10)
    doc.save(path)
    return path


def test_pdf_chunker_strips_headers_and_merges_page_split(tmp_path):
    cl, dropped = chunk_pdf(_pdf_with_split_clause(tmp_path / "a.pdf"))
    assert dropped == 6
    ld = cl[0]
    assert ld.id == "6" and ld.text.endswith("for each day of delay, without any cap or limit.")
    assert {s.page for s in ld.spans} == {0, 1}


def test_pdf_strikeout_across_page_break_reads_back(tmp_path):
    src = _pdf_with_split_clause(tmp_path / "a.pdf")
    cl, _ = chunk_pdf(src)
    ld = cl[0]
    s = ld.text.index("for each day of delay")
    e = VerifiedEdit(clause_id="6", rule_id="LD", start=s, end=len(ld.text), original_text=ld.text[s:],
                     replacement_text="capped at 10%.", comment="cap", match="exact")
    res = pdf_writer.apply_edits(str(src), str(tmp_path / "b.pdf"), cl, [e])
    assert (res.written, res.readback_ok, res.log) == (1, 1, [])
    out = pymupdf.open(tmp_path / "b.pdf")
    assert [len(list(p.annots())) for p in out] == [1, 1, 0]


def test_printed_redline_pdf_is_mined(tmp_path):
    src = _doc(tmp_path / "a.docx", PIECES)
    text = paragraph_text(Document(src).paragraphs[1]._p)
    s = text.index("ten percent (10%)")
    apply_edits(str(src), str(tmp_path / "b.docx"), [DocxEdit(1, s, s + 17, "five percent (5%)")])
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp_path),
                    str(tmp_path / "b.docx")], capture_output=True, timeout=120, check=True)
    ev = mine_pdf(tmp_path / "b.pdf")
    assert [(e.kind, e.deleted, e.inserted) for e in ev] == [("replace", "ten percent (10%)", "five percent (5%)")]


CLAUSE = Clause(id="10", heading="Claims", text="Subcontractor shall give notice of any claim within three (3) "
                "days after the event. Claims not made within that time are waived.")


@pytest.mark.parametrize("needle,how", [
    ("within three (3) days", "exact"),
    ("within  three (3) days", "normalized"),
    ("within thre (3) days after the event", "fuzzy"),
])
def test_locate(needle, how):
    s, e, got, _ = locate(needle, CLAUSE.text)
    assert got == how and "three (3)" in CLAUSE.text[s:e]


def test_verify_rejects_text_not_in_clause():
    v = verify_edit(CLAUSE, ProposedEdit(rule_id="NOT-1", original_text="within seven (7) business days",
                                         replacement_text="within fourteen (14) days", comment=""))
    assert v.__class__.__name__ == "RejectedEdit"


def test_trim_keeps_redline_minimal():
    lead, tail, new = trim_to_change("within three (3) days after", "within fourteen (14) days after")
    assert new == "fourteen (14) " and lead == len("within ") and tail == len("days after")


def test_json_repair_bare_list_and_cut_off():
    assert _repair('[{"rule_id":"A","original_text":"x","replacement_text":"y","comment":""}]', ClauseReview).edits[0].rule_id == "A"
    assert _repair('{"edits": []', ClauseReview).edits == []
    assert _repair('{"edits": [{"rule_id": "A"}]}', ClauseReview) is None     # missing fields stay a failure


def test_strict_schema_closes_objects():
    sch = strict_schema(ClauseReview)
    item = sch["$defs"]["ProposedEdit"]
    assert item["additionalProperties"] is False and set(item["required"]) == set(item["properties"])
