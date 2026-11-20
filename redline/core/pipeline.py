"""Engine entry points used by the CLI (and later by a web worker). No CLI code in here."""
from __future__ import annotations

import time
from pathlib import Path

from redline.core.models import Clause, Playbook, VerifiedEdit
from redline.core.review import ContractReview, review_clauses
from redline.llm.client import LLM
from redline.parsers.chunker import chunk
from redline.writers import docx_writer, pdf_writer


def docx_edits(clauses: list[Clause], edits: list[VerifiedEdit]) -> tuple[list[docx_writer.DocxEdit], list[str]]:
    """Clause offsets -> paragraph offsets. An edit that crosses paragraphs is logged, not written."""
    by_id = {c.id: c for c in clauses}
    out, log = [], []
    for e in edits:
        c = by_id[e.clause_id]
        sp = next((s for s in c.spans if s.start <= e.start and e.end <= s.end), None)
        if sp is None:
            log.append(f"clause {e.clause_id}: edit crosses paragraphs, not written")
            continue
        out.append(docx_writer.DocxEdit(sp.paragraph, e.start - sp.start, e.end - sp.start,
                                        e.replacement_text, e.comment))
    return out, log


async def redline_file(llm: LLM, playbook: Playbook, src: Path, out_dir: Path,
                      author: str = "Redline Engine") -> dict:
    t0 = time.perf_counter()
    clauses = chunk(src)
    review: ContractReview = await review_clauses(llm, playbook, clauses)
    out_dir.mkdir(parents=True, exist_ok=True)
    dst = out_dir / f"{src.stem}_redlined{src.suffix}"
    if src.suffix.lower() == ".docx":
        edits, log = docx_edits(clauses, review.verified)
        log += docx_writer.apply_edits(str(src), str(dst), edits, author=author)
        written, readback = len(edits) - sum("skipped" in l for l in log), None
    else:
        res = pdf_writer.apply_edits(str(src), str(dst), clauses, review.verified, author=author)
        written, readback, log = res.written, res.readback_ok, res.log
    return {
        "file": src.name, "output": str(dst), "clauses": len(clauses),
        "verified": [v.model_dump() for v in review.verified],
        "rejected": [r.model_dump() for r in review.rejected],
        "errors": review.errors, "written": written, "pdf_readback_ok": readback, "log": log,
        "seconds": round(time.perf_counter() - t0, 1),
    }
