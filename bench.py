"""Blind benchmark: mine the historical set, redline the held-out set, score against key.json.

key.json is written by make_contracts.py before any model call. Scoring is per redline a human
reviewer made: a key edit counts as found when the engine wrote an edit in the same clause whose
span overlaps the key's span (for a key insertion, any edit in that clause).
"""
from __future__ import annotations

import asyncio
import json
import os
import statistics
import subprocess
import sys
import time
from pathlib import Path

from docx import Document
from dotenv import load_dotenv

from redline.core.models import Playbook
from redline.core.ooxml import paragraph_text, view_text
from redline.core.pipeline import redline_file
from redline.core.playbook import synthesize, to_markdown
from redline.llm.client import LLM, LLMConfig
from redline.parsers.chunker import chunk
from redline.parsers.miner import mine

DATA = Path(os.environ.get("DATA_DIR", "data"))


def overlap(a0, a1, b0, b1) -> bool:
    return a0 < b1 and b0 < a1


def score(rep: dict, key: list[dict], clauses) -> dict:
    by_id = {c.id: c for c in clauses}
    found, matched_v = 0, set()
    rows = []
    for k in key:
        c = by_id.get(k["clause"])
        hit = None
        if c is not None:
            if k["original"]:
                ks = c.text.find(k["original"])
                ke = ks + len(k["original"])
            for i, v in enumerate(rep["verified"]):
                if v["clause_id"] != k["clause"]:
                    continue
                if not k["original"] or (ks >= 0 and (overlap(v["start"], max(v["end"], v["start"] + 1), ks, ke))):
                    hit = i
                    break
        if hit is not None:
            found += 1
            matched_v.add(hit)
        rows.append({"clause": k["clause"], "rule": k["rule"], "found": hit is not None})
    # an engine edit is "right" if it matches a key edit
    for i, v in enumerate(rep["verified"]):
        if i in matched_v:
            continue
        for k in key:
            if k["clause"] == v["clause_id"]:
                matched_v.add(i)
                break
    return {"key_edits": len(key), "found": found, "engine_edits": len(rep["verified"]),
            "engine_edits_matching_key": len(matched_v), "rows": rows}


def docx_integrity(src: Path, dst: Path) -> dict:
    """Reject-all must give back the original text of every paragraph; the file must reopen and convert."""
    a, b = Document(src), Document(dst)
    orig = [paragraph_text(p._p) for p in a.paragraphs]
    rej = [view_text(p._p, "reject") for p in b.paragraphs]
    ok_reject = orig == rej
    conv = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(dst.parent / "pdf"),
                           str(dst)], capture_output=True, timeout=120)
    pdf = dst.parent / "pdf" / (dst.stem + ".pdf")
    return {"reopens": True, "reject_all_equals_original": ok_reject,
            "libreoffice_converts": conv.returncode == 0 and pdf.exists(),
            "comments": len(list(b.comments))}


async def main(tag: str):
    load_dotenv(DATA / ".env")
    load_dotenv()
    if len(sys.argv) > 2:
        os.environ["LLM_MODEL"] = sys.argv[2]
    run = DATA / "runs" / tag
    run.mkdir(parents=True, exist_ok=True)
    llm = LLM(LLMConfig.from_env())

    t0 = time.perf_counter()
    hist = sorted((DATA / "historical").iterdir())
    edits = [e for f in hist for e in mine(f)]
    (run / "mined_edits.json").write_text(json.dumps([e.model_dump() for e in edits], indent=1))
    pb_path = run / "playbook.json"
    if pb_path.exists():
        pb = Playbook.model_validate_json(pb_path.read_text())
    else:
        pb = await synthesize(llm, edits)
        pb_path.write_text(pb.model_dump_json(indent=1))
        (run / "playbook.md").write_text(to_markdown(pb))
    t_mine = time.perf_counter() - t0
    mine_calls = llm.stats.calls

    key = json.loads((DATA / "key.json").read_text())
    results = []
    for src in sorted((DATA / "incoming").iterdir()):
        rep = await redline_file(llm, pb, src, run / "out")
        clauses = chunk(src)
        rep["score"] = score(rep, key[src.stem], clauses)
        if src.suffix == ".docx":
            rep["integrity"] = docx_integrity(src, Path(rep["output"]))
        results.append(rep)
        s = rep["score"]
        print(f"{src.name}: key {s['key_edits']}, found {s['found']}, engine {s['engine_edits']} "
              f"(matching {s['engine_edits_matching_key']}), rejected {len(rep['rejected'])}, "
              f"errors {len(rep['errors'])}, {rep['seconds']}s", flush=True)

    st = llm.stats
    summary = {
        "model": llm.cfg.model, "historical_files": len(hist), "mined_edits": len(edits),
        "playbook_rules": len(pb.rules), "mine_seconds": round(t_mine, 1), "mine_calls": mine_calls,
        "key_edits": sum(r["score"]["key_edits"] for r in results),
        "found": sum(r["score"]["found"] for r in results),
        "engine_edits": sum(r["score"]["engine_edits"] for r in results),
        "engine_edits_matching_key": sum(r["score"]["engine_edits_matching_key"] for r in results),
        "written": sum(r["written"] for r in results),
        "rejected": sum(len(r["rejected"]) for r in results),
        "errors": sum(len(r["errors"]) for r in results),
        "match_types": {m: sum(v["match"] == m for r in results for v in r["verified"])
                        for m in ("exact", "normalized", "fuzzy")},
        "pdf_readback_ok": sum(r["pdf_readback_ok"] or 0 for r in results if r["file"].endswith(".pdf")),
        "pdf_written": sum(r["written"] for r in results if r["file"].endswith(".pdf")),
        "docx_integrity": {r["file"]: r.get("integrity") for r in results if r["file"].endswith(".docx")},
        "calls": st.calls, "retries": st.retries, "failures": st.failures,
        "schema_violations": st.schema_violations, "repaired_in_code": st.repaired,
        "max_input_tokens": max(st.input_tokens or [0]), "max_output_tokens": max(st.output_tokens or [0]),
        "median_call_seconds": round(statistics.median(st.seconds), 2) if st.seconds else None,
        "p95_call_seconds": round(sorted(st.seconds)[int(0.95 * (len(st.seconds) - 1))], 2) if st.seconds else None,
    }
    (run / "results.json").write_text(json.dumps({"summary": summary, "contracts": results}, indent=1))
    print(json.dumps(summary, indent=1))


if __name__ == "__main__":
    asyncio.run(main(sys.argv[1] if len(sys.argv) > 1 else "run1"))
