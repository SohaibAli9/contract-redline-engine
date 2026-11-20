# Contract Redline Engine

Python engine that mines review rules from marked-up construction contracts (`.docx`, `.pdf`), chunks incoming files into clauses, reviews them via an OpenAI-compatible LLM under 16,000 tokens per call, and writes Word Track Changes, margin comments, and PDF strikeouts with popup notes.

## Layout

- `redline/core/`: models, verification and trimming, playbook synthesis, clause review, pipeline.
- `redline/parsers/`: `lxml` OOXML miner (`w:del`, `w:ins`, `w:comment`), PDF and printed-redline miner, chunker.
- `redline/writers/`: `.docx` Track Changes and comments, isolated `PyMuPDF` writer.
- `redline/llm/`: async client, token counting, JSON validation and repair (`redline/llm/client.py`).
- `redline/cli/`: `mine` and `review` commands, separate from core.

`samples/` holds `historical/` (2 PDFs, 1 Word), `incoming/` (1 Word, 1 PDF), and `output/` (redlined pair, `playbook.json`, `playbook.md`, `benchmark_summary.json`). Root: `Dockerfile`, `.env.example`, `make_contracts.py`, `bench.py`.

## Quick start

```bash
uv sync
cp .env.example .env
# Required: LLM_BASE_URL, LLM_API_KEY, LLM_MODEL
uv run python -m redline.cli mine samples/historical -o playbook.json
uv run python -m redline.cli review samples/incoming --playbook samples/output/playbook.json -o out
uv run pytest
```

Optional `.env` keys: `LLM_MAX_CONCURRENCY`, `LLM_MAX_INPUT_TOKENS`, `LLM_MAX_OUTPUT_TOKENS`, `LLM_MAX_ATTEMPTS`, `LLM_EXTRA_BODY`. `uv run pytest` runs 16 offline unit tests (no model calls), separate from the benchmark scripts.

Docker (ran and wrote 9 edits on one sample contract):

```bash
docker build -t redline .
docker run --rm --env-file .env -v "$PWD/samples:/samples:ro" -v "$PWD/out:/out" redline review /samples/incoming --playbook /samples/output/playbook.json -o /out
```

## How each check works

- **Miner:** reads `w:del` XML via `lxml` (`python-docx` skips deletions), plus PDF annotations and printed-redline colors.
- **Header and footer removal:** strips repeating PDF headers, footers, and page numbers, merges cross-page clauses, and resolves Word numbering.
- **JSON validation and repair:** validates every reply in code. Broken replies (bare lists, missing fields, cut-off JSON) are repaired only when valid after repair, so none reach a file.
- **Verbatim verification and trimming:** requires an exact clause match before writing, then trims to changed words (`"five (5) days"` to `"fourteen (14) days"`).
- **Reject All test:** written `.docx` files keep bold defined terms inside strikes and restore exact paragraphs on Reject All.
- **PDF readback:** writes strikeouts and popup notes (one per page across breaks), then reads back struck text for exact matches.

## Benchmark

15 invented subcontracts (10 historical: 8 Word, 1 annotated PDF, 1 printed-redline PDF, and 5 held-out: 3 Word, 2 PDF) against a pre-written key. Rebuild with `uv run python make_contracts.py` then `uv run python bench.py run1` (writes to `data/`, override via `DATA_DIR`, needs LibreOffice for PDF exports and checks).

`gpt-oss-120b` (cloud API):

- **Miner and chunker:** 64 edits into 10 source-linked rules, 14 of 14 clauses per held-out file.
- **LLM core:** 78 calls (8 playbook, 70 reviews), 1,667 max prompt tokens, 1,271 max output tokens, 1.36 s median. 7 of 78 strict-JSON replies broke schema and were repaired.
- **Accuracy and writers:** 31 of 31 key redlines, 36 edits written (35 on key clauses, 1 extra deleted an Owner payment-timing sentence), 36 of 36 exact matches (0 rejected), 17 of 17 PDF readbacks, 3 Word files reopen, pass Reject All, and render in LibreOffice (10, 2, 7 comments).

`deepseek-r1` (full model, cloud API, same playbook and contracts):

- **Accuracy:** 26 of 31 key redlines found, 27 edits written, 0 wrong edits, 0 schema violations.
- **Latency and tokens:** 9.32 s median (p95 47.86 s), longest reply 5,155 output tokens for one clause, mostly thinking.

## Running on your own GPU

Set `LLM_BASE_URL`, `LLM_API_KEY`, and `LLM_MODEL` in `.env` for vLLM or SGLang.

On one 80 GB GPU, `gpt-oss-120b` fits in 4-bit MXFP4. Full DeepSeek-R1 (671B) needs roughly 336 GB at 4 bits, so benchmark a distill (`R1-Distill-Qwen-32B` or 8-bit Llama-70B) in the POC. Reasoning tokens count against the 16,000-token budget, so token counting and JSON validation catch cut-offs.

## Notes

- **Invented data:** all 15 contracts are invented (no AIA or ConsensusDocs text) and similar in shape. Test on real company contracts.
- **Microsoft Word:** `.docx` output passed Reject All and LibreOffice checks, not yet opened in Microsoft Word.
- **PyMuPDF license:** AGPL without a commercial Artifex license. The PDF writer sits in one small module for Phase 2 swapping.
