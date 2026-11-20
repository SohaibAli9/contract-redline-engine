"""Command line: thin layer over redline.core. Config comes from .env (see .env.example)."""
from __future__ import annotations

import argparse
import asyncio
import json
import logging
from pathlib import Path

from dotenv import load_dotenv

from redline.core.models import Playbook
from redline.core.pipeline import redline_file
from redline.core.playbook import synthesize, to_markdown
from redline.llm.client import LLM, LLMConfig
from redline.parsers.miner import mine


def _files(p: Path) -> list[Path]:
    return sorted(f for f in (p.iterdir() if p.is_dir() else [p]) if f.suffix.lower() in (".docx", ".pdf"))


async def _mine(args):
    edits = [e for f in _files(Path(args.src)) for e in mine(f)]
    Path(args.out).with_suffix(".edits.json").write_text(json.dumps([e.model_dump() for e in edits], indent=1))
    pb = await synthesize(LLM(LLMConfig.from_env()), edits)
    Path(args.out).write_text(pb.model_dump_json(indent=1))
    Path(args.out).with_suffix(".md").write_text(to_markdown(pb))
    print(f"{len(edits)} past edits -> {len(pb.rules)} rules -> {args.out}")


async def _review(args):
    llm = LLM(LLMConfig.from_env())
    pb = Playbook.model_validate_json(Path(args.playbook).read_text())
    for f in _files(Path(args.src)):
        rep = await redline_file(llm, pb, f, Path(args.out))
        print(f"{f.name}: {len(rep['verified'])} edits written to {rep['output']}, "
              f"{len(rep['rejected'])} rejected, {len(rep['errors'])} errors")
        Path(args.out, f"{f.stem}_report.json").write_text(json.dumps(rep, indent=1))


def main():
    load_dotenv()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    ap = argparse.ArgumentParser(prog="redline")
    sub = ap.add_subparsers(dest="cmd", required=True)
    m = sub.add_parser("mine", help="mine past markups into a playbook")
    m.add_argument("src"); m.add_argument("-o", "--out", default="playbook.json")
    r = sub.add_parser("review", help="redline incoming contracts")
    r.add_argument("src"); r.add_argument("--playbook", default="playbook.json"); r.add_argument("-o", "--out", default="out")
    args = ap.parse_args()
    asyncio.run(_mine(args) if args.cmd == "mine" else _review(args))


if __name__ == "__main__":
    main()
