"""Synthesize a structured company playbook from mined historical edits.

Edits are grouped by clause heading and each group goes to the model in its own call, so every
call stays far below the token budget however many contracts are mined.
"""
from __future__ import annotations

import asyncio
import json
import re
from collections import defaultdict

from pydantic import BaseModel

from redline.core.models import HistoricalEdit, Playbook, PlaybookRule
from redline.llm.client import LLM

SYSTEM = """You turn a construction subcontractor's past contract markups into playbook rules.
You get every past edit made under one clause heading, from several contracts. Write the rules
those edits show: what language in an incoming contract triggers the edit, the company's position,
and the replacement language it inserts (copy the company's own inserted wording when it is
consistent). One rule per distinct position. Do not invent positions the edits do not show.
Reply with one JSON object: {"rules": [{"topic", "trigger", "position", "preferred_language",
"rationale", "sources"}]}."""


class _Rules(BaseModel):
    class R(BaseModel):
        topic: str
        trigger: str
        position: str
        preferred_language: str
        rationale: str
        sources: list[str]
    rules: list[R]


def _slug(s: str) -> str:
    return re.sub(r"[^A-Z]", "", s.upper())[:4] or "GEN"


async def synthesize(llm: LLM, edits: list[HistoricalEdit]) -> Playbook:
    failed: list[str] = []
    groups: dict[str, list[HistoricalEdit]] = defaultdict(list)
    for e in edits:
        groups[e.heading or "General"].append(e)

    async def one(heading: str, evs: list[HistoricalEdit]) -> list[PlaybookRule]:
        payload = [{"source": e.source, "kind": e.kind, "deleted": e.deleted, "inserted": e.inserted,
                    "comment": e.comment} for e in evs]
        msgs = [{"role": "system", "content": SYSTEM},
                {"role": "user", "content": f"Clause heading: {heading}\nPast edits ({len(evs)}):\n"
                                             + json.dumps(payload, indent=1)}]
        try:
            res = await llm.structured(msgs, _Rules, "playbook_rules")
        except ValueError as exc:      # reported below; one bad heading must not hide the others
            failed.append(f"{heading}: {exc}")
            return []
        return [PlaybookRule(id=f"{_slug(heading)}-{i + 1}", topic=r.topic, trigger=r.trigger,
                             position=r.position, preferred_language=r.preferred_language,
                             rationale=r.rationale, examples=sorted(set(r.sources)))
                for i, r in enumerate(res.rules)]

    results = await asyncio.gather(*(one(h, evs) for h, evs in groups.items()))
    if failed:
        raise RuntimeError("playbook synthesis failed for: " + "; ".join(failed))
    return Playbook(rules=[r for rs in results for r in rs])


def to_markdown(pb: Playbook) -> str:
    out = ["# Company playbook (mined from past markups)", ""]
    for r in pb.rules:
        out += [f"## {r.id}: {r.topic}", f"- **Trigger:** {r.trigger}", f"- **Position:** {r.position}",
                f"- **Insert:** {r.preferred_language}", f"- **Why:** {r.rationale}",
                f"- **Seen in:** {', '.join(r.examples)}", ""]
    return "\n".join(out)
