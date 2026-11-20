"""Review an incoming contract clause by clause against the playbook. One bounded call per clause."""
from __future__ import annotations

import asyncio
import json
from dataclasses import dataclass, field

from redline.core.models import Clause, ClauseReview, Playbook, RejectedEdit, VerifiedEdit
from redline.core.verify import verify_edit
from redline.llm.client import LLM

SYSTEM = """You review one clause of an incoming construction subcontract for the subcontractor,
using the company's playbook below. Propose an edit only where the clause conflicts with a
playbook rule. If the clause already meets the company's position, return no edits: most
clauses need none. Never rewrite language that already says what the playbook asks for, and
never edit a clause whose topic no rule covers. Return {{"edits": []}} when nothing conflicts.

For each edit:
- original_text: the exact words to change, copied character for character from the clause.
  Keep it short: the phrase or sentence that changes, not the whole clause.
- replacement_text: the text that replaces original_text, using the playbook's preferred language.
  To add a sentence, set original_text to the last sentence of the clause and replacement_text
  to that same sentence followed by the new one.
- comment: one short line for the margin, saying why, in the company's voice.
- rule_id: the playbook rule applied.

PLAYBOOK:
{playbook}"""


def playbook_prompt(pb: Playbook) -> str:
    return json.dumps([{"rule_id": r.id, "topic": r.topic, "trigger": r.trigger, "position": r.position,
                        "preferred_language": r.preferred_language} for r in pb.rules], indent=1)


@dataclass
class ContractReview:
    clauses: list[Clause]
    verified: list[VerifiedEdit] = field(default_factory=list)
    rejected: list[RejectedEdit] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


async def review_clauses(llm: LLM, pb: Playbook, clauses: list[Clause]) -> ContractReview:
    system = SYSTEM.format(playbook=playbook_prompt(pb))
    out = ContractReview(clauses=clauses)

    async def one(c: Clause):
        msgs = [{"role": "system", "content": system},
                {"role": "user", "content": f"Clause {c.id} ({c.heading}):\n{c.text}"}]
        try:
            res = await llm.structured(msgs, ClauseReview, "clause_review")
        except Exception as exc:          # logged, never silently dropped
            out.errors.append(f"clause {c.id}: {exc}")
            return
        taken: list[tuple[int, int]] = []
        for e in res.edits:
            v = verify_edit(c, e)
            if isinstance(v, VerifiedEdit):
                if any(v.start < b and a < v.end or (v.start == v.end and a <= v.start <= b) for a, b in taken):
                    out.rejected.append(RejectedEdit(clause_id=c.id, rule_id=e.rule_id,
                                                     proposed_original=e.original_text,
                                                     reason="overlaps another edit in the clause"))
                    continue
                taken.append((v.start, v.end))
                out.verified.append(v)
            else:
                out.rejected.append(v)

    await asyncio.gather(*(one(c) for c in clauses))
    out.verified.sort(key=lambda v: (int(v.clause_id.split(".")[0]) if v.clause_id[0].isdigit() else 0, v.start))
    return out
