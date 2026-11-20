"""Verbatim substring verification: map a model's proposed edit back to exact document offsets.

The model is told to copy text exactly. It often doesn't (curly vs straight quotes, a dropped
space, a word swapped). Order of attempts:
  1. exact substring
  2. after normalising quotes, dashes and whitespace, with an index map back to the original
  3. fuzzy alignment (rapidfuzz), accepted only at a high score and snapped to word boundaries
Anything else is rejected and logged; nothing is written for it.
Then the edit is trimmed to the words that actually change, so the redline is as small as a
person would make it.
"""
from __future__ import annotations

import re

from rapidfuzz import fuzz

from redline.core.models import Clause, ProposedEdit, RejectedEdit, VerifiedEdit

_TRANS = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"', "–": "-",
                        "—": "-", " ": " ", "‑": "-"})
FUZZY_MIN = 92.0


def _normalise(s: str) -> tuple[str, list[int]]:
    """Normalised text plus, for each normalised char, its index in s."""
    out, idx = [], []
    prev_space = False
    for i, ch in enumerate(s.translate(_TRANS)):
        if ch.isspace():
            if prev_space:
                continue
            ch, prev_space = " ", True
        else:
            prev_space = False
        out.append(ch)
        idx.append(i)
    return "".join(out), idx


def locate(needle: str, hay: str) -> tuple[int, int, str, float] | None:
    if not needle.strip():
        return None
    i = hay.find(needle)
    if i >= 0:
        return i, i + len(needle), "exact", 100.0
    nh, hmap = _normalise(hay)
    nn, _ = _normalise(needle.strip())
    j = nh.find(nn)
    if j >= 0:
        return hmap[j], hmap[j + len(nn) - 1] + 1, "normalized", 100.0
    al = fuzz.partial_ratio_alignment(nn, nh, score_cutoff=FUZZY_MIN)
    if al is None:
        return None
    s, e = al.dest_start, al.dest_end
    while s > 0 and nh[s - 1].isalnum():       # snap to word boundaries
        s -= 1
    while e < len(nh) and nh[e].isalnum():
        e += 1
    return hmap[s], hmap[e - 1] + 1, "fuzzy", float(al.score)


def trim_to_change(orig: str, repl: str) -> tuple[int, int, str]:
    """Drop the words orig and repl share at both ends. Returns (lead, tail) cut from orig and the new repl."""
    tok = lambda s: re.findall(r"\S+\s*|\s+", s)
    a, b = tok(orig), tok(repl)
    p = 0
    while p < min(len(a), len(b)) and a[p] == b[p]:
        p += 1
    q = 0
    while q < min(len(a), len(b)) - p and a[-1 - q] == b[-1 - q]:
        q += 1
    lead = len("".join(a[:p]))
    tail = len("".join(a[len(a) - q:])) if q else 0
    new = "".join(b[p:len(b) - q if q else len(b)])
    return lead, tail, new


def verify_edit(clause: Clause, e: ProposedEdit) -> VerifiedEdit | RejectedEdit:
    hit = locate(e.original_text, clause.text)
    if hit is None:
        return RejectedEdit(clause_id=clause.id, rule_id=e.rule_id, proposed_original=e.original_text,
                            reason="original_text not found in clause")
    s, t, how, score = hit
    found = clause.text[s:t]
    lead, tail, repl = trim_to_change(found, e.replacement_text)
    s, t = s + lead, t - tail
    if s == t and not repl:
        return RejectedEdit(clause_id=clause.id, rule_id=e.rule_id, proposed_original=e.original_text,
                            reason="replacement equals original")
    return VerifiedEdit(clause_id=clause.id, rule_id=e.rule_id, start=s, end=t,
                        original_text=clause.text[s:t], replacement_text=repl, comment=e.comment,
                        match=how, score=score)
