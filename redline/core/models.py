"""Data shapes shared by the parsers, the LLM core and the writers."""
from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class Span(BaseModel):
    """Where a piece of clause text sits in the source file."""
    page: int | None = None          # PDF page index (0-based), None for DOCX
    paragraph: int | None = None     # DOCX body paragraph index
    start: int                       # offset in Clause.text
    end: int
    bbox: tuple[float, float, float, float] | None = None   # PDF line box


class Clause(BaseModel):
    id: str                          # e.g. "6" or "6.2"
    heading: str = ""
    text: str
    spans: list[Span] = Field(default_factory=list)


class HistoricalEdit(BaseModel):
    """One past markup found in a historical contract."""
    source: str
    clause_id: str
    heading: str
    kind: Literal["replace", "delete", "insert", "comment"]
    deleted: str = ""
    inserted: str = ""
    comment: str = ""
    context: str = ""                # paragraph text before the edit
    author: str = ""


class PlaybookRule(BaseModel):
    id: str
    topic: str
    trigger: str                     # what language in an incoming clause sets the rule off
    position: str                    # the company's position
    preferred_language: str          # language the company inserts
    rationale: str
    examples: list[str] = Field(default_factory=list)   # source files the rule was mined from


class Playbook(BaseModel):
    rules: list[PlaybookRule]


class ProposedEdit(BaseModel):
    """What the model proposes for one clause. original_text must be copied verbatim."""
    rule_id: str
    original_text: str
    replacement_text: str
    comment: str


class ClauseReview(BaseModel):
    edits: list[ProposedEdit]


class VerifiedEdit(BaseModel):
    clause_id: str
    rule_id: str
    start: int                       # offsets into Clause.text
    end: int
    original_text: str               # exactly as it appears in the document
    replacement_text: str
    comment: str
    match: Literal["exact", "normalized", "fuzzy"]
    score: float = 100.0


class RejectedEdit(BaseModel):
    clause_id: str
    rule_id: str
    proposed_original: str
    reason: str
