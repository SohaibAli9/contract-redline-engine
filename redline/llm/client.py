"""Async client for any OpenAI-compatible endpoint (cloud API now, vLLM / SGLang on-prem later).

Configuration comes only from the environment (.env): LLM_BASE_URL, LLM_API_KEY, LLM_MODEL.
Every call: counts input tokens against a hard budget, asks for a strict JSON schema, validates the
reply with Pydantic, and retries once with the validation error when the reply does not parse.
Reasoning models think before they answer, so the output budget leaves room for the thinking.
"""
from __future__ import annotations

import asyncio
import copy
import json
import logging
import os
import time
from dataclasses import dataclass, field
from typing import TypeVar

import tiktoken
from openai import AsyncOpenAI
from pydantic import BaseModel, ValidationError

log = logging.getLogger("redline.llm")
T = TypeVar("T", bound=BaseModel)
_ENC = tiktoken.get_encoding("o200k_base")   # gpt-oss tokenizer family


class TokenBudgetError(ValueError):
    pass


def count_tokens(messages: list[dict]) -> int:
    return sum(len(_ENC.encode(m["content"])) + 4 for m in messages)


def strict_schema(model: type[BaseModel]) -> dict:
    """Pydantic schema made acceptable to strict structured outputs."""
    schema = copy.deepcopy(model.model_json_schema())

    def fix(node):
        if isinstance(node, dict):
            if node.get("type") == "object" and "properties" in node:
                node["additionalProperties"] = False
                node["required"] = list(node["properties"].keys())
            for k in ("title", "default"):
                node.pop(k, None)
            for v in node.values():
                fix(v)
        elif isinstance(node, list):
            for v in node:
                fix(v)
    fix(schema)
    return schema


@dataclass
class CallStats:
    calls: int = 0
    retries: int = 0
    failures: int = 0
    schema_violations: int = 0      # replies that broke the schema the endpoint was asked to enforce
    repaired: int = 0               # of those, fixed in code without another call
    input_tokens: list[int] = field(default_factory=list)
    output_tokens: list[int] = field(default_factory=list)
    seconds: list[float] = field(default_factory=list)


@dataclass
class LLMConfig:
    base_url: str
    api_key: str
    model: str
    max_concurrency: int = 8
    max_input_tokens: int = 16_000
    max_output_tokens: int = 6_000
    max_attempts: int = 3
    extra_body: dict = field(default_factory=dict)

    @classmethod
    def from_env(cls) -> "LLMConfig":
        return cls(
            base_url=os.environ["LLM_BASE_URL"],
            api_key=os.environ["LLM_API_KEY"],
            model=os.environ["LLM_MODEL"],
            max_concurrency=int(os.environ.get("LLM_MAX_CONCURRENCY", "8")),
            max_input_tokens=int(os.environ.get("LLM_MAX_INPUT_TOKENS", "16000")),
            max_output_tokens=int(os.environ.get("LLM_MAX_OUTPUT_TOKENS", "6000")),
            max_attempts=int(os.environ.get("LLM_MAX_ATTEMPTS", "3")),
            extra_body=json.loads(os.environ.get("LLM_EXTRA_BODY", "{}")),
        )


class LLM:
    def __init__(self, cfg: LLMConfig):
        self.cfg = cfg
        self.client = AsyncOpenAI(base_url=cfg.base_url, api_key=cfg.api_key, timeout=180)
        self.sem = asyncio.Semaphore(cfg.max_concurrency)
        self.stats = CallStats()

    async def structured(self, messages: list[dict], out: type[T], name: str) -> T:
        n_in = count_tokens(messages)
        if n_in > self.cfg.max_input_tokens:
            raise TokenBudgetError(f"{name}: {n_in} input tokens > budget {self.cfg.max_input_tokens}")
        schema = strict_schema(out)
        msgs = list(messages)
        for attempt in range(self.cfg.max_attempts):
            async with self.sem:
                t0 = time.perf_counter()
                resp = await self.client.chat.completions.create(
                    model=self.cfg.model, messages=msgs, max_tokens=self.cfg.max_output_tokens,
                    response_format={"type": "json_schema",
                                     "json_schema": {"name": name, "strict": True, "schema": schema}},
                    extra_body=self.cfg.extra_body or None,
                )
                self.stats.seconds.append(time.perf_counter() - t0)
            self.stats.calls += 1
            if resp.usage:
                self.stats.input_tokens.append(resp.usage.prompt_tokens)
                self.stats.output_tokens.append(resp.usage.completion_tokens)
            content = (resp.choices[0].message.content or "").strip()
            try:
                return out.model_validate_json(_strip_fences(content))
            except ValidationError as exc:
                self.stats.schema_violations += 1
                fixed = _repair(_strip_fences(content), out)
                if fixed is not None:
                    self.stats.repaired += 1
                    return fixed
                log.warning("%s: reply failed validation (attempt %d): %s", name, attempt + 1, exc.errors()[:2])
                if attempt < self.cfg.max_attempts - 1:
                    self.stats.retries += 1
                    msgs = msgs + [{"role": "assistant", "content": content[:4000]},
                                   {"role": "user", "content": f"That reply did not match the schema: "
                                    f"{str(exc)[:800]}. Reply again with JSON only."}]
        self.stats.failures += 1
        raise ValueError(f"{name}: no valid reply after {self.cfg.max_attempts} attempts")


def _strip_fences(s: str) -> str:
    if s.startswith("```"):
        s = s.split("\n", 1)[1] if "\n" in s else s
        s = s.rsplit("```", 1)[0]
    return s.strip()


def _repair(content: str, out: type[T]) -> T | None:
    """Two known failure shapes, fixed only when the result then validates in full:
    a bare list where the schema wants {"<only list field>": [...]}, and JSON cut off before its
    closing brackets."""
    fields = list(out.model_fields)
    candidates = [content]
    stack = []
    for ch in content:
        if ch in "[{":
            stack.append("]" if ch == "[" else "}")
        elif ch in "]}" and stack:
            stack.pop()
    if stack:
        candidates.append(content + "".join(reversed(stack)))
    for cand in candidates:
        try:
            data = json.loads(cand)
        except json.JSONDecodeError:
            continue
        if isinstance(data, list) and len(fields) == 1:
            data = {fields[0]: data}
        try:
            return out.model_validate(data)
        except ValidationError:
            continue
    return None
