FROM python:3.12-slim
COPY --from=ghcr.io/astral-sh/uv:0.5.11 /uv /usr/local/bin/uv
WORKDIR /app
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev --no-install-project
COPY redline ./redline
ENV PATH="/app/.venv/bin:$PATH" PYTHONPATH=/app PYTHONUNBUFFERED=1
# config comes from the environment / --env-file .env (LLM_BASE_URL, LLM_API_KEY, LLM_MODEL)
ENTRYPOINT ["python", "-m", "redline.cli"]
CMD ["--help"]
