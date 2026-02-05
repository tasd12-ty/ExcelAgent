# Repository Guidelines

## Project Structure & Module Organization
- `agent/`: Core agent loop and tool dispatching (`agent/core.py`, `agent/dispatcher.py`).
- `tools/`: Excel read/write/format/analyze utilities and `run_python` executor.
- `llm/`: OpenRouter/OpenAI-compatible client and multimodal helpers.
- `skills/xlsx/`: Runtime xlsx skill (formula recalculation and Office helpers).
- `benchmark/` and `SpreadsheetBench-NoDocker/`: Benchmark runner and upstream dataset mirror.
- `examples/`: End-to-end demo script (`examples/demo.py`).
- `tests/`: Pytest suite and fixtures.
- `docs/`: Documentation index.

## Build, Test, and Development Commands
- `pip install -e .`: Install the package in editable mode.
- `pip install -e .[dev]`: Install dev dependencies for testing.
- `python examples/demo.py`: Run a full demo (create workbook, analyze, chart, optional multimodal).
- `pytest`: Run the test suite.
- `python benchmark/run_benchmark.py --dataset sample_data_200 --model <model-id> --max_steps 15`: Run SpreadsheetBench tasks.

## Coding Style & Naming Conventions
- Language: Python 3.10+.
- Indentation: 4 spaces (PEP 8 style).
- Naming: `snake_case` for functions/variables, `PascalCase` for classes, `SCREAMING_SNAKE_CASE` for constants.
- No formatter or linter is configured; keep diffs minimal and readable.

## Testing Guidelines
- Framework: `pytest`.
- Tests live in `tests/` and are named `test_*.py`.
- Prefer small, focused tests (e.g., `tests/test_reader.py`, `tests/test_writer.py`).
- LLM tests in `tests/test_llm.py` are skipped when `OPENROUTER_API_KEY` is not set.

## Commit & Pull Request Guidelines
- Commit messages follow a Conventional Commits-style prefix, e.g., `feat: 初始化 ExcelAgent 智能体项目`.
- Use concise, action-oriented summaries. Prefer `feat:`, `fix:`, `docs:`, `chore:`.
- PRs should include a short description, testing notes (`pytest` output or “not run”), and any config changes.

## Security & Configuration Tips
- Never commit real API keys. Configure via `.env` or environment variables.
- Required: `OPENROUTER_API_KEY`. Optional: `OPENROUTER_BASE_URL`, `OPENROUTER_MODEL`, `LLM_MAX_TOKENS`.
