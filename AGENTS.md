# AGENTS.md (repo instructions for coding agents)

This repository is `xpyxl`, a typed Python 3.12 library for building and rendering Excel (and HTML) reports.

If you are an agent making changes:
- Keep diffs small and focused.
- Prefer the existing patterns in `src/xpyxl/` and tests in `tests/`.
- Do not add new tooling/deps unless requested.

## Repo layout
- `src/xpyxl/`: library code (public API re-exported from `src/xpyxl/__init__.py`).
- `src/xpyxl/engines/`: rendering backends.
- `tests/`: pytest suite.
- `tests/end_to_end/`: demo modules and an end-to-end test that saves files into `.testing/`.
- `scripts/`: ad-hoc scripts (e.g. benchmarking).

## Environment
- Python: `3.12` (see `.python-version` and CI workflow).
- Dependency manager/build tool: `uv` (see `uv.lock` and `.github/workflows/publish.yml`).

## Commands (build / lint / test)

### Install / sync
- Create/sync a dev environment: `uv sync --dev`
- Run a one-off command in the env: `uv run <cmd>`

Notes:
- `uv.lock` exists; prefer workflows that respect the lock file for reproducibility.
- If you add/remove dependencies, update `pyproject.toml` and refresh `uv.lock`.

### Tests (pytest)
- Run all tests: `uv run pytest`
- Run tests with output: `uv run pytest -q` or `uv run pytest -vv`
- Stop on first failure: `uv run pytest -x --maxfail=1`

Run a single file:
- `uv run pytest tests/test_html_engine.py`

Run a single test (node id):
- `uv run pytest tests/test_html_engine.py::test_tailwind_cdn_included`

Run tests matching a substring/expr:
- `uv run pytest -k "html and not missing"`

Run only end-to-end test:
- `uv run pytest tests/end_to_end/test_demo.py`

### Type checking (lint-ish)
There is no dedicated linter configured in this repo, but type checking is part of the dev toolchain.
- Run Pyright: `uv run pyright`

### Build
CI builds with `uv build`.
- Build sdist + wheel: `uv build --sdist --wheel`

### Misc
- Benchmark script: `uv run python scripts/benchmark.py`

### Quick sanity (optional)
These are useful when iterating on core rendering/layout changes:
- Import check: `uv run python -c "import xpyxl"`
- Bytecode compile: `uv run python -m compileall src/xpyxl`
- Minimal save smoke test: `uv run python -c "import xpyxl as x; wb=x.workbook()[x.sheet('S')[x.row()['ok']]]; wb.save(engine='html')"`

### CI parity
Before opening a PR, prefer running:
- `uv run pyright`
- `uv run pytest`

## Cursor/Copilot rules
- No Cursor rules found in `.cursor/rules/` or `.cursorrules` at time of writing.
- No Copilot instructions found in `.github/copilot-instructions.md` at time of writing.

If these files appear later, follow them and update this document accordingly.

## Code style (Python)

### General
- Target Python `>=3.12`.
- Prefer small, composable functions over deeply nested logic.
- Favor determinism: avoid randomness/time dependence in core logic unless explicitly required.

### Imports
Follow the existing import ordering:
1. `from __future__ import annotations` (used widely; include in new modules).
2. Standard library imports.
3. Third-party imports.
4. Local imports (`from .foo import Bar`).

Prefer:
- `collections.abc` for `Sequence`, `Mapping`, etc.
- `pathlib.Path` for filesystem paths.

### Formatting
- No formatter is enforced by configuration files in this repo.
- Keep formatting consistent with the current codebase (Black-compatible style is already used: trailing commas, hanging indents, one-arg-per-line when wrapping).
- Keep line length reasonable (the current code is typically ~88-ish, but follow surrounding code).

### Types
- The codebase is fully typed; keep/add type hints for new/changed public functions.
- Prefer modern syntax:
  - `X | None` instead of `Optional[X]`.
  - `list[str]`, `tuple[int, ...]`, etc.
  - `TypeAlias` for aliases (see `src/xpyxl/nodes.py`).
- Use `assert_never(...)` for exhaustiveness in tagged unions when appropriate.
- Avoid `Any` unless it is truly necessary at API boundaries.

### Naming
- Modules: `snake_case.py`.
- Classes: `CamelCase`.
- Functions/vars: `snake_case`.
- Constants: `UPPER_SNAKE_CASE`.
- Private helpers: prefix `_`.

### Data model & immutability
- Core AST-like nodes are immutable dataclasses (`@dataclass(frozen=True)`).
- Prefer returning new values over mutating existing nodes.

### Errors and validation
- Use `TypeError` for incorrect types/shape (e.g. passing a `RowNode` where a scalar cell is expected).
- Use `ValueError` for invalid values/ranges (e.g. negative sizes).
- Error messages should be actionable and stable (tests may match them).
- Prefer:
  - `msg = "..."; raise ValueError(msg)` when the message is reused or multi-line.
  - `raise ValueError("...")` for simple one-liners.

### Public API surface
- Public API is re-exported in `src/xpyxl/__init__.py`.
- If you add a new user-facing function/type, consider whether it must be exported and added to `__all__`.
- Keep backwards compatibility in mind: avoid renames/breaking changes unless requested.

### Performance
- Rendering can be performance-sensitive. Avoid introducing quadratic behavior in layout/render loops.
- Prefer tuples for immutable collections exposed on nodes (consistent with current dataclasses).

## Tests
- Tests use `pytest`.
- Prefer unit tests in `tests/` and keep them deterministic.
- For exception behavior, use `pytest.raises(..., match=...)` with a stable substring/regex.
- The end-to-end test writes to `.testing/` (ignored by git); do not commit generated output files.

### Test file I/O
- Prefer `tempfile.TemporaryDirectory()` for tests that write outputs.
- If you need stable artifacts for manual inspection, write under `.testing/`.
- Keep file names deterministic to avoid flaky diffs.

## Git hygiene
- Do not commit local artifacts: `__pycache__/`, `.pytest_cache/`, `.testing/`, `*.egg-info/`.
- Keep `uv.lock` in sync with `pyproject.toml` when dependencies change.
