# Phase 3: Structure and Quality - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-04-24
**Phase:** 03-structure-and-quality
**Mode:** --auto (all decisions auto-selected)
**Areas discussed:** Package layout, Module splitting strategy, Test scope and tiering, Linting configuration, CI pipeline design

---

## Package Layout

| Option | Description | Selected |
|--------|-------------|----------|
| Follow requirements spec | Six subpackages matching STRUCT-01: app, bridge, services, workers, platform, settings | ✓ |
| Flatter structure | Fewer subpackages, group by layer (core, ui, platform) | |
| Feature-based | Group by feature (snap/, explorer/, settings/) | |

**User's choice:** [auto] Follow requirements spec (recommended default)
**Notes:** Requirements explicitly name the six subpackages. No reason to deviate.

---

## Module Splitting Strategy

| Option | Description | Selected |
|--------|-------------|----------|
| Bottom-up | Pure functions first, then engines, then MainWindow last | ✓ |
| Top-down | Extract MainWindow shell first, then delegate to modules | |
| All at once | Move everything in one pass | |

**User's choice:** [auto] Bottom-up (recommended, per STATE.md flag about circular import risk)
**Notes:** STATE.md explicitly flags bottom-up as required to avoid circular imports and signal/slot disconnection. workers.py split into two files. explorer_columns.py kept whole (justified exception).

---

## Test Scope and Tiering

| Option | Description | Selected |
|--------|-------------|----------|
| Two-tier split | unit/ (no Qt, CI) + integration/ (PySide6, local only) | ✓ |
| Single tier | All tests require PySide6, run on Windows CI | |
| Three tiers | unit + integration + e2e with full app launch | |

**User's choice:** [auto] Two-tier split (recommended default)
**Notes:** Admin elevation unavailable in GitHub Actions (STATE.md flag). Unit tests cover pure logic (geometry, config, validation). Integration tests need PySide6 and run locally. Frontend tests via Vitest for mapping correctness and smoke tests.

---

## Linting Configuration

| Option | Description | Selected |
|--------|-------------|----------|
| Conservative defaults | E, F, I, UP rules. Line length 100. Ruff formatter. | ✓ |
| Strict | Add D (docstrings), S (bandit security), N (naming). | |
| Minimal | E, F only. No formatter. | |

**User's choice:** [auto] Conservative defaults (recommended default)
**Notes:** Personal project — optimize for catching real bugs, not enforcing style pedantry. Configured in pyproject.toml. No frontend linter in this phase.

---

## CI Pipeline Design

| Option | Description | Selected |
|--------|-------------|----------|
| Single workflow, Ubuntu | One ci.yml, 4 jobs, Ubuntu runner, fast and free | ✓ |
| Matrix with Windows | Ubuntu for lint, Windows for tests | |
| Minimal | Lint only, no test runner in CI | |

**User's choice:** [auto] Single workflow, Ubuntu runner (recommended default)
**Notes:** Unit tests designed to be platform-independent (pure math, config). Windows-specific tests excluded via marker. Ubuntu runner is faster and uses less Actions minutes.

---

## Claude's Discretion

- Ruff rule exceptions for noisy existing code
- pytest conftest.py and fixture structure
- Vitest config details
- pyproject.toml metadata beyond tooling requirements

## Deferred Ideas

None — discussion stayed within phase scope
