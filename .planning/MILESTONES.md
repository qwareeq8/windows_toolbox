# Milestones

## v1.0 — Hardening and Cleanup

**Shipped:** 2026-04-24
**Phases:** 4 | **Plans:** 14 | **Commits:** 77
**Codebase:** 4,355 LOC Python, 1,482 LOC React/JS, 607 LOC tests

### Delivered

Transformed a working-but-rough Windows desktop utility into a clean, tested, installable application with proper architecture.

### Key Accomplishments

1. Removed all stale "Windows Toolbox" naming and consolidated version to single source of truth
2. Built reproducible 6-script PowerShell build pipeline (bootstrap through installer)
3. Locked down WebEngine security, implemented draft/commit state model, removed all fake UI controls
4. Restructured monolithic main.py into testable virelo/ package with 6 subpackages
5. Added Ruff linting, pytest unit tests, Vitest frontend tests, and GitHub Actions CI
6. Separated hotkey detection from window movement, extracted ExplorerService, added multi-monitor snap geometry tests

### Known Deferred Items

2 verification artifacts from earlier phases (Phase 1: human_needed, Phase 3: gaps_found) — acknowledged at close. All requirements checked off and tests pass.

### Archive

- `milestones/v1.0-ROADMAP.md` — full phase details
- `milestones/v1.0-REQUIREMENTS.md` — 52 requirements, all complete
