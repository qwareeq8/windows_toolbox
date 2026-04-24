# Phase 1: Hygiene and Build Pipeline - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-04-24
**Phase:** 01-Hygiene and Build Pipeline
**Mode:** Auto (all decisions auto-selected with recommended defaults)
**Areas discussed:** Version strategy, Build script architecture, Repository documentation, .gitignore strategy

---

## Version Strategy

| Option | Description | Selected |
|--------|-------------|----------|
| Vite define + ISS /D flag | Version in app_config.py, injected via Vite define and Inno Setup /D flag at build time | ✓ |
| Build-time JSON file | Generate a version.json read by all consumers | |
| Manual sync with verification | Keep manual duplication but add a verification script to catch drift | |

**User's choice:** Vite define + ISS /D flag (auto-selected recommended)
**Notes:** Standard approach. Vite define replaces at build time with zero runtime cost. ISS /D flag is the documented way to externalize version in Inno Setup.

---

## Build Script Architecture

| Option | Description | Selected |
|--------|-------------|----------|
| Multiple specialized scripts | bootstrap, clean, build-frontend, build-app, build-installer, verify-release — each with precondition checks | ✓ |
| Single monolithic script | One build-all.ps1 that handles everything | |

**User's choice:** Multiple specialized scripts (auto-selected recommended)
**Notes:** Matches user's analysis document exactly. Each script is independently runnable and validates its own preconditions.

---

## Repository Documentation

| Option | Description | Selected |
|--------|-------------|----------|
| Minimal set | .gitignore, README, LICENSE (MIT), CLAUDE.md | ✓ |
| Full set | Add SECURITY.md, CONTRIBUTING.md, CHANGELOG.md, docs/ directory | |

**User's choice:** Minimal set (auto-selected recommended)
**Notes:** Personal tool with one user. Full documentation suite deferred to later phase (Phase 10 in user's analysis).

---

## .gitignore Strategy

| Option | Description | Selected |
|--------|-------------|----------|
| GitHub Python template + extensions | Python template + Node, PyInstaller, Inno Setup, IDE patterns | ✓ |
| Custom minimal | Only the specific patterns needed for this project | |

**User's choice:** GitHub Python template + extensions (auto-selected recommended)
**Notes:** Comprehensive coverage prevents accidental commits of build artifacts, IDE files, and virtual environments.

---

## Claude's Discretion

- Build script error message formatting
- README section ordering
- CLAUDE.md organization depth
- bootstrap.ps1 venv creation strategy
