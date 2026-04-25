---
phase: 7
slug: ci-and-repo-hygiene
status: draft
nyquist_compliant: false
wave_0_complete: false
created: 2026-04-24
---

# Phase 7 — Validation Strategy

> Per-phase validation contract for feedback sampling during execution.

---

## Test Infrastructure

| Property | Value |
|----------|-------|
| **Framework** | pytest 9.x (Python), vitest (frontend) |
| **Config file** | pyproject.toml (pytest section), frontend/vitest.config.js |
| **Quick run command** | `pytest tests/unit/ -q` |
| **Full suite command** | `pytest tests/unit/ -q && cd frontend && npx vitest run` |
| **Estimated runtime** | ~15 seconds |

---

## Sampling Rate

- **After every task commit:** Run `pytest tests/unit/ -q`
- **After every plan wave:** Run `pytest tests/unit/ -q && cd frontend && npx vitest run`
- **Before `/gsd-verify-work`:** Full suite must be green
- **Max feedback latency:** 15 seconds

---

## Per-Task Verification Map

| Task ID | Plan | Wave | Requirement | Threat Ref | Secure Behavior | Test Type | Automated Command | File Exists | Status |
|---------|------|------|-------------|------------|-----------------|-----------|-------------------|-------------|--------|
| 07-01-01 | 01 | 1 | CI-01 | — | N/A | integration | `git ls-files --error-unmatch icon.ico branding/ frontend/index.html` | ✅ | ⬜ pending |
| 07-01-02 | 01 | 1 | CI-02 | — | N/A | integration | `pwsh scripts/clean.ps1; test ! -d __pycache__` | ✅ | ⬜ pending |
| 07-02-01 | 02 | 1 | CI-03 | — | N/A | integration | `grep -rn "Windows Toolbox" . --include="*.py" --exclude-dir=.github --exclude-dir=node_modules --exclude-dir=.git --exclude-dir=dist --exclude-dir=build --exclude-dir=.planning` | ✅ | ⬜ pending |
| 07-02-02 | 02 | 1 | CI-04 | — | N/A | unit | `pytest tests/unit/ -q` | ✅ | ⬜ pending |
| 07-03-01 | 03 | 2 | CI-05 | — | N/A | integration | `python -c "from virelo.app.config import APP_VERSION; print(APP_VERSION)"` | ✅ | ⬜ pending |
| 07-03-02 | 03 | 2 | CI-06 | — | N/A | grep | `grep "Python 3.12" README.md` | ✅ | ⬜ pending |
| 07-03-03 | 03 | 2 | CI-07 | — | N/A | grep | `grep "github.com" installer/virelo.iss` | ✅ | ⬜ pending |
| 07-04-01 | 04 | 2 | CI-08 | — | N/A | grep | `grep "class SnapRestoreController" virelo/services/snap.py` | ✅ | ⬜ pending |

*Status: ⬜ pending · ✅ green · ❌ red · ⚠️ flaky*

---

## Wave 0 Requirements

Existing infrastructure covers all phase requirements.

---

## Manual-Only Verifications

| Behavior | Requirement | Why Manual | Test Instructions |
|----------|-------------|------------|-------------------|
| CI pipeline passes on GitHub | CI-03, CI-04 | Requires GitHub Actions runner | Push branch, verify all checks pass |

---

## Validation Sign-Off

- [ ] All tasks have `<automated>` verify or Wave 0 dependencies
- [ ] Sampling continuity: no 3 consecutive tasks without automated verify
- [ ] Wave 0 covers all MISSING references
- [ ] No watch-mode flags
- [ ] Feedback latency < 15s
- [ ] `nyquist_compliant: true` set in frontmatter

**Approval:** pending
