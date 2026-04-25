# Roadmap: Virelo

## Milestones

- **v1.0 Hardening and Cleanup** -- Phases 1-4 (shipped 2026-04-24)
- **v1.1 Polish and Correctness** -- Phases 5-8 (in progress)

## Phases

<details>
<summary>v1.0 Hardening and Cleanup (Phases 1-4) -- SHIPPED 2026-04-24</summary>

- [x] Phase 1: Hygiene and Build Pipeline (3/3 plans) -- completed 2026-04-24
- [x] Phase 2: Security and Bridge Hardening (3/3 plans) -- completed 2026-04-24
- [x] Phase 3: Structure and Quality (5/5 plans) -- completed 2026-04-24
- [x] Phase 4: Snap and Explorer Hardening (3/3 plans) -- completed 2026-04-24

Full details: `milestones/v1.0-ROADMAP.md`

</details>

### v1.1 Polish and Correctness (In Progress)

**Milestone Goal:** Make every visible UI control contextual, real, and backed by the same Python state model -- plus fix CI, repo hygiene, and window chrome.

- [x] **Phase 5: Bridge and Settings Correctness** - Python-owned dirty state, strict booleans, key capture via draft, theme coherence, and UI preference persistence -- completed 2026-04-25
- [x] **Phase 6: UI Action Placement** - Relocate controls to their correct pages, wire dead buttons, key capture UX, and command palette truthfulness -- completed 2026-04-25
- [x] **Phase 7: CI and Repo Hygiene** - Commit required assets, fix CI false positives and platform issues, sync versions, rename stale classes (completed 2026-04-25)
- [ ] **Phase 8: Window Chrome and Release** - Frameless window dragging, signed hit testing, smoke test, release verification, public documentation

## Phase Details

### Phase 5: Bridge and Settings Correctness
**Goal**: Every setting flows through a single Python-owned draft/commit model with correct types and coherent signals
**Depends on**: Phase 4 (v1.0 complete)
**Requirements**: BRDG-01, BRDG-02, BRDG-03, BRDG-04, BRDG-05, BRDG-06
**Success Criteria** (what must be TRUE):
  1. User sees a dirty indicator in the footer that updates immediately when any setting is changed, driven by a Python dirty_changed signal rather than frontend inference
  2. User can toggle Launch at login, save, and observe the startup shortcut created or removed -- with an error message if shortcut creation fails
  3. User can press a key capture button, press a new key, and see it reflected as a pending (unsaved) draft change with dirty indicator showing
  4. User can select System, Light, or Dark theme and the frontend updates correctly because Python emits both the chosen mode and the resolved effective theme
  5. Every boolean setting round-trips through the bridge without silent coercion (true/false strings, 1/0 integers all parse to strict Python bools)
**Plans:** 3 plans
Plans:
- [x] 05-01-PLAN.md -- Python settings model foundation (strict bool, new keys, dirty_changed signal)
- [x] 05-02-PLAN.md -- Bridge slot cleanup, key capture via draft, side effects expansion
- [x] 05-03-PLAN.md -- Frontend dirty state subscription, key mappings, theme/accent/density routing

### Phase 6: UI Action Placement
**Goal**: Every visible control lives on the correct page, does what it claims, and nothing fake remains in the interface
**Depends on**: Phase 5 (dirty state and key capture draft must exist for UI-02 and UI-05)
**Requirements**: UI-01, UI-02, UI-03, UI-04, UI-05, UI-06, UI-07
**Success Criteria** (what must be TRUE):
  1. User sees Test Snap only on the Window Snap page (in the Target Size card header) and in the command palette -- it does not appear in the global footer
  2. User can click Test Snap and it snaps to the currently visible (possibly unsaved) draft values, not the last-saved values
  3. User sees a footer containing only the status message, dirty indicator, Discard, and Save -- no Reset Defaults button; Reset Defaults lives on the General page behind a confirmation dialog
  4. User can rebind snap and restore keys using press-to-capture controls instead of segmented SHIFT/CTRL/ALT selectors, with the new binding shown as a dirty draft change
  5. User sees only real, implemented actions in the command palette and accurate descriptions on the Shortcuts page
**Plans:** 3 plans
Plans:
- [x] 06-01-PLAN.md -- Footer cleanup and Test Snap relocation to Target Size card header
- [x] 06-02-PLAN.md -- Key capture controls and Reset confirmation dialog
- [x] 06-03-PLAN.md -- Command palette truthfulness and Shortcuts page accuracy
**UI hint**: yes

### Phase 7: CI and Repo Hygiene
**Goal**: A developer can clone the repo and build, test, and lint successfully on the first try
**Depends on**: Phase 5 (CI-08 renames classes introduced or modified in bridge work)
**Requirements**: CI-01, CI-02, CI-03, CI-04, CI-05, CI-06, CI-07, CI-08
**Success Criteria** (what must be TRUE):
  1. Developer can clone the repo on a fresh machine and run the build pipeline without missing any committed assets (icon.ico, branding files, frontend/index.html)
  2. Developer can run the clean script and it removes all generated artifacts (including __pycache__, *.pyc, .pytest_cache, .ruff_cache) recursively
  3. CI stale-name check passes without false positives on its own workflow file or test docstrings, and unit tests pass on the CI platform
  4. Frontend package.json version, config.py APP_VERSION, installer support URL, and README Python version requirement are all consistent and from a single source of truth
  5. Stale class names (ShiftSnapRestore, HotkeyListener) are renamed to reflect that keys are now configurable (SnapRestoreController, MultiPressHotkeyListener)
**Plans:** 4/4 plans complete
Plans:
- [x] 07-01-PLAN.md -- Commit source assets (icon.ico, branding/*, frontend/index.html) and expand clean script
- [x] 07-02-PLAN.md -- Fix CI stale-name false positive, add version-check job, add platform guard for Win32 test
- [x] 07-03-PLAN.md -- Sync frontend/package.json version, fix README Python version, update installer URL
- [x] 07-04-PLAN.md -- Rename HotkeyListener and ShiftSnapRestore throughout codebase

### Phase 8: Window Chrome and Release
**Goal**: The application window behaves correctly on all monitor configurations and the release pipeline produces a verified, documented artifact
**Depends on**: Phase 6 (chrome polish applies to the UI surface delivered in Phase 6), Phase 7 (release verification depends on clean CI)
**Requirements**: CHRM-01, CHRM-02, CHRM-03, CHRM-04, CHRM-05, CHRM-06, CHRM-07
**Success Criteria** (what must be TRUE):
  1. User can drag the frameless window by clicking and dragging non-interactive title bar areas, using a Python-side event filter for reliable hit testing
  2. User can resize the window correctly even on monitors with negative screen coordinates (signed 16-bit lParam decoding)
  3. Developer can run a non-interactive smoke test (--smoke-test) that verifies resource paths, frontend dist, QWebEngine init, settings, and bridge init without manual interaction
  4. Developer can run release verification that checks version consistency, asset existence, and built content correctness -- and .planning/ is gitignored while public docs exist in README.md and docs/
  5. Public documentation covers build instructions, troubleshooting, and release checklist with the correct product name and accurate descriptions
**Plans**: TBD
**UI hint**: yes

## Progress

**Execution Order:**
Phases execute in numeric order: 5 -> 6 -> 7 -> 8

| Phase | Milestone | Plans Complete | Status | Completed |
|-------|-----------|----------------|--------|-----------|
| 1. Hygiene and Build Pipeline | v1.0 | 3/3 | Complete | 2026-04-24 |
| 2. Security and Bridge Hardening | v1.0 | 3/3 | Complete | 2026-04-24 |
| 3. Structure and Quality | v1.0 | 5/5 | Complete | 2026-04-24 |
| 4. Snap and Explorer Hardening | v1.0 | 3/3 | Complete | 2026-04-24 |
| 5. Bridge and Settings Correctness | v1.1 | 3/3 | Complete | 2026-04-25 |
| 6. UI Action Placement | v1.1 | 3/3 | Complete | 2026-04-25 |
| 7. CI and Repo Hygiene | v1.1 | 4/4 | Complete   | 2026-04-25 |
| 8. Window Chrome and Release | v1.1 | 0/0 | Not started | - |
