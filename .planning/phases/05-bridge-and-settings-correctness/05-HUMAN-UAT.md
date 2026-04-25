---
status: partial
phase: 05-bridge-and-settings-correctness
source: [05-VERIFICATION.md]
started: 2026-04-25
updated: 2026-04-25
---

## Current Test

[awaiting human testing]

## Tests

### 1. Launch at Login End-to-End
expected: Toggle ON + Save creates Virelo.lnk in shell:Startup; toggle OFF + Save removes it; blocking creation shows error in status bar
result: [pending]

### 2. Theme System/Light/Dark Visual Behavior
expected: Theme applies visually immediately on selection (draft preview); System follows OS theme; Discard reverts to previously saved theme
result: [pending]

### 3. Key Capture + Dirty Indicator
expected: Dirty indicator appears after capture; old hotkey works until Save; new hotkey active after Save
result: [pending]

### 4. Accent/Density Persistence Across Restart
expected: After changing accent to teal + density to compact, saving, and restarting, values persist
result: [pending]

## Summary

total: 4
passed: 0
issues: 0
pending: 4
skipped: 0
blocked: 0

## Gaps
