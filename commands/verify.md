---
name: verify
description: Run full VBA verification loop before committing. Checks compile readiness, anti-patterns, error handling, and code quality.
---

# /verify - VBA Verification Loop

Run all quality gates before committing code. Produces a READY or NOT READY verdict.

## When to Use

- Before every commit
- After implementing a feature or fix
- Before distributing a workbook or add-in

## Verification Phases

Run each phase in order. A failure in any phase blocks the verdict.

### Phase 1: Option Explicit Check

Scan all modules for `Option Explicit`. Every module must have it.

### Phase 2: Anti-Pattern Scan

Search for known anti-patterns:
- `Select` / `Activate` / `Selection` / `ActiveSheet` / `ActiveWorkbook`
- Blanket `On Error Resume Next` (without targeted scope)
- Unqualified `Range()` / `Cells()` references
- `Integer` declarations (should be `Long`)

### Phase 3: Error Handling & State Restoration

Verify that:
- Public procedures use structured error handling (`CleanExit` / `CleanFail`)
- `Application.ScreenUpdating`, `.Calculation`, `.EnableEvents` are restored on exit and error
- No procedures leave application state modified on failure

### Phase 4: Code Review

Invoke `/code-review` to check all uncommitted changes.

### Phase 5: Diff Review

```bash
git diff --stat HEAD
git diff HEAD
```

Confirm:
- No unintended files changed
- No debug code left in (`Debug.Print`, `Stop`, `MsgBox` for debugging)
- No hardcoded paths or credentials
- Commit message ready

## Verdict Format

```
## Verification Report

| Phase                | Status  | Notes                        |
|----------------------|---------|------------------------------|
| Option Explicit      | PASS    | All 12 modules               |
| Anti-pattern scan    | PASS    | 0 findings                   |
| Error handling       | PASS    | All public procedures covered |
| Code review          | PASS    | No CRITICAL/HIGH issues      |
| Diff review          | PASS    | Changes look correct         |

## Verdict: READY TO COMMIT
```

## Related Commands

- `/plan` — plan before implementing
- `/code-review` — standalone code review
