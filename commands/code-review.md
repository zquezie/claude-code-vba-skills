---
name: code-review
description: Review VBA code for security, quality, and best practices. Uses the vba_code_reviewer agent.
---

# /code-review - VBA Code Quality Review

Review uncommitted code changes for security vulnerabilities, code quality issues, and VBA best practices.

## Process

1. Identify changed files (`git diff --name-only HEAD`)
2. Run automated scans (Option Explicit, Select/Activate, error handling, state restoration)
3. Review each file against quality criteria
4. Generate severity-rated report
5. Issue verdict: APPROVE / NEEDS CHANGES / BLOCK

## Severity Levels

| Level | Action | Examples |
|-------|--------|----------|
| **CRITICAL** | Must fix | Hardcoded secrets, unsafe Shell calls |
| **HIGH** | Should fix | Missing error handling, no Option Explicit |
| **MEDIUM** | Fix when possible | Style issues, unqualified references |
| **LOW** | Consider | Naming suggestions, documentation |

## Usage

```
/code-review
```

Review will automatically check all uncommitted changes.
