# VBA Development Guide

*This document captures high‑level best practices for modern, maintainable VBA development. Detailed examples and implementation patterns are provided elsewhere.*

## Core Principles

1. **Prefer clarity over cleverness** – readable code is maintainable code.
2. **Validate assumptions early** – make invalid states hard to represent.
3. **Minimize hidden dependencies** – avoid relying on active/selected objects.
4. **Separate logic from Office UI** – keep business logic independent from worksheet I/O.
5. **Optimize only after measuring** – most bottlenecks come from Excel object‑model calls.
6. **Always restore application state** – never leave events, calculation, or screen updating disabled.
7. **Write for the next maintainer** – be explicit, structured, and predictable.

## Project & Environment Setup

- Enable `Option Explicit` and consistent indentation.
- Organize modules by responsibility (entry points, utilities, domain classes, I/O logic).
- Use the correct module type: standard modules, class modules, sheet/workbook modules, UserForms.

## VBA Development Skill Areas

| Topic | Skill File | Description |
|-------|------------|-------------|
| **Core Architecture** | [vba-core-architecture](skills/vba-core-architecture/SKILL.md) | architecture, module structure, and project setup |
| **Error Handling** | [vba-error-events](skills/vba-error-events/SKILL.md) | robust error handling, event safety, debugging, logging |
| **Naming and Data** | [vba-naming-and-data](skills/vba-naming-and-data/SKILL.md) | naming, variable typing, enums, data structures, and clarity |
| **Object Model and Performance** | [vba-object-model-and-performance](skills/vba-object-model-and-performance/SKILL.md) | Safe Office object model usage, speed scalability, array and batching operations |
| **Procedures and Encapsulation** | [vba-procedures-and-encapsulation](skills/vba-procedures-and-encapsulation/SKILL.md) | patterns for building small, readable, encapsulated procedures |


## Contexts

Launch Claude with a specific focus using `--system-prompt`:

| Context | Description |
|---------|-------------|
| **dev** | Implementation mode — write VBA first, short responses |
| **research** | Exploration mode — read and understand before acting |
| **review** | Audit mode — find issues, rate severity, give verdict |

## Rules

Always-active constraints enforced throughout the session:

| Rule | Description |
|------|-------------|
| **security** | Credential handling, safe APIs, macro trust |
| **testing** | Compile checks, Debug.Assert, test harness patterns |
| **git-workflow** | Module export, conventional commits, branch naming |

## Commands

| Command | Description |
|---------|-------------|
| `/plan` | Create implementation plans before coding |
| `/code-review` | Review code for security and quality |
| `/verify` | Full quality gate before committing |

## Agents

| Agent | Model | Description |
|-------|-------|-------------|
| **planner** | Opus | Implementation planning specialist |
| **code_reviewer** | Opus | Security and quality review |
| **testing-guide** | Opus | Testing strategies and test harness guidance |

## Quick Reference

### Essential Modern VBA Patterns

```vb
' Fully qualify references
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Data")

' Use arrays for performance
Dim arr As Variant
arr = ws.Range("A1:A1000").Value2

' Guard clauses
If ws Is Nothing Then Exit Sub

' Structured error handling
On Error GoTo CleanFail
```

### Key Anti‑Patterns to Avoid

```vb
' Avoid unqualified references
Range("A1").Value = "Bad"

' Avoid Select/Activate
Selection.Value = 123

' Avoid blanket Resume Next
On Error Resume Next

' Avoid giant do-everything macros
```

## Checklist

### VBA That Meets Professional Standards

- Clear naming, consistent formatting
- Fully qualified references
- Short, single‑responsibility procedures
- Structured error handling
- Array‑based worksheet interaction
- Encapsulated logic using classes
- Documented assumptions and deployment requirements
- No reliance on Selection/ActiveSheet
- Safe cleanup restoring application state
