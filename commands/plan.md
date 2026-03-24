---
name: plan
description: Create an implementation plan before coding. Uses the vba_planner agent for structured analysis.
---

# /plan - VBA Implementation Planning

Create a detailed implementation plan before writing any code.

## When to Use

- New features or functionality
- Architectural changes
- Complex refactoring
- Multi-module modifications
- Requirements that need clarification

## Process

1. Pass your feature description to the `vba_planner` agent
2. The planner performs codebase reconnaissance (scans for patterns, risks, dependencies)
3. Produces a structured plan with phased steps, testing strategy, and risk mitigation
4. Review and confirm before implementation begins

## Usage

```
/plan Add a data validation layer for user input worksheets
```

## After Receiving the Plan

- **Confirm**: "yes" or "proceed" to start implementation
- **Modify**: "modify: [changes]" to adjust the plan
- **Alternative**: "alternative: [approach]" to explore a different direction

**The planner will NOT write any code until you explicitly confirm the plan.**
