# Claude Code VBA Skills

Toolkit of Claude Code configurations supporting VBA development, using skills, rules, and agent contexts. This project is modeled after [claude-code-r-skills](https://github.com/ab604/claude-code-r-skills) by Alistair Bailey, and adapted  for VBA-specific development.

## Features

### Skills (5)

| Skill | Description |
|-------|-------------|
| **vba-core-architecture** | Module structure, project setup, separation of concerns |
| **vba-error-events** | Error handling, event safety, debugging, logging |
| **vba-naming-and-data** | Naming conventions, typing, enums, data structures |
| **vba-object-model-and-performance** | Safe Office object model, arrays, batching |
| **vba-procedures-and-encapsulation** | Small procedures, scope rules, class patterns |

### Commands

| Command | Description |
|---------|-------------|
| `/plan` | Create implementation plans before coding |
| `/code-review` | Review code for security and quality |
| `/verify` | Full quality gate before committing |

### Rules

| Rule | Description |
|------|-------------|
| **security** | Credential handling, safe APIs, macro trust |
| **testing** | Compile checks, Debug.Assert, test harness patterns |
| **git-workflow** | Module export, conventional commits, branch naming |

### Agents

| Agent | Model | Description |
|-------|-------|-------------|
| **planner** | Opus | Implementation planning specialist |
| **code_reviewer** | Opus | Security and quality review |
| **testing-guide** | Opus | Testing strategies and test harness guidance |

### Contexts

System prompt contexts for launching Claude with a specific focus:

| Context | Description |
|---------|-------------|
| **dev** | Implementation mode — write VBA first, short responses |
| **research** | Exploration mode — read and understand before acting |
| **review** | Audit mode — find issues, rate severity, give verdict |

## Installation

### Option 1: Copy to your project

```bash
git clone https://github.com/your-username/claude-code-vba-skills.git

# Copy to your VBA project directory
cp -r claude-code-vba-skills/.claude/ /path/to/your/project/
cp -r claude-code-vba-skills/rules/ /path/to/your/project/
cp -r claude-code-vba-skills/commands/ /path/to/your/project/
cp -r claude-code-vba-skills/agents/ /path/to/your/project/
cp -r claude-code-vba-skills/contexts/ /path/to/your/project/
```

### Option 2: Copy to user configuration

```bash
mkdir -p ~/.claude/skills ~/.claude/rules ~/.claude/commands ~/.claude/agents ~/.claude/contexts

cp -r .claude/skills/* ~/.claude/skills/
cp -r rules/* ~/.claude/rules/
cp -r commands/* ~/.claude/commands/
cp -r agents/* ~/.claude/agents/
cp -r contexts/* ~/.claude/contexts/
```

### Local Settings

Create your own `.claude/settings.local.json` (gitignored) with preferred model and token settings:

```json
{
  "model": "sonnet",
  "env": {
    "MAX_THINKING_TOKENS": "10000",
    "CLAUDE_AUTOCOMPACT_PCT_OVERRIDE": "50"
  }
}
```

## Directory Structure

```
claude-code-vba-skills/
├── README.md
├── LICENSE
├── .gitignore
├── .claude/
│   ├── CLAUDE.md                              # Project instructions
│   ├── settings.json                          # Project-level settings (schema only)
│   ├── settings.local.json                    # Local preferences (gitignored)
│   └── skills/
│       ├── vba-core-architecture/
│       ├── vba-error-events/
│       ├── vba-naming-and-data/
│       ├── vba-object-model-and-performance/
│       └── vba-procedures-and-encapsulation/
├── contexts/
│   ├── dev.md                                 # Coding mode
│   ├── research.md                            # Exploration mode
│   └── review.md                              # Audit mode
├── rules/
│   ├── security.md
│   ├── testing.md
│   └── git-workflow.md
├── commands/
│   ├── plan.md
│   ├── code-review.md
│   └── verify.md
└── agents/
    ├── planner.md
    ├── code_reviewer.md
    └── testing-guide.md
```

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

### Key Anti-Patterns to Avoid

```vb
' Avoid unqualified references
Range("A1").Value = "Bad"

' Avoid Select/Activate
Selection.Value = 123

' Avoid blanket Resume Next
On Error Resume Next

' Avoid giant do-everything macros
```

## Core Principles

1. **Prefer clarity over cleverness** — readable code is maintainable code
2. **Validate assumptions early** — make invalid states hard to represent
3. **Minimize hidden dependencies** — avoid relying on active/selected objects
4. **Separate logic from Office UI** — keep business logic independent from worksheet I/O
5. **Optimize only after measuring** — most bottlenecks come from Excel object-model calls
6. **Always restore application state** — never leave events, calculation, or screen updating disabled
7. **Write for the next maintainer** — be explicit, structured, and predictable

## Token Optimization

| Setting | Default | Recommended | Impact |
|---------|---------|-------------|--------|
| `model` | opus | **sonnet** | ~60% cost reduction; handles 80%+ of coding tasks |
| `MAX_THINKING_TOKENS` | 31,999 | **10,000** | ~70% reduction in hidden thinking cost |
| `CLAUDE_AUTOCOMPACT_PCT_OVERRIDE` | 95 | **50** | Compacts earlier — better quality in long sessions |

Switch to Opus for deep architectural reasoning:
```
/model opus
```

## Recommended Workflow

```
contexts/research.md  → /plan → contexts/dev.md → implement → /verify → /code-review → commit
```

## License

MIT License — see [LICENSE](LICENSE) for details.
