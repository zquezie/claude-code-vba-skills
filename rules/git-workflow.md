# VBA Git Workflow Rules

VBA code lives inside binary Office files (`.xlsm`, `.xlam`, `.docm`). These rules ensure VBA projects remain version-controllable.

## Module Export (Mandatory)

- Export all VBA modules to text files (`.bas`, `.cls`, `.frm`) before committing
- Keep exported files in a `src/` directory as the source of truth
- Never commit binary Office files without also committing exported text modules

```
src/
├── modEntryPoints.bas
├── modConfig.bas
├── modUtilities.bas
├── clsApplication.cls
├── clsLogger.cls
├── Sheet1.cls
└── ThisWorkbook.cls
```

## Commit Messages

Use conventional commit format:

```
feat: add invoice generation workflow
fix: correct date parsing for non-US locales
refactor: extract validation into clsValidator
docs: update deployment instructions
chore: export modules after refactor
```

## Branch Naming

```
feature/invoice-generator
fix/date-parsing-locale
refactor/extract-validator
docs/deployment-guide
```

## Pre-Commit Checklist

Before committing:
- [ ] Project compiles without errors (Debug > Compile VBAProject)
- [ ] All modules exported to `src/`
- [ ] No hardcoded credentials or sensitive data
- [ ] Commit message follows conventional format
- [ ] No Office temp files (`~$*`) in staging

## Files to Commit

- `src/*.bas`, `src/*.cls`, `src/*.frm` — exported VBA source
- Documentation (`.md` files)
- Configuration files

## Files to Never Commit

- Office temp files (`~$*.xl*`, `~$*.doc*`)
- `.claude/settings.local.json`
- Binary Office files without corresponding exported source (when practical)
- Files containing credentials or PII

## Tools

Consider these tools for automated export/import:
- [vbaDeveloper](https://github.com/nicetl/vbaDeveloper) — automatic export on save
- [Rubberduck VBA](https://rubberduckvba.com/) — source control integration
- Manual export via VBE: File > Export File
