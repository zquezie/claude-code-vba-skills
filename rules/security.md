# VBA Security Rules

These rules are always active. Claude must enforce them throughout the session.

## Credential & Secret Management

- **Never** hardcode API keys, passwords, tokens, or connection strings in VBA modules
- **Never** store secrets in worksheet cells, named ranges, or document properties
- Use environment variables or encrypted configuration files for secrets
- If credentials must be collected, use a UserForm with password masking — never `InputBox`

## File & System Safety

- Validate all file paths before using `Open`, `Kill`, `FileCopy`, `MkDir`, or `RmDir`
- Never pass user-provided strings directly to `Shell` without validation
- Avoid `SendKeys` — it is fragile and can execute unintended commands
- Be cautious with `SaveSetting` / `GetSetting` (registry access)
- Never write to or read from arbitrary registry keys

## Windows API Declarations

- All `Declare` statements must use `PtrSafe` for 64-bit compatibility
- Use `LongPtr` for handles and pointers, not `Long`
- Use conditional compilation (`#If VBA7 Then`) when supporting both 32-bit and 64-bit

## Macro Trust & Distribution

- Document required Trust Center settings for end users
- Prefer code signing (`.pfx` certificates) for distributed macros
- Never programmatically disable macro security
- Never instruct users to lower security settings

## Data Protection

- Do not log or `Debug.Print` sensitive data (PII, credentials, financial data)
- Sanitize data before writing to log sheets or external files
- Be cautious with `CreateObject("MSXML2.XMLHTTP")` — validate URLs and handle errors

## Response Protocol

When security issues are discovered:
1. **STOP** — halt current work
2. **ASSESS** — determine scope and severity
3. **FIX** — address critical vulnerabilities first
4. **REVIEW** — check codebase for similar issues

**Security vulnerabilities are blocking issues. No code with security vulnerabilities should be committed.**
