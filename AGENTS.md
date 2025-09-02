# Repository Guidelines

## Project Structure & Module Organization
- Root: `vbac.wsf` (JScript/WSF tool to import/export VBA code), `README.md`, `LICENSE.txt`.
- Source: `src/` holds VBA modules (`.bas`), classes (`.cls`), forms (`.frm` + `.frx`), and document modules (`.dcm`). Optional: `src/App.vbaproj` for references/metadata.
- Binaries: `bin/` contains macro-enabled Office files (e.g., `.xlsm`, `.docm`, `.accdb`). Created/updated by commands below.
- Backups: `bak/` is auto-created when combining (timestamped copies of `bin/`).

## Build, Test, and Development Commands
- Export code from binaries to `src/`:
  - `cscript //nologo vbac.wsf decombine`
- Import code from `src/` into `bin/` (build artifacts):
  - `cscript //nologo vbac.wsf combine`
- Clear VBA components in binaries (use with care):
  - `cscript //nologo vbac.wsf clear`
- Useful options: `/binary:bin`, `/source:src`, `/vbaproj` (use project file), `/dbcompact` (Access only).
  - Example: `cscript //nologo vbac.wsf combine /source:src /binary:bin /vbaproj`

## Coding Style & Naming Conventions
- VBA style: 4-space indent; `Option Explicit` required; avoid implicit variants.
- Naming: PascalCase for procedures/types; camelCase for locals; UPPER_SNAKE for constants.
- Modules: prefix standard modules with `Mod` (e.g., `ModStrings`), classes with `Cls` (e.g., `ClsTokenizer`).
- One public type per file; file extensions must match component type.

## Testing Guidelines
- No automated test harness yet; verify manually in Office:
  - Open the artifact in `bin/`, run key entry points, check immediate window output and form behavior.
- Keep changes small and reversible; use `decombine` to review diffs in `src/`.

## Commit & Pull Request Guidelines
- Commits: imperative present tense, focused scope.
  - Example: `Add tokenizer for string literals`.
- PRs: include summary, scope, affected Office app(s), and manual test notes/screenshots (for forms/Access objects). Link related issues.
- Ensure `decombine` output is clean (no stray generated files) and that `combine` succeeds locally.

## Security & Configuration Tips
- In Office, enable “Trust access to the VBA project object model” and allow macros for development. Close Office apps before running scripts.
- Windows-only tooling: commands require `cscript.exe`. Paths can be overridden with `/binary:` and `/source:` as needed.

