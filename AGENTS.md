# AsiToPix

## Purpose

This project prepares directory trees containing symbolic links to
astrophotography files for processing in PixInsight.

Main scripts:

- `CreateProject.ps1` — creates a PixInsight processing project tree with symlinks to original files, creates project_meta.json with metadata.
- `ExportMasters.ps1` — copies integration masters to the Calibration directory (untested, TODO).
- `CombineSeasons.ps1` — combines data collected during multiple seasons into a single project tree.

## Environment

- Target platform: Windows.
- Shell: PowerShell 5 and higher.
- Source files may reside on different Windows volumes, including network drives.
- Paths may contain spaces, Unicode characters, but not Cyrillic characters.
- Night dates use the local noon-to-noon convention and the `yy.MM.dd`
  label of the night start date. Captures before 12:00 belong to the
  previous calendar date; captures at or after 12:00 belong to the current
  date. Apply this convention to lights and calibration frames.
- The scripts create symbolic links and must preserve the original files.
- Never modify, rename, move, or delete source astrophotography files.

## Development rules

- Read `TODO.md`, patch it after refactoring or bug fixes.
- All commentaries in the code in English.
- Interactive Y/N prompts must accept `yY/nN/дД/нН` variants.
- Preserve existing behavior unless the task explicitly changes it.
- Prefer `Join-Path`, `Resolve-Path`, and `System.IO.Path` over manual
  path concatenation.
- Avoid hard-coded user-specific paths.
- Use `[CmdletBinding()]` and explicit `param()` blocks for executable scripts.
- Use `Set-StrictMode -Version Latest`.
- Use `$ErrorActionPreference = 'Stop'` in executable entry points.
- Validate parameters before modifying the output directory.
- Support `-WhatIf` for operations that create, replace, or remove files.
- Produce actionable error messages containing the affected path.
- Use approved PowerShell verbs for function names.
- Put reusable functions in module files under `src/`.
- Add Pester tests under `tests/`.

## Safety requirements

- Treat source directories as read-only.
- Never recursively delete a directory unless it was created by this project
  and its resolved path was validated.
- Before replacing an existing link, verify that it is a symbolic link.
- Do not overwrite ordinary files.
- Do not silently ignore failed link creation.
- Do not introduce administrator requirements without documenting them.

## Verification

For every code change:

1. Parse all PowerShell files to detect syntax errors.
2. Run PSScriptAnalyzer when available.
3. Run Pester tests when available.
4. Show the resulting Git diff.
5. State any behavior that could not be tested locally.
