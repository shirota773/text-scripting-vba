text-scripting-vba
==================

Forked and maintained for Excel workbook/add-in workflows.

Overview
--------
This project lets you edit VBA components as text files and reload them into Excel.
Supported component types:
- Standard Module (`.bas`)
- Class Module (`.cls`)
- UserForm (`.frm`, with paired `.frx`)

Fork Note
---------
This repository is a fork/customized continuation of the original project:
- Original page: http://rsh.csh.sh/text-scripting-vba/

Safety-first Defaults
---------------------
All safety options are enabled by default in `src/ThisWorkbook.cls`:
- Managed-only update: only components listed in `libdef.txt` are replaced.
- Preflight validation: checks files/extensions/duplicates before apply.
- Backup snapshot: exports all removable components before changes.
- Rollback on failure: restores backup automatically if import fails.
- Form pair check: `.frm` requires paired `.frx`.

Requirements
------------
In Excel, enable:
1. `Developer` tab
2. `Trust access to the VBA project object model`

Usage
-----
1. Place workbook/add-in and `libdef.txt` in the same folder.
2. List managed component files in `libdef.txt`.
3. Open workbook/add-in and run `reloadComponents`.
4. Run `exportComponents` to export current modules/classes/forms to files.
5. If a component is not listed in `libdef.txt`, a Yes/No/Cancel dialog asks whether to add it.

Release Mode
------------
- `enableReleaseMode`: disables reload operations for safer distribution.
- On workbook close, a prompt asks whether to switch to release mode.
- If release mode is already ON, the close prompt is not shown.
- Release mode state is stored in a hidden workbook name.

Operation Log
---------------------
- `reload-log.txt` is appended in the workbook folder when `reloadComponents` runs.
- It records timestamp, status, component snapshot, and details.

Notes
-----
- Auto reload on open is disabled by default (`ENABLE_WORKBOOK_OPEN = False`).
- Editing VBProject may invalidate VBA digital signatures.
- Keep source text English-only to avoid mojibake issues on Windows Excel.
