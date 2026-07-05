---
name: windows-qt-app-release
description: Build, polish, package, document, and publish Windows desktop applications built with Python Qt/PySide6 or similar Qt stacks. Use when Codex is asked to create or continue a Windows Qt app, optimize table-heavy desktop UI performance, configure reusable UI patterns, package an exe with PyInstaller, create GitHub repository documentation, generate a real-app README GIF, or upload source code and Release assets to GitHub.
---

# Windows Qt App Release

## Workflow

Use this workflow for Windows Qt desktop apps, especially data-heavy business tools.

1. Clarify the business workflow and write it down before publishing.
2. Inspect the existing project structure, data files, packaging files, and README.
3. Implement UI and backend changes using the project’s existing patterns.
4. Validate with syntax checks, Qt startup checks, and targeted workflow tests.
5. Generate real screenshots or GIFs from the actual app, not mockups.
6. Package with PyInstaller and verify the output path, icon, app name, and data folder behavior.
7. Prepare GitHub documentation and Release notes.
8. Upload source files and Release assets, excluding local databases, build caches, and user data.

For detailed checklists, read `references/windows-qt-release-checklist.md`.

## UI Rules

- Prefer store-first or entity-first layouts for operational finance tools.
- Keep data tables dense, readable, and stable; avoid marketing-style panels.
- Use splitters where users need to resize table/control/status regions.
- Put pagination and row actions on the same row when they serve the same table.
- Let long filter summaries wrap or move to their own row.
- Use explicit action labels in Chinese for Chinese business apps, such as `开始导入`, `保存设置`, `退出`.
- For large tables, page data by default and avoid refreshing on every intermediate filter edit.

## Performance Rules

- Never load full large-month detail data into a UI table by default.
- Use database-level filtering, sorting, counts, and pagination.
- Cache expensive distinct-value lists for filter dropdowns.
- Batch import/export operations and show progress in the status area.
- During store or report-type switches, clear stale data first and load only after the user queries.
- Disable table updates/sorting while repopulating rows, then restore them.

## Packaging Rules

- Use PyInstaller with `--windowed`, `--onefile`, a stable app name, and an `.ico`.
- Prefer an ASCII Release asset name for GitHub downloads even when the local exe has a Chinese display name.
- Keep generated `.spec`, `build/`, `dist/`, `data/`, `__pycache__/`, and SQLite files out of source unless the user explicitly wants them.
- Package docs with the zip: executable, usage guide, and release notes.

## GitHub Release Rules

- Do not commit local user data or store databases.
- Update README, CHANGELOG, release notes, and design/requirements docs before release.
- For README GIFs, launch the actual app and capture the real UI. Do not draw an illustrative GIF unless the user explicitly asks for a mockup.
- If GitHub mangles non-ASCII asset names, upload a duplicate asset with an ASCII file name.
- Verify final Release asset list and links after upload.

## Validation

Run at least:

```powershell
python -B -c "import ast,pathlib; ast.parse(pathlib.Path('qt_app.py').read_text(encoding='utf-8')); print('AST_OK')"
```

For Qt startup checks, set `QT_QPA_PLATFORM=offscreen` when visual inspection is not needed. For real screenshots or GIFs, use the native Windows platform so Chinese fonts render correctly.
