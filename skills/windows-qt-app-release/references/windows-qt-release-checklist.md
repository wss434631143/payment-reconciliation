# Windows Qt App Release Checklist

## 1. Requirements And Design

- Capture the user’s domain workflow in a repository document.
- List actors, inputs, outputs, calculations, data storage, import/export formats, and release expectations.
- For finance tools, document every money formula and the meaning of manually entered balances.
- Record data isolation rules, such as one SQLite database per store or tenant.

## 2. Qt Application Structure

- Keep UI code and persistence/business logic separate when possible.
- Use `QSplitter` for user-adjustable table/control regions.
- Use `QTableWidget` or model/view tables with pagination for large datasets.
- Use `QComboBox` for configured values, report types, months, stores, and filter values.
- Use business-specific Chinese button labels when the app is Chinese-facing.

## 3. Large Data UX

- Import: show progress text with processed rows and final imported/skipped counts.
- Query: avoid loading detail rows until the user selects a summary row or clicks query.
- Filter: collect multiple filter conditions first, then execute once.
- Export: stream XLSX writes and read database rows in batches.
- Status area: keep messages short, one-line when possible, and use tooltip for full paths.

## 4. SQLite Practices

- Create stable indexes for store, report type, month, transaction time, flow ID, and common filter columns.
- Store raw row payloads when different stores have different original columns.
- Separate master configuration from tenant/store data when data volume or schema differences matter.
- Provide backup/restore for all stores and selected stores.

## 5. UI Polish

- Use consistent spacing and button placement.
- Keep row actions near pagination when both operate on the same table.
- Let long filter summaries wrap or display in multiple lines.
- Avoid modal progress windows for long imports if they flash or block awkwardly; prefer a stable status/progress area.
- Use real icons and an application `.ico`.

## 6. Packaging

- Confirm dependencies are installed in a reproducible runtime.
- Build with PyInstaller `--windowed --onefile`.
- Include hidden imports for PySide6 modules when needed.
- Include an app icon with `--icon`.
- If overwrite fails, check whether the old exe is running or build with a versioned name.
- Create a release zip containing the exe and essential docs.

## 7. GitHub Publishing

- Update README, CHANGELOG, release notes, and requirements/design docs.
- Keep `.gitignore` excluding `data/`, `build/`, `dist/`, caches, logs, and SQLite files.
- Commit source and docs.
- Create or update a GitHub Release.
- Upload both a zip and an ASCII-named exe.
- Read back the Release asset list to confirm download URLs.

## 8. README Demo GIF

- Generate the demo from the actual app window.
- On Windows, use the native Qt platform for screenshots so Chinese fonts render.
- Use offscreen only for startup smoke tests.
- Keep the GIF lightweight and focused on the main workflow.
- If no real user data should appear, use a seeded sample database before recording.
