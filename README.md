# Cyberville Offline Excel

A fully offline spreadsheet PWA that opens, edits and re-saves `.xlsx` files **without losing formatting, images, charts or comments**. Built by [Cyberville](https://cyberville.tech), Juba.

## Features

- **Works everywhere, works offline** — installable PWA (Windows, macOS, Android, iOS). Once loaded, the whole app runs from the service-worker cache with zero internet.
- **Full-fidelity save** — edits are patched surgically into the original Excel file, preserving cell styles, merged cells, images, charts and comments. No SheetJS re-serialize, no formatting loss.
- **No account, no server, no tracking** — your data never leaves the device.
- Import: `.xlsx`, `.xls`, `.xlsm`, `.csv` · Export: `.xlsx`, `.csv`
- Formulas, undo/redo, auto-save to your last session.

## Two ways to use it

1. **Online / installable (recommended):** open the Vercel URL, click **Install**. After the first load it works with no connection at all.
2. **Fully local:** `index.html` is a single self-contained file — copy it (and the optional `lib/` folder) anywhere and double-click to run. It also ships as `ExcelOffline-Windows.zip` for Windows users.

## Source layout

- `index.html` — the app (self-contained build; source lives with external `lib/` references in the sibling `ExcelOffline/` workspace)
- `lib/` — SheetJS CE (`xlsx.full.min.js`) and the surgical OOXML engine (`ooxml.js`)
- `sw.js` — offline service worker (stale-while-revalidate + offline navigation fallback)
- `manifest.webmanifest`, `icon-*.png` — PWA install metadata and icons
- `tests/` — headless Node test suites (see below)

## Tests

```bash
./tests/run-tests.sh
```

- `ooxml-test.js` — zip/XML surgical engine (27 checks)
- `spreadsheet-test.js` — core edit/serialize behavior (19 checks)
- `fidelity-test.js` — end-to-end "open → edit → re-save preserves formatting" (30 checks)

## Deploy

Static site, zero config:

```bash
vercel --prod
```

or import this repository in the Vercel dashboard (framework preset: **Other/Static**).

## Legal

Cyberville-branded app for internal and client use. SheetJS Community Edition is Apache-2.0. All processing happens locally in the browser.