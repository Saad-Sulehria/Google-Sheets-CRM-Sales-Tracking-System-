# Repository Guidelines

## Project Structure & Module Organization
- Root-level `.gs.js`/`.js` files contain the Google Apps Script code (e.g., `UpdateVendoBlock.js`, `Menu.gs.js`, `createNewVendor.gs.js`).
- `tests/` holds Jest tests and mocks (e.g., `tests/UpdateVendoBlock.test.js`, `tests/gas-mocks.js`).
- `Specifications/` and docs (`ARCHITECTURE.md`, `TECHNICAL_DOCS.md`, `USER_GUIDE.md`, `instruction.md`) capture requirements and system behavior.
- `appsscript.json` is the Apps Script manifest; `.clasp.json` configures CLASP sync.

## Build, Test, and Development Commands
- `npm test` — run Jest test suite locally.
- `clasp push --force` — push local Apps Script files to the linked Google project (see `.clasp.json`).
- `clasp open` — open the Apps Script project in the browser.

## Coding Style & Naming Conventions
- Use 2-space indentation and double quotes (matches existing style).
- Prefer `const`/`let` and plain function declarations for Apps Script.
- Keep Apps Script files in the repository root and name tests `*.test.js`.
- No formatter or linter is configured; keep changes minimal and consistent with nearby code.

## Testing Guidelines
- Framework: Jest (see `package.json`).
- Tests live in `tests/` and use `gas-mocks.js` for Apps Script stubs.
- Run all tests with `npm test`; no explicit coverage requirement is defined.

## Commit & Pull Request Guidelines
- Git history is not available in this directory, so no project-specific commit convention could be detected.
- Suggested default: short, imperative subject lines (e.g., “Fix vendor header detection”), include scope if helpful.
- PRs should describe changes, list tests run, and link relevant docs/issues; include screenshots when UI/Sheets behavior changes.

## Configuration & Docs Tips
- Do not change `.clasp.json` script IDs unless intentionally relinking the project.
- For architecture and deployment context, start with `ARCHITECTURE.md` and `TECHNICAL_DOCS.md`.
