# AGENTS.md

- `taskpane.*` is the Office.js client; `broker/` is the zero-dependency localhost bridge.
- `install/` owns Windows installation and rollback; `PROTOCOL.md` is the model-to-workbook contract.
- Never execute model-authored JavaScript or tool calls. Mutations stay structured, bounded, reviewable, and undoable.
- Never commit workbook data, client information, tokens, attachments, exports, or logs.
- Preserve token auth, Host/Origin checks, limits, CSV-injection protection, formula rebasing, verification, review-before-apply, rollback, and the single supervisor.

Run `npm test`, `npm run check:syntax`, `npm run check:install`, and—against a confirmed `hermes-excel-bridge` on a collision-free port—`npm run smoke`. Distinguish browser-only checks from live Excel/model verification. Disclose AI assistance and never invent results.
