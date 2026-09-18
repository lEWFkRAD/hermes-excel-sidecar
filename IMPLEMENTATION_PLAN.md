# Hermes Excel parity implementation plan

This branch is staging-only. It is not deployed to the live Home PC, Work Laptop, VPS gateway, or Excel bridge.

## Current slice

Implement and test the low-risk UI foundation without changing the gateway protocol:

- compact responsive header with overflow menu;
- workbook/sheet/selection scope indicator;
- structured job-state/status surface;
- multiline composer with `+` attachment affordance;
- bottom `Review edits` control;
- model/connector/skill controls reserved behind `+`;
- session ID and connection status remain visible in diagnostics, not as horizontal header clutter;
- accessible labels, focus states, narrow-pane layout, reduced motion.

## Later slices (not in this staging change)

1. Full-workbook context service with selection as target hint.
2. Versioned typed tool loop with bounded repair and checkpoints.
3. Cell citations and change-review drawer.
4. Real model routing and curated catalog.
5. Skills/connectors/secret approval UX.
6. Long-job pause/resume and session compaction/restore.
7. Native table/chart/pivot/error-tracing operations.
8. Canary deployment to Home PC, then Work Laptop.

## Safety rules

- Do not run deployment scripts.
- Do not edit live installed payloads.
- Do not restart or change either gateway.
- Preserve token, origin/host, size, formula, verification, review, rollback, and supervisor controls.
- Do not claim live Excel verification from browser/unit tests.

## Verification

Run after each vertical slice:

- `npm test`
- `npm run check:syntax`
- `npm run check:install` only on Windows (Linux staging reports PowerShell unavailable)
- `git diff --check`
- static secret/injection scan
- visual inspection at normal and narrow pane widths

Before any future deployment, require a fresh release bundle, rollback artifact, version/hash handshake, Home PC canary, and Work Laptop hash match.
