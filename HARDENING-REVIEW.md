# Excel sidecar hardening and gap analysis — 2026-09-21

AI-assisted review and implementation by Codex. Source: upstream main
`6f664ab95d03b616021eb2de4cc6653b22188748`, fetched September 21. Local branch:
`fix/excel-compatibility-hardening-20260921`. This report records the source
review and validation for publication. No release, deployment, runtime patch,
service restart or model-route change was performed.

## Implemented repairs

| Finding | Repair and evidence |
| --- | --- |
| Newer Office APIs were invoked without checking the host; later unsupported steps could follow successful writes | Declare/check ExcelApi 1.1, centralize optional action requirements, and preflight complete immediate/reviewed action lists before work. Tests reject a write followed by each unsupported action without touching Excel |
| Basic context/sizing and sheet lookups unnecessarily required newer APIs | Absolute range helper uses 1.1 cell/bounding APIs; pre-1.4 hosts use ordinary used ranges and collection-based sheet lookup. Tests cover non-A1 dimensions, large selections, empty sheets and deletion |
| Every new sheet used autofit and freeze panes | Presentation checks 1.2/1.7 support independently; explicit requests remain fail-closed. Tests omit unsupported members entirely from old-host mocks |
| Older modern webviews may lack AbortSignal.timeout | Polling uses AbortController plus a cleared timer, including stalled JSON bodies; timeout/failure regressions pass |
| Undo barriers were added only after successful non-reversible operations | Failed/uncertain formatting, structural actions, sheet creation and reviewed commit/verification preserve a barrier, preventing an earlier record from being undone unexpectedly |
| Unsaved workbooks shared `hermes-chat:default` | Persist/load history only for a stable document URL; unnamed workbooks keep history in memory. Regression demonstrates zero storage reads/writes for both missing/default URLs |
| Product support was stated broadly without qualifications | Added API-based support matrix, browser startup guidance and host qualification checklist; corrected the README's implication that model-authored Office.js executes |

These are focused refactors of range sizing, capability policy, timeout handling
and history identity. The structured action protocol, authentication,
Host/Origin protection, upload limits, formula rebasing and immediate-mode
default remain intact. Capability hints reach the typed adapter so it can avoid
unsupported actions, while the pane enforces actual host support.

## Validation

- Original baseline: 96 Node tests passed.
- Final `npm run validate`: **passed**, with 130 Node behavior tests (50
  task-pane tests), 4 package/release tests, and 93 Python tests; syntax,
  installer/rollback and release metadata checks passed. `git diff --check`
  passed. The final run is recorded in the workspace validation log.
- Installed Hermes Python runtime: 93 tests passed without skips.
- API-tier mocks: 1.1, 1.2, 1.4, 1.6, 1.7, 1.12 and 1.20. These exercise the
  changed source, not physical Office products.
- Live **existing deployment**: authenticated identity confirmed
  `hermes-excel-bridge` on 8788; bridge, typed adapter and parser reported ready;
  raw fallback was disabled. Synthetic smoke: **10 passed, 1 failed, 6 skipped**.
  The simple model request returned `source=fallback` with zero actions, so
  anchor/history/medium-model cases were skipped. This does not identify the
  provider root cause. Health readiness is not sufficient release evidence.
- Live security/export checks passed. Manual Host-header, true model-down and
  cross-service containment checks were skipped by the smoke harness. TLS
  verification was bypassed only in the localhost test process; certificate
  trust was not certified. No live workbook was edited. Smoke created synthetic
  export artifacts through the existing endpoint.
- Local logs outside the repository: `../excel-compat-validation.log`,
  `../excel-compat-python.log`, `../excel-compat-live-smoke.log`.
- The configured Onyx helper was attempted for bounded read-only review but
  returned no usable text after one retry; findings were verified directly.

## Remaining gaps, in priority order

1. **P1 — deployed typed chat readiness:** diagnose the fallback response and
   rerun synthetic live smoke before any promotion. Keep service/config changes
   proposal-only under the shared wiki rules. New pane changes do not repair
   an unavailable or misconfigured inference path.
2. **P1 — real Excel qualification:** run the matrix in
   [EXCEL-COMPATIBILITY.md](EXCEL-COMPATIBILITY.md). The implementation broadens
   API compatibility; it does not certify Excel 2016/2019/2021/2024, Mac, web,
   32/64-bit, legacy file formats or localized worksheet functions.
3. **P1 — transactional/concurrent-edit guarantees:** `applyReviewedProposal`
   still uses optimistic snapshots and best-effort restoration. Office sync is
   not a database transaction. Exact duplicate addresses are checked, but
   differently spelled or partially overlapping ranges need explicit overlap
   analysis and real-host tests. Restoration amid coauthor edits and sheet-name
   reuse requires a separate design; the new Undo barriers do not solve that.
4. **P2 — broader Undo coverage:** created sheets, formatting and structural
   edits remain deliberately non-undoable barriers. Full undo requires format,
   merge, conditional-rule and structural snapshots plus conflict checks, not
   simply removing barriers.
5. **P2 — memory/load bounds:** 100 MiB files and the 210 MiB JSON envelope can
   produce multiple in-memory copies. Profile peak RSS and concurrent requests
   before increasing capacity. Worksheet context also walks all sheets for
   metadata; large-sheet-count workbooks need a capped discovery/pagination
   design even though cell samples are bounded.
6. **P2 — installation and browser scope:** Windows-only installation needs
   explicit Mac tooling and web localhost/TLS qualification. Supporting IE or
   legacy Edge would require a transpilation/polyfill/security maintenance
   commitment; the present change intentionally reports them as unsupported.
7. **P2 — formula compatibility and privacy controls:** API versions cannot
   establish worksheet-function support. Add a traditional-formula profile and
   function probes where supported. Named-workbook chat still persists by URL;
   storage retention and an opt-out deserve a separate user-facing policy.

## Next refactor boundary

After host qualification, extract the capability/range helpers and workbook
executor into independently tested modules, followed by separate review/Undo
and chat/session modules. Preserve the static-file allowlist and archive/install
manifest while doing so. Splitting the large pane without those package tests
would introduce a new deployment risk for little immediate user benefit.
