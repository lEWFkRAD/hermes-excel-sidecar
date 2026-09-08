# Changelog

All notable changes are documented here. This project follows Semantic
Versioning and uses GitHub Releases for distributable plugin archives.

## [Unreleased]

### Fixed

- Undo now refuses targets edited or replaced since application and preserves
  existing formatting for in-place writes. Failed Undo records remain available.
- Selection previews read at most 100 rows by 16 columns, including whole-sheet
  selections, while retaining the original selection dimensions.
- Typed Excel chat preserves request correlation across validation retries,
  uses provider-appropriate tool choice, and gives precise action-schema errors.
- Model instructions retain the selected destination and formula anchoring;
  smoke scenarios use isolated request and conversation identities.
- Rollback matches exact script paths, retains profile ownership safeguards,
  and verifies the payload path and stopped supervisor before deletion.
- Rollback process enumeration works with Windows PowerShell 5.1 generic lists.

### Runtime requirement

- `excel_response` requires Hermes `always_visible` tool registration and
  discovery support. See `compat/README.md` for the companion patch and tests.

### Added

- Strict, non-secret profile ownership receipts and call-time profile path
  resolution as the foundation for safe multi-profile Excel ownership.
- Profile-bound install, check, status, deploy, sideload, task, and rollback
  entry points, including explicit one-time legacy adoption.
- Authenticated bridge health attestation for canonical profile name and owner
  receipt fingerprint, plus profile-bound adapter health and request routing.

### Security

- Archive validation now rejects tracked profile receipts, tokens, workbooks,
  logs, TLS material, and generated runtime data.
- A per-user exclusive transaction lock serializes the fixed Office/task/port
  singleton; unreceipted legacy artifacts and corrupt, interrupted, or
  different-profile ownership state fail closed before credentials or singleton
  state are touched.
- Secrets are read only from the owning profile's ACL-restricted files; legacy
  Windows User-scope token values are retired after ownership is established.
- Rollback retains the shared receipt until the owned process tree, port, Office
  registration, shortcuts, and payload are quiescent and removed.
- Unicode-safe launchers preserve non-ASCII profile and install paths.

## [0.2.0] - Unreleased

### Added

- Independently installable Hermes plugin entry point and CLI lifecycle.
- Authenticated, loopback-only Office.js bridge with structured workbook actions.
- Review-before-apply, post-write verification, bounded reads, and rollback support.
- Per-user Windows installation, sideload registration, and one supervised bridge.
- Package, protocol, broker, formula-rebasing, and installer validation.

### Security

- Token authentication plus `Host`, `Origin`, size, and request-budget checks.
- Model-authored JavaScript and agent tool calls remain non-executable.
- Uploads, exports, and logs stay outside the HTTP-served root.
