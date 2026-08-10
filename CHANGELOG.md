# Changelog

All notable changes are documented here. This project follows Semantic
Versioning and uses GitHub Releases for distributable plugin archives.

## [Unreleased]

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
