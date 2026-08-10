# Changelog

All notable changes are documented here. This project follows Semantic
Versioning and uses GitHub Releases for distributable plugin archives.

## [Unreleased]

### Added

- Strict, non-secret profile ownership receipts and call-time profile path
  resolution as the foundation for safe multi-profile Excel ownership.

### Security

- Archive validation now rejects tracked profile receipts, tokens, workbooks,
  logs, TLS material, and generated runtime data.

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
