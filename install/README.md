# Hermes for Excel — install / packaging

Per-box, per-user installer for the Hermes for Excel add-in (Office.js task pane
+ local Node bridge). **Windows-first** (the fleet is Windows + WSL). Nothing
here requires elevation — profile-owned files live under the selected
`HERMES_HOME`, while Office registration and the supervisor live in per-user
Windows state.

## Prerequisites

- **Windows 10/11**, Excel for Windows (desktop) with add-in sideloading allowed.
- **Node.js LTS** — the installer auto-installs it via `winget install OpenJS.NodeJS.LTS`
  if `node` is not on PATH. If `winget` is unavailable, install Node.js LTS
  manually from <https://nodejs.org> and re-run.
- A running **Hermes gateway** (`api_server` on `http://127.0.0.1:8642/v1`) for
  the add-in to be useful. The installer's Hermes check is non-fatal — install
  still succeeds if the gateway is down; you just won't get model replies until
  it's up.
- Optional: **Docling** for PDF/Office parsing. Default mode is `wsl` (distro
  `Ubuntu-24.04`); see env vars below.

## One-line install

```powershell
powershell -ExecutionPolicy Bypass -File install\apply.ps1
```

By default the installer selects the first free port starting at `8788` and
rewrites the installed manifest to match. Add `-Port 8790` to require a
specific port (the install fails if it is occupied), or
`-SkipSideload` to skip the Excel developer-catalog registration, or
`-SkipExcelAddin` to opt a box out entirely.

When invoked through Hermes, `ProfileName`, `HermesHome`, and the shared owner
receipt path are supplied automatically. Direct script callers must provide
the same identity explicitly for named profiles. For example:

```powershell
powershell -ExecutionPolicy Bypass -File install\apply.ps1 `
  -ProfileName finance `
  -HermesHome "$env:LOCALAPPDATA\hermes\profiles\finance" `
  -OwnerReceiptPath "$env:LOCALAPPDATA\hermes\shared\excel-sidecar\owner-v1.json"
```

The Office catalog, fixed Scheduled Task, and localhost port are singletons for
the Windows user. The installer takes an exclusive per-user transaction lock
and validates the shared receipt before inspecting secrets or mutating any of
those resources. Another profile is refused. For an older installation with no
receipt, `-AdoptLegacy` is an explicit, install-only migration and succeeds only
when the payload, task, shortcut, and Office registration consistently point to
the requested profile install. There is no automatic takeover path.

After install, **restart Excel** and open the pane from the **Hermes** ribbon
group.

## What each script does

| Script | Purpose |
| --- | --- |
| `apply.ps1` | Fleet entry point. Resolves the selected profile and calls `addin-install.ps1`; `-SkipExcelAddin` opts a box out; passes `-Port`, `-SkipSideload`, and explicit legacy adoption through. |
| `addin-install.ps1` | The installer. Locks and verifies singleton ownership, ensures Node, copies the profile payload, generates ACL-restricted bridge and adapter tokens, and registers autostart + sideload. It commits the new same-owner receipt before activating that profile's gateway, then requires an authenticated typed-proposal and owner-identity certification turn. An interrupted install retains authoritative ownership; a clean failed upgrade restores the exact prior receipt. |
| `run-bridge.cmd.template` | Template for the bridge launcher. Tokens are read at runtime from ACL-restricted files under the data directory; raw fallback is disabled. |
| `run-bridge.vbs.template` | Hidden launcher for `run-bridge.cmd` (no console window, no elevation). |
| `register-task.ps1` | Registers the logon-triggered Scheduled Task `Hermes_Excel_Bridge` (per-user, hidden, no elevation) running the supervisor. `-Unregister` removes it. |
| `register-sideload.ps1` | Writes a port-adjusted manifest and registers it with Microsoft's pinned `office-addin-dev-settings` tool. It also removes the obsolete legacy registry value. `-Unregister` reverses it. |
| `service\bridge-service.cmd` | Restart-loop supervisor: relaunches the bridge on exit, backs off (stands down) when it crashes instantly, logs restart events with timestamps to `bridge-restarts.log` under the **data dir**. |
| `rollback.ps1` | Verifies the selected profile owns the singleton, then reverses its task, shortcuts, WEF registration, and install directory. The owner receipt is removed last and is retained on cancellation or failure. |

## Environment variables (set by the launchers, read by the bridge)

| Var | Default | Meaning |
| --- | --- | --- |
| `PORT` | `8788` | Bridge port. |
| `HERMES_EXCEL_DATA_DIR` | `%LOCALAPPDATA%\hermes\excel-addin\data` | Writable per-user data root for uploads/exports/logs, **outside** the served web root. |
| `HERMES_EXCEL_BRIDGE_TOKEN` | generated per install | Random session token the bridge requires on `/api/*` and the pane sends. Persisted only to the owning profile's `data\.bridge-token` with a user-only ACL. |
| `HERMES_EXCEL_INGEST_TOKEN` | generated per install | Separate 64-hex secret shared only by the bridge and the owning profile's Hermes Excel platform adapter. Persisted only to `data\.ingest-token` with current-user/SYSTEM ACLs. |
| `HERMES_HOME` / `HERMES_CONFIG` | selected profile | Absolute profile home and `config.yaml` baked into the launcher. |
| `HERMES_EXCEL_PROFILE_NAME` | selected profile | Canonical, non-secret profile identity reported by authenticated health checks. |
| `HERMES_EXCEL_OWNER_FINGERPRINT` | owner receipt | Non-secret SHA-256 identity binding the bridge process to the validated receipt. |
| `HERMES_EXCEL_DOCLING_MODE` | `wsl` | `wsl` \| `native` \| `docker`. |
| `HERMES_EXCEL_WSL_DISTRO` | `Ubuntu-24.04` | WSL distro used when docling mode is `wsl`. |

## Manual sideload fallback

If the registry catalog isn't picked up (policy-locked box, or a different Excel
build), load the manifest by hand:

> **Excel → Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in →**
> browse to `<selected HERMES_HOME>\excel-addin\OfficeAddinManifests\hermes-excel-addin.xml`

The Upload-My-Add-in path works even when the Developer catalog does not.

## Uninstall

```powershell
powershell -ExecutionPolicy Bypass -File install\rollback.ps1          # prompts before deleting data
powershell -ExecutionPolicy Bypass -File install\rollback.ps1 -Force   # no prompt
```

Run rollback with the same profile identity that owns the receipt. A missing,
malformed, or different-profile receipt fails closed before credentials are
read or state is removed.

## Deployer assumptions to verify

- **Office registry version is `16.0`** (Office 2016/2019/2021/365). Older Excel
  uses a different `WEF\Developer` hive — pass `-OfficeVersion` to
  `register-sideload.ps1` / `rollback.ps1` if needed.
- **`winget` availability** — required only if Node isn't already installed.
- **WSL distro name** (`Ubuntu-24.04`) and **docling mode** match the box; retune
  the config block at the top of `addin-install.ps1`.
- The hardened `broker\server.mjs` consumes launcher-injected
  `HERMES_EXCEL_DATA_DIR` and token values loaded from the owning profile's
  ACL-restricted files. The plugin manifest does not solicit or persist either
  generated token as user configuration.
