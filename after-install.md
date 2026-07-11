# Hermes Excel Sidecar installed

Run the following from a Windows PowerShell session:

```powershell
hermes excel-sidecar check
hermes excel-sidecar install
hermes excel-sidecar status
```

The installer selects a free local port, writes the sideloaded manifest,
registers one tracked Scheduled Task supervisor, and validates the bridge.
Use `hermes excel-sidecar rollback` to remove the per-user installation.
