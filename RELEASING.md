# Releasing

Releases are built from immutable tags by `.github/workflows/release.yml`. The
workflow validates the repository, builds and re-tests the exact plugin archive,
requires the dereferenced tag commit to be reachable from `origin/main`, attests
its provenance, and then creates the GitHub Release. Do not upload a locally
built replacement asset to an existing version.

## Prepare

1. Choose a Semantic Version without a leading `v`.
2. Set the same version in `plugin.yaml` and `package.json`.
3. Run `npm install --package-lock-only --ignore-scripts` so `package-lock.json`
   records the same version.
4. Replace the target version section's `Unreleased` marker with its release
   date, and move any later work into the top-level `Unreleased` section.
5. Run the complete local validation:

   ```powershell
   npm ci --ignore-scripts
   npm run validate
   ```

6. Open a pull request and merge only after the stable `Required PR checks`
   status from the `CI` workflow passes. Confirm the release commit is on `main`.

## Tag

Create one signed or annotated tag whose name is exactly `v` plus the synchronized
manifest/package version. Put concise release notes in the tag message because the
release workflow uses them for the GitHub Release.

```powershell
git switch main
git pull --ff-only
git tag -s v0.2.0 -m "Hermes Excel Sidecar v0.2.0"
git push origin v0.2.0
```

Use the intended release version in place of `0.2.0`. Never move or reuse a
published tag; fix forward with a new patch version.

## Verify

After the Release workflow succeeds:

1. Download `hermes-excel-sidecar-vX.Y.Z.zip` and its `.sha256` file from the
   GitHub Release, then compare the recorded digest with `Get-FileHash`.
2. Verify the archive's GitHub artifact attestation:

   ```powershell
   gh attestation verify .\hermes-excel-sidecar-vX.Y.Z.zip --repo lEWFkRAD/hermes-excel-sidecar
   ```

3. Extract it into a clean directory and run `npm ci --ignore-scripts` followed
   by `npm run validate`.
4. On a disposable Windows profile with a collision-free port and a confirmed
   `hermes-excel-bridge`, perform the live `npm run smoke` check separately.

The live Excel/model check is intentionally not part of GitHub-hosted CI because
it requires an interactive Excel session and a real local Hermes gateway.
