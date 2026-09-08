# Hermes runtime compatibility

This sidecar registers `excel_response` with `always_visible=True`. Hermes
must support this argument in both `PluginContext.register_tool` and the tool
registry, and tool search must honor it. Otherwise plugin registration fails
or progressive disclosure hides the terminal response tool and Excel chat
cannot complete. Do not upgrade an unpatched runtime to this sidecar version.

The companion [patch](hermes-always-visible.patch) adds that generic capability.
It does not enable disabled toolsets, bypass availability checks, or disable
tool search globally. This is a local compatibility patch, not a claim that
upstream Hermes has merged or released the capability.

The patch was prepared against Hermes commit
`f1ccf436a27522c1bb5d36383a6f13b950676338`. From your Hermes source checkout,
using absolute paths to the downloaded sidecar files:

```powershell
git apply --check /path/to/hermes-excel-sidecar/compat/hermes-always-visible.patch
git apply /path/to/hermes-excel-sidecar/compat/hermes-always-visible.patch
Copy-Item /path/to/hermes-excel-sidecar/compat/test_terminal_tool_visibility.py tests/tools/test_terminal_tool_visibility.py
python -m pytest tests/tools/test_terminal_tool_visibility.py tests/tools/test_tool_search.py -q
```

Use the Hermes virtual environment for the test command. If the check fails,
compare the runtime API with the patch and adapt it to that checkout; do not
force-apply it or overwrite unrelated changes. Restart the gateway after the
runtime patch and sidecar update. The sidecar does not modify Hermes source
automatically. Preserve the patch when upgrading Hermes until upstream provides
equivalent support.

Validation performed September 6, 2026: 46 Hermes visibility/tool-search tests
passed. The local bridge/model smoke run passed 21 scenarios, including a
formula table anchored at H23, follow-up chat, formatting proposals, and CSV
export. Three smoke scenarios were skipped; the bad-Host scenario was checked
separately (HTTP 421). Model-down behavior and cross-service toolset containment
were not verified in that run. Workbook application and Undo tests use mocked
Office.js; these results do not claim a live Excel workbook mutation test.
