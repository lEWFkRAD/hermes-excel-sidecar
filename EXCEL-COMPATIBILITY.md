# Excel compatibility

Compatibility depends on **both the Excel JavaScript API and the embedded
browser**. A product name such as "Excel 2016" is insufficient to decide either.
The pane checks `Office.context.requirements.isSetSupported` at runtime, and the
manifest declares ExcelApi 1.1. Unknown support fails closed.

## Version and platform targets

| Host | Intended behavior | Verification status |
| --- | --- | --- |
| Microsoft 365 Excel, Windows | All implemented actions when ExcelApi 1.7 is present; current WebView2 required | API profiles tested with mocks; live release qualification pending |
| Excel 2024 / LTSC 2024, Windows | Same capability checks; no subscription-only API dependency | Physical host qualification pending |
| Excel 2021 / LTSC 2021, Windows | Same capability checks | Physical host qualification pending |
| Excel 2019, Windows | Full implemented action set when 1.7 and a modern webview are available | Physical host and webview qualification pending |
| Excel 2016, Windows | Basic operations on 1.1; newer actions gated separately | Conditional API target only; IE/legacy Edge cannot run this pane |
| Excel 2013 and earlier | No supported ExcelApi 1.1 path | Unsupported |
| Excel on Mac | API detection and fallbacks are portable; modern WKWebView required | Manual installation/TLS and live qualification pending; Windows installer does not apply |
| Excel for the web | Office API surface may suffice | Not certified: localhost HTTPS, browser network permissions and deployment need a dedicated integration test |
| iPad / Android | No supported local bridge installation | Unsupported deployment |

Windows 32-bit and 64-bit Excel use the same Office.js protocol. This is not a
COM/VSTO plugin, and no bitness-specific native adapter has been added. Both
architectures still need live testing. No compatibility claim here implies a
vendor lifecycle/support commitment or a completed test on that product.

## API tiers

| Requirement | Operations |
| --- | --- |
| ExcelApi 1.1 | Bounded selection/read/verification, cell writes, basic formatting, sheet creation/rename/delete, insert/delete/clear, reviewed writes and guarded Undo |
| ExcelApi 1.2 | Explicit autofit, row/column sizing, merge/unmerge and sorting |
| ExcelApi 1.4 | Nullable used-range discovery; earlier hosts use `getUsedRange()` |
| ExcelApi 1.6 | Conditional formatting |
| ExcelApi 1.7 | Freeze/unfreeze panes |

Before either immediate execution or review binding/application, the complete
action list is checked. An unsupported explicit operation rejects the whole
response before any workbook changes or exports. This is a compatibility
preflight, **not an atomic transaction guarantee for later runtime failures**.

Default presentation formatting skips automatic autofit below 1.2 and automatic
header freezing below 1.7. Explicit requests for those features are rejected,
not silently dropped. Basic formatting and newly created sheet data remain
available on 1.1. Range sizing uses `getCell` plus `getBoundingRect`, both 1.1,
with absolute dimensions anchored at the top-left cell.

Before 1.4, an empty sheet's used range is conservatively reported as A1 instead
of claiming a zero-sized range. Selection and follow-up reads remain bounded.
The workbook envelope includes the supported relevant API tiers so Hermes can
avoid unavailable actions. These hints are advisory; local preflight enforces
the actual host capabilities.

## Browser and formula limitations

The pane requires JavaScript modules, optional chaining/nullish coalescing,
fetch, AbortController, FileReader, Blob and Web Crypto. IE and legacy Edge are
unsupported. A persistent startup message explains failures before the module
can initialize. Activity polling uses a cancellable timer rather than requiring
the newer `AbortSignal.timeout`; the deadline includes response-body reading.

ExcelApi versions do not certify worksheet functions. XLOOKUP, LET, dynamic
arrays and other newer formulas need separate host testing; the plugin does not
rewrite them into older equivalents. Existing formula-result verification
remains a bounded spot check. Macro execution and VBA are not provided. Legacy
`.xls`, protected workbooks, shared/coauthored files and locale-specific formulas
need separate fixtures; accepting an attachment extension does not certify live
editing of that workbook format.

## Release qualification checklist

For each product/build, webview and bitness, record versions and actual API
checks, then test synthetic workbooks: non-A1 write/rebasing, whole-sheet bounded
selection, empty sheets, create-sheet styling, each optional operation, review
staleness, Undo after user edits, cancellation, attachment limits and export.
Include protected sheets, `.xlsx`/`.xlsm`/`.xls`, localized formulas, and an
interrupted write. Do not use client workbooks. Unit tests emulate API tiers
1.1, 1.2, 1.4, 1.6, 1.7, 1.12 and 1.20; they do not emulate complete Office hosts.

## Microsoft references

- [Excel API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets?view=common-js)
- [Range APIs and requirement annotations](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-preview)
- [Worksheet used-range behavior](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-preview)
- [Range formatting requirements](https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeformat?view=excel-js-preview)
- [Freeze panes: ExcelApi 1.7](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetfreezepanes?view=excel-js-preview)
- [Office webviews](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)

Reviewed against Microsoft documentation on 2026-09-21. API annotations support
the implementation choices; product certification requires the live checks above.
