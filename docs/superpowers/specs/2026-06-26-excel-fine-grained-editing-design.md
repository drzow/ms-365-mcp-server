# Fine-grained Excel (xlsx) editing via the Graph API

**Date:** 2026-06-26
**Status:** Approved design — pending implementation plan

## Summary

Extend the MS-365 MCP server with fine-grained Excel workbook editing: writing
individual cell/range values, formulas, and number formats; font and fill
formatting; worksheet management; table management; and workbook session
management for fast, consistent multi-edit operations.

Word/docx fine-grained editing was explicitly **dropped from scope**: Microsoft
Graph exposes no per-paragraph/per-run or text-formatting API for `.docx`. Word
documents are only reachable as drive items (download/overwrite-whole-file/
convert-format). Fine-grained Word editing exists only via client-side Office.js
add-ins or server-side OpenXML byte manipulation — neither is "through the Graph
API," so it is out of scope here.

## Background: how tools are defined in this repo

Tools are **declarative**. Each tool is an entry in `src/endpoints.json`
(`pathPattern`, `method`, `toolName`, `scopes`, optional `llmTip`, etc.).
`npm run generate`:

1. Loads the cached Microsoft Graph OpenAPI spec (`openapi/openapi.yaml`, 36 MB,
   already on disk — generation works offline).
2. Trims it to only the paths listed in `endpoints.json`
   (`bin/modules/simplified-openapi.mjs`), simplifies/prunes schemas, and writes
   `openapi/openapi-trimmed.yaml`.
3. Runs `openapi-zod-client` to produce `src/generated/client.ts`.

At runtime, `src/graph-tools.ts` reads `api.endpoints` (generated) merged with
`endpoints.json` metadata and registers every entry as an MCP tool via
`executeGraphTool`, which already handles: path/query/body/header param mapping,
read-only filtering, multi-account token resolution, body auto-wrapping, OData
params, pagination, and error handling.

## The constraint that shapes the design

Microsoft's OpenAPI metadata is **incomplete for Excel writes**. Verified against
the cached spec:

| Capability                                                                    | Spec status                               |
| ----------------------------------------------------------------------------- | ----------------------------------------- |
| `range(address='…')/clear`, `/merge`, `/unmerge`, `/insert`, `/delete` (POST) | present                                   |
| `range(address='…')/format` PATCH (alignment, column width, row height, wrap) | present                                   |
| `worksheets/add` (POST), `worksheets/{id}` (PATCH, DELETE)                    | present                                   |
| `worksheets/{id}/tables/add` (POST), `tables/{id}` (PATCH, DELETE)            | present                                   |
| `tables/{id}/rows/add`, `tables/{id}/columns/add` (POST)                      | present                                   |
| `workbook/createSession`, `closeSession`, `refreshSession` (POST)             | present                                   |
| `worksheets/{id}/usedRange()` (GET)                                           | present                                   |
| **`range(address='…')` PATCH (set values/formulas/numberFormat)**             | **GET only — missing**                    |
| **`range(address='…')/format/font` and `/fill` (PATCH)**                      | **missing (modeled as nested nav props)** |
| **`range(address='…')/sort/apply` (POST)**                                    | **missing**                               |

`simplified-openapi.mjs` _throws_ if an `endpoints.json` path is absent from the
spec, and silently drops a tool if the requested method is absent. So the
headline features (set cell values, font/fill, addressed sort) cannot be added by
listing them in `endpoints.json` alone.

Two pre-existing bugs are in scope to fix: `format-excel-range` and
`sort-excel-range` use the empty-parens `range()/format` and `range()/sort`
forms, which provide no way to specify _which_ range — they effectively target
nothing.

## Chosen approach: spec augmentation

Add a generator step that injects the missing operations into the in-memory
OpenAPI object **before** trimming. The new operations use slim, self-contained
inline request-body schemas (no `$ref` to the bloated full workbook schemas), so
the generated tools are tight and LLM-friendly. Everything then flows through the
existing declarative machinery — auto-registration, read-only filtering,
multi-account, body auto-wrap, error handling, LLM tips — with no per-tool custom
code.

**Rejected alternatives:**

- **Hand-written custom tools** calling `graphClient` directly: bypasses all
  cross-cutting machinery (read-only, multi-account token resolution, scopes,
  body handling) and introduces a one-off pattern unlike the other ~100 tools.
- **Download → edit bytes → re-upload:** not fine-grained, not atomic, loses
  concurrent-edit safety, and re-uploads can clobber others' changes.

### Augmentation module

New module `bin/modules/excel-augmentations.mjs` exporting
`augmentExcelPaths(openApiSpec)`, invoked from `simplified-openapi.mjs` at the
start of `createAndSaveSimplifiedOpenAPI` (before the "path not found" check). It
inserts these path items (each with `operationId` matching the `endpoints.json`
`toolName`, the `Files.ReadWrite` scope context, the `driveItem-id`/`drive-id`/
`workbookWorksheet-id`/`address` path params, and slim inline bodies):

1. `PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')`
   — body: `{ values?: any[][], formulas?: any[][], numberFormat?: any[][] }`
2. `PATCH …/range(address='{address}')/format/font`
   — body: `{ name?, size?, color?, bold?, italic?, underline? }`
3. `PATCH …/range(address='{address}')/format/fill`
   — body: `{ color? }`
4. `POST …/range(address='{address}')/sort/apply`
   — body: `{ fields: [{ key, ascending?, sortOn? }], matchCase?, hasHeaders?, orientation? }`

If an augmented path already exists in the spec (e.g. the GET on
`range(address='…')`), the module merges the new method into the existing path
item rather than overwriting it.

### Session header support (repurpose `isExcelOp`)

The `isExcelOp: true` flag in `endpoints.json` is currently dead (referenced
nowhere in code). Repurpose it: in `graph-tools.ts`, when an endpoint's config
has `isExcelOp: true`, inject an optional `workbookSessionId` parameter into the
tool schema. In `executeGraphTool`, when present, set the
`workbook-session-id` request header. This gives optional fast/consistent
multi-edit sessions without touching the spec. All Excel tools (existing and new)
get `isExcelOp: true`.

## Tool inventory

All tools live under `…/workbook/worksheets/{workbookWorksheet-id}/…` (or
`…/workbook/…` for sessions) on `/drives/{drive-id}/items/{driveItem-id}`, take
`scopes: ["Files.ReadWrite"]` (reads use `Files.Read`), `isExcelOp: true`, and
`skipEncoding: ["address"]` where an `address` path param is present.

### Cell / range content

| Tool                            | Method · path                             | Body / notes                                                                                           |
| ------------------------------- | ----------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| `set-excel-range` _(augmented)_ | PATCH `range(address='{address}')`        | `values` / `formulas` / `numberFormat` as 2-D arrays. The headline "edit individual cells" capability. |
| `clear-excel-range`             | POST `range(address='{address}')/clear`   | `{ applyTo: All\|Formats\|Contents\|Hyperlinks }`                                                      |
| `insert-excel-range`            | POST `range(address='{address}')/insert`  | `{ shift: Down\|Right }`                                                                               |
| `delete-excel-range`            | POST `range(address='{address}')/delete`  | `{ shift: Up\|Left }`                                                                                  |
| `merge-excel-range`             | POST `range(address='{address}')/merge`   | `{ across: bool }`                                                                                     |
| `unmerge-excel-range`           | POST `range(address='{address}')/unmerge` | none                                                                                                   |

### Formatting

| Tool                                    | Method · path                                  | Body / notes                                                               |
| --------------------------------------- | ---------------------------------------------- | -------------------------------------------------------------------------- |
| `format-excel-range` _(fixed)_          | PATCH `range(address='{address}')/format`      | Re-point from `range()/format`. Alignment, wrap, column width, row height. |
| `format-excel-range-font` _(augmented)_ | PATCH `range(address='{address}')/format/font` | `name`, `size`, `color`, `bold`, `italic`, `underline`                     |
| `format-excel-range-fill` _(augmented)_ | PATCH `range(address='{address}')/format/fill` | `color`                                                                    |
| `sort-excel-range` _(fixed, augmented)_ | POST `range(address='{address}')/sort/apply`   | Re-point from PATCH `range()/sort`. `fields`, `matchCase`, `hasHeaders`.   |

Borders are out of scope (per-edge style/color/weight; deferred).

### Worksheets

| Tool                     | Method · path                                       | Body / notes                            |
| ------------------------ | --------------------------------------------------- | --------------------------------------- |
| `add-excel-worksheet`    | POST `worksheets/add`                               | `{ name? }`                             |
| `update-excel-worksheet` | PATCH `worksheets/{workbookWorksheet-id}`           | `name`, `position`, `visibility`        |
| `delete-excel-worksheet` | DELETE `worksheets/{workbookWorksheet-id}`          | none                                    |
| `get-excel-used-range`   | GET `worksheets/{workbookWorksheet-id}/usedRange()` | Read the populated area (`Files.Read`). |

(`list-excel-worksheets` and `get-excel-range` already exist.)

### Tables

| Tool                     | Method · path                                       | Body / notes                                 |
| ------------------------ | --------------------------------------------------- | -------------------------------------------- |
| `add-excel-table`        | POST `worksheets/{workbookWorksheet-id}/tables/add` | `{ address, hasHeaders }`                    |
| `add-excel-table-row`    | POST `tables/{workbookTable-id}/rows/add`           | `{ values, index? }`                         |
| `add-excel-table-column` | POST `tables/{workbookTable-id}/columns/add`        | `{ values?, index?, name? }`                 |
| `update-excel-table`     | PATCH `tables/{workbookTable-id}`                   | `name`, `showHeaders`, `showTotals`, `style` |
| `delete-excel-table`     | DELETE `tables/{workbookTable-id}`                  | none                                         |

Path note: `add-excel-table` creates the table on a specific worksheet
(`…/worksheets/{workbookWorksheet-id}/tables/add`). The table-item operations
(`rows/add`, `columns/add`, update, delete) address the table directly at the
workbook level (`…/workbook/tables/{workbookTable-id}/…`, no worksheet segment),
since a table id is unique within the workbook.

### Sessions

| Tool                    | Method · path                  | Body / notes                                       |
| ----------------------- | ------------------------------ | -------------------------------------------------- |
| `create-excel-session`  | POST `workbook/createSession`  | `{ persistChanges: bool }` → returns session `id`. |
| `close-excel-session`   | POST `workbook/closeSession`   | Pass `workbookSessionId`.                          |
| `refresh-excel-session` | POST `workbook/refreshSession` | Keeps a session alive.                             |

Plus the `workbookSessionId` param on all Excel tools (via `isExcelOp`).

## Files changed

- `bin/modules/excel-augmentations.mjs` — **new**: injects the 4 missing operations.
- `bin/modules/simplified-openapi.mjs` — call `augmentExcelPaths()` before trimming.
- `src/endpoints.json` — ~20 new entries + `isExcelOp`/path fixes for the 2 existing tools + LLM tips.
- `src/graph-tools.ts` — `isExcelOp` → inject `workbookSessionId` param; set `workbook-session-id` header in `executeGraphTool`; add the `isExcelOp` field to the `EndpointConfig` interface.
- `src/generated/client.ts` — regenerated (not hand-edited).
- Tests — see below.
- `README.md` — document the new Excel tools and session usage.

## Testing

- **Augmentation unit tests:** feed a minimal OpenAPI fixture to
  `augmentExcelPaths()` and assert the 4 operations are injected with correct
  methods, params, and body schemas; assert merge-not-overwrite when a path
  already has a method (the GET on `range(address='…')`).
- **Generation smoke test:** run `npm run generate` and assert each new
  `toolName` appears as an alias in `src/generated/client.ts` with the expected
  method and a body param where applicable.
- **Tool-registration tests** (extend `src/__tests__/graph-tools.test.ts`):
  - `isExcelOp` tools expose a `workbookSessionId` param; non-Excel tools do not.
  - `executeGraphTool` sets the `workbook-session-id` header when
    `workbookSessionId` is passed, and omits it otherwise.
  - `set-excel-range` issues PATCH to the addressed range with the body
    untouched; `address` is not URL-encoded (skipEncoding).
  - Read-only mode skips all new write tools (POST/PATCH/DELETE).
  - The fixed `format-excel-range`/`sort-excel-range` target the addressed form.
- Follow existing mocking patterns in `graph-tools.test.ts` (no live Graph
  calls).

## Out of scope / YAGNI

- Word/docx fine-grained editing (no Graph API).
- The ~480 spreadsheet-function endpoints (`workbook/functions/*`).
- Cell borders, named ranges, pivot tables, charts beyond the existing
  `create-excel-chart`, conditional formatting, data validation, comments.
- Cross-cell formula recalculation control beyond what sessions provide
  (`workbook/application/calculate` could be a trivial future add).
