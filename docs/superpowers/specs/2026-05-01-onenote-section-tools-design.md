# Design: create-onenote-section tools

**Date:** 2026-05-01
**Branch:** feat/onenote-section-tools
**Status:** Ready for implementation

## Problem

The MCP exposes OneNote read tools (`list-onenote-notebooks`, `list-onenote-notebook-sections`, `list-onenote-section-pages`, `get-onenote-page-content`) and two page-creation tools (`create-onenote-page`, `create-onenote-section-page`), but **no tool to create a section**. A user asking the model to "make a new OneNote section called Q2 Planning under my Work notebook" has no path to fulfilment without falling back to raw HTTP — which is not exposed.

Microsoft Graph offers two POST endpoints:

1. `POST /me/onenote/notebooks/{notebook-id}/sections` — create a section directly under a notebook
2. `POST /me/onenote/sectionGroups/{sectionGroup-id}/sections` — create a section nested in a section group

Both have request body `{ "displayName": "..." }` and require `Notes.Create` scope.

## Design

Add two declarative entries to `src/endpoints.json`. The existing generic executor in `src/graph-tools.ts` handles POST bodies (lines 202-222 for `Body`-typed parameters, lines 308-319 for JSON serialization) — **no executor changes required**. Tool naming mirrors the `create-mail-folder` / `create-mail-child-folder` precedent.

### New endpoint entries

```json
{
  "pathPattern": "/me/onenote/notebooks/{notebook-id}/sections",
  "method": "post",
  "toolName": "create-onenote-section",
  "scopes": ["Notes.Create"],
  "llmTip": "Creates a section directly under a notebook. Use create-onenote-section-in-group when the parent is a section group instead of a notebook. Use list-onenote-notebooks to find the notebook-id. Body must be {\"displayName\": \"Section Name\"}."
}
```

```json
{
  "pathPattern": "/me/onenote/sectionGroups/{sectionGroup-id}/sections",
  "method": "post",
  "toolName": "create-onenote-section-in-group",
  "scopes": ["Notes.Create"],
  "llmTip": "Creates a section inside a section group (a section group is a folder-like container that holds sections within a notebook). Use create-onenote-section when the parent is a notebook directly. Body must be {\"displayName\": \"Section Name\"}."
}
```

Place them in `endpoints.json` immediately after `create-onenote-section-page` (currently the last OneNote entry, ending around line 468), keeping the OneNote block contiguous.

### Body handling — confirmed

The Graph OpenAPI spec (verified at `openapi/openapi.yaml` line 467129 for the notebook path and line 469746 for the section-group path) declares both POST operations with `requestBody` referencing schema `microsoft.graph.onenoteSection`. After regeneration, `src/generated/client.ts` will produce a `Body`-typed `body` parameter for each tool, schema-bound to `microsoft_graph_onenoteSection`. The executor's `Body` branch (`graph-tools.ts:202`) accepts the parameter and JSON-serializes it at line 318. The auto-correction branch (lines 207-215) only fires when a `body` argument IS supplied but its raw value fails schema parse — in that case it retries with the value re-wrapped as `{[paramName]: value}`. **Stray sibling fields (e.g. a top-level `displayName` with no `body` key) are NOT gathered into a body** — the LLM must invoke with `{ body: { displayName: "..." } }`. The `llmTip` for each tool reflects this. **No `contentType` override needed** — these tools use the default `application/json`, unlike `create-onenote-page` which needs `text/html`.

### Scopes

`Notes.Create` is already declared on `create-onenote-page` and `create-onenote-section-page`. No changes to consent scopes, scope rollups, or login flow needed — adding new tools with this scope is a no-op for users already authorized for OneNote writes.

### Tool registration

Tools will appear automatically once added to `endpoints.json` and the client is regenerated, because `registerGraphTools` (in `graph-tools.ts:480`) iterates `api.endpoints` and registers each one. They will be filtered out under `--read-only` mode (graph-tools.ts:488), as expected.

## Out of scope (explicit non-goals)

- **Section deletion** (`DELETE /me/onenote/sections/{onenoteSection-id}`) — not requested; do not add.
- **Section moves / copies** (`copyToNotebook`, `copyToSectionGroup`) — not requested; do not add.
- **Section group CRUD** (creating/deleting section groups) — not requested; do not add.
- **Section rename** (PATCH on a section) — not requested; do not add.
- **Validation of `displayName`** beyond what Graph enforces server-side — let Graph return its 400 if the name is invalid; do not duplicate validation client-side.

If any of these come up in implementation discussion, push back and treat as a separate change.

## Acceptance criteria

1. Both tool entries are present in `src/endpoints.json` with the exact JSON shapes above.
2. After running the generator (`npm run generate`), `src/generated/client.ts` contains entries with `alias: 'create-onenote-section'` and `alias: 'create-onenote-section-in-group'`.
3. Starting the server (`npm run dev` or similar) registers both tools; they appear in MCP tool listings.
4. Unit test (in `src/__tests__/graph-tools.test.ts` or a new file) exercises the executor against a synthetic `create-onenote-section` config and verifies: POST method, correct path with `notebook-id` substituted, `application/json` Content-Type (default), JSON body with `displayName`.
5. Optional integration-style test in `test/onenote-sections.test.ts` mirrors `test/mail-folders.test.ts` patterns, verifying request shape via mocked `global.fetch`.
6. `npm test` is green; `npm run lint` is clean.

## Open questions for the implementer

1. **Generator offline capability.** `bin/modules/download-openapi.mjs` skips download if `openapi/openapi.yaml` exists (verified — it does). However, `bin/modules/generate-mcp-tools.mjs:19` invokes `npx -y openapi-zod-client`, which may need network if not cached. **Investigate before running**: try `npm run generate` and observe — if it fails offline, document the requirement.
2. **OneNote section creation eventual consistency.** Graph's OneNote APIs are known for asynchronous behavior (sections sometimes don't appear immediately in subsequent list calls). Do not add retry/polling logic; surface this in the `llmTip` only if testing reveals it as a frequent UX problem.
3. **Body wrapping behavior.** The executor's body auto-correction (graph-tools.ts:207-215) wraps unwrapped values when the schema parse fails. Verify with the unit test that passing `{ displayName: "X" }` directly (without a `body` wrapper) works — this is the common case from LLMs.
