# Implementation Plan: create-onenote-section tools

**Date:** 2026-05-01
**Spec:** `docs/superpowers/specs/2026-05-01-onenote-section-tools-design.md`
**Approach:** TDD — failing test first, then implementation, then green.

## Steps

### 1. Verify generator offline capability

Run `npm run generate` once **before** changing anything. Two outcomes:

- **Succeeds** → generator is offline-friendly; proceed normally.
- **Fails on `npx -y openapi-zod-client`** → flag to user, request network access, then proceed. Do not pre-edit the generated client by hand.

Confirmed already: `openapi/openapi.yaml` is committed to the repo, so step 1 of the generator (download) is skipped.

### 2. Write the failing unit test FIRST

Add to `src/__tests__/graph-tools.test.ts` a new `describe('create-onenote-section')` block, following the `returnDownloadUrl` and `kebab-case path param normalization` patterns already in the file. The synthetic endpoint should be:

- `alias: 'create-onenote-section'`
- `method: 'post'`
- `path: '/me/onenote/notebooks/:notebookId/sections'`
- `parameters: [{ name: 'notebookId', type: 'Path', schema: z.string() }, { name: 'body', type: 'Body', schema: z.object({ displayName: z.string() }) }]`

Config:

```ts
{ toolName: 'create-onenote-section', pathPattern: '/me/onenote/notebooks/{notebook-id}/sections', method: 'post', scopes: ['Notes.Create'] }
```

Assertions for the first test (path + method + body):

- `graphRequest` called once.
- First arg (path) ends with `/me/onenote/notebooks/nb-abc/sections` (no `:notebookId`, no `{notebook-id}`).
- Second arg (options) has `method: 'POST'`, `body: '{"displayName":"Q2 Planning"}'`.

Add a second test invoking `tool!.handler({ 'notebook-id': 'nb-abc', body: { displayName: 'Q2 Planning' } })` to confirm kebab-case path-param normalization works (the codebase already supports both forms; this protects against regression).

Add a parallel pair of tests for `create-onenote-section-in-group` using `/me/onenote/sectionGroups/:sectionGroupId/sections` and `{sectionGroup-id}` placeholder.

Run `npm test` — these tests **must fail** because the executor has no endpoint matching the synthetic config until we add one. (Actually, they may pass if the synthetic endpoint is fully self-contained in the test — that's fine, the test still validates the executor behavior. The endpoints.json change in step 3 is what makes the real tools appear.)

### 3. Add the two endpoints to `src/endpoints.json`

Insert the two JSON objects from the spec immediately after the `create-onenote-section-page` entry (currently around line 468). Keep the OneNote block contiguous. Validate JSON syntax (`node -e "JSON.parse(require('fs').readFileSync('src/endpoints.json'))"`).

### 4. Regenerate the client

Run `npm run generate`. Verify:

- `src/generated/client.ts` now contains two new entries with `alias: 'create-onenote-section'` and `alias: 'create-onenote-section-in-group'`. Quick check: `grep "create-onenote-section\b\|create-onenote-section-in-group" src/generated/client.ts`.
- The schema reference is `microsoft_graph_onenoteSection` for the `body` parameter on both.

If the grep returns nothing, the generator did not pick up the new endpoints — re-check the JSON edit.

### 5. (Optional) Add integration-style test

Create `test/onenote-sections.test.ts` mirroring `test/mail-folders.test.ts`. Two `describe` blocks (`create-onenote-section`, `create-onenote-section-in-group`), each verifying:

- POST shape via mocked `global.fetch`.
- 404 / 403 error path returns falsy.

This is **optional** — the unit tests in step 2 cover the executor contract; this file is for regression protection of request shape.

### 6. Run full verification

```
npm test
npm run lint
npm run format:check
npm run build
```

All four must pass.

### 7. Manual smoke test (implementer's discretion)

If a real Microsoft 365 dev tenant is available, start the server (`npm run dev`) and invoke `create-onenote-section` against a real notebook ID obtained from `list-onenote-notebooks`. Verify the section appears in OneNote.

If no dev tenant — skip; the unit + integration tests carry the load.

### 8. Commit

```
git add src/endpoints.json src/generated/client.ts src/__tests__/graph-tools.test.ts
# only add test/onenote-sections.test.ts if step 5 was done
git commit -m "feat: add create-onenote-section and create-onenote-section-in-group tools"
```

## Acceptance criteria

- [ ] Both new tools appear in `src/generated/client.ts` after regeneration.
- [ ] Server registration (`registerGraphTools` log line) reports two additional tools registered.
- [ ] Both tools have `Notes.Create` scope.
- [ ] Both tools are filtered out by `--read-only` mode.
- [ ] `llmTip` for each tool clearly distinguishes the two and points at `list-onenote-notebooks` / `list-onenote-notebook-sections` for parent ID discovery.
- [ ] Full test suite passes; lint clean; format clean; build succeeds.

## Hard constraints (do not violate)

- Do NOT modify `src/graph-tools.ts` — the generic executor already handles POST bodies.
- Do NOT add `contentType: "text/html"` — these endpoints use default JSON.
- Do NOT add section delete, section move, or section group CRUD tools.
- Do NOT skip the failing-test-first step.
