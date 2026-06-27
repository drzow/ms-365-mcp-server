# Fine-grained Excel Editing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add ~20 fine-grained Excel workbook editing tools (cell/range value & formula writes, font/fill formatting, worksheet & table management, workbook sessions) to the MS-365 MCP server.

**Architecture:** Tools are declarative — entries in `src/endpoints.json` are trimmed from the cached Graph OpenAPI spec by `bin/modules/simplified-openapi.mjs`, generated into `src/generated/client.ts`, and auto-registered by `src/graph-tools.ts`. Microsoft's metadata omits four Excel write operations (range value PATCH, font/fill PATCH, addressed sort); a new augmentation module injects those into the spec before trimming so they flow through the same machinery. The dead `isExcelOp` flag is repurposed to inject an optional workbook-session header param.

**Tech Stack:** TypeScript, Zod, `@modelcontextprotocol/sdk`, Vitest, `openapi-zod-client` (generation), Node ESM generator scripts (`bin/modules/*.mjs`).

## Global Constraints

- Node `>=18` (engines); target Node `>=20`.
- Never hardcode secrets; not relevant here (no new config).
- All new Excel tools use scope `Files.ReadWrite` (the one read-only tool, `get-excel-used-range`, uses `Files.Read`).
- All Excel tool paths are rooted at `/drives/{drive-id}/items/{driveItem-id}/workbook`.
- Every endpoint touching an `address` path segment sets `"skipEncoding": ["address"]` (range addresses like `A1:B2` must not be URL-encoded).
- Every Excel endpoint sets `"isExcelOp": true`.
- Run `npm run format` before committing; the repo uses Prettier + ESLint. `npm test` must stay green.
- Generation runs offline — the 36 MB spec is already at `openapi/openapi.yaml`; do **not** pass `--force`.

---

### Task 1: Workbook session header support (`isExcelOp`)

Repurpose the currently-dead `isExcelOp` flag: when set, inject an optional `workbookSessionId` tool param that becomes the `workbook-session-id` request header. Pure code change, unit-testable without regeneration.

**Files:**
- Modify: `src/graph-tools.ts` (EndpointConfig interface; param injection in `registerGraphTools`; header set + skip-list in `executeGraphTool`)
- Test: `src/__tests__/graph-tools.test.ts`

**Interfaces:**
- Produces: `EndpointConfig.isExcelOp?: boolean`; tools whose config has `isExcelOp: true` expose an optional `workbookSessionId: string` param; when supplied, `executeGraphTool` sets header `workbook-session-id`.

- [ ] **Step 1: Write the failing tests**

Add to the end of `src/__tests__/graph-tools.test.ts`, before the final closing `});` of the top-level `describe('graph-tools', ...)`:

```ts
  // ---- 8. isExcelOp workbook session header ----
  describe('isExcelOp workbook session header', () => {
    it('exposes a workbookSessionId param when isExcelOp is true', async () => {
      const endpoint = makeEndpoint({
        alias: 'set-excel-range',
        method: 'patch',
        path: "/drives/:driveId/items/:driveItemId/workbook/worksheets/:workbookWorksheetId/range(address=':address')",
        parameters: [{ name: 'body', type: 'Body', schema: z.object({ values: z.any() }).passthrough() }],
      });
      const config = makeConfig({
        toolName: 'set-excel-range',
        method: 'patch',
        pathPattern: "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')",
        scopes: ['Files.ReadWrite'],
        isExcelOp: true,
        skipEncoding: ['address'],
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any);

      const tool = server.tools.get('set-excel-range');
      expect(tool).toBeDefined();
      expect(tool!.schema['workbookSessionId']).toBeDefined();
      expect(tool!.schema['workbookSessionId'].description).toContain('workbook-session-id');
    });

    it('does NOT add workbookSessionId when isExcelOp is absent', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig(); // no isExcelOp
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any);

      const tool = server.tools.get('test-tool');
      expect(tool!.schema['workbookSessionId']).toBeUndefined();
    });

    it('sets the workbook-session-id header when workbookSessionId is passed', async () => {
      const endpoint = makeEndpoint({
        alias: 'set-excel-range',
        method: 'patch',
        path: '/drives/:driveId/items/:driveItemId/workbook',
        parameters: [{ name: 'body', type: 'Body', schema: z.object({ values: z.any() }).passthrough() }],
      });
      const config = makeConfig({
        toolName: 'set-excel-range',
        method: 'patch',
        pathPattern: '/drives/{drive-id}/items/{driveItem-id}/workbook',
        scopes: ['Files.ReadWrite'],
        isExcelOp: true,
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ ok: true }) }] },
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('set-excel-range');
      await tool!.handler({ body: { values: [[1]] }, workbookSessionId: 'SESSION-123' });

      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers['workbook-session-id']).toBe('SESSION-123');
    });

    it('omits the workbook-session-id header when no session id is passed', async () => {
      const endpoint = makeEndpoint({
        alias: 'set-excel-range',
        method: 'patch',
        path: '/drives/:driveId/items/:driveItemId/workbook',
        parameters: [{ name: 'body', type: 'Body', schema: z.object({ values: z.any() }).passthrough() }],
      });
      const config = makeConfig({
        toolName: 'set-excel-range',
        method: 'patch',
        pathPattern: '/drives/{drive-id}/items/{driveItem-id}/workbook',
        scopes: ['Files.ReadWrite'],
        isExcelOp: true,
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ ok: true }) }] },
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('set-excel-range');
      await tool!.handler({ body: { values: [[1]] } });

      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers['workbook-session-id']).toBeUndefined();
    });
  });
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "isExcelOp"`
Expected: FAIL — `workbookSessionId` is undefined and the header is never set.

- [ ] **Step 3: Add `isExcelOp` to the EndpointConfig interface**

In `src/graph-tools.ts`, add the field to the `EndpointConfig` interface (after `acceptType`, around line 29):

```ts
  acceptType?: string; // Custom Accept header for endpoints returning non-JSON content (e.g., text/vtt)
  isExcelOp?: boolean; // Excel workbook op — inject optional workbookSessionId param → workbook-session-id header
```

- [ ] **Step 4: Skip `workbookSessionId` in the param loop**

In `executeGraphTool`, add `'workbookSessionId'` to the control-parameter skip list (the array around lines 130-138):

```ts
        [
          'account',
          'fetchAllPages',
          'includeHeaders',
          'excludeResponse',
          'timezone',
          'expandExtendedProperties',
          'workbookSessionId',
        ].includes(paramName)
```

- [ ] **Step 5: Set the `workbook-session-id` header**

In `executeGraphTool`, just after the `config?.acceptType` block (around line 285, before the `queryParams` assembly), add:

```ts
    if (config?.isExcelOp && params.workbookSessionId) {
      headers['workbook-session-id'] = String(params.workbookSessionId);
      logger.info('Setting workbook-session-id header for Excel operation');
    }
```

- [ ] **Step 6: Inject the `workbookSessionId` param at registration**

In `registerGraphTools`, just after the `excludeResponse` param block (around line 603, before the `supportsTimezone` block), add:

```ts
    // Excel workbook session support (endpoints flagged isExcelOp in endpoints.json).
    // Lets a batch of edits share one persistent workbook session for speed + consistency.
    if (endpointConfig?.isExcelOp) {
      paramSchema['workbookSessionId'] = z
        .string()
        .describe(
          'Optional Excel workbook session ID from create-excel-session, sent as the ' +
            'workbook-session-id header so a batch of edits shares one fast, consistent session.'
        )
        .optional();
    }
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `npx vitest run src/__tests__/graph-tools.test.ts`
Expected: PASS (all existing + 4 new tests).

- [ ] **Step 8: Commit**

```bash
npm run format
git add src/graph-tools.ts src/__tests__/graph-tools.test.ts
git commit -m "feat: workbook session header support via isExcelOp flag

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 2: Spec augmentation module

Inject the four Excel write operations Microsoft's metadata omits, before the spec is trimmed: `PATCH range(address)` (values/formulas/numberFormat), `PATCH range(address)/format/font`, `PATCH range(address)/format/fill`, `POST range(address)/sort/apply`. Uses slim inline body schemas and clones the base range path's four path params.

**Files:**
- Create: `bin/modules/excel-augmentations.mjs`
- Modify: `bin/modules/simplified-openapi.mjs` (import + call before the path-existence check)
- Test: `test/excel-augmentations.test.ts`

**Interfaces:**
- Produces: `augmentExcelPaths(openApiSpec)` — mutates and returns `openApiSpec`, adding `.patch` to the existing `range(address='{address}')` path item and three new path items: `…/range(address='{address}')/format/font` (`patch`), `…/format/fill` (`patch`), `…/sort/apply` (`post`). Each new path item carries cloned path-level `parameters` including `address`. Throws if the base range path is absent.

- [ ] **Step 1: Write the failing test**

Create `test/excel-augmentations.test.ts`:

```ts
import { describe, it, expect } from 'vitest';
import { augmentExcelPaths } from '../bin/modules/excel-augmentations.mjs';

const RANGE =
  "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')";

function minimalSpec() {
  return {
    paths: {
      [RANGE]: {
        get: { operationId: 'get-x', responses: { '2XX': { description: 'ok' } } },
        parameters: [
          { name: 'drive-id', in: 'path', required: true, schema: { type: 'string' } },
          { name: 'driveItem-id', in: 'path', required: true, schema: { type: 'string' } },
          { name: 'workbookWorksheet-id', in: 'path', required: true, schema: { type: 'string' } },
          { name: 'address', in: 'path', required: true, schema: { type: 'string', nullable: true } },
        ],
      },
    },
    components: { responses: { error: { description: 'error' } } },
  };
}

describe('augmentExcelPaths', () => {
  it('adds a PATCH to the existing range path without dropping GET', () => {
    const spec = minimalSpec();
    augmentExcelPaths(spec);
    const rp = spec.paths[RANGE];
    expect(rp.get).toBeDefined(); // preserved
    expect(rp.patch).toBeDefined();
    expect(rp.patch.operationId).toBe('set-excel-range');
    const props = rp.patch.requestBody.content['application/json'].schema.properties;
    expect(Object.keys(props).sort()).toEqual(['formulas', 'numberFormat', 'values']);
  });

  it('adds font, fill (PATCH) and sort/apply (POST) path items with address param', () => {
    const spec = minimalSpec();
    augmentExcelPaths(spec);
    const font = spec.paths[`${RANGE}/format/font`];
    const fill = spec.paths[`${RANGE}/format/fill`];
    const sort = spec.paths[`${RANGE}/sort/apply`];
    expect(font.patch.operationId).toBe('format-excel-range-font');
    expect(fill.patch.operationId).toBe('format-excel-range-fill');
    expect(sort.post.operationId).toBe('sort-excel-range');
    // path params cloned (including address) onto each new path item
    for (const p of [font, fill, sort]) {
      const names = p.parameters.map((x) => x.name);
      expect(names).toContain('address');
      expect(names).toContain('drive-id');
    }
    // font schema carries the expected formatting properties
    expect(
      Object.keys(font.patch.requestBody.content['application/json'].schema.properties).sort()
    ).toEqual(['bold', 'color', 'italic', 'name', 'size', 'underline']);
    // sort requires fields
    expect(sort.post.requestBody.content['application/json'].schema.required).toContain('fields');
  });

  it('throws if the base range path is missing', () => {
    expect(() => augmentExcelPaths({ paths: {} })).toThrow(/base range path/);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run test/excel-augmentations.test.ts`
Expected: FAIL — cannot import `augmentExcelPaths` (module does not exist).

- [ ] **Step 3: Create the augmentation module**

Create `bin/modules/excel-augmentations.mjs`:

```js
// Injects Excel workbook write operations that Microsoft's Graph OpenAPI metadata
// omits, so the declarative endpoints.json/generator pipeline can expose them.
// Called from simplified-openapi.mjs before the spec is trimmed.

const RANGE =
  "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')";

// 2-D array of arbitrary cell values (string | number | boolean | null per cell).
const TWO_D_ARRAY = { type: 'array', items: { type: 'array', items: {} } };

function clone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function buildOperation({ operationId, summary, schema }) {
  return {
    tags: ['drives.driveItem'],
    summary,
    operationId,
    requestBody: {
      description: 'Augmented Excel write body',
      required: true,
      content: { 'application/json': { schema } },
    },
    responses: {
      '2XX': {
        description: 'Success',
        content: { 'application/json': { schema: { type: 'object' } } },
      },
      '4XX': { $ref: '#/components/responses/error' },
      '5XX': { $ref: '#/components/responses/error' },
    },
    'x-ms-docs-operation-type': 'operation',
  };
}

export function augmentExcelPaths(openApiSpec) {
  const paths = openApiSpec.paths || {};
  const rangePath = paths[RANGE];
  if (!rangePath || !rangePath.parameters) {
    throw new Error(`augmentExcelPaths: base range path not found in spec: ${RANGE}`);
  }
  // Path-level params for the range: drive-id, driveItem-id, workbookWorksheet-id, address.
  const params = clone(rangePath.parameters);

  // 1. PATCH range(address) — set values / formulas / number formats. Merge into the
  //    existing path item (which currently only has GET) so GET is preserved.
  rangePath.patch = buildOperation({
    operationId: 'set-excel-range',
    summary: 'Set values, formulas, or number formats on a range',
    schema: {
      type: 'object',
      properties: { values: TWO_D_ARRAY, formulas: TWO_D_ARRAY, numberFormat: TWO_D_ARRAY },
    },
  });

  // 2. PATCH range(address)/format/font
  paths[`${RANGE}/format/font`] = {
    parameters: clone(params),
    patch: buildOperation({
      operationId: 'format-excel-range-font',
      summary: 'Set font properties on a range',
      schema: {
        type: 'object',
        properties: {
          name: { type: 'string' },
          size: { type: 'number' },
          color: { type: 'string' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          underline: { type: 'string' },
        },
      },
    }),
  };

  // 3. PATCH range(address)/format/fill
  paths[`${RANGE}/format/fill`] = {
    parameters: clone(params),
    patch: buildOperation({
      operationId: 'format-excel-range-fill',
      summary: 'Set fill (background) color on a range',
      schema: { type: 'object', properties: { color: { type: 'string' } } },
    }),
  };

  // 4. POST range(address)/sort/apply
  paths[`${RANGE}/sort/apply`] = {
    parameters: clone(params),
    post: buildOperation({
      operationId: 'sort-excel-range',
      summary: 'Sort a range by one or more columns',
      schema: {
        type: 'object',
        required: ['fields'],
        properties: {
          fields: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                key: { type: 'integer' },
                ascending: { type: 'boolean' },
                sortOn: { type: 'string' },
              },
            },
          },
          matchCase: { type: 'boolean' },
          hasHeaders: { type: 'boolean' },
          orientation: { type: 'string' },
        },
      },
    }),
  };

  return openApiSpec;
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run test/excel-augmentations.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 5: Wire augmentation into the generator**

In `bin/modules/simplified-openapi.mjs`, add the import at the top (after the existing imports, lines 1-2):

```js
import { augmentExcelPaths } from './excel-augmentations.mjs';
```

Then in `createAndSaveSimplifiedOpenAPI`, call it immediately after the spec is parsed and before the path-existence check (between the `const openApiSpec = yaml.load(spec);` line and the `for (const endpoint of endpoints)` existence loop):

```js
  const openApiSpec = yaml.load(spec);

  // Inject Excel write operations missing from Microsoft's metadata so the
  // endpoints.json entries below resolve against real path items.
  augmentExcelPaths(openApiSpec);

  for (const endpoint of endpoints) {
    if (!openApiSpec.paths[endpoint.pathPattern]) {
      throw new Error(`Path "${endpoint.pathPattern}" not found in OpenAPI spec.`);
    }
  }
```

- [ ] **Step 6: Commit**

```bash
npm run format
git add bin/modules/excel-augmentations.mjs bin/modules/simplified-openapi.mjs test/excel-augmentations.test.ts
git commit -m "feat: inject missing Excel write operations into OpenAPI spec

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 3: Excel endpoint entries + regenerate client

Add 20 new endpoint entries, repoint the 2 broken existing ones, regenerate the client, and assert the tools appear via an integration smoke test.

**Files:**
- Modify: `src/endpoints.json` (replace the 5 existing Excel entries' block with the full 22-entry set)
- Modify: `src/generated/client.ts` (regenerated — never hand-edited)
- Modify: `openapi/openapi-trimmed.yaml` (regenerated artifact)
- Test: `test/excel-tools-generated.test.ts`

**Interfaces:**
- Consumes: `augmentExcelPaths` (Task 2) — required for `format-excel-range-font`, `format-excel-range-fill`, `sort-excel-range` to resolve during generation.
- Produces: generated `api.endpoints` aliases for all tools listed below, with the stated HTTP methods.

- [ ] **Step 1: Replace the Excel block in `src/endpoints.json`**

Locate the 5 existing Excel entries (`create-excel-chart`, `format-excel-range`, `sort-excel-range`, `get-excel-range`, `list-excel-worksheets`) — currently a contiguous block. Replace that entire block with the following entries (keep `create-excel-chart` as-is at the top; the rest are new or modified). Preserve surrounding non-Excel entries and valid JSON (commas):

```json
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/charts/add",
    "method": "post",
    "toolName": "create-excel-chart",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"]
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')",
    "method": "get",
    "toolName": "get-excel-range",
    "isExcelOp": true,
    "scopes": ["Files.Read"],
    "skipEncoding": ["address"]
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')",
    "method": "patch",
    "toolName": "set-excel-range",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"],
    "llmTip": "Edit cells. Body sets values, formulas, and/or numberFormat as 2-D arrays matching the address dimensions, e.g. {\"values\": [[\"Name\", 1], [\"Bob\", 2]]}. address is A1-style like 'A1:B2' (worksheet already in the path). Use formulas for cells like \"=SUM(A1:A10)\". Pass workbookSessionId from create-excel-session to batch many edits efficiently."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/clear",
    "method": "post",
    "toolName": "clear-excel-range",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"],
    "llmTip": "Clears a range. Body: {\"applyTo\": \"All\"|\"Formats\"|\"Contents\"|\"Hyperlinks\"}. 'Contents' keeps formatting."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/insert",
    "method": "post",
    "toolName": "insert-excel-range",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"],
    "llmTip": "Inserts cells, shifting existing ones. Body: {\"shift\": \"Down\"|\"Right\"}."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/delete",
    "method": "post",
    "toolName": "delete-excel-range",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"],
    "llmTip": "Deletes cells, shifting remaining ones. Body: {\"shift\": \"Up\"|\"Left\"}."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/merge",
    "method": "post",
    "toolName": "merge-excel-range",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"],
    "llmTip": "Merges cells. Body: {\"across\": true} merges each row separately; false merges the whole range into one cell."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/unmerge",
    "method": "post",
    "toolName": "unmerge-excel-range",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"]
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/format",
    "method": "patch",
    "toolName": "format-excel-range",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"],
    "llmTip": "Range-level format: columnWidth, rowHeight, horizontalAlignment (General/Left/Center/Right), verticalAlignment (Top/Center/Bottom), wrapText (bool). For font use format-excel-range-font; for background color use format-excel-range-fill."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/format/font",
    "method": "patch",
    "toolName": "format-excel-range-font",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"],
    "llmTip": "Sets font on a range. Body fields (all optional): name, size (number), color (hex like \"#FF0000\"), bold, italic, underline (\"None\"|\"Single\"|\"Double\")."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/format/fill",
    "method": "patch",
    "toolName": "format-excel-range-fill",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"],
    "llmTip": "Sets cell background fill. Body: {\"color\": \"#FFFF00\"} (hex)."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/sort/apply",
    "method": "post",
    "toolName": "sort-excel-range",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "skipEncoding": ["address"],
    "llmTip": "Sorts a range. Body: {\"fields\": [{\"key\": 0, \"ascending\": true}], \"hasHeaders\": true}. key is the 0-based column index within the range."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets",
    "method": "get",
    "toolName": "list-excel-worksheets",
    "isExcelOp": true,
    "scopes": ["Files.Read"]
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/usedRange()",
    "method": "get",
    "toolName": "get-excel-used-range",
    "isExcelOp": true,
    "scopes": ["Files.Read"],
    "llmTip": "Returns the populated area of a worksheet (values, formulas, address) — use this before editing to see existing data and find the next empty row."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/add",
    "method": "post",
    "toolName": "add-excel-worksheet",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "llmTip": "Adds a worksheet. Body: {\"name\": \"Sheet name\"} (optional)."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}",
    "method": "patch",
    "toolName": "update-excel-worksheet",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "llmTip": "Rename/move/hide a worksheet. Body fields (optional): name, position (0-based), visibility (\"Visible\"|\"Hidden\"|\"VeryHidden\")."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}",
    "method": "delete",
    "toolName": "delete-excel-worksheet",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"]
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/tables/add",
    "method": "post",
    "toolName": "add-excel-table",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "llmTip": "Creates a table on a worksheet. Body: {\"address\": \"A1:C10\", \"hasHeaders\": true}."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/tables/{workbookTable-id}/rows/add",
    "method": "post",
    "toolName": "add-excel-table-row",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "llmTip": "Appends/inserts table rows. Body: {\"values\": [[\"a\", 1], [\"b\", 2]], \"index\": null}. index null appends to the end. Tables are addressed by id at the workbook level (no worksheet in path); list tables to get the id."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/tables/{workbookTable-id}/columns/add",
    "method": "post",
    "toolName": "add-excel-table-column",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "llmTip": "Adds a table column. Body: {\"name\": \"Header\", \"values\": [[\"x\"], [\"y\"]], \"index\": null}."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/tables/{workbookTable-id}",
    "method": "patch",
    "toolName": "update-excel-table",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "llmTip": "Updates table properties. Body fields (optional): name, showHeaders (bool), showTotals (bool), style (e.g. \"TableStyleMedium2\")."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/tables/{workbookTable-id}",
    "method": "delete",
    "toolName": "delete-excel-table",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"]
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/createSession",
    "method": "post",
    "toolName": "create-excel-session",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "llmTip": "Starts a workbook session for fast batch edits. Body: {\"persistChanges\": true} writes changes to the file; false is a sandbox. Returns an id — pass it as workbookSessionId on subsequent Excel tools, then call close-excel-session."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/closeSession",
    "method": "post",
    "toolName": "close-excel-session",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "llmTip": "Closes a workbook session. Pass the session id as workbookSessionId."
  },
  {
    "pathPattern": "/drives/{drive-id}/items/{driveItem-id}/workbook/refreshSession",
    "method": "post",
    "toolName": "refresh-excel-session",
    "isExcelOp": true,
    "scopes": ["Files.ReadWrite"],
    "llmTip": "Keeps a workbook session alive. Pass the session id as workbookSessionId."
  },
```

- [ ] **Step 2: Validate JSON**

Run: `node -e "JSON.parse(require('fs').readFileSync('src/endpoints.json','utf8')); console.log('valid')"`
Expected: `valid`

- [ ] **Step 3: Write the failing smoke test**

Create `test/excel-tools-generated.test.ts`:

```ts
import { describe, it, expect } from 'vitest';
import { api } from '../src/generated/client.js';

const EXPECTED: Record<string, string> = {
  'set-excel-range': 'patch',
  'clear-excel-range': 'post',
  'insert-excel-range': 'post',
  'delete-excel-range': 'post',
  'merge-excel-range': 'post',
  'unmerge-excel-range': 'post',
  'format-excel-range': 'patch',
  'format-excel-range-font': 'patch',
  'format-excel-range-fill': 'patch',
  'sort-excel-range': 'post',
  'get-excel-used-range': 'get',
  'add-excel-worksheet': 'post',
  'update-excel-worksheet': 'patch',
  'delete-excel-worksheet': 'delete',
  'add-excel-table': 'post',
  'add-excel-table-row': 'post',
  'add-excel-table-column': 'post',
  'update-excel-table': 'patch',
  'delete-excel-table': 'delete',
  'create-excel-session': 'post',
  'close-excel-session': 'post',
  'refresh-excel-session': 'post',
};

describe('generated Excel tools', () => {
  it('exposes every new/modified Excel tool with the expected method', () => {
    for (const [alias, method] of Object.entries(EXPECTED)) {
      const ep = api.endpoints.find((e: any) => e.alias === alias);
      expect(ep, `missing generated tool: ${alias}`).toBeDefined();
      expect(ep!.method.toLowerCase(), `wrong method for ${alias}`).toBe(method);
    }
  });

  it('set-excel-range exposes a body parameter', () => {
    const ep = api.endpoints.find((e: any) => e.alias === 'set-excel-range');
    const hasBody = (ep!.parameters || []).some((p: any) => p.type === 'Body');
    expect(hasBody).toBe(true);
  });
});
```

- [ ] **Step 4: Run the smoke test to verify it fails**

Run: `npx vitest run test/excel-tools-generated.test.ts`
Expected: FAIL — the new aliases are not yet in the (stale) generated client.

- [ ] **Step 5: Regenerate the client**

Run: `npm run generate`
Expected: completes with "Successfully generated client code" and no thrown "Path ... not found" error. (If it throws on a `font`/`fill`/`sort/apply` path, Task 2's augmentation wiring is missing — fix before continuing.)

- [ ] **Step 6: Run the smoke test to verify it passes**

Run: `npx vitest run test/excel-tools-generated.test.ts`
Expected: PASS (2 tests).

- [ ] **Step 7: Run the full suite**

Run: `npm test`
Expected: PASS (no regressions in existing tests).

- [ ] **Step 8: Commit**

```bash
npm run format
git add src/endpoints.json src/generated/ openapi/openapi-trimmed.yaml test/excel-tools-generated.test.ts
git commit -m "feat: add fine-grained Excel editing tools (values, format, tables, sessions)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 4: Excel tool behavior tests

Validate the Excel-specific runtime behavior with synthetic endpoints (mocked client): addressed-range PATCH sends the body, `address` is not URL-encoded, POST actions send their body, and writes are filtered out in read-only mode.

**Files:**
- Test: `src/__tests__/graph-tools.test.ts` (append a new `describe` block)

**Interfaces:**
- Consumes: `registerGraphTools`, the test helpers `makeEndpoint`/`makeConfig`/`createMockGraphClient`/`createMockServer`/`loadModule` already in the file.

- [ ] **Step 1: Write the failing tests**

Append inside the top-level `describe('graph-tools', ...)` (before its closing `});`):

```ts
  // ---- 9. Excel range editing behavior ----
  describe('excel range editing', () => {
    function setRangeEndpoint() {
      const endpoint = makeEndpoint({
        alias: 'set-excel-range',
        method: 'patch',
        path: "/drives/:driveId/items/:driveItemId/workbook/worksheets/:workbookWorksheetId/range(address=':address')",
        parameters: [
          { name: 'driveId', type: 'Path', schema: z.string() },
          { name: 'driveItemId', type: 'Path', schema: z.string() },
          { name: 'workbookWorksheetId', type: 'Path', schema: z.string() },
          { name: 'address', type: 'Path', schema: z.string() },
          { name: 'body', type: 'Body', schema: z.object({ values: z.any() }).passthrough() },
        ],
      });
      const config = makeConfig({
        toolName: 'set-excel-range',
        method: 'patch',
        pathPattern:
          "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')",
        scopes: ['Files.ReadWrite'],
        isExcelOp: true,
        skipEncoding: ['address'],
      });
      return { endpoint, config };
    }

    it('PATCHes the addressed range with the body and leaves the address un-encoded', async () => {
      const { endpoint, config } = setRangeEndpoint();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ address: 'Sheet1!A1:B2' }) }] },
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('set-excel-range');
      await tool!.handler({
        driveId: 'd1',
        driveItemId: 'item1',
        workbookWorksheetId: 'ws1',
        address: 'A1:B2',
        body: { values: [['Name', 1]] },
      });

      const [requestedPath, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.method).toBe('PATCH');
      expect(options.body).toBe('{"values":[["Name",1]]}');
      // address contains ':' — skipEncoding keeps it literal (no %3A)
      expect(requestedPath).toContain("range(address='A1:B2')");
      expect(requestedPath).not.toContain('%3A');
    });

    it('POSTs a clear action body to the addressed range', async () => {
      const endpoint = makeEndpoint({
        alias: 'clear-excel-range',
        method: 'post',
        path: "/drives/:driveId/items/:driveItemId/workbook/worksheets/:workbookWorksheetId/range(address=':address')/clear",
        parameters: [
          { name: 'driveId', type: 'Path', schema: z.string() },
          { name: 'driveItemId', type: 'Path', schema: z.string() },
          { name: 'workbookWorksheetId', type: 'Path', schema: z.string() },
          { name: 'address', type: 'Path', schema: z.string() },
          { name: 'body', type: 'Body', schema: z.object({ applyTo: z.string() }).passthrough() },
        ],
      });
      const config = makeConfig({
        toolName: 'clear-excel-range',
        method: 'post',
        pathPattern:
          "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')/clear",
        scopes: ['Files.ReadWrite'],
        isExcelOp: true,
        skipEncoding: ['address'],
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({}) }] },
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('clear-excel-range');
      await tool!.handler({
        driveId: 'd1',
        driveItemId: 'item1',
        workbookWorksheetId: 'ws1',
        address: 'A1:B2',
        body: { applyTo: 'Contents' },
      });

      const [requestedPath, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.method).toBe('POST');
      expect(options.body).toBe('{"applyTo":"Contents"}');
      expect(requestedPath).toContain("range(address='A1:B2')/clear");
    });

    it('is filtered out in read-only mode (PATCH is non-GET)', async () => {
      const { endpoint, config } = setRangeEndpoint();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any, /* readOnly */ true);

      expect(server.tools.has('set-excel-range')).toBe(false);
    });
  });
```

- [ ] **Step 2: Run to verify it passes**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "excel range editing"`
Expected: PASS (3 tests). (These exercise already-implemented behavior from Tasks 1 & 3; they encode the Excel-specific contract.)

- [ ] **Step 3: Commit**

```bash
npm run format
git add src/__tests__/graph-tools.test.ts
git commit -m "test: Excel range editing behavior (addressed PATCH, skipEncoding, read-only)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 5: README documentation

Document the new Excel editing surface and the session workflow.

**Files:**
- Modify: `README.md` (the **Excel Operations** entry around lines 112-113)

**Interfaces:** none.

- [ ] **Step 1: Replace the Excel Operations tool list**

In `README.md`, replace the existing two-line Excel Operations block:

```markdown
**Excel Operations**  
<sub>list-excel-worksheets, get-excel-range, create-excel-chart, format-excel-range, sort-excel-range</sub>
```

with:

```markdown
**Excel Operations**  
<sub>list-excel-worksheets, get-excel-range, get-excel-used-range, set-excel-range, clear-excel-range, insert-excel-range, delete-excel-range, merge-excel-range, unmerge-excel-range, format-excel-range, format-excel-range-font, format-excel-range-fill, sort-excel-range, create-excel-chart, add-excel-worksheet, update-excel-worksheet, delete-excel-worksheet, add-excel-table, add-excel-table-row, add-excel-table-column, update-excel-table, delete-excel-table, create-excel-session, close-excel-session, refresh-excel-session</sub>

> **Fine-grained Excel editing.** `set-excel-range` writes cell values, formulas, and number formats to an A1-style range; `format-excel-range`/`-font`/`-fill` control alignment, font, and fill. For a batch of edits, call `create-excel-session` (with `persistChanges: true`), pass the returned id as `workbookSessionId` on each subsequent Excel tool, then `close-excel-session` — this is faster and keeps the edits consistent.
>
> _Note: Microsoft Graph offers no equivalent fine-grained editing API for Word/`.docx` files — only whole-file operations — so Word editing is not provided._
```

- [ ] **Step 2: Verify formatting**

Run: `npm run format:check`
Expected: passes (or run `npm run format` then re-check).

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: document fine-grained Excel editing tools and sessions

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Final verification

- [ ] Run `npm run verify` (generate + lint + format:check + build + test). Expected: all green.
- [ ] Spot-check: `grep -c '"toolName": ".*excel' src/endpoints.json` shows 25 Excel tool entries (5 pre-existing incl. chart + 20 new).

## Notes for the implementer

- **Why augmentation, not hand-written tools:** the generator throws if an `endpoints.json` path is absent from the spec and silently drops a tool if the method is absent. Microsoft's metadata exposes only `GET` on `range(address=…)` and has no font/fill/sort-apply paths. Injecting them (Task 2) lets all the cross-cutting machinery (read-only filter, multi-account, body auto-wrap, error handling, session header) apply uniformly.
- **`range()` vs `range(address='…')`:** the two pre-existing tools used the empty-parens `range()` form, which cannot target a specific range. Task 3 repoints `format-excel-range` to the addressed `…/format` path (native PATCH) and `sort-excel-range` to the augmented `…/sort/apply` POST.
- **Tables are workbook-scoped for item ops:** `add-excel-table` is worksheet-scoped (`…/worksheets/{id}/tables/add`), but row/column/update/delete address the table directly at `…/workbook/tables/{workbookTable-id}/…`. Use `list` to obtain a table id (an existing list-tables tool is out of scope here; `get-excel-used-range` + Graph table listing can be added later if needed).
- **Regeneration is deterministic and offline:** never pass `--force` to the generator; the cached spec is pinned.
```
