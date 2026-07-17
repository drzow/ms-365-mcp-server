# Group Calendar Tools Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add nine tools that let the server discover Microsoft 365 groups and read/write a group's calendar (the immediate need: create events on the Aleaquo group calendar).

**Architecture:** Tools are declared as entries in `src/endpoints.json`; a generator (`npm run generate`) downloads the Graph OpenAPI, trims it to the referenced paths, and emits the zod client. We add nine entries (all `workScopes`-gated so they only appear in org mode), regenerate, and test the config + gating. No changes to `auth.ts`/`graph-tools.ts` are expected because scope derivation and org-mode gating already read `workScopes` generically. A contingency augmentation module is included in case any group path is absent from Microsoft's metadata.

**Tech Stack:** TypeScript (ESM), zod, `@modelcontextprotocol/sdk`, vitest, MSAL. Generator is Node ESM in `bin/`.

## Global Constraints

- Node `>=18` (engines); code is ESM (`"type": "module"`) — imports use `.js` extensions.
- `src/endpoints.json` is the ONLY git-tracked build artifact. `openapi/*.yaml` and `src/generated/client.ts` are git-ignored and regenerated — never commit them.
- Group tools are work-account features: declare `workScopes` (NOT `scopes`) so they are org-mode-gated and their OAuth scopes are only requested in org mode.
- Derived scope set for this feature: `Group.Read.All`, `Group.ReadWrite.All`. Do not add any other scopes.
- Do not hardcode tenant/client IDs or tokens anywhere; they come from `.env` already.
- Prettier + eslint must pass (`npm run format:check`, `npm run lint`). Run `npm run format` before committing.

---

### Task 1: De-risk — regenerate against the full upstream spec and verify group paths

**Files:**

- No source changes. Touches (git-ignored) `openapi/openapi.yaml`, `openapi/openapi-trimmed.yaml`, `src/generated/client.ts`.

**Interfaces:**

- Consumes: nothing.
- Produces: a verified answer to "do all nine group paths exist in the upstream Graph v1.0 metadata?" — which decides whether Task 1b (augmentation) is needed.

**Why this task exists:** The committed `openapi/openapi.yaml` is a stale partial (1302 paths) that is missing 125 of the 158 paths currently in `endpoints.json`. `npm run generate` therefore throws against the local file and must re-download the full spec. This task does that download and confirms the group paths are present before we wire them in.

- [ ] **Step 1: Re-download the full upstream spec and regenerate the client from the CURRENT (unchanged) endpoints.json**

Run:

```bash
npm run generate -- --force
```

Expected: completes with "✅ Successfully generated client code". This proves the pipeline works and refreshes `openapi/openapi.yaml` to the full upstream metadata. (This step makes NO endpoints.json change yet.)

- [ ] **Step 2: Verify the nine group paths exist as path keys in the freshly downloaded spec**

Run:

```bash
for p in \
  "/me/memberOf" \
  "/groups" \
  "/groups/{group-id}" \
  "/groups/{group-id}/calendar/events" \
  "/groups/{group-id}/calendar/calendarView" \
  "/groups/{group-id}/calendar/events/{event-id}" ; do
  grep -qF "  $p:" openapi/openapi.yaml && echo "FOUND   $p" || echo "MISSING $p"
done
```

Expected: all six FOUND. (These six path keys cover all nine tools — list and create share `/calendar/events`; get/update/delete share `/calendar/events/{event-id}`.)

- [ ] **Step 3: Decision**

- If ALL six are FOUND → skip Task 1b entirely; proceed to Task 2.
- If ANY is MISSING → do Task 1b (augmentation) before Task 2, and note which paths were missing.

- [ ] **Step 4: No commit**

Only git-ignored files changed. Nothing to commit. Confirm with:

```bash
git status --porcelain
```

Expected: empty (or only untracked git-ignored files, which won't be listed).

---

### Task 1b (CONDITIONAL — only if Task 1 Step 2 reported a MISSING path): inject group calendar paths via augmentation

**Files:**

- Create: `bin/modules/group-augmentations.mjs`
- Modify: `bin/modules/simplified-openapi.mjs` (call the new augmenter next to `augmentExcelPaths`)
- Test: `test/group-augmentations.test.ts`

**Interfaces:**

- Consumes: the loaded `openApiSpec` object (mutated in place), same contract as `augmentExcelPaths`.
- Produces: `augmentGroupCalendarPaths(openApiSpec)` — ensures the group calendar path items exist so `endpoints.json` entries resolve.

**Only build this if Task 1 found a missing path.** If everything was FOUND, these paths come from Microsoft's metadata and this module is unnecessary.

- [ ] **Step 1: Write the failing test**

Create `test/group-augmentations.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { augmentGroupCalendarPaths } from '../bin/modules/group-augmentations.mjs';

describe('augmentGroupCalendarPaths', () => {
  it('adds group calendar path items with get/post/patch/delete operations', () => {
    const spec: any = { paths: {} };
    augmentGroupCalendarPaths(spec);

    const events = spec.paths['/groups/{group-id}/calendar/events'];
    expect(events).toBeDefined();
    expect(events.get).toBeDefined();
    expect(events.post).toBeDefined();

    const oneEvent = spec.paths['/groups/{group-id}/calendar/events/{event-id}'];
    expect(oneEvent.get).toBeDefined();
    expect(oneEvent.patch).toBeDefined();
    expect(oneEvent.delete).toBeDefined();

    const view = spec.paths['/groups/{group-id}/calendar/calendarView'];
    expect(view.get).toBeDefined();
  });

  it('does not clobber an already-present path item', () => {
    const spec: any = {
      paths: { '/groups/{group-id}/calendar/events': { get: { operationId: 'preexisting' } } },
    };
    augmentGroupCalendarPaths(spec);
    // Existing get preserved; post added alongside it.
    expect(spec.paths['/groups/{group-id}/calendar/events'].get.operationId).toBe('preexisting');
    expect(spec.paths['/groups/{group-id}/calendar/events'].post).toBeDefined();
  });
});
```

- [ ] **Step 2: Run it to verify it fails**

Run: `npx vitest run test/group-augmentations.test.ts`
Expected: FAIL — cannot resolve `../bin/modules/group-augmentations.mjs`.

- [ ] **Step 3: Create the augmentation module**

Create `bin/modules/group-augmentations.mjs`:

```javascript
// Injects Microsoft 365 group calendar path items in case Microsoft's Graph
// OpenAPI metadata omits them, so the declarative endpoints.json/generator
// pipeline can expose group calendar tools. Called from simplified-openapi.mjs
// before the spec is trimmed. Mirrors excel-augmentations.mjs.

const GROUP_ID_PARAM = {
  name: 'group-id',
  in: 'path',
  required: true,
  description: 'The unique identifier of the group.',
  schema: { type: 'string' },
};
const EVENT_ID_PARAM = {
  name: 'event-id',
  in: 'path',
  required: true,
  description: 'The unique identifier of the event.',
  schema: { type: 'string' },
};
const DATE_TIME_TZ = {
  type: 'object',
  properties: { dateTime: { type: 'string' }, timeZone: { type: 'string' } },
};
const EVENT_SCHEMA = {
  type: 'object',
  properties: {
    subject: { type: 'string' },
    body: {
      type: 'object',
      properties: { contentType: { type: 'string' }, content: { type: 'string' } },
    },
    start: DATE_TIME_TZ,
    end: DATE_TIME_TZ,
    location: { type: 'object', properties: { displayName: { type: 'string' } } },
    attendees: { type: 'array', items: {} },
    isAllDay: { type: 'boolean' },
    recurrence: { type: 'object' },
  },
};

function getOp(operationId, summary, extraParams = []) {
  return {
    tags: ['groups.calendar'],
    summary,
    operationId,
    parameters: [GROUP_ID_PARAM, ...extraParams],
    responses: {
      '2XX': { description: 'Success', content: { 'application/json': { schema: EVENT_SCHEMA } } },
      '4XX': { $ref: '#/components/responses/error' },
      '5XX': { $ref: '#/components/responses/error' },
    },
    'x-ms-docs-operation-type': 'operation',
  };
}

function writeOp(operationId, summary, extraParams = []) {
  return {
    tags: ['groups.calendar'],
    summary,
    operationId,
    parameters: [GROUP_ID_PARAM, ...extraParams],
    requestBody: {
      description: 'Group calendar event',
      required: true,
      content: { 'application/json': { schema: EVENT_SCHEMA } },
    },
    responses: {
      '2XX': { description: 'Success', content: { 'application/json': { schema: EVENT_SCHEMA } } },
      '4XX': { $ref: '#/components/responses/error' },
      '5XX': { $ref: '#/components/responses/error' },
    },
    'x-ms-docs-operation-type': 'operation',
  };
}

export function augmentGroupCalendarPaths(openApiSpec) {
  const paths = (openApiSpec.paths = openApiSpec.paths || {});

  const eventsKey = '/groups/{group-id}/calendar/events';
  paths[eventsKey] = paths[eventsKey] || {};
  paths[eventsKey].get =
    paths[eventsKey].get || getOp('groups.ListEvents', 'List group calendar events');
  paths[eventsKey].post =
    paths[eventsKey].post || writeOp('groups.CreateEvent', 'Create a group calendar event');

  const oneEventKey = '/groups/{group-id}/calendar/events/{event-id}';
  paths[oneEventKey] = paths[oneEventKey] || {};
  paths[oneEventKey].get =
    paths[oneEventKey].get ||
    getOp('groups.GetEvent', 'Get a group calendar event', [EVENT_ID_PARAM]);
  paths[oneEventKey].patch =
    paths[oneEventKey].patch ||
    writeOp('groups.UpdateEvent', 'Update a group calendar event', [EVENT_ID_PARAM]);
  paths[oneEventKey].delete = paths[oneEventKey].delete || {
    tags: ['groups.calendar'],
    summary: 'Delete a group calendar event',
    operationId: 'groups.DeleteEvent',
    parameters: [GROUP_ID_PARAM, EVENT_ID_PARAM],
    responses: {
      '2XX': { description: 'Success' },
      '4XX': { $ref: '#/components/responses/error' },
      '5XX': { $ref: '#/components/responses/error' },
    },
    'x-ms-docs-operation-type': 'operation',
  };

  const viewKey = '/groups/{group-id}/calendar/calendarView';
  paths[viewKey] = paths[viewKey] || {};
  paths[viewKey].get =
    paths[viewKey].get ||
    getOp('groups.CalendarView', 'Get expanded group calendar view', [
      { name: 'startDateTime', in: 'query', required: true, schema: { type: 'string' } },
      { name: 'endDateTime', in: 'query', required: true, schema: { type: 'string' } },
    ]);
}
```

Note: only inject the specific keys Task 1 reported MISSING; the `|| {}` / `|| getOp(...)` guards make it safe to run even when Microsoft already defines some of them.

- [ ] **Step 4: Wire it into the generator**

In `bin/modules/simplified-openapi.mjs`, add the import at the top alongside the Excel one:

```javascript
import { augmentGroupCalendarPaths } from './group-augmentations.mjs';
```

And call it right after `augmentExcelPaths(openApiSpec);` inside `createAndSaveSimplifiedOpenAPI`:

```javascript
augmentExcelPaths(openApiSpec);
augmentGroupCalendarPaths(openApiSpec);
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `npx vitest run test/group-augmentations.test.ts`
Expected: PASS (both tests).

- [ ] **Step 6: Commit**

```bash
npm run format
git add bin/modules/group-augmentations.mjs bin/modules/simplified-openapi.mjs test/group-augmentations.test.ts
git commit -m "feat: inject group calendar paths when absent from Graph metadata

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 2: Add the nine group endpoints to endpoints.json and regenerate

**Files:**

- Modify: `src/endpoints.json` (add nine entries)
- Test: `test/group-calendar-tools.test.ts` (create)

**Interfaces:**

- Consumes: the generator pipeline (Task 1 confirmed it works; Task 1b augmentation if it was needed).
- Produces: nine registered tools — `list-my-groups`, `list-groups`, `get-group`, `list-group-calendar-events`, `get-group-calendar-view`, `get-group-calendar-event`, `create-group-calendar-event`, `update-group-calendar-event`, `delete-group-calendar-event` — and the derived scopes `Group.Read.All`, `Group.ReadWrite.All` in org mode.

- [ ] **Step 1: Write the failing test**

Create `test/group-calendar-tools.test.ts`:

```typescript
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { buildScopesFromEndpoints } from '../src/auth.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const endpoints = JSON.parse(
  readFileSync(path.join(__dirname, '..', 'src', 'endpoints.json'), 'utf8')
) as Array<{
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
}>;

function find(toolName: string) {
  const e = endpoints.find((x) => x.toolName === toolName);
  expect(e, `endpoint ${toolName} should exist`).toBeDefined();
  return e!;
}

describe('group calendar endpoints.json entries', () => {
  const expected: Array<[string, string, string, string]> = [
    ['list-my-groups', 'get', '/me/memberOf', 'Group.Read.All'],
    ['list-groups', 'get', '/groups', 'Group.Read.All'],
    ['get-group', 'get', '/groups/{group-id}', 'Group.Read.All'],
    ['list-group-calendar-events', 'get', '/groups/{group-id}/calendar/events', 'Group.Read.All'],
    [
      'get-group-calendar-view',
      'get',
      '/groups/{group-id}/calendar/calendarView',
      'Group.Read.All',
    ],
    [
      'get-group-calendar-event',
      'get',
      '/groups/{group-id}/calendar/events/{event-id}',
      'Group.Read.All',
    ],
    [
      'create-group-calendar-event',
      'post',
      '/groups/{group-id}/calendar/events',
      'Group.ReadWrite.All',
    ],
    [
      'update-group-calendar-event',
      'patch',
      '/groups/{group-id}/calendar/events/{event-id}',
      'Group.ReadWrite.All',
    ],
    [
      'delete-group-calendar-event',
      'delete',
      '/groups/{group-id}/calendar/events/{event-id}',
      'Group.ReadWrite.All',
    ],
  ];

  it.each(expected)(
    '%s is declared with method %s, path %s, workScope %s',
    (toolName, method, pathPattern, workScope) => {
      const e = find(toolName);
      expect(e.method).toBe(method);
      expect(e.pathPattern).toBe(pathPattern);
      // Group tools MUST use workScopes (org-mode gated), never plain scopes.
      expect(e.scopes, `${toolName} must not use plain scopes`).toBeUndefined();
      expect(e.workScopes).toContain(workScope);
    }
  );
});

describe('scope derivation', () => {
  it('includes Group scopes only in org mode', () => {
    const orgScopes = buildScopesFromEndpoints(true);
    expect(orgScopes).toContain('Group.Read.All');
    expect(orgScopes).toContain('Group.ReadWrite.All');

    const personalScopes = buildScopesFromEndpoints(false);
    expect(personalScopes).not.toContain('Group.Read.All');
    expect(personalScopes).not.toContain('Group.ReadWrite.All');
  });
});
```

- [ ] **Step 2: Run it to verify it fails**

Run: `npx vitest run test/group-calendar-tools.test.ts`
Expected: FAIL — endpoints not found / scopes missing.

- [ ] **Step 3: Add the nine entries to `src/endpoints.json`**

Insert this block just before the closing `]` of the array (append as the last entries; JSON array order does not matter):

```json
,
  {
    "pathPattern": "/me/memberOf",
    "method": "get",
    "toolName": "list-my-groups",
    "workScopes": ["Group.Read.All"],
    "llmTip": "Lists the groups and directory roles the signed-in user belongs to. Keep only objects where @odata.type is '#microsoft.graph.group'. Use $select=id,displayName,mail to reduce noise. Microsoft 365 group calendars belong to groups whose groupTypes contains 'Unified'."
  },
  {
    "pathPattern": "/groups",
    "method": "get",
    "toolName": "list-groups",
    "workScopes": ["Group.Read.All"],
    "llmTip": "Find a group's id. By mail: $filter=mail eq 'Aleaquo@scalesology.com'. By name: $filter=startswith(displayName,'Alea'). Always $select=id,displayName,mail,groupTypes. Unified (Microsoft 365) groups have groupTypes containing 'Unified' and are the ones with a group calendar."
  },
  {
    "pathPattern": "/groups/{group-id}",
    "method": "get",
    "toolName": "get-group",
    "workScopes": ["Group.Read.All"],
    "llmTip": "Get a single group by id. Use list-groups to find the id. Recommended $select=id,displayName,mail,groupTypes,visibility."
  },
  {
    "pathPattern": "/groups/{group-id}/calendar/events",
    "method": "get",
    "toolName": "list-group-calendar-events",
    "workScopes": ["Group.Read.All"],
    "supportsTimezone": true,
    "llmTip": "Lists events on a group's calendar. WARNING: does NOT expand recurring events — returns seriesMaster only. Use get-group-calendar-view with a date range for expanded instances. Get group-id from list-groups. Recommended $select=id,subject,start,end,organizer,location and $orderby=start/dateTime."
  },
  {
    "pathPattern": "/groups/{group-id}/calendar/calendarView",
    "method": "get",
    "toolName": "get-group-calendar-view",
    "workScopes": ["Group.Read.All"],
    "supportsTimezone": true,
    "llmTip": "Returns expanded recurring event instances on a group's calendar within a date range. Requires startDateTime and endDateTime query params in ISO 8601 (e.g. 2026-07-01T00:00:00Z). Get group-id from list-groups."
  },
  {
    "pathPattern": "/groups/{group-id}/calendar/events/{event-id}",
    "method": "get",
    "toolName": "get-group-calendar-event",
    "workScopes": ["Group.Read.All"],
    "supportsTimezone": true,
    "llmTip": "Gets a single event from a group's calendar. Get group-id from list-groups and event-id from list-group-calendar-events or get-group-calendar-view."
  },
  {
    "pathPattern": "/groups/{group-id}/calendar/events",
    "method": "post",
    "toolName": "create-group-calendar-event",
    "workScopes": ["Group.ReadWrite.All"],
    "supportsTimezone": true,
    "llmTip": "Creates an event owned by the group mailbox (survives individual accounts). Get group-id from list-groups. Body example: {\"subject\":\"...\",\"start\":{\"dateTime\":\"2026-07-10T15:00:00\",\"timeZone\":\"America/New_York\"},\"end\":{\"dateTime\":\"2026-07-10T16:00:00\",\"timeZone\":\"America/New_York\"},\"body\":{\"contentType\":\"HTML\",\"content\":\"...\"},\"location\":{\"displayName\":\"...\"},\"attendees\":[{\"emailAddress\":{\"address\":\"a@b.com\",\"name\":\"A\"},\"type\":\"required\"}]}. Requires org mode and Group.ReadWrite.All admin consent — a 403 means that consent or org mode is missing."
  },
  {
    "pathPattern": "/groups/{group-id}/calendar/events/{event-id}",
    "method": "patch",
    "toolName": "update-group-calendar-event",
    "workScopes": ["Group.ReadWrite.All"],
    "supportsTimezone": true,
    "llmTip": "Updates fields on an existing group calendar event. Send only the fields to change. Get ids from list-group-calendar-events."
  },
  {
    "pathPattern": "/groups/{group-id}/calendar/events/{event-id}",
    "method": "delete",
    "toolName": "delete-group-calendar-event",
    "workScopes": ["Group.ReadWrite.All"],
    "llmTip": "Deletes an event from a group's calendar. This is irreversible. Get ids from list-group-calendar-events."
  }
```

- [ ] **Step 4: Regenerate the client and build**

Run:

```bash
npm run generate && npm run build
```

Expected: both succeed. If `generate` throws "Path ... not found in OpenAPI spec" for a group calendar path, Task 1b (augmentation) was required and skipped — go do it, then re-run.

- [ ] **Step 5: Run the config test to verify it passes**

Run: `npx vitest run test/group-calendar-tools.test.ts`
Expected: PASS (all 9 parametrized cases + scope derivation).

- [ ] **Step 6: Confirm the tools actually generated into the client**

Run:

```bash
grep -c "create-group-calendar-event\|list-groups\|get-group-calendar-view" src/generated/client.ts
```

Expected: a count `>= 3` (client.ts is git-ignored; this just confirms generation worked).

- [ ] **Step 7: Commit (endpoints.json only — client.ts is git-ignored)**

```bash
npm run format
git add src/endpoints.json test/group-calendar-tools.test.ts
git commit -m "feat: add group discovery and group calendar tools

Nine org-mode tools: list-my-groups, list-groups, get-group, and full
group calendar CRUD (list/view/get/create/update/delete). Scopes
Group.Read.All + Group.ReadWrite.All, derived automatically.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 3: Unit-test org-mode gating for a workScopes-only tool

**Files:**

- Modify: `src/__tests__/graph-tools.test.ts` (add a `describe` block)

**Interfaces:**

- Consumes: `registerGraphTools(server, graphClient, readOnly, enabledToolsPattern, orgMode)` — the 5th positional arg is `orgMode`.
- Produces: regression coverage guaranteeing group tools (workScopes-only) are hidden unless org mode is on.

**Why:** The gating logic already exists in `graph-tools.ts` (`if (!orgMode && ... !scopes && workScopes) skip`). This locks the contract so a future refactor can't silently expose group tools (and their scopes) outside org mode.

- [ ] **Step 1: Add the failing/guard test**

Append this `describe` block inside the top-level `describe('graph-tools', ...)` in `src/__tests__/graph-tools.test.ts` (place it just before the final closing `});` of that block):

```typescript
// ---- 11. org-mode gating for workScopes-only tools ----
describe('org-mode gating', () => {
  function groupToolEndpointAndConfig() {
    const endpoint = makeEndpoint({
      alias: 'list-groups',
      method: 'get',
      path: '/groups',
      parameters: [],
    });
    // workScopes ONLY (no `scopes`) — this is what makes it org-mode gated.
    const config = makeConfig({
      toolName: 'list-groups',
      pathPattern: '/groups',
      scopes: undefined,
      workScopes: ['Group.Read.All'],
    });
    return { endpoint, config };
  }

  it('is skipped when orgMode is false (default)', async () => {
    const { endpoint, config } = groupToolEndpointAndConfig();
    mockEndpoints.push(endpoint);
    mockEndpointsJson = [config];

    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    // args: (server, graphClient, readOnly=false, enabledToolsPattern=undefined, orgMode=false)
    registerGraphTools(server as any, createMockGraphClient() as any, false, undefined, false);

    expect(server.tools.has('list-groups')).toBe(false);
  });

  it('is registered when orgMode is true', async () => {
    const { endpoint, config } = groupToolEndpointAndConfig();
    mockEndpoints.push(endpoint);
    mockEndpointsJson = [config];

    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    registerGraphTools(server as any, createMockGraphClient() as any, false, undefined, true);

    expect(server.tools.has('list-groups')).toBe(true);
  });
});
```

- [ ] **Step 2: Run the new block**

Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "org-mode gating"`
Expected: PASS (2 tests). The gating logic already exists, so these characterize and guard it.

- [ ] **Step 3: Run the full test suite**

Run: `npm test`
Expected: all tests pass (no regressions).

- [ ] **Step 4: Commit**

```bash
npm run format
git add src/__tests__/graph-tools.test.ts
git commit -m "test: guard org-mode gating for workScopes-only tools

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 4: Documentation and enabling org mode for the local server

**Files:**

- Modify: `README.md` (add a group-calendar subsection)
- Modify: `start-mcp.sh` (pass `--org-mode` so the group tools are actually available at runtime)

**Interfaces:**

- Consumes: nothing.
- Produces: user-facing docs + a runtime that surfaces the new tools.

- [ ] **Step 1: Add a README subsection**

Find the calendar-related section in `README.md` (search for "calendar" or the tools table). Add this subsection near it:

```markdown
### Group calendars (org mode)

Microsoft 365 **group** calendars (shared team calendars) are separate from your
personal calendars and require **org mode**. Tools:

- `list-my-groups`, `list-groups`, `get-group` — discover groups and their IDs.
  Find a group by mail with `$filter=mail eq 'team@contoso.com'`.
- `list-group-calendar-events`, `get-group-calendar-view`, `get-group-calendar-event`
- `create-group-calendar-event`, `update-group-calendar-event`, `delete-group-calendar-event`

**Setup:**

1. Grant the app the delegated scopes `Group.Read.All` and `Group.ReadWrite.All`
   (admin consent required) on your Entra ID app registration.
2. Run the server in org mode (`--org-mode`, or set `MS365_MCP_ORG_MODE=true`).
3. Re-run `--login` so the new scopes are consented into your token.

A `403` on write means the scopes were not admin-consented or the server is not
in org mode.
```

- [ ] **Step 2: Enable org mode in the launch wrapper**

In `start-mcp.sh`, change the final `exec` line from:

```bash
exec "$SERVER_DIR/node_modules/.bin/tsx" "$SERVER_DIR/src/index.ts" "$@" 2>> "$LOG"
```

to:

```bash
exec "$SERVER_DIR/node_modules/.bin/tsx" "$SERVER_DIR/src/index.ts" --org-mode "$@" 2>> "$LOG"
```

(This surfaces all org-mode/work-account tools — Teams, SharePoint, and now group calendars. Intended, since group calendar is a work-account feature.)

- [ ] **Step 3: Verify formatting**

Run: `npm run format:check && npm run lint`
Expected: pass. If `format:check` fails, run `npm run format` and re-check.

- [ ] **Step 4: Commit**

```bash
git add README.md start-mcp.sh
git commit -m "docs: document group calendar tools; enable org mode in launcher

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 5: Full verification and manual smoke test

**Files:** none (verification only).

**Interfaces:**

- Consumes: everything above.
- Produces: a green `npm run verify` and a manually-confirmed group-calendar write.

- [ ] **Step 1: Run the full verify pipeline**

Run: `npm run verify`
Expected: `generate` → `lint` → `format:check` → `build` → `test` all pass.
(Note: `verify` runs `npm run generate`, which re-downloads/uses the full spec. This must succeed end-to-end.)

- [ ] **Step 2: Re-authenticate with the new scopes**

After confirming a Global Admin has admin-consented `Group.Read.All` + `Group.ReadWrite.All` on the app registration, run the server login in org mode. For the local wrapper this is now automatic; for a manual run:

```bash
npm run dev -- --org-mode --login
```

Then complete device-code login. Expected log line: `Granted scopes:` includes `Group.Read.All` and `Group.ReadWrite.All`.

- [ ] **Step 3: Manual smoke test (via MCP client / inspector)**

1. `list-groups` with `$filter=mail eq 'Aleaquo@scalesology.com'` and `$select=id,displayName,mail` → capture the returned `id`.
2. `create-group-calendar-event` with that `group-id` and a test event (subject "MCP smoke test", a start/end an hour apart) → expect a `201` with an event `id`.
3. `delete-group-calendar-event` with the same `group-id` and the new `event-id` → cleanup.

Expected: create returns an event owned by the group; delete succeeds. A `403` indicates missing admin consent or org mode (see README).

- [ ] **Step 4: Finish the branch**

Use the `superpowers:finishing-a-development-branch` skill to decide merge/PR. Do not merge to `main` without the user's go-ahead.

---

## Self-Review

**Spec coverage:**

- Discovery (find Aleaquo id) → Task 2 (`list-my-groups`, `list-groups`, `get-group`). ✓
- Group calendar read (list/view/get) → Task 2. ✓
- Group calendar write (create/update/delete) → Task 2. ✓
- Scopes `Group.Read.All` + `Group.ReadWrite.All`, org-mode gated, auto-derived → Task 2 (config test) + Task 3 (gating test). ✓
- Stale-spec / regeneration caveat → Task 1 (forced re-download) + Task 1b (augmentation fallback). ✓
- Re-auth / admin-consent flow → already performed via `az` (done); documented in Task 4 + Task 5 Step 2. ✓
- README docs → Task 4. ✓
- Testing (register-only-in-org-mode; scope derivation) → Tasks 2 & 3. ✓
- Out-of-scope tiers (collaboration/management) → intentionally not in any task, per spec Non-Goals. ✓

**Placeholder scan:** No TBD/TODO; all code blocks and commands are concrete. Task 1b is conditional but fully specified (complete module + wiring + tests).

**Type/name consistency:** Tool names, path patterns, methods, and workScopes are identical across the spec, Task 2 endpoints block, and Task 2/Task 3 tests. `augmentGroupCalendarPaths` is defined in Task 1b and consumed only there and in its test. `registerGraphTools` 5th-arg `orgMode` matches `graph-tools.ts`.
