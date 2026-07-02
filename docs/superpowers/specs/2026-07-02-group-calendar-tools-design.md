# Group Calendar Tools — Design Spec

**Date:** 2026-07-02
**Status:** DRAFT — pending user review
**Author:** Terry Brugger (with Claude Code)

## Problem

There is currently no way to write to (or discover) a Microsoft 365 **group**
calendar through this MCP server. Concretely, we want to create an event on the
**Aleaquo** group calendar (`Aleaquo@scalesology.com`) so the event is owned by
the group mailbox and survives any individual account.

Today the server exposes only personal-calendar tools:

- `list-calendars` returns only the six calendars owned by the signed-in user;
  group calendars never appear there.
- `create-specific-calendar-event` targets `/me/calendars/{calendar-id}/events`
  — a **personal** path. There is no group-calendar equivalent.
- There is no directory/group discovery tool, so the group's `id` (a GUID)
  cannot be found from inside the server.

## Goals

1. Discover Microsoft 365 groups the user can see (find Aleaquo's `id` by mail).
2. Read and write the group's calendar: list, view (expanded recurring), get,
   create, update, delete events on `/groups/{group-id}/calendar/...`.
3. Do this within the server's existing patterns (endpoints.json → generated
   client), gated behind org mode, with scopes derived automatically.

## Non-Goals (documented future phases)

Deliberately out of scope for this change to keep the consent footprint small:

- **Collaboration tier** — group SharePoint files (`/groups/{id}/drive/...`) and
  group Planner plans/tasks.
- **Management tier** — membership admin (list/add/remove members & owners),
  group settings, create/delete groups. These pull in higher-privilege scopes
  (`GroupMember.ReadWrite.All`, `Directory.ReadWrite.All`) and are better as a
  separate, explicitly-requested phase.

If desired later, each is an additive set of endpoints.json entries following
the same pattern established here.

## How the server works (context)

- **`src/endpoints.json`** (git-tracked) is the source of truth: each entry is a
  Graph `pathPattern` + `method` + `toolName` + `scopes`/`workScopes` +
  optional `llmTip` and flags (`supportsTimezone`, etc.).
- **`npm run generate`** downloads the upstream Graph OpenAPI, trims it to the
  paths named in endpoints.json (`openapi/openapi-trimmed.yaml`), and generates
  the zod client (`src/generated/client.ts`). The openapi files and
  `client.ts` are **git-ignored** — only `endpoints.json` is committed. CI
  regenerates on every build.
- **Scopes are derived automatically** from the enabled endpoints
  (`buildScopesFromEndpoints` in `src/auth.ts`). Adding a `workScopes` entry is
  what causes that scope to be requested at login.
- **Org mode** (`--org-mode`, `--work-mode`, or `MS365_MCP_ORG_MODE=true`) gates
  every tool that has only `workScopes`. Group tools are work-account features,
  so they use `workScopes` and only appear/consent in org mode. Existing group
  tools (e.g. `list-group-conversations`) already follow this convention.

### Known repo-state caveat (must handle during implementation)

The committed `openapi/openapi.yaml` is a **stale partial** (1302 paths; missing
125 of the 158 paths currently referenced by endpoints.json — including
`/me/calendars/{calendar-id}/events`). Because `createAndSaveSimplifiedOpenAPI`
throws on any endpoints.json path absent from the spec, **`npm run generate`
does not currently succeed against the local file** — it requires a fresh
`--force` download of the full upstream spec (which is what CI does). This is
pre-existing and not caused by this change, but the implementation must
re-download the full spec to regenerate.

## Scopes and re-authentication

**New delegated scopes:** `Group.Read.All` (read) and `Group.ReadWrite.All`
(create/update/delete). Both are **admin-consent** scopes in Entra ID.

Because the server uses a **custom app registration** (`MS365_MCP_CLIENT_ID` in
`.env`), the cleanest path is:

1. A Global Admin opens that app registration in Entra ID → **API permissions**
   → adds **Microsoft Graph → Delegated → `Group.Read.All` and
   `Group.ReadWrite.All`** → **Grant admin consent**.
2. Run the server in **org mode** (add `--org-mode` to `start-mcp.sh`, or set
   `MS365_MCP_ORG_MODE=true` in `.env`).
3. Re-run `--login` (device code). The new scopes are now requestable; consent
   is already granted, so login succeeds silently on scope.

A `403` on create means step 1 or 3 was not completed; the create tool's
`llmTip` will say so explicitly.

## Tools to add (Calendar-focused tier)

All entries use `workScopes` (org-mode gated). `supportsTimezone: true` on the
event tools mirrors the existing personal-calendar tools.

| Tool name | Method | Path | workScopes |
|---|---|---|---|
| `list-my-groups` | GET | `/me/memberOf` | `Group.Read.All` |
| `list-groups` | GET | `/groups` | `Group.Read.All` |
| `get-group` | GET | `/groups/{group-id}` | `Group.Read.All` |
| `list-group-calendar-events` | GET | `/groups/{group-id}/calendar/events` | `Group.Read.All` |
| `get-group-calendar-view` | GET | `/groups/{group-id}/calendar/calendarView` | `Group.Read.All` |
| `get-group-calendar-event` | GET | `/groups/{group-id}/calendar/events/{event-id}` | `Group.Read.All` |
| `create-group-calendar-event` | POST | `/groups/{group-id}/calendar/events` | `Group.ReadWrite.All` |
| `update-group-calendar-event` | PATCH | `/groups/{group-id}/calendar/events/{event-id}` | `Group.ReadWrite.All` |
| `delete-group-calendar-event` | DELETE | `/groups/{group-id}/calendar/events/{event-id}` | `Group.ReadWrite.All` |

Derived scope set: **`Group.Read.All`, `Group.ReadWrite.All`**.

### LLM tips (key ones)

- `list-groups`: "Find a group's id by mail:
  `$filter=mail eq 'Aleaquo@scalesology.com'`, or by name with
  `$filter=startswith(displayName,'Alea')`. Always
  `$select=id,displayName,mail,groupTypes`. Unified (Microsoft 365) groups have
  `groupTypes` containing `Unified` and are the ones with a group calendar."
- `list-my-groups`: "Returns the groups AND directory roles you belong to; keep
  only objects where `@odata.type` is `#microsoft.graph.group`. Use
  `$select=id,displayName,mail` to reduce noise."
- `list-group-calendar-events`: "Does NOT expand recurring events — returns
  seriesMaster only. Use `get-group-calendar-view` with a date range for
  expanded instances."
- `get-group-calendar-view`: "Requires `startDateTime` and `endDateTime` query
  params in ISO 8601 (e.g. `2026-07-01T00:00:00Z`). Returns expanded recurring
  instances."
- `create-group-calendar-event`: "Creates an event owned by the group mailbox
  (survives individual accounts). Get `group-id` from `list-groups`. Requires
  org mode and `Group.ReadWrite.All` admin consent — a `403` means that consent
  or org mode is missing."

## Architecture / approach

Add the nine entries to `src/endpoints.json`; everything else flows from the
existing pipeline. No changes to `auth.ts`, `graph-tools.ts`, or the generator
are expected **unless** the upstream spec is missing a group path.

### De-risking spike (first implementation step)

Because we cannot regenerate against the stale local spec, step one is:

1. `npm run generate --force` to re-download the full upstream Graph v1.0 spec.
2. Grep the freshly downloaded `openapi/openapi.yaml` for each of the nine
   paths. The group calendar paths are expected to be present in Microsoft's
   published v1.0 metadata.
3. **Contingency:** if any needed path is genuinely absent upstream, add a
   `bin/modules/group-augmentations.mjs` that injects the missing path items
   (mirroring the existing `excel-augmentations.mjs` pattern) and call it from
   `simplified-openapi.mjs` alongside `augmentExcelPaths`. This keeps the
   feature self-contained and independent of upstream drift.

### Generated body schema

`create/update-group-calendar-event` get their request-body zod schema from the
upstream `event` schema on the group path — the same mechanism that gives
`create-specific-calendar-event` a rich event body (subject, start, end,
attendees, body, location, recurrence). No hand-written schema needed.

## Error handling

Unchanged from existing tools: Graph errors surface as `isError: true` with the
Graph message. The `403` guidance lives in the create tool's `llmTip`.

## Testing

- Regenerate + `npm run build` must succeed with the nine new paths.
- Extend `src/__tests__/graph-tools.test.ts` (or add a focused test) to assert:
  - the nine tools register **only** when `orgMode = true`, and are skipped when
    `orgMode = false` (they have `workScopes` only);
  - `buildScopesFromEndpoints(true)` includes `Group.Read.All` and
    `Group.ReadWrite.All`, and `buildScopesFromEndpoints(false)` does **not**.
- Manual verification (post admin-consent + org-mode login): `list-groups`
  filtered by `Aleaquo@scalesology.com` returns an id; `create-group-calendar-event`
  creates a test event on that calendar; delete it to clean up.

## Documentation

Update `README.md`: a short "Group calendars (org mode)" section covering the
admin-consent scopes, enabling org mode, and the discover-then-write flow.

## Rollout / runtime checklist (for the user)

1. Global Admin: add `Group.Read.All` + `Group.ReadWrite.All` (delegated) to the
   app registration and grant admin consent.
2. Enable org mode (`--org-mode` in `start-mcp.sh` or `MS365_MCP_ORG_MODE=true`).
3. Re-run `--login`.
4. `list-groups` (filter by mail) → copy Aleaquo's `id` →
   `create-group-calendar-event`.
