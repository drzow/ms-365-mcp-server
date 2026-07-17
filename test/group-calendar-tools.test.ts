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
