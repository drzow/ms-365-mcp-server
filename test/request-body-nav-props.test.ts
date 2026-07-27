import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import yaml from 'js-yaml';

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

interface Schema {
  $ref?: string;
  type?: string;
  properties?: Record<string, Record<string, unknown>>;
  required?: string[];
}

interface Operation {
  operationId?: string;
  requestBody?: { content?: Record<string, { schema?: Schema }> };
}

const spec = yaml.load(
  readFileSync(path.join(rootDir, 'openapi/openapi-trimmed.yaml'), 'utf8')
) as {
  paths: Record<string, Record<string, Operation>>;
  components: { schemas: Record<string, Schema> };
};

const endpoints = JSON.parse(
  readFileSync(path.join(rootDir, 'src/endpoints.json'), 'utf8')
) as Array<{ toolName: string; bodyNavProps?: string[] }>;

function bodySchema(toolName: string): Schema | undefined {
  for (const pathItem of Object.values(spec.paths)) {
    for (const operation of Object.values(pathItem)) {
      if (operation?.operationId === toolName) {
        return operation.requestBody?.content?.['application/json']?.schema;
      }
    }
  }
  return undefined;
}

function bodyProps(toolName: string): string[] {
  return Object.keys(bodySchema(toolName)?.properties ?? {});
}

describe('request body navigation-property stripping', () => {
  it('drops read-only nav properties from a write body', () => {
    // microsoft.graph.list carries drive/items/subscriptions/etc. as expansions that
    // Graph ignores on POST — they were ~19k of this tool's 22k tokens.
    const props = bodyProps('create-sharepoint-list');
    expect(props).not.toContain('drive');
    expect(props).not.toContain('items');
    expect(props).not.toContain('subscriptions');
    expect(props).not.toContain('operations');
    expect(props).not.toContain('createdByUser');
  });

  it('keeps the settable scalar properties alongside them', () => {
    expect(bodyProps('create-sharepoint-list')).toEqual(
      expect.arrayContaining(['displayName', 'description', 'list'])
    );
  });

  it('preserves nav properties named in bodyNavProps', () => {
    // Each of these is genuinely settable on write; dropping them would cost real capability.
    expect(bodyProps('create-sharepoint-list')).toContain('columns');
    expect(bodyProps('create-sharepoint-list-item')).toContain('fields');
    expect(bodyProps('update-sharepoint-list-item')).toContain('fields');
    expect(bodyProps('create-chat')).toContain('members');
    expect(bodyProps('create-team-channel')).toContain('members');
    expect(bodyProps('create-draft-email')).toContain('attachments');
    expect(bodyProps('create-calendar-event')).toContain('attachments');
    expect(bodyProps('create-todo-task')).toContain('checklistItems');
    expect(bodyProps('send-channel-message')).toContain('hostedContents');
  });

  it('inlines the filtered body instead of mutating the shared component', () => {
    // GET /sites/{id}/lists/{id} must still be able to expand the nav properties.
    expect(bodySchema('create-sharepoint-list')?.$ref).toBeUndefined();
    const component = spec.components.schemas['microsoft.graph.list'];
    expect(Object.keys(component.properties ?? {})).toEqual(
      expect.arrayContaining(['drive', 'items', 'subscriptions'])
    );
  });

  it('leaves refs reachable from surviving properties in the pruned spec', () => {
    // Pruning runs after stripping; nested refs in an inlined body must survive it.
    const columns = bodySchema('create-sharepoint-list')?.properties?.columns as {
      items?: { $ref?: string };
    };
    const ref = columns?.items?.$ref?.replace('#/components/schemas/', '');
    expect(ref).toBe('microsoft.graph.columnDefinition');
    expect(spec.components.schemas[ref!]).toBeDefined();
  });

  it('declares bodyNavProps only for tools that have a JSON request body', () => {
    for (const endpoint of endpoints) {
      if (!endpoint.bodyNavProps?.length) continue;
      expect(bodySchema(endpoint.toolName), `${endpoint.toolName} has no JSON body`).toBeDefined();
    }
  });
});
