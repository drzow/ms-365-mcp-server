import { beforeEach, describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'fs';
import { join } from 'path';
import { registerGraphTools } from '../src/graph-tools.js';
import { api } from '../src/generated/client.js';
import type { GraphClient } from '../src/graph-client.js';

/**
 * Edge-case coverage for create-onenote-section / create-onenote-section-in-group
 * that requires the REAL generated client (microsoft_graph_onenoteSection schema)
 * and the REAL endpoints.json. The mocked-loader tests in
 * src/__tests__/graph-tools.test.ts use synthetic schemas, so this file
 * complements them by validating that the implementer's wiring works against
 * the actual generated artifacts.
 */

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
  },
}));

interface EndpointConfig {
  toolName: string;
  pathPattern: string;
  method: string;
  scopes?: string[];
  workScopes?: string[];
}

describe('create-onenote-section: real generated client wiring', () => {
  let mockServer: {
    tool: ReturnType<typeof vi.fn>;
    handlers: Map<string, (params: Record<string, unknown>) => Promise<unknown>>;
  };
  let mockGraphClient: GraphClient;

  beforeEach(() => {
    vi.clearAllMocks();
    const handlers = new Map<string, (params: Record<string, unknown>) => Promise<unknown>>();
    mockServer = {
      tool: vi.fn(
        (
          name: string,
          _description: string,
          _schema: unknown,
          _annotations: unknown,
          handler: (params: Record<string, unknown>) => Promise<unknown>
        ) => {
          handlers.set(name, handler);
        }
      ),
      handlers,
    };
    mockGraphClient = {
      graphRequest: vi.fn().mockResolvedValue({
        content: [{ type: 'text', text: JSON.stringify({ id: 'sec-real', displayName: 'X' }) }],
      }),
    } as unknown as GraphClient;
  });

  it('should register create-onenote-section and create-onenote-section-in-group from the real client', () => {
    registerGraphTools(mockServer, mockGraphClient, /* readOnly */ false);

    expect(mockServer.handlers.has('create-onenote-section')).toBe(true);
    expect(mockServer.handlers.has('create-onenote-section-in-group')).toBe(true);
  });

  it('should call POST with the real generated microsoft_graph_onenoteSection schema accepting { displayName: X }', async () => {
    // Confirm the tool's body parameter parses { displayName: 'X' } against the
    // ACTUAL generated schema (microsoft_graph_onenoteSection has all-optional
    // fields, so a bare displayName satisfies it).
    const realEndpoint = api.endpoints.find((e) => e.alias === 'create-onenote-section');
    expect(realEndpoint).toBeDefined();
    const bodyParam = realEndpoint!.parameters!.find((p) => p.type === 'Body');
    expect(bodyParam).toBeDefined();
    expect(bodyParam!.schema).toBeDefined();
    const result = bodyParam!.schema!.safeParse({ displayName: 'Q2 Planning' });
    expect(result.success).toBe(true);

    // Now register and exercise end-to-end.
    registerGraphTools(mockServer, mockGraphClient, /* readOnly */ false);

    const handler = mockServer.handlers.get('create-onenote-section')!;
    await handler({ notebookId: 'nb-real', body: { displayName: 'Q2 Planning' } });

    expect(mockGraphClient.graphRequest).toHaveBeenCalledTimes(1);
    const [requestedPath, options] = (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock
      .calls[0];
    expect(requestedPath).toContain('/me/onenote/notebooks/nb-real/sections');
    expect(requestedPath).not.toContain(':notebookId');
    expect(options.method).toBe('POST');
    // Default JSON content type (the bodyFormat=text Prefer header is added by
    // the executor for outlook bodies; here we just verify Content-Type is not
    // overridden to text/html).
    expect(options.headers['Content-Type']).not.toBe('text/html');
    // The displayName should round-trip through serialization.
    expect(options.body).toContain('"displayName"');
    expect(options.body).toContain('"Q2 Planning"');
  });

  it('should preserve realistic Graph-style notebook IDs (with !) end-to-end via the real client', async () => {
    registerGraphTools(mockServer, mockGraphClient, false);

    const handler = mockServer.handlers.get('create-onenote-section')!;
    const realisticId = '0-A1B2C3D4E5F60718-2!1-A1B2C3D4E5F60718!7777';
    await handler({ notebookId: realisticId, body: { displayName: 'Real ID Test' } });

    const [requestedPath] = (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock
      .calls[0];
    expect(requestedPath).toContain(`/me/onenote/notebooks/${realisticId}/sections`);
    expect(requestedPath).not.toContain('%21'); // ! must not be percent-encoded
  });

  it('should preserve realistic Graph-style sectionGroup IDs end-to-end via the real client', async () => {
    registerGraphTools(mockServer, mockGraphClient, false);

    const handler = mockServer.handlers.get('create-onenote-section-in-group')!;
    const realisticId = '0-FEDCBA9876543210-1!2-FEDCBA9876543210!4242';
    await handler({ sectionGroupId: realisticId, body: { displayName: 'Nested Real' } });

    const [requestedPath] = (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock
      .calls[0];
    expect(requestedPath).toContain(`/me/onenote/sectionGroups/${realisticId}/sections`);
    expect(requestedPath).not.toContain('%21');
  });

  it('should accept kebab-case notebook-id from LLMs against the real client', async () => {
    // The endpoints.json placeholder is {notebook-id} (kebab) but the generated
    // client uses :notebookId (camel). LLMs may pass either form. The executor
    // normalizes — confirm with the real client wiring.
    registerGraphTools(mockServer, mockGraphClient, false);

    const handler = mockServer.handlers.get('create-onenote-section')!;
    await handler({ 'notebook-id': 'nb-kebab', body: { displayName: 'Kebab Path' } });

    const [requestedPath] = (mockGraphClient.graphRequest as ReturnType<typeof vi.fn>).mock
      .calls[0];
    expect(requestedPath).toContain('/me/onenote/notebooks/nb-kebab/sections');
    expect(requestedPath).not.toContain(':notebookId');
  });

  it('should be filtered out of the registered tools under read-only mode', () => {
    registerGraphTools(mockServer, mockGraphClient, /* readOnly */ true);

    expect(mockServer.handlers.has('create-onenote-section')).toBe(false);
    expect(mockServer.handlers.has('create-onenote-section-in-group')).toBe(false);
  });
});

describe('create-onenote-section: endpoints.json scope and config assertions', () => {
  // Read the real endpoints.json directly. This test does not mock fs, so it
  // exercises the file as committed — the source of truth for scope wiring.
  const endpointsPath = join(process.cwd(), 'src', 'endpoints.json');
  const endpoints = JSON.parse(readFileSync(endpointsPath, 'utf8')) as EndpointConfig[];

  it('should declare Notes.Create as the only scope on create-onenote-section', () => {
    const config = endpoints.find((e) => e.toolName === 'create-onenote-section');
    expect(config).toBeDefined();
    expect(config!.scopes).toEqual(['Notes.Create']);
    expect(config!.workScopes).toBeUndefined();
    expect(config!.method).toBe('post');
    expect(config!.pathPattern).toBe('/me/onenote/notebooks/{notebook-id}/sections');
  });

  it('should declare Notes.Create as the only scope on create-onenote-section-in-group', () => {
    const config = endpoints.find((e) => e.toolName === 'create-onenote-section-in-group');
    expect(config).toBeDefined();
    expect(config!.scopes).toEqual(['Notes.Create']);
    expect(config!.workScopes).toBeUndefined();
    expect(config!.method).toBe('post');
    expect(config!.pathPattern).toBe('/me/onenote/sectionGroups/{sectionGroup-id}/sections');
  });
});
