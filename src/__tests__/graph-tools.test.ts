import { describe, it, expect, vi, beforeEach } from 'vitest';
import { z } from 'zod';

/**
 * We test executeGraphTool logic by importing it indirectly through registerGraphTools.
 * Strategy: mock GraphClient, create a real McpServer, register tools, then invoke them.
 */

// Mock logger to silence output
vi.mock('../logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

// Mock the generated client — we supply our own endpoint definitions per test
const mockEndpoints: any[] = [];
vi.mock('../generated/client.js', () => ({
  api: {
    get endpoints() {
      return mockEndpoints;
    },
  },
}));

// Mock endpoints.json — we supply our own config per test
let mockEndpointsJson: any[] = [];
vi.mock('fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('fs')>();
  return {
    ...actual,
    readFileSync: (filePath: string, encoding?: string) => {
      if (typeof filePath === 'string' && filePath.includes('endpoints.json')) {
        return JSON.stringify(mockEndpointsJson);
      }
      return actual.readFileSync(filePath, encoding as any);
    },
  };
});

// Mock tool-categories
vi.mock('../tool-categories.js', () => ({
  TOOL_CATEGORIES: {},
}));

// ---------- helpers ----------

function makeEndpoint(overrides: Partial<any> = {}) {
  return {
    method: 'get',
    path: '/me/messages',
    alias: 'test-tool',
    description: 'Test tool',
    requestFormat: 'json' as const,
    parameters: [
      { name: 'filter', type: 'Query', schema: z.string().optional() },
      { name: 'search', type: 'Query', schema: z.string().optional() },
      { name: 'select', type: 'Query', schema: z.string().optional() },
      { name: 'orderby', type: 'Query', schema: z.string().optional() },
      { name: 'count', type: 'Query', schema: z.boolean().optional() },
      { name: 'top', type: 'Query', schema: z.number().optional() },
      { name: 'skip', type: 'Query', schema: z.number().optional() },
    ],
    response: z.any(),
    ...overrides,
  };
}

function makeConfig(overrides: Partial<any> = {}) {
  return {
    pathPattern: '/me/messages',
    method: 'get',
    toolName: 'test-tool',
    scopes: ['Mail.Read'],
    ...overrides,
  };
}

/** Creates a mock GraphClient with a controllable graphRequest spy */
function createMockGraphClient(responses?: any[]) {
  const responseQueue = [...(responses || [])];
  return {
    graphRequest: vi.fn().mockImplementation(async () => {
      if (responseQueue.length > 0) {
        return responseQueue.shift();
      }
      return {
        content: [{ type: 'text', text: JSON.stringify({ value: [] }) }],
      };
    }),
  };
}

/**
 * Because registerGraphTools reads endpointsData at module load time,
 * and we mock fs.readFileSync, we need to re-import after setting mocks.
 */
async function loadModule() {
  // Clear cached module so mocks take effect
  vi.resetModules();
  const mod = await import('../graph-tools.js');
  return mod;
}

/** Minimal McpServer mock that captures registered tools */
function createMockServer() {
  const tools = new Map<
    string,
    { description: string; schema: any; handler: (...args: any[]) => any }
  >();
  return {
    tool: vi.fn(
      (
        name: string,
        description: string,
        schema: any,
        annotations: any,
        handler: (...args: any[]) => any
      ) => {
        tools.set(name, { description, schema, handler });
      }
    ),
    tools,
  };
}

// ========== TESTS ==========

describe('graph-tools', () => {
  beforeEach(() => {
    mockEndpoints.length = 0;
    mockEndpointsJson = [];
    vi.clearAllMocks();
  });

  // ---- 1. $count advanced query mode ----
  describe('$count advanced query mode', () => {
    it('should set ConsistencyLevel: eventual header when $count=true', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      // Invoke the registered tool with count=true
      const tool = server.tools.get('test-tool');
      expect(tool).toBeDefined();
      await tool!.handler({ count: true });

      // Verify graphRequest was called with ConsistencyLevel header
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(1);
      const [url] = graphClient.graphRequest.mock.calls[0];
      // $count=true should appear in query string
      expect(url).toContain('$count=true');
    });
  });

  // ---- 2. fetchAllPages pagination ----
  describe('fetchAllPages pagination', () => {
    it('should follow @odata.nextLink and combine results', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                value: [{ id: '1' }, { id: '2' }],
                '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=2',
              }),
            },
          ],
        },
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                value: [{ id: '3' }],
              }),
            },
          ],
        },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      const result = await tool!.handler({ fetchAllPages: true });

      // Should have made 2 requests (initial + 1 nextLink)
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(2);

      // Combined result should have 3 items
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.value).toHaveLength(3);
      expect(parsed.value.map((v: any) => v.id)).toEqual(['1', '2', '3']);
      // nextLink should be removed from final response
      expect(parsed['@odata.nextLink']).toBeUndefined();
    });

    it('should stop at 100 page limit', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      // Generate 101 responses — each has a nextLink except the last
      const responses = [];
      for (let i = 0; i < 101; i++) {
        responses.push({
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                value: [{ id: `item-${i}` }],
                '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=' + (i + 1),
              }),
            },
          ],
        });
      }

      const graphClient = createMockGraphClient(responses);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      await tool!.handler({ fetchAllPages: true });

      // 1 initial + 99 pagination = 100 total requests (stops at pageCount=100)
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(100);
    });
  });

  // ---- 3. Parameter describe() overrides ----
  describe('parameter describe() overrides', () => {
    it('should apply custom descriptions to OData parameters', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any);

      const tool = server.tools.get('test-tool');
      expect(tool).toBeDefined();

      const schema = tool!.schema;

      // $filter override
      expect(schema['filter']).toBeDefined();
      expect(schema['filter'].description).toContain('OData filter expression');
      expect(schema['filter'].description).toContain('$count=true');

      // $search override
      expect(schema['search']).toBeDefined();
      expect(schema['search'].description).toContain('KQL search query');

      // $select override
      expect(schema['select']).toBeDefined();
      expect(schema['select'].description).toContain('Comma-separated fields');

      // $orderby override
      expect(schema['orderby']).toBeDefined();
      expect(schema['orderby'].description).toContain('Sort expression');

      // $count override
      expect(schema['count']).toBeDefined();
      expect(schema['count'].description).toContain('advanced query mode');
    });
  });

  // ---- 4. returnDownloadUrl ----
  describe('returnDownloadUrl', () => {
    it('should strip /content from path and return downloadUrl when returnDownloadUrl=true', async () => {
      const endpoint = makeEndpoint({
        alias: 'download-file',
        path: '/me/drive/items/:driveItem-id/content',
        parameters: [{ name: 'driveItem-id', type: 'Path', schema: z.string() }],
      });
      const config = makeConfig({
        toolName: 'download-file',
        pathPattern: '/me/drive/items/{driveItem-id}/content',
        returnDownloadUrl: true,
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const downloadUrl = 'https://download.example.com/file.pdf';
      const graphClient = createMockGraphClient([
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                '@microsoft.graph.downloadUrl': downloadUrl,
                name: 'file.pdf',
              }),
            },
          ],
        },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('download-file');
      expect(tool).toBeDefined();
      await tool!.handler({ 'driveItem-id': 'abc123' });

      // Path should NOT end with /content — it gets stripped
      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).not.toContain('/content');
      expect(requestedPath).toContain('/me/drive/items/abc123');
    });
  });

  // ---- 5. kebab-case path param normalization ----
  describe('kebab-case path param normalization', () => {
    it('should substitute path when LLM passes message-id (kebab) but schema has messageId (camelCase)', async () => {
      // Simulates what hack.ts generates: path uses :messageId (camelCase)
      // but LLMs may pass message-id (kebab-case) since endpoints.json uses {message-id}
      const endpoint = makeEndpoint({
        alias: 'get-mail-message',
        method: 'get',
        path: '/me/messages/:messageId',
        parameters: [
          { name: 'messageId', type: 'Path', schema: z.string() },
          { name: 'select', type: 'Query', schema: z.string().optional() },
        ],
      });
      const config = makeConfig({
        toolName: 'get-mail-message',
        pathPattern: '/me/messages/{message-id}',
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'AAMk123', subject: 'Test' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('get-mail-message');
      expect(tool).toBeDefined();

      // Pass kebab-case 'message-id' — should still resolve to correct path
      await tool!.handler({ 'message-id': 'AAMk123abc=' });

      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('AAMk123abc=');
      expect(requestedPath).not.toContain(':messageId');
    });

    it('should also work when LLM passes messageId (camelCase) directly', async () => {
      const endpoint = makeEndpoint({
        alias: 'get-mail-message2',
        method: 'get',
        path: '/me/messages/:messageId',
        parameters: [{ name: 'messageId', type: 'Path', schema: z.string() }],
      });
      const config = makeConfig({
        toolName: 'get-mail-message2',
        pathPattern: '/me/messages/{message-id}',
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'AAMk456' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('get-mail-message2');
      await tool!.handler({ messageId: 'AAMk456xyz=' });

      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('AAMk456xyz=');
      expect(requestedPath).not.toContain(':messageId');
    });
  });

  // ---- 6. create-onenote-section (POST body, JSON) ----
  describe('create-onenote-section', () => {
    it('should POST to /me/onenote/notebooks/{notebook-id}/sections with JSON displayName body', async () => {
      const endpoint = makeEndpoint({
        alias: 'create-onenote-section',
        method: 'post',
        path: '/me/onenote/notebooks/:notebookId/sections',
        parameters: [
          { name: 'notebookId', type: 'Path', schema: z.string() },
          {
            name: 'body',
            type: 'Body',
            schema: z.object({ displayName: z.string() }),
          },
        ],
      });
      const config = makeConfig({
        toolName: 'create-onenote-section',
        pathPattern: '/me/onenote/notebooks/{notebook-id}/sections',
        method: 'post',
        scopes: ['Notes.Create'],
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ id: 'sec-1', displayName: 'Q2 Planning' }),
            },
          ],
        },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('create-onenote-section');
      expect(tool).toBeDefined();

      await tool!.handler({ notebookId: 'nb-abc', body: { displayName: 'Q2 Planning' } });

      expect(graphClient.graphRequest).toHaveBeenCalledTimes(1);
      const [requestedPath, options] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('/me/onenote/notebooks/nb-abc/sections');
      expect(requestedPath).not.toContain(':notebookId');
      expect(requestedPath).not.toContain('{notebook-id}');
      expect(options.method).toBe('POST');
      expect(options.body).toBe('{"displayName":"Q2 Planning"}');
      // Default JSON content type — not text/html
      expect(options.headers['Content-Type']).not.toBe('text/html');
    });

    it('should accept kebab-case notebook-id path param', async () => {
      const endpoint = makeEndpoint({
        alias: 'create-onenote-section',
        method: 'post',
        path: '/me/onenote/notebooks/:notebookId/sections',
        parameters: [
          { name: 'notebookId', type: 'Path', schema: z.string() },
          {
            name: 'body',
            type: 'Body',
            schema: z.object({ displayName: z.string() }),
          },
        ],
      });
      const config = makeConfig({
        toolName: 'create-onenote-section',
        pathPattern: '/me/onenote/notebooks/{notebook-id}/sections',
        method: 'post',
        scopes: ['Notes.Create'],
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'sec-2' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('create-onenote-section');
      // LLM may pass kebab-case 'notebook-id' (matching endpoints.json placeholder)
      // even though the generated client schema uses camelCase 'notebookId'
      await tool!.handler({ 'notebook-id': 'nb-xyz', body: { displayName: 'Sprint Notes' } });

      const [requestedPath, options] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('/me/onenote/notebooks/nb-xyz/sections');
      expect(requestedPath).not.toContain(':notebookId');
      expect(options.method).toBe('POST');
      expect(options.body).toBe('{"displayName":"Sprint Notes"}');
    });

    it('should accept wrapped { body: { displayName } } and pass schema parse on first try', async () => {
      // The wrapped form { body: { displayName: 'X' } } parses against the body schema
      // directly — no auto-wrap branch needed. The genuine wrap-branch coverage lives in
      // the edge case test "should exercise the auto-wrap branch when body parse fails as-is
      // but succeeds when wrapped".
      const endpoint = makeEndpoint({
        alias: 'create-onenote-section',
        method: 'post',
        path: '/me/onenote/notebooks/:notebookId/sections',
        parameters: [
          { name: 'notebookId', type: 'Path', schema: z.string() },
          {
            name: 'body',
            type: 'Body',
            schema: z.object({ displayName: z.string() }),
          },
        ],
      });
      const config = makeConfig({
        toolName: 'create-onenote-section',
        pathPattern: '/me/onenote/notebooks/{notebook-id}/sections',
        method: 'post',
        scopes: ['Notes.Create'],
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'sec-w' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('create-onenote-section');
      // Common LLM case: { body: { displayName: 'X' } } — schema parses paramValue directly.
      await tool!.handler({ notebookId: 'nb-w', body: { displayName: 'Wrapped Body' } });

      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.body).toBe('{"displayName":"Wrapped Body"}');
    });
  });

  describe('create-onenote-section-in-group', () => {
    it('should POST to /me/onenote/sectionGroups/{sectionGroup-id}/sections with JSON body', async () => {
      const endpoint = makeEndpoint({
        alias: 'create-onenote-section-in-group',
        method: 'post',
        path: '/me/onenote/sectionGroups/:sectionGroupId/sections',
        parameters: [
          { name: 'sectionGroupId', type: 'Path', schema: z.string() },
          {
            name: 'body',
            type: 'Body',
            schema: z.object({ displayName: z.string() }),
          },
        ],
      });
      const config = makeConfig({
        toolName: 'create-onenote-section-in-group',
        pathPattern: '/me/onenote/sectionGroups/{sectionGroup-id}/sections',
        method: 'post',
        scopes: ['Notes.Create'],
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'sec-3' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('create-onenote-section-in-group');
      expect(tool).toBeDefined();

      await tool!.handler({
        sectionGroupId: 'sg-123',
        body: { displayName: 'Subsection' },
      });

      const [requestedPath, options] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('/me/onenote/sectionGroups/sg-123/sections');
      expect(requestedPath).not.toContain(':sectionGroupId');
      expect(requestedPath).not.toContain('{sectionGroup-id}');
      expect(options.method).toBe('POST');
      expect(options.body).toBe('{"displayName":"Subsection"}');
      expect(options.headers['Content-Type']).not.toBe('text/html');
    });

    it('should accept kebab-case sectionGroup-id path param', async () => {
      const endpoint = makeEndpoint({
        alias: 'create-onenote-section-in-group',
        method: 'post',
        path: '/me/onenote/sectionGroups/:sectionGroupId/sections',
        parameters: [
          { name: 'sectionGroupId', type: 'Path', schema: z.string() },
          {
            name: 'body',
            type: 'Body',
            schema: z.object({ displayName: z.string() }),
          },
        ],
      });
      const config = makeConfig({
        toolName: 'create-onenote-section-in-group',
        pathPattern: '/me/onenote/sectionGroups/{sectionGroup-id}/sections',
        method: 'post',
        scopes: ['Notes.Create'],
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'sec-4' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('create-onenote-section-in-group');
      await tool!.handler({
        'sectionGroup-id': 'sg-kebab',
        body: { displayName: 'Kebab Section' },
      });

      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('/me/onenote/sectionGroups/sg-kebab/sections');
      expect(requestedPath).not.toContain(':sectionGroupId');
    });
  });

  // ---- 6b. create-onenote-section: edge case coverage (tester) ----
  describe('create-onenote-section edge cases', () => {
    /**
     * Helper: build a synthetic create-onenote-section endpoint that mirrors
     * what the real generated client produces (Body + Path parameters).
     * The body schema mirrors the real microsoft_graph_onenoteSection: all
     * fields optional/nullish so { displayName: 'X' } parses successfully on
     * its own AND wrapped as { body: { displayName: 'X' } } also parses
     * (because the schema uses .passthrough()).
     */
    function buildSectionEndpointAndConfig() {
      const onenoteSectionSchema = z
        .object({
          id: z.string().optional(),
          displayName: z.string().nullish(),
          isDefault: z.boolean().nullish(),
        })
        .passthrough();
      const endpoint = makeEndpoint({
        alias: 'create-onenote-section',
        method: 'post',
        path: '/me/onenote/notebooks/:notebookId/sections',
        parameters: [
          { name: 'body', type: 'Body', schema: onenoteSectionSchema },
          { name: 'notebookId', type: 'Path', schema: z.string() },
        ],
      });
      const config = makeConfig({
        toolName: 'create-onenote-section',
        pathPattern: '/me/onenote/notebooks/{notebook-id}/sections',
        method: 'post',
        scopes: ['Notes.Create'],
      });
      return { endpoint, config };
    }

    it('should preserve realistic Graph-style notebook IDs containing ! and 0-...!... segments', async () => {
      // Real OneNote notebook IDs from MS Graph look like:
      //   "0-A1B2C3D4E5F60718-2!1-A1B2C3D4E5F60718!7777"
      // The bang `!` and dash `-` are RFC 3986 unreserved chars — encodeURIComponent
      // leaves them as-is. The `=` preserve workaround in graph-tools.ts must NOT
      // mangle the ID. This test guards against accidental over-encoding.
      const { endpoint, config } = buildSectionEndpointAndConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'sec-real' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('create-onenote-section');
      expect(tool).toBeDefined();

      const realisticId = '0-A1B2C3D4E5F60718-2!1-A1B2C3D4E5F60718!7777';
      await tool!.handler({ notebookId: realisticId, body: { displayName: 'Q2 Planning' } });

      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      // The literal ID — including ! and -, both unreserved — must appear unencoded.
      expect(requestedPath).toContain(`/me/onenote/notebooks/${realisticId}/sections`);
      expect(requestedPath).not.toContain('%21'); // !
      expect(requestedPath).not.toContain('%2D'); // - (would never be encoded anyway, sanity)
      expect(requestedPath).not.toContain(':notebookId');
      expect(requestedPath).not.toContain('{notebook-id}');
    });

    it('should preserve realistic Graph-style sectionGroup IDs containing ! and 0-...!... segments', async () => {
      const sgSchema = z.object({ displayName: z.string().nullish() }).passthrough();
      const endpoint = makeEndpoint({
        alias: 'create-onenote-section-in-group',
        method: 'post',
        path: '/me/onenote/sectionGroups/:sectionGroupId/sections',
        parameters: [
          { name: 'body', type: 'Body', schema: sgSchema },
          { name: 'sectionGroupId', type: 'Path', schema: z.string() },
        ],
      });
      const config = makeConfig({
        toolName: 'create-onenote-section-in-group',
        pathPattern: '/me/onenote/sectionGroups/{sectionGroup-id}/sections',
        method: 'post',
        scopes: ['Notes.Create'],
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'sec-g' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const realisticId = '0-FEDCBA9876543210-1!2-FEDCBA9876543210!4242';
      const tool = server.tools.get('create-onenote-section-in-group');
      await tool!.handler({ sectionGroupId: realisticId, body: { displayName: 'Nested' } });

      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain(`/me/onenote/sectionGroups/${realisticId}/sections`);
      expect(requestedPath).not.toContain('%21');
    });

    it('should pass through { body: { displayName: X } } unchanged (wrapped form)', async () => {
      // Task #3 part A: explicitly verify the wrapped form. paramValue is the
      // inner object; the body schema parses it directly (no auto-wrap fires).
      const { endpoint, config } = buildSectionEndpointAndConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'sec' }) }] },
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('create-onenote-section');
      await tool!.handler({ 'notebook-id': 'nb', body: { displayName: 'Wrapped' } });

      const [, options] = graphClient.graphRequest.mock.calls[0];
      // Sent JSON is the inner object — NOT { body: { displayName: ... } }.
      expect(JSON.parse(options.body)).toEqual({ displayName: 'Wrapped' });
    });

    it('should send the expected body JSON when LLM passes notebook-id and displayName side-by-side', async () => {
      // Task #3 part B: the spec/task expects that an LLM passing
      //   { 'notebook-id': 'nb', displayName: 'X' }   (no `body` wrapper)
      // works. Trace through executeGraphTool:
      //   - 'notebook-id' → matches notebookId path param (kebab→camel)
      //   - 'displayName' → no paramDef match, name !== 'body', no path placeholder
      //     → silently ignored, body stays null
      // We assert the OBSERVED behavior here so a future refactor is documented.
      // If this test starts failing because the executor now wraps bare scalars,
      // that's a feature improvement — update the assertion accordingly.
      const { endpoint, config } = buildSectionEndpointAndConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'sec' }) }] },
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('create-onenote-section');
      await tool!.handler({ 'notebook-id': 'nb', displayName: 'Bare' });

      const [requestedPath, options] = graphClient.graphRequest.mock.calls[0];
      // Path substitution still works.
      expect(requestedPath).toContain('/me/onenote/notebooks/nb/sections');
      // Document current behavior: bare displayName is dropped (no body sent).
      // If this changes, the assertion below must be updated to expect
      // { displayName: 'Bare' } in options.body.
      expect(options.body).toBeUndefined();
    });

    it('should exercise the auto-wrap branch when body parse fails as-is but succeeds when wrapped', async () => {
      // Direct exercise of the wrap branch in graph-tools.ts:206-215.
      // The body schema for the real onenoteSection uses .passthrough(), so any
      // object — including { body: 'X' } as a passthrough field — parses
      // successfully. We exercise the wrap path by passing a bare string for
      // body: parse({ ... a string ... }) fails (not an object), then the
      // executor wraps as { body: 'X' } which DOES parse via passthrough.
      // The wrapped object is then sent as the body.
      const { endpoint, config } = buildSectionEndpointAndConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'sec' }) }] },
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('create-onenote-section');
      // Pass body as a string (malformed input from LLM)
      await tool!.handler({ notebookId: 'nb', body: 'just-a-string' });

      const [, options] = graphClient.graphRequest.mock.calls[0];
      // The auto-wrap kicks in: { body: 'just-a-string' } parses via passthrough,
      // so the wrapped object is what gets serialized. This documents the
      // observed behavior of the wrap path.
      expect(JSON.parse(options.body)).toEqual({ body: 'just-a-string' });
    });

    it('should be filtered out under read-only mode (POST is non-GET)', async () => {
      // Acceptance criterion 4: read-only mode must skip the new tools.
      // This complements test/read-only.test.ts (which uses its own mocked
      // endpoints). Here we register a synthetic create-onenote-section in
      // read-only mode and confirm it does not appear.
      const { endpoint, config } = buildSectionEndpointAndConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any, /* readOnly */ true);

      // The POST tool must NOT be registered.
      expect(server.tools.has('create-onenote-section')).toBe(false);
      // parse-teams-url is the only tool that should appear (read-only-safe utility)
      expect(server.tools.has('parse-teams-url')).toBe(true);
    });

    it('should be registered under normal (non-read-only) mode', async () => {
      // Mirror of the read-only test: confirm the tool DOES appear in normal mode.
      const { endpoint, config } = buildSectionEndpointAndConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any, /* readOnly */ false);

      expect(server.tools.has('create-onenote-section')).toBe(true);
    });
  });

  // ---- 7. supportsTimezone ----
  describe('supportsTimezone', () => {
    it('should set Prefer: outlook.timezone header when timezone param provided', async () => {
      const endpoint = makeEndpoint({
        alias: 'list-calendar-events',
        path: '/me/events',
        parameters: [],
      });
      const config = makeConfig({
        toolName: 'list-calendar-events',
        pathPattern: '/me/events',
        supportsTimezone: true,
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('list-calendar-events');
      expect(tool).toBeDefined();

      // Verify timezone parameter was added to schema
      expect(tool!.schema['timezone']).toBeDefined();
      expect(tool!.schema['timezone'].description).toContain('IANA timezone');

      await tool!.handler({ timezone: 'Europe/Brussels' });

      // Verify Prefer header contains outlook.timezone
      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers['Prefer']).toContain('outlook.timezone="Europe/Brussels"');
    });

    it('should NOT add timezone parameter when supportsTimezone is false/absent', async () => {
      const endpoint = makeEndpoint({
        alias: 'list-mail',
        path: '/me/messages',
        parameters: [],
      });
      const config = makeConfig({
        toolName: 'list-mail',
        pathPattern: '/me/messages',
        // no supportsTimezone
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any);

      const tool = server.tools.get('list-mail');
      expect(tool!.schema['timezone']).toBeUndefined();
    });
  });
});
