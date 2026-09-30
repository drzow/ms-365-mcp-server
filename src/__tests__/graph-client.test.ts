import { describe, it, expect, vi } from 'vitest';

/**
 * Communication #10: the fetchAllPages loop lives in graph-tools.ts, but it reads the
 * body that GraphClient already serialised. These tests sit on that seam: they pin
 * that the collection annotations a caller needs to detect truncation survive
 * formatJsonResponse(), while item-level @odata.* noise is still stripped.
 */

vi.mock('../logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

import GraphClient from '../graph-client.js';

// formatJsonResponse() touches no auth state, so skip the constructor's credential setup.
function clientWithoutAuth(): any {
  const client = Object.create(GraphClient.prototype) as any;
  client.outputFormat = 'json';
  return client;
}

function page(items: number, nextToken?: string, count?: number) {
  const body: Record<string, unknown> = {
    value: Array.from({ length: items }, (_, i) => ({
      id: `folder-${i}`,
      '@odata.type': '#microsoft.graph.mailFolder',
    })),
  };
  if (nextToken !== undefined) {
    body['@odata.nextLink'] =
      `https://graph.microsoft.com/v1.0/me/mailFolders?$skiptoken=${nextToken}`;
  }
  if (count !== undefined) {
    body['@odata.count'] = count;
  }
  return body;
}

function textOf(response: any): Record<string, unknown> {
  return JSON.parse(response.content[0].text);
}

describe('GraphClient.formatJsonResponse OData annotations', () => {
  it('keeps top-level @odata.nextLink so a caller can detect truncation', () => {
    const parsed = textOf(clientWithoutAuth().formatJsonResponse(page(10, 'ABC')));

    expect(parsed['@odata.nextLink']).toBe(
      'https://graph.microsoft.com/v1.0/me/mailFolders?$skiptoken=ABC'
    );
    expect(parsed.value).toHaveLength(10);
  });

  it('keeps top-level @odata.count', () => {
    const parsed = textOf(clientWithoutAuth().formatJsonResponse(page(10, 'ABC', 26)));

    expect(parsed['@odata.count']).toBe(26);
  });

  it('keeps top-level @odata.nextLink on the includeHeaders (_headers) response shape', () => {
    const parsed = textOf(
      clientWithoutAuth().formatJsonResponse(
        { data: page(10, 'ABC'), _headers: { 'content-type': 'application/json' } },
        false,
        false
      )
    );

    expect(parsed['@odata.nextLink']).toBe(
      'https://graph.microsoft.com/v1.0/me/mailFolders?$skiptoken=ABC'
    );
  });

  it('still strips item-level @odata.* and non-pagination top-level annotations', () => {
    const body = page(2, 'ABC');
    (body as any)['@odata.context'] = 'https://graph.microsoft.com/v1.0/$metadata#mailFolders';

    const parsed = textOf(clientWithoutAuth().formatJsonResponse(body));

    expect(parsed).not.toHaveProperty('@odata.context');
    const items = parsed.value as Record<string, unknown>[];
    expect(items[0]).not.toHaveProperty('@odata.type');
    expect(items[0]).toHaveProperty('id', 'folder-0');
  });
});
