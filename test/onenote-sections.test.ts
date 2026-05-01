import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

global.fetch = vi.fn();

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const MOCK_TOKEN = 'mock-access-token';

function makeHeaders() {
  return expect.objectContaining({
    Authorization: `Bearer ${MOCK_TOKEN}`,
    'Content-Type': 'application/json',
  });
}

async function graphPost(path: string, body: object) {
  const response = await fetch(`${GRAPH_BASE}${path}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${MOCK_TOKEN}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!response.ok) return null;
  return response.json();
}

describe('OneNote Section Tools', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    (global.fetch as ReturnType<typeof vi.fn>).mockImplementation(async () => ({
      ok: true,
      status: 200,
      json: async () => ({}),
      text: async () => '',
    }));
  });

  afterEach(() => {
    vi.resetAllMocks();
  });

  describe('create-onenote-section', () => {
    const NOTEBOOK_ID = 'notebook-abc';

    it('should POST to /me/onenote/notebooks/{id}/sections with displayName', async () => {
      const section = { id: 'sec-1', displayName: 'Q2 Planning' };
      (global.fetch as ReturnType<typeof vi.fn>).mockImplementationOnce(async () => ({
        ok: true,
        status: 201,
        json: async () => section,
      }));

      const result = await graphPost(`/me/onenote/notebooks/${NOTEBOOK_ID}/sections`, {
        displayName: 'Q2 Planning',
      });

      expect(result).toEqual(section);
      expect(global.fetch).toHaveBeenCalledWith(
        `${GRAPH_BASE}/me/onenote/notebooks/${NOTEBOOK_ID}/sections`,
        expect.objectContaining({ method: 'POST', headers: makeHeaders() })
      );
      const body = JSON.parse(
        (global.fetch as ReturnType<typeof vi.fn>).mock.calls[0][1].body as string
      );
      expect(body.displayName).toBe('Q2 Planning');
    });

    it('should return null when notebook does not exist', async () => {
      (global.fetch as ReturnType<typeof vi.fn>).mockImplementationOnce(async () => ({
        ok: false,
        status: 404,
        json: async () => ({ error: { message: 'The specified notebook does not exist.' } }),
      }));

      const result = await graphPost('/me/onenote/notebooks/nonexistent/sections', {
        displayName: 'Whatever',
      });
      expect(result).toBeNull();
    });

    it('should return null when scope is insufficient', async () => {
      (global.fetch as ReturnType<typeof vi.fn>).mockImplementationOnce(async () => ({
        ok: false,
        status: 403,
        json: async () => ({
          error: { message: 'Insufficient privileges to complete the operation.' },
        }),
      }));

      const result = await graphPost(`/me/onenote/notebooks/${NOTEBOOK_ID}/sections`, {
        displayName: 'Forbidden',
      });
      expect(result).toBeNull();
    });
  });

  describe('create-onenote-section-in-group', () => {
    const SECTION_GROUP_ID = 'sg-xyz';

    it('should POST to /me/onenote/sectionGroups/{id}/sections with displayName', async () => {
      const section = { id: 'sec-2', displayName: 'Subsection' };
      (global.fetch as ReturnType<typeof vi.fn>).mockImplementationOnce(async () => ({
        ok: true,
        status: 201,
        json: async () => section,
      }));

      const result = await graphPost(`/me/onenote/sectionGroups/${SECTION_GROUP_ID}/sections`, {
        displayName: 'Subsection',
      });

      expect(result).toEqual(section);
      expect(global.fetch).toHaveBeenCalledWith(
        `${GRAPH_BASE}/me/onenote/sectionGroups/${SECTION_GROUP_ID}/sections`,
        expect.objectContaining({ method: 'POST', headers: makeHeaders() })
      );
      const body = JSON.parse(
        (global.fetch as ReturnType<typeof vi.fn>).mock.calls[0][1].body as string
      );
      expect(body.displayName).toBe('Subsection');
    });

    it('should return null when section group does not exist', async () => {
      (global.fetch as ReturnType<typeof vi.fn>).mockImplementationOnce(async () => ({
        ok: false,
        status: 404,
        json: async () => ({ error: { message: 'The specified section group does not exist.' } }),
      }));

      const result = await graphPost('/me/onenote/sectionGroups/nonexistent/sections', {
        displayName: 'Whatever',
      });
      expect(result).toBeNull();
    });
  });
});
