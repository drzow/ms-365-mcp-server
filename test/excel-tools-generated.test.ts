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

  it('format/sort tools resolve to addressed range paths (bug-fix regression guard)', () => {
    // The two pre-existing tools used the rangeless range() form; Task 3 repoints
    // them to addressed paths so they can target a specific range.
    for (const alias of ['format-excel-range', 'sort-excel-range']) {
      const ep = api.endpoints.find((e: any) => e.alias === alias);
      expect(ep, `missing generated tool: ${alias}`).toBeDefined();
      expect(ep!.path).toContain('range(address=');
      expect(ep!.path).not.toContain('range()');
    }
  });
});
