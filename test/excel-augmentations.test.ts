import { describe, it, expect } from 'vitest';
import { augmentExcelPaths } from '../bin/modules/excel-augmentations.mjs';

const RANGE =
  "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')";

// Note: the real Graph spec params also carry extra keys (x-ms-docs-key-type,
// description). The generator/zod-client ignores them; this minimal fixture omits
// them, but the assertions only check param `name`s so they stay robust either way.
function minimalSpec() {
  return {
    paths: {
      [RANGE]: {
        get: { operationId: 'get-x', responses: { '2XX': { description: 'ok' } } },
        parameters: [
          { name: 'drive-id', in: 'path', required: true, schema: { type: 'string' } },
          { name: 'driveItem-id', in: 'path', required: true, schema: { type: 'string' } },
          { name: 'workbookWorksheet-id', in: 'path', required: true, schema: { type: 'string' } },
          {
            name: 'address',
            in: 'path',
            required: true,
            schema: { type: 'string', nullable: true },
          },
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
    // sort requires fields, and each field requires key (Graph rejects a keyless field)
    const sortSchema = sort.post.requestBody.content['application/json'].schema;
    expect(sortSchema.required).toContain('fields');
    expect(sortSchema.properties.fields.items.required).toContain('key');
  });

  it('throws if the base range path is missing', () => {
    expect(() => augmentExcelPaths({ paths: {} })).toThrow(/base range path/);
  });
});
