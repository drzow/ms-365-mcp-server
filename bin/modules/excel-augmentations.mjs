// Injects Excel workbook write operations that Microsoft's Graph OpenAPI metadata
// omits, so the declarative endpoints.json/generator pipeline can expose them.
// Called from simplified-openapi.mjs before the spec is trimmed.

const RANGE =
  "/drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')";

// 2-D array of arbitrary cell values (string | number | boolean | null per cell).
const TWO_D_ARRAY = { type: 'array', items: { type: 'array', items: {} } };

function clone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function buildOperation({ operationId, summary, schema }) {
  return {
    tags: ['drives.driveItem'],
    summary,
    operationId,
    requestBody: {
      description: 'Augmented Excel write body',
      required: true,
      content: { 'application/json': { schema } },
    },
    responses: {
      // The '2XX' key is intentional and matches the existing get-excel-range
      // convention in the trimmed spec — openapi-zod-client maps it to the 200
      // path. Do NOT "fix" it to '200'.
      '2XX': {
        description: 'Success',
        content: { 'application/json': { schema: { type: 'object' } } },
      },
      '4XX': { $ref: '#/components/responses/error' },
      '5XX': { $ref: '#/components/responses/error' },
    },
    'x-ms-docs-operation-type': 'operation',
  };
}

export function augmentExcelPaths(openApiSpec) {
  const paths = openApiSpec.paths || {};
  const rangePath = paths[RANGE];
  if (!rangePath || !rangePath.parameters) {
    throw new Error(`augmentExcelPaths: base range path not found in spec: ${RANGE}`);
  }
  // Path-level params for the range: drive-id, driveItem-id, workbookWorksheet-id, address.
  const params = clone(rangePath.parameters);

  // 1. PATCH range(address) — set values / formulas / number formats. Merge into the
  //    existing path item (which currently only has GET) so GET is preserved.
  rangePath.patch = buildOperation({
    operationId: 'set-excel-range',
    summary: 'Set values, formulas, or number formats on a range',
    schema: {
      type: 'object',
      properties: { values: TWO_D_ARRAY, formulas: TWO_D_ARRAY, numberFormat: TWO_D_ARRAY },
    },
  });

  // 2. PATCH range(address)/format/font
  paths[`${RANGE}/format/font`] = {
    parameters: clone(params),
    patch: buildOperation({
      operationId: 'format-excel-range-font',
      summary: 'Set font properties on a range',
      schema: {
        type: 'object',
        properties: {
          name: { type: 'string' },
          size: { type: 'number' },
          color: { type: 'string' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          underline: { type: 'string' },
        },
      },
    }),
  };

  // 3. PATCH range(address)/format/fill
  paths[`${RANGE}/format/fill`] = {
    parameters: clone(params),
    patch: buildOperation({
      operationId: 'format-excel-range-fill',
      summary: 'Set fill (background) color on a range',
      schema: { type: 'object', properties: { color: { type: 'string' } } },
    }),
  };

  // 4. POST range(address)/sort/apply
  paths[`${RANGE}/sort/apply`] = {
    parameters: clone(params),
    post: buildOperation({
      operationId: 'sort-excel-range',
      summary: 'Sort a range by one or more columns',
      schema: {
        type: 'object',
        required: ['fields'],
        properties: {
          fields: {
            type: 'array',
            items: {
              type: 'object',
              // key (0-based column index within the range) is required by the Graph
              // range/sort/apply action — a field without it is rejected server-side.
              required: ['key'],
              properties: {
                key: { type: 'integer' },
                ascending: { type: 'boolean' },
                sortOn: { type: 'string' },
              },
            },
          },
          matchCase: { type: 'boolean' },
          hasHeaders: { type: 'boolean' },
          orientation: { type: 'string' },
        },
      },
    }),
  };

  return openApiSpec;
}
