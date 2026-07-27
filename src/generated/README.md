# MS 365 OpenAPI Client Generation

This directory contains the generated TypeScript client for the Microsoft 365 API based on the OpenAPI specification.

## The Evolution

### Version 1: AI-Generated Mappings

Our initial approach used AI to map Microsoft 365 documentation and OpenAPI specifications directly into MCP tools with
Zod mappings. While conceptually appealing, this approach didn't work well in practice and created several problems.

### Version 2: Direct OpenAPI Spec Usage

We then moved to using the full MS 365 OpenAPI specification file directly. This improved reliability but created new
significant problems:

- The spec file was a whopping 45MB in size
- It had to be included in the npm package
- Startup time was painfully slow due to parsing the large spec file

### Version 3: Current Solution (Trimmed Spec + Generated Client)

We eventually settled on a combined approach:

- Trim the OpenAPI spec to only what we need
- Generate static TypeScript client code using [openapi-zod-client](https://github.com/astahmer/openapi-zod-client)

### Benefits

- **Dramatically faster startup time** - No need to parse a large spec file
- **Significantly smaller package size** - No more bundling a 45MB spec file
- **Type safety** - Full TypeScript types generated from the OpenAPI spec
- **Validation** - Zod schemas for request/response validation

## Request Body Trimming

Microsoft's metadata describes every write endpoint's request body with the _entire_ entity,
navigation properties included. Those nav props are read-only expansions — Graph ignores them on
POST/PATCH — but each one drags its whole sub-entity graph into the generated Zod schema, and from
there into every `tools/list` payload the client pays for. `create-sharepoint-list` alone was 22k
tokens, ~19k of it fields the API cannot accept. Across the server they were 44% of the tool-list
budget.

So the generator drops properties marked `x-ms-navigationProperty: true` from **request bodies
only**. Responses keep theirs, since expanding them on read is exactly what they're for. The
filtered body is inlined at the path rather than mutating the shared component, which the
corresponding GET responses still reference.

A few nav properties genuinely are settable on write — a list's `columns`, a chat's required
`members`, a listItem's `fields`, a draft's `attachments`. Name those per-endpoint in
`src/endpoints.json` and they survive the pass:

```json
{
  "pathPattern": "/sites/{site-id}/lists",
  "method": "post",
  "toolName": "create-sharepoint-list",
  "bodyNavProps": ["columns"]
}
```

A `bodyNavProps` entry naming an unknown tool fails the build. `test/request-body-nav-props.test.ts`
asserts the pass against the checked-in trimmed spec.

### Current Limitations & Future Improvements

While this approach is a significant improvement, it's not perfect. The MCP server might still struggle to understand
the MS 365 endpoints correctly, and there's room for improvements in how the API is exposed and documented for AI
assistants. However, with the current foundation of generated TypeScript clients and proper type safety, these
improvements should now be much easier to implement and maintain.

## Regenerating the Client

To regenerate the client code (e.g., after API changes or to update the supported endpoints):

```
npm run bin/generate-graph-client.mjs
```

This command does the following:

1. Fetches/processes the OpenAPI spec
2. Generates the TypeScript client with Zod validation
3. Outputs the result to `client.ts` in this directory

No complex build scripts needed - the generation is handled by openapi-zod-client.
