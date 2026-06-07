# AI PPT Builder Backend

Fastify API service for the AI PPT Builder plugin.

## Current Phase

This is the phase 2 API skeleton. It accepts report/template uploads and validates the request shape, but PPTX generation is intentionally left for phase 3.

## Technology

- Fastify for the HTTP API.
- `@fastify/multipart` for report/template upload handling.
- `@fastify/cors` for local task pane integration.
- TypeScript for request/response contracts.
- Zod for environment parsing.

## Local Development

```bash
npm install
npm run dev
```

Default server:

```text
http://127.0.0.1:3000
```

Health check:

```text
GET /health
```

Reserved generation endpoint:

```text
POST /api/decks/generate
```

Multipart fields:

- `report`: source report file.
- `template`: PPTX template file.
- `instruction`: optional generation instruction.
