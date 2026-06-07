# AI PPT Builder Backend

Fastify API service for the AI PPT Builder plugin.

## Current Phase

This service now covers the phase 2 API skeleton and phase 3 PPTX smoke-test loop. It accepts report/template uploads, validates the request shape, generates a one-slide PPTX smoke-test deck, and returns it as Base64.

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

Response:

```json
{
  "deckId": "deck_...",
  "pptxBase64": "base64-encoded-pptx",
  "summary": "Generated a one-slide PPTX smoke test deck.",
  "qa": "Smoke test only: template parsing and content QA are not implemented yet."
}
```
