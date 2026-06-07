# AI PPT Builder Backend

Fastify API service for the AI PPT Builder plugin.

## Current Phase

This service now covers the phase 2 API skeleton, phase 3 PPTX smoke-test loop, and phase 5 template parsing script/API. It accepts report/template uploads, validates the request shape, generates a one-slide PPTX smoke-test deck, returns it as Base64, and can analyze PPTX templates into reusable template profiles.

## Technology

- Fastify for the HTTP API.
- `@fastify/multipart` for report/template upload handling.
- `@fastify/cors` for local task pane integration.
- JSZip and `fast-xml-parser` for PPTX Open XML template parsing.
- TypeScript for request/response contracts.
- Zod for environment parsing.

## Local Development

```bash
npm install
npm run dev
```

## AI API Configuration

The AI API key is never stored in source code and must not be committed to GitHub.

Create a local env file:

```bash
cp .env.example .env
```

Then edit `backend/.env`:

```text
AI_PROVIDER=openai
AI_API_KEY=your_api_key_here
AI_BASE_URL=https://api.openai.com/v1
AI_MODEL=gpt-4.1-mini
```

The frontend never receives `AI_API_KEY`. It calls the backend only, and the backend is responsible for future AI requests.

Public AI settings endpoint:

```text
GET /api/settings/ai
```

Response example:

```json
{
  "provider": "openai",
  "configured": true,
  "baseUrlHost": "api.openai.com",
  "model": "gpt-4.1-mini"
}
```

Default server:

```text
http://127.0.0.1:3000
```

Health check:

```text
GET /health
```

Generation endpoint:

```text
POST /api/decks/generate
```

Multipart fields:

- `report`: source report file.
- `template`: optional PPTX template file. If omitted, the backend uses the saved default template, falling back to the bundled Tsinghua template.
- `instruction`: optional generation instruction.

Response:

```json
{
  "deckId": "deck_...",
  "pptxBase64": "base64-encoded-pptx",
  "summary": "Generated a one-slide PPTX smoke test deck.",
  "qa": "Smoke test only: report parsing, template page replacement, and content QA are not implemented yet."
}
```

Template endpoints:

```text
GET  /api/templates/default
POST /api/templates/default
POST /api/templates/analyze
```

CLI template analysis:

```bash
npm run analyze:template -- "../清华大学2025年度演示文稿系列模板2-通用主题.pptx" "../templates/tsinghua-2025-general-2"
```

Output:

```text
templates/tsinghua-2025-general-2/template-profile.json
```
