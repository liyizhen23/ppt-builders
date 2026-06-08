# AI PPT Builder

AI PPT Builder is a local PowerPoint Office Add-in prototype for generating and editing presentation slides with a backend API.

## Current Capabilities

- PowerPoint task pane built with React, TypeScript, Vite, and Office.js.
- Backend API built with Fastify and TypeScript.
- PPTX generation returns Base64 and inserts through `insertSlidesFromBase64`.
- Tsinghua template parsing and Template Profile generation.
- Default template persistence, with the bundled Tsinghua template as fallback.
- DOCX/text report parsing with a local Evidence Index.
- DeckPlan schema and template-slot replacement rendering.
- Local report reuse after the first upload.
- Local asset library for images and tables under `asset-library/`.
- Local editing modes:
  - selected text rewrite
  - selected image replacement
  - AI image selection from saved local assets
  - current-slide reflow into a replacement slide
  - QA check and text-structure autofix

## Project Structure

```text
backend/                 Fastify API service
frontend/                PowerPoint Office Add-in task pane
templates/               Bundled/default template profiles
asset-library/           Local ignored asset store, created at runtime
reports/current/         Local ignored current report cache
AI_PPT_PLUGIN_BUILD_PLAN.md
```

## Local Setup

Install dependencies:

```bash
cd backend
npm install

cd ../frontend
npm install
```

Configure backend AI settings by copying:

```bash
copy backend\.env.example backend\.env
```

Then edit `backend/.env`. Real API keys must only live in `.env`; `.env` files are ignored by Git.

Example for DeepSeek-compatible configuration:

```text
AI_PROVIDER=custom
AI_API_KEY=your-local-key
AI_BASE_URL=https://api.deepseek.com
AI_MODEL=deepseek-v4-pro
```

Start backend:

```bash
cd backend
npm run dev
```

Start frontend:

```bash
cd frontend
npm run dev
```

The add-in manifest points to:

```text
https://localhost:5173/index.html
```

Office Add-ins require HTTPS for local sideloading. This project uses local Office development certificates configured in `frontend/vite.config.ts`.

## Runtime Data And Secrets

These local paths are intentionally ignored and must not be pushed:

```text
backend/.env
asset-library/
reports/current/
templates/default/
```

The frontend never receives the full API key. It can only read non-secret status from:

```text
GET /api/settings/ai
```

## Useful API Endpoints

```text
GET  /health
GET  /api/settings/ai
GET  /api/templates/default
POST /api/templates/default
POST /api/templates/analyze
GET  /api/reports/current
POST /api/reports/current
POST /api/decks/plan
POST /api/decks/generate

GET    /api/assets
POST   /api/assets
GET    /api/assets/:assetId/base64
DELETE /api/assets/:assetId

POST /api/edits/selection/text
POST /api/edits/selection/image
POST /api/edits/selection/image/select
POST /api/edits/slide/reflow

POST /api/qa/check
POST /api/qa/autofix
```

## Validation

Run:

```bash
cd backend
npm run build

cd ../frontend
npm run build
```

Both builds should pass before pushing.
