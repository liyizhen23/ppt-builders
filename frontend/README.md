# AI PPT Builder Frontend

PowerPoint Office Add-in task pane for the AI PPT Builder plugin.

## Technology

- Office.js for PowerPoint host integration.
- React for task pane UI.
- TypeScript for typed API and Office integration boundaries.
- Vite for local development and production builds.
- Lightweight custom CSS for the first milestone; Fluent UI can be introduced after the interaction surface grows.

## Local Development

```bash
npm install
npm run dev
```

The add-in manifest currently points to:

```text
https://localhost:5173/index.html
```

Office add-ins normally require HTTPS for sideloading. If the dev server is plain HTTP, configure local HTTPS certificates before sideloading into PowerPoint.

## API Contract Reserved For Phase 2

The task pane posts report/template files to:

```text
POST /api/decks/generate
```

Expected JSON response:

```json
{
  "deckId": "optional-id",
  "pptxBase64": "base64-encoded-pptx",
  "summary": "optional generation summary",
  "qa": "optional QA summary"
}
```

The insert button calls PowerPoint's `insertSlidesFromBase64` with `keepSourceFormatting`.
