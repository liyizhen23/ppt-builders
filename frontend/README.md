# AI PPT Builder Frontend

PowerPoint Office Add-in task pane for the AI PPT Builder plugin.

## Technology

- Office.js for PowerPoint host integration.
- React for task pane UI.
- TypeScript for typed API and Office integration boundaries.
- Vite for local development and production builds.
- Lightweight Apple-inspired custom CSS for the first milestones; Fluent UI can be introduced only if the control surface outgrows this system.

## Interface Direction

The task pane follows an Apple-inspired utility surface:

- Single interactive accent: Action Blue `#0066cc`.
- System/SF-style font stack with 17px body rhythm where space allows.
- White utility panels on `#f5f5f7` parchment canvas.
- Pill-shaped primary and secondary actions.
- Hairline borders instead of chrome shadows.
- No decorative gradients, no card/button shadows, no second accent color.
- Focus and active states are quiet: blue focus outline and `scale(0.96)` press feedback.
- Runtime status stays compact so report/template controls remain the primary surface.

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

The task pane calls the backend through relative `/api/...` paths during local development. Vite proxies those requests to:

```text
http://127.0.0.1:3000
```

The task pane posts report/template files to:

```text
POST /api/decks/generate
```

The report file is required only for the first generation in a session. Once uploaded, the backend saves it as the current report and future generate/plan calls can omit `report`.

The task pane now has two modes:

- Generate: creates PPT pages from a report and optional template. This mode needs a report unless a current report is already saved on the backend.
- Edit: modifies the current PowerPoint selection through a chat-style confirmation flow. This mode does not show report/template upload controls and does not require a report.

The template file is optional. If the user does not choose one, the backend uses the saved default template, falling back to the bundled Tsinghua template.

The task pane can save a selected template as the future default through:

```text
POST /api/templates/default
```

Expected JSON response:

```json
{
  "deckId": "optional-id",
  "pptxBase64": "base64-encoded-pptx",
  "summary": "optional generation summary",
  "qa": "optional QA summary",
  "deckPlan": "schema-validated planning object",
  "templateReplacement": "selected template slide and replaced slots"
}
```

The insert button calls PowerPoint's `insertSlidesFromBase64` with `keepSourceFormatting`.

AI API keys are not used in frontend code. The task pane can only read non-secret AI status from:

```text
GET /api/settings/ai
```

Selected text editing uses:

```text
POST /api/edits/selection/text
```

The backend returns an `editPlan` with `replacementText`, `needsConfirmation`, optional `clarificationQuestion`, and QA metadata. The frontend applies the replacement only after the user confirms.

Selected image replacement and first-pass image format guidance use:

```text
POST /api/edits/selection/image
```

The first image MVP supports selecting a local replacement image, generating an image edit plan, and applying the image to the current PowerPoint selection through Office.js `setSelectedDataAsync` with image coercion. Fine-grained crop, rounded corners, and exact shape alignment remain constrained by PowerPoint Office.js and are represented as plan/QA guidance until the renderer can operate on PPTX XML directly.

When the user does not know which existing image belongs on a slide, the task pane can accept multiple candidate images and page text, then call:

```text
POST /api/edits/selection/image/select
```

This MVP selects from local candidate images using page text, the user instruction, image file names, and optional metadata. The actual image bytes stay in the frontend; the backend only receives candidate metadata, then returns the selected image id and reason. Visual image understanding is not enabled yet, so descriptive file names such as `POI分类图表.png` materially improve selection quality.

The task pane also has a persistent local asset library:

```text
POST /api/assets
GET  /api/assets
GET  /api/assets/:assetId/base64
DELETE /api/assets/:assetId
```

Images and table files are stored under local `asset-library/`, which is ignored by Git. The image editing page can save files into that library and later ask AI to choose from the saved library without re-uploading all candidates every time.

Current-slide reflow uses:

```text
POST /api/edits/slide/reflow
```

It returns a replacement PPTX slide as Base64. The frontend inserts it into PowerPoint for side-by-side comparison with the original current slide.

Current-page QA uses:

```text
POST /api/qa/check
POST /api/qa/autofix
```

The QA MVP checks excessive text, missing hierarchy, too many page-level points, and missing visual intent. Autofix rewrites the page text into a compact title plus bullet structure before generating a replacement slide.
