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

## Install The Add-in Into PowerPoint

This section is written so Codex or Claude Code can read it and install the local add-in into PowerPoint on Windows.

### 1. Prepare Dependencies And Certificates

Run from the repository root:

```powershell
cd D:\Users\LYZ20\Documents\ppt-builders

cd backend
npm.cmd install

cd ..\frontend
npm.cmd install
npx.cmd office-addin-dev-certs install
```

The manifest loads the task pane from:

```text
https://localhost:5173/index.html
```

PowerPoint must trust the local HTTPS certificate, and both local servers must be running.

### 2. Start The Local Services

Open two terminals, or let an agent start both hidden processes:

```powershell
Start-Process -FilePath npm.cmd -ArgumentList 'run','dev' -WorkingDirectory 'D:\Users\LYZ20\Documents\ppt-builders\backend' -WindowStyle Hidden
Start-Process -FilePath npm.cmd -ArgumentList 'run','dev' -WorkingDirectory 'D:\Users\LYZ20\Documents\ppt-builders\frontend' -WindowStyle Hidden
```

Verify:

```powershell
Invoke-WebRequest -Uri http://127.0.0.1:3000/health -UseBasicParsing
Invoke-WebRequest -Uri https://localhost:5173/index.html -UseBasicParsing
```

Both should return HTTP 200. If PowerPoint shows "加载项错误", the most common cause is that `https://localhost:5173/index.html` is not reachable.

### 3. Copy The Manifest To A Local Catalog

```powershell
$catalog = 'C:\OfficeAddinCatalog\ppt-builders'
New-Item -ItemType Directory -Force -Path $catalog
Copy-Item -Force 'D:\Users\LYZ20\Documents\ppt-builders\frontend\manifest.xml' "$catalog\manifest.xml"
```

### 4. Share The Catalog Folder

Windows desktop Office sideloading expects a trusted network shared folder catalog. Creating an SMB share usually requires an elevated PowerShell session.

Run PowerShell as Administrator:

```powershell
New-SmbShare -Name ppt-builders-addin -Path 'C:\OfficeAddinCatalog\ppt-builders' -ChangeAccess $env:USERNAME
```

The catalog URL is then:

```text
\\localhost\ppt-builders-addin
```

If `New-SmbShare` fails because the terminal is not elevated, create the share manually in Windows folder properties, or run this step from an Administrator terminal.

### 5. Trust The Catalog In PowerPoint

In PowerPoint:

1. Open `File -> Options -> Trust Center -> Trust Center Settings`.
2. Open `Trusted Add-in Catalogs`.
3. Add `\\localhost\ppt-builders-addin`.
4. Check `Show in Menu`.
5. Restart PowerPoint.
6. Open `Home -> Add-ins -> Shared Folder`.
7. Select `AI PPT Builder`.

After the task pane opens, use the add-in while the backend and frontend dev servers keep running.

### 6. Agent Checklist

When Codex or Claude Code is asked to install this add-in locally, it should:

```text
1. Install backend and frontend npm dependencies.
2. Install Office dev certificates with office-addin-dev-certs.
3. Start backend on http://127.0.0.1:3000.
4. Start frontend on https://localhost:5173.
5. Copy frontend/manifest.xml to C:\OfficeAddinCatalog\ppt-builders\manifest.xml.
6. Create or verify the SMB share \\localhost\ppt-builders-addin.
7. Tell the user to add that share under PowerPoint Trusted Add-in Catalogs and restart PowerPoint.
8. Confirm /health and https://localhost:5173/index.html return HTTP 200 before debugging PowerPoint.
```

### 7. Office On The Web Alternative

For PowerPoint on the web, use `Home -> Add-ins -> More Add-ins -> Upload My Add-in`, then upload:

```text
frontend/manifest.xml
```

The local frontend and backend still need to be reachable from the browser.

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
