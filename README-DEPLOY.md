# File Translator Deploy Guide

This folder is the clean deployable version of the real file translator.

It contains:

- `backend/`: FastAPI backend for real translation
- `frontend/`: upload web page
- `render.yaml`: Render backend deployment config
- `frontend/vercel.json`: Vercel static frontend config

Do not upload your local `.venv`, translated test files, or API key batch files.

## Step 1: Upload This Folder To GitHub

Create a GitHub repository, for example:

```text
file-translator-real
```

Upload everything inside this `file-translator-deploy` folder.

## Step 2: Deploy Backend On Render

1. Open https://render.com
2. Sign in with GitHub
3. Click `New +`
4. Choose `Blueprint`
5. Select your GitHub repository
6. Render will detect `render.yaml`
7. When Render asks for `OPENAI_API_KEY`, paste your API key there
8. Click deploy

After deployment, Render will give you a backend URL like:

```text
https://file-translator-backend.onrender.com
```

Open this URL with `/health`:

```text
https://file-translator-backend.onrender.com/health
```

You should see:

```json
{"status":"ok","translation_mode":"openai"}
```

## Step 3: Connect Frontend To Backend

Open:

```text
frontend/config.js
```

Replace:

```js
window.FILE_TRANSLATOR_API_BASE_URL = "https://your-render-backend-url.onrender.com";
```

with your real Render backend URL.

Commit this change to GitHub.

## Step 4: Deploy Frontend On Vercel

1. Open https://vercel.com
2. Sign in with GitHub
3. Click `Add New Project`
4. Select the same GitHub repository
5. Set `Root Directory` to:

```text
frontend
```

6. Deploy

Vercel will give you a frontend URL like:

```text
https://file-translator-real.vercel.app
```

## Step 5: Test Online Translation

1. Open the Vercel frontend URL
2. Upload `.docx`, `.xlsx`, or text-based `.pdf`
3. Choose target language
4. For PDF, choose:
   - Professional PDF overlay
   - Or Word export
5. Click translate
6. The translated file should download

## Important Security Notes

- Never put `OPENAI_API_KEY` in frontend files.
- Only store `OPENAI_API_KEY` in Render environment variables.
- If your API key was ever uploaded to GitHub, delete it from OpenAI and create a new key.

## Current PDF Limitations

Professional PDF mode works best for text-based engineering drawings.

It does not yet handle:

- Scanned image-only PDF without OCR
- Perfect CAD-grade text fitting
- Every possible rotated or curved annotation

For scanned drawings, add OCR in a later stage.
