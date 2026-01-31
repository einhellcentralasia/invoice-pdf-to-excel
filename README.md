# Invoice PDF → Excel

Web-based extractor that converts invoice PDFs into structured Excel files.

## 🚀 Features
- Upload → Convert → Auto-download workflow
- Detects `Art. No`, `Qty`, `Price` + AU/Invoice ID
- Generates 3 Excel tables (main + summaries)
- Supports live formulas
- Two languages (RU default, EN optional)
- Style/theme from `styles/style.css`
- Deployable via Docker on Render

## 🌐 Cloudflare Pages (UI) + Render (API)

### 1) UI on Cloudflare Pages (static)
- Build command: *(leave empty)*
- Output directory: `public`
- Set API base URL in `public/config.js`:
  - `window.API_BASE = "https://YOUR-API-URL";`

### 2) API on Render (Docker)
1. Push repo to GitHub
2. In Render:
   - New Web Service → Connect this repo
   - Runtime = **Docker**
   - Port = `8000`
3. Add env var for CORS (comma-separated if multiple):
   - `CORS_ORIGIN=https://YOUR-CLOUDFLARE_PAGES_DOMAIN`
4. Deploy — your API will be live!

## 🛠 Local run (API)
```bash
docker build -t invoice-pdf-to-excel .
docker run -p 8000:8000 invoice-pdf-to-excel
