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

## 🧱 Deployment on Render
1. Push repo to GitHub
2. In Render:
   - New Web Service → Connect this repo
   - Runtime = **Docker**
   - Port = `8000`
3. Deploy — your app will be live!

## 🛠 Local run
```bash
docker build -t invoice-pdf-to-excel .
docker run -p 8000:8000 invoice-pdf-to-excel
