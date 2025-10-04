Sheet Access Pro (Next.js) – Remove worksheet/workbook protection from .xlsx/.xlsm via browser uploads.

## Getting Started

First, run the development server:

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

## How it works
- Client uploads multiple files via drag-and-drop or file picker
- API route `/api/unlock` processes uploads in-memory using JSZip + xmldom + xpath
- Removes `<sheetProtection>`, `<workbookProtection>`, and `<fileSharing>` nodes
- Returns base64 data URLs for immediate download of `_unlocked.xlsx`

## Limits and privacy
- Designed for Vercel’s serverless environment
- Files are processed in-memory only; nothing is persisted
- Request parsing uses multipart form-data; increase limits in Pro plans if needed
- For very large workbooks, consider self-hosting on platforms with higher limits

## Deploy on Vercel
- Push the `next-app/` directory as your project root to a new repo or set it as the root in Vercel
- Ensure `next.config.ts` has `api.bodyParser = false` to support multipart form-data
