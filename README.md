# Bluebeam Like Plan Viewer

This workspace now uses your React + PDF.js viewer as the base instead of the earlier static prototype.

## Stack

- React 18
- TypeScript
- Vite
- Tailwind CSS
- `pdfjs-dist`
- `framer-motion`
- `lucide-react`

## Main Files

- [src/components/BluebeamLikePlanViewer.tsx](C:\Users\Shawn Graham\OneDrive - Sigal Development\Desktop\Codex\src\components\BluebeamLikePlanViewer.tsx)
- [src/App.tsx](C:\Users\Shawn Graham\OneDrive - Sigal Development\Desktop\Codex\src\App.tsx)
- [src/components/ui/button.tsx](C:\Users\Shawn Graham\OneDrive - Sigal Development\Desktop\Codex\src\components\ui\button.tsx)
- [package.json](C:\Users\Shawn Graham\OneDrive - Sigal Development\Desktop\Codex\package.json)

## What Was Added Around Your Base

- A complete Vite app scaffold so the component has a real project to live in
- Local `@/components/ui/*` replacements so the imports resolve cleanly
- Tailwind and alias configuration
- Zoom in, zoom out, fit-to-page, and mouse-wheel zoom support
- A compare-page selector in the inspector

## Notes

- Project/session saves persist markup data, layer state, users, and OCR text metadata.
- The current export is still a JSON manifest mock, not a true burned-in annotated PDF.
- Loading a saved project restores markup state, but it does not restore the original PDF binary itself.

## Deploy To Vercel

This project is prepared for Vercel with [vercel.json](C:\Users\Shawn Graham\OneDrive - Sigal Development\Desktop\Codex\vercel.json).

### Deploy Settings

- Framework preset: `Vite`
- Build command: `npm run build`
- Output directory: `dist`

### Quick Publish Flow

1. Push this project to GitHub.
2. Create a new project in Vercel and import the GitHub repo.
3. Confirm the detected settings:
   - Build Command: `npm run build`
   - Output Directory: `dist`
4. Deploy.
5. Share the generated Vercel URL with your colleague.

### Notes For Sharing

- The app itself will be shareable by link once deployed.
- PDF files are not bundled by default, so each user still needs access to the plan files unless we add hosted sample PDFs into the app.
