# River & Sediment Discharge Data Processor (GitHub Pages)

This folder contains a **pure HTML/CSS/JavaScript** version that runs entirely in the browser. No Python or server required — works on GitHub Pages.

## Deploy to GitHub Pages

1. Create a GitHub repository
2. Upload the contents of this `upload` folder:
   - `index.html`
   - `styles.css`
   - `script.js`
3. In repo **Settings → Pages**, set source to your branch (e.g. `main`) and folder (e.g. `/upload` or root)
4. Your page will be at: `https://yourusername.github.io/your-repo/upload/` (or `/` if you put files in root)

## Files

- `index.html` - Main page (includes SheetJS from CDN)
- `styles.css` - Styling
- `script.js` - All processing logic (runs in browser)

## How it works

- Uses [SheetJS](https://sheetjs.com/) (loaded from CDN) to read/write Excel files
- All processing happens in your browser — no data is sent to any server
- Files are processed locally and downloaded directly
