# MMI Uploader — Deploy package

Upload this **entire folder** to your server. Everything needed to run the app is inside this folder.

## Contents

- `server.js` — Node app entry point
- `package.json` / `package-lock.json` — dependencies
- `lib/` — RTF parsing and HTML generation
- `public/` — upload form (index.html)
- `og code/` — template (`dec_mmi_2025_fmg.html`)
- `generated reports/` — where generated HTML files are saved (created automatically if missing)

## On your server

1. Upload the whole `mmi-uploader-deploy` folder (or rename it; the app runs from this folder).
2. In that folder, run:
   ```bash
   npm install
   npm start
   ```
3. The app listens on port 3000 (or `process.env.PORT` if set). Point your domain or reverse proxy at this port, or open `http://your-server:3000`.

## Requirements

- Node.js (v14 or newer)
- No `node_modules` in the zip — run `npm install` on the server after upload.
