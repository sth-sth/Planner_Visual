# Planner_Visual

Planner PMO Workbench is now structured as a pure static website that can be deployed directly to Vercel. The app runs entirely in the browser: users upload a local Planner Excel file, the workbook is parsed client-side, and no backend service is required.

## Project Structure

- `index.html`: app entry page
- `styles.css`: all page styles
- `app.js`: client-side parsing, filtering, charts, and rendering
- `xlsx.full.min.js`: local SheetJS runtime bundled in the repo
- `vercel.json`: Vercel static deployment configuration

## Local Preview

Because this is a static site, you only need a simple local HTTP server.

```bash
python3 -m http.server 4173
```

Then open:

```text
http://localhost:4173
```

## Deploy to Vercel

### Option 1: Import Git Repository

1. Push this repository to GitHub.
2. In Vercel, click `Add New -> Project`.
3. Import the repository.
4. Keep the default settings:
	- Framework Preset: `Other`
	- Build Command: empty
	- Output Directory: empty
5. Click `Deploy`.

### Option 2: Deploy with Vercel CLI

```bash
npm i -g vercel
vercel
```

For production deployment:

```bash
vercel --prod
```

## Deployment Notes

- No server-side rendering is needed.
- No API route is needed.
- All Excel processing stays in the browser.
- Static assets are referenced with relative paths, so the site can be hosted directly from the repository root.

## Recommended Vercel Settings

- Root Directory: repository root
- Install Command: empty
- Build Command: empty
- Output Directory: empty

## Browser Usage

1. Open the deployed site.
2. Upload a Planner-exported `.xlsx` or `.xls` file.
3. Review the Gantt view, analytics, PMO overview, and task table.
4. Export the current filtered task set as CSV if needed.