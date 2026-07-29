# Tool Tutorial Dashboard

A web dashboard where each internal tool has its own tutorial page with step-by-step instructions and linked PowerPoint presentations.

## Quick start

```bash
cd tutorial-dashboard
npm install
npm run dev
```

Open http://localhost:5173

## Add a new tool

1. Copy your `.pptx` file into `public/presentations/` (e.g. `my-tool-tutorial.pptx`).
2. Add an entry to `src/data/tools.json`:

```json
{
  "id": "my-tool",
  "name": "My Tool",
  "description": "Short description of what this tool does.",
  "category": "Analysis",
  "icon": "🔧",
  "pptFile": "/presentations/my-tool-tutorial.pptx",
  "pptEmbedUrl": "",
  "estimatedMinutes": 15,
  "steps": [
    {
      "title": "Step title",
      "description": "What the user should do in this step.",
      "tips": ["Optional helpful tip"]
    }
  ]
}
```

3. Restart the dev server if it is already running.

## Linking your existing PPTs

| Method | How |
|--------|-----|
| **Local file** | Place `.pptx` in `public/presentations/` and set `pptFile` |
| **OneDrive / SharePoint** | Upload PPT, get embed link, set `pptEmbedUrl` |
| **Google Slides** | File → Share → Publish to web → embed, set `pptEmbedUrl` |

When `pptEmbedUrl` is set, the presentation is shown inline in the tutorial page.

## Build for production

```bash
npm run build
npm run preview
```

Deploy the `dist/` folder to any static host (GitHub Pages, Netlify, Vercel, internal server).

## Project structure

```
tutorial-dashboard/
├── public/presentations/   ← put your .pptx files here
├── src/
│   ├── data/tools.json     ← tool list + steps (edit this)
│   ├── components/
│   └── pages/
└── package.json
```
