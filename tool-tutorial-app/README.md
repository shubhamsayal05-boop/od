# Tool Tutorial Dashboard (Tkinter)

Desktop app for step-by-step tool tutorials. Each tool has its own dashboard with guided steps and a linked PowerPoint file.

## Requirements

- Python 3.10+
- **tkinter** (included with standard Python on Windows; on Linux install `python3-tk`)

## Run

```bash
cd tool-tutorial-app
python main.py
```

On Windows you can also double-click `main.py` if Python is associated with `.py` files.

## Add your tools and PPTs

1. Copy each `.pptx` into `presentations/`
   - Example: `presentations/odriv-tutorial.pptx`

2. Add an entry in `tools.json`:

```json
{
  "id": "my-tool",
  "name": "My Tool",
  "description": "Short description of the tool.",
  "category": "Analysis",
  "pptFile": "presentations/my-tool-tutorial.pptx",
  "estimatedMinutes": 15,
  "steps": [
    {
      "title": "Open the tool",
      "description": "What the user should do in this step.",
      "tips": ["Optional tip"]
    }
  ]
}
```

3. Restart the app.

## Using your existing PPTs

- Put all PowerPoint files in the `presentations/` folder next to `main.py`
- Set `pptFile` to the relative path for each tool
- Click **Open Presentation** on any tutorial page to launch the PPT in PowerPoint (or your default app)

## Folder layout

```
tool-tutorial-app/
├── main.py              ← run this
├── tools.json           ← tool list + steps
└── presentations/       ← your .pptx files here
```
