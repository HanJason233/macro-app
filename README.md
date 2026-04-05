# Macro App

Windows desktop macro workflow tool built with PySide6.

## Features

- Workflow editor with big-node and small-node execution model
- Window-scoped automation
- Absolute click, relative click, image matching, and OCR text matching
- Window management actions such as resize, minimize, and close
- Screenshot overlay tools for capture and template creation

## Project Structure

```text
macro_app/
  app.py                 # Application entrypoint
  constants.py           # Action definitions and labels
  models.py              # Workflow data helpers
  services/              # OCR, capture, window, and runner services
  ui/                    # Main window, overlays, dialogs, and panels
examples/                # Optional sample workflow files
run.py                   # Recommended local start script
MacroApp.spec            # PyInstaller build config
```

## Requirements

- Windows
- Python 3.12 or 3.13

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

Run locally:

```bash
python run.py
```

Open the sample workflow if needed:

```text
examples/sample_workflow.json
```

Build an executable:

```bash
pyinstaller MacroApp.spec
```

## Notes

- `build/`, `dist/`, and `.macro_app_state.json` are runtime/generated files and are ignored by Git.
- OCR is powered by `rapidocr_onnxruntime`.
