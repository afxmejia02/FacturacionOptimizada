# Web UI for Validation

This folder contains a simple Flask web interface that reuses the extraction and comparison
functions defined in `facturacion/gui_validation_app.py`.

Run locally:

```bash
python -m pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:8500/ and upload a PDF and an Excel file. Select the type (equipos/servicios/perfiles).

Notes:
- The app imports the existing module from the parent folder — don't move `facturacion`.
- No GUI windows are created; the code instantiates the validator class without running the tkinter init.
