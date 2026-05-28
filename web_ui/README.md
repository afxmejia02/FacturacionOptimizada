# Web UI for Validation

This folder contains a Streamlit web interface that reuses the extraction and comparison
functions defined in `facturacion/gui_validation_app.py`.

Run locally:

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Open the Streamlit URL shown in the terminal and upload the required files. Select the type
(equipos/servicios/perfiles) or the reconciliation mode.

Notes:
- The app imports the existing module from the parent folder — don't move `facturacion`.
- No GUI windows are created; the code instantiates the validator class without running the tkinter init.
