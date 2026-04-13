# TraCuuNNT Streamlit

Streamlit app for Vietnamese taxpayer lookup.

## Deploy on Streamlit Community Cloud

Use this app entrypoint:

```text
streamlit_ui/app.py
```

The app reads MST values from pasted text or uploaded `.txt`, `.csv`, `.xlsx`, `.xlsm` files, queries the GDT taxpayer lookup page, and returns an Excel file.

If dependency install fails, redeploy with Python 3.12 in Advanced settings.

## Local Run

```bash
pip install -r requirements.txt
streamlit run streamlit_ui/app.py
```
