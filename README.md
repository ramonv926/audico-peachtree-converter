# Audico EC → Peachtree Quote Converter

Web app to convert monthly EC (estado de cuenta) xlsx files into Peachtree/Sage 50
quote CSVs ready for import.

**Live app:** _deploy it first, then paste your URL here_

## For developers

- `app.py` — Streamlit web interface
- `convert.py` — core conversion logic (also works as CLI: `python convert.py file.xlsx`)
- `config/hotel_mapping.json` — hotel → Peachtree customer mapping + skip rules + defaults
- `requirements.txt` — Python dependencies

## Deploying

See [`DEPLOYMENT.md`](DEPLOYMENT.md) for step-by-step instructions.
First deploy: ~15 minutes. Subsequent updates: ~30 seconds.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open http://localhost:8501 in your browser.
