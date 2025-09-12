
# Estrattore Consumi Bollette (Enel/Repower) → Excel

Prototype Streamlit app to parse Enel/Repower electricity bills (PDF) and export an Excel with one sheet per file:
- **Grafico (kWh)**: values as shown in the bill's "consumi / andamento storico" chart
- **Fatturati (kWh)**: values multiplied by **Costante di misura** (detected from the PDF; defaults to x1)

## Quick start (local)
```bash
python -m venv .venv
. .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```
Open the URL shown by Streamlit and upload your PDFs.

## Notes
- Supports **Enel** (block "Consumi in kWh degli ultimi ... mesi") and **Repower** (page 4 "Andamento storico – Energia").
- Keeps **last 12 months** if more are present.
- Validates **Totale = F1+F2+F3** (±1 kWh), fixes if off.
- Tries to auto-detect **Costante di misura** (e.g., "Costante Mis. 25,00").
