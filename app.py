import io
import re
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pdfplumber
import pandas as pd

# ---------------------------
# Config
# ---------------------------
st.set_page_config(page_title="Estrattore Consumi Bollette (Enel/Repower) → Excel", page_icon="⚡")
st.title("⚡ Estrattore Consumi Bollette (Enel/Repower) → Excel")
st.caption("Carica una o più bollette in PDF. Per ogni file verrà creato un foglio Excel con due tabelle: Grafico (kWh) e Fatturati (kWh).")

DEBUG = False  # metti True per vedere il testo grezzo dei PDF quando il parsing fallisce

# ---------------------------
# Helpers
# ---------------------------

MONTHS_IT = [
    "gennaio","febbraio","marzo","aprile","maggio","giugno",
    "luglio","agosto","settembre","ottobre","novembre","dicembre"
]
ABBR = ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"]
MONTH_MAP = {m:i for i,m in enumerate(MONTHS_IT)}
NORM_MONTH = {m:abbr for m,abbr in zip(MONTHS_IT, ABBR)}

NUM_RE = r"(?:\d{1,3}(?:[.\s’']\d{3})*|\d+)"  # supporta 1.234, 1 234, 1’234
SEP = r"[ \t]+"

def norm_int(x: str) -> int:
    x = x.strip().replace(" ", "").replace(".", "").replace(",", "").replace("’","").replace("'","")
    return int(x) if x else 0

def take_last_12(df: pd.DataFrame) -> pd.DataFrame:
    if len(df) > 12:
        return df.tail(12).reset_index(drop=True)
    return df.reset_index(drop=True)

def validate_totals(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    notes = []
    if {"F1","F2","F3","Totale"}.issubset(df.columns):
        calc = (df["F1"] + df["F2"] + df["F3"]).astype(int)
        diff = (calc - df["Totale"]).abs()
        mask = diff > 1
        if mask.any():
            n = mask.sum()
            notes.append(f"Corretto 'Totale' in {n} righe (tolleranza 1 kWh).")
            df.loc[mask, "Totale"] = calc.loc[mask]
    return df, notes

def detect_constant(text: str) -> int:
    m = re.search(r"Costante\s*Mis\.?\s*[:=]?\s*(\d+(?:[.,]\d+)?)", text, re.IGNORECASE)
    if m:
        val = m.group(1).replace(",", ".")
        try:
            return int(round(float(val)))
        except:
            pass
    m = re.search(r"\bcostante\b\s*(\d+(?:[.,]\d+)?)", text, re.IGNORECASE)
    if m:
        val = m.group(1).replace(",", ".")
        try:
            return int(round(float(val)))
        except:
            pass
    return 1

# ---------------------------
# Parsers
# ---------------------------

def parse_enel(text: str) -> Optional[pd.DataFrame]:
    title_pat = re.compile(r"Consumi\s+(?:in\s+kWh\s+)?degli?\s+ultimi\s+\d{1,2}\s+mesi", re.IGNORECASE)
    m_title = title_pat.search(text.replace("\r", ""))
    if not m_title:
        return None
    start = m_title.start()
    block = text[start:start+2500]

    month_pat = re.compile(rf"\b({'|'.join(MONTHS_IT)})\b(?:\s+20\d{{2}})?", re.IGNORECASE)
    labels = []
    for mm in month_pat.finditer(block):
        mlow = mm.group(1).lower()
        lab = NORM_MONTH.get(mlow, mlow.title())
        tail = block[mm.end():mm.end()+6]
        y = re.search(r"20\d{2}", tail)
        if y:
            lab = f"{lab} {y.group(0)}"
        labels.append(lab)
    uniq, seen = [], set()
    for l in labels:
        if l not in seen:
            seen.add(l); uniq.append(l)
    labels = uniq or [f"Mese {i+1}" for i in range(12)]

    def grab_line(prefixes):
        pfx = r"|".join([re.escape(p) for p in prefixes])
        m = re.search(rf"(?m)^\s*(?:{pfx})\s*[:-]?\s*((?:{NUM_RE}\s*)+)", block, re.IGNORECASE)
        if not m:
            return []
        nums = re.findall(NUM_RE, m.group(1))
        return [norm_int(x) for x in nums]

    f1_vals = grab_line(["F1","F 1","Fascia 1","FASCIA 1"])
    f2_vals = grab_line(["F2","F 2","Fascia 2","FASCIA 2"])
    f3_vals = grab_line(["F3","F 3","Fascia 3","FASCIA 3"])
    tot_vals = grab_line(["Tot","Totale","TOTALE","TOT"])

    lengths = [len(f1_vals), len(f2_vals), len(f3_vals), len(tot_vals), len(labels)]
    L = min([l for l in lengths if l>0] + [0])
    if L == 0:
        return None

    df = pd.DataFrame({
        "Mese": labels[:L],
        "F1": f1_vals[:L],
        "F2": f2_vals[:L],
        "F3": f3_vals[:L],
        "Totale": tot_vals[:L]
    })
    return df

def parse_repower(text: str) -> Optional[pd.DataFrame]:
    m = re.search(r"Andamento\s+storico.*?Energia.*?(?=Potenza|Cosφ|Legenda|$)", text, re.IGNORECASE | re.DOTALL)
    if not m:
        return None
    block = m.group(0)

    rows = []
    pat = re.compile(
        rf"\b({'|'.join(MONTHS_IT)}){SEP}(20\d{{2}}){SEP}({NUM_RE}){SEP}({NUM_RE}){SEP}({NUM_RE}){SEP}({NUM_RE})",
        re.IGNORECASE
    )
    for mm in pat.finditer(block):
        mese = mm.group(1).lower()
        anno = mm.group(2)
        f1 = norm_int(mm.group(3)); f2 = norm_int(mm.group(4)); f3 = norm_int(mm.group(5)); tot = norm_int(mm.group(6))
        rows.append((int(anno), MONTH_MAP[mese], f"{NORM_MONTH[mese]} {anno}", f1, f2, f3, tot))

    if not rows:
        return None

    rows.sort()
    df = pd.DataFrame([
        dict(Mese=label, F1=f1, F2=f2, F3=f3, Totale=tot)
        for _,_,label,f1,f2,f3,tot in rows
    ])
    return df

def parse_pdf(file_bytes: bytes) -> Tuple[Optional[pd.DataFrame], str, int, List[str]]:
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        text_pages = [p.extract_text() or "" for p in pdf.pages]
    full_text = "\n".join(text_pages)

    notes: List[str] = []

    df = parse_repower(full_text)
    brand = "Repower" if df is not None else None

    if df is None:
        df = parse_enel(full_text)
        if df is not None:
            brand = "Enel"

    if df is None:
        return None, "Non riconosciuto", 1, ["Layout non riconosciuto (né Repower né Enel)."]

    df = take_last_12(df)
    df, val_notes = validate_totals(df)
    notes.extend(val_notes)

    const = detect_constant(full_text)
    notes.append(f"Costante di misura: x{const}")

    sheet = brand

    grafico = df.copy(); grafico.insert(1, "Tipo", "Grafico (kWh)")
    fatturati = df.copy()
    fatturati[["F1","F2","F3","Totale"]] = (fatturati[["F1","F2","F3","Totale"]] * const).round().astype(int)
    fatturati.insert(1, "Tipo", "Fatturati (kWh)")

    return (pd.concat([grafico, fatturati], axis=0, ignore_index=True), sheet, const, notes)

# ---------------------------
# UI
# ---------------------------

uploaded = st.file_uploader("Trascina qui i PDF", type=["pdf"], accept_multiple_files=True)

if uploaded:
    logs = []
    sheets: Dict[str, pd.DataFrame] = {}
    for up in uploaded:
        try:
            data = up.read()
            df_sheet, sheet_name, const, notes = parse_pdf(data)
            if df_sheet is None:
                logs.append(f"❌ {up.name}: {notes[0] if notes else 'Errore parsing'}")
                if DEBUG:
                    with pdfplumber.open(io.BytesIO(data)) as pdf:
                        raw = "\n".join([p.extract_text() or "" for p in pdf.pages])
                    st.expander(f"Testo grezzo: {up.name}").code(raw[:8000])
                continue
            sheets[sheet_name[:31]] = df_sheet
            logs.append(f"✅ {up.name} → foglio: '{sheet_name[:31]}', {'; '.join(notes)}")
        except Exception as e:
            logs.append(f"❌ {up.name}: errore inatteso - {e}")

    if sheets:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet, df in sheets.items():
                g = df[df['Tipo']=="Grafico (kWh)"].drop(columns=["Tipo"])
                b = df[df['Tipo']=="Fatturati (kWh)"].drop(columns=["Tipo"])
                g.to_excel(writer, sheet_name=sheet, index=False, startrow=2)
                b.to_excel(writer, sheet_name=sheet, index=False, startrow=len(g)+5)
        st.download_button(
            "⬇️ Scarica Excel (Consumi_F1F2F3_per_Fattura.xlsx)",
            data=output.getvalue(),
            file_name="Consumi_F1F2F3_per_Fattura.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("### Log elaborazione")
    for line in logs:
        st.write(line)
else:
    st.info("Carica uno o più PDF per iniziare.")
