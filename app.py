import io
import re
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pdfplumber
import pandas as pd

# =========================
# Config / UI
# =========================
st.set_page_config(page_title="Estrattore Consumi Bollette (Enel/Repower) → Excel", page_icon="⚡", layout="wide")
st.title("⚡ Estrattore Consumi Bollette (Enel/Repower) → Excel")
st.caption("Carica una o più bollette in PDF. Per ogni file verrà creato un foglio Excel con due tabelle: Grafico (kWh) e Fatturati (kWh).")

DEBUG = False  # True per mostrare il testo grezzo dei PDF quando il parsing fallisce

# =========================
# Helpers
# =========================

MONTHS_IT = [
    "gennaio","febbraio","marzo","aprile","maggio","giugno",
    "luglio","agosto","settembre","ottobre","novembre","dicembre"
]
ABBR = ["Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"]
MONTH_MAP = {m: i for i, m in enumerate(MONTHS_IT)}
NORM_MONTH = {m: abbr for m, abbr in zip(MONTHS_IT, ABBR)}

# Numeri: 1.518 / 1 518 / 1’518 / 1. 518 / 1518
NUM_RE = r"(?:\d{1,3}(?:[.’'\s]\s*\d{3})*|\d+)"
SEP = r"[ \t]+"

def _normalize(txt: str) -> str:
    """Normalizza testo estratto dai PDF (NBSP, CRLF, spazi multipli)."""
    t = txt.replace("\r", "\n").replace("\xa0", " ")
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{2,}", "\n", t)
    return t

def norm_int(x: str) -> int:
    x = x.strip().replace(" ", "").replace(".", "").replace(",", "").replace("’", "").replace("'", "")
    return int(x) if x else 0

def take_last_12(df: pd.DataFrame) -> pd.DataFrame:
    return df.tail(12).reset_index(drop=True) if len(df) > 12 else df.reset_index(drop=True)

def validate_totals(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    notes: List[str] = []
    if {"F1","F2","F3","Totale"}.issubset(df.columns):
        calc = (df["F1"] + df["F2"] + df["F3"]).astype(int)
        diff = (calc - df["Totale"]).abs()
        mask = diff > 1
        if mask.any():
            notes.append(f"Corretto 'Totale' in {int(mask.sum())} righe (tolleranza 1 kWh).")
            df.loc[mask, "Totale"] = calc.loc[mask]
    return df, notes

def detect_constant(text: str) -> int:
    m = re.search(r"Costante\s*Mis\.?\s*[:=]?\s*(\d+(?:[.,]\d+)?)", text, re.IGNORECASE)
    if m:
        try:
            return int(round(float(m.group(1).replace(",", "."))))
        except Exception:
            pass
    m = re.search(r"\bcostante\b\s*(\d+(?:[.,]\d+)?)", text, re.IGNORECASE)
    if m:
        try:
            return int(round(float(m.group(1).replace(",", "."))))
        except Exception:
            pass
    return 1

# =========================
# Parsers
# =========================

def parse_enel(text: str) -> Optional[pd.DataFrame]:
    """
    Parser robusto per Enel:
    - isola il blocco vicino al titolo
    - cattura le serie tra header e header successivo (F1→F2, F2→F3, F3→Tot, Tot→sezione successiva)
    - taglia l’eventuale riga dei totali di colonna in coda
    - allinea agli ultimi 12 valori coerenti
    """
    # Titolo
    title_pat = re.compile(r"Consumi\s+(?:in\s+kWh\s+)?degli?\s+ultimi\s+\d{1,2}\s+mesi", re.IGNORECASE)
    m_title = title_pat.search(text)
    block = text[m_title.start():m_title.start()+4000] if m_title else text

    # Etichette (best effort)
    month_pat = re.compile(rf"\b({'|'.join(MONTHS_IT)})\b(?:\s+20\d{{2}})?", re.IGNORECASE)
    labels: List[str] = []
    if m_title:
        for mm in month_pat.finditer(block):
            mlow = mm.group(1).lower()
            lab = NORM_MONTH.get(mlow, mlow.title())
            tail = block[mm.end():mm.end()+6]
            y = re.search(r"20\d{2}", tail)
            if y:
                lab = f"{lab} {y.group(0)}"
            labels.append(lab)
        seen = set()
        labels = [l for l in labels if not (l in seen or seen.add(l))]

    # Serie tra header e prossimo header
    def series_between(starts: List[str], stops: List[str], hay: str) -> List[int]:
        P = r"|".join([re.escape(s) for s in starts])
        Q = r"|".join([re.escape(s) for s in stops])
        m = re.search(
            rf"(?:^|\n)\s*(?:{P})\s*[:-]?\s*(?P<body>.*?)"
            rf"(?=(?:^|\n)\s*(?:{Q})\b|$)",
            hay, re.IGNORECASE | re.DOTALL | re.MULTILINE
        )
        if not m:
            m = re.search(rf"(?:{P})\s*[:-]?\s*(?P<body>.*)", hay, re.IGNORECASE | re.DOTALL)
            if not m:
                return []
        nums = re.findall(NUM_RE, m.group("body"))
        return [norm_int(x) for x in nums]

    H1 = ["F1","F 1","Fascia 1","FASCIA 1"]
    H2 = ["F2","F 2","Fascia 2","FASCIA 2"]
    H3 = ["F3","F 3","Fascia 3","FASCIA 3"]
    HT = ["Tot","Totale","TOTALE","TOT"]
    STOP_AFTER_TOT = ["Potenza","kW max","GLOSSARIO","Legenda","Dettaglio","Energia reattiva","REATTIVA","POTENZA"]

    f1_vals = series_between(H1, H2, block)
    f2_vals = series_between(H2, H3, block)
    f3_vals = series_between(H3, HT, block)
    tot_vals = series_between(HT, STOP_AFTER_TOT, block)

    # Fallback sull'intero testo se serve
    if not (f1_vals and f2_vals and f3_vals and tot_vals):
        f1_vals = f1_vals or series_between(H1, H2, text)
        f2_vals = f2_vals or series_between(H2, H3, text)
        f3_vals = f3_vals or series_between(H3, HT, text)
        tot_vals = tot_vals or series_between(HT, STOP_AFTER_TOT, text)

    # ---- TAGLIO dei totali di colonna in coda (se presenti) ----
    def drop_column_total(vals: List[int]) -> List[int]:
        if len(vals) >= 13:
            s = sum(vals[:-1])
            last = vals[-1]
            # Tolleranza 2% o 10 kWh (copre colonne con 12 valori)
            if abs(last - s) <= max(10, int(0.02 * s)):
                return vals[:-1]
        return vals

    f1_vals = drop_column_total(f1_vals)
    f2_vals = drop_column_total(f2_vals)
    f3_vals = drop_column_total(f3_vals)
    tot_vals = drop_column_total(tot_vals)

    # Allineamento: prendi la minima lunghezza > 0 e tieni gli ULTIMI L valori
    lengths = [len(f1_vals), len(f2_vals), len(f3_vals), len(tot_vals)]
    L = min([l for l in lengths if l > 0], default=0)
    if L == 0:
        return None

    f1_vals, f2_vals, f3_vals, tot_vals = f1_vals[-L:], f2_vals[-L:], f3_vals[-L:], tot_vals[-L:]

    # Mantieni al massimo 12 (ultimi 12 mesi)
    if L > 12:
        f1_vals, f2_vals, f3_vals, tot_vals = f1_vals[-12:], f2_vals[-12:], f3_vals[-12:], tot_vals[-12:]
        L = 12

    # Etichette: se esistono, usa le ultime L; altrimenti M1..ML
    if labels and len(labels) >= L:
        labels = labels[-L:]
    else:
        labels = [f"M{i+1}" for i in range(L)]

    return pd.DataFrame({
        "Mese": labels,
        "F1": f1_vals,
        "F2": f2_vals,
        "F3": f3_vals,
        "Totale": tot_vals,
    })

def parse_repower(text: str) -> Optional[pd.DataFrame]:
    """Parser Repower centrato sulla sezione ENERGIA (mese anno F1 F2 F3 Totale)."""
    m = re.search(
        r"\bENERGIA\b.*?(?=\b(POTENZA|POTENZE|COS.?φ|LEGENDA|NOTE|TARIFFE|ALTRE\s+VOCI)\b|$)",
        text, re.IGNORECASE | re.DOTALL
    )
    if not m:
        return None
    block = m.group(0)

    pat = re.compile(
        rf"\b({'|'.join(MONTHS_IT)}){SEP}(20\d{{2}}){SEP}"
        rf"({NUM_RE}){SEP}({NUM_RE}){SEP}({NUM_RE}){SEP}({NUM_RE})\b",
        re.IGNORECASE
    )

    rows = []
    for mm in pat.finditer(block):
        mese = mm.group(1).lower()
        anno = int(mm.group(2))
        f1 = norm_int(mm.group(3))
        f2 = norm_int(mm.group(4))
        f3 = norm_int(mm.group(5))
        tot = norm_int(mm.group(6))
        rows.append((anno, MONTH_MAP.get(mese, 0), f"{NORM_MONTH.get(mese, mese.title())} {anno}", f1, f2, f3, tot))

    if not rows:
        return None

    rows.sort()
    df = pd.DataFrame(
        [dict(Mese=label, F1=f1, F2=f2, F3=f3, Totale=tot) for _,_,label,f1,f2,f3,tot in rows]
    )
    return df

# =========================
# Orchestrator
# =========================

def parse_pdf(file_bytes: bytes) -> Tuple[Optional[pd.DataFrame], str, int, List[str]]:
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        text_pages = [p.extract_text() or "" for p in pdf.pages]
    full_text = _normalize("\n".join(text_pages))

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

    grafico = df.copy(); grafico.insert(1, "Tipo", "Grafico (kWh)")
    fatturati = df.copy()
    fatturati[["F1","F2","F3","Totale"]] = (fatturati[["F1","F2","F3","Totale"]] * const).round().astype(int)
    fatturati.insert(1, "Tipo", "Fatturati (kWh)")

    return pd.concat([grafico, fatturati], axis=0, ignore_index=True), brand, const, notes

# =========================
# UI
# =========================

uploaded = st.file_uploader("Trascina qui i PDF", type=["pdf"], accept_multiple_files=True)

if uploaded:
    logs: List[str] = []
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
                    st.expander(f"Testo grezzo: {up.name}").code(_normalize(raw)[:8000])
                continue
            sheets[sheet_name[:31]] = df_sheet
            logs.append(f"✅ {up.name} → foglio: '{sheet_name[:31]}', {'; '.join(notes)}")
        except Exception as e:
            logs.append(f"❌ {up.name}: errore inatteso - {e}")

    if sheets:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet, df in sheets.items():
                g = df[df["Tipo"]=="Grafico (kWh)"].drop(columns=["Tipo"])
                b = df[df["Tipo"]=="Fatturati (kWh)"].drop(columns=["Tipo"])
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
