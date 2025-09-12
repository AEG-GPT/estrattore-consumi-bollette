
import io
import re
import sys
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pdfplumber
import pandas as pd

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

NUM_RE = r"(?:\d{1,3}(?:[\.\s]\d{3})*|\d+)"  # 1.234 or 1 234 or 1234
SEP = r"[ \t]+"

def norm_int(x: str) -> int:
    x = x.strip().replace(" ", "").replace(".", "").replace(",", "")
    return int(x)

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
    # Enel: "Costante Mis. 25,00" / "Costante Mis. 25.00"
    m = re.search(r"Costante\s*Mis\.?\s*[:=]?\s*(\d+(?:[.,]\d+)?)", text, re.IGNORECASE)
    if m:
        val = m.group(1).replace(",", ".")
        try:
            return int(round(float(val)))
        except:
            pass
    # Repower: "costante 1,00" nelle tabelle misure
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

def parse_repower(text: str) -> Optional[pd.DataFrame]:
    """
    Cerca il blocco 'Andamento storico dei prelievi' sezione 'Energia' con righe:
    'giugno 2024 368 358 792 1.518'
    """
    # Trova blocco con intestazione
    m = re.search(r"Andamento\s+storico.*?Energia.*?(?=Potenza|Cosφ|Legenda|$)", text, re.IGNORECASE | re.DOTALL)
    if not m:
        return None
    block = m.group(0)

    rows = []
    # pattern: "<mese> <anno> F1 F2 F3 Totale"
    pat = re.compile(
        rf"\b({'|'.join(MONTHS_IT)}){SEP}(20\d{{2}}){SEP}({NUM_RE}){SEP}({NUM_RE}){SEP}({NUM_RE}){SEP}({NUM_RE})",
        re.IGNORECASE
    )
    for mm in pat.finditer(block):
        mese = mm.group(1).lower()
        anno = mm.group(2)
        f1 = norm_int(mm.group(3))
        f2 = norm_int(mm.group(4))
        f3 = norm_int(mm.group(5))
        tot = norm_int(mm.group(6))
        rows.append((int(anno), MONTH_MAP[mese], f"{NORM_MONTH[mese]} {anno}", f1, f2, f3, tot))

    if not rows:
        return None

    rows.sort()  # anno, mese index
    data = [dict(Mese=label, F1=f1, F2=f2, F3=f3, Totale=tot) for _,_,label,f1,f2,f3,tot in rows]
    df = pd.DataFrame(data)
    return df

def parse_enel(text: str) -> Optional[pd.DataFrame]:
    """
    Parser per blocco 'Consumi in kWh degli ultimi ... mesi' in bolletta Enel.
    Si aspetta:
      - elenco mesi in ordine
      - righe con F1/F2/F3/Tot valori per ciascun mese
    """
    m = re.search(r"Consumi\s+in\s+kWh\s+degli\s+ultimi.*?(?=Potenza|kW\smax|GLOSSARIO|Legenda|Dettaglio|$)",
                  text, re.IGNORECASE | re.DOTALL)
    if not m:
        return None
    block = m.group(0)

    # Estrai sequenza mesi con anni (possono essere verticalizzati nel PDF; proviamo pattern mese anno ripetuti)
    months = []
    month_pat = re.compile(rf"\b({'|'.join(MONTHS_IT)})\s*(20\d{{2}})\b", re.IGNORECASE)
    for mm in month_pat.finditer(block):
        mese = mm.group(1).lower()
        anno = mm.group(2)
        months.append((int(anno), MONTH_MAP[mese], f"{NORM_MONTH[mese]} {anno}"))

    # Dedup e ordina
    months = sorted(list({x:None for x in months}.keys()))
    labels = [lbl for _,_,lbl in months]

    if len(labels) == 0:
        # fallback: se i mesi non sono accoppiati all'anno, prova solo i nomi (assume stesso anno a blocchi)
        names = re.findall(rf"\b({'|'.join(MONTHS_IT)})\b", block, re.IGNORECASE)
        # heuristica: prendi ultimi 12 univoci in ordine di apparizione
        seen = set()
        labels = []
        for n in names:
            mlow = n.lower()
            if mlow not in seen:
                seen.add(mlow)
                labels.append(NORM_MONTH.get(mlow, mlow.title()))
        # anno ignoto → non ottimale, ma proseguiamo
    # Estrai righe F1/F2/F3/Tot ciascuna con una sequenza di numeri
    def grab_line(prefix: str) -> List[int]:
        mm = re.search(rf"^{prefix}\s*[-–]?\s*((?:{NUM_RE}\s*)+)$", block, re.IGNORECASE | re.MULTILINE)
        if not mm:
            return []
        nums = re.findall(NUM_RE, mm.group(1))
        return [norm_int(x) for x in nums]

    f1_vals = grab_line("F1")
    f2_vals = grab_line("F2")
    f3_vals = grab_line("F3")
    tot_vals = grab_line("Tot")

    # Allinea per lunghezza minima comune
    lengths = [len(f1_vals), len(f2_vals), len(f3_vals), len(tot_vals), len(labels)]
    L = min([l for l in lengths if l > 0] + [0])
    if L == 0:
        return None

    f1_vals, f2_vals, f3_vals, tot_vals, labels = f1_vals[:L], f2_vals[:L], f3_vals[:L], tot_vals[:L], labels[:L]
    df = pd.DataFrame({
        "Mese": labels,
        "F1": f1_vals,
        "F2": f2_vals,
        "F3": f3_vals,
        "Totale": tot_vals
    })
    return df

def parse_pdf(file_bytes: bytes) -> Tuple[Optional[pd.DataFrame], str, int, List[str]]:
    """
    Ritorna: (df_grafico, titolo_sheet, costante, note_log)
    """
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        text_pages = [p.extract_text() or "" for p in pdf.pages]
    full_text = "\n".join(text_pages)

    notes = []

    # Tenta parsing Repower
    df = parse_repower(full_text)
    brand = "Repower" if df is not None else None

    # Tenta parsing Enel se Repower fallisce
    if df is None:
        df = parse_enel(full_text)
        if df is not None:
            brand = "Enel"

    if df is None:
        return None, "Non riconosciuto", 1, ["Layout non riconosciuto (né Repower né Enel)."]

    # Ultimi 12 mesi
    df = take_last_12(df)

    # Validazione Totale
    df, val_notes = validate_totals(df)
    notes.extend(val_notes)

    # Costante
    const = detect_constant(full_text)
    if const != 1:
        notes.append(f"Costante di misura rilevata: x{const}")
    else:
        notes.append("Costante di misura: x1 (default)")

    # Titolo foglio (POD o indirizzo se presenti, altrimenti brand)
    sheet = brand
    m = re.search(r"POD\s*([A-Z0-9]{14,})", full_text, re.IGNORECASE)
    if m:
        sheet = f"{brand} {m.group(1)[-8:]}"
    else:
        m2 = re.search(r"(?:punto di prelievo|INDIRIZZO DI FORNITURA)[:\s]+(.+)", full_text, re.IGNORECASE)
        if m2:
            sheet = f"{brand} - {m2.group(1)[:24].strip()}"

    # Aggiunge colonna Tipo e duplica per Fatturati
    grafico = df.copy()
    grafico.insert(1, "Tipo", "Grafico (kWh)")

    fatturati = df.copy()
    fatturati[["F1","F2","F3","Totale"]] = (fatturati[["F1","F2","F3","Totale"]] * const).round().astype(int)
    fatturati.insert(1, "Tipo", "Fatturati (kWh)")

    return (pd.concat([grafico, fatturati], axis=0, ignore_index=True), sheet, const, notes)

# ---------------------------
# UI
# ---------------------------

st.set_page_config(page_title="Estrattore Consumi Bollette (Enel/Repower) → Excel", page_icon="⚡")
st.title("⚡ Estrattore Consumi Bollette (Enel/Repower) → Excel")
st.caption("Carica una o più bollette in PDF. Per ogni file verrà creato un foglio Excel con due tabelle: Grafico (kWh) e Fatturati (kWh).")

uploaded = st.file_uploader("Trascina qui i PDF", type=["pdf"], accept_multiple_files=True)

if uploaded:
    logs = []
    sheets: Dict[str, pd.DataFrame] = {}
    for up in uploaded:
        try:
            df_sheet, sheet_name, const, notes = parse_pdf(up.read())
            if df_sheet is None:
                logs.append(f"❌ {up.name}: {notes[0] if notes else 'Errore parsing'}")
                continue
            # Prepara tabella pivot per comodità di lettura in Excel (due blocchi)
            # Manteniamo una singola tabella con colonna Tipo; in Excel appariranno due blocchi consecutivi.
            sheets[sheet_name[:31]] = df_sheet
            notes_str = "; ".join(notes)
            logs.append(f"✅ {up.name} → foglio: '{sheet_name[:31]}', {notes_str}")
        except Exception as e:
            logs.append(f"❌ {up.name}: errore inatteso - {e}")

    if sheets:
        # Compone Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet, df in sheets.items():
                # Riga 1: nota costante come commento nel log; qui teniamo solo i dati
                # Scriviamo due blocchi separati usando il campo Tipo
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
