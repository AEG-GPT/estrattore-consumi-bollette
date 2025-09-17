def parse_enel(text: str) -> Optional[pd.DataFrame]:
    """
    Parser robusto per Enel:
    - isola il blocco vicino al titolo
    - cattura le serie tra header e header successivo (F1→F2, F2→F3, F3→Tot, Tot→sezione successiva)
    - **filtra gli anni 20xx** che talvolta finiscono dentro i blocchi numerici
    - taglia l’eventuale riga dei totali di colonna
    - allinea alle ultime 12 osservazioni
    """
    # 1) Trova il titolo e isola un blocco ragionevole
    title_pat = re.compile(r"Consumi\s+(?:in\s+kWh\s+)?degli?\s+ultimi\s+\d{1,2}\s+mesi", re.IGNORECASE)
    m_title = title_pat.search(text)
    block = text[m_title.start():m_title.start()+4000] if m_title else text

    # 2) Etichette mesi (best effort)
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

    # 3) utilità: prendi numeri tra header corrente e prossimo header, **filtrando anni**
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
        body = m.group("body")
        # 🔴 filtra anni 20xx che entrano nel blocco
        body = re.sub(r"\b20\d{2}\b", " ", body)
        nums = re.findall(NUM_RE, body)
        return [norm_int(x) for x in nums]

    # 4) header e stop-words
    H1 = ["F1","F 1","Fascia 1","FASCIA 1"]
    H2 = ["F2","F 2","Fascia 2","FASCIA 2"]
    H3 = ["F3","F 3","Fascia 3","FASCIA 3"]
    HT = ["Tot","Totale","TOTALE","TOT"]
    STOP_AFTER_TOT = ["Potenza","kW max","GLOSSARIO","Legenda","Dettaglio","Energia reattiva","REATTIVA","POTENZA"]

    # 5) estrai serie
    f1_vals = series_between(H1, H2, block)
    f2_vals = series_between(H2, H3, block)
    f3_vals = series_between(H3, HT, block)
    tot_vals = series_between(HT, STOP_AFTER_TOT, block)

    # fallback sul testo intero, se mancano
    if not (f1_vals and f2_vals and f3_vals and tot_vals):
        f1_vals = f1_vals or series_between(H1, H2, text)
        f2_vals = f2_vals or series_between(H2, H3, text)
        f3_vals = f3_vals or series_between(H3, HT, text)
        tot_vals = tot_vals or series_between(HT, STOP_AFTER_TOT, text)

    # 6) togli l'eventuale riga dei totali di colonna in coda
    def drop_column_total(vals: List[int]) -> List[int]:
        if len(vals) >= 13:
            s = sum(vals[:-1])
            last = vals[-1]
            if abs(last - s) <= max(10, int(0.02 * s)):  # 2% o 10 kWh
                return vals[:-1]
        return vals

    f1_vals = drop_column_total(f1_vals)
    f2_vals = drop_column_total(f2_vals)
    f3_vals = drop_column_total(f3_vals)
    tot_vals = drop_column_total(tot_vals)

    # 7) allineamento (ultimi 12)
    lengths = [len(f1_vals), len(f2_vals), len(f3_vals), len(tot_vals)]
    L = min([l for l in lengths if l > 0], default=0)
    if L == 0:
        return None
    f1_vals, f2_vals, f3_vals, tot_vals = f1_vals[-L:], f2_vals[-L:], f3_vals[-L:], tot_vals[-L:]
    if L > 12:
        f1_vals, f2_vals, f3_vals, tot_vals = f1_vals[-12:], f2_vals[-12:], f3_vals[-12:], tot_vals[-12:]
        L = 12

    # 8) etichette
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
