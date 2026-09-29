#!/usr/bin/env python3
"""
filtra_dcpensioni.py
Uso: python filtra_dcpensioni.py <file_xlsx> <file_prodotti_csv> <output_dir>

Usato dalla macro VBA "Filtra". A differenza di filtra_consolidato.py
(usato da ImportaDatiMensili), qui Python fa TUTTO il lavoro pesante:
  - legge il file di produzione con calamine
  - filtra le righe: Prodotto Outcome -> DC Pensioni (foglio "prodotti",
    colonna B) E (Totale Pervenuti > 0 OPPURE Totale Definiti > 0)
  - calcola mese/anno dalla cella A1
  - scrive DIRETTAMENTE l'xlsx finale, gia' nel layout colonne del
    Consolidato (Anno, Mese, Area, Prodotto, Codice, Descrizione,
    Regione, CodSede, Sede, GiacIniz, P0..P23, GiacFin, Omog, GiacGG),
    con valori NUMERICI VERI (non testo) per le colonne numeriche,
    necessari alle formule XLOOKUP/coefficienti che VBA aggiungera' dopo
  - determina da solo il nome del file di output:
    SIMP_DCPensioni_<AAAA>_<MM>.xlsx dentro <output_dir>

VBA, dopo aver ricevuto OK, apre semplicemente il file scritto da questo
script, copia i fogli "territorio" e "prodotti" da ThisWorkbook, scrive
le formule delle colonne AM:AR e salva. In questo modo tutto il lavoro
di lettura/filtro/scrittura dati (la parte lenta su file grandi) resta
a Python/calamine/openpyxl, molto piu' veloce del ciclo riga-per-riga
in VBA usato in precedenza.

Output su stdout: OK|n_tot|n_out|n_sca|cellA1|output_path
oppure, se non ci sono righe: OK|n_tot|0|n_sca|cellA1|
"""
import sys
import os

# --- Colonne SORGENTE nel file di produzione (1-based, stesse costanti
#     S_* usate in VBA/filtra_consolidato.py) ---
S_AREA     = 1
S_PRODOTTO = 2
S_CODICE   = 3
S_DESCRIZ  = 4
S_REGIONE  = 5
S_CODSEDE  = 6
S_SEDE     = 7
S_GIACINIZ = 8
S_P0       = 9
S_TOTPERV  = 20
S_TOTDEF   = 32
S_GIACFIN  = 33
S_OMOG     = 34
S_GIACGG   = 35
N_DIFF     = 24  # numero di colonne P0..P23


def to_float(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0


def find_col(headers, *names):
    h = [str(c).strip().lower() if c is not None else '' for c in headers]
    for name in names:
        try:
            return h.index(name.lower())
        except ValueError:
            pass
    return None


def parse_mese_anno(testo):
    """Replica la logica di ParseMeseAnno in VBA: cerca l'ultimo '/',
    poi lo spazio immediatamente precedente, e legge mese/anno in mezzo."""
    if not testo:
        return None
    p = testo.rfind('/')
    if p == -1:
        return None
    sp = testo.rfind(' ', 0, p)
    if sp == -1:
        return None
    s_m = testo[sp + 1:p].strip()
    s_a = testo[p + 1:].strip()
    if not (s_m.isdigit() and s_a.isdigit()):
        return None
    mv = int(s_m)
    av = int(s_a)
    if mv < 1 or mv > 12 or av < 2000 or av > 2100:
        return None
    return mv, av


def main():
    if len(sys.argv) != 4:
        print("USO: python filtra_dcpensioni.py <xlsx> <prodotti_csv> <output_dir>")
        sys.exit(1)

    xlsx_path    = sys.argv[1]
    prodotti_csv = sys.argv[2]
    output_dir   = sys.argv[3]

    # 1. Leggi prodotti DC Pensioni dal CSV dei prodotti (foglio "prodotti", colonna B)
    import csv
    dc_prodotti = set()
    with open(prodotti_csv, newline='', encoding='cp1252') as f:
        reader = csv.reader(f, delimiter=';')
        next(reader, None)
        for row in reader:
            if len(row) >= 2 and 'DC Pensioni' in row[1]:
                dc_prodotti.add(row[0].strip())

    if not dc_prodotti:
        print("ERRORE: nessun prodotto DC Pensioni trovato in " + prodotti_csv)
        sys.exit(2)

    # 2. Leggi xlsx con calamine (veloce)
    try:
        from python_calamine import CalamineWorkbook
    except ImportError:
        print("ERRORE: python-calamine non installato. Eseguire: pip install python-calamine")
        sys.exit(3)

    wb = CalamineWorkbook.from_path(xlsx_path)
    sheet_names = [s.lower() for s in wb.sheet_names]
    if 'consolidato' in sheet_names:
        sheet = wb.get_sheet_by_name(wb.sheet_names[sheet_names.index('consolidato')])
    else:
        sheet = wb.get_sheet_by_index(0)
    rows = sheet.to_python(skip_empty_area=False)

    # 3. Trova riga A1 (mese/anno) e riga header
    cell_a1 = ''
    hdr_idx = None
    for i, row in enumerate(rows):
        if i == 0:
            cell_a1 = str(row[0]).strip() if row and row[0] is not None else ''
        if not row:
            continue
        row_vals = [str(c).strip().lower() if c is not None else '' for c in row]
        if 'area' in row_vals:
            hdr_idx = i
            break

    if hdr_idx is None:
        print("ERRORE: intestazione non trovata nel file xlsx")
        sys.exit(4)

    header_row = [str(c).strip() if c is not None else '' for c in rows[hdr_idx]]

    IDX_PROD = find_col(header_row, 'prodotto outcome', 'prodotto')
    IDX_PERV = find_col(header_row, 'totale pervenuti', 'tot pervenuti')
    IDX_DEF  = find_col(header_row, 'totale definiti', 'tot definiti')

    if IDX_PROD is None or IDX_PERV is None or IDX_DEF is None:
        print(f"ERRORE: colonne obbligatorie non trovate. Header: {header_row}")
        sys.exit(5)

    # 4. Mese/Anno da cella A1
    ma = parse_mese_anno(cell_a1)
    if ma is None:
        print(f"ERRORE: impossibile leggere Mese/Anno dalla cella A1: '{cell_a1}'")
        sys.exit(6)
    mese, anno = ma

    # 5. Filtra + costruisci righe gia' nel layout Consolidato (D_*),
    #    con valori assoluti (NESSUN differenziale) e tipi nativi
    n_tot = 0
    n_out = 0
    n_sca = 0
    out_rows = []

    for row in rows[hdr_idx + 1:]:
        if not row or row[0] is None:
            continue
        n_tot += 1

        prod = str(row[IDX_PROD]).strip() if IDX_PROD < len(row) and row[IDX_PROD] else ''
        if prod not in dc_prodotti:
            n_sca += 1
            continue

        tp = to_float(row[IDX_PERV]) if IDX_PERV < len(row) else 0.0
        td = to_float(row[IDX_DEF])  if IDX_DEF  < len(row) else 0.0
        if tp == 0 and td == 0:
            n_sca += 1
            continue

        def col(i):
            return row[i - 1] if i - 1 < len(row) else None

        area      = str(col(S_AREA)).strip()      if col(S_AREA)      is not None else ''
        prodotto  = str(col(S_PRODOTTO)).strip()   if col(S_PRODOTTO)  is not None else ''
        codice    = str(col(S_CODICE)).strip()     if col(S_CODICE)    is not None else ''
        descriz   = str(col(S_DESCRIZ)).strip()    if col(S_DESCRIZ)   is not None else ''
        regione   = str(col(S_REGIONE)).strip()    if col(S_REGIONE)   is not None else ''
        codsede   = str(col(S_CODSEDE)).strip()    if col(S_CODSEDE)   is not None else ''
        sede      = str(col(S_SEDE)).strip()       if col(S_SEDE)      is not None else ''
        giaciniz  = to_float(col(S_GIACINIZ))
        p_vals    = [to_float(col(S_P0 + k)) for k in range(N_DIFF)]
        giacfin   = to_float(col(S_GIACFIN))
        omog      = to_float(col(S_OMOG))
        giacgg    = to_float(col(S_GIACGG))

        out_rows.append(
            [anno, mese, area, prodotto, codice, descriz, regione, codsede, sede, giaciniz]
            + p_vals + [giacfin, omog, giacgg]
        )
        n_out += 1

    if n_out == 0:
        print(f"OK|{n_tot}|0|{n_sca}|{cell_a1}|")
        return

    # 6. Scrivi l'xlsx finale (dati soltanto, senza formule: le aggiunge VBA)
    try:
        from openpyxl import Workbook
    except ImportError:
        print("ERRORE: openpyxl non installato. Eseguire: pip install openpyxl")
        sys.exit(7)

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"SIMP_DCPensioni_{anno:04d}_{mese:02d}.xlsx")

    wbo = Workbook()
    wso = wbo.active
    wso.title = "DCPensioni"

    header = ["Anno", "Mese", "Area", "Prodotto Outcome", "Codice", "Descrizione",
              "Regione", "Codice Sede", "Sede", "Giacenza Iniziale"]
    header += ["P0", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10",
               "Totale Pervenuti",
               "D0", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10",
               "Totale Definiti"]
    header += ["Giacenza Finale", "Totale Omogeneizzato", "Giacenza Gg"]
    wso.append(header)

    for r in out_rows:
        wso.append(r)

    wbo.save(output_path)

    print(f"OK|{n_tot}|{n_out}|{n_sca}|{cell_a1}|{output_path}")


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        import traceback
        print("ERRORE|" + str(e))
        traceback.print_exc()
