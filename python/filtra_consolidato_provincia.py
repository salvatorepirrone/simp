#!/usr/bin/env python3
"""
filtra_consolidato_provincia.py
Uso: python filtra_consolidato_provincia.py <file_xlsx> <file_prodotti_csv> <file_territorio_csv> <file_output_csv> <file_output_xlsx>

Variante di filtra_consolidato.py (usato da ImportaDatiMensili) che, oltre a
filtrare le righe DC Pensioni (Prodotto Outcome -> DC Pensioni E (Totale
Pervenuti > 0 OPPURE Totale Definiti > 0)), AGGREGA i risultati a livello di
Descrizione Provincia invece di lasciarli a livello di singola Sede:

  - le colonne "Codice Sede"/"Sede" del file sorgente vengono sostituite da
    UNA colonna "Descrizione Provincia", ottenuta con lo stesso aggancio
    gia' usato dalla formula "Territorio" nelle altre macro (Codice Sede
    cercato nel foglio "territorio", colonna D, con restituzione dalla
    colonna C)
  - tutte le righe che condividono la stessa combinazione
    Area/Prodotto/Codice/Descrizione/Regione/Descrizione Provincia vengono
    sommate (Giacenza Iniziale, P0..P23, Giacenza Finale, Totale
    Omogeneizzato, Giacenza Gg)

Questo riduce drasticamente il numero di righe scritte ogni mese nel
Consolidato (una riga per provincia anziche' una per sede), che è la causa
principale del peso eccessivo del file.

L'aggancio Codice Sede -> Descrizione Provincia viene fatto qui in Python
(non con una formula XLOOKUP in Excel) perche' deve avvenire PRIMA di
aggregare, per poter decidere quali righe finiscono nello stesso gruppo.
Il CSV del foglio "territorio" (tutte le colonne, comprese C=Descrizione
Provincia e D=Codice Sede) viene esportato dalla macro VBA con la stessa
routine gia' usata da FiltraProvincia (EsportaFoglioCsv).

Layout di output (CSV e XLSX), SENZA Anno/Mese (aggiunti da VBA in fase di
scrittura nel Consolidato), identico a quello del file prodotto da
FiltraProvincia/filtra_dcpensioni_provincia.py una volta tolti Anno/Mese:
  Area, Prodotto Outcome, Codice, Descrizione, Regione, Descrizione Provincia,
  Giacenza Iniziale, P0..P10, Totale Pervenuti, D0..D10, Totale Definiti,
  Giacenza Finale, Totale Omogeneizzato, Giacenza Gg
(7 colonne descrittive + 1 Giacenza Iniziale + 24 valori P/D + 3 colonne
finali = 35 colonne)

Output su stdout: OK|n_tot|n_out|n_sca|n_gruppi|cellA1
"""
import sys
import os
import csv

# --- Colonne SORGENTE nel file di produzione (1-based, stesse costanti
#     S_* usate in VBA/filtra_consolidato.py/filtra_dcpensioni.py) ---
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

# --- Colonne del foglio "territorio" esportato da EsportaFoglioCsv
#     (dump completo colonne A,B,C,... ; 0-based sull'elenco letto da csv) ---
T_PROVINCIA = 2  # colonna C: Descrizione Provincia
T_CODSEDE   = 3  # colonna D: Codice Sede


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
    except Exception:
        return 0.0


def format_val(v):
    """Scrive i valori nel CSV in formato compatibile con Excel italiano."""
    if v is None:
        return ''
    if isinstance(v, bool):
        return str(int(v))
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        if v != v:  # NaN
            return ''
        if v == int(v):
            return str(int(v))
        return str(v).replace('.', ',')
    try:
        f = float(v)
        if f != f:
            return ''
        if f == int(f):
            return str(int(f))
        return str(f).replace('.', ',')
    except (ValueError, TypeError):
        pass
    return str(v)


def find_col(headers, *names):
    h = [str(c).strip().lower() if c is not None else '' for c in headers]
    for name in names:
        try:
            return h.index(name.lower())
        except ValueError:
            pass
    return None


def norm_codsede(v):
    """Replica TEXT(v,"000000"): stringa numerica su 6 cifre con zeri iniziali."""
    if v is None:
        return ''
    s = str(v).strip()
    if s == '':
        return ''
    try:
        n = int(float(s.replace(',', '.'))) if ('.' in s or ',' in s) else int(s)
        return f"{n:06d}"
    except Exception:
        return s.zfill(6)


def carica_territorio(territorio_csv):
    """Legge il CSV del foglio 'territorio' (dump completo, header in riga 1)
    e restituisce un dict {CodSede normalizzato (6 cifre) -> Descrizione Provincia}."""
    # Il CSV e' scritto da VBA con "Print #" (EsportaFoglioCsv), quindi nella
    # codifica ANSI di sistema (cp1252 su Windows in italiano) - stessa
    # codifica gia' usata per leggere il CSV prodotti in filtra_dcpensioni.py.
    mapping = {}
    with open(territorio_csv, newline='', encoding='cp1252') as f:
        reader = csv.reader(f, delimiter=';')
        next(reader, None)  # salta header
        for row in reader:
            if len(row) <= max(T_PROVINCIA, T_CODSEDE):
                continue
            cod = norm_codsede(row[T_CODSEDE])
            prov = row[T_PROVINCIA].strip()
            if cod and prov:
                mapping[cod] = prov
    return mapping


def main():
    if len(sys.argv) != 6:
        print("USO: python filtra_consolidato_provincia.py <xlsx> <prodotti_csv> <territorio_csv> <output_csv> <output_xlsx>")
        sys.exit(1)

    xlsx_path      = sys.argv[1]
    prodotti_csv   = sys.argv[2]
    territorio_csv = sys.argv[3]
    output_csv     = sys.argv[4]
    output_xlsx    = sys.argv[5]

    # 1. Prodotti DC Pensioni
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

    # 2. Mappa Codice Sede -> Descrizione Provincia
    prov_by_codsede = carica_territorio(territorio_csv)
    if not prov_by_codsede:
        print("ERRORE: nessuna riga letta dal CSV territorio " + territorio_csv)
        sys.exit(3)

    # 3. Leggi xlsx con calamine
    try:
        from python_calamine import CalamineWorkbook
    except ImportError:
        print("ERRORE: python-calamine non installato. Eseguire: pip install python-calamine")
        sys.exit(4)

    wb = CalamineWorkbook.from_path(xlsx_path)
    sheet_names = [s.lower() for s in wb.sheet_names]
    if 'consolidato' in sheet_names:
        sheet = wb.get_sheet_by_name(wb.sheet_names[sheet_names.index('consolidato')])
    else:
        sheet = wb.get_sheet_by_index(0)
    rows = sheet.to_python(skip_empty_area=False)

    # 4. Riga A1 (mese/anno) e riga header
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
        sys.exit(5)

    header_row = [str(c).strip() if c is not None else '' for c in rows[hdr_idx]]

    IDX_PROD = find_col(header_row, 'prodotto outcome', 'prodotto')
    IDX_PERV = find_col(header_row, 'totale pervenuti', 'tot pervenuti')
    IDX_DEF  = find_col(header_row, 'totale definiti', 'tot definiti')

    if IDX_PROD is None or IDX_PERV is None or IDX_DEF is None:
        print(f"ERRORE: colonne obbligatorie non trovate. Header: {header_row}")
        sys.exit(6)

    # 5. Filtra e aggrega per (Area,Prodotto,Codice,Descrizione,Regione,Provincia)
    n_tot = 0
    n_out = 0
    n_sca = 0
    n_senza_provincia = 0
    gruppi = {}   # chiave -> [descrittivi..., somme...]
    ordine = []   # ordine di prima comparsa, per output deterministico

    for row in rows[hdr_idx + 1:]:
        if not row or row[0] is None:
            continue
        n_tot += 1

        def col(i):
            return row[i - 1] if i - 1 < len(row) else None

        prod = str(col(S_PRODOTTO)).strip() if col(S_PRODOTTO) is not None else ''
        if prod not in dc_prodotti:
            n_sca += 1
            continue

        tp = to_float(col(S_TOTPERV))
        td = to_float(col(S_TOTDEF))
        if tp == 0 and td == 0:
            n_sca += 1
            continue

        area     = str(col(S_AREA)).strip()     if col(S_AREA)     is not None else ''
        prodotto = str(col(S_PRODOTTO)).strip()  if col(S_PRODOTTO) is not None else ''
        codice   = str(col(S_CODICE)).strip()    if col(S_CODICE)   is not None else ''
        descriz  = str(col(S_DESCRIZ)).strip()   if col(S_DESCRIZ)  is not None else ''
        regione  = str(col(S_REGIONE)).strip()   if col(S_REGIONE)  is not None else ''

        cod_sede = norm_codsede(col(S_CODSEDE))
        provincia = prov_by_codsede.get(cod_sede, '')
        if not provincia:
            n_senza_provincia += 1
            provincia = '(Provincia sconosciuta)'

        chiave = (area, prodotto, codice, descriz, regione, provincia)

        giaciniz = to_float(col(S_GIACINIZ))
        p_vals   = [to_float(col(S_P0 + k)) for k in range(N_DIFF)]
        giacfin  = to_float(col(S_GIACFIN))
        omog     = to_float(col(S_OMOG))
        giacgg   = to_float(col(S_GIACGG))

        if chiave not in gruppi:
            gruppi[chiave] = [giaciniz] + p_vals + [giacfin, omog, giacgg]
            ordine.append(chiave)
        else:
            acc = gruppi[chiave]
            acc[0] += giaciniz
            for k in range(N_DIFF):
                acc[1 + k] += p_vals[k]
            acc[1 + N_DIFF]     += giacfin
            acc[1 + N_DIFF + 1] += omog
            acc[1 + N_DIFF + 2] += giacgg

        n_out += 1

    n_gruppi = len(ordine)

    if n_senza_provincia > 0:
        sys.stderr.write(
            f"ATTENZIONE: {n_senza_provincia} righe con Codice Sede non trovato "
            f"nel foglio territorio (raggruppate sotto '(Provincia sconosciuta)')\n")

    header_out = ["Area", "Prodotto Outcome", "Codice", "Descrizione", "Regione",
                  "Descrizione Provincia", "Giacenza Iniziale"]
    header_out += ["P0", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10",
                    "Totale Pervenuti",
                    "D0", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10",
                    "Totale Definiti"]
    header_out += ["Giacenza Finale", "Totale Omogeneizzato", "Giacenza Gg"]

    out_rows = []
    for chiave in ordine:
        area, prodotto, codice, descriz, regione, provincia = chiave
        out_rows.append([area, prodotto, codice, descriz, regione, provincia] + gruppi[chiave])

    # 6. Crea le cartelle di destinazione se non esistono
    for out_path in (output_csv, output_xlsx):
        out_dir = os.path.dirname(out_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

    # 7. Scrivi il CSV aggregato (formato Excel italiano)
    with open(output_csv, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow([cell_a1])
        writer.writerow(header_out)
        for row in out_rows:
            writer.writerow([format_val(c) for c in row])

    # 8. Scrivi lo stesso output in formato XLSX
    try:
        from openpyxl import Workbook
    except ImportError:
        print("ERRORE: openpyxl non installato. Eseguire: pip install openpyxl")
        sys.exit(7)

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "ConsolidatoProvincia"
    ws_out.append([cell_a1])
    ws_out.append(header_out)
    for row in out_rows:
        ws_out.append([format_val(c) for c in row])
    wb_out.save(output_xlsx)

    print(f"OK|{n_tot}|{n_out}|{n_sca}|{n_gruppi}|{cell_a1}")


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        import traceback
        print("ERRORE|" + str(e))
        traceback.print_exc()
