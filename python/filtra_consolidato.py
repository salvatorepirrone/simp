#!/usr/bin/env python3
"""
filtra_consolidato.py  (v2)
Uso: python filtra_consolidato.py <file_xlsx> <file_prodotti_csv> <file_output_csv> <file_output_xlsx>

Legge il file consolidato mensile (xlsx), filtra le righe DC Pensioni
con Totale Pervenuti > 0 o Totale Definiti > 0, scrive:
  - un CSV ridotto (file_output_csv)
  - lo stesso contenuto in formato XLSX (file_output_xlsx)
Usato come pre-processore dalla macro VBA ImportaDatiMensili.

Novita' v2:
  - aggiunto il 4o argomento file_output_xlsx: lo stesso output filtrato
    viene salvato anche come workbook .xlsx (oltre al CSV), cosi' la
    macro VBA puo' collocare i due file in due cartelle differenti
    senza dover riaprire/convertire nulla via Excel.
  - le cartelle di destinazione (sia del CSV sia dell'XLSX) vengono
    create automaticamente se non esistono.
  - l'XLSX usa la stessa formattazione "italiana" del CSV (numeri come
    testo con virgola decimale, es. 24,5): le celle numeriche nell'XLSX
    sono quindi testo, non numeri nativi, per coerenza col CSV.
"""
import sys
import os
import csv

def to_float(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    # Formato italiano (punto=migliaia, virgola=decimale) solo se c'e' la virgola
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
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
        if v != v:          # NaN
            return ''
        if v == int(v):
            return str(int(v))          # 24.0 -> "24"
        return str(v).replace('.', ',') # 24.5 -> "24,5"
    # Decimal o altri tipi numerici: converti via float
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
    """Trova l'indice di colonna cercando tra i nomi alternativi (case-insensitive)."""
    h = [str(c).strip().lower() if c is not None else '' for c in headers]
    for name in names:
        try:
            return h.index(name.lower())
        except ValueError:
            pass
    return None

def main():
    if len(sys.argv) != 5:
        print("USO: python filtra_consolidato.py <xlsx> <prodotti_csv> <output_csv> <output_xlsx>")
        sys.exit(1)

    xlsx_path    = sys.argv[1]
    prodotti_csv = sys.argv[2]
    output_csv   = sys.argv[3]
    output_xlsx  = sys.argv[4]

    # 1. Leggi prodotti DC Pensioni dal CSV dei prodotti
    dc_prodotti = set()
    with open(prodotti_csv, newline='', encoding='cp1252') as f:
        reader = csv.reader(f, delimiter=';')
        next(reader, None)  # salta header
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
    # Cerca il foglio per nome (case-insensitive), fallback al primo foglio
    sheet_names = [s.lower() for s in wb.sheet_names]
    if 'consolidato' in sheet_names:
        sheet = wb.get_sheet_by_name(wb.sheet_names[sheet_names.index('consolidato')])
    else:
        sheet = wb.get_sheet_by_index(0)
    rows = sheet.to_python(skip_empty_area=False)

    # 3. Trova riga A1 (mese/anno) e riga header
    #    L'header e' la prima riga che contiene "area" in qualsiasi colonna
    cell_a1 = ''
    hdr_idx  = None
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

    # 4. Rileva indici colonne dalla riga header (robusto a variazioni di layout)
    IDX_PROD = find_col(header_row, 'prodotto outcome', 'prodotto')
    IDX_PERV = find_col(header_row, 'totale pervenuti', 'tot pervenuti')
    IDX_DEF  = find_col(header_row, 'totale definiti', 'tot definiti')

    if IDX_PROD is None or IDX_PERV is None or IDX_DEF is None:
        print(f"ERRORE: colonne obbligatorie non trovate. Header: {header_row}")
        sys.exit(5)

    # 5. Filtra le righe (tenute in memoria per scrivere sia CSV sia XLSX)
    n_tot  = 0
    n_out  = 0
    n_sca  = 0
    filtered_rows = []

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

        filtered_rows.append(row)
        n_out += 1

    # 6. Crea le cartelle di destinazione se non esistono
    for out_path in (output_csv, output_xlsx):
        out_dir = os.path.dirname(out_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

    # 7. Scrivi il CSV ridotto (formato Excel italiano)
    with open(output_csv, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow([cell_a1])
        writer.writerow(header_row)
        for row in filtered_rows:
            writer.writerow([format_val(c) for c in row])

    # 8. Scrivi lo stesso output in formato XLSX
    try:
        from openpyxl import Workbook
    except ImportError:
        print("ERRORE: openpyxl non installato. Eseguire: pip install openpyxl")
        sys.exit(6)

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Consolidato"
    ws_out.append([cell_a1])
    ws_out.append(header_row)
    for row in filtered_rows:
        ws_out.append([format_val(c) for c in row])
    wb_out.save(output_xlsx)

    print(f"OK|{n_tot}|{n_out}|{n_sca}|{cell_a1}")

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        import traceback
        print("ERRORE|" + str(e))
        traceback.print_exc()
