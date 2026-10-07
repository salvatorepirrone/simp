Attribute VB_Name = "Modulo1"
Option Explicit

'==============================================================================
' Importazione periodica di Memo (file "Progetti") e Demand (file "data")
'
' - AGGIORNAMENTO: le righe dei file selezionati (anche piu' di uno) sostituiscono
'   quelle con lo stesso codice o vengono aggiunte; le righe del foglio che non
'   compaiono nei file importati NON vengono cancellate. Se lo stesso codice
'   compare piu' volte, vale l'ultima occorrenza.
' - Vengono scritti SOLO i record per i quali esiste la trascodifica dell'area
'   nel foglio "Aree_trascodifica" (colonna A = archivio "Memo" / "Demand",
'   colonna B = valore da cercare, colonna C = valore trascodificato).
' - Memo   : AC = area richiedente omogeneizzata (da "Aree richiedenti e coinvolte")
' - Demand : U  = area richiedente omogeneizzata (da "Area Richiedente")
'            T  = codici Memo (col. A foglio Memo) che contengono il codice
'                 Demand in col. B ("Codice Richiesta"), separati da ";"
'==============================================================================

Private Const COL_DEMAND_MEMO As Long = 20   'T
Private Const COL_DEMAND_AREA As Long = 21   'U
Private Const COL_MEMO_AREA As Long = 29     'AC
Private Const MAX_RIGHE_INTESTAZIONE As Long = 15

Public Sub Importa_Memo_Da_File()
    'Memo: aggiorna/aggiunge le righe dei file importati e CONSERVA le altre
    ImportaArchivio "Memo", "Codice", "Aree richiedenti e coinvolte", COL_MEMO_AREA, False
End Sub

Public Sub Importa_Demand_Da_File()
    'Demand: aggiorna/aggiunge le righe dei file importati e CONSERVA le altre
    '(cosi' si possono importare i file data uno alla volta)
    ImportaArchivio "Demand", "Codice richiesta", "Area Richiedente", COL_DEMAND_AREA, False
End Sub

'------------------------------------------------------------------------------
' Routine comune di importazione
'------------------------------------------------------------------------------
Private Sub ImportaArchivio(ByVal foglio As String, ByVal hdrCodice As String, _
                            ByVal hdrAreaSrc As String, ByVal colAreaDst As Long, _
                            ByVal sostituisciTutto As Boolean)
    Dim wsDst As Worksheet, wsReport As Worksheet
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim fd As FileDialog
    Dim nomiFile As String
    Dim iFile As Long

    Dim dictTrasc As Object, dictOld As Object, dictNew As Object, dictScarti As Object, dictRun As Object
    Dim dstHeaders As Variant, srcHeaders As Variant, dati As Variant
    Dim srcToDst() As Long
    Dim vec() As Variant
    Dim itemsNew As Variant, outArr() As Variant
    Dim tipoData() As Long

    Dim hdrRowDst As Long, lastColDst As Long, colCodDst As Long, lastUsed As Long
    Dim hdrRowSrc As Long, lastRowSrc As Long, lastColSrc As Long
    Dim colCodSrc As Long, colAreaSrc As Long
    Dim c As Long, r As Long, k As Long, n As Long
    Dim codice As String, areaSrc As String, areaDst As String
    Dim v As Variant
    Dim nAgg As Long, nAdd As Long, nScart As Long
    Dim ok As Boolean
    Dim wsScartiEsist As Worksheet

    On Error GoTo CleanFail

    Set wsDst = ThisWorkbook.Worksheets(foglio)
    Set wsReport = ThisWorkbook.Worksheets("Report")

    '--- intestazione di destinazione
    hdrRowDst = TrovaRigaIntestazione(wsDst, hdrCodice, MAX_RIGHE_INTESTAZIONE)
    If hdrRowDst = 0 Then
        MsgBox "Non trovo l'intestazione '" & hdrCodice & "' nel foglio '" & foglio & "'.", vbCritical
        Exit Sub
    End If
    lastColDst = wsDst.Cells(hdrRowDst, wsDst.Columns.Count).End(xlToLeft).Column
    If lastColDst < colAreaDst Then lastColDst = colAreaDst
    If foglio = "Demand" And lastColDst < COL_DEMAND_MEMO Then lastColDst = COL_DEMAND_MEMO

    'intestazioni delle colonne calcolate (se mancanti)
    ImpostaIntestazione wsDst, hdrRowDst, colAreaDst, "Area richiedente omogeneizzata"
    If foglio = "Demand" Then ImpostaIntestazione wsDst, hdrRowDst, COL_DEMAND_MEMO, "memo collegati"

    dstHeaders = wsDst.Range(wsDst.Cells(hdrRowDst, 1), wsDst.Cells(hdrRowDst, lastColDst)).Value
    colCodDst = FindHeaderCol(dstHeaders, hdrCodice)

    '--- tabella di trascodifica (per l'archivio corrente)
    Set dictTrasc = CaricaTrascodifica(foglio)
    If dictTrasc.Count = 0 Then
        MsgBox "Nel foglio 'Aree_trascodifica' non ci sono trascodifiche per l'archivio '" & foglio & "'.", vbCritical
        Exit Sub
    End If

    '--- scelta file (anche multipla)
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleziona il/i file da importare nel foglio " & foglio
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Sub
    End With

    '--- codici attualmente presenti (solo per le statistiche del report)
    Set dictOld = CreateObject("Scripting.Dictionary")
    dictOld.CompareMode = vbTextCompare
    lastUsed = wsDst.Cells(wsDst.Rows.Count, colCodDst).End(xlUp).Row
    If lastUsed > hdrRowDst Then
        dati = wsDst.Range(wsDst.Cells(hdrRowDst + 1, colCodDst), wsDst.Cells(lastUsed, colCodDst)).Value
        For r = 1 To UBound(dati, 1)
            codice = Trim$(CStr(dati(r, 1)))
            If Len(codice) > 0 Then dictOld(codice) = True
        Next r
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set dictNew = CreateObject("Scripting.Dictionary")
    dictNew.CompareMode = vbTextCompare
    Set dictScarti = CreateObject("Scripting.Dictionary")
    dictScarti.CompareMode = vbTextCompare
    ReDim tipoData(1 To lastColDst)
    Set dictRun = CreateObject("Scripting.Dictionary")   'codici scritti da questa esecuzione
    dictRun.CompareMode = vbTextCompare

    'modalita' aggiornamento: si parte dalle righe gia' presenti (scritte e scartate)
    If Not sostituisciTutto Then
        CaricaRighe wsDst, hdrRowDst + 1, colCodDst, lastColDst, dictNew, tipoData
        Set wsScartiEsist = Nothing
        On Error Resume Next
        Set wsScartiEsist = ThisWorkbook.Worksheets(LCase$(foglio) & "_scarti")
        On Error GoTo CleanFail
        If Not wsScartiEsist Is Nothing Then
            CaricaRighe wsScartiEsist, 2, colCodDst, lastColDst, dictScarti, tipoData
        End If
    End If

    '--- legge tutti i file in memoria (la destinazione non viene toccata finche'
    '    tutti i file non sono stati letti correttamente)
    For iFile = 1 To fd.SelectedItems.Count
        If Len(nomiFile) > 0 Then nomiFile = nomiFile & ", "
        nomiFile = nomiFile & Dir(fd.SelectedItems(iFile))

        Set wbSrc = Workbooks.Open(fileName:=fd.SelectedItems(iFile), ReadOnly:=True, UpdateLinks:=0)

        Set wsSrc = Nothing
        On Error Resume Next
        Set wsSrc = wbSrc.Worksheets("Export")
        On Error GoTo CleanFail
        If wsSrc Is Nothing Then Set wsSrc = wbSrc.Worksheets(1)

        'i file di export possono avere righe iniziali di filtro / righe vuote
        hdrRowSrc = TrovaRigaIntestazione(wsSrc, hdrCodice, MAX_RIGHE_INTESTAZIONE)
        If hdrRowSrc = 0 Then
            MsgBox "Nel file '" & wbSrc.Name & "' non trovo l'intestazione '" & hdrCodice & "'." & vbCrLf & _
                   "Nessuna modifica effettuata.", vbCritical
            GoTo CleanExit
        End If

        lastColSrc = wsSrc.Cells(hdrRowSrc, wsSrc.Columns.Count).End(xlToLeft).Column
        srcHeaders = wsSrc.Range(wsSrc.Cells(hdrRowSrc, 1), wsSrc.Cells(hdrRowSrc, lastColSrc)).Value
        colCodSrc = FindHeaderCol(srcHeaders, hdrCodice)
        colAreaSrc = FindHeaderCol(srcHeaders, hdrAreaSrc)
        If colAreaSrc = 0 Then
            MsgBox "Nel file '" & wbSrc.Name & "' non trovo l'intestazione '" & hdrAreaSrc & "'." & vbCrLf & _
                   "Nessuna modifica effettuata.", vbCritical
            GoTo CleanExit
        End If

        'mappa colonne sorgente -> destinazione tramite intestazione
        '(le colonne calcolate non vengono mai sovrascritte dal sorgente)
        ReDim srcToDst(1 To lastColSrc)
        For c = 1 To lastColSrc
            srcToDst(c) = 0
            If Len(Trim$(CStr(srcHeaders(1, c)))) > 0 Then
                srcToDst(c) = FindHeaderCol(dstHeaders, CStr(srcHeaders(1, c)))
                If srcToDst(c) = colAreaDst Then srcToDst(c) = 0
                If foglio = "Demand" And srcToDst(c) = COL_DEMAND_MEMO Then srcToDst(c) = 0
            End If
        Next c

        lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, colCodSrc).End(xlUp).Row
        If lastRowSrc > hdrRowSrc Then
            dati = wsSrc.Range(wsSrc.Cells(hdrRowSrc + 1, 1), wsSrc.Cells(lastRowSrc, lastColSrc)).Value

            For r = 1 To UBound(dati, 1)
                codice = Trim$(CStr(dati(r, colCodSrc)))
                If Len(codice) > 0 Then
                    areaSrc = CStr(dati(r, colAreaSrc))
                    areaDst = TrascodificaAree(areaSrc, dictTrasc)

                    'riga nel layout del foglio di destinazione
                    ReDim vec(1 To lastColDst)
                    For c = 1 To lastColSrc
                        If srcToDst(c) > 0 Then
                            v = dati(r, c)
                            If IsError(v) Then
                                v = Empty
                            ElseIf VarType(v) = vbString Then
                                If Left$(v, 1) = "=" Then v = "'" & v
                            ElseIf VarType(v) = vbDate Then
                                If CDbl(v) <> Int(CDbl(v)) Then
                                    tipoData(srcToDst(c)) = 2
                                ElseIf tipoData(srcToDst(c)) = 0 Then
                                    tipoData(srcToDst(c)) = 1
                                End If
                            End If
                            vec(srcToDst(c)) = v
                        End If
                    Next c

                    If Len(areaDst) = 0 Then
                        'nessuna trascodifica: record non scritto nel foglio,
                        'ma riportato nel foglio degli scarti per verifica
                        If dictNew.Exists(codice) Then dictNew.Remove codice
                        If dictRun.Exists(codice) Then dictRun.Remove codice
                        dictScarti(codice) = vec
                    Else
                        vec(colAreaDst) = areaDst
                        dictNew(codice) = vec
                        dictRun(codice) = True
                        If dictScarti.Exists(codice) Then dictScarti.Remove codice
                    End If
                End If
            Next r
        End If

        wbSrc.Close SaveChanges:=False
        Set wbSrc = Nothing
    Next iFile

    '--- nulla da scrivere: non si cancella niente (probabile file sbagliato)
    If dictRun.Count = 0 Then
        MsgBox "Nessun record dei file selezionati ha una trascodifica dell'area." & vbCrLf & _
               "Nessuna modifica effettuata.", vbExclamation
        GoTo CleanExit
    End If

    nScart = dictScarti.Count

    '--- statistiche
    For Each v In dictRun.Keys
        If dictOld.Exists(CStr(v)) Then nAgg = nAgg + 1 Else nAdd = nAdd + 1
    Next v

    '--- riscrittura del foglio (in modalita' aggiornamento dictNew contiene
    '    anche le righe preesistenti non presenti nei file importati)
    lastUsed = wsDst.UsedRange.Row + wsDst.UsedRange.Rows.Count - 1
    If lastUsed > hdrRowDst Then
        wsDst.Range(wsDst.Cells(hdrRowDst + 1, 1), wsDst.Cells(lastUsed, lastColDst)).ClearContents
    End If

    n = dictNew.Count
    itemsNew = dictNew.Items
    ReDim outArr(1 To n, 1 To lastColDst)
    For k = 0 To n - 1
        vec = itemsNew(k)
        For c = 1 To lastColDst
            outArr(k + 1, c) = vec(c)
        Next c
    Next k

    For c = 1 To lastColDst
        If tipoData(c) = 2 Then
            wsDst.Range(wsDst.Cells(hdrRowDst + 1, c), wsDst.Cells(hdrRowDst + n, c)).NumberFormat = "dd/mm/yyyy hh:mm"
        ElseIf tipoData(c) = 1 Then
            wsDst.Range(wsDst.Cells(hdrRowDst + 1, c), wsDst.Cells(hdrRowDst + n, c)).NumberFormat = "dd/mm/yyyy"
        End If
    Next c
    wsDst.Range(wsDst.Cells(hdrRowDst + 1, 1), wsDst.Cells(hdrRowDst + n, lastColDst)).Value = outArr

    '--- righe scartate -> foglio "<archivio>_scarti" (stesso layout del foglio di destinazione)
    ScriviScarti LCase$(foglio) & "_scarti", wsDst, hdrRowDst, lastColDst, dictScarti, tipoData

    '--- colonna "memo collegati" del foglio Demand (T), sempre riallineata
    '    sia dopo l'import di Demand sia dopo quello di Memo
    AggiornaMemoCollegati

    ScriviLogReport wsReport, Now, nomiFile, foglio, nAgg, nAdd, nScart
    ok = True

    MsgBox "Import completato nel foglio " & foglio & "." & vbCrLf & _
           "Righe nel foglio: " & n & vbCrLf & _
           "  dai file: " & (nAgg + nAdd) & " (gia' presenti, sostituite: " & nAgg & "; nuove: " & nAdd & ")" & vbCrLf & _
           "Righe negli scarti (area senza trascodifica): " & nScart & _
           "  -> foglio " & LCase$(foglio) & "_scarti", vbInformation

CleanExit:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Errore durante l'import: " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'------------------------------------------------------------------------------
' Carica in un dizionario (codice -> riga nel layout del foglio) le righe gia'
' presenti in un foglio, a partire da primaRiga
'------------------------------------------------------------------------------
Private Sub CaricaRighe(ByVal ws As Worksheet, ByVal primaRiga As Long, ByVal colCod As Long, _
                        ByVal lastCol As Long, ByVal d As Object, ByRef tipoData() As Long)
    Dim lastRow As Long, r As Long, c As Long
    Dim dati As Variant, vec() As Variant
    Dim codice As String
    Dim v As Variant

    lastRow = ws.Cells(ws.Rows.Count, colCod).End(xlUp).Row
    If lastRow < primaRiga Then Exit Sub

    dati = ws.Range(ws.Cells(primaRiga, 1), ws.Cells(lastRow, lastCol)).Value
    For r = 1 To UBound(dati, 1)
        codice = Trim$(CStr(dati(r, colCod)))
        If Len(codice) > 0 Then
            ReDim vec(1 To lastCol)
            For c = 1 To lastCol
                v = dati(r, c)
                If IsError(v) Then
                    v = Empty
                ElseIf VarType(v) = vbString Then
                    If Left$(v, 1) = "=" Then v = "'" & v
                ElseIf VarType(v) = vbDate Then
                    If CDbl(v) <> Int(CDbl(v)) Then
                        tipoData(c) = 2
                    ElseIf tipoData(c) = 0 Then
                        tipoData(c) = 1
                    End If
                End If
                vec(c) = v
            Next c
            d(codice) = vec
        End If
    Next r
End Sub

'------------------------------------------------------------------------------
' Scrive i record scartati nel foglio indicato (lo crea se non esiste),
' sostituendone il contenuto. Intestazione copiata dal foglio di destinazione
' (riga 1); la colonna "Area richiedente omogeneizzata" resta vuota, l'area
' originale e' nella colonna di area del file sorgente.
'------------------------------------------------------------------------------
Private Sub ScriviScarti(ByVal nomeFoglio As String, ByVal wsDst As Worksheet, _
                         ByVal hdrRowDst As Long, ByVal lastColDst As Long, _
                         ByVal dictScarti As Object, ByRef tipoData() As Long)
    Dim ws As Worksheet
    Dim itemsS As Variant, outArr() As Variant, vec() As Variant
    Dim n As Long, k As Long, c As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomeFoglio)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nomeFoglio
    End If

    ws.Cells.ClearContents
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastColDst)).Value = _
        wsDst.Range(wsDst.Cells(hdrRowDst, 1), wsDst.Cells(hdrRowDst, lastColDst)).Value
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastColDst)).Font.Bold = True

    n = dictScarti.Count
    If n = 0 Then Exit Sub

    itemsS = dictScarti.Items
    ReDim outArr(1 To n, 1 To lastColDst)
    For k = 0 To n - 1
        vec = itemsS(k)
        For c = 1 To lastColDst
            outArr(k + 1, c) = vec(c)
        Next c
    Next k

    For c = 1 To lastColDst
        If tipoData(c) = 2 Then
            ws.Range(ws.Cells(2, c), ws.Cells(n + 1, c)).NumberFormat = "dd/mm/yyyy hh:mm"
        ElseIf tipoData(c) = 1 Then
            ws.Range(ws.Cells(2, c), ws.Cells(n + 1, c)).NumberFormat = "dd/mm/yyyy"
        End If
    Next c
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, lastColDst)).Value = outArr
End Sub

'------------------------------------------------------------------------------
' Foglio Demand, colonna T: codici Memo (col. A del foglio Memo) il cui campo
' "Codice Richiesta" (col. B) contiene il codice Demand; separati da ";"
'------------------------------------------------------------------------------
Private Sub AggiornaMemoCollegati()
    Dim wsMemo As Worksheet, wsDem As Worksheet
    Dim hdrM As Long, hdrD As Long
    Dim colMemoCod As Long, colMemoReq As Long, colDemCod As Long
    Dim lastM As Long, lastD As Long
    Dim hM As Variant, hD As Variant
    Dim datiM As Variant, datiD As Variant, outArr() As Variant
    Dim memoCodes() As String, memoReq() As String
    Dim nM As Long, r As Long, m As Long
    Dim codD As String, res As String, memoCod As String

    Set wsMemo = ThisWorkbook.Worksheets("Memo")
    Set wsDem = ThisWorkbook.Worksheets("Demand")

    hdrM = TrovaRigaIntestazione(wsMemo, "Codice", MAX_RIGHE_INTESTAZIONE)
    hdrD = TrovaRigaIntestazione(wsDem, "Codice richiesta", MAX_RIGHE_INTESTAZIONE)
    If hdrM = 0 Or hdrD = 0 Then Exit Sub

    hM = wsMemo.Range(wsMemo.Cells(hdrM, 1), wsMemo.Cells(hdrM, wsMemo.Cells(hdrM, wsMemo.Columns.Count).End(xlToLeft).Column)).Value
    hD = wsDem.Range(wsDem.Cells(hdrD, 1), wsDem.Cells(hdrD, wsDem.Cells(hdrD, wsDem.Columns.Count).End(xlToLeft).Column)).Value
    colMemoCod = FindHeaderCol(hM, "Codice")
    colMemoReq = FindHeaderCol(hM, "Codice Richiesta")
    colDemCod = FindHeaderCol(hD, "Codice richiesta")
    If colMemoCod = 0 Or colMemoReq = 0 Or colDemCod = 0 Then Exit Sub

    'Memo in memoria: codice Memo e testo integrale della colonna "Codice Richiesta"
    lastM = wsMemo.Cells(wsMemo.Rows.Count, colMemoCod).End(xlUp).Row
    If lastM > hdrM Then
        datiM = wsMemo.Range(wsMemo.Cells(hdrM + 1, 1), wsMemo.Cells(lastM, IIf(colMemoCod > colMemoReq, colMemoCod, colMemoReq))).Value
        ReDim memoCodes(1 To UBound(datiM, 1))
        ReDim memoReq(1 To UBound(datiM, 1))
        For r = 1 To UBound(datiM, 1)
            memoCod = Trim$(CStr(datiM(r, colMemoCod)))
            If Len(memoCod) > 0 Then
                nM = nM + 1
                memoCodes(nM) = memoCod
                memoReq(nM) = UCase$(CStr(datiM(r, colMemoReq)))
            End If
        Next r
    End If

    lastD = wsDem.Cells(wsDem.Rows.Count, colDemCod).End(xlUp).Row
    If lastD <= hdrD Then Exit Sub

    datiD = wsDem.Range(wsDem.Cells(hdrD + 1, colDemCod), wsDem.Cells(lastD, colDemCod)).Value
    ReDim outArr(1 To UBound(datiD, 1), 1 To 1)
    For r = 1 To UBound(datiD, 1)
        codD = UCase$(Trim$(CStr(datiD(r, 1))))
        If Len(codD) > 0 And nM > 0 Then
            res = ""
            For m = 1 To nM
                'il codice Demand e' semplicemente CONTENUTO nel campo del Memo
                '(da solo o insieme ad altri codici / testo)
                If InStr(1, memoReq(m), codD, vbBinaryCompare) > 0 Then
                    If Len(res) > 0 Then res = res & ";"
                    res = res & memoCodes(m)
                End If
            Next m
            If Len(res) > 0 Then outArr(r, 1) = res
        End If
    Next r
    wsDem.Range(wsDem.Cells(hdrD + 1, COL_DEMAND_MEMO), wsDem.Cells(lastD, COL_DEMAND_MEMO)).Value = outArr
End Sub

'------------------------------------------------------------------------------
' Dizionario valore-sorgente (normalizzato) -> valore trascodificato
' per l'archivio indicato ("Memo" / "Demand")
'------------------------------------------------------------------------------
Private Function CaricaTrascodifica(ByVal archivio As String) As Object
    Dim ws As Worksheet
    Dim d As Object
    Dim lastRow As Long, r As Long
    Dim dati As Variant
    Dim k As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Set ws = ThisWorkbook.Worksheets("Aree_trascodifica")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow >= 2 Then
        dati = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 3)).Value
        For r = 1 To UBound(dati, 1)
            If StrComp(Trim$(CStr(dati(r, 1))), archivio, vbTextCompare) = 0 Then
                k = NormKey(CStr(dati(r, 2)))
                If Len(k) > 0 And Len(Trim$(CStr(dati(r, 3)))) > 0 Then
                    d(k) = Trim$(CStr(dati(r, 3)))
                End If
            End If
        Next r
    End If

    Set CaricaTrascodifica = d
End Function

'------------------------------------------------------------------------------
' Trascodifica un valore di area (eventualmente multiplo, separato da ";").
' Restituisce i soli valori trascodificabili (senza duplicati) separati da "; ";
' stringa vuota se nessun valore ha una trascodifica.
'------------------------------------------------------------------------------
Private Function TrascodificaAree(ByVal areaSrc As String, ByVal dictT As Object) As String
    Dim parti As Variant
    Dim i As Long
    Dim k As String, v As String, res As String, seen As String

    k = NormKey(areaSrc)
    If Len(k) = 0 Then Exit Function

    'prima il valore intero (alcune denominazioni contengono gia' un ";")
    If dictT.Exists(k) Then
        TrascodificaAree = dictT(k)
        Exit Function
    End If

    parti = Split(areaSrc, ";")
    seen = "|"
    For i = LBound(parti) To UBound(parti)
        k = NormKey(CStr(parti(i)))
        If Len(k) > 0 Then
            If dictT.Exists(k) Then
                v = dictT(k)
                If InStr(1, seen, "|" & v & "|", vbTextCompare) = 0 Then
                    seen = seen & v & "|"
                    If Len(res) > 0 Then res = res & "; "
                    res = res & v
                End If
            End If
        End If
    Next i

    TrascodificaAree = res
End Function

'------------------------------------------------------------------------------
' Normalizza una denominazione per il confronto: spazi non separabili,
' apostrofi tipografici, ritorni a capo e spazi multipli; maiuscolo.
'------------------------------------------------------------------------------
Private Function NormKey(ByVal s As String) As String
    s = Replace(s, ChrW$(160), " ")
    s = Replace(s, ChrW$(8217), "'")
    s = Replace(s, ChrW$(8216), "'")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormKey = UCase$(Trim$(s))
End Function

Private Sub ImpostaIntestazione(ByVal ws As Worksheet, ByVal hdrRow As Long, _
                                ByVal col As Long, ByVal testo As String)
    If Len(Trim$(CStr(ws.Cells(hdrRow, col).Value))) = 0 Then
        ws.Cells(hdrRow, col).Value = testo
    End If
End Sub

'--- cerca in quale riga (entro maxRows) compare l'intestazione indicata
Private Function TrovaRigaIntestazione(ws As Worksheet, nomeColonna As String, maxRows As Long) As Long
    Dim r As Long, c As Long, lastCol As Long
    Dim testo As String

    For r = 1 To maxRows
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If lastCol > 0 Then
            For c = 1 To lastCol
                testo = Trim$(CStr(ws.Cells(r, c).Value))
                If StrComp(testo, nomeColonna, vbTextCompare) = 0 Then
                    TrovaRigaIntestazione = r
                    Exit Function
                End If
            Next c
        End If
    Next r

    TrovaRigaIntestazione = 0
End Function

Private Function FindHeaderCol(ByVal headersRow As Variant, ByVal headerName As String) As Long
    Dim c As Long, n As Long
    n = UBound(headersRow, 2)
    For c = 1 To n
        If StrComp(Trim$(CStr(headersRow(1, c))), Trim$(headerName), vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

'--- Log nel foglio Report:
'    data | file importato | archivio aggiornato | righe aggiornate | righe aggiunte | righe scartate
Private Sub ScriviLogReport(ByVal ws As Worksheet, ByVal dt As Date, ByVal fileName As String, _
                            ByVal foglio As String, ByVal updated As Long, ByVal added As Long, _
                            ByVal scartate As Long)
    Dim nextRow As Long

    If Len(Trim$(CStr(ws.Cells(1, 1).Value))) = 0 And _
       Len(Trim$(CStr(ws.Cells(1, 2).Value))) = 0 And _
       Len(Trim$(CStr(ws.Cells(1, 3).Value))) = 0 And _
       Len(Trim$(CStr(ws.Cells(1, 4).Value))) = 0 Then
        ws.Cells(1, 1).Value = "data"
        ws.Cells(1, 2).Value = "file importato"
        ws.Cells(1, 3).Value = "archivio aggiornato"
        ws.Cells(1, 4).Value = "righe aggiornate"
        ws.Cells(1, 5).Value = "righe aggiunte"
    End If
    If Len(Trim$(CStr(ws.Cells(1, 6).Value))) = 0 Then
        ws.Cells(1, 6).Value = "righe scartate (senza trascodifica)"
    End If

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = dt
    ws.Cells(nextRow, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(nextRow, 2).Value = fileName
    ws.Cells(nextRow, 3).Value = foglio
    ws.Cells(nextRow, 4).Value = updated
    ws.Cells(nextRow, 5).Value = added
    ws.Cells(nextRow, 6).Value = scartate
End Sub
