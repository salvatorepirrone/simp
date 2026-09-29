Attribute VB_Name = "ImportaMensile"
'==============================================================================
' MODULO: ImportaMensile  (v13)
'
' MACRO 1: ImportaDatiMensili
'   1. VBA esporta foglio "prodotti" in CSV temporaneo e foglio "territorio"
'      in CSV temporaneo (serve per l'aggancio Codice Sede -> Provincia)
'   2. VBA chiama filtra_consolidato_provincia.py via bat:
'      - legge le righe del file mensile con calamine (Rust, veloce)
'      - filtra DC Pensione e righe non-zero (Tot.Pervenuti>0 o Tot.Definiti>0)
'      - AGGREGA le righe filtrate a livello di Descrizione Provincia
'        (una riga per provincia anziche' una per sede: e' la modifica
'        principale della v13, per ridurre il peso del Consolidato)
'      - scrive CSV ridotto/aggregato + il corrispondente XLSX + restituisce
'        mese/anno e conteggi nell'output
'   3. VBA legge mese/anno dall'output Python (NO apertura xlsx)
'   4. VBA legge CSV ridotto, calcola il DIFFERENZIALE vs mese precedente
'      (chiave Codice+Provincia, non piu' Codice+CodSede), scrive in
'      blocco nel Consolidato di ThisWorkbook
'
' MACRO 2: Filtra / FiltraRegione / FiltraProvincia
'   Invariate nella logica (vedi versioni precedenti del modulo): estraggono
'   un file autonomo (non toccano il Consolidato). Aggiornato solo il modo
'   in cui vengono risolti i percorsi (vedi CONFIGURAZIONE SHAREPOINT sotto).
'
' MACRO 3 (NUOVA): MigraStoricoProvincia
'   Da lanciare UNA SOLA VOLTA per convertire i mesi gia' presenti nel
'   Consolidato (che sono a livello di sede) al nuovo layout a livello di
'   provincia, cosi' l'intero foglio torna ad avere granularita' uniforme.
'   Operazione distruttiva sul contenuto dati del foglio Consolidato: fare
'   un backup del file PRIMA di lanciarla. Vedi commenti nella Sub per i
'   dettagli.
'
' CAMBIO DI SCHEMA (v13): il Consolidato non ha piu' le colonne "Codice
' Sede"/"Sede" (livello sede): al loro posto c'e' un'unica colonna
' "Descrizione Provincia" (livello provincia), ottenuta con lo stesso
' aggancio gia' usato dalla formula "Territorio" nelle macro Filtra/
' FiltraProvincia (Codice Sede -> foglio "territorio" colonna D -> colonna
' C = Descrizione Provincia). Questo riduce drasticamente il numero di
' righe scritte ogni mese (una per provincia anziche' una per sede) ed e'
' la causa principale per cui il file era troppo pesante.
'
' CONFIGURAZIONE SHAREPOINT (v13):
'   In precedenza SCRIPT_DIR e BASE_DIR erano percorsi locali fissi
'   (es. "C:\INPS\...") perche' ThisWorkbook.Path restituisce un URL (non
'   utilizzabile con Dir()/MkDir()) quando il file con le macro e' aperto
'   da SharePoint/OneDrive. La soluzione NON e' usare l'URL del sito, ma
'   il percorso locale della cartella di sincronizzazione OneDrive/
'   SharePoint (che esiste sempre, anche quando il file e' "online"):
'   qualcosa come "C:\Users\<utente>\<Organizzazione>\<Sito> - <Libreria>\...".
'   Da qui in poi SCRIPT_DIR e BASE_DIR vengono risolti automaticamente A
'   RUNTIME (funzioni GetScriptDir()/GetBaseDir()), individuando quella
'   cartella tramite il registro di Windows (chiave dove OneDrive registra
'   ogni libreria SharePoint sincronizzata, con l'indirizzo del sito e il
'   percorso locale corrispondente). Occorre solo indicare, nelle costanti
'   SHAREPOINT_SITE_MATCH / SHAREPOINT_LIB_MATCH qui sotto, una porzione
'   (case-insensitive) dell'indirizzo del sito e del nome della libreria
'   SharePoint da usare, cosi' la macro individua la cartella giusta anche
'   se cambiano PC/utente/lettera di unita'. Se l'autorilevamento
'   fallisse sul tuo PC (es. politiche aziendali particolari), puoi sempre
'   forzare il percorso a mano con SHAREPOINT_ROOT_OVERRIDE.
'
' PREREQUISITI:
'   - py -m pip install python-calamine openpyxl
'   - filtra_consolidato_provincia.py, filtra_dcpensioni.py,
'     filtra_dcpensioni_regione.py, filtra_dcpensioni_provincia.py nella
'     cartella restituita da GetScriptDir() (vedi sopra)
'   - fogli "territorio" e "prodotti" presenti in ThisWorkbook
'   - OneDrive in esecuzione, con la libreria SharePoint scelta
'     effettivamente sincronizzata (non solo "disponibile online")
'==============================================================================

Option Explicit

' --- Interprete e script Python ---
' PYTHON_EXE non e' piu' un percorso fisso con lo username in chiaro (che
' cambierebbe da un PC/utente all'altro): viene individuato a runtime
' (funzione GetPythonExe(), cache in mPythonExe) cercando "py"/"python" nel
' PATH di sistema. Se sul tuo PC l'autorilevamento non funzionasse, puoi
' comunque forzare un percorso fisso valorizzando PYTHON_EXE_OVERRIDE.
Private Const PYTHON_EXE_OVERRIDE As String = ""
Private Const SCRIPT_NAME2 As String = "filtra_consolidato_provincia.py"   ' usato da ImportaDatiMensili (v13)
Private Const SCRIPT_NAME3 As String = "filtra_dcpensioni.py"              ' usato da Filtra
Private Const SCRIPT_NAME4 As String = "filtra_dcpensioni_regione.py"      ' usato da FiltraRegione
Private Const SCRIPT_NAME5 As String = "filtra_dcpensioni_provincia.py"    ' usato da FiltraProvincia

Private mPythonExe As String
Private mPythonExeResolved As Boolean

' --- Sottocartelle (relative alla cartella SharePoint sincronizzata) ---
Private Const SCRIPT_SUBDIR As String = "Scripts\"
Private Const DATA_SUBDIR  As String = "SIMP_Pensioni\"
Private Const CSV_SUBDIR   As String = "csv"
Private Const XLSX_SUBDIR  As String = "DCPensioni"

' --- Individuazione della cartella SharePoint sincronizzata ---
' Porzione (case-insensitive) dell'indirizzo del sito e del nome della
' libreria SharePoint da usare: MODIFICARE con i valori corretti per il
' proprio tenant (es. "sites/DCPensioni" e "Documenti condivisi").
Private Const SHAREPOINT_SITE_MATCH As String = "sites/"
Private Const SHAREPOINT_LIB_MATCH  As String = ""
' Se valorizzata, salta l'autorilevamento e usa direttamente questo
' percorso locale (es. "C:\Users\mario.rossi\Contoso\DCPensioni - Documenti\").
Private Const SHAREPOINT_ROOT_OVERRIDE As String = ""

Private mSharePointRoot As String
Private mSharePointResolved As Boolean

' --- Colonne SORGENTE nel CSV filtrato NON aggregato (1-based) - usate da
'     Filtra/FiltraRegione tramite gli script filtra_dcpensioni*.py ---
Private Const S_AREA     As Integer = 1
Private Const S_PRODOTTO As Integer = 2
Private Const S_CODICE   As Integer = 3
Private Const S_DESCRIZ  As Integer = 4
Private Const S_REGIONE  As Integer = 5
Private Const S_CODSEDE  As Integer = 6
Private Const S_SEDE     As Integer = 7
Private Const S_GIACINIZ As Integer = 8
Private Const S_P0       As Integer = 9
Private Const S_TOTPERV  As Integer = 20
Private Const S_TOTDEF   As Integer = 32
Private Const S_GIACFIN  As Integer = 33
Private Const S_OMOG     As Integer = 34
Private Const S_GIACGG   As Integer = 35

' --- Colonne nel CSV aggregato per PROVINCIA scritto da
'     filtra_consolidato_provincia.py (letto da ImportaDatiMensili) ---
Private Const SP_AREA      As Integer = 1
Private Const SP_PRODOTTO  As Integer = 2
Private Const SP_CODICE    As Integer = 3
Private Const SP_DESCRIZ   As Integer = 4
Private Const SP_REGIONE   As Integer = 5
Private Const SP_PROVINCIA As Integer = 6
Private Const SP_GIACINIZ  As Integer = 7
Private Const SP_P0        As Integer = 8
Private Const SP_GIACFIN   As Integer = 32
Private Const SP_OMOG      As Integer = 33
Private Const SP_GIACGG    As Integer = 34

' --- Colonne DESTINATARIO Consolidato - NUOVO layout a Provincia (1-based) ---
Private Const D_ANNO      As Integer = 1
Private Const D_MESE      As Integer = 2
Private Const D_AREA      As Integer = 3
Private Const D_PRODOTTO  As Integer = 4
Private Const D_CODICE    As Integer = 5
Private Const D_DESCRIZ   As Integer = 6
Private Const D_REGIONE   As Integer = 7
Private Const D_PROVINCIA As Integer = 8
Private Const D_GIACINIZ  As Integer = 9
Private Const D_P0        As Integer = 10
Private Const D_TOTPERV   As Integer = 21
Private Const D_TOTDEF    As Integer = 33
Private Const D_GIACFIN   As Integer = 34
Private Const D_OMOG      As Integer = 35
Private Const D_GIACGG    As Integer = 36
Private Const D_AM        As Integer = 37   ' Coefficiente
Private Const N_DIFF      As Integer = 24
Private Const D_NCOLS     As Integer = 36
Private Const PRIMA_RIGA  As Integer = 5

' --- Colonne Consolidato - VECCHIO layout a Sede, usate SOLO da
'     MigraStoricoProvincia per leggere i dati storici pre-migrazione ---
Private Const OLD_D_ANNO     As Integer = 1
Private Const OLD_D_MESE     As Integer = 2
Private Const OLD_D_AREA     As Integer = 3
Private Const OLD_D_PRODOTTO As Integer = 4
Private Const OLD_D_CODICE   As Integer = 5
Private Const OLD_D_DESCRIZ  As Integer = 6
Private Const OLD_D_REGIONE  As Integer = 7
Private Const OLD_D_CODSEDE  As Integer = 8
Private Const OLD_D_GIACINIZ As Integer = 10
Private Const OLD_D_P0       As Integer = 11
Private Const OLD_D_GIACFIN  As Integer = 35
Private Const OLD_D_OMOG     As Integer = 36
Private Const OLD_D_GIACGG   As Integer = 37
Private Const OLD_D_NCOLS    As Integer = 37


'==============================================================================
Sub ImportaDatiMensili()

    Dim spRoot As String
    spRoot = GetSharePointRoot()
    If spRoot = "" Then Exit Sub
    If GetPythonExe() = "" Then Exit Sub

    Dim wbDest        As Workbook
    Dim wsConsolidato As Worksheet
    Dim wsProdotti    As Worksheet
    Dim wsTerritorio  As Worksheet
    Set wbDest = ThisWorkbook
    Set wsConsolidato = wbDest.Sheets("Consolidato")
    Set wsProdotti = wbDest.Sheets("prodotti")
    Set wsTerritorio = wbDest.Sheets("territorio")

    ' 1. SELEZIONE FILE
    Dim filePath As String
    filePath = Application.GetOpenFilename( _
        FileFilter:="File Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="Seleziona il file di produzione mensile da importare")
    If filePath = "False" Then
        MsgBox "Operazione annullata.", vbInformation, "Annullato"
        Exit Sub
    End If

    ' 2. VERIFICA SCRIPT PYTHON
    Dim scriptDir As String
    Dim scriptPath As String
    scriptDir = GetScriptDir()
    scriptPath = scriptDir & SCRIPT_NAME2
    AssicuraCartella scriptDir
    If Dir(scriptPath) = "" Then
        MsgBox "Script non trovato: " & scriptPath & vbCrLf & vbCrLf & _
               "Copiare filtra_consolidato_provincia.py in: " & scriptDir, _
               vbCritical, "Script mancante"
        Exit Sub
    End If

    ' 3. ESPORTA PRODOTTI E TERRITORIO IN CSV TEMPORANEO
    Application.StatusBar = "Esportazione prodotti e territorio..."
    Dim prodottiCsv As String
    Dim territorioCsv As String
    prodottiCsv = GetTempDir() & "dc_prodotti_tmp.csv"
    territorioCsv = GetTempDir() & "dc_territorio_tmp.csv"
    EsportaProdottiCsv wsProdotti, prodottiCsv
    EsportaFoglioCsv wsTerritorio, territorioCsv

    ' 4. CHIAMA PYTHON PER FILTRAGGIO + AGGREGAZIONE PER PROVINCIA
    '    Python legge mese/anno da A1 e li restituisce nell'output
    '    -> nessuna apertura xlsx da VBA
    Dim xlsxName As String
    xlsxName = Mid(filePath, InStrRev(filePath, "\") + 1)
    xlsxName = Left(xlsxName, InStrRev(xlsxName, ".") - 1)

    Dim csvDir  As String
    Dim xlsxDir As String
    csvDir = GetBaseDir() & CSV_SUBDIR & "\"
    xlsxDir = GetBaseDir() & XLSX_SUBDIR & "\"

    AssicuraCartella csvDir
    AssicuraCartella xlsxDir

    Dim outputCsv  As String
    Dim outputXlsx As String
    Dim logPath    As String
    Dim batPath    As String
    Dim errPath    As String
    outputCsv = csvDir & xlsxName & "_filtrato_provincia.csv"
    outputXlsx = xlsxDir & xlsxName & "_filtrato_provincia.xlsx"
    logPath = GetTempDir() & "dc_filtra.log"
    batPath = GetTempDir() & "dc_filtra.bat"
    errPath = GetTempDir() & "dc_filtra_err.log"

    If Dir(outputCsv) <> "" Then Kill outputCsv
    If Dir(outputXlsx) <> "" Then Kill outputXlsx
    If Dir(logPath) <> "" Then Kill logPath
    If Dir(errPath) <> "" Then Kill errPath

    Dim batLine As String
    batLine = "@echo off" & vbCrLf & _
              """" & GetPythonExe() & """ " & _
              """" & scriptPath & """ " & _
              """" & filePath & """ " & _
              """" & prodottiCsv & """ " & _
              """" & territorioCsv & """ " & _
              """" & outputCsv & """ " & _
              """" & outputXlsx & """ " & _
              "1> """ & logPath & """ " & _
              "2> """ & errPath & """" & vbCrLf & _
              "if %ERRORLEVEL% NEQ 0 type """ & errPath & """ >> """ & logPath & """"

    Dim iFile As Integer
    iFile = FreeFile
    Open batPath For Output As #iFile
    Print #iFile, batLine
    Close #iFile

    Application.StatusBar = "Fase 1/4 - Pre-filtraggio e aggregazione con Python (attendere)..."
    Shell "cmd.exe /c """ & batPath & """", vbHide

    ' Attendi che Python scriva OK o errore nel log (max 180s)
    Dim t0 As Single
    Dim logContent As String
    Dim elapsed As Long
    t0 = Timer
    Do
        DoEvents
        Application.Wait Now + TimeValue("00:00:02")
        logContent = LeggiFile(logPath)
        elapsed = CLng(Timer - t0)
        Application.StatusBar = "Fase 1/4 - Pre-filtraggio Python... " & elapsed & "s"
        If Left(Trim(logContent), 2) = "OK" Then Exit Do
        If Left(Trim(logContent), 6) = "ERRORE" Then Exit Do
        If elapsed > 180 Then Exit Do
    Loop

    If Left(Trim(logContent), 2) <> "OK" Then
        MsgBox "Errore nello script Python:" & vbCrLf & Left(logContent, 600), _
               vbCritical, "Errore Python"
        GoTo CleanupTemp
    End If

    ' Output formato: OK|n_tot|n_out|n_sca|n_gruppi|cellA1
    Dim parts() As String
    parts = Split(Trim(logContent), "|")
    If UBound(parts) < 5 Then
        MsgBox "Output Python non valido: " & logContent, vbCritical, "Errore"
        GoTo CleanupTemp
    End If
    ' Pulizia difensiva: rimuove eventuali CR/LF residui in ciascun campo
    Dim iPulisci As Integer
    For iPulisci = LBound(parts) To UBound(parts)
        Do While Len(parts(iPulisci)) > 0 And _
                 (Right(parts(iPulisci), 1) = vbCr Or Right(parts(iPulisci), 1) = vbLf)
            parts(iPulisci) = Left(parts(iPulisci), Len(parts(iPulisci)) - 1)
        Loop
        parts(iPulisci) = Trim(parts(iPulisci))
    Next iPulisci
    Dim nTot As Long
    Dim nOut As Long
    Dim nSca As Long
    Dim nGruppi As Long
    Dim cellA1 As String
    nTot = CLng(parts(1))
    nOut = CLng(parts(2))
    nSca = CLng(parts(3))
    nGruppi = CLng(parts(4))
    cellA1 = parts(5)

    ' 5. PARSING MESE/ANNO dall'output Python (non apriamo xlsx)
    Dim mese As Integer
    Dim anno As Integer
    If Not ParseMeseAnno(cellA1, mese, anno) Then
        MsgBox "Impossibile leggere Mese/Anno dall'output Python." & vbCrLf & _
               "Valore ricevuto: """ & cellA1 & """" & vbCrLf & _
               "Formato atteso: ""Consolidato M/AAAA""", vbCritical, "Errore"
        GoTo CleanupTemp
    End If

    ' 6. CONTROLLI SUI MESI
    Dim mesiPresenti(1 To 12) As Boolean
    AnalizzaMesiPresenti wsConsolidato, anno, mesiPresenti

    If mesiPresenti(mese) Then
        Dim risp As Integer
        risp = MsgBox("Il mese " & mese & "/" & anno & " e' gia' presente." & vbCrLf & _
                      "Eliminare e reimportare?", vbYesNo + vbExclamation, "Mese gia' presente")
        If risp = vbNo Then GoTo CleanupTemp
        EliminaRigheMese wsConsolidato, mese, anno
        AnalizzaMesiPresenti wsConsolidato, anno, mesiPresenti
    End If

    If mese > 1 Then
        If Not mesiPresenti(mese - 1) Then
            MsgBox "Mese precedente " & (mese - 1) & "/" & anno & " non presente." & vbCrLf & _
                   "Importarlo prima.", vbCritical, "Mese mancante"
            GoTo CleanupTemp
        End If
    End If

    ' 7. CARICA DATI PRECEDENTI PER DIFFERENZIALE (solo se mese > 1)
    Application.StatusBar = "Fase 2/4 - Caricamento dati precedenti..."
    Dim dictPrec As Object
    Set dictPrec = CreateObject("Scripting.Dictionary")
    If mese > 1 Then
        Set dictPrec = CaricaDatiPrecedenti(wsConsolidato, mese - 1, anno)
    End If

    ' 8. LEGGI CSV FILTRATO/AGGREGATO E COSTRUISCI ARRAY DI OUTPUT
    Application.StatusBar = "Fase 3/4 - Lettura e calcolo differenziale (" & nGruppi & " righe per provincia)..."

    Dim outData() As Variant
    If nGruppi > 0 Then ReDim outData(1 To nGruppi, 1 To D_NCOLS)

    Dim nLetti As Long
    nLetti = 0

    iFile = FreeFile
    Open outputCsv For Input As #iFile

    ' Riga 1 = cellA1, riga 2 = header: scarta entrambe
    Dim sLine As String
    Line Input #iFile, sLine
    Line Input #iFile, sLine

    Do While Not EOF(iFile) And nLetti < nGruppi
        Line Input #iFile, sLine
        If Trim(sLine) = "" Then GoTo NextLinea

        Dim cols() As String
        cols = Split(sLine, ";")
        If UBound(cols) < SP_GIACGG - 1 Then GoTo NextLinea

        nLetti = nLetti + 1

        outData(nLetti, D_ANNO) = anno
        outData(nLetti, D_MESE) = mese

        ' Copia campi descrittivi: Area, Prodotto, Codice, Descriz, Regione, Provincia
        Dim c As Integer
        For c = 0 To 5
            outData(nLetti, D_AREA + c) = cols(SP_AREA + c - 1)
        Next c

        ' Chiave: Codice (col E) + Provincia (col H, testo)
        Dim chiave As String
        chiave = NormCodice(cols(SP_CODICE - 1)) & "|" & Trim(cols(SP_PROVINCIA - 1))

        Dim hasPrev As Boolean
        Dim prevArr() As Double
        hasPrev = dictPrec.Exists(chiave)
        If hasPrev Then prevArr = dictPrec(chiave)

        ' GiacIniz: differenziale
        Dim vGiac As Double
        vGiac = StrToDouble(cols(SP_GIACINIZ - 1))
        If hasPrev Then
            outData(nLetti, D_GIACINIZ) = vGiac - prevArr(N_DIFF + 1)
        Else
            outData(nLetti, D_GIACINIZ) = vGiac
        End If

        ' P0-P23: differenziale
        Dim idx As Integer
        Dim vC As Double
        For idx = 0 To N_DIFF - 1
            vC = StrToDouble(cols(SP_P0 + idx - 1))
            If hasPrev Then
                outData(nLetti, D_P0 + idx) = vC - prevArr(idx)
            Else
                outData(nLetti, D_P0 + idx) = vC
            End If
        Next idx

        ' GiacFin: valore assoluto
        outData(nLetti, D_GIACFIN) = StrToDouble(cols(SP_GIACFIN - 1))

        ' Omog: differenziale
        Dim vOmog As Double
        vOmog = StrToDouble(cols(SP_OMOG - 1))
        If hasPrev Then
            outData(nLetti, D_OMOG) = vOmog - prevArr(N_DIFF)
        Else
            outData(nLetti, D_OMOG) = vOmog
        End If

        ' GiacGG: valore assoluto
        outData(nLetti, D_GIACGG) = StrToDouble(cols(SP_GIACGG - 1))

NextLinea:
    Loop
    Close #iFile

    ' 9. SCRITTURA MASSIVA NEL CONSOLIDATO
    If nLetti = 0 Then
        MsgBox "Nessuna riga DC Pensione trovata.", vbInformation
        GoTo CleanupTemp
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Fase 4/4 - Scrittura " & nLetti & " righe nel Consolidato..."

    Dim lastDst As Long
    lastDst = wsConsolidato.Cells(wsConsolidato.Rows.Count, D_PRODOTTO).End(xlUp).Row
    If lastDst < PRIMA_RIGA - 1 Then lastDst = PRIMA_RIGA - 1
    Dim rDest As Long
    rDest = lastDst + 1

    wsConsolidato.Range( _
        wsConsolidato.Cells(rDest, 1), _
        wsConsolidato.Cells(rDest + nLetti - 1, D_NCOLS) _
    ).Value = outData

    Application.StatusBar = "Fase 4/4 - Scrittura formule..."

    Dim rFirst As Long
    Dim rLast As Long
    rFirst = rDest
    rLast = rDest + nLetti - 1
    ScriviFormuleProvincia wsConsolidato, rFirst, rLast

    ' 10. AGGIORNA RANGE E REFRESH PIVOT
    Application.StatusBar = "Aggiornamento pivot..."
    Dim lastPivotRow As Long
    lastPivotRow = wsConsolidato.Cells(wsConsolidato.Rows.Count, D_PRODOTTO).End(xlUp).Row
    Dim srcRange As String
    srcRange = "Consolidato!$A$4:$AP$" & lastPivotRow
    Dim ws2 As Worksheet
    Dim pt As PivotTable
    For Each ws2 In wbDest.Worksheets
        For Each pt In ws2.PivotTables
            On Error Resume Next
            pt.ChangePivotCache wbDest.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=srcRange)
            pt.RefreshTable
            On Error GoTo 0
        Next pt
    Next ws2

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Importazione completata!" & vbCrLf & vbCrLf & _
           "Mese: " & mese & "/" & anno & vbCrLf & _
           "Righe totali nel file:      " & nTot & vbCrLf & _
           "Righe DC Pensioni filtrate: " & nOut & vbCrLf & _
           "Righe scartate:             " & nSca & vbCrLf & _
           "Righe aggregate per prov.:  " & nLetti & vbCrLf & vbCrLf & _
           "CSV:  " & outputCsv & vbCrLf & _
           "XLSX: " & outputXlsx, vbInformation, "OK"

CleanupTemp:
    On Error Resume Next
    If Dir(prodottiCsv) <> "" Then Kill prodottiCsv
    If Dir(territorioCsv) <> "" Then Kill territorioCsv
    If Dir(batPath) <> "" Then Kill batPath
    If Dir(errPath) <> "" Then Kill errPath
    On Error GoTo 0
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


'==============================================================================
' MACRO: MigraStoricoProvincia   (NUOVA, v13)
'
' Da lanciare UNA SOLA VOLTA (dopo aver fatto un backup del file!) per
' convertire i mesi gia' presenti nel Consolidato (livello sede, vecchio
' layout con colonne "Codice Sede"/"Sede") al nuovo layout a livello di
' provincia (colonna unica "Descrizione Provincia"), lo stesso usato da
' ImportaDatiMensili a partire dalla v13.
'
' Logica:
'   1. Legge tutte le righe dati esistenti nel Consolidato (vecchio layout)
'   2. Per ognuna, risale alla Provincia tramite il foglio "territorio"
'      (stesso aggancio Codice Sede -> Provincia usato dalle formule
'      "Territorio" in Filtra/FiltraProvincia)
'   3. Somma (per Anno+Mese+Area+Prodotto+Codice+Descrizione+Regione+
'      Provincia) Giacenza Iniziale, P0..P23, Giacenza Finale, Totale
'      Omogeneizzato, Giacenza Gg
'   4. Sostituisce il contenuto dati del Consolidato con le righe
'      aggregate, nel nuovo layout, con le formule Coefficiente/
'      IndiceDeflusso/GiacenzaOmog/PervenutoOmog/Aggregati/Gestione
'
' ATTENZIONE: dopo la migrazione, i campi delle eventuali tabelle pivot che
' facevano riferimento a "Codice Sede"/"Sede" non sono piu' validi e vanno
' sostituiti manualmente con "Descrizione Provincia".
'==============================================================================
Sub MigraStoricoProvincia()

    Dim wsConsolidato As Worksheet
    Dim wsTerritorio  As Worksheet
    Set wsConsolidato = ThisWorkbook.Sheets("Consolidato")
    Set wsTerritorio = ThisWorkbook.Sheets("territorio")

    Dim lastRow As Long
    lastRow = wsConsolidato.Cells(wsConsolidato.Rows.Count, OLD_D_PRODOTTO).End(xlUp).Row
    If lastRow < PRIMA_RIGA Then
        MsgBox "Nessun dato storico da migrare nel Consolidato.", vbInformation
        Exit Sub
    End If

    Dim nRigheOriginarie As Long
    nRigheOriginarie = lastRow - PRIMA_RIGA + 1

    If MsgBox("ATTENZIONE: questa operazione riscrive TUTTO il contenuto dati del foglio " & _
              "'Consolidato', aggregando le righe esistenti (per sede) a livello di provincia." & vbCrLf & vbCrLf & _
              "Operazione da lanciare UNA SOLA VOLTA. Assicurati di avere un BACKUP del file " & _
              "PRIMA di continuare (l'Annulla di Excel potrebbe non essere sufficiente)." & vbCrLf & vbCrLf & _
              "Righe attuali nel Consolidato: " & nRigheOriginarie & vbCrLf & vbCrLf & _
              "Continuare?", vbYesNo + vbExclamation + vbDefaultButton2, "Migrazione a Provincia") = vbNo Then
        Exit Sub
    End If

    Application.StatusBar = "Migrazione: lettura dati storici..."
    Dim datiVecchi As Variant
    datiVecchi = wsConsolidato.Range(wsConsolidato.Cells(PRIMA_RIGA, 1), wsConsolidato.Cells(lastRow, OLD_D_NCOLS)).Value

    ' Mappa Codice Sede -> Descrizione Provincia dal foglio "territorio"
    ' (colonna D = Codice Sede, colonna C = Descrizione Provincia; stesso
    ' aggancio della formula "Territorio" usata in Filtra/FiltraProvincia)
    Dim mapProv As Object
    Set mapProv = CreateObject("Scripting.Dictionary")
    Dim lastRowTerr As Long
    lastRowTerr = wsTerritorio.Cells(wsTerritorio.Rows.Count, 4).End(xlUp).Row
    Dim rt As Long
    Dim codT As String
    Dim provT As String
    For rt = 2 To lastRowTerr
        codT = Norm6(wsTerritorio.Cells(rt, 4).Value)
        provT = Trim(CStr(wsTerritorio.Cells(rt, 3).Value))
        If codT <> "" And provT <> "" Then
            If Not mapProv.Exists(codT) Then mapProv.Add codT, provT
        End If
    Next rt
    If mapProv.Count = 0 Then
        MsgBox "Impossibile leggere la mappa Codice Sede -> Provincia dal foglio 'territorio'.", vbCritical
        Exit Sub
    End If

    Application.StatusBar = "Migrazione: aggregazione per provincia..."
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim ordine As Collection
    Set ordine = New Collection

    Dim r As Long
    Dim n_senza_prov As Long
    Dim anno As Integer, mese As Integer
    Dim area As String, prodotto As String, codice As String, descriz As String, regione As String
    Dim codSede As String, provincia As String
    Dim chiave As String
    Dim acc() As Double
    Dim idx As Integer

    For r = 1 To UBound(datiVecchi, 1)
        If Trim(CStr(datiVecchi(r, OLD_D_PRODOTTO))) = "" Then GoTo NextR

        anno = CInt(datiVecchi(r, OLD_D_ANNO))
        mese = CInt(datiVecchi(r, OLD_D_MESE))
        area = CStr(datiVecchi(r, OLD_D_AREA))
        prodotto = CStr(datiVecchi(r, OLD_D_PRODOTTO))
        codice = CStr(datiVecchi(r, OLD_D_CODICE))
        descriz = CStr(datiVecchi(r, OLD_D_DESCRIZ))
        regione = CStr(datiVecchi(r, OLD_D_REGIONE))

        codSede = Norm6(datiVecchi(r, OLD_D_CODSEDE))
        If mapProv.Exists(codSede) Then
            provincia = mapProv(codSede)
        Else
            provincia = "(Provincia sconosciuta)"
            n_senza_prov = n_senza_prov + 1
        End If

        chiave = anno & "|" & mese & "|" & area & "|" & prodotto & "|" & codice & "|" & descriz & "|" & regione & "|" & provincia

        If dict.Exists(chiave) Then
            acc = dict(chiave)
        Else
            ReDim acc(0 To N_DIFF + 3)  ' 0=GiacIniz, 1..24=P0..P23, 25=GiacFin, 26=Omog, 27=GiacGG
            ordine.Add chiave
        End If

        acc(0) = acc(0) + ToDouble(datiVecchi(r, OLD_D_GIACINIZ))
        For idx = 0 To N_DIFF - 1
            acc(1 + idx) = acc(1 + idx) + ToDouble(datiVecchi(r, OLD_D_P0 + idx))
        Next idx
        acc(1 + N_DIFF) = acc(1 + N_DIFF) + ToDouble(datiVecchi(r, OLD_D_GIACFIN))
        acc(1 + N_DIFF + 1) = acc(1 + N_DIFF + 1) + ToDouble(datiVecchi(r, OLD_D_OMOG))
        acc(1 + N_DIFF + 2) = acc(1 + N_DIFF + 2) + ToDouble(datiVecchi(r, OLD_D_GIACGG))

        dict(chiave) = acc
NextR:
    Next r

    If ordine.Count = 0 Then
        MsgBox "Nessuna riga valida trovata da migrare.", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "Migrazione: scrittura " & ordine.Count & " righe aggregate..."
    Dim outData() As Variant
    ReDim outData(1 To ordine.Count, 1 To D_NCOLS)

    Dim i As Long
    Dim k As String
    Dim parti() As String
    For i = 1 To ordine.Count
        k = ordine(i)
        parti = Split(k, "|")
        acc = dict(k)

        outData(i, D_ANNO) = CInt(parti(0))
        outData(i, D_MESE) = CInt(parti(1))
        outData(i, D_AREA) = parti(2)
        outData(i, D_PRODOTTO) = parti(3)
        outData(i, D_CODICE) = parti(4)
        outData(i, D_DESCRIZ) = parti(5)
        outData(i, D_REGIONE) = parti(6)
        outData(i, D_PROVINCIA) = parti(7)
        outData(i, D_GIACINIZ) = acc(0)
        For idx = 0 To N_DIFF - 1
            outData(i, D_P0 + idx) = acc(1 + idx)
        Next idx
        outData(i, D_GIACFIN) = acc(1 + N_DIFF)
        outData(i, D_OMOG) = acc(1 + N_DIFF + 1)
        outData(i, D_GIACGG) = acc(1 + N_DIFF + 2)
    Next i

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' Pulisce TUTTO il vecchio contenuto dati, comprese le vecchie colonne
    ' formula (il vecchio layout arrivava fino alla colonna 43 = AR)
    wsConsolidato.Range(wsConsolidato.Cells(PRIMA_RIGA, 1), wsConsolidato.Cells(lastRow, 43)).Clear

    wsConsolidato.Range( _
        wsConsolidato.Cells(PRIMA_RIGA, 1), _
        wsConsolidato.Cells(PRIMA_RIGA + ordine.Count - 1, D_NCOLS) _
    ).Value = outData

    ScriviIntestazioniProvincia wsConsolidato, PRIMA_RIGA - 1

    Dim rFirst As Long, rLast As Long
    rFirst = PRIMA_RIGA
    rLast = PRIMA_RIGA + ordine.Count - 1
    ScriviFormuleProvincia wsConsolidato, rFirst, rLast

    ' Aggiorna pivot
    Dim lastPivotRow As Long
    lastPivotRow = wsConsolidato.Cells(wsConsolidato.Rows.Count, D_PRODOTTO).End(xlUp).Row
    Dim srcRange As String
    srcRange = "Consolidato!$A$4:$AP$" & lastPivotRow
    Dim ws2 As Worksheet
    Dim pt As PivotTable
    For Each ws2 In ThisWorkbook.Worksheets
        For Each pt In ws2.PivotTables
            On Error Resume Next
            pt.ChangePivotCache ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, SourceData:=srcRange)
            pt.RefreshTable
            On Error GoTo 0
        Next pt
    Next ws2

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False

    Dim msgExtra As String
    If n_senza_prov > 0 Then
        msgExtra = "Righe con Codice Sede non riconosciuto: " & n_senza_prov & vbCrLf
    End If

    MsgBox "Migrazione completata." & vbCrLf & vbCrLf & _
           "Righe storiche originarie: " & nRigheOriginarie & vbCrLf & _
           "Righe aggregate per provincia: " & ordine.Count & vbCrLf & _
           msgExtra & vbCrLf & _
           "Verifica ora i campi delle tabelle pivot: i riferimenti a 'Codice Sede'/'Sede' non " & _
           "sono piu' validi e vanno sostituiti con 'Descrizione Provincia'.", _
           vbInformation, "Migrazione completata"
End Sub


'==============================================================================
' MACRO: Filtra
'
' Estrae dal file di produzione selezionato le righe DC Pensioni con lo
' STESSO filtro usato da ImportaDatiMensili:
'   Prodotto Outcome riconducibile a "DC Pensioni" (foglio "prodotti",
'   colonna B) E (Totale Pervenuti > 0 OPPURE Totale Definiti > 0)
'
' A differenza di ImportaDatiMensili:
'   - NON scrive nel Consolidato di ThisWorkbook: crea un file xlsx
'     autonomo e a se stante
'   - NON calcola alcun differenziale rispetto al mese precedente: i
'     valori riportati sono esattamente quelli del file sorgente
'   - Le colonne AM:AR contengono le STESSE formule (XLOOKUP e
'     coefficienti) usate da ImportaDatiMensili nel Consolidato. Per
'     farle funzionare, il file include anche una copia dei fogli
'     "territorio" e "prodotti" (necessari alle formule XLOOKUP)
'
' Output: <SharePoint>\SIMP_Pensioni\DCPensioni\SIMP_DCPensioni_<AAAA>_<MM>.xlsx
'==============================================================================
Sub Filtra()

    Dim spRoot As String
    spRoot = GetSharePointRoot()
    If spRoot = "" Then Exit Sub
    If GetPythonExe() = "" Then Exit Sub

    Dim wsProdotti As Worksheet
    Set wsProdotti = ThisWorkbook.Sheets("prodotti")

    ' 1. SELEZIONE FILE
    Dim filePath As String
    filePath = Application.GetOpenFilename( _
        FileFilter:="File Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="Seleziona il file di produzione da cui estrarre DC Pensioni")
    If filePath = "False" Then
        MsgBox "Operazione annullata.", vbInformation, "Annullato"
        Exit Sub
    End If

    ' 2. VERIFICA SCRIPT PYTHON
    Dim scriptDir As String
    Dim scriptPath As String
    scriptDir = GetScriptDir()
    scriptPath = scriptDir & SCRIPT_NAME3
    AssicuraCartella scriptDir
    If Dir(scriptPath) = "" Then
        MsgBox "Script non trovato: " & scriptPath & vbCrLf & vbCrLf & _
               "Copiare filtra_dcpensioni.py in: " & scriptDir, _
               vbCritical, "Script mancante"
        Exit Sub
    End If

    ' 3. ESPORTA PRODOTTI IN CSV TEMPORANEO
    Application.StatusBar = "Esportazione prodotti..."
    Dim prodottiCsv As String
    prodottiCsv = GetTempDir() & "dc_prodotti_tmp3.csv"
    EsportaProdottiCsv wsProdotti, prodottiCsv

    ' 4. CARTELLA DI OUTPUT (Python vi scrive direttamente il file finale)
    Dim xlsxDir As String
    xlsxDir = GetBaseDir() & XLSX_SUBDIR & "\"
    AssicuraCartella xlsxDir

    ' 5. CHIAMA PYTHON: fa tutto (filtro, mese/anno, scrittura dati)
    Dim logPath As String
    Dim batPath As String
    Dim errPath As String
    logPath = GetTempDir() & "dc_filtra3.log"
    batPath = GetTempDir() & "dc_filtra3.bat"
    errPath = GetTempDir() & "dc_filtra3_err.log"

    If Dir(logPath) <> "" Then Kill logPath
    If Dir(errPath) <> "" Then Kill errPath

    ' Un backslash finale prima delle virgolette di chiusura confonde il
    ' parser degli argomenti di Windows (backslash "dispari" prima di una
    ' virgoletta = virgoletta letterale, non di chiusura): l'argomento
    ' passato a Python risulta corrotto (causa del WinError 123). Per
    ' l'argomento da riga di comando si usa quindi il percorso SENZA lo
    ' slash finale; internamente in VBA (AssicuraCartella, ecc.) resta
    ' invariato con lo slash.
    Dim xlsxDirArg As String
    xlsxDirArg = xlsxDir
    If Right(xlsxDirArg, 1) = "\" Then xlsxDirArg = Left(xlsxDirArg, Len(xlsxDirArg) - 1)

    Dim batLine As String
    batLine = "@echo off" & vbCrLf & _
              """" & GetPythonExe() & """ " & _
              """" & scriptPath & """ " & _
              """" & filePath & """ " & _
              """" & prodottiCsv & """ " & _
              """" & xlsxDirArg & """ " & _
              "1> """ & logPath & """ " & _
              "2> """ & errPath & """" & vbCrLf & _
              "if %ERRORLEVEL% NEQ 0 type """ & errPath & """ >> """ & logPath & """"

    Dim iFile As Integer
    iFile = FreeFile
    Open batPath For Output As #iFile
    Print #iFile, batLine
    Close #iFile

    Application.StatusBar = "Filtraggio con Python (attendere)..."
    Shell "cmd.exe /c """ & batPath & """", vbHide

    Dim t0 As Single
    Dim logContent As String
    Dim elapsed As Long
    t0 = Timer
    Do
        DoEvents
        Application.Wait Now + TimeValue("00:00:02")
        logContent = LeggiFile(logPath)
        elapsed = CLng(Timer - t0)
        Application.StatusBar = "Filtraggio Python... " & elapsed & "s"
        If Left(Trim(logContent), 2) = "OK" Then Exit Do
        If Left(Trim(logContent), 6) = "ERRORE" Then Exit Do
        If elapsed > 180 Then Exit Do
    Loop

    If Left(Trim(logContent), 2) <> "OK" Then
        MsgBox "Errore nello script Python:" & vbCrLf & Left(logContent, 600), _
               vbCritical, "Errore Python"
        GoTo CleanupTemp3
    End If

    ' Output formato: OK|n_tot|n_out|n_sca|cellA1|output_path (path vuoto se n_out=0)
    Dim parts() As String
    parts = Split(Trim(logContent), "|")
    If UBound(parts) < 4 Then
        MsgBox "Output Python non valido: " & logContent, vbCritical, "Errore"
        GoTo CleanupTemp3
    End If
    Dim iPulisci As Integer
    For iPulisci = LBound(parts) To UBound(parts)
        Do While Len(parts(iPulisci)) > 0 And _
                 (Right(parts(iPulisci), 1) = vbCr Or Right(parts(iPulisci), 1) = vbLf)
            parts(iPulisci) = Left(parts(iPulisci), Len(parts(iPulisci)) - 1)
        Loop
        parts(iPulisci) = Trim(parts(iPulisci))
    Next iPulisci
    Dim nTot As Long
    Dim nOut As Long
    Dim nSca As Long
    Dim cellA1 As String
    Dim outputXlsx As String
    nTot = CLng(parts(1))
    nOut = CLng(parts(2))
    nSca = CLng(parts(3))
    cellA1 = parts(4)
    If UBound(parts) >= 5 Then outputXlsx = parts(5) Else outputXlsx = ""

    If nOut = 0 Or outputXlsx = "" Then
        MsgBox "Nessuna riga DC Pensioni trovata (con Tot.Pervenuti > 0 o Tot.Definiti > 0)." & vbCrLf & _
               "Riferimento: " & cellA1 & vbCrLf & _
               "Righe totali nel file: " & nTot, vbInformation
        GoTo CleanupTemp3
    End If

    ' 6. APRI IL FILE SCRITTO DA PYTHON E AGGIUNGI FOGLI + FORMULE
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.StatusBar = "Aggiunta formule e fogli di supporto..."

    Dim tOpen As Single, tCopy As Single, tFormule As Single
    Dim tAutofit As Single, tSalva As Single, tChiudi As Single
    Dim tX As Single

    Dim tAttesaFile As Single
    tX = Timer
    Dim tentativi As Integer
    tentativi = 0
    Do While Dir(outputXlsx) = "" And tentativi < 15
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        tentativi = tentativi + 1
    Loop
    tAttesaFile = Timer - tX

    If Dir(outputXlsx) = "" Then
        MsgBox "Il file generato da Python non e' (ancora) visibile dopo " & tentativi & _
               " secondi di attesa:" & vbCrLf & outputXlsx & vbCrLf & vbCrLf & _
               "Verifica che antivirus o altri programmi non lo blocchino, poi riprova.", _
               vbCritical, "File non trovato"
        GoTo CleanupTemp3
    End If

    tX = Timer
    Dim wbOut As Workbook
    Set wbOut = ApriWorkbookRobusto(outputXlsx)
    If wbOut Is Nothing Then GoTo CleanupTemp3
    tOpen = Timer - tX

    tX = Timer
    ThisWorkbook.Sheets("territorio").Copy After:=wbOut.Sheets(wbOut.Sheets.Count)
    ThisWorkbook.Sheets("prodotti").Copy After:=wbOut.Sheets(wbOut.Sheets.Count)
    tCopy = Timer - tX

    Dim wsOut As Worksheet
    Set wsOut = wbOut.Sheets(1)

    Dim rFirst As Long
    Dim rLast As Long
    rFirst = 2
    rLast = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row

    Const F_AM As Integer = 39
    wsOut.Cells(1, F_AM - 1).Value = "Territorio"
    wsOut.Cells(1, F_AM).Value = "Coefficiente"
    wsOut.Cells(1, F_AM + 1).Value = "IndiceDeflusso"
    wsOut.Cells(1, F_AM + 2).Value = "GiacenzaOmog"
    wsOut.Cells(1, F_AM + 3).Value = "PervenutoOmog"
    wsOut.Cells(1, F_AM + 4).Value = "Aggregati"
    wsOut.Cells(1, F_AM + 5).Value = "Gestione"

    tX = Timer
    wsOut.Range(wsOut.Cells(rFirst, F_AM - 1), wsOut.Cells(rLast, F_AM - 1)).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(TEXT(RC8,""000000""),territorio!C4:C4,territorio!C3:C3),"""")"
    wsOut.Range(wsOut.Cells(rFirst, F_AM), wsOut.Cells(rLast, F_AM)).FormulaR1C1 = _
        "=IF(RC34 > 0,RC36/RC34,0)"
    wsOut.Range(wsOut.Cells(rFirst, F_AM + 1), wsOut.Cells(rLast, F_AM + 1)).FormulaR1C1 = _
        "=IF(RC22=0,0,RC34/RC22)"
    wsOut.Range(wsOut.Cells(rFirst, F_AM + 2), wsOut.Cells(rLast, F_AM + 2)).FormulaR1C1 = _
        "=RC35*RC39"
    wsOut.Range(wsOut.Cells(rFirst, F_AM + 3), wsOut.Cells(rLast, F_AM + 3)).FormulaR1C1 = _
        "=RC22*RC39"
    wsOut.Range(wsOut.Cells(rFirst, F_AM + 4), wsOut.Cells(rLast, F_AM + 4)).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(LEFT(RC4,4),prodotti!C3:C3,prodotti!C4:C4),"""")"
    wsOut.Range(wsOut.Cells(rFirst, F_AM + 5), wsOut.Cells(rLast, F_AM + 5)).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(LEFT(RC4,4),prodotti!C3:C3,prodotti!C5:C5),"""")"
    wsOut.Calculate
    tFormule = Timer - tX

    tX = Timer
    wsOut.Columns.AutoFit
    tAutofit = Timer - tX

    Dim tPivot As Single
    tX = Timer

    Dim lastColPivot As Long
    lastColPivot = F_AM + 5   ' AR

    Dim pvCache As PivotCache
    Set pvCache = wbOut.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(rLast, lastColPivot)))

    Dim wsPivot As Worksheet
    Set wsPivot = wbOut.Sheets.Add(After:=wsOut)
    wsPivot.Name = "Pivot"

    Dim pt As PivotTable
    Set pt = pvCache.CreatePivotTable( _
        TableDestination:=wsPivot.Cells(3, 1), _
        TableName:="PivotDCPensioni")

    With pt
        .PivotFields("Area").Orientation = xlRowField
        .PivotFields("Prodotto").Orientation = xlRowField
        .PivotFields("GiacIniz").Orientation = xlDataField
        .PivotFields("GiacFin").Orientation = xlDataField
        .PivotFields("Omog").Orientation = xlDataField
        .PivotFields("GiacenzaOmog").Orientation = xlDataField
        .PivotFields("PervenutoOmog").Orientation = xlDataField
    End With

    tPivot = Timer - tX

    tX = Timer
    wbOut.Save
    tSalva = Timer - tX

    tX = Timer
    wbOut.Close SaveChanges:=False
    tChiudi = Timer - tX

    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim mese As Integer
    Dim anno As Integer
    ParseMeseAnno cellA1, mese, anno

    MsgBox "Estrazione completata!" & vbCrLf & vbCrLf & _
           "Mese: " & mese & "/" & anno & vbCrLf & _
           "Righe totali nel file:  " & nTot & vbCrLf & _
           "Righe DC Pensioni:      " & nOut & vbCrLf & _
           "Righe scartate:         " & nSca & vbCrLf & vbCrLf & _
           "XLSX: " & outputXlsx & vbCrLf & vbCrLf & _
           "--- Tempi (secondi) ---" & vbCrLf & _
           "Python (filtro+scrittura dati): ~" & elapsed & vbCrLf & _
           "Attesa visibilita' file:        " & Format(tAttesaFile, "0.0") & vbCrLf & _
           "Apertura file:                  " & Format(tOpen, "0.0") & vbCrLf & _
           "Copia fogli territorio/prodotti: " & Format(tCopy, "0.0") & vbCrLf & _
           "Scrittura formule + calcolo:     " & Format(tFormule, "0.0") & vbCrLf & _
           "Autofit colonne:                 " & Format(tAutofit, "0.0") & vbCrLf & _
           "Creazione pivot:                 " & Format(tPivot, "0.0") & vbCrLf & _
           "Salvataggio:                     " & Format(tSalva, "0.0") & vbCrLf & _
           "Chiusura:                        " & Format(tChiudi, "0.0"), _
           vbInformation, "OK"

CleanupTemp3:
    On Error Resume Next
    If Dir(prodottiCsv) <> "" Then Kill prodottiCsv
    If Dir(batPath) <> "" Then Kill batPath
    If Dir(errPath) <> "" Then Kill errPath
    On Error GoTo 0
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


'==============================================================================
' MACRO: FiltraRegione
'
' Come "Filtra", ma i dati vengono AGGREGATI a livello di Regione.
' Output: <SharePoint>\SIMP_Pensioni\DCPensioni\SIMP_DCPensioni_Regione_<AAAA>_<MM>.xlsx
'==============================================================================
Sub FiltraRegione()

    Dim spRoot As String
    spRoot = GetSharePointRoot()
    If spRoot = "" Then Exit Sub
    If GetPythonExe() = "" Then Exit Sub

    Dim wsProdotti As Worksheet
    Set wsProdotti = ThisWorkbook.Sheets("prodotti")

    Const AGG_PRODOTTO  As Integer = 4
    Const AGG_TOTPERV   As Integer = 20
    Const AGG_TOTDEF    As Integer = 32
    Const AGG_GIACFIN   As Integer = 33
    Const AGG_OMOG      As Integer = 34
    Const AGG_NCOLS     As Integer = 35
    Const AGG_AM        As Integer = 36

    Dim filePath As String
    filePath = Application.GetOpenFilename( _
        FileFilter:="File Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="Seleziona il file di produzione da cui estrarre DC Pensioni per Regione")
    If filePath = "False" Then
        MsgBox "Operazione annullata.", vbInformation, "Annullato"
        Exit Sub
    End If

    Dim scriptDir As String
    Dim scriptPath As String
    scriptDir = GetScriptDir()
    scriptPath = scriptDir & SCRIPT_NAME4
    AssicuraCartella scriptDir
    If Dir(scriptPath) = "" Then
        MsgBox "Script non trovato: " & scriptPath & vbCrLf & vbCrLf & _
               "Copiare filtra_dcpensioni_regione.py in: " & scriptDir, _
               vbCritical, "Script mancante"
        Exit Sub
    End If

    Application.StatusBar = "Esportazione prodotti..."
    Dim prodottiCsv As String
    prodottiCsv = GetTempDir() & "dc_prodotti_tmp4.csv"
    EsportaProdottiCsv wsProdotti, prodottiCsv

    Dim xlsxDir As String
    xlsxDir = GetBaseDir() & XLSX_SUBDIR & "\"
    AssicuraCartella xlsxDir

    Dim xlsxDirArg As String
    xlsxDirArg = xlsxDir
    If Right(xlsxDirArg, 1) = "\" Then xlsxDirArg = Left(xlsxDirArg, Len(xlsxDirArg) - 1)

    Dim logPath As String
    Dim batPath As String
    Dim errPath As String
    logPath = GetTempDir() & "dc_filtra4.log"
    batPath = GetTempDir() & "dc_filtra4.bat"
    errPath = GetTempDir() & "dc_filtra4_err.log"

    If Dir(logPath) <> "" Then Kill logPath
    If Dir(errPath) <> "" Then Kill errPath

    Dim batLine As String
    batLine = "@echo off" & vbCrLf & _
              """" & GetPythonExe() & """ " & _
              """" & scriptPath & """ " & _
              """" & filePath & """ " & _
              """" & prodottiCsv & """ " & _
              """" & xlsxDirArg & """ " & _
              "1> """ & logPath & """ " & _
              "2> """ & errPath & """" & vbCrLf & _
              "if %ERRORLEVEL% NEQ 0 type """ & errPath & """ >> """ & logPath & """"

    Dim iFile As Integer
    iFile = FreeFile
    Open batPath For Output As #iFile
    Print #iFile, batLine
    Close #iFile

    Application.StatusBar = "Filtraggio e aggregazione con Python (attendere)..."
    Shell "cmd.exe /c """ & batPath & """", vbHide

    Dim t0 As Single
    Dim logContent As String
    Dim elapsed As Long
    t0 = Timer
    Do
        DoEvents
        Application.Wait Now + TimeValue("00:00:02")
        logContent = LeggiFile(logPath)
        elapsed = CLng(Timer - t0)
        Application.StatusBar = "Filtraggio Python... " & elapsed & "s"
        If Left(Trim(logContent), 2) = "OK" Then Exit Do
        If Left(Trim(logContent), 6) = "ERRORE" Then Exit Do
        If elapsed > 180 Then Exit Do
    Loop

    If Left(Trim(logContent), 2) <> "OK" Then
        MsgBox "Errore nello script Python:" & vbCrLf & Left(logContent, 600), _
               vbCritical, "Errore Python"
        GoTo CleanupTemp4
    End If

    Dim parts() As String
    parts = Split(Trim(logContent), "|")
    If UBound(parts) < 5 Then
        MsgBox "Output Python non valido: " & logContent, vbCritical, "Errore"
        GoTo CleanupTemp4
    End If
    Dim iPulisci As Integer
    For iPulisci = LBound(parts) To UBound(parts)
        Do While Len(parts(iPulisci)) > 0 And _
                 (Right(parts(iPulisci), 1) = vbCr Or Right(parts(iPulisci), 1) = vbLf)
            parts(iPulisci) = Left(parts(iPulisci), Len(parts(iPulisci)) - 1)
        Loop
        parts(iPulisci) = Trim(parts(iPulisci))
    Next iPulisci

    Dim nTot As Long
    Dim nOut As Long
    Dim nSca As Long
    Dim nGruppi As Long
    Dim cellA1 As String
    Dim outputXlsx As String
    nTot = CLng(parts(1))
    nOut = CLng(parts(2))
    nSca = CLng(parts(3))
    nGruppi = CLng(parts(4))
    cellA1 = parts(5)
    If UBound(parts) >= 6 Then outputXlsx = parts(6) Else outputXlsx = ""

    If nGruppi = 0 Or outputXlsx = "" Then
        MsgBox "Nessuna riga DC Pensioni trovata (con Tot.Pervenuti > 0 o Tot.Definiti > 0)." & vbCrLf & _
               "Riferimento: " & cellA1 & vbCrLf & _
               "Righe totali nel file: " & nTot, vbInformation
        GoTo CleanupTemp4
    End If

    Dim tentativi As Integer
    tentativi = 0
    Do While Dir(outputXlsx) = "" And tentativi < 15
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        tentativi = tentativi + 1
    Loop
    If Dir(outputXlsx) = "" Then
        MsgBox "Il file generato da Python non e' (ancora) visibile dopo " & tentativi & _
               " secondi di attesa:" & vbCrLf & outputXlsx, vbCritical, "File non trovato"
        GoTo CleanupTemp4
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.StatusBar = "Aggiunta formule e foglio prodotti..."

    Dim wbOut As Workbook
    Set wbOut = ApriWorkbookRobusto(outputXlsx)
    If wbOut Is Nothing Then GoTo CleanupTemp4

    ThisWorkbook.Sheets("prodotti").Copy After:=wbOut.Sheets(wbOut.Sheets.Count)

    Dim wsOut As Worksheet
    Set wsOut = wbOut.Sheets(1)

    Dim rFirst As Long
    Dim rLast As Long
    rFirst = 2
    rLast = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row

    wsOut.Cells(1, AGG_AM).Value = "Coefficiente"
    wsOut.Cells(1, AGG_AM + 1).Value = "IndiceDeflusso"
    wsOut.Cells(1, AGG_AM + 2).Value = "GiacenzaOmog"
    wsOut.Cells(1, AGG_AM + 3).Value = "PervenutoOmog"
    wsOut.Cells(1, AGG_AM + 4).Value = "Aggregati"
    wsOut.Cells(1, AGG_AM + 5).Value = "Gestione"

    wsOut.Range(wsOut.Cells(rFirst, AGG_AM), wsOut.Cells(rLast, AGG_AM)).FormulaR1C1 = _
        "=IF(RC" & AGG_TOTDEF & " > 0,RC" & AGG_OMOG & "/RC" & AGG_TOTDEF & ",0)"
    wsOut.Range(wsOut.Cells(rFirst, AGG_AM + 1), wsOut.Cells(rLast, AGG_AM + 1)).FormulaR1C1 = _
        "=IF(RC" & AGG_TOTPERV & "=0,0,RC" & AGG_TOTDEF & "/RC" & AGG_TOTPERV & ")"
    wsOut.Range(wsOut.Cells(rFirst, AGG_AM + 2), wsOut.Cells(rLast, AGG_AM + 2)).FormulaR1C1 = _
        "=RC" & AGG_GIACFIN & "*RC" & AGG_AM
    wsOut.Range(wsOut.Cells(rFirst, AGG_AM + 3), wsOut.Cells(rLast, AGG_AM + 3)).FormulaR1C1 = _
        "=RC" & AGG_TOTPERV & "*RC" & AGG_AM
    wsOut.Range(wsOut.Cells(rFirst, AGG_AM + 4), wsOut.Cells(rLast, AGG_AM + 4)).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(LEFT(RC" & AGG_PRODOTTO & ",4),prodotti!C3:C3,prodotti!C4:C4),"""")"
    wsOut.Range(wsOut.Cells(rFirst, AGG_AM + 5), wsOut.Cells(rLast, AGG_AM + 5)).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(LEFT(RC" & AGG_PRODOTTO & ",4),prodotti!C3:C3,prodotti!C5:C5),"""")"

    wsOut.Calculate
    wsOut.Columns.AutoFit

    wbOut.Save
    wbOut.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim mese As Integer
    Dim anno As Integer
    ParseMeseAnno cellA1, mese, anno

    MsgBox "Estrazione ed aggregazione per Regione completata!" & vbCrLf & vbCrLf & _
           "Mese: " & mese & "/" & anno & vbCrLf & _
           "Righe totali nel file:  " & nTot & vbCrLf & _
           "Righe DC Pensioni:      " & nOut & vbCrLf & _
           "Righe scartate:         " & nSca & vbCrLf & _
           "Righe aggregate:        " & nGruppi & vbCrLf & vbCrLf & _
           "XLSX: " & outputXlsx, vbInformation, "OK"

CleanupTemp4:
    On Error Resume Next
    If Dir(prodottiCsv) <> "" Then Kill prodottiCsv
    If Dir(batPath) <> "" Then Kill batPath
    If Dir(errPath) <> "" Then Kill errPath
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


'==============================================================================
' MACRO: FiltraProvincia
'
' Come "FiltraRegione", ma aggregato a livello di Descrizione Provincia.
' Output: <SharePoint>\SIMP_Pensioni\DCPensioni\SIMP_DCPensioni_Provincia_<AAAA>_<MM>.xlsx
'==============================================================================
Sub FiltraProvincia()

    Dim spRoot As String
    spRoot = GetSharePointRoot()
    If spRoot = "" Then Exit Sub
    If GetPythonExe() = "" Then Exit Sub

    Dim wsTerritorio As Worksheet
    Set wsTerritorio = ThisWorkbook.Sheets("territorio")

    Const PRV_PRODOTTO  As Integer = 4
    Const PRV_TOTPERV   As Integer = 21
    Const PRV_TOTDEF    As Integer = 33
    Const PRV_GIACFIN   As Integer = 34
    Const PRV_OMOG      As Integer = 35
    Const PRV_NCOLS     As Integer = 36
    Const PRV_AM        As Integer = 37

    Dim filePath As String
    filePath = Application.GetOpenFilename( _
        FileFilter:="File Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="Seleziona il file di produzione da cui estrarre DC Pensioni per Provincia")
    If filePath = "False" Then
        MsgBox "Operazione annullata.", vbInformation, "Annullato"
        Exit Sub
    End If

    Dim scriptDir As String
    Dim scriptPath As String
    scriptDir = GetScriptDir()
    scriptPath = scriptDir & SCRIPT_NAME5
    AssicuraCartella scriptDir
    If Dir(scriptPath) = "" Then
        MsgBox "Script non trovato: " & scriptPath & vbCrLf & vbCrLf & _
               "Copiare filtra_dcpensioni_provincia.py in: " & scriptDir, _
               vbCritical, "Script mancante"
        Exit Sub
    End If

    Application.StatusBar = "Esportazione territorio..."
    Dim territorioCsv As String
    territorioCsv = GetTempDir() & "dc_territorio_tmp5.csv"
    EsportaFoglioCsv wsTerritorio, territorioCsv

    Dim xlsxDir As String
    xlsxDir = GetBaseDir() & XLSX_SUBDIR & "\"
    AssicuraCartella xlsxDir

    Dim xlsxDirArg As String
    xlsxDirArg = xlsxDir
    If Right(xlsxDirArg, 1) = "\" Then xlsxDirArg = Left(xlsxDirArg, Len(xlsxDirArg) - 1)

    Dim logPath As String
    Dim batPath As String
    Dim errPath As String
    logPath = GetTempDir() & "dc_filtra5.log"
    batPath = GetTempDir() & "dc_filtra5.bat"
    errPath = GetTempDir() & "dc_filtra5_err.log"

    If Dir(logPath) <> "" Then Kill logPath
    If Dir(errPath) <> "" Then Kill errPath

    Dim batLine As String
    batLine = "@echo off" & vbCrLf & _
              """" & GetPythonExe() & """ " & _
              """" & scriptPath & """ " & _
              """" & filePath & """ " & _
              """" & territorioCsv & """ " & _
              """" & xlsxDirArg & """ " & _
              "1> """ & logPath & """ " & _
              "2> """ & errPath & """" & vbCrLf & _
              "if %ERRORLEVEL% NEQ 0 type """ & errPath & """ >> """ & logPath & """"

    Dim iFile As Integer
    iFile = FreeFile
    Open batPath For Output As #iFile
    Print #iFile, batLine
    Close #iFile

    Application.StatusBar = "Filtraggio e aggregazione per Provincia con Python (attendere)..."
    Shell "cmd.exe /c """ & batPath & """", vbHide

    Dim t0 As Single
    Dim logContent As String
    Dim elapsed As Long
    t0 = Timer
    Do
        DoEvents
        Application.Wait Now + TimeValue("00:00:02")
        logContent = LeggiFile(logPath)
        elapsed = CLng(Timer - t0)
        Application.StatusBar = "Filtraggio Python... " & elapsed & "s"
        If Left(Trim(logContent), 2) = "OK" Then Exit Do
        If Left(Trim(logContent), 6) = "ERRORE" Then Exit Do
        If elapsed > 180 Then Exit Do
    Loop

    If Left(Trim(logContent), 2) <> "OK" Then
        MsgBox "Errore nello script Python:" & vbCrLf & Left(logContent, 600), _
               vbCritical, "Errore Python"
        GoTo CleanupTemp5
    End If

    Dim parts() As String
    parts = Split(Trim(logContent), "|")
    If UBound(parts) < 5 Then
        MsgBox "Output Python non valido: " & logContent, vbCritical, "Errore"
        GoTo CleanupTemp5
    End If
    Dim iPulisci As Integer
    For iPulisci = LBound(parts) To UBound(parts)
        Do While Len(parts(iPulisci)) > 0 And _
                 (Right(parts(iPulisci), 1) = vbCr Or Right(parts(iPulisci), 1) = vbLf)
            parts(iPulisci) = Left(parts(iPulisci), Len(parts(iPulisci)) - 1)
        Loop
        parts(iPulisci) = Trim(parts(iPulisci))
    Next iPulisci

    Dim nTot As Long
    Dim nOut As Long
    Dim nSca As Long
    Dim nGruppi As Long
    Dim cellA1 As String
    Dim outputXlsx As String
    nTot = CLng(parts(1))
    nOut = CLng(parts(2))
    nSca = CLng(parts(3))
    nGruppi = CLng(parts(4))
    cellA1 = parts(5)
    If UBound(parts) >= 6 Then outputXlsx = parts(6) Else outputXlsx = ""

    If nGruppi = 0 Or outputXlsx = "" Then
        MsgBox "Nessuna riga DC Pensioni trovata (con Tot.Pervenuti > 0 o Tot.Definiti > 0)." & vbCrLf & _
               "Riferimento: " & cellA1 & vbCrLf & _
               "Righe totali nel file: " & nTot, vbInformation
        GoTo CleanupTemp5
    End If

    Dim tentativi As Integer
    tentativi = 0
    Do While Dir(outputXlsx) = "" And tentativi < 15
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        tentativi = tentativi + 1
    Loop
    If Dir(outputXlsx) = "" Then
        MsgBox "Il file generato da Python non e' (ancora) visibile dopo " & tentativi & _
               " secondi di attesa:" & vbCrLf & outputXlsx, vbCritical, "File non trovato"
        GoTo CleanupTemp5
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.StatusBar = "Aggiunta formule e foglio prodotti..."

    Dim wbOut As Workbook
    Set wbOut = ApriWorkbookRobusto(outputXlsx)
    If wbOut Is Nothing Then GoTo CleanupTemp5

    ThisWorkbook.Sheets("prodotti").Copy After:=wbOut.Sheets(wbOut.Sheets.Count)

    Dim wsOut As Worksheet
    Set wsOut = wbOut.Sheets(1)

    Dim rFirst As Long
    Dim rLast As Long
    rFirst = 2
    rLast = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row

    wsOut.Cells(1, PRV_AM).Value = "Coefficiente"
    wsOut.Cells(1, PRV_AM + 1).Value = "IndiceDeflusso"
    wsOut.Cells(1, PRV_AM + 2).Value = "GiacenzaOmog"
    wsOut.Cells(1, PRV_AM + 3).Value = "PervenutoOmog"
    wsOut.Cells(1, PRV_AM + 4).Value = "Aggregati"
    wsOut.Cells(1, PRV_AM + 5).Value = "Gestione"

    wsOut.Range(wsOut.Cells(rFirst, PRV_AM), wsOut.Cells(rLast, PRV_AM)).FormulaR1C1 = _
        "=IF(RC" & PRV_TOTDEF & " > 0,RC" & PRV_OMOG & "/RC" & PRV_TOTDEF & ",0)"
    wsOut.Range(wsOut.Cells(rFirst, PRV_AM + 1), wsOut.Cells(rLast, PRV_AM + 1)).FormulaR1C1 = _
        "=IF(RC" & PRV_TOTPERV & "=0,0,RC" & PRV_TOTDEF & "/RC" & PRV_TOTPERV & ")"
    wsOut.Range(wsOut.Cells(rFirst, PRV_AM + 2), wsOut.Cells(rLast, PRV_AM + 2)).FormulaR1C1 = _
        "=RC" & PRV_GIACFIN & "*RC" & PRV_AM
    wsOut.Range(wsOut.Cells(rFirst, PRV_AM + 3), wsOut.Cells(rLast, PRV_AM + 3)).FormulaR1C1 = _
        "=RC" & PRV_TOTPERV & "*RC" & PRV_AM
    wsOut.Range(wsOut.Cells(rFirst, PRV_AM + 4), wsOut.Cells(rLast, PRV_AM + 4)).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(LEFT(RC" & PRV_PRODOTTO & ",4),prodotti!C3:C3,prodotti!C4:C4),"""")"
    wsOut.Range(wsOut.Cells(rFirst, PRV_AM + 5), wsOut.Cells(rLast, PRV_AM + 5)).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(LEFT(RC" & PRV_PRODOTTO & ",4),prodotti!C3:C3,prodotti!C5:C5),"""")"

    wsOut.Calculate
    wsOut.Columns.AutoFit

    wbOut.Save
    wbOut.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim mese As Integer
    Dim anno As Integer
    ParseMeseAnno cellA1, mese, anno

    MsgBox "Estrazione ed aggregazione per Provincia completata!" & vbCrLf & vbCrLf & _
           "Mese: " & mese & "/" & anno & vbCrLf & _
           "Righe totali nel file:  " & nTot & vbCrLf & _
           "Righe DC Pensioni:      " & nOut & vbCrLf & _
           "Righe scartate:         " & nSca & vbCrLf & _
           "Righe aggregate:        " & nGruppi & vbCrLf & vbCrLf & _
           "XLSX: " & outputXlsx, vbInformation, "OK"

CleanupTemp5:
    On Error Resume Next
    If Dir(territorioCsv) <> "" Then Kill territorioCsv
    If Dir(batPath) <> "" Then Kill batPath
    If Dir(errPath) <> "" Then Kill errPath
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


'==============================================================================
Sub EliminaMese()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Consolidato")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, D_PRODOTTO).End(xlUp).Row
    If lastRow < PRIMA_RIGA Then
        MsgBox "Nessun dato nel Consolidato.", vbInformation
        Exit Sub
    End If

    Dim mesiSet As Object
    Set mesiSet = CreateObject("Scripting.Dictionary")
    Dim r As Long
    Dim rAnno As Integer
    Dim rMese As Integer
    For r = PRIMA_RIGA To lastRow
        On Error Resume Next
        rAnno = CInt(ws.Cells(r, D_ANNO).Value)
        rMese = CInt(ws.Cells(r, D_MESE).Value)
        On Error GoTo 0
        If rAnno > 0 And rMese >= 1 And rMese <= 12 Then
            Dim tag As String
            tag = rMese & "/" & rAnno
            If Not mesiSet.Exists(tag) Then mesiSet.Add tag, Array(rMese, rAnno)
        End If
    Next r

    If mesiSet.Count = 0 Then
        MsgBox "Nessun mese identificabile.", vbInformation
        Exit Sub
    End If

    Dim lista As String
    Dim k As Variant
    For Each k In mesiSet.keys
        lista = lista & k & vbCrLf
    Next k

    Dim sel As String
    sel = InputBox("Mesi presenti:" & vbCrLf & lista & vbCrLf & _
                   "Inserisci il mese da eliminare (M/AAAA):", "Elimina mese")
    If Trim(sel) = "" Then Exit Sub
    If Not mesiSet.Exists(Trim(sel)) Then
        MsgBox """" & sel & """ non trovato.", vbExclamation
        Exit Sub
    End If

    Dim conta As Long
    Dim me2 As Integer
    Dim an2 As Integer
    me2 = mesiSet(Trim(sel))(0)
    an2 = mesiSet(Trim(sel))(1)
    conta = 0
    For r = PRIMA_RIGA To lastRow
        On Error Resume Next
        rAnno = CInt(ws.Cells(r, D_ANNO).Value)
        rMese = CInt(ws.Cells(r, D_MESE).Value)
        On Error GoTo 0
        If rAnno = an2 And rMese = me2 Then conta = conta + 1
    Next r

    If MsgBox("Eliminare " & conta & " righe del mese " & sel & "?", _
              vbYesNo + vbExclamation, "Conferma") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    EliminaRigheMese ws, me2, an2
    Application.ScreenUpdating = True
    MsgBox "Eliminate " & conta & " righe del mese " & sel & ".", vbInformation
End Sub


'==============================================================================
Sub InstallaDipendenze()
    Dim py As String
    py = GetPythonExe()
    If py = "" Then Exit Sub
    If MsgBox("Verranno installati python-calamine e openpyxl." & vbCrLf & "Continuare?", _
              vbYesNo + vbQuestion, "Installa dipendenze") = vbNo Then Exit Sub
    Shell "cmd.exe /c " & py & " -m pip install python-calamine openpyxl && pause", vbNormalFocus
End Sub


'==============================================================================
' Cartella temporanea per-utente (Environ$("TEMP")): MAI un percorso fisso,
' perche' conterrebbe lo username di chi ha scritto la macro e non
' esisterebbe sul PC di un altro utente.
Private Function GetTempDir() As String
    Dim t As String
    t = Environ$("TEMP")
    If t = "" Then t = Environ$("TMP")
    If Right(t, 1) <> "\" Then t = t & "\"
    GetTempDir = t
End Function

' Risolve (con cache) e restituisce il percorso dell'interprete Python da
' usare. Analogo a GetSharePointRoot(): niente piu' percorso fisso con lo
' username in chiaro (non funzionerebbe su un altro PC/utente), individuato
' invece a runtime cercando "py" (il launcher ufficiale, preferito perche'
' sceglie da solo la versione installata) o, in mancanza, "python" nel PATH
' di sistema. Se l'autorilevamento fallisse, valorizza PYTHON_EXE_OVERRIDE.
Private Function GetPythonExe() As String
    If mPythonExeResolved Then
        GetPythonExe = mPythonExe
        Exit Function
    End If

    Dim p As String
    If Len(PYTHON_EXE_OVERRIDE) > 0 Then
        p = PYTHON_EXE_OVERRIDE
    Else
        p = RilevaPythonExe()
    End If

    If p = "" Then
        MsgBox "Impossibile individuare automaticamente l'interprete Python (py.exe o python.exe) " & _
               "nel PATH di sistema." & vbCrLf & vbCrLf & _
               "Installa Python (da python.org o dal Microsoft Store) assicurandoti che l'opzione " & _
               """Add python.exe to PATH""/""Aggiungi al PATH"" sia selezionata durante l'installazione." & vbCrLf & vbCrLf & _
               "In alternativa, valorizza la costante PYTHON_EXE_OVERRIDE in cima al modulo con il " & _
               "percorso completo di python.exe.", _
               vbCritical, "Python non trovato"
    End If

    mPythonExe = p
    mPythonExeResolved = True
    GetPythonExe = p
End Function

' Cerca "py" (launcher ufficiale) e, in mancanza, "python" nel PATH di
' sistema tramite il comando "where". Restituisce stringa vuota se nessuno
' dei due e' installato/nel PATH: il chiamante (GetPythonExe) gestisce il
' fallback mostrando l'errore.
Private Function RilevaPythonExe() As String
    On Error GoTo ErrHandler

    Dim batPath As String
    Dim outPath As String
    batPath = GetTempDir() & "dc_python_lookup.bat"
    outPath = GetTempDir() & "dc_python_lookup.txt"

    On Error Resume Next
    If Dir(batPath) <> "" Then Kill batPath
    If Dir(outPath) <> "" Then Kill outPath
    On Error GoTo ErrHandler

    Dim iFile As Integer
    iFile = FreeFile
    Open batPath For Output As #iFile
    Print #iFile, "@echo off"
    Print #iFile, "(where py || where python) > """ & outPath & """ 2>nul"
    Close #iFile

    Shell "cmd.exe /c """ & batPath & """", vbHide

    Dim t0 As Single
    t0 = Timer
    Do While Dir(outPath) = "" And Timer - t0 < 10
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Loop

    Dim risultato As String
    risultato = LeggiFile(outPath)
    ' "where" puo' restituire piu' righe (piu' versioni installate): prendi solo la prima
    If InStr(risultato, vbCrLf) > 0 Then risultato = Left(risultato, InStr(risultato, vbCrLf) - 1)
    risultato = Trim(risultato)

    On Error Resume Next
    If Dir(batPath) <> "" Then Kill batPath
    If Dir(outPath) <> "" Then Kill outPath
    On Error GoTo 0

    RilevaPythonExe = risultato
    Exit Function

ErrHandler:
    RilevaPythonExe = ""
End Function


'==============================================================================
' Risolve (con cache) e restituisce la cartella script sulla SharePoint
' sincronizzata: <radice SharePoint>\Scripts\
Private Function GetScriptDir() As String
    Dim root As String
    root = GetSharePointRoot()
    If root = "" Then
        GetScriptDir = ""
    Else
        GetScriptDir = root & SCRIPT_SUBDIR
    End If
End Function

' Risolve (con cache) e restituisce la cartella dati sulla SharePoint
' sincronizzata: <radice SharePoint>\SIMP_Pensioni\
Private Function GetBaseDir() As String
    Dim root As String
    root = GetSharePointRoot()
    If root = "" Then
        GetBaseDir = ""
    Else
        GetBaseDir = root & DATA_SUBDIR
    End If
End Function

' Individua (una sola volta per sessione, poi in cache) la cartella locale
' di sincronizzazione OneDrive/SharePoint da usare come radice per script e
' dati. Percorso SEMPRE locale (mai un URL), quindi utilizzabile con
' Dir()/MkDir()/Open anche quando ThisWorkbook e' aperto direttamente da
' SharePoint (dove ThisWorkbook.Path restituirebbe un URL).
Private Function GetSharePointRoot() As String
    If mSharePointResolved Then
        GetSharePointRoot = mSharePointRoot
        Exit Function
    End If

    Dim root As String
    If Len(SHAREPOINT_ROOT_OVERRIDE) > 0 Then
        root = SHAREPOINT_ROOT_OVERRIDE
    Else
        root = RilevaCartellaSharePointSincronizzata(SHAREPOINT_SITE_MATCH, SHAREPOINT_LIB_MATCH)
    End If

    If Len(root) = 0 Or Dir(root, vbDirectory) = "" Then
        MsgBox "Impossibile individuare automaticamente la cartella SharePoint sincronizzata." & vbCrLf & vbCrLf & _
               "Verifica che:" & vbCrLf & _
               " - OneDrive sia in esecuzione e la libreria SharePoint desiderata sia " & _
               "effettivamente sincronizzata (non solo ""disponibile online"")" & vbCrLf & _
               " - le costanti SHAREPOINT_SITE_MATCH / SHAREPOINT_LIB_MATCH in cima al modulo " & _
               "corrispondano a una porzione dell'indirizzo del sito/libreria SharePoint" & vbCrLf & vbCrLf & _
               "In alternativa, valorizza la costante SHAREPOINT_ROOT_OVERRIDE con il percorso " & _
               "locale della cartella sincronizzata (es. C:\Users\<utente>\<Azienda>\<Sito> - <Libreria>\).", _
               vbCritical, "Cartella SharePoint non trovata"
        mSharePointRoot = ""
        mSharePointResolved = True
        GetSharePointRoot = ""
        Exit Function
    End If

    If Right(root, 1) <> "\" Then root = root & "\"
    mSharePointRoot = root
    mSharePointResolved = True
    GetSharePointRoot = root
End Function

' Interroga (via PowerShell) il registro di Windows dove OneDrive tiene
' traccia di ogni libreria SharePoint sincronizzata (indirizzo del sito +
' percorso locale corrispondente), e restituisce il percorso locale della
' prima libreria il cui indirizzo contiene sia siteMatch sia libMatch
' (case-insensitive). Restituisce stringa vuota se non trovata o in caso
' di errore: il chiamante (GetSharePointRoot) gestisce il fallback.
Private Function RilevaCartellaSharePointSincronizzata(siteMatch As String, libMatch As String) As String
    On Error GoTo ErrHandler

    Dim psPath As String
    Dim outPath As String
    psPath = GetTempDir() & "dc_sp_lookup.ps1"
    outPath = GetTempDir() & "dc_sp_lookup.txt"

    On Error Resume Next
    If Dir(outPath) <> "" Then Kill outPath
    If Dir(psPath) <> "" Then Kill psPath
    On Error GoTo ErrHandler

    Dim ps As String
    ps = "$ErrorActionPreference = 'SilentlyContinue'" & vbCrLf & _
         "$rows = Get-ItemProperty 'HKCU:\Software\SyncEngines\Providers\OneDrive\*'" & vbCrLf & _
         "$m = $rows | Where-Object { $_.UrlNamespace -like '*" & siteMatch & "*' -and $_.UrlNamespace -like '*" & libMatch & "*' } | Select-Object -First 1" & vbCrLf & _
         "if ($m) { [System.IO.File]::WriteAllText('" & outPath & "', $m.MountPoint) }"

    Dim iFile As Integer
    iFile = FreeFile
    Open psPath For Output As #iFile
    Print #iFile, ps
    Close #iFile

    Shell "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & psPath & """", vbHide

    Dim t0 As Single
    t0 = Timer
    Do While Dir(outPath) = "" And Timer - t0 < 15
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Loop

    Dim risultato As String
    risultato = Trim(LeggiFile(outPath))

    On Error Resume Next
    If Dir(psPath) <> "" Then Kill psPath
    If Dir(outPath) <> "" Then Kill outPath
    On Error GoTo 0

    RilevaCartellaSharePointSincronizzata = risultato
    Exit Function

ErrHandler:
    RilevaCartellaSharePointSincronizzata = ""
End Function


'==============================================================================
' Crea, se necessario, tutti i livelli mancanti di un percorso di cartelle.
Private Sub AssicuraCartella(percorso As String)
    Dim p As String
    p = percorso
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)

    Dim parti() As String
    parti = Split(p, "\")
    If UBound(parti) < 1 Then Exit Sub

    Dim cur As String
    Dim i As Integer
    cur = parti(0) ' es. "C:"
    For i = 1 To UBound(parti)
        cur = cur & "\" & parti(i)
        If Dir(cur, vbDirectory) = "" Then
            On Error Resume Next
            MkDir cur
            On Error GoTo 0
        End If
    Next i
End Sub

Private Sub EsportaProdottiCsv(ws As Worksheet, outPath As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim iFile As Integer
    iFile = FreeFile
    Open outPath For Output As #iFile
    Print #iFile, "Prodotto;Competenza"
    Dim r As Long
    For r = 2 To lastRow
        Dim p As String
        Dim comp As String
        p = Trim(CStr(ws.Cells(r, 1).Value))
        comp = Trim(CStr(ws.Cells(r, 2).Value))
        If p <> "" Then Print #iFile, p & ";" & comp
    Next r
    Close #iFile
End Sub

' Esporta l'intero foglio (righe/colonne usate) in CSV punto e virgola,
' intestazioni comprese. Usata da ImportaDatiMensili e FiltraProvincia per
' esportare il foglio "territorio" (Codice Sede -> Descrizione Provincia).
Private Sub EsportaFoglioCsv(ws As Worksheet, outPath As String)
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim iFile As Integer
    iFile = FreeFile
    Open outPath For Output As #iFile
    Dim r As Long
    Dim c As Long
    Dim riga As String
    For r = 1 To lastRow
        riga = ""
        For c = 1 To lastCol
            If c > 1 Then riga = riga & ";"
            riga = riga & Trim(CStr(ws.Cells(r, c).Value))
        Next c
        Print #iFile, riga
    Next r
    Close #iFile
End Sub

' Apre un workbook con qualche tentativo di retry.
Private Function ApriWorkbookRobusto(percorso As String) As Workbook
    Dim wb As Workbook
    Dim tentativi As Integer
    Dim ultimoErrNum As Long
    Dim ultimoErrDesc As String
    tentativi = 0
    Do
        Set wb = Nothing
        On Error Resume Next
        Set wb = Workbooks.Open(percorso)
        ultimoErrNum = Err.Number
        ultimoErrDesc = Err.Description
        On Error GoTo 0
        If Not wb Is Nothing Then Exit Do
        tentativi = tentativi + 1
        If tentativi >= 5 Then Exit Do
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    If wb Is Nothing Then
        MsgBox "Impossibile aprire il file dopo " & tentativi & " tentativi:" & vbCrLf & _
               percorso & vbCrLf & vbCrLf & _
               "Errore VBA: " & ultimoErrNum & " - " & ultimoErrDesc, _
               vbCritical, "Errore apertura file"
    End If
    Set ApriWorkbookRobusto = wb
End Function

Private Function LeggiFile(path As String) As String
    If Dir(path) = "" Then LeggiFile = "": Exit Function
    Dim iFile As Integer
    iFile = FreeFile
    Dim result As String
    Dim sLine As String
    Dim isFirst As Boolean
    isFirst = True
    Open path For Input As #iFile
    Do While Not EOF(iFile)
        Line Input #iFile, sLine
        If isFirst Then
            result = sLine
            isFirst = False
        Else
            result = result & vbCrLf & sLine
        End If
    Loop
    Close #iFile
    Do While Len(result) > 0 And (Right(result, 1) = vbCr Or Right(result, 1) = vbLf)
        result = Left(result, Len(result) - 1)
    Loop
    LeggiFile = Trim(result)
End Function

Private Function ParseMeseAnno(testo As String, ByRef mese As Integer, ByRef anno As Integer) As Boolean
    ParseMeseAnno = False
    mese = 0
    anno = 0
    Dim p As Integer
    p = InStr(testo, "/")
    If p = 0 Then Exit Function
    Dim sp As Integer
    Dim i As Integer
    sp = 0
    For i = p - 1 To 1 Step -1
        If Mid(testo, i, 1) = " " Then sp = i: Exit For
    Next i
    If sp = 0 Then Exit Function
    Dim sM As String
    Dim sA As String
    sM = Trim(Mid(testo, sp + 1, p - sp - 1))
    sA = Trim(Mid(testo, p + 1))
    If Not IsNumeric(sM) Or Not IsNumeric(sA) Then Exit Function
    Dim mv As Integer
    Dim av As Integer
    mv = CInt(sM)
    av = CInt(sA)
    If mv < 1 Or mv > 12 Or av < 2000 Or av > 2100 Then Exit Function
    mese = mv
    anno = av
    ParseMeseAnno = True
End Function

Private Sub AnalizzaMesiPresenti(ws As Worksheet, anno As Integer, ByRef mp() As Boolean)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, D_PRODOTTO).End(xlUp).Row
    If lastRow < PRIMA_RIGA Then Exit Sub
    Dim r As Long
    Dim rAnno As Integer
    Dim rMese As Integer
    For r = PRIMA_RIGA To lastRow
        On Error Resume Next
        rAnno = CInt(ws.Cells(r, D_ANNO).Value)
        rMese = CInt(ws.Cells(r, D_MESE).Value)
        On Error GoTo 0
        If rAnno = anno And rMese >= 1 And rMese <= 12 Then mp(rMese) = True
    Next r
End Sub

Private Sub EliminaRigheMese(ws As Worksheet, mese As Integer, anno As Integer)
    Dim r As Long
    For r = ws.Cells(ws.Rows.Count, D_PRODOTTO).End(xlUp).Row To PRIMA_RIGA Step -1
        On Error Resume Next
        Dim rAnno As Integer
        Dim rMese As Integer
        rAnno = CInt(ws.Cells(r, D_ANNO).Value)
        rMese = CInt(ws.Cells(r, D_MESE).Value)
        On Error GoTo 0
        If rAnno = anno And rMese = mese Then ws.Rows(r).Delete
    Next r
End Sub

Private Function CaricaDatiPrecedenti(ws As Worksheet, meseMax As Integer, anno As Integer) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, D_PRODOTTO).End(xlUp).Row
    If lastRow < PRIMA_RIGA Then
        Set CaricaDatiPrecedenti = dict
        Exit Function
    End If
    Dim dstData As Variant
    dstData = ws.Range(ws.Cells(PRIMA_RIGA, 1), ws.Cells(lastRow, D_OMOG)).Value
    Dim r As Long
    Dim chiave As String
    Dim arr() As Double
    Dim idx As Integer
    Dim rAnno As Integer
    Dim rMese As Integer
    For r = 1 To UBound(dstData, 1)
        rAnno = CInt(dstData(r, D_ANNO))
        rMese = CInt(dstData(r, D_MESE))
        If rAnno <> anno Then GoTo SkipRow
        If rMese > meseMax Then GoTo SkipRow
        If Trim(CStr(dstData(r, D_PRODOTTO))) = "" Then GoTo SkipRow
        ' Chiave: Codice + Provincia (prima della migrazione a provincia era Codice + CodSede)
        chiave = NormCodice(dstData(r, D_CODICE)) & "|" & Trim(CStr(dstData(r, D_PROVINCIA)))
        If dict.Exists(chiave) Then
            arr = dict(chiave)
        Else
            ReDim arr(0 To N_DIFF + 1)
        End If
        For idx = 0 To N_DIFF - 1
            arr(idx) = arr(idx) + ToDouble(dstData(r, D_P0 + idx))
        Next idx
        arr(N_DIFF) = arr(N_DIFF) + ToDouble(dstData(r, D_OMOG))
        arr(N_DIFF + 1) = arr(N_DIFF + 1) + ToDouble(dstData(r, D_GIACINIZ))
        dict(chiave) = arr
SkipRow:
    Next r
    Set CaricaDatiPrecedenti = dict
End Function

' Scrive le formule Coefficiente/IndiceDeflusso/GiacenzaOmog/PervenutoOmog/
' Aggregati/Gestione per l'intervallo di righe [rFirst,rLast] del
' Consolidato (layout a Provincia). Usata sia da ImportaDatiMensili sia da
' MigraStoricoProvincia, cosi' le formule restano identiche ovunque.
' NOTA: a differenza della vecchia ImportaDatiMensili (dove per un bug la
' formula "Gestione" veniva scritta per errore nello stesso range di
' "Aggregati", sovrascrivendola), qui le due formule sono scritte in range
' distinti, come gia' corretto nelle macro Filtra/FiltraRegione/FiltraProvincia.
Private Sub ScriviFormuleProvincia(ws As Worksheet, rFirst As Long, rLast As Long)
    ws.Range(ws.Cells(rFirst, D_AM), ws.Cells(rLast, D_AM)).FormulaR1C1 = _
        "=IF(RC" & D_TOTDEF & " > 0,RC" & D_OMOG & "/RC" & D_TOTDEF & ",0)"
    ws.Range(ws.Cells(rFirst, D_AM + 1), ws.Cells(rLast, D_AM + 1)).FormulaR1C1 = _
        "=IF(RC" & D_TOTPERV & "=0,0,RC" & D_TOTDEF & "/RC" & D_TOTPERV & ")"
    ws.Range(ws.Cells(rFirst, D_AM + 2), ws.Cells(rLast, D_AM + 2)).FormulaR1C1 = _
        "=RC" & D_GIACFIN & "*RC" & D_AM
    ws.Range(ws.Cells(rFirst, D_AM + 3), ws.Cells(rLast, D_AM + 3)).FormulaR1C1 = _
        "=RC" & D_TOTPERV & "*RC" & D_AM
    ws.Range(ws.Cells(rFirst, D_AM + 4), ws.Cells(rLast, D_AM + 4)).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(LEFT(RC" & D_PRODOTTO & ",4),prodotti!C3:C3,prodotti!C4:C4),"""")"
    ws.Range(ws.Cells(rFirst, D_AM + 5), ws.Cells(rLast, D_AM + 5)).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(LEFT(RC" & D_PRODOTTO & ",4),prodotti!C3:C3,prodotti!C5:C5),"""")"
End Sub

' Scrive le intestazioni di colonna del Consolidato (layout a Provincia)
' sulla riga indicata (tipicamente PRIMA_RIGA - 1). Usata da
' MigraStoricoProvincia; ImportaDatiMensili non la richiama perche' le
' intestazioni, dopo la prima migrazione, restano gia' corrette.
Private Sub ScriviIntestazioniProvincia(ws As Worksheet, headerRow As Long)
    ws.Cells(headerRow, D_ANNO).Value = "Anno"
    ws.Cells(headerRow, D_MESE).Value = "Mese"
    ws.Cells(headerRow, D_AREA).Value = "Area"
    ws.Cells(headerRow, D_PRODOTTO).Value = "Prodotto Outcome"
    ws.Cells(headerRow, D_CODICE).Value = "Codice"
    ws.Cells(headerRow, D_DESCRIZ).Value = "Descrizione"
    ws.Cells(headerRow, D_REGIONE).Value = "Regione"
    ws.Cells(headerRow, D_PROVINCIA).Value = "Descrizione Provincia"
    ws.Cells(headerRow, D_GIACINIZ).Value = "Giacenza Iniziale"

    Dim labelsPD As Variant
    labelsPD = Array("P0", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "Totale Pervenuti", _
                      "D0", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "Totale Definiti")
    Dim i As Integer
    For i = 0 To N_DIFF - 1
        ws.Cells(headerRow, D_P0 + i).Value = labelsPD(i)
    Next i

    ws.Cells(headerRow, D_GIACFIN).Value = "Giacenza Finale"
    ws.Cells(headerRow, D_OMOG).Value = "Totale Omogeneizzato"
    ws.Cells(headerRow, D_GIACGG).Value = "Giacenza Gg"
    ws.Cells(headerRow, D_AM).Value = "Coefficiente"
    ws.Cells(headerRow, D_AM + 1).Value = "IndiceDeflusso"
    ws.Cells(headerRow, D_AM + 2).Value = "GiacenzaOmog"
    ws.Cells(headerRow, D_AM + 3).Value = "PervenutoOmog"
    ws.Cells(headerRow, D_AM + 4).Value = "Aggregati"
    ws.Cells(headerRow, D_AM + 5).Value = "Gestione"
End Sub

Private Function StrToDouble(v As String) As Double
    Dim s As String
    s = Trim(v)
    If s = "" Then StrToDouble = 0: Exit Function
    If InStr(s, ",") > 0 Then
        s = Replace(s, ".", "")
        s = Replace(s, ",", ".")
    End If
    If IsNumeric(s) Then StrToDouble = val(s) Else StrToDouble = 0
End Function

Private Function ToDouble(v As Variant) As Double
    If IsNumeric(v) Then ToDouble = CDbl(v) Else ToDouble = 0
End Function

' Normalizza un codice numerico (stringa o numero) in stringa intera senza leading zero
' "040184" -> "40184", 230000 -> "230000", "230000" -> "230000"
Private Function NormCodice(v As Variant) As String
    On Error Resume Next
    NormCodice = CStr(CLng(v))
    If Err.Number <> 0 Then NormCodice = Trim(CStr(v))
    On Error GoTo 0
End Function

' Normalizza un Codice Sede su 6 cifre con zeri iniziali, replicando
' TEXT(v,"000000") usata dalle formule XLOOKUP di Territorio/Provincia.
Private Function Norm6(v As Variant) As String
    On Error Resume Next
    Dim n As Long
    n = CLng(v)
    If Err.Number = 0 Then
        Norm6 = Format(n, "000000")
    Else
        Err.Clear
        Norm6 = Trim(CStr(v))
    End If
    On Error GoTo 0
End Function
