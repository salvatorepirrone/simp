# SIMP Pensioni - Macro ImportaMensile

## Cosa e' cambiato (v13)

1. **Aggregazione a livello di Provincia.** `ImportaDatiMensili` non scrive
   piu' una riga per ogni singola Sede: aggrega i dati a livello di
   Descrizione Provincia (la stessa aggregazione gia' usata dalla macro
   `FiltraProvincia`). Questo riduce drasticamente il numero di righe
   scritte ogni mese nel foglio `Consolidato`, che era la causa principale
   del peso eccessivo del file.
2. **Percorsi SharePoint autorilevati.** `SCRIPT_DIR` e `BASE_DIR` non sono
   piu' costanti fisse con un percorso locale (`C:\INPS\...`): vengono
   risolte a runtime individuando la cartella locale di sincronizzazione
   OneDrive/SharePoint, cosi' sia gli script Python sia i file prodotti
   possono risiedere su una libreria SharePoint condivisa dal team.
3. **Nessun percorso con lo username in chiaro.** Anche `PYTHON_EXE` e la
   cartella temporanea non sono piu' percorsi fissi con lo username di chi
   ha scritto la macro (`C:\Users\spirrone01\...`): vengono risolti a
   runtime per l'utente che sta effettivamente eseguendo Excel, cosi' la
   macro funziona invariata sul PC di qualunque collega.

## File

- `vba/ImportaMensile.bas` - il modulo VBA da importare nel workbook Excel
  (VBA Editor -> click destro sul progetto -> Importa file).
- `python/filtra_consolidato_provincia.py` - **nuovo**, usato da
  `ImportaDatiMensili`: filtra + aggrega per Provincia.
- `python/filtra_consolidato.py` - il vecchio script (aggregazione per
  sede), non piu' richiamato da nessuna macro dopo questa modifica.
  Tenuto solo come riferimento/rollback.
- `python/filtra_dcpensioni.py` - usato da `Filtra` (invariato).

Nota: `filtra_dcpensioni_regione.py` e `filtra_dcpensioni_provincia.py`
(usati rispettivamente da `FiltraRegione` e `FiltraProvincia`) non sono
stati toccati e restano quelli gia' presenti nella cartella script del
team: non li ho ricreati perche' non facevano parte di questa richiesta e
non erano stati forniti.

## Setup richiesto prima del primo utilizzo

### 1. Configura l'aggancio alla libreria SharePoint

In cima a `ImportaMensile.bas` valorizza:

```vb
Private Const SHAREPOINT_SITE_MATCH As String = "sites/NOME-SITO"
Private Const SHAREPOINT_LIB_MATCH  As String = "Documenti condivisi"
```

con una porzione (case-insensitive, basta un frammento) dell'indirizzo
del sito e del nome della libreria SharePoint che ospita `Scripts\` e
`SIMP_Pensioni\`. La macro individua da sola il percorso locale
sincronizzato interrogando il registro di Windows dove OneDrive registra
ogni libreria SharePoint sincronizzata (indirizzo del sito + percorso
locale corrispondente): funziona su qualsiasi PC/utente senza modificare
il codice, a patto che OneDrive abbia quella libreria effettivamente
**sincronizzata** (non solo "disponibile online").

Se sul tuo PC l'autorilevamento non funzionasse (politiche aziendali
particolari, versione OneDrive diversa, ecc.), puoi sempre forzare il
percorso a mano:

```vb
Private Const SHAREPOINT_ROOT_OVERRIDE As String = "C:\Users\<utente>\<Azienda>\<Sito> - <Libreria>\"
```

Con `SHAREPOINT_ROOT_OVERRIDE` valorizzato, l'autorilevamento viene
saltato del tutto.

### 2. Copia gli script Python

Nella cartella `<radice SharePoint>\Scripts\` copia:
`filtra_consolidato_provincia.py`, `filtra_dcpensioni.py`,
`filtra_dcpensioni_regione.py`, `filtra_dcpensioni_provincia.py`.

### 3. Python

Non serve piu' configurare nulla in genere: la macro cerca da sola `py`
(il launcher ufficiale) o `python` nel PATH di sistema, per l'utente che
sta usando Excel in quel momento (funziona quindi automaticamente anche
per i colleghi, ognuno con il proprio Python installato). Se
sul tuo PC l'autorilevamento non trovasse nulla (Python non installato, o
installato senza l'opzione "Aggiungi al PATH"), valorizza
`PYTHON_EXE_OVERRIDE` in cima al modulo con il percorso completo del tuo
`python.exe`.

### 4. Migra lo storico (una tantum)

I mesi gia' presenti nel `Consolidato` sono a livello di Sede (vecchio
layout). Prima di importare il primo mese nuovo con il layout a Provincia,
**fai un backup del file**, poi lancia la macro `MigraStoricoProvincia`
(una sola volta): riaggrega tutte le righe storiche a livello di
Provincia, cosi' l'intero foglio torna ad avere granularita' uniforme.

Dopo la migrazione, verifica manualmente i campi delle eventuali tabelle
pivot: i riferimenti a "Codice Sede"/"Sede" non sono piu' validi e vanno
sostituiti con "Descrizione Provincia".

### 5. Uso ordinario

Da qui in poi, `ImportaDatiMensili` funziona come prima (stesso flusso:
selezione file, calcolo differenziale rispetto al mese precedente,
scrittura nel Consolidato), solo con una riga per Provincia invece che per
Sede, e con i percorsi risolti automaticamente sulla libreria SharePoint
configurata.
