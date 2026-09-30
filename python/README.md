# kpimeta: metadati dei KPI nel DB Excel (versione Python)

Versione Python dell'app MATLAB `IC_KPIs_App`. Serve a creare una riga di metadati per ogni riga di KPI e a scriverle entrambe nel DB Excel letto da Power BI.

- **Stesso DB:** stesso file, stesso nome, stessi fogli, stesse tabelle (`TableMet`, `TableKPI`) e stesse colonne. Power BI non deve cambiare nulla.
- **Flusso guidato in 5 passi.** La revisione delle righe si fa dentro l'app: non serve più aprire l'Excel intermedio.
- **Uno schema unico** (`config/schema_metadata.csv`) descrive i 71 metadati. Da lì derivano form, validazione e liste.
- **Mapping LoLa in CSV** al posto del `.mat`, così si può tenere in git.
- **Scrittura sicura:** backup automatico, una sola scrittura per entrambe le tabelle, verifica del file prima di sostituire l'originale.

L'app MATLAB resta nella cartella superiore e non è stata modificata.

## Installazione (Windows)

Serve Python 3.10 o superiore (testato con 3.10, 3.11, 3.12 e 3.13). `py -0` elenca le versioni installate.

In PowerShell:

```powershell
cd python
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

- Se PowerShell blocca `Activate.ps1` ("esecuzione di script disabilitata"), esegui una volta `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.
- Nel Prompt dei comandi (cmd) l'attivazione è `.venv\Scripts\activate.bat`.
- Se è già attivo un altro ambiente (il prompt inizia con `(.venv)`), esegui prima `deactivate`.

## Primo avvio

1. **Converti il mapping LoLa** (una volta sola). Poi aggiungi `config\lola_mapping.csv` al repository.
   ```bat
   python -m kpimeta convert-mapping C:\percorso\Mapping_LoLa_DB_mapping.mat -o config\lola_mapping.csv
   ```
   Se il `.mat` è salvato in formato v7.3, il messaggio d'errore mostra il comando MATLAB per esportarlo in CSV. Si può convertire anche da `.xlsx`, con le stesse colonne usate da `load_mapping_from_excel.m`.
2. **Controlla il DB**, in sola lettura:
   ```bat
   python -m kpimeta doctor C:\percorso\KPI_DB.xlsx
   ```
3. **Avvia l'app:**
   ```bat
   streamlit run app.py
   ```
   Si apre il browser su http://localhost:8501. Il DB, la cartella LoLa e l'Excel usati l'ultima volta vengono ricordati. Il pulsante "Sfoglia…" apre la finestra di scelta file di Windows.

## Provarlo in 5 minuti senza toccare il DB vero

```bat
python -m kpimeta demo C:\temp\kpimeta_demo
streamlit run app.py
```

Nell'app usa questi percorsi:
- DB: `C:\temp\kpimeta_demo\KPI_DB_demo.xlsx`
- mapping: `C:\temp\kpimeta_demo\lola_mapping_demo.csv`
- cartella LoLa: `C:\temp\kpimeta_demo\lola`

Cosa aspettarsi:
- vengono lette 8 righe;
- c'è 1 errore, il valore `n.a.` in `KPI_C` alla riga 1, da correggere nella tabella del passo 4;
- `KPI_D` non esiste ancora nel DB e al passo 5 si sceglie di aggiungerlo;
- la scrittura crea gli ID 3–10.

I dati demo sono **sintetici**: servono a provare lo strumento, non a validarlo.

## Collaudo sul DB vero (prima di usarlo davvero)

1. Copia il DB in una cartella che Power BI **non** legge.
2. Esegui `python -m kpimeta doctor copia.xlsx` e controlla due cose:
   - che la scrittura sia possibile;
   - la tabella "Come sono salvate oggi le colonne data". L'app MATLAB ha scritto date in tre modi diversi (vedi i bug 1 e 2 più sotto). Le date anomale vengono segnalate con la data corretta probabile, ma **non sono corrette in automatico**.
3. Con l'app, collega la copia, leggi un output LoLa reale, rivedi le righe e scrivi.
4. Apri la copia in Excel e controlla righe, ID, date e nuove colonne KPI.
5. Punta una copia del report Power BI alla copia del DB e aggiorna. Verifica tipi (per esempio `Date`) e relazioni.
6. Confronta con MATLAB. Passa gli stessi file LoLa nell'app MATLAB e in `python -m kpimeta lola file1.xlsx file2.xlsx -m config\lola_mapping.csv -o ALL_DB.xlsx`. Le differenze attese sono i decimali completi (MATLAB li troncava) e le colonne con "date"/"data" nel nome non più svuotate.

## Distribuire l'app ai colleghi

Non serve compilare nulla e non serve un server. Si crea una **cartella portatile** che contiene un suo Python, le librerie e l'app.

Ogni collega:
1. la copia sul proprio PC;
2. fa doppio clic su `Avvia KPI metadata.bat`.

L'app gira sul PC del collega e il browser si collega solo a quel PC (`http://127.0.0.1:8501`). Non ci sono link da condividere né porte aperte sulla rete, e non servono installazioni né permessi di amministratore.

Per creare la cartella, sul tuo PC:

```powershell
cd python
py build_portable.py
```

- **Python da usare:** lo script copia il Python con cui lo lanci, senza i pacchetti installati. Deve essere un Python di python.org; quello del Microsoft Store o di Anaconda viene rifiutato con un messaggio. Le librerie vengono scaricate da PyPI come con `pip install`.
- **Risultato:** `dist\KPI_metadata\` e `dist\KPI_metadata.zip`, di qualche centinaio di MB (lo zip è circa un terzo).
- **Configurazione:** `config\` viene copiata così com'è, con `settings.toml`, lo schema e soprattutto `lola_mapping.csv`. Aggiornala prima del build.
- **Distribuzione:** metti lo zip in una cartella condivisa. Ogni collega lo estrae sul proprio PC e segue `LEGGIMI.txt`, che è nella cartella.
- **Nuova versione:** rifai il build e ridistribuisci lo zip. Preferenze e backup di ciascuno restano in `%LOCALAPPDATA%\kpimeta`.
- **Riga di comando:** c'è anche `kpimeta.bat`, per esempio `kpimeta.bat doctor percorso\DB.xlsx`.

## Il flusso dell'app

| Passo | Cosa fa |
|---|---|
| 1 · DB | Legge il DB senza modificarlo. Mostra righe, prossimo ID, KPI, template e controlli di coerenza (ID, duplicati, Metadata ↔ KPIs). I problemi bloccanti disabilitano la scrittura. |
| 2 · Sorgente | Quattro sorgenti possibili: **output LoLa** (cartella, scelta dei file, mapping); **Excel Metadata + KPI** (per esempio `Metadata_ALL_DB.xlsx`); **nuovo test** (una riga); **modifica di un ID esistente**. |
| 3 · Metadati | Form generato dallo schema, con le schede General e IC. Le tendine prendono i valori dai fogli di supporto del DB; il pulsante "+" aggiunge un valore al foglio. Si possono applicare template da `Metadata_Templates` e importare o esportare il form come Excel di una riga. Opzioni: sovrascrivere o solo riempire, dedurre i dati dal nome file, applicare un template per riga. |
| 4 · Revisione | Tabella modificabile, con i problemi indicati riga per riga e ricontrollati a ogni modifica. Filtri "solo righe con problemi" e "solo colonne con problemi", scelta delle colonne visibili, pulsante "Accetta come vuoti i valori non validi". I valori che mancano da un foglio di supporto (Brand, Cluster…) si aggiungono al foglio da qui, senza tornare al passo 3. C'è anche il download Excel nel formato `Metadata_ALL_DB`, da modificare in Excel e reimportare al passo 2. |
| 5 · Scrittura | Per i KPI assenti da `TableKPI` si sceglie: nuova colonna, colonna esistente oppure scarta. Poi anteprima, backup e una sola scrittura per entrambe le tabelle. |

**Precedenza dei valori:** colonne protette della sorgente (`Repetition_ID`, `FileName`, `StartSpeed` se valido) > nome file > template > form. È la regola di `writeMetadataToNewALLDB.m`. I campi vuoti non sovrascrivono nulla.

## Riga di comando

Tutte le operazioni dell'app sono disponibili anche da riga di comando: `python -m kpimeta <comando> -h`.

```bat
python -m kpimeta doctor DB.xlsx
python -m kpimeta lola C:\LoLa\*.xlsx -o C:\LoLa\ALL_DB.xlsx
python -m kpimeta prepare --lola C:\LoLa\*.xlsx --db DB.xlsx --template PRJ1_VEH1 --set IDName="Prova" --infer-filename -o da_rivedere.xlsx
python -m kpimeta check da_rivedere.xlsx --db DB.xlsx --new-kpi add
python -m kpimeta write da_rivedere.xlsx --db DB.xlsx --new-kpi add --dry-run
python -m kpimeta write da_rivedere.xlsx --db DB.xlsx --new-kpi add
```

- `lola` è l'equivalente di `process_excel`. I caratteri jolly li espande il programma; i file intermedi (`ALL_DB.xlsx`, `Metadata_ALL_DB.xlsx`, `*_DB.xlsx`) vengono saltati.
- `prepare` compone le righe in un Excel da rivedere; `check` lo controlla contro il DB; `write` lo scrive.
- Codici di uscita: `0` tutto ok; `1` errori nei dati, niente scritto; `2` errore di file, DB o configurazione.

## Compatibilità con il flusso attuale e con Power BI

- **Il DB è lo stesso file.** Viene sostituito sul posto, con lo stesso nome. Fogli, tabelle e colonne restano gli stessi e nello stesso ordine. Le righe vengono aggiunte in fondo alle tabelle.
- **Nuove colonne KPI** solo se lo scegli al passo 5. Vengono aggiunte in fondo a `TableKPI`.
- **ID:** massimo + 1. Lo stesso ID e lo stesso ordine in `TableMet` e `TableKPI`, come richiede il controllo di coerenza dell'app MATLAB.
- **Date:** per default sono testo `gg/mm/aaaa`, il formato che il flusso di import MATLAB voleva scrivere. Con `date_storage = "excel"` in `config/settings.toml` diventano date Excel vere.
- **`ALL_DB.xlsx` e `Metadata_ALL_DB.xlsx`** non servono più al flusso. Si possono ancora creare (casella al passo 2, comando `lola`, download al passo 4) e reimportare.
- **Nomi** di fogli, tabelle, colonna ID, template, regola del nome file e mapping sono in `config/settings.toml`.

## Sicurezza della scrittura

- Rifiuta di scrivere se il DB è aperto in Excel (file `~$…` o file bloccato).
- Prima di ogni scrittura crea un backup con data e ora in `%LOCALAPPDATA%\kpimeta\backups\<nome DB>\`. La cartella è configurabile con `backup_dir` e non va messa dentro la cartella letta da Power BI. I backup non vengono cancellati in automatico.
- Scrive su un file temporaneo nella stessa cartella e lo verifica: tabelle, intestazioni, valori scritti, nessuna parte del file persa. Solo dopo sostituisce l'originale, in un'unica operazione. Se qualcosa va storto l'originale resta intatto.
- Al passo 5 rilegge il DB se è stato modificato dopo la lettura. Se il file cambia mentre la scrittura è in corso, la scrittura viene annullata.
- Prima di scrivere controlla tutto: tipi, limiti, liste, duplicati (nel batch e contro il DB) e coerenza Metadata ↔ KPIs.

## Limiti noti (dovuti a openpyxl)

- **Formule dentro `TableMet` o `TableKPI`:** la scrittura è bloccata.
- **Formule in altri fogli:** restano, ma senza il valore calcolato finché il file non viene aperto e salvato in Excel. Se Power BI legge quei fogli vedrebbe celle vuote. L'app lo segnala al passo 1.
- **Modello dati Power Pivot, Power Query, connessioni dati, slicer:** la scrittura è disabilitata, perché openpyxl non li conserva.
- **Grafici e immagini** vengono conservati ma riscritti: controllane l'aspetto dopo la prima scrittura.
- **Convalide che puntano ad altri fogli, sparkline e formattazioni condizionali avanzate** potrebbero andare perse: l'app lo segnala.
- **Riga dei totali o colonne calcolate** nelle tabelle: la scrittura è bloccata.

**Tempi indicativi**, misurati su Linux:

| Dimensione del DB | Lettura | Scrittura |
|---|---|---|
| 1.000 righe × 131 colonne | circa 1 s | circa 3 s |
| 5.000 righe × 221 colonne | circa 10 s | circa 31 s |

## Schema verità: `config/schema_metadata.csv`

Una riga per ogni metadato, nello stesso ordine di `TableMet`. È ricavato dalle mappe `GeneralMetadataMap` e `ICMetadataMap` dell'app MATLAB. Il separatore è `;`.

| Colonna | Significato |
|---|---|
| `column` | Nome della colonna nel DB. |
| `label`, `tab`, `section` | Etichetta, scheda e sezione nel form. |
| `type` | `id`, `text`, `int`, `number`, `date`, `choice` (una voce da lista), `multichoice` (più voci separate da `;`). |
| `choices_sheet`, `choices_column` | Foglio e colonna del DB che forniscono la lista (i vecchi pulsanti "DB"). |
| `choices` | Lista fissa, per esempio `ON\|OFF`. |
| `min`, `max` | Limiti per i numeri. |
| `default` | Valore iniziale nel form. |
| `required` | Se il campo è obbligatorio. Oggi nessun campo lo è, come nell'app MATLAB. |
| `lola_policy` | `keep`: il valore della sorgente non viene mai sovrascritto. `keep_if_valid`: viene sovrascritto solo se non è valido. |
| `in_kpi_table` | Colonne copiate anche in `TableKPI` (`ID`, `IDName`, `Date`). |
| `notes` | Aiuto mostrato nel form. |

**Per aggiungere un metadato:** aggiungi la colonna a `TableMet` in Excel, poi aggiungi una riga allo schema. Nell'app premi "Ricarica configurazione".

**Mapping LoLa (`config/lola_mapping.csv`):** colonne `KPI_LoLa_Name;KPI_DB_Name;Type`.
- `Type` è facoltativa e vale `auto`, `text`, `number` o `date`.
- Solo le colonne con `Type` = `date` vengono lette come date.
- Separatore `;` oppure `,`; codifica UTF-8 oppure Windows-1252.

## Bug dell'app MATLAB e come sono gestiti qui

| # | Bug nell'app MATLAB | Qui |
|---|---|---|
| 1 | Date scritte con `datenum()`, spostate di 693.960 giorni; una data vuota diventa 1800-01-01. | Date scritte come testo `gg/mm/aaaa` o come date Excel; una data vuota resta vuota. Le date già sbagliate nel DB vengono segnalate con la data corretta. |
| 2 | Testo `gg/mm` interpretato come `mm/gg` quando passa per COM. | Niente COM. Le date si leggono prima come `gg/mm`; `mm/gg` solo in ultima istanza e con avviso. |
| 3 | `num2str` tronca i KPI a circa 5 cifre significative e li salva come testo. | I numeri restano numeri, a precisione piena. |
| 4 | Ogni colonna con "date" o "data" nel nome viene trattata come data e svuotata. | Sono date solo le colonne con `Type` = `date` nel mapping. |
| 5 | File intermedi sovrascritti senza chiedere. | Non servono più; vengono creati solo su richiesta. |
| 6 | "Close connected DBs" chiude anche gli Excel dell'utente. | Rimosso: niente COM. |
| 7 | Se un import fallisce, la scrittura parte con il batch precedente. | Una lettura fallita azzera le righe. Si scrive solo dopo la validazione. |
| 8 | Metadati scritti senza la riga KPI. | Le due tabelle sono scritte nella stessa operazione; la riga KPI c'è sempre. |
| 9 | Excel aperto e salvato due volte per riga; scritture parziali. | Una sola scrittura atomica, con verifica e backup. |
| 10 | Overwrite non funzionante. | Sorgente "Modifica un ID esistente": mostra le differenze, sovrascrive e allinea `IDName` e `Date` della riga KPI. |
| 11–12 | Reset KPIs table e Cancel dei dialog "DB" vanno in errore. | Sostituiti dal flusso a passi, dalle tendine e dal pulsante "+". |
| 13 | Mapping `.mat` con percorso relativo e fuori dal repository. | CSV in `config/`, percorso in `settings.toml`, errori mostrati nell'app. |
| 14 | Le mappe dei campi esistono solo dopo il Connect. | Lo schema si carica all'avvio e i passi guidano l'ordine. |
| – | `str2double("1,5")` restituisce 15. | Errore con il suggerimento di usare il punto. |
| – | `sanitizeKpiValue` trasforma in NaN, senza dirlo, i valori non validi. | I valori non validi vengono segnalati. Si correggono oppure si accettano come vuoti in modo esplicito. |
| – | CoOpType non viene caricato da template e import. | Gestito come `multichoice`. |

## Struttura del codice

| File | Contenuto | Equivalente MATLAB |
|---|---|---|
| `app.py` | Interfaccia Streamlit | `IC_KPIs_App.mlapp` (GUI) |
| `kpimeta/lola.py` | Lettura degli output LoLa | `process_excel.m` |
| `kpimeta/mapping.py` | Mapping LoLa → DB (CSV, xlsx, mat) | `load_mapping_from_excel.m` |
| `kpimeta/batch.py`, `infer.py` | Composizione delle righe, deduzione dal nome file | `writeMetadataToNewALLDB.m` e la logica di import dell'app |
| `kpimeta/validate.py` | Validazione, duplicati, KPI nuovi | `validateImportedMetadata` e i controlli dell'app |
| `kpimeta/excel_db.py` | Lettura e scrittura del DB | Le funzioni COM dell'app |
| `kpimeta/values.py` | Conversioni di tipo (unica regola per numeri, date e valori mancanti) | `sanitizeKpiValue.m` e le conversioni sparse |
| `kpimeta/schema.py`, `config.py` | Schema e impostazioni | Mappe dei campi nell'app |
| `kpimeta/cli.py`, `demo.py` | Riga di comando e dati demo | – |
| `build_portable.py` | Cartella portatile da distribuire ai colleghi | `App_Compiler.m` |

## Test

```bat
pip install -r requirements-dev.txt
pytest
```

Sono 82 test, circa 15 secondi, più il test facoltativo nel browser descritto sotto. Coprono conversioni, lettura LoLa, composizione, validazione, scrittura del DB (incluse interruzioni, file aperto, DB cambiato nel frattempo e parti del file da conservare), CLI e flusso dell'app.

La tabella di revisione va provata anche in un browser vero, perché le modifiche fatte nella griglia sono gestite dal frontend. Il test richiede Playwright:

```bat
pip install playwright
playwright install chromium
set KPIMETA_BROWSER_TEST=1
pytest tests\test_browser.py
```
