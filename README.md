# File-Maneger — file explorer in stile Windows Explorer

Un file explorer completo in **Python** basato sulla **libreria standard**
(`tkinter`). Funziona ovunque ci sia Python; l'unica dipendenza **opzionale**
è **mutagen** (pura Python, compatibile con Nuitka) per l'anteprima dei
metadati audio/video — i build .bat la installano automaticamente.

## Avvio rapido

```bash
python file_explorer.py
```

Puoi anche passare una cartella iniziale:

```bash
python file_explorer.py "C:\Users\TuoNome\Download"
```

## Creare il file .exe (Windows)

Doppio clic su **`build_exe.bat`** (PyInstaller) o **`build_nuitka.bat`**
(Nuitka: exe più veloce e protetto dalla decompilazione), oppure eseguili
dal terminale. I batch installano le dipendenze e generano l'eseguibile
standalone in:

```
dist\EsploraFile.exe
```

L'.exe non richiede Python installato per funzionare.

> **⚠ Antivirus**: alcuni antivirus (in particolare **Bitdefender**) bloccano
> la fase finale della compilazione Nuitka (`Failed to add resources`) e
> possono mettere in quarantena l'exe generato (falso positivo sui binari
> non firmati). Se la compilazione fallisce in quella fase, aggiungi la
> cartella del progetto alle **esclusioni** dell'antivirus e rilancia il
> build.

## Funzionalità

- **Schede multiple come nel browser** (barra disegnata, stile Chrome/Edge):
  Ctrl+T nuova scheda, Ctrl+W chiudi, **Ctrl+Maiusc+T riapri l'ultima scheda
  chiusa**, Ctrl+Tab / Ctrl+Maiusc+Tab per cambiare,
  clic centrale o tasto destro su una scheda per chiuderla / duplicarla,
  doppio clic sullo spazio vuoto per aprire una nuova scheda, **Ctrl+doppio
  clic su una cartella** o **clic centrale (rotella)** su una cartella (o
  "Apri in nuova scheda" dal menu) per aprirla in una scheda separata,
  frecce di scorrimento e rotella del mouse
- **Drag & drop**: trascina file e cartelle su una cartella, sulla barra
  laterale o sull'intestazione di un'altra scheda per **spostarli**;
  tieni premuto **Ctrl** mentre rilasci per **copiarli**
- **Copia/sposta in background con barra di avanzamento**: Incolla e il
  drag & drop non bloccano più l'interfaccia: l'operazione procede in un
  thread separato mostrando nella barra di stato il numero di file elaborati
  e una barra di avanzamento. Anche compressione, estrazione, ricerca
  duplicati e statistiche mostrano una **barra animata (indeterminata)**
- **Navigazione**: Indietro / Avanti (con cronologia per ogni scheda), cartella
  superiore, home, aggiorna e menu **Cartelle recenti** (le ultime cartelle
  visitate, persistite tra le sessioni)
- **Barra laterale**: Accesso rapido (Desktop, Documenti, Download, ...),
  unità disco e **Preferiti** (aggiungi/rimuovi una cartella dal menu
  contestuale; tasto destro su un Preferito per rimuoverlo, salvato tra le sessioni)
- **Filtro per estensione**: un menu a tendina nella barra degli strumenti
  mostra solo i file con l'estensione scelta (le cartelle restano visibili)
- **Barra degli indirizzi con breadcrumb cliccabile**: clicca un segmento del
  percorso per raggiungerlo; il tasto ✎ (o il percorso in modifica) permette
  di digitare un percorso manualmente (Invio per andare, Esc per annullare)
- **Ricerca RICORSIVA**: il filtro cerca anche in tutte le sottocartelle,
  in background (la UI resta fluida), con colonna "Percorso" nei risultati;
  spunta **Contenuto** per cercare il testo DENTRO i file (grep), e **Regex**
  per interpretare il filtro come espressione regolare
- **Anteprima file** (pannello laterale, 👁 o Alt+P): immagini PNG/GIF/PPM,
  JPEG/BMP/TIFF/ICO/WebP (via GDI+ di Windows), file di testo e codice
  (decine di formati: script, sottotitoli, email, Markdown, LaTeX, ...),
  testo estratto dai PDF e dai documenti **Office** (DOCX/PPTX/XLSX;
  DOC/PPT/XLS binari best-effort), contenuto degli **archivi ZIP/TAR**,
  aggiornata cliccando un file
- **Presentazione (slideshow)**: a tutto schermo delle immagini della cartella
  (Visualizza > Presentazione o menu contestuale su una cartella); frecce/click
  per avanzare, Spazio per pausa, Esc per chiudere
- **Anteprima audio/video** (metadati): selezionando un file MP3, FLAC, OGG,
  M4A, WMA, WAV, OPUS, MP4, MKV o WebM il pannello mostra durata, bitrate,
  frequenza, canali, profondità, risoluzione (video) e tag (titolo, artista,
  album, anno, genere). Richiede la libreria **mutagen**: se manca, l'app
  funziona comunque e mostra un avviso
- **Risoluzione anteprima**: dal menu Visualizza scegli come aprire le
  immagini nell'anteprima: **Bassa** (adattate al pannello), **Media**
  (100%, dimensione reale) o **Alta** (200%, dettaglio). La scelta è
  salvata tra le sessioni e limita anche lo zoom massimo per risparmiare
  memoria
- **Zoom immagini**: rotella del mouse sopra l'anteprima per ingrandire /
  rimpicciolire, doppio clic per tornare al 100%
- **Ordinamento** per Nome / Tipo / Dimensioni / Data (clic sulle intestazioni),
  con **ordinamento naturale** ('file2' prima di 'file10') e **freccia ▲/▼**
  sulla colonna ordinata per indicare la direzione
- **Doppio clic / Invio** per aprire cartelle e file con il programma predefinito
- **Menu contestuale** (tasto destro): Apri, **Apri tutti i selezionati**,
  Taglia, Copia, **Copia percorso** (tutti i selezionati, uno per riga),
  **Copia nome** (i soli nomi, uno per riga), **Seleziona per dimensione**,
  **Seleziona per data** e **Seleziona per regex** (menu Modifica),
  **Duplica** (nella stessa cartella, in background),
  **Crea collegamento** (.lnk, Windows), Rinomina,
  **Rinomina in batch** (modello con {n}/{name}/{ext}, anteprima in tempo
  reale), **Sostituisci nei nomi** (trova → sostituisci, con anteprima),
  **Elimina definitivamente** (irreversibile, con conferma),
  **Distruggi (sovrascrivi)**: sovrascrive il contenuto con byte casuali
  (3 passaggi, fsync) prima di cancellarlo, anche per intere cartelle, con
  protezioni: radici di unità e cartelle di sistema (anche le loro
  sottocartelle, es. C:\Windows\System32) bloccate, rifiuto delle selezioni
  oltre 5 GB e doppia conferma (digitare "elimina" oltre 500 MB),
  **Sposta nel cestino** (Cestino di Windows, recuperabile, senza conferma),
  Proprietà, **Apri terminale qui** (o nella cartella selezionata),
  **Apri in Esplora Risorse** (o Finder/file manager su altri OS): apre il
  file manager di sistema mostrando la cartella corrente o il file selezionato
- **Barra di stato**: conta gli elementi, la selezione, la **dimensione totale
  degli elementi selezionati** e lo **spazio libero/totale sull'unità** corrente
- **Confronta cartelle** (menu Navigazione): mostra i nomi presenti solo in
  una cartella, solo nell'altra e in comune
- **Nuova finestra** (menu File): apre una seconda istanza dell'app sulla
  cartella corrente
- **Menu Crea**: dal menu contestuale crea cartelle, file di testo (.txt),
  documenti Word (.docx), presentazioni PowerPoint (.pptx) e file Python
  (.py), chiedendo il nome; i documenti Office creati sono già validi e
  mostrano l'anteprima del testo
- **Eliminazione nel Cestino** via API Windows nativa (nessuna libreria extra),
  incluso lo **svuotamento del Cestino** (menu File, con conferma)
- **Esporta elenco (CSV)**: dal menu File salva l'elenco visualizzato (cartella
  o risultati di ricerca) in CSV con BOM, apribile direttamente in Excel
- **Comprimi in ZIP / 7z / RAR ed Estrai qui**: dal menu contestuale, in
  background (file e cartelle, anche più elementi; l'estrazione ignora i nomi
  non sicuri). ZIP/TAR via libreria standard; **7z** via 7-Zip o py7zr
  (pura Python); **RAR** via 7-Zip/WinRAR. La **creazione RAR** richiede
  WinRAR (formato proprietario)
- **Dimensione cartelle**: dal menu contestuale su una cartella, calcolata in
  background con **cache persistente**: le dimensioni vengono ricordate tra
  una sessione e l'altra, la scansione riprende se interrotta e viene
  ricalcolata automaticamente quando la cartella cambia; durante il calcolo
  la barra di stato mostra il **numero che sale in tempo reale**
  (dimensione e numero di file scanditi finora)
- **Dimensioni automatiche**: appena navighi in una cartella, le dimensioni
  delle sottocartelle visibili vengono calcolate da sole in background (una
  alla volta, senza bloccare l'interfaccia) e salvate nella cache: al ritorno
  o al riavvio il valore compare subito
- **Tema scuro**: dal menu Visualizza attiva l'interfaccia scura (albero,
  anteprima, barra laterale, tooltip); la scelta viene salvata tra le sessioni
- **Statistiche cartella**: dal menu contestuale su una cartella, mostra in
  una finestra il riepilogo (file, cartelle, byte totali, dimensione media,
  file più grande) e la tabella per tipo/estensione (conteggio e byte,
  ordinata per dimensione)
- **Apri con…** (Windows): scegli un altro programma dal menu contestuale
- **Confronta file**: seleziona due file di testo e confrontali dal menu
  contestuale (diff unificato colorato, via libreria standard `difflib`)
- **Cerca duplicati**: dal menu contestuale su una cartella, trova i file con
  lo stesso contenuto (raggruppati per dimensione, confermati con hash MD5,
  in background) e li mostra in una finestra con doppio clic per aprire,
  oltre a **eliminare i selezionati** (nel Cestino o definitivamente)
- **Tooltip con dettagli**: passa il mouse su un file o una cartella per
  vedere subito tipo, dimensione, data di modifica e percorso completo
- **File nascosti** opzionali, vista dettagli / solo nomi
- **Proprietà** con nome, posizione, dimensione, date di creazione/modifica/accesso,
  **hash MD5 e SHA-256** calcolati in background (con percentuale di avanzamento)
  e **valori selezionabili e copiabili** (Ctrl+C, tasto destro o "Copia tutto");
  con più elementi selezionati mostra file/cartelle e dimensione totale aggregata;
  **Calcola hash** calcola MD5 e SHA-256 di più file in background
- **Sessione persistente**: geometria della finestra, larghezza delle colonne,
  ordinamento, pannelli visibili (anteprima, dettagli, file nascosti) e schede
  aperte vengono ricordati e ripristinati al riavvio

## Scorciatoie da tastiera

| Tasto | Azione |
|---|---|
| Invio / doppio clic | Apri |
| Alt + ← / Alt + → | Indietro / Avanti |
| Tasti laterali del mouse (XBUTTON1/2) | Indietro / Avanti (come le frecce) |
| Alt + ↑ | Cartella superiore |
| F2 | Rinomina |
| Canc | Elimina definitivamente (irreversibile) |
| Ctrl + C / X / V | Copia / Taglia / Incolla |
| Ctrl + A | Seleziona tutto |
| Ctrl + Maiusc + A | Seleziona stesso tipo |
| Ctrl + I | Inverti selezione |
| Ctrl + Maiusc + M | Seleziona per modello (wildcard) |
| Ctrl + F | Cerca (focus sul filtro) |
| Ctrl + L | Barra indirizzi |
| F5 | Aggiorna |
| Backspace | Cartella superiore |
| Ctrl + T | Nuova scheda |
| Ctrl + W | Chiudi scheda |
| Ctrl + Tab / Ctrl + Maiusc + Tab | Scheda successiva / precedente |
| Ctrl + click | Copia tramite drag & drop |
| Alt + P | Mostra / nascondi anteprima |
| Rotella (sopra l'anteprima) | Zoom avanti / indietro sulle immagini |
| Doppio clic (sopra l'anteprima) | Ripristina zoom 100% |

## Test

Suite funzionale (schede, drag & drop, ricerca ricorsiva):

```bash
python test_file_explorer.py
```
