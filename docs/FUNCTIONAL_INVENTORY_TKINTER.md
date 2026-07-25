# Inventario funzionale della versione Tkinter

Questo inventario definisce le funzioni che la futura interfaccia PySide6/QML dovrà mantenere prima della rimozione della GUI Tkinter.

## Nuovo progetto

- Selezione di presentazioni `.pptx`.
- Lettura degli script da Excel o dalle note PowerPoint.
- Scelta di foglio e colonna Excel.
- Scelta della voce e regolazione di velocità, tonalità e volume.
- Anteprima vocale.
- Generazione di un PowerPoint narrato.
- Generazione video MP4.
- Sottotitoli e report delle durate.
- Cache degli audio e gestione della concorrenza.

## Ripara PowerPoint

- Analisi di audio e script mancanti o non validi.
- Filtri e selezione delle slide da correggere.
- Rigenerazione della narrazione.
- Salvataggio di una nuova presentazione corretta.

## Elaborazione batch

- Importazione di un manifest JSON.
- Avvio di più elaborazioni.
- Stato, progressione ed errori per ogni lavoro.
- Report finale.

## Libreria voci

- Voci Microsoft Edge TTS.
- Importazione e registrazione di campioni vocali.
- Voci locali PocketTTS, Chatterbox e XTTS quando installate.
- Anteprima, rinomina ed eliminazione.

## Cronologia

- Ricerca dei report di esecuzione.
- Apertura del report e del risultato.
- Apertura della cartella di output.

## Impostazioni

- Tema chiaro, scuro o di sistema.
- Percorsi e preferenze locali.
- Controllo di FFmpeg, LibreOffice, PowerPoint ed Edge TTS.
- Installazione guidata degli strumenti supportati su Windows.

## Installazione, avvio e manutenzione

- Installer e launcher Windows `.bat` / `.ps1`.
- Riparazione dell'ambiente virtuale.
- Disinstallazione dei componenti senza rimuovere Python.
- Build portable e installer Windows.
- Firma, checksum e scansione della release.

## Requisiti di parità per la nuova UI

Una funzione può essere considerata migrata soltanto quando:

1. produce lo stesso risultato della baseline su file di prova equivalenti;
2. mostra errori comprensibili;
3. non blocca l'interfaccia durante le operazioni lunghe;
4. dispone di test automatici;
5. dispone di screenshot nei temi chiaro e scuro;
6. è stata verificata sui sistemi operativi dichiarati compatibili.
