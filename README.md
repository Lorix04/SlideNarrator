# PPTX TTS — Audio e video per presentazioni PowerPoint

Applicazione desktop in **Python** per trasformare una presentazione PowerPoint in:

- un nuovo file **PPTX con narrazione audio incorporata**;
- un **video MP4** composto dalle slide e dalla relativa narrazione;
- sottotitoli sincronizzati e report delle durate.

Il testo da leggere viene fornito tramite un file Excel: ogni riga della colonna **A** corrisponde a una slide della presentazione.

> Il progetto è pensato principalmente per Windows e può utilizzare sia voci Microsoft online sia modelli locali di sintesi e clonazione vocale.

## Funzionalità principali

- Importazione di presentazioni `.pptx` e script `.xlsx`.
- Associazione automatica tra righe Excel e slide PowerPoint.
- Sintesi vocale con voci neurali Microsoft tramite Edge TTS.
- Clonazione vocale locale con PocketTTS, Chatterbox TTS e XTTS v2.
- Registrazione e importazione di campioni vocali.
- Libreria locale per aggiungere, rinominare ed eliminare voci personalizzate.
- Anteprima della voce prima della generazione completa.
- Regolazione della velocità di lettura.
- Inserimento automatico dell'audio nelle slide.
- Riproduzione automatica della narrazione all'apertura della slide.
- Avanzamento opzionale alla slide successiva al termine dell'audio.
- Esportazione video in 720p o 1080p.
- Dissolvenza incrociata opzionale tra le slide.
- Creazione di sottotitoli `.srt`, con possibilità di incorporarli nel video.
- Generazione di un report con la durata di ogni slide.
- Esportazione di un file `_captions.json` con i timing delle frasi, utilizzabile da applicazioni esterne come uno SCORM Builder.
- Cache degli audio già generati.
- Elaborazione parallela delle slide per ridurre i tempi di sintesi.
- Modalità sperimentale INT8 per PocketTTS su CPU.

## Flusso di elaborazione

### PowerPoint con audio

```text
Presentazione PPTX + script XLSX
              ↓
        Sintesi vocale
              ↓
 Inserimento audio e configurazione autoplay
              ↓
 PPTX narrato + report durate + captions JSON
```

### Video MP4

```text
Presentazione PPTX + script XLSX
              ↓
        Sintesi vocale
              ↓
LibreOffice: PPTX → PDF
              ↓
PyMuPDF: PDF → immagini PNG
              ↓
FFmpeg: immagini + audio + sottotitoli → MP4
```

Il video utilizza immagini statiche delle slide e non riproduce le animazioni native di PowerPoint.

## Requisiti

### Requisiti di base

- Windows 10 o Windows 11.
- Python dalla versione **3.10 alla 3.14**.
- Connessione Internet per le voci Microsoft Edge TTS.

Durante l'installazione di Python, attivare l'opzione **Add Python to PATH**.

### Requisiti aggiuntivi

| Funzione | Requisito |
|---|---|
| PowerPoint con voci Microsoft | Python e dipendenze base |
| Compatibilità audio avanzata | FFmpeg |
| Voci clonate locali | FFmpeg, PocketTTS o Chatterbox TTS |
| Registrazione dal microfono | `sounddevice` |
| Esportazione video | FFmpeg, LibreOffice e PyMuPDF |
| PocketTTS | Accesso al modello `kyutai/pocket-tts` su Hugging Face |
| XTTS v2 | `coqui-tts` e accettazione della relativa licenza |

Le voci clonate possono richiedere diversi gigabyte di spazio e una quantità significativa di RAM. In assenza di una GPU compatibile, i modelli vengono eseguiti sulla CPU.

## Installazione su Windows

1. Scaricare o clonare il repository:

   ```bash
   git clone https://github.com/Lorix04/TTS_PPTX-MP4.git
   cd TTS_PPTX-MP4
   ```

2. Eseguire con un doppio clic:

   ```text
   Installa_PPTX_TTS.bat
   ```

3. Lo script di installazione:

   - crea l'ambiente virtuale `.venv`;
   - installa le dipendenze principali;
   - propone l'installazione dei motori per le voci clonate;
   - propone l'installazione opzionale di XTTS v2;
   - verifica la presenza di FFmpeg e LibreOffice.

4. Al termine, avviare l'applicazione con:

   ```text
   Avvia_PPTX_TTS.bat
   ```

## Installazione manuale

Creare e attivare un ambiente virtuale:

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
```

Installare le dipendenze di base:

```bash
pip install edge-tts python-pptx openpyxl lxml mutagen pygame-ce pymupdf
```

Installare facoltativamente i motori di clonazione vocale:

```bash
pip install pocket-tts sounddevice
pip install chatterbox-tts
```

Installare XTTS v2 solo quando necessario:

```bash
pip install coqui-tts
```

FFmpeg e LibreOffice sono programmi esterni e devono essere installati separatamente o tramite le opzioni offerte dallo script `Installa_PPTX_TTS.bat`.

## Preparazione dei file di input

### Presentazione PowerPoint

Utilizzare una presentazione originale in formato `.pptx`.

Non usare come input un file già generato da PPTX TTS. Per cambiare voce o script, ripartire sempre dal PowerPoint originale senza audio incorporato.

### File Excel

Il file `.xlsx` deve contenere un testo per ogni slide nella colonna **A**:

| Cella | Contenuto |
|---|---|
| A1 | Testo da leggere nella slide 1 |
| A2 | Testo da leggere nella slide 2 |
| A3 | Testo da leggere nella slide 3 |

Non inserire un'intestazione, a meno che non debba essere letta come testo della prima slide.

- Le celle vuote producono slide senza narrazione.
- Se gli script sono meno delle slide, le slide finali rimangono mute.
- Se gli script sono più delle slide, quelli in eccesso vengono ignorati.

## Utilizzo tramite interfaccia grafica

1. Scegliere il tipo di output:
   - **PowerPoint con audio**;
   - **Video MP4**.
2. Selezionare il file PowerPoint originale.
3. Selezionare il file Excel contenente gli script.
4. Scegliere il percorso del file di output.
5. Selezionare la voce e la velocità di lettura.
6. Ascoltare eventualmente un'anteprima.
7. Configurare le opzioni specifiche dell'output.
8. Premere il pulsante di generazione e attendere il completamento.

### Opzioni PowerPoint

- Riproduzione automatica dell'audio.
- Avanzamento automatico al termine della narrazione.
- Riconversione degli MP3 a 44.100 Hz mono per migliorare la compatibilità con alcune versioni di PowerPoint.

### Opzioni video

- Risoluzione 720p o 1080p.
- Sottotitoli incorporati nel video.
- Dissolvenza incrociata tra le slide.

Un file `.srt` viene salvato accanto al video anche quando i sottotitoli non vengono incorporati.

## Voci disponibili

### Voci Microsoft

Le voci Microsoft vengono generate online tramite Edge TTS:

- Isabella — `it-IT-IsabellaNeural`;
- Elsa — `it-IT-ElsaNeural`;
- Diego — `it-IT-DiegoNeural`;
- Giuseppe — `it-IT-GiuseppeNeural`.

### Voci clonate locali

Sono supportati i seguenti motori:

- **PocketTTS**: soluzione consigliata per velocità e utilizzo su CPU;
- **Chatterbox TTS**: maggiore espressività, ma tempi più lunghi su CPU;
- **XTTS v2**: motore opzionale con italiano nativo, soggetto a specifiche condizioni di licenza.

Dal gestore delle voci è possibile:

- importare un campione audio;
- registrare un campione dal microfono;
- scegliere il motore TTS;
- salvare la voce nella libreria locale;
- rinominare o eliminare le voci registrate.

Per ottenere risultati migliori, utilizzare una registrazione pulita, in un ambiente silenzioso e con ritmo naturale. Il progetto raccomanda un campione di circa 30 secondi.

## File generati

Con un output chiamato `corso_audio.pptx`, l'applicazione può produrre:

```text
corso_audio.pptx
corso_audio_durate.txt
corso_audio_captions.json
```

Con un output chiamato `corso_video.mp4`, può produrre:

```text
corso_video.mp4
corso_video.srt
corso_video_durate.txt
```

Il file `_captions.json` contiene, per ogni slide, il testo e i timing iniziali e finali delle frasi.

## Utilizzo da riga di comando

### Generazione di un PowerPoint narrato

```bash
python pptx_tts.py input.pptx scripts.xlsx output_audio.pptx
```

Esempio con voce, velocità e avanzamento automatico:

```bash
python pptx_tts.py input.pptx scripts.xlsx output_audio.pptx \
  --voice it-IT-IsabellaNeural \
  --rate +5% \
  --auto-advance
```

In PowerShell o nel Prompt dei comandi di Windows, il comando può essere scritto su un'unica riga.

### Generazione di un video

```bash
python pptx_tts.py input.pptx scripts.xlsx output_video.mp4 \
  --video \
  --resolution 1080p \
  --subtitles \
  --transition
```

### Opzioni principali

| Opzione | Descrizione |
|---|---|
| `--voice` | Voce Microsoft o identificativo di una voce clonata |
| `--rate` | Velocità della voce, ad esempio `-10%`, `+0%` o `+15%` |
| `--no-autoplay` | Disattiva la riproduzione automatica |
| `--auto-advance` | Avanza alla slide successiva al termine dell'audio |
| `--concurrency` | Numero di sintesi Edge TTS eseguite contemporaneamente |
| `--transcode-audio` | Converte gli MP3 tramite FFmpeg per PowerPoint |
| `--video` | Genera un MP4 invece di un PPTX |
| `--resolution` | Imposta `720p` o `1080p` |
| `--subtitles` | Incorpora i sottotitoli nel video |
| `--transition` | Applica una dissolvenza tra le slide |
| `--pocket-variant` | Seleziona PocketTTS veloce o ad alta qualità |
| `--clone-workers` | Numero di slide elaborate in parallelo con voci clonate |
| `--no-cache` | Disattiva il riutilizzo degli audio già generati |
| `--pocket-quantize` | Attiva la modalità INT8 sperimentale di PocketTTS |

Per visualizzare tutte le opzioni:

```bash
python pptx_tts.py --help
```

## Cache

Per le voci clonate, la cache è attiva per impostazione predefinita. Gli audio vengono riutilizzati quando non cambiano:

- testo;
- voce;
- velocità;
- motore e variante del modello.

La directory predefinita è:

```text
%USERPROFILE%\.pptx_tts_cache
```

È possibile cambiarla tramite la variabile d'ambiente `PPTXTTS_CACHE_DIR` oppure svuotarla dall'interfaccia grafica.

## Struttura del progetto

```text
TTS_PPTX-MP4/
├── pptx_tts.py          # Pipeline principale e interfaccia CLI
├── pptx_tts_gui.py      # Interfaccia desktop Tkinter
├── video_export.py      # Rendering delle slide ed esportazione MP4/SRT
├── voice_clone.py       # Motori di sintesi e clonazione vocale
├── voice_library.py     # Archivio locale delle voci personalizzate
├── voice_manager.py     # Interfaccia di gestione e registrazione delle voci
├── test_limite_token.py # Test sperimentale per PocketTTS
├── Installa_PPTX_TTS.bat
├── Avvia_PPTX_TTS.bat
└── LEGGIMI_PRIMA.txt
```

La cartella `voices/` contiene i campioni vocali locali e può essere esclusa dal repository per ragioni di privacy.

## Tecnologie utilizzate

| Tecnologia | Utilizzo nel progetto |
|---|---|
| Python | Linguaggio principale e orchestrazione della pipeline |
| Tkinter / ttk | Interfaccia grafica desktop |
| Edge TTS | Sintesi vocale tramite voci Microsoft |
| PocketTTS | Clonazione vocale locale orientata alla velocità |
| Chatterbox TTS | Clonazione vocale locale più espressiva |
| Coqui XTTS v2 | Motore TTS opzionale multilingua |
| PyTorch / Torchaudio | Esecuzione dei modelli AI e trattamento dell'audio |
| Transformers | Componenti utilizzati dai modelli di sintesi |
| python-pptx | Lettura e modifica delle presentazioni PowerPoint |
| openpyxl | Lettura degli script dal file Excel |
| lxml / OOXML | Modifica avanzata dell'XML interno dei file PPTX |
| Mutagen | Lettura della durata degli MP3 |
| SoundFile / sounddevice | Gestione e registrazione dei campioni audio |
| pygame-ce | Riproduzione delle anteprime vocali |
| FFmpeg | Conversione audio, sottotitoli e montaggio video |
| LibreOffice Headless | Rendering delle presentazioni in PDF |
| PyMuPDF | Conversione delle pagine PDF in immagini PNG |
| Asyncio / multiprocessing | Sintesi concorrente e parallelizzazione |

## Limitazioni note

- Il progetto è orientato principalmente a Windows.
- Edge TTS richiede una connessione Internet.
- L'esportazione video non conserva animazioni, transizioni native o interattività del PowerPoint.
- L'attivazione dell'autoplay può sostituire i timing e le animazioni già presenti nelle slide. Conservare sempre una copia del file originale.
- Un PowerPoint già elaborato dall'applicazione non deve essere riutilizzato come input.
- Un numero elevato di processi paralleli per le voci clonate può consumare molta RAM.
- La resa delle slide esportate da LibreOffice può differire leggermente da quella di Microsoft PowerPoint.
- La modalità INT8 di PocketTTS è sperimentale e può ridurre la qualità della voce.

## Privacy e uso responsabile delle voci

Utilizzare la clonazione vocale esclusivamente con il consenso della persona interessata e nel rispetto delle normative applicabili.

Non pubblicare nel repository:

- registrazioni vocali personali;
- token di Hugging Face;
- file `.env`;
- credenziali o altri dati riservati.

È consigliabile aggiungere `voices/`, `.venv/` e la cache al file `.gitignore`.

## Licenze dei modelli

Le librerie e i modelli integrati possono avere licenze diverse da quella del codice del progetto.

In particolare, **XTTS v2** è indicato nel progetto per uso non commerciale. Verificare sempre le condizioni aggiornate delle singole dipendenze e dei modelli prima di distribuire o vendere contenuti generati.

Il repository non include attualmente un file `LICENSE` generale: aggiungerne uno prima di consentire il riutilizzo o la distribuzione del codice.

## Contributi

Segnalazioni, correzioni e proposte di miglioramento possono essere inviate tramite le issue o le pull request del repository.
