"""
pptx_tts_gui.py — Interfaccia grafica per pptx_tts.py

Permette di:
- selezionare il file PowerPoint e il file Excel tramite finestre di dialogo
- scegliere la voce italiana
- ascoltare un'anteprima della voce prima di generare
- regolare la velocità di lettura
- vedere il progresso della generazione in tempo reale

Requisiti:
    pip install edge-tts python-pptx openpyxl lxml mutagen pygame-ce

Opzionali (per le voci clonate locali):
    pip install pocket-tts        # motore PocketTTS (CPU, veloce)
    pip install chatterbox-tts    # motore Chatterbox (qualità più alta)
    pip install sounddevice       # solo per registrare la voce dall'app
I moduli voice_library.py, voice_clone.py e voice_manager.py devono trovarsi
nella stessa cartella. Se mancano, la GUI funziona con le sole voci Microsoft.

Uso:
    python pptx_tts_gui.py

Il file pptx_tts.py deve trovarsi nella stessa cartella.
"""

import os
import platform
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pptx_tts  # engine

# Moduli per le voci clonate. Import opzionale: se mancano, la GUI funziona
# con le sole voci Microsoft e il pannello di gestione voci viene disabilitato.
try:
    import voice_library
    import voice_clone
    from voice_manager import VoiceManagerDialog
    _CLONE_AVAILABLE = True
except Exception as _e:
    _CLONE_AVAILABLE = False
    _CLONE_IMPORT_ERROR = str(_e)

# Esportazione video: la GUI mostra l'opzione "Video" solo se il motore ha il
# modulo video_export (e quindi le sue dipendenze) disponibile.
_VIDEO_AVAILABLE = getattr(pptx_tts, "_VIDEO_AVAILABLE", False)

# Audio playback (optional — only used for the preview button).
# L'import del modulo avviene al boot, ma pygame.mixer.init() è LAZY:
# viene chiamato solo al primo utilizzo dell'anteprima vocale, in modo da
# non occupare risorse audio (e da non fallire) per chi non usa la preview.
try:
    import pygame
    PYGAME_INSTALLED = True
    PYGAME_IMPORT_ERROR = ""
except Exception as e:
    pygame = None  # type: ignore[assignment]
    PYGAME_INSTALLED = False
    PYGAME_IMPORT_ERROR = str(e)

# Stato del mixer: None = non ancora tentato, True = ok, False = fallito
_mixer_ready: bool | None = None
_mixer_error: str = ""


def _ensure_mixer_initialized() -> bool:
    """Inizializza il mixer pygame al primo utilizzo. Ritorna True se ok."""
    global _mixer_ready, _mixer_error
    if _mixer_ready is not None:
        return _mixer_ready
    if not PYGAME_INSTALLED:
        _mixer_ready = False
        _mixer_error = PYGAME_IMPORT_ERROR
        return False
    try:
        pygame.mixer.init()
        _mixer_ready = True
        return True
    except Exception as e:
        _mixer_ready = False
        _mixer_error = str(e)
        return False


# (voice_id, display_name, description)
VOICE_OPTIONS = [
    ("it-IT-IsabellaNeural", "Isabella", "femminile, calda"),
    ("it-IT-ElsaNeural",     "Elsa",     "femminile, chiara"),
    ("it-IT-DiegoNeural",    "Diego",    "maschile, professionale"),
    ("it-IT-GiuseppeNeural", "Giuseppe", "maschile, maturo"),
]

DEFAULT_PREVIEW_TEXT = (
    "Ciao, questa è una prova della voce italiana. "
    "Se ti piace, possiamo procedere con la generazione."
)


def get_voice_entries(library=None):
    """
    Elenco unificato delle voci selezionabili: prima le voci Microsoft,
    poi le voci clonate dell'utente. Ogni voce è un dict:
        {"voice_id", "name", "engine", "is_clone"}
    Funzione pura (niente Tk) per poterla testare facilmente.
    """
    entries = [
        {"voice_id": vid, "name": name, "engine": "Microsoft", "is_clone": False}
        for (vid, name, _desc) in VOICE_OPTIONS
    ]
    if _CLONE_AVAILABLE:
        lib = library if library is not None else voice_library.VoiceLibrary()
        for v in lib.list_voices():
            entries.append({
                "voice_id": v.voice_id,
                "name": v.name,
                "engine": v.engine_label,
                "is_clone": True,
            })
    return entries


class StdoutRedirector:
    """
    Pipe lo stdout del motore (chiamate a print) verso il widget Text della
    GUI. Tk non è thread-safe: ogni write avviene da un thread di lavoro,
    quindi l'aggiornamento del widget è marshalled sul thread principale di
    Tk con .after(0, ...).

    Usato sostituendo direttamente sys.stdout (vedi work() sotto): non con
    contextlib.redirect_stdout, perché quello non è thread-safe e su Windows
    salta gli output di asyncio/subprocess (ffmpeg).
    """
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        if not message:
            return len(message) if message else 0
        # Marshalling sul thread principale di Tk
        try:
            self.text_widget.after(0, self._append, message)
        except Exception:
            # Se la finestra è già stata distrutta, ignoriamo
            pass
        # Convenzione Python: write() restituisce il numero di caratteri scritti
        return len(message)

    def _append(self, message):
        try:
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", message)
            self.text_widget.see("end")
            self.text_widget.configure(state="disabled")
        except Exception:
            pass

    def flush(self):
        # I print() su sys.stdout di norma chiamano flush implicitamente.
        # Lasciamo no-op: l'append è già stato schedulato sul main thread.
        pass

    def isatty(self):
        # Alcune librerie (ad es. tqdm) controllano isatty per cambiare
        # comportamento; rispondiamo False per restare prevedibili.
        return False


class App:
    def __init__(self, root):
        self.root = root
        root.title("PPTX TTS — audio e video per le slide")
        root.geometry("780x740")
        root.minsize(700, 640)

        self._setup_style()

        # bound state
        self.input_pptx_var = tk.StringVar()
        self.input_xlsx_var = tk.StringVar()
        self.output_pptx_var = tk.StringVar()
        self.voice_var = tk.StringVar(value=VOICE_OPTIONS[0][0])
        self.rate_var = tk.IntVar(value=0)
        # Qualità del modello PocketTTS per le voci clonate:
        #   "italian"     -> veloce (default)
        #   "italian_24l" -> qualità più alta, più lenta
        self.pocket_quality_var = tk.StringVar(value="italian")
        # Slide sintetizzate in parallelo con le voci clonate (1 = sequenziale).
        self.clone_workers_var = tk.IntVar(value=1)
        # Riusa da disco l'audio di slide identiche già generate.
        self.use_cache_var = tk.BooleanVar(value=True)
        # Quantizzazione INT8 sperimentale di PocketTTS (spenta di default).
        self.pocket_quantize_var = tk.BooleanVar(value=False)
        self.autoplay_var = tk.BooleanVar(value=True)
        # Avanzamento automatico alla slide successiva a fine audio. Default
        # OFF: all'apertura l'unica opzione attiva è "Riproduci automaticamente".
        self.auto_advance_var = tk.BooleanVar(value=False)
        self.transcode_audio_var = tk.BooleanVar(value=False)
        # Tipo di output: "pptx" (default) o "video".
        self.output_mode_var = tk.StringVar(value="pptx")
        self.resolution_var = tk.StringVar(value="1080p")
        self.subtitles_var = tk.BooleanVar(value=False)
        self.transition_var = tk.BooleanVar(value=False)
        self.preview_text_var = tk.StringVar(value=DEFAULT_PREVIEW_TEXT)

        # internal state
        self._preview_audio_path = None
        self._is_busy = False
        # Stato per percentuale di progresso e calcolo ETA
        self._synthesis_start_time: float | None = None
        self._synthesis_total: int = 0

        self._build_ui()

        # Pulisco i file temp di anteprima alla chiusura della finestra
        root.protocol("WM_DELETE_WINDOW", self._on_close)

        if not PYGAME_INSTALLED:
            self._log(
                f"⚠ Modulo pygame non disponibile ({PYGAME_IMPORT_ERROR}).\n"
                f"L'anteprima vocale è disabilitata; la generazione delle slide funziona ugualmente.\n\n"
            )

    # ----------------------------------------------------------------- style
    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Generate.TButton", font=("Segoe UI", 11, "bold"), padding=10)

    # -------------------------------------------------------------------- ui
    def _build_ui(self):
        # Contenitore scorrevole: un canvas con barra verticale a lato, così
        # tutta la finestra si può far scorrere su e giù anche su schermi bassi.
        container = ttk.Frame(self.root)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, highlightthickness=0, borderwidth=0)
        vbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self._scroll_canvas = canvas

        outer = ttk.Frame(canvas, padding=15)
        outer.columnconfigure(0, weight=1)
        self._scroll_window = canvas.create_window((0, 0), window=outer,
                                                   anchor="nw")

        # La scrollregion segue l'altezza del contenuto; la larghezza del
        # contenuto segue quella del canvas, così i riquadri si allargano bene.
        def _sync_scrollregion(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        outer.bind("<Configure>", _sync_scrollregion)

        def _sync_width(e):
            canvas.itemconfigure(self._scroll_window, width=e.width)
        canvas.bind("<Configure>", _sync_width)

        # Rotellina del mouse (Windows).
        canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # Header
        ttk.Label(
            outer,
            text="PPTX TTS — audio e video per le slide",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 12))

        # Tipo di output (PowerPoint con audio  /  Video MP4)
        out_frame = ttk.Labelframe(outer, text=" Tipo di output ", padding=12)
        out_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        out_frame.columnconfigure(0, weight=1)
        out_frame.columnconfigure(1, weight=1)
        ttk.Radiobutton(
            out_frame, text="PowerPoint con audio",
            variable=self.output_mode_var, value="pptx",
            command=self._on_output_mode_change,
        ).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            out_frame, text="Video (MP4)",
            variable=self.output_mode_var, value="video",
            command=self._on_output_mode_change,
        ).grid(row=0, column=1, sticky="w")
        if not _VIDEO_AVAILABLE:
            # Senza il modulo video (o LibreOffice/PyMuPDF) resta solo il pptx.
            for child in out_frame.winfo_children():
                if isinstance(child, ttk.Radiobutton) and child.cget("value") == "video":
                    child.configure(state="disabled")
            ttk.Label(
                out_frame,
                text="(modulo video non disponibile)",
                foreground="#888", font=("Segoe UI", 9, "italic"),
            ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # 1. Files
        files_frame = ttk.Labelframe(outer, text=" 1. File ", padding=12)
        files_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        files_frame.columnconfigure(1, weight=1)

        self._file_row(files_frame, 0, "Presentazione PowerPoint:",
                       self.input_pptx_var, self._browse_pptx, "Sfoglia…")
        self._file_row(files_frame, 1, "File Excel con script (col. A):",
                       self.input_xlsx_var, self._browse_xlsx, "Sfoglia…")
        self._file_row(files_frame, 2, "Salva risultato come:",
                       self.output_pptx_var, self._browse_output, "Salva…")

        # 2. Voice
        voice_frame = ttk.Labelframe(outer, text=" 2. Voce e velocità ", padding=12)
        voice_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        voice_frame.columnconfigure(0, weight=1)

        # Picker unico delle voci (Microsoft + clonate), ricostruibile.
        voice_picker = ttk.Frame(voice_frame)
        voice_picker.grid(row=0, column=0, sticky="ew")
        voice_picker.columnconfigure(0, weight=1)

        top = ttk.Frame(voice_picker)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text="Scegli la voce:").grid(row=0, column=0, sticky="w")
        self.manage_voices_btn = ttk.Button(
            top, text="Gestisci voci…", command=self._open_voice_manager)
        self.manage_voices_btn.grid(row=0, column=1, sticky="e")
        if not _CLONE_AVAILABLE:
            self.manage_voices_btn.configure(state="disabled")

        self.voice_list_frame = ttk.Frame(voice_picker)
        self.voice_list_frame.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.voice_list_frame.columnconfigure(0, weight=1)
        self._rebuild_voice_list()

        ttk.Label(voice_frame, text="Testo di anteprima:").grid(
            row=1, column=0, sticky="w", pady=(12, 2))
        ttk.Entry(voice_frame, textvariable=self.preview_text_var).grid(
            row=2, column=0, sticky="ew")

        preview_btns = ttk.Frame(voice_frame)
        preview_btns.grid(row=3, column=0, sticky="w", pady=(8, 0))
        self.preview_btn = ttk.Button(
            preview_btns, text="▶ Ascolta anteprima", command=self._on_preview)
        self.preview_btn.grid(row=0, column=0, padx=(0, 6))
        self.stop_btn = ttk.Button(
            preview_btns, text="⏹ Stop", command=self._on_stop_preview)
        self.stop_btn.grid(row=0, column=1)

        # Se pygame non è installato, anteprima e stop non hanno senso: li disabilito
        # subito. (Se pygame c'è ma il mixer fallirà al primo init, lo scopriremo
        # al click di anteprima e mostreremo l'errore allora.)
        if not PYGAME_INSTALLED:
            self.preview_btn.configure(state="disabled")
            self.stop_btn.configure(state="disabled")

        rate_frame = ttk.Frame(voice_frame)
        rate_frame.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        rate_frame.columnconfigure(1, weight=1)
        ttk.Label(rate_frame, text="Velocità:").grid(row=0, column=0, sticky="w")
        self.rate_label = ttk.Label(rate_frame, text="+0%", width=6, anchor="e")
        self.rate_label.grid(row=0, column=2, sticky="e")
        ttk.Scale(rate_frame, from_=-50, to=50, variable=self.rate_var,
                  command=lambda _: self._on_rate_change()).grid(
            row=0, column=1, sticky="ew", padx=10)

        # Qualità del modello PocketTTS (solo per le voci clonate PocketTTS).
        self.pq_frame = ttk.Frame(voice_frame)
        self.pq_frame.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(self.pq_frame, text="Qualità voce clonata (PocketTTS):").grid(
            row=0, column=0, sticky="w")
        ttk.Radiobutton(self.pq_frame, text="Veloce", value="italian",
                        variable=self.pocket_quality_var).grid(
            row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Radiobutton(self.pq_frame, text="Qualità alta (più lenta)",
                        value="italian_24l",
                        variable=self.pocket_quality_var).grid(
            row=0, column=2, sticky="w", padx=(10, 0))
        ttk.Label(
            self.pq_frame,
            text="     ↳ La qualità alta usa il modello «24 layer»: resa migliore, generazione più lenta.",
            foreground="#666", font=("Segoe UI", 9, "italic"),
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Checkbutton(
            self.pq_frame,
            text="Modalità sperimentale INT8 (più veloce, qualità da verificare)",
            variable=self.pocket_quantize_var,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(
            self.pq_frame,
            text="     ↳ Comprime il modello per generare più in fretta su CPU. Prova su qualche slide e confronta: se la voce peggiora, lasciala spenta.",
            foreground="#666", font=("Segoe UI", 9, "italic"), wraplength=560, justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 0))

        # Slide in parallelo (solo per le voci clonate, qualsiasi motore).
        self.parallel_frame = ttk.Frame(voice_frame)
        self.parallel_frame.grid(row=6, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(self.parallel_frame,
                  text="Slide in parallelo:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            self.parallel_frame, width=18, state="readonly",
            textvariable=self.clone_workers_var,
            values=list(range(1, 11)),
        ).grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Label(
            self.parallel_frame,
            text="     ↳ 1 = una alla volta. 2-10 sfrutta più core della CPU: più veloce, ma ogni processo carica una copia del modello (più RAM).",
            foreground="#666", font=("Segoe UI", 9, "italic"), wraplength=560, justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

        # Cache dei risultati (solo per le voci clonate).
        self.cache_frame = ttk.Frame(voice_frame)
        self.cache_frame.grid(row=7, column=0, sticky="ew", pady=(10, 0))
        ttk.Checkbutton(
            self.cache_frame,
            text="Riusa l'audio di slide identiche già generate (cache)",
            variable=self.use_cache_var,
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(self.cache_frame, text="Svuota cache",
                   command=self._on_clear_cache).grid(
            row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Label(
            self.cache_frame,
            text="     ↳ Rigenerando un corso dopo piccole modifiche, le slide invariate sono immediate. Si aggiorna da sé se cambi testo, voce o velocità.",
            foreground="#666", font=("Segoe UI", 9, "italic"), wraplength=560, justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))
        # Ora che pq_frame esiste, imposto la visibilità in base alla voce di
        # default (il primo rebuild è avvenuto prima che questo frame esistesse).
        self._update_pocket_quality_visibility()

        # --- Opzioni che cambiano in base al tipo di output ----------------
        # Un contenitore unico ospita due gruppi sovrapposti: le opzioni del
        # pptx e quelle del video. Ne mostro uno solo, in base alla scelta in
        # cima alla finestra.
        opts_holder = ttk.Frame(voice_frame)
        opts_holder.grid(row=8, column=0, sticky="ew", pady=(12, 0))
        opts_holder.columnconfigure(0, weight=1)

        # ---- Opzioni PowerPoint ----
        self.pptx_opts_frame = ttk.Frame(opts_holder)
        self.pptx_opts_frame.grid(row=0, column=0, sticky="ew")
        self.pptx_opts_frame.columnconfigure(0, weight=1)
        po = self.pptx_opts_frame

        ttk.Checkbutton(
            po,
            text="Riproduci automaticamente all'apertura della slide (consigliato)",
            variable=self.autoplay_var,
            command=self._on_autoplay_toggle,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            po,
            text="     ↳ Stile «Riproduci in background»: l'audio parte da solo, l'icona è nascosta.",
            foreground="#666", font=("Segoe UI", 9, "italic"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        self.auto_advance_chk = ttk.Checkbutton(
            po,
            text="      Avanza alla slide successiva a fine audio",
            variable=self.auto_advance_var,
        )
        self.auto_advance_chk.grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.auto_advance_lbl = ttk.Label(
            po,
            text="     ↳ La presentazione va avanti da sola, slide dopo slide.",
            foreground="#666", font=("Segoe UI", 9, "italic"),
        )
        self.auto_advance_lbl.grid(row=3, column=0, sticky="w", pady=(2, 0))

        ttk.Checkbutton(
            po,
            text="Riconverti l'audio per massima compatibilità PowerPoint (richiede ffmpeg)",
            variable=self.transcode_audio_var,
        ).grid(row=4, column=0, sticky="w", pady=(12, 0))
        ttk.Label(
            po,
            text="     ↳ Riconverte gli MP3 a 44100Hz mono. Default OFF: gli audio Edge TTS\n"
                 "       (24kHz) sono più compatti (~50% in meno) e funzionano già nella maggior\n"
                 "       parte dei casi. Attiva solo se vedi avvisi di compatibilità in PowerPoint.",
            foreground="#666", font=("Segoe UI", 9, "italic"), justify="left",
        ).grid(row=5, column=0, sticky="w", pady=(2, 0))

        # ---- Opzioni Video ----
        self.video_opts_frame = ttk.Frame(opts_holder)
        self.video_opts_frame.grid(row=0, column=0, sticky="ew")
        self.video_opts_frame.columnconfigure(1, weight=1)
        vo = self.video_opts_frame

        ttk.Label(vo, text="Risoluzione:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            vo, textvariable=self.resolution_var, state="readonly",
            width=14, values=["720p", "1080p"],
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))

        ttk.Checkbutton(
            vo, text="Mostra i sottotitoli nel video",
            variable=self.subtitles_var,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(12, 0))
        ttk.Label(
            vo,
            text="     ↳ I sottotitoli vengono comunque salvati anche in un file .srt a fianco al video.",
            foreground="#666", font=("Segoe UI", 9, "italic"),
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 0))

        ttk.Checkbutton(
            vo, text="Dissolvenza incrociata tra le slide",
            variable=self.transition_var,
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(12, 0))
        ttk.Label(
            vo,
            text="     ↳ Transizione morbida tra una slide e l'altra; il parlato resta sequenziale.",
            foreground="#666", font=("Segoe UI", 9, "italic"),
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(2, 0))

        ttk.Label(
            vo,
            text="ℹ Il video richiede LibreOffice installato per disegnare le slide.",
            foreground="#666", font=("Segoe UI", 9, "italic"),
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(12, 0))

        # Stato iniziale: mostro le opzioni giuste e la sotto-opzione autoplay.
        self._on_output_mode_change()

        # 3. Generate
        gen_frame = ttk.Labelframe(outer, text=" 3. Genera ", padding=12)
        gen_frame.grid(row=4, column=0, sticky="nsew")
        gen_frame.columnconfigure(0, weight=1)
        # rowconfigure: ora il log è in row=3 (era row=2 prima dell'aggiunta
        # della label di progresso)
        gen_frame.rowconfigure(3, weight=1)

        self.generate_btn = ttk.Button(
            gen_frame, text="Genera presentazione con audio",
            style="Generate.TButton", command=self._on_generate)
        self.generate_btn.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        # Barra di avanzamento in modalità "determinate" (0-100). Per
        # l'anteprima vocale (operazione molto breve) si passa
        # temporaneamente a "indeterminate". Vedi _set_progress_*.
        self.progress = ttk.Progressbar(
            gen_frame, mode="determinate", maximum=100, value=0)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(0, 4))

        # Etichetta sotto la barra: mostra percentuale corrente e ETA.
        # Esempio: "47% — Sintesi 7/15 — ETA 1m 24s"
        self.progress_label = ttk.Label(
            gen_frame, text="Pronto",
            foreground="#444", font=("Segoe UI", 9))
        self.progress_label.grid(row=2, column=0, sticky="ew", pady=(0, 8))

        self.log_text = scrolledtext.ScrolledText(
            gen_frame, height=10, state="disabled",
            font=("Consolas", 9), wrap="word")
        self.log_text.grid(row=3, column=0, sticky="nsew")

    def _file_row(self, parent, row, label, var, cmd, btn_text):
        ttk.Label(parent, text=label).grid(
            row=row, column=0, sticky="w", pady=4, padx=(0, 8))
        ttk.Entry(parent, textvariable=var).grid(
            row=row, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text=btn_text, command=cmd).grid(
            row=row, column=2, sticky="e", padx=(8, 0), pady=4)

    # --------------------------------------------------------------- voci
    def _rebuild_voice_list(self):
        """Ricostruisce il picker delle voci (Microsoft + clonate). Chiamato
        all'avvio e ogni volta che il pannello di gestione voci cambia qualcosa."""
        for child in self.voice_list_frame.winfo_children():
            child.destroy()

        entries = get_voice_entries()
        valid_ids = {e["voice_id"] for e in entries}
        # Se la voce selezionata non esiste più (es. clone eliminato), torno
        # alla prima disponibile.
        if self.voice_var.get() not in valid_ids and entries:
            self.voice_var.set(entries[0]["voice_id"])

        r = 0
        last_group = None
        for e in entries:
            group = "Le mie voci" if e["is_clone"] else "Voci Microsoft"
            if group != last_group:
                ttk.Label(self.voice_list_frame, text=group,
                          foreground="#888", font=("Segoe UI", 9)).grid(
                    row=r, column=0, sticky="w", pady=((8 if r else 0), 2))
                r += 1
                last_group = group
            text = e["name"]
            if e["is_clone"]:
                text += f"   ·   {e['engine']}"
            ttk.Radiobutton(
                self.voice_list_frame, text=text,
                variable=self.voice_var, value=e["voice_id"],
                command=self._update_pocket_quality_visibility,
            ).grid(row=r, column=0, sticky="w", padx=(8, 0), pady=2)
            r += 1

        # Il selettore "Veloce / Qualità alta" ha senso solo per PocketTTS:
        # lo mostro/nascondo in base alla voce ora selezionata.
        self._update_pocket_quality_visibility()

    def _update_pocket_quality_visibility(self):
        """Mostra il selettore di qualità PocketTTS solo per voci PocketTTS, e il
        selettore 'Slide in parallelo' per qualsiasi voce clonata (locale);
        entrambi restano nascosti per le voci Microsoft."""
        if not hasattr(self, "pq_frame"):
            return
        vid = self.voice_var.get()
        is_clone = _CLONE_AVAILABLE and voice_library.is_clone_voice(vid)
        is_pocket = False
        if is_clone:
            try:
                engine, _ = voice_library.parse_clone_id(vid)
                is_pocket = (engine == voice_library.ENGINE_POCKET)
            except Exception:
                is_pocket = False
        # qualità PocketTTS: solo per PocketTTS
        if is_pocket:
            self.pq_frame.grid()
        else:
            self.pq_frame.grid_remove()
        # slide in parallelo: per ogni voce clonata locale
        if hasattr(self, "parallel_frame"):
            if is_clone:
                self.parallel_frame.grid()
            else:
                self.parallel_frame.grid_remove()
        # cache: per ogni voce clonata locale
        if hasattr(self, "cache_frame"):
            if is_clone:
                self.cache_frame.grid()
            else:
                self.cache_frame.grid_remove()

    def _on_clear_cache(self):
        """Svuota la cache dell'audio clonato, con conferma e riepilogo."""
        if not _CLONE_AVAILABLE:
            return
        try:
            n = voice_clone.clear_cache()
            messagebox.showinfo(
                "Cache svuotata",
                f"Rimossi {n} file audio dalla cache." if n else
                "La cache era già vuota.")
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile svuotare la cache: {e}")

    def _open_voice_manager(self):
        if not _CLONE_AVAILABLE:
            messagebox.showinfo(
                "Voci clonate non disponibili",
                "I moduli per le voci clonate non sono presenti in questa "
                f"cartella.\n({_CLONE_IMPORT_ERROR})")
            return
        VoiceManagerDialog(self.root, on_change=self._rebuild_voice_list)

    def _on_autoplay_toggle(self):
        """Mostra/nasconde la sotto-opzione di avanzamento automatico in base
        allo stato di "Riproduci automaticamente"."""
        if self.autoplay_var.get():
            self.auto_advance_chk.grid()
            self.auto_advance_lbl.grid()
        else:
            # Senza autoplay l'avanzamento automatico non ha senso: lo spengo
            # e nascondo.
            self.auto_advance_var.set(False)
            self.auto_advance_chk.grid_remove()
            self.auto_advance_lbl.grid_remove()

    def _generate_btn_label(self) -> str:
        return ("Genera video" if self.output_mode_var.get() == "video"
                else "Genera presentazione con audio")

    def _on_output_mode_change(self):
        """Adatta la finestra al tipo di output: mostra le opzioni giuste,
        aggiorna il pulsante e l'estensione del file di output."""
        if self.output_mode_var.get() == "video":
            self.pptx_opts_frame.grid_remove()
            self.video_opts_frame.grid()
        else:
            self.video_opts_frame.grid_remove()
            self.pptx_opts_frame.grid()
            self._on_autoplay_toggle()
        # Il pulsante esiste solo dopo la costruzione di gen_frame: alla prima
        # chiamata (durante _build_ui) potrebbe non esserci ancora.
        if hasattr(self, "generate_btn") and not self._is_busy:
            self.generate_btn.configure(text=self._generate_btn_label())
        self._sync_output_extension()

    def _sync_output_extension(self):
        """Allinea l'estensione del file di output al tipo scelto (.mp4/.pptx)."""
        cur = self.output_pptx_var.get().strip()
        if not cur:
            return
        want = ".mp4" if self.output_mode_var.get() == "video" else ".pptx"
        p = Path(cur)
        if p.suffix.lower() != want:
            self.output_pptx_var.set(str(p.with_suffix(want)))

    def _on_mousewheel(self, event):
        """Scorre la finestra con la rotellina. Se il puntatore è sopra il log
        (che ha già la sua barra), lascia scorrere quello."""
        try:
            w = self.root.winfo_containing(event.x_root, event.y_root)
        except (KeyError, tk.TclError):
            # Le Combobox usano popup interni che Tkinter non sempre riesce a
            # convertire in widget Python. In quel caso non facciamo scrollare
            # la finestra principale.
            return "break"
        if w is None:
            return
        node = w
        while node is not None:
            if node is self.log_text:
                return
            node = getattr(node, "master", None)
        self._scroll_canvas.yview_scroll(int(-event.delta / 120), "units")

    def _synthesize_preview(self, text, voice, rate_str, path):
        """Sintetizza l'anteprima con il motore giusto: clone o Microsoft."""
        if _CLONE_AVAILABLE and voice_library.is_clone_voice(voice):
            cv = voice_library.VoiceLibrary().get(voice)
            voice_clone.synthesize_clone(
                cv, text, rate_str, path,
                pocket_variant=self.pocket_quality_var.get())
        else:
            pptx_tts.synthesize_to_file(text, voice, rate_str, path)

    # ------------------------------------------------------------- browsers
    def _browse_pptx(self):
        path = filedialog.askopenfilename(
            title="Seleziona file PowerPoint",
            filetypes=[("PowerPoint", "*.pptx"), ("Tutti i file", "*.*")])
        if path:
            self.input_pptx_var.set(path)
            if not self.output_pptx_var.get():
                p = Path(path)
                if self.output_mode_var.get() == "video":
                    self.output_pptx_var.set(str(p.with_name(f"{p.stem}_video.mp4")))
                else:
                    self.output_pptx_var.set(str(p.with_name(f"{p.stem}_audio.pptx")))

    def _browse_xlsx(self):
        path = filedialog.askopenfilename(
            title="Seleziona file Excel con gli script",
            filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")])
        if path:
            self.input_xlsx_var.set(path)

    def _browse_output(self):
        if self.output_mode_var.get() == "video":
            path = filedialog.asksaveasfilename(
                title="Salva video",
                defaultextension=".mp4",
                filetypes=[("Video MP4", "*.mp4")])
        else:
            path = filedialog.asksaveasfilename(
                title="Salva presentazione con audio",
                defaultextension=".pptx",
                filetypes=[("PowerPoint", "*.pptx")])
        if path:
            self.output_pptx_var.set(path)

    # ----------------------------------------------------------------- rate
    def _on_rate_change(self):
        v = self.rate_var.get()
        self.rate_label.configure(text=f"{'+' if v >= 0 else ''}{v}%")

    def _rate_string(self):
        v = self.rate_var.get()
        return f"{'+' if v >= 0 else ''}{v}%"

    # -------------------------------------------------------------- preview
    def _on_preview(self):
        if self._is_busy:
            return
        if not _ensure_mixer_initialized():
            messagebox.showerror(
                "Audio non disponibile",
                f"Impossibile inizializzare il riproduttore audio: {_mixer_error}\n\n"
                "L'anteprima non è possibile, ma la generazione delle slide funziona comunque.")
            # Da ora in poi disabilito i pulsanti, tanto non servono
            self.preview_btn.configure(state="disabled")
            self.stop_btn.configure(state="disabled")
            return
        text = self.preview_text_var.get().strip()
        if not text:
            messagebox.showwarning("Testo vuoto", "Inserisci un testo da ascoltare.")
            return

        self._is_busy = True
        self.preview_btn.configure(state="disabled", text="Generando…")
        # Per l'anteprima vocale (operazione molto breve) uso modalità
        # indeterminate: non posso stimare quanto durerà visto che è una
        # singola sintesi atomica.
        self._set_progress_indeterminate("Generando anteprima…")

        try:
            pygame.mixer.music.stop()
        except Exception:
            pass

        voice = self.voice_var.get()
        rate_str = self._rate_string()

        def work():
            try:
                if self._preview_audio_path and os.path.exists(self._preview_audio_path):
                    # Su Windows pygame.mixer.music.load() tiene un handle al file:
                    # serve stop + unload PRIMA dell'unlink, altrimenti
                    # PermissionError e il file resta in %TEMP%.
                    try:
                        pygame.mixer.music.stop()
                        if hasattr(pygame.mixer.music, "unload"):
                            pygame.mixer.music.unload()
                    except Exception:
                        pass
                    try:
                        os.unlink(self._preview_audio_path)
                    except Exception:
                        pass
                fd, path = tempfile.mkstemp(suffix=".mp3", prefix="pptx_tts_preview_")
                os.close(fd)
                self._synthesize_preview(text, voice, rate_str, path)
                self._preview_audio_path = path
                self.root.after(0, self._play_preview)
            except Exception as e:
                # Catturo il messaggio SUBITO: dopo l'except la variabile 'e'
                # viene cancellata, e la lambda (eseguita più tardi dal loop
                # tkinter) altrimenti solleverebbe NameError mascherando l'errore
                # vero. Includo anche il tipo per diagnosticare meglio (es. XTTS).
                import traceback as _tb
                msg = f"{type(e).__name__}: {e}"
                self._log("Dettaglio errore anteprima:\n" + _tb.format_exc() + "\n")
                self.root.after(0, lambda m=msg: self._preview_failed(m))

        threading.Thread(target=work, daemon=True).start()

    def _play_preview(self):
        try:
            pygame.mixer.music.load(self._preview_audio_path)
            pygame.mixer.music.play()
        except Exception as e:
            self._log(f"Errore riproduzione: {e}\n")
        finally:
            self._set_progress_idle("Pronto")
            self.preview_btn.configure(state="normal", text="▶ Ascolta anteprima")
            self._is_busy = False

    def _preview_failed(self, err):
        self._set_progress_idle("Pronto")
        self.preview_btn.configure(state="normal", text="▶ Ascolta anteprima")
        self._is_busy = False
        messagebox.showerror("Errore TTS", f"Generazione anteprima fallita:\n{err}")

    def _on_stop_preview(self):
        if _mixer_ready:
            try:
                pygame.mixer.music.stop()
            except Exception:
                pass

    # -------------------------------------------------------------- generate
    def _on_generate(self):
        if self._is_busy:
            return

        self._sync_output_extension()
        in_pptx = self.input_pptx_var.get().strip()
        in_xlsx = self.input_xlsx_var.get().strip()
        out_pptx = self.output_pptx_var.get().strip()

        if not in_pptx or not Path(in_pptx).exists():
            messagebox.showerror("File mancante", "Seleziona un file PowerPoint valido.")
            return
        if not in_xlsx or not Path(in_xlsx).exists():
            messagebox.showerror("File mancante", "Seleziona un file Excel valido.")
            return
        if not out_pptx:
            messagebox.showerror("File mancante", "Specifica dove salvare il risultato.")
            return
        if Path(out_pptx).resolve() == Path(in_pptx).resolve():
            messagebox.showerror("Conflitto", "Il file di output deve essere diverso dall'input.")
            return
        if Path(out_pptx).exists():
            if not messagebox.askyesno(
                    "Sovrascrivi?",
                    f"Il file '{Path(out_pptx).name}' esiste già. Sovrascriverlo?"):
                return

        self._is_busy = True
        self.generate_btn.configure(state="disabled", text="Generazione in corso…")
        # Durante la generazione blocco l'anteprima: edge-tts viene martellato
        # dal motore, non ha senso lanciare richieste in parallelo dalla UI.
        self._set_preview_enabled(False)
        # Se l'anteprima sta suonando, la fermo
        if _mixer_ready:
            try:
                pygame.mixer.music.stop()
            except Exception:
                pass
        # Reset progress: barra a 0%, label "Avvio..." (e niente più animazione
        # indeterminate). I successivi aggiornamenti arriveranno dal motore via
        # progress_callback.
        self._set_progress_determinate(0, "Avvio…")
        self._synthesis_start_time = None
        self._synthesis_total = 0
        self._clear_log()

        voice = self.voice_var.get()
        rate_str = self._rate_string()
        autoplay = self.autoplay_var.get()
        auto_advance = self.auto_advance_var.get()
        transcode_audio = self.transcode_audio_var.get()
        output_mode = self.output_mode_var.get()
        resolution = self.resolution_var.get()
        subtitles = self.subtitles_var.get()
        transition = self.transition_var.get()
        pocket_variant = self.pocket_quality_var.get()
        clone_workers = max(1, min(10, int(self.clone_workers_var.get() or 1)))
        use_cache = bool(self.use_cache_var.get())
        pocket_quantize = bool(self.pocket_quantize_var.get())

        # Marshalling progress events from the worker thread (where the engine
        # runs) onto the Tk main thread. lambda+default-arg captures the event
        # by value to avoid the classic late-binding bug.
        def _progress_cb(event):
            self.root.after(0, lambda e=event: self._handle_progress_event(e))

        def work():
            redirector = StdoutRedirector(self.log_text)
            # Salvo lo stdout originale e lo sostituisco a livello di
            # processo. Non uso contextlib.redirect_stdout perché non è
            # thread-safe e in alcuni casi su Windows gli output di asyncio
            # / subprocess (ffmpeg) non vengono catturati. Sostituendo
            # sys.stdout direttamente, ogni print() del motore trova subito
            # il redirector. Inoltre forzo write_through-like: dopo ogni
            # print il redirector chiama già self._append via Tk.after,
            # quindi il widget si aggiorna appena Tk fa il prossimo idle.
            old_stdout = sys.stdout
            sys.stdout = redirector
            try:
                pptx_tts.process(
                    input_pptx=in_pptx,
                    scripts_xlsx=in_xlsx,
                    output_pptx=out_pptx,
                    voice=voice,
                    rate=rate_str,
                    autoplay=autoplay,
                    auto_advance=auto_advance,
                    progress_callback=_progress_cb,
                    transcode_audio=transcode_audio,
                    output_mode=output_mode,
                    resolution=resolution,
                    subtitles=subtitles,
                    transition=transition,
                    pocket_variant=pocket_variant,
                    clone_workers=clone_workers,
                    use_cache=use_cache,
                    pocket_quantize=pocket_quantize,
                )
                self.root.after(0, lambda: self._generate_done(out_pptx))
            except Exception as e:
                err = str(e)
                self.root.after(0, lambda: self._generate_failed(err))
            finally:
                sys.stdout = old_stdout

        threading.Thread(target=work, daemon=True).start()

    def _generate_done(self, out_path):
        self._set_progress_determinate(100, "Completato")
        self.generate_btn.configure(state="normal", text=self._generate_btn_label())
        self._set_preview_enabled(True)
        self._is_busy = False
        out_p = Path(out_path)
        report_path = out_p.with_name(f"{out_p.stem}_durate.txt")
        is_video = self.output_mode_var.get() == "video"
        what = "Video generato" if is_video else "Presentazione generata"
        open_q = ("Vuoi aprire ora il video?" if is_video
                  else "Vuoi aprire ora la presentazione?")
        msg = (
            f"{what} con successo:\n  {out_path}\n\n"
            f"Report durate:\n  {report_path}\n\n"
            f"{open_q}"
        )
        if messagebox.askyesno("Fatto!", msg):
            self._open_file(out_path)

    def _generate_failed(self, err):
        self._set_progress_idle("Errore")
        self.generate_btn.configure(state="normal", text=self._generate_btn_label())
        self._set_preview_enabled(True)
        self._is_busy = False
        self._log(f"\n⚠ ERRORE: {err}\n")
        messagebox.showerror("Errore", f"Generazione fallita:\n{err}")

    # ------------------------------------------------------ progress helpers
    def _set_progress_indeterminate(self, label_text: str = ""):
        """Switch the progress bar to the spinning 'indeterminate' mode.
        Used for short, unmeasurable operations like the voice preview."""
        try:
            self.progress.stop()
        except Exception:
            pass
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)
        self.progress_label.configure(text=label_text)

    def _set_progress_determinate(self, value: float, label_text: str):
        """Switch to a measurable 0-100 progress with a status label."""
        try:
            self.progress.stop()
        except Exception:
            pass
        self.progress.configure(mode="determinate", value=max(0, min(100, value)))
        self.progress_label.configure(text=label_text)

    def _set_progress_idle(self, label_text: str = "Pronto"):
        """Reset the bar to 0 and show an idle/ready label."""
        self._set_progress_determinate(0, label_text)

    @staticmethod
    def _format_eta(seconds: float) -> str:
        """Format an ETA in seconds. Examples: '23s', '1m 04s', '12m 30s'."""
        seconds = max(0, int(round(seconds)))
        if seconds < 60:
            return f"{seconds}s"
        return f"{seconds // 60}m {seconds % 60:02d}s"

    def _handle_progress_event(self, event: dict):
        """Called on the Tk main thread for every progress milestone from
        the engine. Updates the bar value and the percentage/ETA label.

        Allocation of the 0-100 bar:
          - loading             →  1%
          - synthesis_start     →  3%
          - synthesis_progress  →  3% .. 93% (90 percentage points = the
                                   dominant phase)
          - synthesis_end       → 93%
          - embedding_progress  → 93% .. 98%
          - saving              → 98%
          - done                → 100%
        """
        stage = event.get("stage")

        if stage == "loading":
            self._set_progress_determinate(1, "1% — Apertura presentazione…")

        elif stage == "synthesis_start":
            total = event.get("total", 0)
            self._synthesis_start_time = time.monotonic()
            self._synthesis_total = total
            self._set_progress_determinate(
                3, f"3% — Sintesi audio 0/{total}…")

        elif stage == "synthesis_progress":
            done = event.get("done", 0)
            total = event.get("total", self._synthesis_total or 1)
            pct = 3 + 90 * (done / total) if total else 3
            # ETA: regola del tre. Saltiamo il calcolo finché abbiamo solo 1
            # campione (troppo rumoroso). Aggiungiamo un piccolo buffer
            # costante (~5s) per tenere conto delle fasi finali (embedding,
            # save, scrittura report).
            eta_label = "calcolo…"
            if self._synthesis_start_time is not None and done >= 2:
                elapsed = time.monotonic() - self._synthesis_start_time
                remaining = total - done
                if remaining > 0 and elapsed > 0:
                    eta_secs = remaining * elapsed / done + 5
                    eta_label = self._format_eta(eta_secs)
                else:
                    eta_label = self._format_eta(5)
            self._set_progress_determinate(
                pct, f"{int(pct)}% — Sintesi {done}/{total} — ETA {eta_label}")

        elif stage == "synthesis_end":
            self._set_progress_determinate(93, "93% — Sintesi completata")

        elif stage == "embedding_progress":
            done = event.get("done", 0)
            total = event.get("total", 1)
            pct = 93 + 5 * (done / total) if total else 93
            self._set_progress_determinate(
                pct, f"{int(pct)}% — Incorporamento audio {done}/{total}")

        elif stage == "render_start":
            self._set_progress_determinate(94, "94% — Rendering slide…")

        elif stage == "clip_progress":
            done = event.get("done", 0)
            total = event.get("total", 1)
            pct = 94 + 4 * (done / total) if total else 94
            self._set_progress_determinate(
                pct, f"{int(pct)}% — Creazione video {done}/{total}")

        elif stage == "muxing":
            self._set_progress_determinate(99, "99% — Montaggio del video…")

        elif stage == "subtitles":
            self._set_progress_determinate(99, "99% - Sottotitoli nel video...")

        elif stage == "saving":
            self._set_progress_determinate(98, "98% — Salvataggio file…")

        elif stage == "done":
            self._set_progress_determinate(100, "100% — Completato")

    # ----------------------------------------------------------------- misc
    def _set_preview_enabled(self, enabled: bool):
        """Abilita o disabilita i pulsanti dell'anteprima.
        Se pygame non è disponibile, restano comunque disabilitati."""
        if not PYGAME_INSTALLED or _mixer_ready is False:
            self.preview_btn.configure(state="disabled")
            self.stop_btn.configure(state="disabled")
            return
        new_state = "normal" if enabled else "disabled"
        self.preview_btn.configure(state=new_state)
        self.stop_btn.configure(state=new_state)

    def _on_close(self):
        """Pulizia all'uscita: ferma l'audio e cancella il file MP3 temporaneo."""
        if _mixer_ready:
            try:
                pygame.mixer.music.stop()
                if hasattr(pygame.mixer.music, "unload"):
                    pygame.mixer.music.unload()
            except Exception:
                pass
            try:
                pygame.mixer.quit()
            except Exception:
                pass
        if self._preview_audio_path and os.path.exists(self._preview_audio_path):
            try:
                os.unlink(self._preview_audio_path)
            except Exception:
                pass
        self.root.destroy()

    def _log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _open_file(self, path):
        try:
            if platform.system() == "Windows":
                os.startfile(path)  # type: ignore[attr-defined]
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire il file: {e}")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
