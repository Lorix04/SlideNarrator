"""
voice_manager.py — Finestra di gestione delle voci clonate (Tkinter)

Apre un dialogo dove l'utente può:
- vedere le voci clonate salvate (nome, motore, durata del campione)
- aggiungerne una nuova importando un file audio OPPURE registrandosi al volo
- rinominarle ed eliminarle

Si appoggia a voice_library (storage) e voice_clone (per sapere quali motori
sono installati). La registrazione usa sounddevice, importato in modo pigro:
se non è installato, il pulsante "Registra…" lo segnala senza far crashare nulla.

Uso dalla GUI principale:
    from voice_manager import VoiceManagerDialog
    VoiceManagerDialog(parent_window, on_change=callback_per_aggiornare_il_picker)
"""

from __future__ import annotations

import tempfile
import threading
import time
import wave
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import voice_library as vl
import voice_clone


RECORD_SAMPLE_RATE_HZ = 24000  # mono, coerente col formato di riferimento


# ----------------------------------------------------------------------------
# Registratore audio (Toplevel) — opzionale, richiede sounddevice
# ----------------------------------------------------------------------------
class _RecordDialog(tk.Toplevel):
    """Piccolo registratore start/stop. Al termine, self.result contiene il
    path del wav registrato (o None se annullato)."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Registra la tua voce")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result: str | None = None

        self._sd = None
        self._stream = None
        self._frames: list = []
        self._recording = False
        self._t0 = 0.0

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(
            frm,
            text="Leggi un testo naturale per circa un minuto.\n"
                 "Ambiente silenzioso, microfono vicino, ritmo normale.",
            justify="left",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        self._time_lbl = ttk.Label(frm, text="00:00", font=("Segoe UI", 16))
        self._time_lbl.grid(row=1, column=0, columnspan=2, pady=(0, 12))

        self._rec_btn = ttk.Button(frm, text="● Registra", command=self._toggle)
        self._rec_btn.grid(row=2, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(frm, text="Annulla", command=self._cancel).grid(
            row=2, column=1, sticky="ew")

        self._save_btn = ttk.Button(frm, text="Usa registrazione",
                                    command=self._accept, state="disabled")
        self._save_btn.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self._ensure_sounddevice()

    def _ensure_sounddevice(self) -> bool:
        if self._sd is not None:
            return True
        try:
            import sounddevice as sd
            self._sd = sd
            return True
        except Exception:
            messagebox.showerror(
                "Registrazione non disponibile",
                "Per registrare serve il pacchetto 'sounddevice'.\n"
                "Installa con:  pip install sounddevice\n\n"
                "In alternativa importa un file audio già pronto.",
                parent=self,
            )
            self._rec_btn.configure(state="disabled")
            return False

    def _toggle(self):
        if not self._recording:
            self._start()
        else:
            self._stop()

    def _start(self):
        if not self._ensure_sounddevice():
            return
        self._frames = []
        self._recording = True
        self._t0 = time.monotonic()

        def cb(indata, frames, time_info, status):
            # Copio i campioni: indata è riutilizzato dal driver audio.
            self._frames.append(indata.copy())

        self._stream = self._sd.InputStream(
            samplerate=RECORD_SAMPLE_RATE_HZ, channels=1,
            dtype="int16", callback=cb,
        )
        self._stream.start()
        self._rec_btn.configure(text="■ Ferma")
        self._save_btn.configure(state="disabled")
        self._tick()

    def _tick(self):
        if not self._recording:
            return
        elapsed = int(time.monotonic() - self._t0)
        self._time_lbl.configure(text=f"{elapsed // 60:02d}:{elapsed % 60:02d}")
        self.after(250, self._tick)

    def _stop(self):
        self._recording = False
        try:
            self._stream.stop()
            self._stream.close()
        except Exception:
            pass
        self._rec_btn.configure(text="● Registra")
        # Salvo il wav.
        if not self._frames:
            return
        import numpy as np
        data = np.concatenate(self._frames, axis=0).reshape(-1)
        fd, path = tempfile.mkstemp(suffix=".wav", prefix="voce_reg_")
        import os
        os.close(fd)
        with wave.open(path, "wb") as w:
            w.setnchannels(1)
            w.setsampwidth(2)
            w.setframerate(RECORD_SAMPLE_RATE_HZ)
            w.writeframes(data.astype("<i2").tobytes())
        self._recorded_path = path
        self._save_btn.configure(state="normal")

    def _accept(self):
        self.result = getattr(self, "_recorded_path", None)
        self._cleanup_and_close()

    def _cancel(self):
        self.result = None
        self._cleanup_and_close()

    def _cleanup_and_close(self):
        if self._recording:
            self._stop()
        self.grab_release()
        self.destroy()


# ----------------------------------------------------------------------------
# Dialogo "aggiungi voce" (nome + motore + sorgente audio)
# ----------------------------------------------------------------------------
class _AddVoiceDialog(tk.Toplevel):
    """Raccoglie nome, motore e campione audio, poi crea la voce in libreria.
    Al termine, self.created contiene la ClonedVoice creata (o None)."""

    def __init__(self, parent, library: vl.VoiceLibrary, availability: dict):
        super().__init__(parent)
        self.title("Aggiungi voce clonata")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.library = library
        self.availability = availability
        self.created = None
        self._source_path: str | None = None

        self.name_var = tk.StringVar()
        # Default motore: il primo installato, altrimenti pocket.
        default_engine = next(
            (k for k in vl.SUPPORTED_ENGINES if availability.get(k, (False, ""))[0]),
            vl.ENGINE_POCKET,
        )
        self.engine_var = tk.StringVar(value=default_engine)

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Nome voce:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.name_var, width=32).grid(
            row=0, column=1, columnspan=2, sticky="ew", pady=4)

        ttk.Label(frm, text="Motore:").grid(row=1, column=0, sticky="nw", pady=(8, 0))
        eng_frm = ttk.Frame(frm)
        eng_frm.grid(row=1, column=1, columnspan=2, sticky="w", pady=(8, 0))
        for engine in vl.SUPPORTED_ENGINES:
            installed, _msg = availability.get(engine, (False, ""))
            label = vl.ENGINE_LABELS.get(engine, engine)
            if not installed:
                label += "  (non installato)"
            ttk.Radiobutton(
                eng_frm, text=label, variable=self.engine_var, value=engine,
                state=("normal" if installed else "disabled"),
            ).pack(anchor="w")

        ttk.Label(frm, text="Campione audio:").grid(
            row=2, column=0, sticky="w", pady=(8, 0))
        self.source_lbl = ttk.Label(frm, text="(nessuno)", foreground="#666")
        self.source_lbl.grid(row=2, column=1, columnspan=2, sticky="w", pady=(8, 0))

        src_btns = ttk.Frame(frm)
        src_btns.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Button(src_btns, text="Scegli file…", command=self._choose_file).pack(
            side="left", padx=(0, 6))
        ttk.Button(src_btns, text="Registra…", command=self._record).pack(side="left")

        ttk.Label(
            frm,
            text="Suggerimento: ~1 minuto di parlato pulito dà i risultati migliori.",
            foreground="#666", font=("Segoe UI", 9, "italic"),
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=3, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self._cancel).pack(
            side="right", padx=(6, 0))
        ttk.Button(btns, text="Salva voce", command=self._save).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _set_source(self, path: str | None):
        self._source_path = path
        self.source_lbl.configure(
            text=Path(path).name if path else "(nessuno)",
            foreground="#000" if path else "#666",
        )

    def _choose_file(self):
        path = filedialog.askopenfilename(
            parent=self, title="Scegli un campione audio",
            filetypes=[("Audio", "*.wav *.mp3 *.m4a *.flac *.ogg"),
                       ("Tutti i file", "*.*")],
        )
        if path:
            self._set_source(path)

    def _record(self):
        dlg = _RecordDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            self._set_source(dlg.result)

    def _save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Nome mancante", "Inserisci un nome.", parent=self)
            return
        if not self._source_path:
            messagebox.showwarning("Campione mancante",
                                   "Scegli un file o registra un campione.",
                                   parent=self)
            return
        try:
            self.created = self.library.add_voice(
                name=name, engine=self.engine_var.get(),
                source_audio=self._source_path,
            )
        except vl.VoiceLibraryError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return
        self.grab_release()
        self.destroy()

    def _cancel(self):
        self.created = None
        self.grab_release()
        self.destroy()


# ----------------------------------------------------------------------------
# Finestra principale di gestione voci
# ----------------------------------------------------------------------------
class VoiceManagerDialog(tk.Toplevel):
    """Elenco delle voci clonate con aggiungi / rinomina / elimina."""

    def __init__(self, parent, on_change=None):
        super().__init__(parent)
        self.title("Gestisci voci clonate")
        self.minsize(460, 320)
        self.transient(parent)
        self.grab_set()
        self.on_change = on_change
        self.library = vl.VoiceLibrary()
        self.availability = voice_clone.available_backends()

        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        # Stato motori installati
        self._engines_lbl = ttk.Label(outer, foreground="#444",
                                       font=("Segoe UI", 9))
        self._engines_lbl.grid(row=0, column=0, sticky="w", pady=(0, 8))
        self._update_engines_label()

        # Elenco voci (contenitore ricostruibile)
        self.list_frame = ttk.Labelframe(outer, text=" Voci salvate ", padding=8)
        self.list_frame.grid(row=1, column=0, sticky="nsew")
        self.list_frame.columnconfigure(0, weight=1)

        # Pulsanti
        btns = ttk.Frame(outer)
        btns.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(btns, text="+ Aggiungi voce…", command=self._add_voice).pack(
            side="left")
        ttk.Button(btns, text="Chiudi", command=self._close).pack(side="right")

        self._rebuild_list()
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _update_engines_label(self):
        parts = []
        for engine in vl.SUPPORTED_ENGINES:
            installed = self.availability.get(engine, (False, ""))[0]
            label = vl.ENGINE_LABELS.get(engine, engine)
            parts.append(f"{label}: {'installato' if installed else 'non installato'}")
        self._engines_lbl.configure(text="Motori — " + "   ·   ".join(parts))

    def _rebuild_list(self):
        for child in self.list_frame.winfo_children():
            child.destroy()

        voices = self.library.list_voices()
        if not voices:
            ttk.Label(self.list_frame,
                      text="Nessuna voce clonata. Aggiungine una col pulsante qui sotto.",
                      foreground="#666").grid(row=0, column=0, sticky="w", pady=6)
            return

        for i, v in enumerate(voices):
            row = ttk.Frame(self.list_frame)
            row.grid(row=i, column=0, sticky="ew", pady=3)
            row.columnconfigure(0, weight=1)

            dur = f"{int(v.duration_s)}s"
            short = "  ⚠ campione corto" if v.is_short else ""
            ttk.Label(
                row, text=f"{v.name}   ·   {v.engine_label}   ·   rif. {dur}{short}",
            ).grid(row=0, column=0, sticky="w")

            ttk.Button(row, text="Rinomina",
                       command=lambda s=v.slug, n=v.name: self._rename(s, n)).grid(
                row=0, column=1, padx=(6, 0))
            ttk.Button(row, text="Elimina",
                       command=lambda s=v.slug, n=v.name: self._delete(s, n)).grid(
                row=0, column=2, padx=(6, 0))

    def _add_voice(self):
        dlg = _AddVoiceDialog(self, self.library, self.availability)
        self.wait_window(dlg)
        if dlg.created is not None:
            self._rebuild_list()
            self._notify()

    def _rename(self, slug: str, current_name: str):
        from tkinter import simpledialog
        new_name = simpledialog.askstring(
            "Rinomina voce", "Nuovo nome:", initialvalue=current_name, parent=self)
        if not new_name:
            return
        try:
            self.library.rename_voice(slug, new_name)
        except vl.VoiceLibraryError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return
        self._rebuild_list()
        self._notify()

    def _delete(self, slug: str, name: str):
        if not messagebox.askyesno(
                "Eliminare la voce?",
                f"Eliminare definitivamente «{name}»? L'operazione non è reversibile.",
                parent=self):
            return
        try:
            self.library.delete_voice(slug)
        except vl.VoiceLibraryError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return
        self._rebuild_list()
        self._notify()

    def _notify(self):
        if callable(self.on_change):
            try:
                self.on_change()
            except Exception:
                pass

    def _close(self):
        self.grab_release()
        self.destroy()
