"""Slide Narrator — interfaccia moderna.

La logica operativa e i callback consolidati restano in
``slide_narrator_gui_legacy.py``. Questa classe ne ridisegna completamente la GUI
con navigazione laterale, procedura guidata, pagina Fix dedicata, temi e
pannello tecnico richiudibile.

Il progetto usa ttkbootstrap 2.x quando disponibile. È previsto un fallback
su ttk nativo per mostrare un errore comprensibile anche prima
­dell'installazione completa.
"""
from __future__ import annotations

import json
import ntpath
import os
import platform
import re
import zipfile
from xml.etree import ElementTree as ET
import queue
import shutil
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


try:
    from PIL import Image, ImageDraw, ImageTk
    _PIL_AVAILABLE = True
except Exception:
    Image = ImageDraw = ImageTk = None  # type: ignore[assignment]
    _PIL_AVAILABLE = False

try:
    import ttkbootstrap as tb
    _BOOTSTRAP_AVAILABLE = True
except Exception:
    tb = None  # type: ignore[assignment]
    _BOOTSTRAP_AVAILABLE = False

import slide_narrator
import slide_narrator_batch
import slide_narrator_gui_legacy as legacy

# Re-export per compatibilità con eventuali integrazioni esistenti.
StdoutRedirector = legacy.StdoutRedirector
VOICE_OPTIONS = legacy.VOICE_OPTIONS
DEFAULT_PREVIEW_TEXT = legacy.DEFAULT_PREVIEW_TEXT
get_voice_entries = legacy.get_voice_entries


APP_NAME = "Slide Narrator"
APP_VERSION = "2.9"
BRAND = "#19769A"
BRAND_DARK = "#11556F"
SIDEBAR_LIGHT = "#123E4D"
SIDEBAR_DARK = "#0A2631"
SETTINGS_DIR = Path.home() / ".slide_narrator"
SETTINGS_FILE = SETTINGS_DIR / "gui_settings.json"


def _resource_path(*parts: str) -> Path:
    """Trova una risorsa sia dal sorgente sia da un eventuale bundle PyInstaller."""
    bundle_root = getattr(sys, "_MEIPASS", None)
    base = Path(bundle_root) if bundle_root else Path(__file__).resolve().parent
    return base.joinpath(*parts)


APP_ICON_PNG = _resource_path("assets", "slide_narrator_icon.png")
APP_ICON_SIDEBAR_PNG = _resource_path("assets", "slide_narrator_icon_sidebar.png")
APP_ICON_ICO = _resource_path("assets", "slide_narrator_icon.ico")


AI_SCRIPT_PROMPT = (
    "Leggi tutte le pagine del file in allegato. Genera un file .xlsx inserendo lo script in ITALIANO "
    "per una voce TTS che spiega tutte le pagine, con tono professionale e incoraggiante, usando un "
    "registro da formazione aziendale. Per OGNI pagina lo script deve essere intorno alle 130 parole. "
    "NON usare niente di estetico o titoli nel foglio. DEVE essere tutto il testo nella cella A e poi "
    "per ogni pagina ci si sposta su ogni numero (ES: A1: Script prima pagina, A2: Script seconda pagina)\n"
    "RUOLO: Sei un tutor didattico autorevole e pragmatico con impostazione tradizionale. Parli in modo "
    "discorsivo, lineare e rassicurante. NON porre domande. NON chiedere conferme. NON auto-referenze. "
    "NON inserire il titolo di ogni pagina. Limitati a spiegare in modo chiaro e consequenziale. NON dire "
    "all’inizio “Qui”, “Questa pagina”, “Con questa pagina”, “In questa parte” all’inizio di ogni script "
    "di ogni pagina. La PRIMA pagina dai il benvenuto ai Corsisti. Fai le ricerche sul WEB sugli argomenti "
    "di cui tratti per confermare che quello che stai dicendo sia coretto e aggiornato.\n"
)


def _path_identity(value: str | os.PathLike[str]) -> str:
    """Restituisce una chiave stabile anche con slash Windows differenti.

    I filedialog Tk su Windows possono restituire ``C:/...`` mentre
    ``Path``/``str`` restituisce ``C:\\...``. Il confronto letterale lasciava
    quindi il menu fogli bloccato nello stato "Lettura...".
    """
    raw = os.path.expandvars(os.path.expanduser(str(value).strip()))
    if not raw:
        return ""
    if re.match(r"^[A-Za-z]:[\\/]", raw) or raw.startswith(("\\\\", "//")):
        return ntpath.normcase(ntpath.normpath(raw.replace("/", "\\")))
    return os.path.normcase(os.path.abspath(os.path.normpath(raw)))


def _read_excel_sheet_names_fast(path: Path) -> list[str]:
    """Legge i nomi dei fogli direttamente dal pacchetto XLSX/XLSM.

    Per ottenere i soli nomi non serve inizializzare l'intero workbook con
    openpyxl: ``xl/workbook.xml`` contiene già l'elenco ordinato dei fogli.
    Questo approccio è molto più rapido sui file grandi e sui percorsi cloud.
    """
    with zipfile.ZipFile(path, "r") as archive:
        try:
            payload = archive.read("xl/workbook.xml")
        except KeyError as exc:
            raise ValueError("Il file non contiene xl/workbook.xml e non sembra un Excel valido.") from exc
    try:
        root = ET.fromstring(payload)
    except ET.ParseError as exc:
        raise ValueError("Il file workbook.xml non è leggibile.") from exc
    names: list[str] = []
    for element in root.iter():
        if element.tag.rsplit("}", 1)[-1] != "sheet":
            continue
        name = element.attrib.get("name")
        if name is not None:
            names.append(name)
    if not names:
        raise ValueError("Nessun foglio Excel trovato nel file.")
    return names


def _enable_windows_dpi_awareness() -> None:
    """Abilita Per-Monitor DPI Awareness v2 prima della creazione di Tk."""
    if platform.system() != "Windows":
        return
    try:
        import ctypes
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            pass


def _set_windows_app_user_model_id() -> None:
    """Assegna un'identità propria all'app nella barra delle applicazioni."""
    if platform.system() != "Windows":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "Lorix04.SlideNarrator"
        )
    except Exception:
        pass


def _center_window(root, width: int = 1280, height: int = 820) -> None:
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    width = min(width, max(980, sw - 80))
    height = min(height, max(680, sh - 80))
    x = max(0, (sw - width) // 2)
    y = max(0, (sh - height) // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")


def _load_preferences() -> dict:
    defaults = {
        "theme_mode": "system",
        "history_folder": str(Path.cwd()),
        "start_page": "generate",
        "show_log_on_error": True,
    }
    try:
        saved = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        if isinstance(saved, dict):
            defaults.update(saved)
    except Exception:
        pass
    return defaults


def _system_prefers_dark() -> bool:
    """Rilevazione prudente della preferenza scura del sistema operativo."""
    if platform.system() == "Windows":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            )
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return int(value) == 0
        except Exception:
            return False
    if platform.system() == "Darwin":
        try:
            proc = subprocess.run(
                ["defaults", "read", "-g", "AppleInterfaceStyle"],
                capture_output=True, text=True, timeout=2,
            )
            return proc.returncode == 0 and "dark" in proc.stdout.lower()
        except Exception:
            return False
    desktop = " ".join(
        str(os.environ.get(k, ""))
        for k in ("GTK_THEME", "COLORFGBG", "XDG_CURRENT_DESKTOP")
    ).lower()
    return "dark" in desktop


def _theme_for_mode(mode: str) -> str:
    dark = mode == "dark" or (mode == "system" and _system_prefers_dark())
    return "bootstrap-dark" if dark else "bootstrap-light"


class ScrollablePage(ttk.Frame):
    """Pagina scorrevole ottimizzata per il ridimensionamento della finestra.

    Tk genera molti eventi ``<Configure>`` mentre l'utente trascina un bordo.
    La pagina accorpa gli eventi ravvicinati, aggiorna soltanto la pagina attiva
    e modifica canvas/scrollbar solo quando le dimensioni cambiano davvero.
    """

    MAX_CONTENT_WIDTH = 1180
    RESIZE_SETTLE_MS = 220
    REGION_SETTLE_MS = 40

    def __init__(self, master, app: "App"):
        super().__init__(master, style="App.TFrame")
        self.app = app
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self._on_yview)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.body = ttk.Frame(self.canvas, style="App.TFrame", padding=(24, 16, 24, 28))
        self.window_id = self.canvas.create_window((0, 0), window=self.body, anchor="n")
        self.body.columnconfigure(0, weight=1)

        self._active = False
        self._pending_canvas_width: int | None = None
        self._width_after_id = None
        self._region_after_id = None
        self._activate_after_id = None
        self._last_window_width: int | None = None
        self._last_window_x: float | None = None
        self._last_scrollregion = None
        self._scrollbar_visible = False
        self._last_yview = (0.0, 1.0)
        self._resize_suspended = False
        self._resize_hint_id = None
        self._resize_saved_yview = 0.0

        self.body.bind("<Configure>", self._queue_region_sync)
        self.canvas.bind("<Configure>", self._queue_width_sync)
        self.refresh_colors()

    def _cancel_scheduled(self, attr: str) -> None:
        job = getattr(self, attr, None)
        if job is None:
            return
        try:
            self.after_cancel(job)
        except Exception:
            pass
        setattr(self, attr, None)

    def set_active(self, active: bool) -> None:
        """Abilita gli aggiornamenti geometrici soltanto per la pagina visibile."""
        active = bool(active)
        if self._active == active:
            if active:
                self._queue_width_sync()
                self._queue_region_sync()
            return
        self._active = active
        if not active:
            self._cancel_scheduled("_activate_after_id")
            self._cancel_scheduled("_width_after_id")
            self._cancel_scheduled("_region_after_id")
            self._resize_suspended = False
            try:
                if self._resize_hint_id is not None:
                    self.canvas.itemconfigure(self._resize_hint_id, state="hidden")
                self.canvas.itemconfigure(self.window_id, state="normal")
            except Exception:
                pass
            return
        self._cancel_scheduled("_activate_after_id")
        try:
            self._activate_after_id = self.after_idle(self._activate_layout)
        except Exception:
            self._activate_after_id = None
            self._activate_layout()

    def _activate_layout(self) -> None:
        self._activate_after_id = None
        if not self._active:
            return
        self._queue_width_sync(immediate=True)
        self._queue_region_sync(immediate=True)

    def _on_yview(self, first, last):
        # Il callback della scrollbar può essere invocato molte volte durante un
        # resize. Aggiorniamo il thumb, ma la visibilità viene decisa in un solo
        # punto (_apply_region_sync) per evitare continui grid/grid_remove.
        self.scrollbar.set(first, last)
        try:
            self._last_yview = (float(first), float(last))
        except Exception:
            self._last_yview = (0.0, 1.0)

    def _set_scrollbar_visible(self, visible: bool) -> None:
        visible = bool(visible and self._active)
        if visible == self._scrollbar_visible:
            return
        self._scrollbar_visible = visible
        if visible:
            self.scrollbar.grid(row=0, column=1, sticky="ns")
        else:
            self.scrollbar.grid_remove()
            try:
                self.canvas.yview_moveto(0)
            except Exception:
                pass

    def _queue_region_sync(self, _event=None, *, immediate: bool = False):
        if not self._active or self._resize_suspended:
            return
        self._cancel_scheduled("_region_after_id")
        delay = 0 if immediate else self.REGION_SETTLE_MS
        try:
            self._region_after_id = self.after(delay, self._apply_region_sync)
        except Exception:
            self._region_after_id = None

    def _apply_region_sync(self):
        self._region_after_id = None
        if not self._active or not self.winfo_exists():
            return
        bbox = self.canvas.bbox("all")
        if bbox:
            normalised = tuple(int(v) for v in bbox)
            if normalised != self._last_scrollregion:
                self.canvas.configure(scrollregion=normalised)
                self._last_scrollregion = normalised
        try:
            content_h = max(self.body.winfo_reqheight(), (bbox[3] - bbox[1]) if bbox else 0)
            viewport_h = self.canvas.winfo_height()
            self._set_scrollbar_visible(content_h > viewport_h + 2)
        except Exception:
            pass

    def _queue_width_sync(self, event=None, *, immediate: bool = False):
        if event is not None:
            try:
                self._pending_canvas_width = int(event.width)
            except Exception:
                self._pending_canvas_width = None
        if not self._active:
            return
        if self._resize_suspended and not immediate:
            # Durante il trascinamento aggiorniamo solo il messaggio leggero;
            # il complesso albero di widget viene ridimensionato una volta sola.
            self._position_resize_hint()
            return
        # Debounce reale: nessun relayout del contenuto finché gli eventi di
        # ridimensionamento continuano ad arrivare.
        self._cancel_scheduled("_width_after_id")
        delay = 0 if immediate else self.RESIZE_SETTLE_MS
        try:
            self._width_after_id = self.after(delay, self._apply_width_sync)
        except Exception:
            self._width_after_id = None

    def _apply_window_geometry(self, available: int | float) -> bool:
        """Applica subito la geometria del contenuto senza attendere ``<Configure>``.

        Quando una pagina nascosta viene rimappata, il canvas non ha ancora
        ricevuto il nuovo evento ``<Configure>`` e il suo window-item conserva
        per un fotogramma la larghezza richiesta dai widget interni. Questo
        produce il visibile effetto "pagina stretta -> pagina corretta".
        Precalcolando la larghezza dal contenitore stabile ``page_host`` prima
        della mappatura, la pagina compare direttamente nella geometria finale.
        """
        try:
            available = int(available)
        except Exception:
            return False
        if available <= 1:
            return False
        width = min(max(760, available - 8), self.MAX_CONTENT_WIDTH)
        x = available / 2
        changed = False
        if self._last_window_width is None or abs(width - self._last_window_width) >= 2:
            self.canvas.itemconfigure(self.window_id, width=width)
            self._last_window_width = width
            changed = True
        if self._last_window_x is None or abs(x - self._last_window_x) >= 1:
            self.canvas.coords(self.window_id, x, 0)
            self._last_window_x = x
        return changed

    def prepare_for_show(self, available_width: int | None = None) -> None:
        """Prepara una pagina nascosta prima che venga resa visibile.

        Il metodo è intenzionalmente sincrono e leggero: non chiama ``update``
        né ``update_idletasks`` e quindi non forza un ridisegno intermedio.
        """
        self._cancel_scheduled("_activate_after_id")
        self._cancel_scheduled("_width_after_id")
        self._cancel_scheduled("_region_after_id")
        try:
            available = int(available_width or self.master.winfo_width())
        except Exception:
            available = 0
        if available <= 1:
            try:
                available = int(self.app.page_host.winfo_width())
            except Exception:
                available = 0
        if available > 1:
            self._pending_canvas_width = available
            self._apply_window_geometry(available)
            self._pending_canvas_width = None
        try:
            self.canvas.itemconfigure(self.window_id, state="normal")
        except Exception:
            pass

    def _apply_width_sync(self):
        self._width_after_id = None
        if not self._active or not self.winfo_exists():
            return
        available = self._pending_canvas_width or self.canvas.winfo_width()
        self._pending_canvas_width = None
        if self._apply_window_geometry(available):
            self._queue_region_sync()

    def _position_resize_hint(self) -> None:
        if self._resize_hint_id is None:
            return
        try:
            width = self._pending_canvas_width or self.canvas.winfo_width()
            height = self.canvas.winfo_height()
            self.canvas.coords(self._resize_hint_id, max(1, width) / 2, max(70, height / 2))
        except Exception:
            pass

    def begin_interactive_resize(self) -> None:
        """Isola il contenuto pesante mentre l'utente trascina il bordo."""
        if not self._active or self._resize_suspended:
            return
        self._resize_suspended = True
        self._cancel_scheduled("_width_after_id")
        self._cancel_scheduled("_region_after_id")
        try:
            self._resize_saved_yview = float(self.canvas.yview()[0])
        except Exception:
            self._resize_saved_yview = 0.0
        try:
            self.canvas.itemconfigure(self.window_id, state="hidden")
        except Exception:
            pass
        try:
            muted = self.app.style.lookup("MutedCard.TLabel", "foreground") or "#60717c"
            bg = self.app.style.lookup("App.TFrame", "background") or "#f5f7fa"
            if self._resize_hint_id is None:
                self._resize_hint_id = self.canvas.create_text(
                    0, 0, text="Ridimensionamento in corso…",
                    fill=muted, font=("Segoe UI", 10), anchor="center",
                )
            self.canvas.itemconfigure(self._resize_hint_id, state="normal", fill=muted)
            self.canvas.configure(background=bg)
            self._position_resize_hint()
        except Exception:
            pass

    def end_interactive_resize(self) -> None:
        """Ripristina e ridimensiona il contenuto una sola volta a drag finito."""
        if not self._active:
            self._resize_suspended = False
            return
        self._resize_suspended = False
        try:
            if self._resize_hint_id is not None:
                self.canvas.itemconfigure(self._resize_hint_id, state="hidden")
            self.canvas.itemconfigure(self.window_id, state="normal")
        except Exception:
            pass
        self._queue_width_sync(immediate=True)
        try:
            self.canvas.yview_moveto(self._resize_saved_yview)
        except Exception:
            pass
        try:
            self._region_after_id = self.after_idle(self._apply_region_sync)
        except Exception:
            self._queue_region_sync(immediate=True)

    # Alias mantenuti per compatibilità con test/integrazioni precedenti.
    def _sync_region(self, event=None):
        self._queue_region_sync(event)

    def _sync_width(self, event=None):
        self._queue_width_sync(event)

    def refresh_colors(self):
        try:
            bg = self.app.style.lookup("App.TFrame", "background") or "#f5f7fa"
            self.canvas.configure(background=bg)
        except Exception:
            pass

    def scroll(self, units: int):
        if self._active and self._scrollbar_visible:
            self.canvas.yview_scroll(units, "units")

    def destroy(self):
        self._cancel_scheduled("_activate_after_id")
        self._cancel_scheduled("_width_after_id")
        self._cancel_scheduled("_region_after_id")
        super().destroy()


class CollapsibleCard(ttk.Frame):
    """Accordion pulito: un solo bordo, intestazione uniforme e contenuto separato."""

    def __init__(self, master, title: str, subtitle: str = "", expanded: bool = False):
        super().__init__(master, style="Accordion.TFrame", padding=(16, 12))
        self.expanded = tk.BooleanVar(value=expanded)
        self.columnconfigure(0, weight=1)

        self.header = ttk.Frame(self, style="AccordionInner.TFrame")
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.columnconfigure(0, weight=1)
        title_box = ttk.Frame(self.header, style="AccordionInner.TFrame")
        title_box.grid(row=0, column=0, sticky="ew")
        self.title_label = ttk.Label(title_box, text=title, style="CardTitle.TLabel")
        self.title_label.pack(anchor="w")
        self.subtitle_label = None
        if subtitle:
            self.subtitle_label = ttk.Label(
                title_box, text=subtitle, style="MutedCard.TLabel", wraplength=820
            )
            self.subtitle_label.pack(anchor="w", pady=(2, 0))
        self.toggle_btn = ttk.Button(
            self.header,
            text="Nascondi" if expanded else "Mostra",
            style="Link.TButton",
            command=self.toggle,
        )
        self.toggle_btn.grid(row=0, column=1, sticky="e", padx=(16, 0))

        self.separator = ttk.Separator(self, orient="horizontal")
        self.content = ttk.Frame(
            self, style="AccordionInner.TFrame", padding=(0, 14, 0, 2)
        )
        if expanded:
            self.separator.grid(row=1, column=0, sticky="ew", pady=(12, 0))
            self.content.grid(row=2, column=0, sticky="ew")

        for widget in (self.header, title_box, self.title_label, self.subtitle_label):
            if widget is not None:
                widget.bind("<Button-1>", lambda _e: self.toggle())

    def toggle(self):
        value = not self.expanded.get()
        self.expanded.set(value)
        self.toggle_btn.configure(text="Nascondi" if value else "Mostra")
        if value:
            self.separator.grid(row=1, column=0, sticky="ew", pady=(12, 0))
            self.content.grid(row=2, column=0, sticky="ew")
        else:
            self.content.grid_remove()
            self.separator.grid_remove()


class SegmentedControl(ttk.Frame):
    """Controllo segmentato accessibile basato su pulsanti ttk."""

    def __init__(self, master, variable: tk.Variable, choices, command=None):
        super().__init__(master, style="CardInner.TFrame")
        self.variable = variable
        self.command = command
        self.buttons = {}

        # Tutti i segmenti devono mantenere esattamente la stessa geometria.
        # Cambiare stile tra stato attivo e inattivo non deve ridimensionare il
        # controllo, né la diversa lunghezza delle etichette deve far apparire
        # il pulsante selezionato più piccolo degli altri.
        choices = list(choices)
        uniform_group = f"segment_{id(self)}"
        uniform_chars = max(8, max((len(str(label)) for label, _value in choices), default=8) + 2)
        for idx, (label, value) in enumerate(choices):
            self.columnconfigure(idx, weight=1, uniform=uniform_group)
            btn = ttk.Button(
                self,
                text=label,
                width=uniform_chars,
                style="Segment.TButton",
                command=lambda v=value: self.set(v),
            )
            btn.grid(
                row=0,
                column=idx,
                sticky="nsew",
                padx=(0 if idx == 0 else 2, 0),
            )
            self.buttons[value] = btn
        try:
            variable.trace_add("write", lambda *_: self.refresh())
        except Exception:
            pass
        self.refresh()

    def set(self, value):
        self.variable.set(value)
        self.refresh()
        if self.command:
            self.command()

    def refresh(self):
        current = self.variable.get()
        for value, btn in self.buttons.items():
            btn.configure(style="SegmentActive.TButton" if value == current else "Segment.TButton")


class App(legacy.App):
    """Interfaccia moderna che riusa integralmente il motore consolidato."""

    PAGE_META = {
        "generate": ("Nuovo progetto", "Crea un PowerPoint con audio o un video MP4"),
        "fix": ("Ripara PowerPoint", "Completa o uniforma gli audio di una presentazione"),
        "batch": ("Elaborazione batch", "Esegui più lavori da un manifest JSON"),
        "voices": ("Libreria voci", "Gestisci voci Microsoft e voci clonate locali"),
        "history": ("Cronologia", "Consulta risultati e rapporti delle elaborazioni"),
        "settings": ("Impostazioni", "Personalizza interfaccia e comportamento"),
        "result": ("Operazione completata", "Il risultato è pronto"),
    }

    def _setup_style(self):
        self._prefs = _load_preferences()
        self._theme_mode = str(self._prefs.get("theme_mode", "system"))
        self.root.minsize(980, 680)
        _center_window(self.root, 1280, 820)

        if _BOOTSTRAP_AVAILABLE:
            self.style = getattr(self.root, "style", None) or tb.Style(
                theme=_theme_for_mode(self._theme_mode)
            )
            try:
                self.style.theme_use(_theme_for_mode(self._theme_mode))
            except Exception:
                pass
        else:
            self.style = ttk.Style(self.root)
            try:
                self.style.theme_use("clam")
            except tk.TclError:
                pass
        self._configure_custom_styles()
        self._apply_window_identity()

    def _configure_custom_styles(self):
        dark = self._theme_mode == "dark" or (
            self._theme_mode == "system" and _system_prefers_dark()
        )
        app_bg = "#101820" if dark else "#F4F7F9"
        card_bg = "#18242E" if dark else "#FFFFFF"
        text = "#EEF4F7" if dark else "#17242C"
        muted = "#B2C1C9" if dark else "#526873"
        border = "#344651" if dark else "#D7E0E5"
        sidebar = SIDEBAR_DARK if dark else SIDEBAR_LIGHT
        hover = "#174C5E" if dark else "#185E75"
        selected = "#2797BC" if dark else BRAND
        sidebar_active = "#18596D" if dark else "#1A6075"
        sidebar_accent = "#43BCE5"
        sidebar_icon = "#BED3DA"
        sidebar_group = "#86AAB6"
        sidebar_separator = "#2A5663" if not dark else "#214551"
        danger = "#C44747"
        warning = "#C78317"
        success = "#258758"
        disabled_bg = "#273640" if dark else "#E8EEF1"
        disabled_fg = "#8597A1" if dark else "#72838C"

        self._colors = {
            "app": app_bg, "card": card_bg, "text": text, "muted": muted,
            "border": border, "sidebar": sidebar, "hover": hover,
            "primary": selected, "sidebar_active": sidebar_active,
            "sidebar_accent": sidebar_accent, "sidebar_icon": sidebar_icon,
            "sidebar_group": sidebar_group, "sidebar_separator": sidebar_separator,
            "danger": danger, "warning": warning, "success": success,
            "disabled_bg": disabled_bg, "disabled_fg": disabled_fg,
        }
        self.root.configure(background=app_bg)
        s = self.style
        s.configure("App.TFrame", background=app_bg)
        s.configure("Header.TFrame", background=app_bg)
        s.configure("Card.TFrame", background=card_bg, relief="solid", borderwidth=1)
        s.configure("CardInner.TFrame", background=card_bg, relief="flat", borderwidth=0)
        s.configure("Accordion.TFrame", background=card_bg, relief="solid", borderwidth=1)
        s.configure("AccordionInner.TFrame", background=card_bg, relief="flat", borderwidth=0)
        s.configure("Sidebar.TFrame", background=sidebar)
        s.configure("Footer.TFrame", background=card_bg, relief="flat", borderwidth=0)
        s.configure("Log.TFrame", background=card_bg, relief="solid", borderwidth=1)
        s.configure("PageTitle.TLabel", background=app_bg, foreground=text,
                    font=("Segoe UI Variable Display", 22, "bold"))
        s.configure("PageSubtitle.TLabel", background=app_bg, foreground=muted,
                    font=("Segoe UI Variable Text", 10))
        s.configure("CardTitle.TLabel", background=card_bg, foreground=text,
                    font=("Segoe UI Variable Text", 11, "bold"))
        s.configure("CardText.TLabel", background=card_bg, foreground=text,
                    font=("Segoe UI Variable Text", 10))
        s.configure("MutedCard.TLabel", background=card_bg, foreground=muted,
                    font=("Segoe UI Variable Text", 9))
        s.configure("MetricValue.TLabel", background=card_bg, foreground=text,
                    font=("Segoe UI Variable Display", 22, "bold"))
        s.configure("MetricLabel.TLabel", background=card_bg, foreground=muted,
                    font=("Segoe UI Variable Text", 9))
        s.configure("SidebarBrand.TLabel", background=sidebar, foreground="#FFFFFF",
                    font=("Segoe UI Variable Display", 14, "bold"))
        s.configure("SidebarSmall.TLabel", background=sidebar, foreground="#D0E1E7",
                    font=("Segoe UI Variable Text", 9))
        s.configure("SidebarGroup.TLabel", background=sidebar, foreground=sidebar_group,
                    font=("Segoe UI Variable Text", 8, "bold"))
        s.configure("SidebarVersion.TLabel", background=sidebar, foreground="#AFC7CF",
                    font=("Segoe UI Variable Text", 8))
        # Gli elementi di navigazione condividono esattamente la stessa geometria.
        # Lo stato selezionato cambia solo sfondo, peso e icona.
        nav_padding = (11, 9)
        s.configure("Sidebar.TButton", background=sidebar, foreground="#E3EFF3",
                    borderwidth=0, padding=nav_padding, anchor="w",
                    font=("Segoe UI Variable Text", 10))
        s.map("Sidebar.TButton",
              background=[("active", hover), ("pressed", hover)],
              foreground=[("active", "#FFFFFF"), ("pressed", "#FFFFFF")])
        s.configure("SidebarActive.TButton", background=sidebar_active, foreground="#FFFFFF",
                    borderwidth=0, padding=nav_padding, anchor="w",
                    font=("Segoe UI Variable Text", 10, "bold"))
        s.map("SidebarActive.TButton",
              background=[("active", sidebar_active), ("pressed", sidebar_active)],
              foreground=[("active", "#FFFFFF"), ("pressed", "#FFFFFF")])
        s.configure("Primary.TButton", background=selected, foreground="#FFFFFF",
                    padding=(16, 9), borderwidth=0,
                    font=("Segoe UI Variable Text", 10, "bold"))
        s.map("Primary.TButton",
              background=[("active", BRAND_DARK), ("disabled", disabled_bg)],
              foreground=[("disabled", disabled_fg)])
        s.configure("Secondary.TButton", padding=(13, 8),
                    font=("Segoe UI Variable Text", 10))
        s.map("Secondary.TButton",
              foreground=[("disabled", disabled_fg)],
              background=[("disabled", disabled_bg)])
        s.configure("Link.TButton", borderwidth=0, padding=(6, 3),
                    font=("Segoe UI Variable Text", 9), foreground=selected,
                    background=card_bg)
        s.configure("Danger.TButton", foreground="#FFFFFF", background=danger,
                    padding=(14, 9), font=("Segoe UI Variable Text", 10, "bold"))
        # Stato attivo e inattivo condividono padding, bordo, rilievo e
        # metrica del font. In questo modo il cambio di selezione modifica solo
        # colore e peso visivo, non larghezza o altezza del pulsante.
        segment_padding = (12, 8)
        segment_font = ("Segoe UI Variable Text", 9)
        s.configure("Segment.TButton", padding=segment_padding, borderwidth=1,
                    relief="flat", font=segment_font, anchor="center")
        s.configure("SegmentActive.TButton", background=selected, foreground="#FFFFFF",
                    padding=segment_padding, borderwidth=1, relief="flat",
                    font=segment_font, anchor="center")
        s.map(
            "SegmentActive.TButton",
            background=[("pressed", selected), ("active", selected)],
            foreground=[("pressed", "#FFFFFF"), ("active", "#FFFFFF")],
        )
        s.configure("Step.TButton", padding=(10, 8), borderwidth=0,
                    font=("Segoe UI Variable Text", 9), anchor="center")
        s.configure("StepActive.TButton", background=selected, foreground="#FFFFFF",
                    padding=(10, 8), borderwidth=0,
                    font=("Segoe UI Variable Text", 9, "bold"))
        s.configure("StepDone.TButton", foreground=selected, padding=(10, 8),
                    borderwidth=0, font=("Segoe UI Variable Text", 9, "bold"))
        s.map("StepActive.TButton", background=[("active", selected)])
        s.configure("Output.TRadiobutton", background=card_bg, foreground=text,
                    padding=(16, 14), font=("Segoe UI Variable Text", 10, "bold"))
        s.configure("SummaryKey.TLabel", background=card_bg, foreground=muted,
                    font=("Segoe UI Variable Text", 9))
        s.configure("SummaryValue.TLabel", background=card_bg, foreground=text,
                    font=("Segoe UI Variable Text", 10, "bold"))
        s.configure("Status.TLabel", background=card_bg, foreground=muted,
                    font=("Segoe UI Variable Text", 9))
        s.configure("FooterStatus.TLabel", background=card_bg, foreground=text,
                    font=("Segoe UI Variable Text", 9))
        s.configure("SuccessTitle.TLabel", background=app_bg, foreground=success,
                    font=("Segoe UI Variable Display", 25, "bold"))
        s.configure("Warning.TLabel", background=card_bg, foreground=warning,
                    font=("Segoe UI Variable Text", 9))
        s.configure("Error.TLabel", background=card_bg, foreground=danger,
                    font=("Segoe UI Variable Text", 9, "bold"))
        s.configure("Success.TLabel", background=card_bg, foreground=success,
                    font=("Segoe UI Variable Text", 9, "bold"))
        s.configure("Treeview", rowheight=31, font=("Segoe UI Variable Text", 9),
                    background=card_bg, fieldbackground=card_bg, foreground=text,
                    borderwidth=0)
        s.map("Treeview", background=[("selected", selected)], foreground=[("selected", "#FFFFFF")])
        s.configure("Treeview.Heading", font=("Segoe UI Variable Text", 9, "bold"),
                    padding=(8, 8), relief="flat")
        s.configure("TEntry", padding=7)
        s.configure("TCombobox", padding=6)
        s.configure("Horizontal.TProgressbar", troughcolor=border, background=selected)
        # Non ridefiniamo TRadiobutton/TCheckbutton globalmente: i temi ttkbootstrap
        # mantengono così indicatori corretti in chiaro, scuro e DPI elevati.

    def _load_app_icon_source(self):
        """Carica una sola volta il PNG principale mantenendo il canale alpha."""
        cached = getattr(self, "_app_icon_source", None)
        if cached is not None:
            return cached
        if not _PIL_AVAILABLE or not APP_ICON_PNG.exists():
            return None
        try:
            with Image.open(APP_ICON_PNG) as source:
                cached = source.convert("RGBA").copy()
            self._app_icon_source = cached
            return cached
        except Exception:
            return None

    def _apply_window_identity(self):
        self.root.title(f"{APP_NAME} — Audio e video per le slide")

        # Su Windows il file ICO multirisoluzione migliora barra del titolo,
        # taskbar, Alt+Tab e collegamenti. iconphoto resta il fallback portabile.
        if platform.system() == "Windows" and APP_ICON_ICO.exists():
            # ``bitmap`` assegna l'icona alla finestra corrente. ``default`` da
            # solo non sostituisce sempre l'icona Tcl/Tk gia' associata alla
            # finestra principale e può lasciare la piuma nella taskbar.
            try:
                self.root.iconbitmap(str(APP_ICON_ICO))
            except Exception:
                pass
            try:
                self.root.iconbitmap(default=str(APP_ICON_ICO))
            except Exception:
                pass

        source = self._load_app_icon_source()
        if source is None:
            return
        try:
            sizes = (16, 20, 24, 32, 40, 48, 64, 128, 256)
            self._window_icons = [
                ImageTk.PhotoImage(
                    source.resize((size, size), Image.Resampling.LANCZOS),
                    master=self.root,
                )
                for size in sizes
            ]
            self.root.iconphoto(True, *self._window_icons)
        except Exception:
            pass

        # Tk può conservare l'icona predefinita della propria classe finestra
        # nella taskbar di Windows. WM_SETICON forza esplicitamente le varianti
        # piccola e grande sull'HWND reale della finestra principale.
        self._apply_native_windows_icon()
        try:
            self.root.after_idle(self._apply_native_windows_icon)
        except Exception:
            pass

    def _apply_native_windows_icon(self) -> None:
        """Applica il file ICO direttamente alla finestra Win32 principale."""
        if platform.system() != "Windows" or not APP_ICON_ICO.exists():
            return
        try:
            import ctypes
            from ctypes import wintypes

            user32 = ctypes.windll.user32
            load_image = user32.LoadImageW
            load_image.argtypes = [
                wintypes.HINSTANCE,
                wintypes.LPCWSTR,
                wintypes.UINT,
                ctypes.c_int,
                ctypes.c_int,
                wintypes.UINT,
            ]
            load_image.restype = wintypes.HANDLE
            user32.GetParent.argtypes = [wintypes.HWND]
            user32.GetParent.restype = wintypes.HWND
            user32.SendMessageW.argtypes = [
                wintypes.HWND,
                wintypes.UINT,
                ctypes.c_size_t,
                ctypes.c_ssize_t,
            ]
            user32.SendMessageW.restype = ctypes.c_ssize_t

            self.root.update_idletasks()
            tk_hwnd = wintypes.HWND(int(self.root.winfo_id()))
            hwnd = user32.GetParent(tk_hwnd) or tk_hwnd

            IMAGE_ICON = 1
            LR_LOADFROMFILE = 0x0010
            WM_SETICON = 0x0080
            ICON_SMALL = 0
            ICON_BIG = 1
            SM_CXICON, SM_CYICON = 11, 12
            SM_CXSMICON, SM_CYSMICON = 49, 50

            big_w = max(32, int(user32.GetSystemMetrics(SM_CXICON)))
            big_h = max(32, int(user32.GetSystemMetrics(SM_CYICON)))
            small_w = max(16, int(user32.GetSystemMetrics(SM_CXSMICON)))
            small_h = max(16, int(user32.GetSystemMetrics(SM_CYSMICON)))
            icon_path = str(APP_ICON_ICO.resolve())

            hicon_big = load_image(None, icon_path, IMAGE_ICON, big_w, big_h, LR_LOADFROMFILE)
            hicon_small = load_image(None, icon_path, IMAGE_ICON, small_w, small_h, LR_LOADFROMFILE)
            def _handle_value(value) -> int:
                raw = getattr(value, "value", value)
                return int(raw or 0)

            if hicon_big:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, _handle_value(hicon_big))
            if hicon_small:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, _handle_value(hicon_small))

            # Mantiene vivi gli handle per tutta la durata della finestra.
            self._native_window_icon_handles = (hicon_big, hicon_small)
            self._native_window_hwnd = _handle_value(hwnd)
        except Exception:
            # iconbitmap/iconphoto restano i fallback portabili.
            pass

    def _draw_brand_badge(self) -> None:
        """Disegna l'icona ufficiale al posto del vecchio cerchio con la P."""
        badge = getattr(self, "_brand_canvas", None)
        if badge is None:
            return
        badge.delete("all")

        source_path = APP_ICON_SIDEBAR_PNG if APP_ICON_SIDEBAR_PNG.exists() else APP_ICON_PNG
        if _PIL_AVAILABLE and source_path.exists():
            try:
                with Image.open(source_path) as source:
                    icon = source.convert("RGBA").resize(
                        (44, 44), Image.Resampling.LANCZOS
                    )
                self._brand_icon = ImageTk.PhotoImage(icon, master=self.root)
                badge.create_image(22, 22, image=self._brand_icon, anchor="center")
                return
            except Exception:
                pass

        # Fallback leggibile se Pillow o la risorsa non fossero disponibili.
        badge.create_oval(2, 2, 42, 42, fill=self._colors["primary"], outline="")
        badge.create_text(
            22,
            22,
            text="S",
            fill="white",
            font=("Segoe UI Variable Display", 17, "bold"),
        )

    def _render_nav_icon_image(
        self,
        kind: str,
        color: str,
        *,
        size: int = 20,
        slot_width: int = 30,
    ):
        """Disegna un'icona di navigazione coerente in stile Fluent regular.

        Tutte le icone condividono canvas, spessore, margini e area ottica.
        Il canvas più largo crea uno spazio costante fra icona e testo senza
        dipendere dalle metriche del font o dal tema ttk.
        """
        if not _PIL_AVAILABLE:
            return None
        scale = 4
        W, H = slot_width * scale, size * scale
        img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        stroke = max(5, int(1.65 * scale))
        ox = 1 * scale

        def pt(x: float, y: float) -> tuple[int, int]:
            return (int((x + 0.5) * scale + ox), int((y + 0.5) * scale))

        def line(points, width=stroke, joint="curve"):
            coords = [pt(x, y) for x, y in points]
            d.line(coords, fill=color, width=width, joint=joint)
            radius = max(1, width // 2)
            for x, y in (coords[0], coords[-1]):
                d.ellipse((x-radius, y-radius, x+radius, y+radius), fill=color)

        def ellipse(box, width=stroke):
            x1, y1, x2, y2 = box
            d.ellipse((*pt(x1, y1), *pt(x2, y2)), outline=color, width=width)

        def arc(box, start, end, width=stroke):
            x1, y1, x2, y2 = box
            d.arc((*pt(x1, y1), *pt(x2, y2)), start, end, fill=color, width=width)

        if kind == "plus":
            # Documento con angolo piegato e simbolo +.
            line([(4, 2.5), (12, 2.5), (16, 6.5), (16, 17.5), (4, 17.5), (4, 2.5)])
            line([(12, 2.5), (12, 6.5), (16, 6.5)], width=max(4, stroke-1))
            line([(6.5, 11.5), (12.5, 11.5)])
            line([(9.5, 8.5), (9.5, 14.5)])
        elif kind == "repair":
            # Freccia circolare: ripara/aggiorna una presentazione esistente.
            arc((2.5, 2.5, 17.5, 17.5), 35, 305)
            p1 = pt(16.8, 4.7); p2 = pt(13.2, 4.5); p3 = pt(16.0, 7.4)
            d.polygon([p1, p2, p3], fill=color)
            line([(6.2, 10.2), (8.8, 12.8), (14.2, 7.2)], width=max(4, stroke-1))
        elif kind == "batch":
            # Tre lavori in coda, tutti con la stessa scansione verticale.
            for y in (4.0, 9.5, 15.0):
                x1, y1 = pt(3.0, y-1.4); x2, y2 = pt(5.8, y+1.4)
                d.rounded_rectangle((x1, y1, x2, y2), radius=2*scale,
                                    outline=color, width=max(4, stroke-1))
                line([(8.0, y), (17.0, y)], width=max(4, stroke-1))
        elif kind == "voice":
            # Microfono regular con base centrata.
            x1, y1 = pt(7.0, 2.5); x2, y2 = pt(13.0, 11.5)
            d.rounded_rectangle((x1, y1, x2, y2), radius=3*scale,
                                outline=color, width=stroke)
            arc((4.5, 7.5, 15.5, 16.0), 0, 180)
            line([(10.0, 15.8), (10.0, 18.0)], width=max(4, stroke-1))
            line([(6.8, 18.0), (13.2, 18.0)], width=max(4, stroke-1))
        elif kind == "history":
            # Orologio con ritorno: cronologia delle elaborazioni.
            arc((2.8, 2.8, 17.2, 17.2), 48, 340)
            p1 = pt(3.2, 5.4); p2 = pt(3.4, 2.0); p3 = pt(6.2, 4.0)
            d.polygon([p1, p2, p3], fill=color)
            line([(10.0, 6.0), (10.0, 10.2), (13.3, 12.0)], width=max(4, stroke-1))
        elif kind == "settings":
            # Ingranaggio semplificato, leggibile anche a 100% DPI.
            import math
            ellipse((5.0, 5.0, 15.0, 15.0), width=max(4, stroke-1))
            ellipse((8.2, 8.2, 11.8, 11.8), width=max(4, stroke-1))
            for angle in range(0, 360, 45):
                a = math.radians(angle)
                line([
                    (10 + 5.8*math.cos(a), 10 + 5.8*math.sin(a)),
                    (10 + 8.0*math.cos(a), 10 + 8.0*math.sin(a)),
                ], width=max(4, stroke-1))
        else:
            ellipse((4.0, 4.0, 16.0, 16.0))

        return img.resize((slot_width, size), Image.Resampling.LANCZOS)

    def _make_nav_icon(
        self,
        kind: str,
        *,
        color: str | None = None,
        size: int = 20,
        slot_width: int = 30,
    ):
        if not _PIL_AVAILABLE:
            return None
        try:
            image = self._render_nav_icon_image(
                kind,
                color or self._colors.get("sidebar_icon", "#BED3DA"),
                size=size,
                slot_width=slot_width,
            )
            return ImageTk.PhotoImage(image) if image is not None else None
        except Exception:
            return None

    def _create_nav_item(self, parent, row: int, key: str, text: str, icon_kind: str):
        """Crea una voce sidebar con icone stateful e indicatore laterale."""
        wrapper = tk.Frame(parent, bg=self._colors["sidebar"], height=44, bd=0)
        wrapper.grid(row=row, column=0, sticky="ew", pady=2)
        wrapper.grid_propagate(False)
        wrapper.columnconfigure(1, weight=1)
        indicator = tk.Frame(wrapper, bg=self._colors["sidebar"], width=3, bd=0)
        indicator.grid(row=0, column=0, sticky="ns")

        icon_set = {
            "normal": self._make_nav_icon(icon_kind, color=self._colors["sidebar_icon"]),
            "hover": self._make_nav_icon(icon_kind, color="#FFFFFF"),
            "active": self._make_nav_icon(icon_kind, color="#FFFFFF"),
        }
        self._nav_icons.extend([icon for icon in icon_set.values() if icon is not None])
        btn = ttk.Button(
            wrapper,
            text=text,
            image=icon_set["normal"] or "",
            compound="left",
            style="Sidebar.TButton",
            command=lambda k=key: self._show_page(k),
            takefocus=True,
        )
        btn.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        def on_enter(_event=None, item_key=key):
            if self.current_page != item_key:
                indicator.configure(bg="#2C7E98")
                if icon_set["hover"] is not None:
                    btn.configure(image=icon_set["hover"])

        def on_leave(_event=None, item_key=key):
            if self.current_page != item_key:
                indicator.configure(bg=self._colors["sidebar"])
                if icon_set["normal"] is not None:
                    btn.configure(image=icon_set["normal"])

        for widget in (wrapper, indicator, btn):
            widget.bind("<Enter>", on_enter, add="+")
            widget.bind("<Leave>", on_leave, add="+")

        self.nav_buttons[key] = btn
        self._nav_rows[key] = wrapper
        self._nav_indicators[key] = indicator
        self._nav_icon_sets[key] = icon_set
        self._nav_icon_kinds[key] = icon_kind
        return wrapper

    def _update_sidebar_selection(self, key: str) -> None:
        for name, btn in self.nav_buttons.items():
            active = name == key
            btn.configure(style="SidebarActive.TButton" if active else "Sidebar.TButton")
            indicator = self._nav_indicators.get(name)
            if indicator is not None:
                indicator.configure(
                    bg=self._colors["sidebar_accent"] if active else self._colors["sidebar"]
                )
            icons = self._nav_icon_sets.get(name, {})
            icon = icons.get("active" if active else "normal")
            if icon is not None:
                btn.configure(image=icon)


    def _refresh_sidebar_visuals(self) -> None:
        """Aggiorna superfici, indicatori e icone dopo un cambio tema."""
        sidebar = self._colors["sidebar"]
        for row in getattr(self, "_nav_rows", {}).values():
            row.configure(bg=sidebar)
        for separator_name in ("_sidebar_top_separator", "_sidebar_bottom_separator"):
            separator = getattr(self, separator_name, None)
            if separator is not None:
                separator.configure(bg=self._colors["sidebar_separator"])

        refreshed = []
        for key, kind in getattr(self, "_nav_icon_kinds", {}).items():
            icon_set = {
                "normal": self._make_nav_icon(kind, color=self._colors["sidebar_icon"]),
                "hover": self._make_nav_icon(kind, color="#FFFFFF"),
                "active": self._make_nav_icon(kind, color="#FFFFFF"),
            }
            self._nav_icon_sets[key] = icon_set
            refreshed.extend(icon for icon in icon_set.values() if icon is not None)
        if refreshed:
            self._nav_icons = refreshed
        self._update_sidebar_selection(self.current_page)

    def _build_ui(self):
        self.current_page = "generate"
        self.wizard_step = tk.IntVar(value=1)
        self._max_unlocked_step = 1
        self.fix_wizard_step = tk.IntVar(value=1)
        self._fix_max_unlocked_step = 1
        self.theme_mode_var = tk.StringVar(value=self._theme_mode)
        self.show_log_on_error_var = tk.BooleanVar(
            value=bool(self._prefs.get("show_log_on_error", True))
        )
        self.history_folder_var = tk.StringVar(
            value=str(self._prefs.get("history_folder", Path.cwd()))
        )
        self.batch_manifest_var = tk.StringVar()
        self.batch_force_var = tk.BooleanVar(value=False)
        self.batch_stop_on_error_var = tk.BooleanVar(value=False)
        self.batch_result_var = tk.StringVar(value="Seleziona o crea una coda di lavori.")
        self.batch_jobs_var = tk.StringVar(value="Nessun lavoro caricato")
        self.voice_search_var = tk.StringVar()
        self.voice_engine_filter_var = tk.StringVar(value="Tutti")
        self.fix_filter_var = tk.StringVar(value="all")
        self.fix_total_var = tk.StringVar(value="0")
        self.fix_valid_var = tk.StringVar(value="0")
        self.fix_missing_var = tk.StringVar(value="0")
        self.fix_invalid_var = tk.StringVar(value="0")
        self.result_path_var = tk.StringVar()
        self.result_title_var = tk.StringVar(value="Operazione completata")
        self.result_details_var = tk.StringVar()
        self.summary_vars = {k: tk.StringVar(value="—") for k in (
            "pptx", "scripts", "voice", "output", "options", "preflight", "validation"
        )}
        self.fix_summary_vars = {k: tk.StringVar(value="—") for k in (
            "pptx", "analysis", "scripts", "voice", "output", "actions", "validation"
        )}
        self._log_visible = False
        self._all_fix_rows: dict[str, tuple] = {}
        self._voice_display_to_id: dict[str, str] = {}
        self._voice_id_to_display: dict[str, str] = {}
        self._history_items: dict[str, dict] = {}
        self._last_fix_analysis = None
        self._batch_jobs = []
        # Le analisi di presentazioni grandi vengono eseguite fuori dal thread
        # Tk. Il token impedisce che un risultato ormai superato aggiorni la GUI.
        self._fix_analysis_token = 0
        self._active_fix_analysis_path = None
        self._current_operation = None
        self._history_scan_token = 0
        self._diagnostics_token = 0
        self._summary_preflight_token = 0
        self._summary_preflight_key = None
        self._summary_preflight_result = None
        self._template_running = False
        self._cache_clear_running = False
        self._sheet_scan_tokens = {"generate": 0, "fix": 0}
        self._sheet_scan_after_ids = {"generate": None, "fix": None}
        self._sheet_scan_timeout_ids = {"generate": None, "fix": None}
        self._sheet_combos = {}
        self._sheet_status_labels = {}
        self._ai_prompt_buttons = []
        self._ai_prompt_reset_after_id = None
        # Coda thread-safe: i worker non chiamano mai direttamente Tcl/Tk.
        # Il main thread la svuota periodicamente tramite ``after``.
        self._ui_queue: queue.Queue = queue.Queue()
        self._ui_poller_id = None
        # Il window manager invia una raffica di <Configure> durante il drag.
        # Accorpiamo tali eventi e ricalcoliamo i breakpoint solo a resize fermo.
        self._root_resize_after_id = None
        self._pending_root_width = None
        self._pending_root_height = None
        self._last_root_size = None
        self._interactive_resize = False
        self._resize_settle_ms = 230

        # Stati indipendenti: passare da Genera a Fix non trascina i file.
        self.generate_input_pptx_var = self.input_pptx_var
        self.generate_input_xlsx_var = self.input_xlsx_var
        self.generate_output_pptx_var = self.output_pptx_var
        self.generate_sheet_name_var = self.sheet_name_var
        self.generate_script_column_var = self.script_column_var
        self.generate_has_header_var = self.has_header_var
        self.generate_script_source_var = self.script_source_var
        self.fix_input_pptx_var = tk.StringVar()
        self.fix_input_xlsx_var = tk.StringVar()
        self.fix_output_pptx_var = tk.StringVar()
        self.fix_sheet_name_var = tk.StringVar()
        self.fix_script_column_var = tk.StringVar(value="A")
        self.fix_has_header_var = tk.BooleanVar(value=False)
        self.fix_script_source_var = tk.StringVar(value="excel")
        # La voce del Fix è indipendente dal flusso "Nuovo progetto".
        self.fix_voice_var = tk.StringVar(value=self.voice_var.get())
        self.fix_voice_combo_var = tk.StringVar()

        shell = ttk.Frame(self.root, style="App.TFrame")
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        self.sidebar = ttk.Frame(shell, style="Sidebar.TFrame", width=236)
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self.sidebar.grid_propagate(False)
        self.sidebar.columnconfigure(0, weight=1)

        brand = ttk.Frame(self.sidebar, style="Sidebar.TFrame", padding=(18, 18, 14, 16))
        brand.grid(row=0, column=0, sticky="ew")
        brand.columnconfigure(1, weight=1)
        badge = tk.Canvas(
            brand,
            width=44,
            height=44,
            highlightthickness=0,
            bd=0,
            bg=self._colors["sidebar"],
        )
        badge.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 11))
        self._brand_canvas = badge
        self._draw_brand_badge()
        ttk.Label(brand, text=APP_NAME, style="SidebarBrand.TLabel").grid(
            row=0, column=1, sticky="sw"
        )
        ttk.Label(brand, text="Audio e video per slide", style="SidebarSmall.TLabel").grid(
            row=1, column=1, sticky="nw", pady=(2, 0)
        )

        self._sidebar_top_separator = tk.Frame(
            self.sidebar, bg=self._colors["sidebar_separator"], height=1, bd=0
        )
        self._sidebar_top_separator.grid(row=1, column=0, sticky="ew", padx=16)

        self.nav_buttons = {}
        self._nav_icons = []
        self._nav_icon_sets = {}
        self._nav_icon_kinds = {}
        self._nav_rows = {}
        self._nav_indicators = {}
        nav = ttk.Frame(self.sidebar, style="Sidebar.TFrame", padding=(10, 13, 10, 0))
        nav.grid(row=2, column=0, sticky="new")
        nav.columnconfigure(0, weight=1)

        ttk.Label(nav, text="CREA E RIPARA", style="SidebarGroup.TLabel").grid(
            row=0, column=0, sticky="w", padx=(10, 0), pady=(0, 5)
        )
        self._create_nav_item(nav, 1, "generate", "Nuovo progetto", "plus")
        self._create_nav_item(nav, 2, "fix", "Ripara PPTX", "repair")

        ttk.Label(nav, text="STRUMENTI", style="SidebarGroup.TLabel").grid(
            row=3, column=0, sticky="w", padx=(10, 0), pady=(16, 5)
        )
        self._create_nav_item(nav, 4, "batch", "Elaborazione batch", "batch")
        self._create_nav_item(nav, 5, "voices", "Libreria voci", "voice")
        self._create_nav_item(nav, 6, "history", "Cronologia", "history")

        # Impostazioni nel footer, come nelle NavigationView Windows.
        self.sidebar.rowconfigure(3, weight=1)
        sidebar_bottom = ttk.Frame(
            self.sidebar, style="Sidebar.TFrame", padding=(10, 0, 10, 16)
        )
        sidebar_bottom.grid(row=4, column=0, sticky="sew")
        sidebar_bottom.columnconfigure(0, weight=1)
        self._sidebar_bottom_separator = tk.Frame(
            sidebar_bottom, bg=self._colors["sidebar_separator"], height=1, bd=0
        )
        self._sidebar_bottom_separator.grid(row=0, column=0, sticky="ew", padx=6, pady=(0, 8))
        self._create_nav_item(sidebar_bottom, 1, "settings", "Impostazioni", "settings")
        version_box = ttk.Frame(sidebar_bottom, style="Sidebar.TFrame", padding=(10, 12, 8, 0))
        version_box.grid(row=2, column=0, sticky="ew")
        ttk.Label(version_box, text="Versione 2.9", style="SidebarVersion.TLabel").pack(anchor="w")

        main = ttk.Frame(shell, style="App.TFrame")
        main.grid(row=0, column=1, sticky="nsew")
        main.rowconfigure(1, weight=1)
        main.columnconfigure(0, weight=1)

        header = ttk.Frame(main, style="Header.TFrame", padding=(24, 14, 24, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        title_box = ttk.Frame(header, style="Header.TFrame")
        title_box.grid(row=0, column=0, sticky="w")
        self.page_title = ttk.Label(title_box, text="", style="PageTitle.TLabel")
        self.page_title.pack(anchor="w")
        self.page_subtitle = ttk.Label(title_box, text="", style="PageSubtitle.TLabel")
        self.page_subtitle.pack(anchor="w", pady=(2, 0))
        header_actions = ttk.Frame(header, style="Header.TFrame")
        header_actions.grid(row=0, column=1, sticky="e")
        self.log_toggle_btn = ttk.Button(
            header_actions, text="Dettagli tecnici", style="Secondary.TButton",
            command=self._toggle_log,
        )
        self.log_toggle_btn.grid(row=0, column=0, padx=(0, 8))

        content_shell = ttk.Frame(main, style="App.TFrame")
        content_shell.grid(row=1, column=0, sticky="nsew")
        content_shell.rowconfigure(0, weight=1)
        content_shell.columnconfigure(0, weight=1)
        self.page_host = ttk.Frame(content_shell, style="App.TFrame")
        self.page_host.grid(row=0, column=0, sticky="nsew")
        self.page_host.rowconfigure(0, weight=1)
        self.page_host.columnconfigure(0, weight=1)

        self.pages = {}
        self._build_generate_page()
        self._build_fix_page()
        self._build_batch_page()
        self._build_voices_page()
        self._build_history_page()
        self._build_settings_page()
        self._build_result_page()

        self.log_drawer = ttk.Frame(content_shell, style="Log.TFrame", padding=(16, 10))
        self.log_drawer.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 8))
        self.log_drawer.columnconfigure(0, weight=1)
        log_header = ttk.Frame(self.log_drawer, style="CardInner.TFrame")
        log_header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        log_header.columnconfigure(0, weight=1)
        ttk.Label(log_header, text="Dettagli tecnici", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(log_header, text="Pulisci", style="Link.TButton",
                   command=self._clear_log).grid(row=0, column=1, sticky="e")
        self.log_text = scrolledtext.ScrolledText(
            self.log_drawer, height=9, state="disabled", font=("Cascadia Mono", 9),
            wrap="word", relief="flat", borderwidth=0,
        )
        self.log_text.grid(row=1, column=0, sticky="ew")
        self.log_drawer.grid_remove()

        self._build_action_footer(main)
        self.root.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.root.bind_all("<Button-4>", self._on_mousewheel, add="+")
        self.root.bind_all("<Button-5>", self._on_mousewheel, add="+")
        self.root.bind("<Configure>", self._on_root_resize, add="+")
        self.root.bind("<Destroy>", self._on_root_destroy, add="+")

        # Aggiorna stati e validazione quando cambiano i dati essenziali.
        for var in (
            self.generate_input_pptx_var, self.generate_input_xlsx_var,
            self.generate_output_pptx_var, self.fix_input_pptx_var,
            self.fix_input_xlsx_var, self.fix_output_pptx_var,
            self.generate_script_source_var, self.fix_script_source_var,
            self.voice_var, self.fix_voice_var,
            self.output_mode_var, self.batch_manifest_var,
        ):
            try:
                var.trace_add("write", lambda *_: self._on_form_state_change())
            except Exception:
                pass

        try:
            self.generate_input_xlsx_var.trace_add(
                "write", lambda *_: self._schedule_sheet_names_load("generate")
            )
            self.fix_input_xlsx_var.trace_add(
                "write", lambda *_: self._schedule_sheet_names_load("fix")
            )
        except Exception:
            pass

        self._rebuild_voice_list()
        self._on_script_source_change()
        self._show_page(str(self._prefs.get("start_page", "generate")))
        self._show_wizard_step(1, force=True)
        self._show_fix_wizard_step(1, force=True)
        self._set_progress_idle("Pronto")
        self._schedule_ui_poller()

    def _schedule_ui_poller(self):
        if self._ui_poller_id is None:
            try:
                self._ui_poller_id = self.root.after(25, self._drain_ui_queue)
            except Exception:
                self._ui_poller_id = None

    def _post_ui(self, callback, *args, **kwargs):
        """Accoda un aggiornamento Tk da qualunque thread."""
        try:
            self._ui_queue.put_nowait((callback, args, kwargs))
        except Exception:
            pass

    def _drain_ui_queue(self):
        self._ui_poller_id = None
        try:
            while True:
                callback, args, kwargs = self._ui_queue.get_nowait()
                try:
                    callback(*args, **kwargs)
                except tk.TclError:
                    return
                except Exception as exc:
                    # Non lasciare che un singolo callback interrompa il poller.
                    try:
                        self._log(f"\nErrore aggiornamento interfaccia: {exc}\n")
                    except Exception:
                        pass
        except queue.Empty:
            pass
        try:
            if self.root.winfo_exists():
                # Durante il resize riduciamo anche il polling non essenziale;
                # quando un worker è attivo restiamo reattivi, a riposo evitiamo
                # 40 wake-up al secondo inutili.
                if getattr(self, "_interactive_resize", False):
                    delay = 120
                elif getattr(self, "_is_busy", False):
                    delay = 25
                else:
                    delay = 90
                self._ui_poller_id = self.root.after(delay, self._drain_ui_queue)
        except Exception:
            self._ui_poller_id = None

    def _build_action_footer(self, main):
        self.footer = ttk.Frame(main, style="Footer.TFrame", padding=(22, 10))
        self.footer.grid(row=2, column=0, sticky="ew")
        self.footer.columnconfigure(1, weight=1)
        self.back_btn = ttk.Button(
            self.footer, text="Indietro", style="Secondary.TButton",
            command=self._wizard_back,
        )
        self.back_btn.grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.status_box = ttk.Frame(self.footer, style="CardInner.TFrame")
        self.status_box.grid(row=0, column=1, sticky="ew")
        self.status_box.columnconfigure(0, weight=1)
        self.footer_step_label = ttk.Label(
            self.status_box, text="Passaggio 1 di 4", style="FooterStatus.TLabel"
        )
        self.footer_step_label.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(
            self.status_box, mode="determinate", maximum=100, value=0
        )
        self.progress_label = ttk.Label(
            self.status_box, text="Pronto", style="Status.TLabel"
        )
        self.cancel_btn = ttk.Button(
            self.footer, text="Annulla", style="Secondary.TButton",
            command=self._on_cancel, state="disabled",
        )
        self.generate_btn = ttk.Button(
            self.footer, text="Avanti", style="Primary.TButton",
            command=self._on_primary_action,
        )
        self.generate_btn.grid(row=0, column=3, sticky="e")
        self._show_footer_idle()

    def _show_footer_idle(self):
        try:
            self.progress.grid_remove()
            self.progress_label.grid_remove()
            self.cancel_btn.grid_remove()
            self.footer_step_label.grid(row=0, column=0, sticky="w")
        except Exception:
            pass

    def _show_footer_progress(self):
        try:
            self.footer_step_label.grid_remove()
            self.progress.grid(row=0, column=0, sticky="ew")
            self.progress_label.grid(row=1, column=0, sticky="w", pady=(3, 0))
            self.cancel_btn.grid(row=0, column=2, padx=(14, 8))
        except Exception:
            pass

    def _on_form_state_change(self):
        if getattr(self, "current_page", None) == "fix":
            current = self.fix_input_pptx_var.get().strip()
            active = getattr(self, "_active_fix_analysis_path", None)
            if (getattr(self, "_current_operation", None) == "fix_analysis"
                    and active and current != active):
                self._cancel_event.set()
                self._fix_analysis_token += 1
                self._is_busy = False
                self._current_operation = None
                self._active_fix_analysis_path = None
                if hasattr(self, "fix_analysis_var"):
                    self.fix_analysis_var.set("Analisi annullata: è stato selezionato un altro file.")
            if getattr(self, "_last_analyzed_fix_path", None) != current:
                self._last_fix_analysis = None
                if hasattr(self, "fix_analysis_var"):
                    self.fix_analysis_var.set("Seleziona un PowerPoint e avvia l’analisi.")
                if hasattr(self, "fix_empty_label"):
                    self.fix_empty_label.grid()
        self._update_action_bar()
        if getattr(self, "wizard_step", None) is not None and self.wizard_step.get() == 4:
            self._refresh_summary()
        if (getattr(self, "fix_wizard_step", None) is not None
                and self.fix_wizard_step.get() == 4):
            self._refresh_fix_summary()

    def _on_root_resize(self, event=None):
        """Congela il layout pesante durante il trascinamento del bordo.

        Il contenuto della pagina viene nascosto temporaneamente e ridisegnato
        una sola volta dopo l'ultimo evento di dimensionamento. Gli eventi di
        semplice spostamento della finestra vengono ignorati.
        """
        if event is not None and event.widget is not self.root:
            return
        try:
            width = int(event.width) if event is not None else int(self.root.winfo_width())
            height = int(event.height) if event is not None else int(self.root.winfo_height())
        except Exception:
            return
        size = (width, height)
        if self._last_root_size == size:
            return
        self._last_root_size = size
        self._pending_root_width = width
        self._pending_root_height = height

        if not self._interactive_resize:
            self._interactive_resize = True
            page = self.pages.get(getattr(self, "current_page", "")) if hasattr(self, "pages") else None
            if isinstance(page, ScrollablePage):
                page.begin_interactive_resize()

        if self._root_resize_after_id is not None:
            try:
                self.root.after_cancel(self._root_resize_after_id)
            except Exception:
                pass
        try:
            self._root_resize_after_id = self.root.after(
                self._resize_settle_ms, self._finish_root_resize
            )
        except Exception:
            self._root_resize_after_id = None

    def _finish_root_resize(self):
        self._root_resize_after_id = None
        width = self._pending_root_width or self.root.winfo_width()
        self._pending_root_width = None
        self._pending_root_height = None
        mode = "compact" if width < 1180 else "wide"
        if hasattr(self, "file_cards_frame"):
            if getattr(self, "_last_layout_mode", None) != mode:
                self._last_layout_mode = mode
                self._layout_file_cards(mode)
        if hasattr(self, "fix_files_frame"):
            if getattr(self, "_last_fix_layout_mode", None) != mode:
                self._last_fix_layout_mode = mode
                self._layout_fix_file_cards(mode)
        page = self.pages.get(getattr(self, "current_page", "")) if hasattr(self, "pages") else None
        if isinstance(page, ScrollablePage):
            page.end_interactive_resize()
        self._interactive_resize = False

    # Alias per test e integrazioni delle versioni precedenti.
    def _apply_root_resize(self):
        self._finish_root_resize()

    def _on_root_destroy(self, event=None):
        if event is not None and event.widget is not self.root:
            return
        if self._root_resize_after_id is not None:
            try:
                self.root.after_cancel(self._root_resize_after_id)
            except Exception:
                pass
            self._root_resize_after_id = None
        self._interactive_resize = False
        for mode, after_id in getattr(self, "_sheet_scan_after_ids", {}).items():
            if after_id is not None:
                try:
                    self.root.after_cancel(after_id)
                except Exception:
                    pass
                self._sheet_scan_after_ids[mode] = None
            timeout_id = getattr(self, "_sheet_scan_timeout_ids", {}).get(mode)
            if timeout_id is not None:
                try:
                    self.root.after_cancel(timeout_id)
                except Exception:
                    pass
                self._sheet_scan_timeout_ids[mode] = None
            if mode in getattr(self, "_sheet_scan_tokens", {}):
                self._sheet_scan_tokens[mode] += 1
        if getattr(self, "_ai_prompt_reset_after_id", None) is not None:
            try:
                self.root.after_cancel(self._ai_prompt_reset_after_id)
            except Exception:
                pass
            self._ai_prompt_reset_after_id = None
        for page in getattr(self, "pages", {}).values():
            if isinstance(page, ScrollablePage):
                page.set_active(False)

    def _new_page(self, key: str, scrollable: bool = True):
        page = ScrollablePage(self.page_host, self) if scrollable else ttk.Frame(
            self.page_host, style="App.TFrame", padding=(24, 18, 24, 24)
        )
        page.grid(row=0, column=0, sticky="nsew")
        self.pages[key] = page
        return page.body if scrollable else page

    def _card(self, parent, padding=(18, 16)):
        card = ttk.Frame(parent, style="Card.TFrame", padding=padding)
        card.columnconfigure(0, weight=1)
        return card

    def _normalise_suffixes(self, allowed_suffixes):
        values = allowed_suffixes() if callable(allowed_suffixes) else allowed_suffixes
        return {str(item).lower() for item in (values or set())}

    def _save_path_status(self, value: str, allowed_suffixes) -> tuple[bool, str, str]:
        """Valida una destinazione nuova: deve esistere la cartella, non il file."""
        raw = str(value or "").strip()
        if not raw:
            return False, "Nessun percorso selezionato", "MutedCard.TLabel"
        path = Path(raw).expanduser()
        suffixes = self._normalise_suffixes(allowed_suffixes)
        if suffixes and path.suffix.lower() not in suffixes:
            expected = " oppure ".join(sorted(suffixes))
            return False, f"Estensione non valida · usa {expected}", "Error.TLabel"
        parent = path.parent if str(path.parent) else Path.cwd()
        if not parent.exists() or not parent.is_dir():
            return False, "Cartella di destinazione non trovata", "Error.TLabel"
        if path.exists() and path.is_dir():
            return False, "La destinazione indicata è una cartella", "Error.TLabel"
        if path.exists():
            return True, f"File esistente · {path.name} · verrà chiesta conferma", "Warning.TLabel"
        return True, f"Destinazione pronta · {path.name}", "Success.TLabel"

    def _file_card(
        self, parent, title: str, description: str, var, command,
        button_text="Seleziona", *, path_kind: str = "input",
        allowed_suffixes=None, dependent_vars=(),
    ):
        card = self._card(parent, padding=(16, 14))
        card.columnconfigure(0, weight=1)
        top = ttk.Frame(card, style="CardInner.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        title_label = ttk.Label(top, text=title, style="CardTitle.TLabel")
        title_label.grid(row=0, column=0, sticky="w")
        button = ttk.Button(top, text=button_text, style="Secondary.TButton", command=command)
        button.grid(row=0, column=1, sticky="e")
        ttk.Label(card, text=description, style="MutedCard.TLabel", wraplength=620).grid(
            row=1, column=0, sticky="w", pady=(3, 10)
        )
        entry = ttk.Entry(card, textvariable=var)
        entry.grid(row=2, column=0, sticky="ew")
        status = ttk.Label(card, text="Nessun file selezionato", style="MutedCard.TLabel")
        status.grid(row=3, column=0, sticky="w", pady=(7, 0))

        def refresh(*_):
            value = str(var.get()).strip()
            if path_kind == "output":
                _ok, text, style = self._save_path_status(value, allowed_suffixes)
                status.configure(text=text, style=style)
                return
            if not value:
                status.configure(text="Nessun file selezionato", style="MutedCard.TLabel")
                return
            path = Path(value)
            suffixes = self._normalise_suffixes(allowed_suffixes)
            if not path.is_file():
                status.configure(text="Percorso non valido", style="Error.TLabel")
            elif suffixes and path.suffix.lower() not in suffixes:
                expected = " oppure ".join(sorted(suffixes))
                status.configure(text=f"Formato non valido · usa {expected}", style="Error.TLabel")
            else:
                size = path.stat().st_size / (1024 * 1024)
                status.configure(
                    text=f"File pronto · {path.name} · {size:.1f} MB",
                    style="Success.TLabel",
                )
        try:
            var.trace_add("write", refresh)
            for dep in dependent_vars:
                dep.trace_add("write", refresh)
        except Exception:
            pass
        status.refresh_path_status = refresh  # type: ignore[attr-defined]
        refresh()
        return card, (title_label, entry, button), status

    def _section_heading(self, parent, title: str, subtitle: str = ""):
        ttk.Label(parent, text=title, style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        if subtitle:
            ttk.Label(parent, text=subtitle, style="MutedCard.TLabel", wraplength=760).grid(
                row=1, column=0, sticky="w", pady=(3, 12)
            )

    def _build_generate_page(self):
        body = self._new_page("generate")
        stepper_card = self._card(body, padding=(12, 10))
        stepper_card.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        stepper = ttk.Frame(stepper_card, style="CardInner.TFrame")
        stepper.grid(row=0, column=0, sticky="ew")
        for i in range(4):
            stepper.columnconfigure(i, weight=1)
        self.step_buttons = []
        labels = ["1  File", "2  Voce", "3  Output", "4  Riepilogo"]
        for i, label in enumerate(labels, 1):
            b = ttk.Button(stepper, text=label, style="Step.TButton",
                           command=lambda n=i: self._show_wizard_step(n))
            b.grid(row=0, column=i - 1, sticky="ew", padx=(0 if i == 1 else 4, 0))
            self.step_buttons.append(b)

        self.wizard_host = ttk.Frame(body, style="App.TFrame")
        self.wizard_host.grid(row=1, column=0, sticky="nsew")
        self.wizard_host.columnconfigure(0, weight=1)
        self.wizard_steps = {}
        self._build_step_files()
        self._build_step_voice()
        self._build_step_output()
        self._build_step_summary()

    def _build_step_files(self):
        frame = ttk.Frame(self.wizard_host, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        self.wizard_steps[1] = frame
        self.file_cards_frame = frame

        pptx_card, self.generate_input_pptx_widgets, self.generate_pptx_status = self._file_card(
            frame, "Presentazione PowerPoint",
            "Scegli il file .pptx originale da sonorizzare.",
            self.generate_input_pptx_var, self._browse_pptx,
        )
        script_card, self.generate_input_xlsx_widgets, self.generate_xlsx_status = self._file_card(
            frame, "Script della narrazione",
            "Carica un Excel oppure usa direttamente le Note PowerPoint.",
            self.generate_input_xlsx_var, self._browse_xlsx,
        )
        self._generate_pptx_card = pptx_card
        self._generate_script_card = script_card

        source = self._card(frame)
        source.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        self._section_heading(source, "Origine degli script", "Scegli dove il programma deve leggere i testi.")
        self.script_source_segment = SegmentedControl(
            source, self.generate_script_source_var,
            [("Excel", "excel"), ("Note PowerPoint", "notes")],
            command=self._on_script_source_change,
        )
        self.script_source_segment.grid(row=2, column=0, sticky="w")

        excel_opts = ttk.Frame(source, style="CardInner.TFrame")
        excel_opts.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(excel_opts, text="Foglio Excel", style="CardText.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.generate_sheet_combo = ttk.Combobox(
            excel_opts, width=24, textvariable=self.generate_sheet_name_var, state="disabled"
        )
        self.generate_sheet_combo.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._sheet_combos["generate"] = self.generate_sheet_combo
        ttk.Label(excel_opts, text="Colonna", style="CardText.TLabel").grid(
            row=0, column=1, sticky="w", padx=(18, 0)
        )
        self.column_entry = ttk.Entry(
            excel_opts, width=8, textvariable=self.generate_script_column_var
        )
        self.column_entry.grid(row=1, column=1, sticky="w", padx=(18, 0), pady=(4, 0))
        self.header_chk = ttk.Checkbutton(
            excel_opts, text="Prima riga intestazione", variable=self.generate_has_header_var
        )
        self.header_chk.grid(row=1, column=2, sticky="w", padx=(18, 0), pady=(4, 0))
        actions = ttk.Frame(excel_opts, style="CardInner.TFrame")
        actions.grid(row=1, column=3, sticky="e", padx=(18, 0), pady=(4, 0))
        self.copy_ai_prompt_btn = ttk.Button(
            actions, text="Copia prompt AI", style="Secondary.TButton",
            command=self._copy_ai_prompt,
        )
        self.copy_ai_prompt_btn.pack(side="left")
        self._ai_prompt_buttons.append(self.copy_ai_prompt_btn)
        self.template_btn = ttk.Button(
            actions, text="Crea modello Excel", style="Secondary.TButton",
            command=self._create_script_template,
        )
        self.template_btn.pack(side="left", padx=(8, 0))
        self.generate_sheet_status = ttk.Label(
            excel_opts, text="Carica un file Excel per visualizzare i fogli disponibili.",
            style="MutedCard.TLabel",
        )
        self.generate_sheet_status.grid(row=2, column=0, columnspan=4, sticky="w", pady=(7, 0))
        self._sheet_status_labels["generate"] = self.generate_sheet_status
        excel_opts.columnconfigure(3, weight=1)
        self.script_options_frame = source

        self.input_pptx_widgets = self.generate_input_pptx_widgets
        self.input_xlsx_widgets = self.generate_input_xlsx_widgets
        self._layout_file_cards("wide")

    def _layout_file_cards(self, mode: str):
        if not hasattr(self, "_generate_pptx_card"):
            return
        for card in (self._generate_pptx_card, self._generate_script_card):
            card.grid_forget()
        if mode == "compact":
            self._generate_pptx_card.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
            self._generate_script_card.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
            # source scende di una riga
            for child in self.file_cards_frame.grid_slaves(row=1):
                if child not in (self._generate_script_card,):
                    child.grid_configure(row=2)
        else:
            self._generate_pptx_card.grid(row=0, column=0, sticky="nsew", padx=(0, 7), pady=(0, 12))
            self._generate_script_card.grid(row=0, column=1, sticky="nsew", padx=(7, 0), pady=(0, 12))
            for child in self.file_cards_frame.grid_slaves():
                if child not in (self._generate_pptx_card, self._generate_script_card):
                    if int(child.grid_info().get("columnspan", 1)) == 2:
                        child.grid_configure(row=1)

    def _build_step_voice(self):
        frame = ttk.Frame(self.wizard_host, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        self.wizard_steps[2] = frame

        picker = self._card(frame)
        picker.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        top = ttk.Frame(picker, style="CardInner.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text="Voce della narrazione", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.refresh_voices_btn = ttk.Button(
            top, text="Aggiorna Microsoft", style="Secondary.TButton",
            command=self._refresh_online_voices,
        )
        self.refresh_voices_btn.grid(row=0, column=1, padx=(8, 0))
        self.manage_voices_btn = ttk.Button(
            top, text="Gestisci voci", style="Secondary.TButton",
            command=self._open_voice_manager,
        )
        self.manage_voices_btn.grid(row=0, column=2, padx=(8, 0))
        if not legacy._CLONE_AVAILABLE:
            self.manage_voices_btn.configure(state="disabled")
        ttk.Label(
            picker, text="Scegli una voce Microsoft oppure una voce clonata locale.",
            style="MutedCard.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(3, 10))
        self.voice_combo_var = tk.StringVar()
        self.voice_combo = ttk.Combobox(
            picker, textvariable=self.voice_combo_var, state="readonly"
        )
        self.voice_combo.grid(row=2, column=0, sticky="ew")
        self.voice_combo.bind("<<ComboboxSelected>>", self._on_voice_combo_selected)
        # voice_list_frame resta per compatibilità ma non viene visualizzato.
        self.voice_list_frame = ttk.Frame(picker, style="CardInner.TFrame")

        preview = self._card(frame)
        preview.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(preview, "Anteprima", "Ascolta la voce prima di avviare il progetto.")
        ttk.Entry(preview, textvariable=self.preview_text_var).grid(row=2, column=0, sticky="ew")
        preview_buttons = ttk.Frame(preview, style="CardInner.TFrame")
        preview_buttons.grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.preview_btn = ttk.Button(
            preview_buttons, text="▶  Ascolta anteprima", style="Primary.TButton",
            command=self._on_preview,
        )
        self.preview_btn.grid(row=0, column=0)
        self.stop_btn = ttk.Button(
            preview_buttons, text="Stop", style="Secondary.TButton",
            command=self._on_stop_preview,
        )
        self.stop_btn.grid(row=0, column=1, padx=(8, 0))
        if not legacy.PYGAME_INSTALLED:
            self.preview_btn.configure(state="disabled")
            self.stop_btn.configure(state="disabled")

        controls = self._card(frame)
        controls.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(controls, "Espressività", "Regola ritmo, volume e tonalità della voce.")
        controls.columnconfigure(1, weight=1)
        for row, (label, var, start, end, callback) in enumerate((
            ("Velocità", self.rate_var, -50, 50, self._on_rate_change),
            ("Volume", self.volume_var, -100, 100, self._on_voice_effect_change),
            ("Tonalità", self.pitch_var, -100, 100, self._on_voice_effect_change),
        ), start=2):
            ttk.Label(controls, text=label, style="CardText.TLabel").grid(row=row, column=0, sticky="w", pady=6)
            ttk.Scale(controls, from_=start, to=end, variable=var,
                      command=lambda _x, cb=callback: cb()).grid(
                row=row, column=1, sticky="ew", padx=14, pady=6
            )
        self.rate_label = ttk.Label(controls, text="+0%", style="CardText.TLabel", width=7, anchor="e")
        self.rate_label.grid(row=2, column=2, sticky="e")
        self.volume_label = ttk.Label(controls, text="+0%", style="CardText.TLabel", width=7, anchor="e")
        self.volume_label.grid(row=3, column=2, sticky="e")
        self.pitch_label = ttk.Label(controls, text="+0Hz", style="CardText.TLabel", width=7, anchor="e")
        self.pitch_label.grid(row=4, column=2, sticky="e")

        local_advanced = CollapsibleCard(
            frame, "Prestazioni delle voci locali",
            "Opzioni PocketTTS, processi paralleli e cache.", expanded=False,
        )
        local_advanced.grid(row=3, column=0, sticky="ew")
        la = local_advanced.content
        la.columnconfigure(0, weight=1)
        self.pq_frame = ttk.Frame(la, style="CardInner.TFrame")
        self.pq_frame.grid(row=0, column=0, sticky="ew")
        ttk.Label(self.pq_frame, text="Qualità PocketTTS", style="CardText.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(self.pq_frame, text="Veloce", value="italian",
                        variable=self.pocket_quality_var).grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Radiobutton(self.pq_frame, text="Qualità alta", value="italian_24l",
                        variable=self.pocket_quality_var).grid(row=1, column=1, sticky="w", padx=(20, 0), pady=(6, 0))
        ttk.Checkbutton(
            self.pq_frame, text="Quantizzazione INT8 sperimentale",
            variable=self.pocket_quantize_var,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.parallel_frame = ttk.Frame(la, style="CardInner.TFrame")
        self.parallel_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(self.parallel_frame, text="Slide in parallelo", style="CardText.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            self.parallel_frame, width=12, state="readonly", textvariable=self.clone_workers_var,
            values=list(range(1, 11)),
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))
        self.cache_frame = ttk.Frame(la, style="CardInner.TFrame")
        self.cache_frame.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        ttk.Checkbutton(
            self.cache_frame, text="Riusa audio identici già generati",
            variable=self.use_cache_var,
        ).grid(row=0, column=0, sticky="w")
        self.clear_cache_btn = ttk.Button(
            self.cache_frame, text="Svuota cache", style="Secondary.TButton",
            command=self._on_clear_cache,
        )
        self.clear_cache_btn.grid(row=0, column=1, padx=(12, 0))

    def _build_step_output(self):
        frame = ttk.Frame(self.wizard_host, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        self.wizard_steps[3] = frame

        self.out_frame = self._card(frame)
        self.out_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(self.out_frame, "Formato finale", "Scegli il risultato da produrre.")
        self.output_segment = SegmentedControl(
            self.out_frame, self.output_mode_var,
            [("PowerPoint con audio", "pptx"), ("Video MP4", "video")],
            command=self._on_output_mode_change,
        )
        self.output_segment.grid(row=2, column=0, sticky="ew")
        if not legacy._VIDEO_AVAILABLE:
            self.output_segment.buttons["video"].configure(state="disabled")

        out_card, self.generate_output_widgets, self.generate_output_status = self._file_card(
            frame, "File di destinazione",
            "Il file originale non verrà mai sovrascritto automaticamente.",
            self.generate_output_pptx_var, self._browse_output, "Scegli percorso",
            path_kind="output",
            allowed_suffixes=lambda: {".mp4"} if self.output_mode_var.get() == "video" else {".pptx"},
            dependent_vars=(self.output_mode_var,),
        )
        out_card.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self.output_widgets = self.generate_output_widgets

        opts_card = self._card(frame)
        opts_card.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        opts_card.columnconfigure(0, weight=1)
        self.pptx_opts_frame = ttk.Frame(opts_card, style="CardInner.TFrame")
        self.pptx_opts_frame.grid(row=0, column=0, sticky="ew")
        ttk.Label(self.pptx_opts_frame, text="Comportamento PowerPoint", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(
            self.pptx_opts_frame, text="Riproduci automaticamente l'audio",
            variable=self.autoplay_var, command=self._on_autoplay_toggle,
        ).grid(row=1, column=0, sticky="w", pady=(12, 0))
        self.auto_advance_chk = ttk.Checkbutton(
            self.pptx_opts_frame, text="Avanza alla slide successiva a fine audio",
            variable=self.auto_advance_var,
        )
        self.auto_advance_chk.grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.auto_advance_lbl = ttk.Label(
            self.pptx_opts_frame,
            text="La presentazione procede automaticamente, slide dopo slide.",
            style="MutedCard.TLabel",
        )
        self.auto_advance_lbl.grid(row=3, column=0, sticky="w", pady=(2, 0))
        ttk.Checkbutton(
            self.pptx_opts_frame,
            text="Riconverti audio per massima compatibilità PowerPoint",
            variable=self.transcode_audio_var,
        ).grid(row=4, column=0, sticky="w", pady=(12, 0))

        self.video_opts_frame = ttk.Frame(opts_card, style="CardInner.TFrame")
        self.video_opts_frame.grid(row=0, column=0, sticky="ew")
        self.video_opts_frame.columnconfigure(1, weight=1)
        ttk.Label(self.video_opts_frame, text="Impostazioni video", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(self.video_opts_frame, text="Risoluzione", style="CardText.TLabel").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Combobox(
            self.video_opts_frame, textvariable=self.resolution_var, state="readonly",
            values=["720p", "1080p"], width=16,
        ).grid(row=1, column=1, sticky="w", padx=(14, 0), pady=(12, 0))
        ttk.Checkbutton(
            self.video_opts_frame, text="Mostra sottotitoli nel video",
            variable=self.subtitles_var,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Checkbutton(
            self.video_opts_frame, text="Aggiungi transizione tra le slide",
            variable=self.transition_var,
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(
            self.video_opts_frame,
            text="Il video usa immagini statiche delle slide: animazioni interne e video incorporati non vengono riprodotti.",
            style="Warning.TLabel", wraplength=720,
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 0))

        advanced = CollapsibleCard(
            frame, "Impostazioni avanzate",
            "Backend di rendering, transizioni e durata delle slide mute.", expanded=False,
        )
        advanced.grid(row=3, column=0, sticky="ew")
        adv = advanced.content
        adv.columnconfigure(1, weight=1)
        ttk.Label(adv, text="Rendering slide", style="CardText.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            adv, textvariable=self.render_backend_var, state="readonly",
            values=["auto", "powerpoint", "libreoffice"], width=18,
        ).grid(row=0, column=1, sticky="w", padx=(14, 0))
        ttk.Label(adv, text="Stile transizione", style="CardText.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Combobox(
            adv, textvariable=self.transition_style_var, state="readonly", width=18,
            values=["fade", "wipeleft", "wiperight", "slideleft", "slideright", "dissolve"],
        ).grid(row=1, column=1, sticky="w", padx=(14, 0), pady=(10, 0))
        ttk.Label(adv, text="Slide senza audio (secondi)", style="CardText.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Spinbox(
            adv, from_=0.5, to=60.0, increment=0.5, width=9,
            textvariable=self.silent_slide_s_var,
        ).grid(row=2, column=1, sticky="w", padx=(14, 0), pady=(10, 0))
        self._on_output_mode_change()

    def _build_step_summary(self):
        frame = ttk.Frame(self.wizard_host, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        self.wizard_steps[4] = frame
        summary = self._card(frame)
        summary.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(summary, "Riepilogo del progetto", "Controlla le scelte prima di avviare la generazione.")
        keys = [
            ("PowerPoint", "pptx"), ("Script", "scripts"), ("Voce", "voice"),
            ("Output", "output"), ("Opzioni", "options"),
            ("Completezza progetto", "validation"), ("Controllo del sistema", "preflight"),
        ]
        for row, (label, key) in enumerate(keys, start=2):
            ttk.Label(summary, text=label, style="SummaryKey.TLabel", width=22).grid(
                row=row, column=0, sticky="nw", pady=6
            )
            style = "SummaryValue.TLabel"
            ttk.Label(summary, textvariable=self.summary_vars[key], style=style,
                      wraplength=720, justify="left").grid(row=row, column=1, sticky="nw", pady=6)
        summary.columnconfigure(1, weight=1)
        ttk.Button(
            frame, text="Esegui nuovamente i controlli",
            style="Secondary.TButton", command=self._refresh_summary,
        ).grid(row=1, column=0, sticky="w")

    def _build_fix_page(self):
        """Costruisce il Fix come procedura guidata in quattro passaggi.

        Il flusso separa chiaramente file, analisi, sorgente/voce e riepilogo.
        In questo modo i controlli più pesanti non sono tutti presenti nella
        stessa schermata e l'utente non può avviare il Fix saltando passaggi.
        """
        body = self._new_page("fix")
        body.columnconfigure(0, weight=1)

        stepper_card = self._card(body, padding=(12, 10))
        stepper_card.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        stepper = ttk.Frame(stepper_card, style="CardInner.TFrame")
        stepper.grid(row=0, column=0, sticky="ew")
        for i in range(4):
            stepper.columnconfigure(i, weight=1)
        self.fix_step_buttons = []
        labels = ["1  File", "2  Analisi", "3  Script e voce", "4  Riepilogo"]
        for i, label in enumerate(labels, 1):
            button = ttk.Button(
                stepper, text=label, style="Step.TButton",
                command=lambda n=i: self._show_fix_wizard_step(n),
            )
            button.grid(row=0, column=i - 1, sticky="ew", padx=(0 if i == 1 else 4, 0))
            self.fix_step_buttons.append(button)

        self.fix_wizard_host = ttk.Frame(body, style="App.TFrame")
        self.fix_wizard_host.grid(row=1, column=0, sticky="nsew")
        self.fix_wizard_host.columnconfigure(0, weight=1)
        self.fix_wizard_steps = {}
        self._build_fix_step_files()
        self._build_fix_step_analysis()
        self._build_fix_step_source_voice()
        self._build_fix_step_summary()

    def _build_fix_step_files(self):
        frame = ttk.Frame(self.fix_wizard_host, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        self.fix_wizard_steps[1] = frame
        self.fix_files_frame = frame

        pptx_card, self.fix_input_pptx_widgets, self.fix_input_status = self._file_card(
            frame, "PowerPoint da controllare",
            "Seleziona una presentazione che contiene già alcuni audio.",
            self.fix_input_pptx_var, self._browse_pptx,
            allowed_suffixes={".pptx"},
        )
        out_card, self.fix_output_widgets, self.fix_output_status = self._file_card(
            frame, "PowerPoint corretto",
            "Il risultato sarà salvato in un nuovo file .pptx.",
            self.fix_output_pptx_var, self._browse_output, "Scegli percorso",
            path_kind="output", allowed_suffixes={".pptx"},
        )
        self._fix_input_card = pptx_card
        self._fix_output_card = out_card
        self._layout_fix_file_cards("wide")

        info = self._card(frame)
        info.grid(row=1, column=0, columnspan=2, sticky="ew")
        self._section_heading(
            info, "Come funziona la riparazione",
            "Il programma analizzerà gli audio già presenti, conserverà quelli validi "
            "e genererà soltanto le narrazioni necessarie.",
        )
        ttk.Label(
            info,
            text="Il PowerPoint originale non viene mai sovrascritto automaticamente.",
            style="CardText.TLabel",
        ).grid(row=2, column=0, sticky="w")

    def _layout_fix_file_cards(self, mode: str):
        if not hasattr(self, "_fix_input_card"):
            return
        for card in (self._fix_input_card, self._fix_output_card):
            card.grid_forget()
        if mode == "compact":
            self._fix_input_card.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
            self._fix_output_card.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
            for child in self.fix_files_frame.grid_slaves():
                if child not in (self._fix_input_card, self._fix_output_card):
                    child.grid_configure(row=2)
        else:
            self._fix_input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 7), pady=(0, 12))
            self._fix_output_card.grid(row=0, column=1, sticky="nsew", padx=(7, 0), pady=(0, 12))
            for child in self.fix_files_frame.grid_slaves():
                if child not in (self._fix_input_card, self._fix_output_card):
                    child.grid_configure(row=1)

    def _build_fix_step_analysis(self):
        frame = ttk.Frame(self.fix_wizard_host, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        self.fix_wizard_steps[2] = frame

        analysis = self._card(frame)
        analysis.grid(row=0, column=0, sticky="ew")
        analysis.columnconfigure(0, weight=1)
        analysis_top = ttk.Frame(analysis, style="CardInner.TFrame")
        analysis_top.grid(row=0, column=0, sticky="ew")
        analysis_top.columnconfigure(0, weight=1)
        ttk.Label(
            analysis_top, text="Analisi della presentazione", style="CardTitle.TLabel"
        ).grid(row=0, column=0, sticky="w")
        self.analyze_fix_btn = ttk.Button(
            analysis_top, text="Analizza audio presenti", style="Primary.TButton",
            command=self._on_analyze_fix, state="disabled", width=24,
        )
        self.analyze_fix_btn.grid(row=0, column=1, sticky="e")
        ttk.Label(
            analysis, textvariable=self.fix_analysis_var, style="MutedCard.TLabel",
            wraplength=800,
        ).grid(row=1, column=0, sticky="w", pady=(4, 12))

        metrics = ttk.Frame(analysis, style="CardInner.TFrame")
        metrics.grid(row=2, column=0, sticky="ew", pady=(0, 14))
        for i in range(4):
            metrics.columnconfigure(i, weight=1)
        for col, (value_var, label) in enumerate((
            (self.fix_total_var, "Slide totali"),
            (self.fix_valid_var, "Audio validi"),
            (self.fix_missing_var, "Mancanti"),
            (self.fix_invalid_var, "Da verificare"),
        )):
            mini = ttk.Frame(metrics, style="Card.TFrame", padding=(12, 9))
            mini.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 6, 0))
            ttk.Label(mini, textvariable=value_var, style="MetricValue.TLabel").pack(anchor="w")
            ttk.Label(mini, text=label, style="MetricLabel.TLabel").pack(anchor="w")

        self.fix_filter_segment = SegmentedControl(
            analysis, self.fix_filter_var,
            [("Tutte", "all"), ("Mancanti", "missing"),
             ("Anomale", "invalid"), ("Valide", "valid")],
            command=self._apply_fix_filter,
        )
        self.fix_filter_segment.grid(row=3, column=0, sticky="w", pady=(0, 8))
        tree_wrap = ttk.Frame(analysis, style="CardInner.TFrame")
        tree_wrap.grid(row=4, column=0, sticky="nsew")
        tree_wrap.columnconfigure(0, weight=1)
        tree_wrap.rowconfigure(0, weight=1)
        self.fix_tree = ttk.Treeview(
            tree_wrap,
            columns=("stato", "formato", "durata", "autoplay", "avanza", "titolo"),
            show="tree headings", selectmode="extended", height=11,
        )
        self.fix_tree.heading("#0", text="Slide")
        specs = (("stato", "Stato", 120), ("formato", "Formato", 75),
                 ("durata", "Durata", 78), ("autoplay", "Autoplay", 78),
                 ("avanza", "Avanza", 72), ("titolo", "Titolo", 300))
        for key, title, width in specs:
            self.fix_tree.heading(key, text=title)
            self.fix_tree.column(key, width=width, stretch=(key == "titolo"))
        self.fix_tree.column("#0", width=60, stretch=False)
        self.fix_tree.grid(row=0, column=0, sticky="nsew")
        self.fix_empty_label = ttk.Label(
            tree_wrap,
            text="Nessuna analisi disponibile\nSeleziona un PowerPoint e avvia il controllo.",
            style="MutedCard.TLabel", justify="center",
        )
        self.fix_empty_label.grid(row=0, column=0)
        self.fix_analysis_frame = analysis

    def _build_fix_step_source_voice(self):
        frame = ttk.Frame(self.fix_wizard_host, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        self.fix_wizard_steps[3] = frame

        script_card, self.fix_input_xlsx_widgets, self.fix_script_status = self._file_card(
            frame, "Script per gli audio mancanti",
            "Carica un Excel oppure usa direttamente le Note PowerPoint.",
            self.fix_input_xlsx_var, self._browse_xlsx,
            allowed_suffixes={".xlsx", ".xlsm"},
        )
        script_card.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        source = self._card(frame)
        source.grid(row=1, column=0, sticky="ew")
        source.columnconfigure(0, weight=1)
        self._section_heading(
            source, "Sorgente script e voce",
            "Scegli l'origine dei testi e la voce da usare esclusivamente per i nuovi audio.",
        )
        # Compatibilità con integrazioni precedenti: la sezione è sempre aperta
        # e non espone più alcun comando Mostra/Nascondi.
        self.fix_source_card = source
        self.fix_source_card.expanded = tk.BooleanVar(value=True)  # type: ignore[attr-defined]
        self.fix_source_card.content = source  # type: ignore[attr-defined]

        ttk.Label(source, text="Origine degli script", style="CardText.TLabel").grid(
            row=2, column=0, sticky="w"
        )
        self.fix_script_source_segment = SegmentedControl(
            source, self.fix_script_source_var,
            [("Excel", "excel"), ("Note PowerPoint", "notes")],
            command=self._on_script_source_change,
        )
        self.fix_script_source_segment.grid(row=3, column=0, sticky="w", pady=(6, 12))

        fix_excel_opts = ttk.Frame(source, style="CardInner.TFrame")
        fix_excel_opts.grid(row=4, column=0, sticky="ew", pady=(0, 14))
        ttk.Label(fix_excel_opts, text="Foglio Excel", style="CardText.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.fix_sheet_combo = ttk.Combobox(
            fix_excel_opts, width=24, textvariable=self.fix_sheet_name_var, state="disabled"
        )
        self.fix_sheet_combo.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._sheet_combos["fix"] = self.fix_sheet_combo
        ttk.Label(fix_excel_opts, text="Colonna", style="CardText.TLabel").grid(
            row=0, column=1, sticky="w", padx=(18, 0)
        )
        self.fix_column_entry = ttk.Entry(
            fix_excel_opts, width=8, textvariable=self.fix_script_column_var
        )
        self.fix_column_entry.grid(row=1, column=1, sticky="w", padx=(18, 0), pady=(4, 0))
        self.fix_header_chk = ttk.Checkbutton(
            fix_excel_opts, text="Prima riga intestazione", variable=self.fix_has_header_var
        )
        self.fix_header_chk.grid(row=1, column=2, sticky="w", padx=(18, 0), pady=(4, 0))
        self.fix_sheet_status = ttk.Label(
            fix_excel_opts, text="Carica un file Excel per visualizzare i fogli disponibili.",
            style="MutedCard.TLabel",
        )
        self.fix_sheet_status.grid(row=2, column=0, columnspan=3, sticky="w", pady=(7, 0))
        self._sheet_status_labels["fix"] = self.fix_sheet_status
        fix_excel_opts.columnconfigure(2, weight=1)
        self.fix_excel_options_frame = fix_excel_opts

        ttk.Separator(source, orient="horizontal").grid(
            row=5, column=0, sticky="ew", pady=(0, 14)
        )
        voice_top = ttk.Frame(source, style="CardInner.TFrame")
        voice_top.grid(row=6, column=0, sticky="ew")
        voice_top.columnconfigure(0, weight=1)
        ttk.Label(
            voice_top, text="Voce per gli audio mancanti", style="CardText.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            voice_top, text="Aggiorna voci", style="Secondary.TButton",
            command=self._refresh_online_voices,
        ).grid(row=0, column=1, padx=(10, 0))
        self.fix_manage_voices_btn = ttk.Button(
            voice_top, text="Gestisci voci", style="Secondary.TButton",
            command=self._open_voice_manager,
        )
        self.fix_manage_voices_btn.grid(row=0, column=2, padx=(8, 0))
        if not legacy._CLONE_AVAILABLE:
            self.fix_manage_voices_btn.configure(state="disabled")

        self.fix_voice_combo = ttk.Combobox(
            source, textvariable=self.fix_voice_combo_var, state="readonly"
        )
        self.fix_voice_combo.grid(row=7, column=0, sticky="ew", pady=(8, 14))
        self.fix_voice_combo.bind("<<ComboboxSelected>>", self._on_fix_voice_combo_selected)

        effects = ttk.Frame(source, style="CardInner.TFrame")
        effects.grid(row=8, column=0, sticky="ew")
        effects.columnconfigure(1, weight=1)
        for row, (label, var, start, end, callback) in enumerate((
            ("Velocità", self.rate_var, -50, 50, self._on_rate_change),
            ("Volume", self.volume_var, -100, 100, self._on_voice_effect_change),
            ("Tonalità", self.pitch_var, -100, 100, self._on_voice_effect_change),
        )):
            ttk.Label(effects, text=label, style="CardText.TLabel").grid(
                row=row, column=0, sticky="w", pady=4
            )
            ttk.Scale(
                effects, from_=start, to=end, variable=var,
                command=lambda _x, cb=callback: cb(),
            ).grid(row=row, column=1, sticky="ew", padx=(14, 12), pady=4)
        self.fix_rate_label = ttk.Label(
            effects, text="+0%", style="CardText.TLabel", width=7, anchor="e"
        )
        self.fix_rate_label.grid(row=0, column=2, sticky="e")
        self.fix_volume_label = ttk.Label(
            effects, text="+0%", style="CardText.TLabel", width=7, anchor="e"
        )
        self.fix_volume_label.grid(row=1, column=2, sticky="e")
        self.fix_pitch_label = ttk.Label(
            effects, text="+0Hz", style="CardText.TLabel", width=7, anchor="e"
        )
        self.fix_pitch_label.grid(row=2, column=2, sticky="e")

    def _build_fix_step_summary(self):
        frame = ttk.Frame(self.fix_wizard_host, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        self.fix_wizard_steps[4] = frame

        summary = self._card(frame)
        summary.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(
            summary, "Riepilogo della riparazione",
            "Controlla i file e le operazioni prima di creare il nuovo PowerPoint.",
        )
        rows = (
            ("PowerPoint", "pptx"), ("Analisi", "analysis"),
            ("Script", "scripts"), ("Voce", "voice"),
            ("Risultato", "output"), ("Azioni", "actions"),
            ("Stato", "validation"),
        )
        for row, (label, key) in enumerate(rows, start=2):
            ttk.Label(summary, text=label, style="MutedCard.TLabel").grid(
                row=row, column=0, sticky="nw", pady=5
            )
            style = "Success.TLabel" if key == "validation" else "CardText.TLabel"
            ttk.Label(
                summary, textvariable=self.fix_summary_vars[key], style=style,
                wraplength=720, justify="left",
            ).grid(row=row, column=1, sticky="w", padx=(24, 0), pady=5)
        summary.columnconfigure(1, weight=1)

        advanced = self._card(frame)
        advanced.grid(row=1, column=0, sticky="ew")
        self._section_heading(
            advanced, "Azioni avanzate del Fix",
            "Le opzioni selezionate vengono applicate oltre al completamento degli audio mancanti.",
        )
        self.fix_advanced_card = advanced
        advanced.columnconfigure(0, weight=1)
        advanced.columnconfigure(1, weight=1)
        ttk.Checkbutton(
            advanced, text="Rigenera tutte le slide", variable=self.fix_regenerate_all_var,
            command=self._on_fix_option_change,
        ).grid(row=2, column=0, sticky="w", pady=5)
        ttk.Checkbutton(
            advanced, text="Ripara audio corrotti o anomali", variable=self.fix_repair_invalid_var,
            command=self._on_fix_option_change,
        ).grid(row=2, column=1, sticky="w", padx=(24, 0), pady=5)
        ttk.Checkbutton(
            advanced, text="Uniforma autoplay", variable=self.fix_normalize_autoplay_var,
            command=self._on_fix_option_change,
        ).grid(row=3, column=0, sticky="w", pady=5)
        ttk.Checkbutton(
            advanced, text="Uniforma avanzamento", variable=self.fix_normalize_advance_var,
            command=self._on_fix_option_change,
        ).grid(row=3, column=1, sticky="w", padx=(24, 0), pady=5)
        ttk.Checkbutton(
            advanced, text="Nascondi tutte le icone audio", variable=self.fix_normalize_icon_var,
            command=self._on_fix_option_change,
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=5)

    def _on_fix_option_change(self):
        self._refresh_fix_summary()
        self._update_action_bar()

    def _build_batch_page(self):
        body = self._new_page("batch")
        card = self._card(body)
        card.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(card, "Coda dei lavori", "Carica un manifest JSON oppure usa il modello di esempio.")
        row = ttk.Frame(card, style="CardInner.TFrame")
        row.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        row.columnconfigure(0, weight=1)
        ttk.Entry(row, textvariable=self.batch_manifest_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(row, text="Seleziona manifest", style="Secondary.TButton",
                   command=self._browse_batch_manifest).grid(row=0, column=1, padx=(10, 0))
        ttk.Button(row, text="Apri esempio", style="Secondary.TButton",
                   command=lambda: self._open_file(str(Path(__file__).with_name("ESEMPIO_BATCH.json")))).grid(row=0, column=2, padx=(8, 0))
        ttk.Label(card, textvariable=self.batch_jobs_var, style="MutedCard.TLabel").grid(row=3, column=0, sticky="w", pady=(0, 8))
        self.batch_tree = ttk.Treeview(card, columns=("modo", "input", "script", "output"), show="headings", height=8)
        for key, title, width in (("modo", "Operazione", 100), ("input", "PowerPoint", 250),
                                  ("script", "Script", 210), ("output", "Risultato", 260)):
            self.batch_tree.heading(key, text=title)
            self.batch_tree.column(key, width=width, stretch=key in {"input", "output"})
        self.batch_tree.grid(row=4, column=0, sticky="ew")

        opts = self._card(body)
        opts.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(opts, "Comportamento", "I lavori già completati possono essere saltati automaticamente.")
        ttk.Checkbutton(opts, text="Rielabora anche gli output già completati",
                        variable=self.batch_force_var).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(opts, text="Interrompi al primo errore",
                        variable=self.batch_stop_on_error_var).grid(row=3, column=0, sticky="w", pady=(8, 0))
        status = self._card(body)
        status.grid(row=2, column=0, sticky="ew")
        self._section_heading(status, "Stato batch")
        ttk.Label(status, textvariable=self.batch_result_var, style="CardText.TLabel",
                  wraplength=760).grid(row=2, column=0, sticky="w")

    def _build_voices_page(self):
        body = self._new_page("voices")
        top = self._card(body)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        top.columnconfigure(0, weight=1)
        self._section_heading(top, "Voci disponibili", "Cerca, ascolta e seleziona una voce per il prossimo progetto.")
        filters = ttk.Frame(top, style="CardInner.TFrame")
        filters.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filters.columnconfigure(0, weight=1)
        search = ttk.Entry(filters, textvariable=self.voice_search_var)
        search.grid(row=0, column=0, sticky="ew")
        engine = ttk.Combobox(filters, textvariable=self.voice_engine_filter_var, state="readonly",
                              values=["Tutti", "Microsoft", "PocketTTS", "Chatterbox", "XTTS"], width=16)
        engine.grid(row=0, column=1, padx=(8, 0))
        self.voice_search_var.trace_add("write", lambda *_: self._refresh_voice_library_page())
        engine.bind("<<ComboboxSelected>>", lambda _e: self._refresh_voice_library_page())

        actions = ttk.Frame(top, style="CardInner.TFrame")
        actions.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        ttk.Button(actions, text="Aggiorna voci Microsoft", style="Secondary.TButton",
                   command=self._refresh_online_voices).pack(side="left")
        self.library_manage_btn = ttk.Button(actions, text="Gestisci voci clonate",
                                             style="Secondary.TButton", command=self._open_voice_manager)
        self.library_manage_btn.pack(side="left", padx=(8, 0))
        if not legacy._CLONE_AVAILABLE:
            self.library_manage_btn.configure(state="disabled")
        self.library_preview_btn = ttk.Button(actions, text="Ascolta", style="Secondary.TButton",
                                              command=self._preview_library_voice, state="disabled")
        self.library_preview_btn.pack(side="right", padx=(8, 0))
        self.library_use_btn = ttk.Button(actions, text="Usa nel nuovo progetto", style="Primary.TButton",
                                          command=self._use_library_voice, state="disabled")
        self.library_use_btn.pack(side="right")
        self.library_tree = ttk.Treeview(
            top, columns=("nome", "motore", "tipo", "stato"), show="headings",
            selectmode="browse", height=12,
        )
        for key, text, width in (("nome", "Nome", 260), ("motore", "Motore", 170),
                                 ("tipo", "Tipo", 150), ("stato", "Stato", 110)):
            self.library_tree.heading(key, text=text)
            self.library_tree.column(key, width=width, stretch=(key == "nome"))
        self.library_tree.grid(row=4, column=0, sticky="ew")
        self.library_tree.bind("<<TreeviewSelect>>", lambda _e: self._update_voice_actions())
        self.library_tree.bind("<Double-1>", lambda _e: self._use_library_voice())

    def _build_history_page(self):
        body = self._new_page("history")
        card = self._card(body)
        card.grid(row=0, column=0, sticky="ew")
        card.columnconfigure(0, weight=1)
        self._section_heading(card, "Rapporti di esecuzione", "Consulta i risultati delle elaborazioni precedenti.")
        folder = ttk.Frame(card, style="CardInner.TFrame")
        folder.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        folder.columnconfigure(0, weight=1)
        ttk.Entry(folder, textvariable=self.history_folder_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(folder, text="Cartella", style="Secondary.TButton",
                   command=self._browse_history_folder).grid(row=0, column=1, padx=(8, 0))
        self.history_refresh_btn = ttk.Button(
            folder, text="Aggiorna", style="Primary.TButton", command=self._scan_history
        )
        self.history_refresh_btn.grid(row=0, column=2, padx=(8, 0))
        tree_wrap = ttk.Frame(card, style="CardInner.TFrame")
        tree_wrap.grid(row=3, column=0, sticky="ew")
        tree_wrap.columnconfigure(0, weight=1)
        self.history_tree = ttk.Treeview(
            tree_wrap, columns=("data", "stato", "operazione", "output"), show="headings",
            selectmode="browse", height=12,
        )
        for key, title, width in (("data", "Data", 145), ("stato", "Stato", 100),
                                  ("operazione", "Operazione", 130), ("output", "Risultato", 430)):
            self.history_tree.heading(key, text=title)
            self.history_tree.column(key, width=width, stretch=(key == "output"))
        self.history_tree.grid(row=0, column=0, sticky="ew")
        self.history_empty_label = ttk.Label(
            tree_wrap, text="Nessuna elaborazione trovata\nCompleta il primo progetto o seleziona un'altra cartella.",
            style="MutedCard.TLabel", justify="center"
        )
        self.history_tree.bind("<<TreeviewSelect>>", lambda _e: self._update_history_actions())
        buttons = ttk.Frame(card, style="CardInner.TFrame")
        buttons.grid(row=4, column=0, sticky="e", pady=(10, 0))
        self.history_report_btn = ttk.Button(buttons, text="Apri rapporto", style="Secondary.TButton",
                                             command=self._open_selected_history_report, state="disabled")
        self.history_report_btn.pack(side="left")
        self.history_output_btn = ttk.Button(buttons, text="Apri risultato", style="Primary.TButton",
                                             command=self._open_selected_history_output, state="disabled")
        self.history_output_btn.pack(side="left", padx=(8, 0))

    def _build_settings_page(self):
        body = self._new_page("settings")
        appearance = self._card(body)
        appearance.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(appearance, "Aspetto", "Il tema Sistema segue la preferenza di Windows quando disponibile.")
        self.theme_segment = SegmentedControl(
            appearance, self.theme_mode_var,
            [("Sistema", "system"), ("Chiaro", "light"), ("Scuro", "dark")],
            command=self._on_theme_mode_change,
        )
        self.theme_segment.grid(row=2, column=0, sticky="w")

        behaviour = self._card(body)
        behaviour.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(behaviour, "Comportamento")
        ttk.Checkbutton(
            behaviour, text="Apri automaticamente i dettagli tecnici in caso di errore",
            variable=self.show_log_on_error_var,
        ).grid(row=2, column=0, sticky="w")

        tools = self._card(body)
        tools.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        tools.columnconfigure(0, weight=1)
        self._section_heading(
            tools, "Strumenti e tecnologie",
            "Scarica i componenti esterni usati per creare video e convertire le presentazioni.",
        )
        self.tool_status_vars = {
            "ffmpeg": tk.StringVar(value="Controllo in corso…"),
            "libreoffice": tk.StringVar(value="Controllo in corso…"),
            "powerpoint": tk.StringVar(value="Controllo in corso…"),
            "winget": tk.StringVar(value="Controllo in corso…"),
        }
        tool_rows = [
            ("FFmpeg", "Montaggio audio/video e sottotitoli.", "ffmpeg",
             "Installa FFmpeg", lambda: self._install_windows_tool("ffmpeg")),
            ("LibreOffice", "Rendering delle slide quando PowerPoint non è disponibile.", "libreoffice",
             "Installa LibreOffice", lambda: self._install_windows_tool("libreoffice")),
            ("Microsoft PowerPoint", "Rendering nativo ad alta fedeltà, richiede una licenza Microsoft 365.", "powerpoint",
             "Apri Microsoft 365", self._open_microsoft_365),
        ]
        self.tool_install_buttons = {}
        for row_index, (title, description, key, button_text, command) in enumerate(tool_rows, start=2):
            row = ttk.Frame(tools, style="CardInner.TFrame")
            row.grid(row=row_index, column=0, sticky="ew", pady=(0, 10))
            row.columnconfigure(0, weight=1)
            ttk.Label(row, text=title, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(row, text=description, style="MutedCard.TLabel", wraplength=600).grid(
                row=1, column=0, sticky="w", pady=(2, 0)
            )
            ttk.Label(row, textvariable=self.tool_status_vars[key], style="CardText.TLabel").grid(
                row=2, column=0, sticky="w", pady=(3, 0)
            )
            btn = ttk.Button(row, text=button_text, style="Secondary.TButton", command=command)
            btn.grid(row=0, column=1, rowspan=3, sticky="e", padx=(12, 0))
            self.tool_install_buttons[key] = btn
        self.tools_note_var = tk.StringVar(
            value="I motori vocali locali (PocketTTS, Chatterbox e XTTS) non vengono aggiunti a una build EXE già compilata."
        )
        ttk.Label(tools, textvariable=self.tools_note_var, style="MutedCard.TLabel", wraplength=820).grid(
            row=5, column=0, sticky="w", pady=(2, 0)
        )
        ttk.Button(
            tools, text="Aggiorna stato", style="Secondary.TButton", command=self._refresh_tools_status
        ).grid(row=6, column=0, sticky="w", pady=(10, 0))

        diagnostics = self._card(body)
        diagnostics.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        self._section_heading(diagnostics, "Diagnostica", "Controlla gli strumenti richiesti per PowerPoint e video.")
        self.diagnostics_var = tk.StringVar(value="Controllo non ancora eseguito.")
        ttk.Label(diagnostics, textvariable=self.diagnostics_var, style="CardText.TLabel",
                  wraplength=820).grid(row=2, column=0, sticky="w")
        self.diagnostics_btn = ttk.Button(
            diagnostics, text="Esegui diagnostica", style="Secondary.TButton",
            command=self._run_diagnostics,
        )
        self.diagnostics_btn.grid(row=3, column=0, sticky="w", pady=(10, 0))

        info = self._card(body)
        info.grid(row=4, column=0, sticky="ew")
        self._section_heading(info, "Informazioni")
        ttk.Label(info, text=f"{APP_NAME} — {APP_VERSION}", style="CardText.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Label(info,
            text=("Interfaccia basata su ttkbootstrap 2.x." if _BOOTSTRAP_AVAILABLE
                  else "ttkbootstrap non disponibile: è attivo il tema di compatibilità."),
            style="MutedCard.TLabel").grid(row=3, column=0, sticky="w", pady=(4, 0))
        self.root.after_idle(self._refresh_tools_status)

    def _build_result_page(self):
        body = self._new_page("result", scrollable=False)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        center = ttk.Frame(body, style="App.TFrame")
        center.grid(row=0, column=0)
        ttk.Label(center, text="✓", style="SuccessTitle.TLabel",
                  font=("Segoe UI Variable Display", 50, "bold")).pack()
        ttk.Label(center, textvariable=self.result_title_var, style="SuccessTitle.TLabel").pack(pady=(8, 4))
        ttk.Label(center, textvariable=self.result_details_var, style="PageSubtitle.TLabel",
                  justify="center", wraplength=680).pack(pady=(0, 18))
        result_card = self._card(center, padding=(22, 18))
        result_card.pack(fill="x", pady=(0, 18))
        ttk.Label(result_card, text="File generato", style="SummaryKey.TLabel").pack(anchor="w")
        ttk.Label(result_card, textvariable=self.result_path_var, style="SummaryValue.TLabel",
                  wraplength=650).pack(anchor="w", pady=(4, 0))
        actions = ttk.Frame(center, style="App.TFrame")
        actions.pack()
        ttk.Button(actions, text="Apri file", style="Primary.TButton",
                   command=lambda: self._open_file(self.result_path_var.get())).pack(side="left")
        ttk.Button(actions, text="Apri cartella", style="Secondary.TButton",
                   command=self._open_result_folder).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Nuovo progetto", style="Secondary.TButton",
                   command=lambda: self._show_page("generate")).pack(side="left", padx=(8, 0))

    # ------------------------------------------------------------- navigation
    def _show_page(self, key: str):
        if key not in self.pages:
            key = "generate"
        previous_key = getattr(self, "current_page", None)
        self.current_page = key
        target = self.pages[key]

        # Disattiviamo prima la pagina precedente. Le pagine nascoste non devono
        # partecipare ai ricalcoli geometrici durante resize o cambio sezione.
        for name, page in self.pages.items():
            if name == key:
                continue
            if isinstance(page, ScrollablePage):
                page.set_active(False)
            page.grid_remove()

        # Il page_host mantiene la larghezza finale anche mentre le sue pagine
        # sono nascoste. Usiamola per dimensionare il canvas-window *prima* di
        # mappare la pagina: in questo modo non appare prima stretta e poi larga.
        if isinstance(target, ScrollablePage):
            try:
                host_width = int(self.page_host.winfo_width())
            except Exception:
                host_width = 0
            target.prepare_for_show(host_width)

        target.grid(row=0, column=0, sticky="nsew")
        target.tkraise()
        if isinstance(target, ScrollablePage):
            target.set_active(True)

        # Se l'utente riclicca la sezione corrente, manteniamo posizione e
        # geometria senza innescare una seconda animazione/rimappatura.
        if previous_key == key and isinstance(target, ScrollablePage):
            target._queue_region_sync()
        title, subtitle = self.PAGE_META[key]
        self.page_title.configure(text=title)
        self.page_subtitle.configure(text=subtitle)
        self._update_sidebar_selection(key)
        if key == "generate":
            self.operation_mode_var.set("generate")
            self.input_pptx_var = self.generate_input_pptx_var
            self.input_xlsx_var = self.generate_input_xlsx_var
            self.output_pptx_var = self.generate_output_pptx_var
            self.sheet_name_var = self.generate_sheet_name_var
            self.script_column_var = self.generate_script_column_var
            self.has_header_var = self.generate_has_header_var
            self.script_source_var = self.generate_script_source_var
            self.input_pptx_widgets = self.generate_input_pptx_widgets
            self.input_xlsx_widgets = self.generate_input_xlsx_widgets
            self.output_widgets = self.generate_output_widgets
            self._on_output_mode_change()
        elif key == "fix":
            self.operation_mode_var.set("fix")
            self.output_mode_var.set("pptx")
            self.input_pptx_var = self.fix_input_pptx_var
            self.input_xlsx_var = self.fix_input_xlsx_var
            self.output_pptx_var = self.fix_output_pptx_var
            self.sheet_name_var = self.fix_sheet_name_var
            self.script_column_var = self.fix_script_column_var
            self.has_header_var = self.fix_has_header_var
            self.script_source_var = self.fix_script_source_var
            self.input_pptx_widgets = self.fix_input_pptx_widgets
            self.input_xlsx_widgets = self.fix_input_xlsx_widgets
            self.output_widgets = self.fix_output_widgets
            self._sync_output_extension()
        elif key == "voices":
            self._refresh_voice_library_page()
        elif key == "history":
            self._scan_history(silent=True)
        self._on_script_source_change()
        self._update_action_bar()

    def _show_wizard_step(self, step: int, force: bool = False):
        step = max(1, min(4, int(step)))
        current = self.wizard_step.get()
        if not force and step > current:
            for required in range(current, step):
                errors = self._validate_generate_step(required)
                if errors:
                    messagebox.showwarning("Completa il passaggio", "\n".join(errors))
                    return False
        self.wizard_step.set(step)
        self._max_unlocked_step = max(self._max_unlocked_step, step)
        for number, frame in self.wizard_steps.items():
            if number == step:
                frame.tkraise()
        for i, btn in enumerate(self.step_buttons, 1):
            style = "StepActive.TButton" if i == step else ("StepDone.TButton" if i < step else "Step.TButton")
            btn.configure(style=style)
        if step == 4:
            self._refresh_summary()
        self._update_action_bar()
        return True

    def _show_fix_wizard_step(self, step: int, force: bool = False):
        step = max(1, min(4, int(step)))
        current = self.fix_wizard_step.get()
        if not force and step > current:
            for required in range(current, step):
                errors = self._validate_fix_step(required)
                if errors:
                    messagebox.showwarning("Completa il passaggio", "\n".join(errors))
                    return False
        self.fix_wizard_step.set(step)
        self._fix_max_unlocked_step = max(self._fix_max_unlocked_step, step)
        for number, frame in self.fix_wizard_steps.items():
            if number == step:
                frame.tkraise()
        for i, btn in enumerate(self.fix_step_buttons, 1):
            style = (
                "StepActive.TButton" if i == step
                else ("StepDone.TButton" if i < step else "Step.TButton")
            )
            btn.configure(style=style)
        if step == 4:
            self._refresh_fix_summary()
        self._update_action_bar()
        return True

    def _wizard_back(self):
        if self.current_page == "generate" and self.wizard_step.get() > 1:
            self._show_wizard_step(self.wizard_step.get() - 1)
        elif self.current_page == "fix" and self.fix_wizard_step.get() > 1:
            self._show_fix_wizard_step(self.fix_wizard_step.get() - 1)

    def _on_primary_action(self):
        if self.current_page == "generate":
            step = self.wizard_step.get()
            if step < 4:
                errors = self._validate_generate_step(step)
                if errors:
                    messagebox.showwarning("Dati mancanti", "\n".join(errors))
                    return
                self._show_wizard_step(step + 1, force=True)
            else:
                errors = self._validate_generate_project()
                if errors:
                    messagebox.showerror("Progetto incompleto", "\n".join(errors))
                    return
                self._on_generate()
        elif self.current_page == "fix":
            step = self.fix_wizard_step.get()
            if step < 4:
                errors = self._validate_fix_step(step)
                if errors:
                    messagebox.showwarning("Dati mancanti", "\n".join(errors))
                    return
                self._show_fix_wizard_step(step + 1, force=True)
            else:
                errors = self._validate_fix_project()
                if errors:
                    messagebox.showerror("Fix non pronto", "\n".join(errors))
                    return
                self._on_generate()
        elif self.current_page == "batch":
            errors = self._validate_batch()
            if errors:
                messagebox.showerror("Batch non pronto", "\n".join(errors))
                return
            self._on_run_batch()

    def _update_action_bar(self):
        if not hasattr(self, "footer"):
            return
        page = self.current_page
        if page in {"generate", "fix", "batch"}:
            self.footer.grid()
        else:
            self.footer.grid_remove()
            return
        if self._is_busy:
            self._show_footer_progress()
            self.generate_btn.configure(state="disabled")
            return
        self._show_footer_idle()
        self.cancel_btn.configure(state="disabled")
        if page == "generate":
            step = self.wizard_step.get()
            self.footer_step_label.configure(text=f"Passaggio {step} di 4")
            self.back_btn.configure(state="normal" if step > 1 else "disabled")
            label = "Avanti" if step < 4 else self._generate_btn_label()
            errors = self._validate_generate_step(step) if step < 4 else self._validate_generate_project()
            self.generate_btn.configure(text=label, state="disabled" if errors else "normal")
        elif page == "fix":
            step = self.fix_wizard_step.get()
            self.footer_step_label.configure(text=f"Passaggio {step} di 4")
            self.back_btn.configure(state="normal" if step > 1 else "disabled")
            path_ok = Path(self.fix_input_pptx_var.get().strip()).is_file()
            if hasattr(self, "analyze_fix_btn"):
                self.analyze_fix_btn.configure(state="normal" if path_ok else "disabled")
            label = "Avanti" if step < 4 else "Completa/Ripara PowerPoint"
            errors = self._validate_fix_step(step) if step < 4 else self._validate_fix_project()
            self.generate_btn.configure(text=label, state="disabled" if errors else "normal")
        else:
            self.footer_step_label.configure(text=self.batch_jobs_var.get())
            self.back_btn.configure(state="disabled")
            self.generate_btn.configure(text="Avvia elaborazione batch", state="disabled" if self._validate_batch() else "normal")

    def _validate_generate_step(self, step: int) -> list[str]:
        errors: list[str] = []
        if step == 1:
            pptx = self.generate_input_pptx_var.get().strip()
            if not pptx or not Path(pptx).is_file() or Path(pptx).suffix.lower() != ".pptx":
                errors.append("Seleziona una presentazione PowerPoint .pptx valida.")
            if self.generate_script_source_var.get() == "excel":
                xlsx = self.generate_input_xlsx_var.get().strip()
                if not xlsx or not Path(xlsx).is_file() or Path(xlsx).suffix.lower() not in {".xlsx", ".xlsm"}:
                    errors.append("Seleziona un file Excel valido oppure usa le Note PowerPoint.")
        elif step == 2 and not self.voice_var.get().strip():
            errors.append("Seleziona una voce per la narrazione.")
        elif step == 3:
            out = self.generate_output_pptx_var.get().strip()
            suffixes = {".mp4"} if self.output_mode_var.get() == "video" else {".pptx"}
            valid, message, _style = self._save_path_status(out, suffixes)
            if not valid:
                errors.append(message if out else "Scegli il percorso del file di destinazione.")
            elif self.generate_input_pptx_var.get().strip() and Path(out).resolve() == Path(self.generate_input_pptx_var.get()).resolve():
                errors.append("Il file di destinazione deve essere diverso dall'originale.")
        return errors

    def _validate_generate_project(self) -> list[str]:
        errors: list[str] = []
        for step in (1, 2, 3): errors.extend(self._validate_generate_step(step))
        return list(dict.fromkeys(errors))

    def _fix_requires_scripts(self) -> bool:
        result = self._last_fix_analysis or {}
        return bool(
            result.get("missing") or result.get("invalid")
            or self.fix_regenerate_all_var.get() or self.fix_repair_invalid_var.get()
        )

    def _validate_fix_step(self, step: int) -> list[str]:
        errors: list[str] = []
        inp = self.fix_input_pptx_var.get().strip()
        out = self.fix_output_pptx_var.get().strip()
        if step == 1:
            if not inp or not Path(inp).is_file() or Path(inp).suffix.lower() != ".pptx":
                errors.append("Seleziona il PowerPoint da controllare.")
            valid_out, output_message, _style = self._save_path_status(out, {".pptx"})
            if not valid_out:
                errors.append(output_message if out else "Scegli il percorso del PowerPoint corretto.")
            elif inp and Path(out).resolve() == Path(inp).resolve():
                errors.append("Il risultato deve essere salvato in un nuovo file.")
        elif step == 2:
            if not inp or not Path(inp).is_file():
                errors.append("Seleziona prima il PowerPoint da controllare.")
            elif self._last_fix_analysis is None or self._last_analyzed_fix_path != inp:
                errors.append("Esegui l'analisi degli audio presenti prima di continuare.")
        elif step == 3 and self._fix_requires_scripts():
            if not self.fix_voice_var.get().strip():
                errors.append("Seleziona la voce da usare per gli audio mancanti.")
            if self.fix_script_source_var.get() == "excel":
                xlsx = self.fix_input_xlsx_var.get().strip()
                if (not xlsx or not Path(xlsx).is_file()
                        or Path(xlsx).suffix.lower() not in {".xlsx", ".xlsm"}):
                    errors.append(
                        "Per generare gli audio mancanti seleziona il file Excel degli script."
                    )
                elif not self.fix_sheet_name_var.get().strip():
                    errors.append("Seleziona il foglio Excel che contiene gli script.")
                column = self.fix_script_column_var.get().strip().upper()
                if not re.fullmatch(r"[A-Z]+", column):
                    errors.append("Indica una colonna Excel valida, ad esempio A oppure C.")
        return list(dict.fromkeys(errors))

    def _validate_fix_project(self) -> list[str]:
        errors: list[str] = []
        for step in (1, 2, 3):
            errors.extend(self._validate_fix_step(step))
        return list(dict.fromkeys(errors))

    def _validate_batch(self) -> list[str]:
        manifest = self.batch_manifest_var.get().strip(); errors = []
        if not manifest or not Path(manifest).is_file(): errors.append("Seleziona un manifest JSON valido.")
        elif not self._batch_jobs: errors.append("Il manifest non contiene lavori validi.")
        return errors

    def _rebuild_voice_list(self):
        entries = legacy.get_voice_entries()
        self._voice_display_to_id.clear()
        self._voice_id_to_display.clear()
        displays = []
        for entry in entries:
            suffix = "Microsoft" if not entry["is_clone"] else entry["engine"]
            display = f"{entry['name']}  ·  {suffix}"
            displays.append(display)
            self._voice_display_to_id[display] = entry["voice_id"]
            self._voice_id_to_display[entry["voice_id"]] = display
        valid_ids = set(self._voice_id_to_display)
        if self.voice_var.get() not in valid_ids and entries:
            self.voice_var.set(entries[0]["voice_id"])
        if hasattr(self, "voice_combo"):
            self.voice_combo.configure(values=displays)
            self.voice_combo_var.set(
                self._voice_id_to_display.get(self.voice_var.get(), displays[0] if displays else "")
            )
        if self.fix_voice_var.get() not in valid_ids and entries:
            self.fix_voice_var.set(entries[0]["voice_id"])
        if hasattr(self, "fix_voice_combo"):
            self.fix_voice_combo.configure(values=displays)
            self.fix_voice_combo_var.set(
                self._voice_id_to_display.get(
                    self.fix_voice_var.get(), displays[0] if displays else ""
                )
            )
        self._update_pocket_quality_visibility()
        if hasattr(self, "library_tree"):
            self._refresh_voice_library_page()

    def _on_voice_combo_selected(self, _event=None):
        voice_id = self._voice_display_to_id.get(self.voice_combo_var.get())
        if voice_id:
            self.voice_var.set(voice_id)
            self._update_pocket_quality_visibility()

    def _on_fix_voice_combo_selected(self, _event=None):
        voice_id = self._voice_display_to_id.get(self.fix_voice_combo_var.get())
        if voice_id:
            self.fix_voice_var.set(voice_id)
            self._update_action_bar()

    def _on_rate_change(self):
        legacy.App._on_rate_change(self)
        if hasattr(self, "fix_rate_label"):
            value = int(round(self.rate_var.get()))
            self.fix_rate_label.configure(text=f"{'+' if value >= 0 else ''}{value}%")

    def _on_voice_effect_change(self):
        legacy.App._on_voice_effect_change(self)
        if hasattr(self, "fix_volume_label"):
            volume = int(round(self.volume_var.get()))
            pitch = int(round(self.pitch_var.get()))
            self.fix_volume_label.configure(text=f"{'+' if volume >= 0 else ''}{volume}%")
            self.fix_pitch_label.configure(text=f"{'+' if pitch >= 0 else ''}{pitch}Hz")

    def _on_generate(self):
        """Nel Fix usa sorgente e voce dedicate senza alterare Nuovo progetto."""
        if self.operation_mode_var.get() != "fix":
            return legacy.App._on_generate(self)
        original_voice_var = self.voice_var
        original_source_var = self.script_source_var
        try:
            self.voice_var = self.fix_voice_var
            self.script_source_var = self.fix_script_source_var
            return legacy.App._on_generate(self)
        finally:
            self.voice_var = original_voice_var
            self.script_source_var = original_source_var

    def _on_operation_mode_change(self):
        # La navigazione laterale determina la modalità; non mostriamo più un
        # selettore duplicato nella pagina.
        if self.operation_mode_var.get() == "fix":
            self.output_mode_var.set("pptx")
        self._sync_output_extension()
        self._update_action_bar()

    def _on_output_mode_change(self):
        if not hasattr(self, "video_opts_frame"):
            return
        if self.operation_mode_var.get() == "fix":
            self.output_mode_var.set("pptx")
        if hasattr(self, "output_segment"):
            self.output_segment.refresh()
        if self.output_mode_var.get() == "video" and self.operation_mode_var.get() != "fix":
            self.pptx_opts_frame.grid_remove()
            self.video_opts_frame.grid()
        else:
            self.video_opts_frame.grid_remove()
            self.pptx_opts_frame.grid()
            self._on_autoplay_toggle()
        self._sync_output_extension()
        self._update_action_bar()

    def _on_script_source_change(self):
        def configure_mode(mode: str, excel: bool):
            state = "normal" if excel else "disabled"
            widgets = (
                getattr(self, "generate_input_xlsx_widgets", ())
                if mode == "generate" else getattr(self, "fix_input_xlsx_widgets", ())
            )
            for widget in widgets:
                try:
                    widget.configure(state=state)
                except Exception:
                    pass
            names = (
                ("column_entry", "header_chk", "template_btn")
                if mode == "generate" else ("fix_column_entry", "fix_header_chk")
            )
            for name in names:
                widget = getattr(self, name, None)
                if widget is not None:
                    try:
                        widget.configure(state=state)
                    except Exception:
                        pass
            combo = getattr(self, "_sheet_combos", {}).get(mode)
            if combo is not None:
                try:
                    has_values = bool(combo.cget("values"))
                    combo.configure(state="readonly" if excel and has_values else "disabled")
                except Exception:
                    pass

        configure_mode("generate", self.generate_script_source_var.get() == "excel")
        configure_mode("fix", self.fix_script_source_var.get() == "excel")
        for name in ("script_source_segment", "fix_script_source_segment"):
            segment = getattr(self, name, None)
            if segment is not None:
                segment.refresh()
        self._update_action_bar()

    def _create_script_template(self):
        """Crea il modello Excel in background per non bloccare file grandi."""
        if self._template_running or self._is_busy:
            return
        pptx_path = self.generate_input_pptx_var.get().strip()
        if not pptx_path or not Path(pptx_path).is_file():
            messagebox.showerror("File mancante", "Seleziona prima il PowerPoint.")
            return
        output = filedialog.asksaveasfilename(
            title="Salva modello degli script", defaultextension=".xlsx",
            initialfile=f"{Path(pptx_path).stem}_script.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not output:
            return
        self._template_running = True
        self.template_btn.configure(state="disabled", text="Creazione…")

        def work():
            try:
                slide_narrator.create_script_template(pptx_path, output)
                self._post_ui(self._template_created, output)
            except Exception as exc:
                self._post_ui(self._template_failed, str(exc))

        threading.Thread(target=work, name="pptx-template", daemon=True).start()

    def _template_created(self, output: str):
        self._template_running = False
        self.template_btn.configure(state="normal", text="Crea modello Excel")
        self.generate_input_xlsx_var.set(output)
        self.script_source_var.set("excel")
        self.generate_sheet_name_var.set("Script")
        self.generate_script_column_var.set("E")
        self.generate_has_header_var.set(True)
        self._on_script_source_change()
        messagebox.showinfo("Modello creato", f"Modello Excel creato:\n{output}")

    def _template_failed(self, error: str):
        self._template_running = False
        self.template_btn.configure(state="normal", text="Crea modello Excel")
        messagebox.showerror("Errore", error)

    def _on_clear_cache(self):
        """Svuota la cache fuori dal thread Tk."""
        if self._cache_clear_running or not legacy._CLONE_AVAILABLE:
            return
        self._cache_clear_running = True
        if hasattr(self, "clear_cache_btn"):
            self.clear_cache_btn.configure(state="disabled", text="Pulizia…")

        def work():
            try:
                count = legacy.voice_clone.clear_cache()
                self._post_ui(self._cache_cleared, int(count or 0), None)
            except Exception as exc:
                self._post_ui(self._cache_cleared, 0, str(exc))

        threading.Thread(target=work, name="pptx-cache-cleanup", daemon=True).start()

    def _cache_cleared(self, count: int, error: str | None):
        self._cache_clear_running = False
        if hasattr(self, "clear_cache_btn"):
            self.clear_cache_btn.configure(state="normal", text="Svuota cache")
        if error:
            messagebox.showerror("Errore", f"Impossibile svuotare la cache: {error}")
        else:
            messagebox.showinfo(
                "Cache svuotata",
                f"Rimossi {count} file audio dalla cache." if count else "La cache era già vuota.",
            )

    def _on_analyze_fix(self):
        """Avvia l'analisi senza bloccare il mainloop di Tk.

        Il lavoro su ZIP, XML e tracce audio può durare parecchio su file con
        centinaia di slide. Viene quindi eseguito in un worker e tutte le
        modifiche ai widget vengono riportate al thread principale con
        ``after``. L'avvio è rinviato ad ``after_idle`` così il normale stato
        grafico "pressed" del pulsante viene rilasciato prima di disabilitarlo.
        """
        if self._is_busy:
            return
        path = self.fix_input_pptx_var.get().strip()
        if not path or not Path(path).is_file():
            messagebox.showerror("File mancante", "Seleziona un PowerPoint valido da controllare.")
            return
        # Lascia terminare il ciclo mouse-up del pulsante: evita rendering
        # sovrapposto/ritagliato sui temi Windows ad alto DPI.
        self.root.after_idle(lambda p=path: self._start_fix_analysis(p))

    def _start_fix_analysis(self, path: str):
        if self._is_busy or path != self.fix_input_pptx_var.get().strip():
            return
        self._is_busy = True
        self._current_operation = "fix_analysis"
        self._cancel_event.clear()
        self._fix_analysis_token += 1
        token = self._fix_analysis_token
        self._active_fix_analysis_path = path
        self._last_fix_analysis = None
        self._last_analyzed_fix_path = None
        self._all_fix_rows.clear()
        self.fix_analysis_var.set("Preparazione dell'analisi…")
        try:
            self.analyze_fix_btn.state(["disabled", "!pressed", "!active"])
        except Exception:
            self.analyze_fix_btn.configure(state="disabled")
        self._set_progress_determinate(0, "Preparazione analisi PowerPoint…")
        self.cancel_btn.configure(state="normal")
        self._update_action_bar()
        # _update_action_bar può ridisegnare il footer: riabilito Annulla dopo.
        self.cancel_btn.configure(state="normal")

        def progress(event):
            try:
                self._post_ui(self._fix_analysis_progress, token, dict(event))
            except Exception:
                pass

        def work():
            try:
                result = slide_narrator.analyze_pptx_audio(
                    path, progress_callback=progress, cancel_event=self._cancel_event
                )
                if self._cancel_event.is_set():
                    raise slide_narrator.OperationCancelled("Analisi annullata")
                metadata = {
                    m["slide_num"]: m
                    for m in slide_narrator.get_pptx_slide_metadata(path)
                }
                self._post_ui(self._finish_fix_analysis, token, path, result, metadata)
            except slide_narrator.OperationCancelled:
                self._post_ui(self._cancelled_fix_analysis, token)
            except Exception as exc:
                error = str(exc)
                self._post_ui(self._failed_fix_analysis, token, error)

        self._worker_thread = threading.Thread(
            target=work, name="pptx-audio-analysis", daemon=True
        )
        self._worker_thread.start()

    def _fix_analysis_progress(self, token: int, event: dict):
        if token != self._fix_analysis_token or self._current_operation != "fix_analysis":
            return
        current = int(event.get("current", 0) or 0)
        total = int(event.get("total", 0) or 0)
        percent = (current / total * 100.0) if total else 0.0
        label = event.get("message") or (
            f"Analisi slide {current} di {total}" if total else "Analisi in corso…"
        )
        self.fix_analysis_var.set(label)
        self._set_progress_determinate(percent, label)

    def _finish_fix_analysis(self, token: int, path: str, result: dict, metadata: dict):
        if token != self._fix_analysis_token:
            return
        # Se durante il lavoro è stato scelto un altro file, il risultato non
        # deve mai contaminare la nuova selezione.
        if path != self.fix_input_pptx_var.get().strip():
            self._complete_fix_analysis_ui("Analisi ignorata: il file selezionato è cambiato.")
            return
        self._last_fix_analysis = result
        self._last_analyzed_fix_path = path
        self._all_fix_rows.clear()
        for slide_num in range(1, result["total_slides"] + 1):
            detail = result["details"].get(slide_num, {})
            status = detail.get("status", "missing")
            status_label = {
                "valid": "Valido", "missing": "Mancante", "invalid": "Da verificare"
            }.get(status, status)
            duration = detail.get("duration_s")
            duration_text = slide_narrator.format_duration(duration) if duration is not None else "—"
            fmt = str(detail.get("format") or "—").upper().lstrip(".")
            values = (
                status_label, fmt, duration_text,
                "Sì" if detail.get("autoplay") else "No",
                "Sì" if detail.get("auto_advance") else "No",
                metadata.get(slide_num, {}).get("title", ""),
            )
            self._all_fix_rows[str(slide_num)] = (values, status)
        self.fix_total_var.set(str(result["total_slides"]))
        self.fix_valid_var.set(str(len(result["valid"])))
        self.fix_missing_var.set(str(len(result["missing"])))
        self.fix_invalid_var.set(str(len(result["invalid"])))
        message = (
            f"Analisi completata: {len(result['missing'])} audio mancanti e "
            f"{len(result['invalid'])} anomalie da valutare."
        )
        self.fix_analysis_var.set(message)
        self._apply_fix_filter()
        self._fix_max_unlocked_step = max(self._fix_max_unlocked_step, 3)
        self._refresh_fix_summary()
        self._complete_fix_analysis_ui("Analisi completata")

    def _cancelled_fix_analysis(self, token: int):
        if token != self._fix_analysis_token:
            return
        self._last_fix_analysis = None
        self._last_analyzed_fix_path = None
        self.fix_analysis_var.set("Analisi annullata. Puoi avviarla nuovamente.")
        self._complete_fix_analysis_ui("Analisi annullata")

    def _failed_fix_analysis(self, token: int, error: str):
        if token != self._fix_analysis_token:
            return
        self._last_fix_analysis = None
        self._last_analyzed_fix_path = None
        self.fix_analysis_var.set("Analisi non riuscita.")
        self._complete_fix_analysis_ui("Errore durante l'analisi")
        messagebox.showerror("Errore di analisi", error)

    def _complete_fix_analysis_ui(self, footer_text: str):
        self._is_busy = False
        self._current_operation = None
        self._active_fix_analysis_path = None
        try:
            self.analyze_fix_btn.state(["!disabled", "!pressed", "!active"])
        except Exception:
            self.analyze_fix_btn.configure(state="normal")
        self._set_progress_idle(footer_text)
        self._update_action_bar()

    def _apply_fix_filter(self):
        selected = set(self.fix_tree.selection()) if hasattr(self, "fix_tree") else set()
        for item in self.fix_tree.get_children():
            self.fix_tree.delete(item)
        wanted = self.fix_filter_var.get()
        for slide, (values, status) in sorted(self._all_fix_rows.items(), key=lambda x: int(x[0])):
            if wanted != "all" and status != wanted:
                continue
            self.fix_tree.insert("", "end", iid=slide, text=slide, values=values, tags=(status,))
        dark = self._theme_mode == "dark" or (self._theme_mode == "system" and _system_prefers_dark())
        self.fix_tree.tag_configure("invalid", background="#583537" if dark else "#FBE7E7")
        self.fix_tree.tag_configure("missing", background="#564720" if dark else "#FFF3CF")
        for item in selected:
            if self.fix_tree.exists(item): self.fix_tree.selection_add(item)
        if hasattr(self, "fix_empty_label"):
            if self.fix_tree.get_children(): self.fix_empty_label.grid_remove()
            else:
                self.fix_empty_label.configure(text="Nessuna slide corrisponde al filtro selezionato.")
                self.fix_empty_label.grid(row=0, column=0)
        if hasattr(self, "fix_filter_segment"): self.fix_filter_segment.refresh()

    def _on_mousewheel(self, event):
        # Il log e le tabelle mantengono il proprio scroll.
        try:
            widget = self.root.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return
        node = widget
        while node is not None:
            if node is getattr(self, "log_text", None) or isinstance(node, ttk.Treeview):
                return
            node = getattr(node, "master", None)
        page = self.pages.get(self.current_page)
        if not isinstance(page, ScrollablePage):
            return
        if getattr(event, "num", None) == 4:
            delta = -3
        elif getattr(event, "num", None) == 5:
            delta = 3
        else:
            delta = int(-event.delta / 120) * 3
        page.scroll(delta)
        return "break"

    def _sheet_context(self, mode: str):
        if mode == "fix":
            return (
                self.fix_input_xlsx_var, self.fix_sheet_name_var,
                getattr(self, "fix_sheet_combo", None),
                getattr(self, "fix_sheet_status", None),
            )
        return (
            self.generate_input_xlsx_var, self.generate_sheet_name_var,
            getattr(self, "generate_sheet_combo", None),
            getattr(self, "generate_sheet_status", None),
        )

    def _schedule_sheet_names_load(self, mode: str, immediate: bool = False):
        """Carica in modo asincrono i nomi dei fogli del file Excel selezionato."""
        mode = "fix" if mode == "fix" else "generate"
        previous = self._sheet_scan_after_ids.get(mode)
        if previous is not None:
            try:
                self.root.after_cancel(previous)
            except Exception:
                pass
        delay = 0 if immediate else 250
        try:
            self._sheet_scan_after_ids[mode] = self.root.after(
                delay, lambda m=mode: self._start_sheet_names_load(m)
            )
        except Exception:
            self._sheet_scan_after_ids[mode] = None

    def _cancel_sheet_timeout(self, mode: str) -> None:
        timeout_id = self._sheet_scan_timeout_ids.get(mode)
        if timeout_id is not None:
            try:
                self.root.after_cancel(timeout_id)
            except Exception:
                pass
            self._sheet_scan_timeout_ids[mode] = None

    def _start_sheet_names_load(self, mode: str):
        self._sheet_scan_after_ids[mode] = None
        self._cancel_sheet_timeout(mode)
        path_var, _sheet_var, combo, status = self._sheet_context(mode)
        raw = path_var.get().strip()
        path = Path(raw) if raw else None
        if not path or not path.is_file() or path.suffix.lower() not in {".xlsx", ".xlsm"}:
            self._sheet_scan_tokens[mode] += 1
            self._apply_sheet_names(mode, self._sheet_scan_tokens[mode], raw, [], None)
            return
        self._sheet_scan_tokens[mode] += 1
        token = self._sheet_scan_tokens[mode]
        if combo is not None:
            try:
                combo.configure(state="disabled")
            except Exception:
                pass
        if status is not None:
            status.configure(text="Lettura dei fogli Excel in corso…", style="MutedCard.TLabel")

        # Evita uno stato di caricamento infinito quando il file cloud non è
        # disponibile offline, è bloccato o il filesystem non risponde.
        try:
            self._sheet_scan_timeout_ids[mode] = self.root.after(
                20000,
                lambda m=mode, t=token, p=raw: self._sheet_names_timeout(m, t, p),
            )
        except Exception:
            self._sheet_scan_timeout_ids[mode] = None

        def work():
            try:
                names = _read_excel_sheet_names_fast(path)
                # Passiamo il valore originale della StringVar. Su Windows
                # ``str(Path('C:/...'))`` può trasformarlo in ``C:\\...``.
                self._post_ui(self._apply_sheet_names, mode, token, raw, names, None)
            except Exception as exc:
                self._post_ui(self._apply_sheet_names, mode, token, raw, [], str(exc))

        threading.Thread(
            target=work, name=f"excel-sheets-{mode}", daemon=True
        ).start()

    def _sheet_names_timeout(self, mode: str, token: int, path: str) -> None:
        self._sheet_scan_timeout_ids[mode] = None
        path_var, sheet_var, combo, status = self._sheet_context(mode)
        if token != self._sheet_scan_tokens.get(mode):
            return
        if _path_identity(path_var.get()) != _path_identity(path):
            return
        # Invalida il worker eventualmente ancora bloccato: un suo risultato
        # tardivo non potrà più sovrascrivere la GUI.
        self._sheet_scan_tokens[mode] += 1
        if combo is not None:
            combo.configure(values=(), state="disabled")
        sheet_var.set("")
        if status is not None:
            status.configure(
                text=("Lettura non completata. Se il file è su Dropbox/OneDrive, "
                      "rendilo disponibile offline e selezionalo di nuovo."),
                style="Error.TLabel",
            )
        self._log(f"\n[Excel] Timeout durante la lettura dei fogli di {path}\n")
        self._update_action_bar()

    def _apply_sheet_names(
        self, mode: str, token: int, path: str, names: list[str], error: str | None
    ):
        path_var, sheet_var, combo, status = self._sheet_context(mode)
        if token != self._sheet_scan_tokens.get(mode):
            return
        if _path_identity(path_var.get()) != _path_identity(path):
            return
        self._cancel_sheet_timeout(mode)
        if combo is None:
            return
        if error:
            combo.configure(values=(), state="disabled")
            sheet_var.set("")
            if status is not None:
                status.configure(
                    text="Impossibile leggere i fogli del file Excel.", style="Error.TLabel"
                )
            self._log(f"\n[Excel] Impossibile leggere i fogli di {path}: {error}\n")
            return
        combo.configure(values=tuple(names))
        if names:
            if sheet_var.get().strip() not in names:
                sheet_var.set(names[0])
            source_var = (
                self.generate_script_source_var if mode == "generate"
                else self.fix_script_source_var
            )
            combo.configure(
                state="readonly" if source_var.get() == "excel" else "disabled"
            )
            if status is not None:
                count = len(names)
                label = "1 foglio disponibile" if count == 1 else f"{count} fogli disponibili"
                status.configure(
                    text=f"{label} · seleziona quello che contiene gli script.",
                    style="Success.TLabel",
                )
        else:
            sheet_var.set("")
            combo.configure(state="disabled")
            if status is not None:
                status.configure(
                    text="Carica un file Excel per visualizzare i fogli disponibili.",
                    style="MutedCard.TLabel",
                )
        self._update_action_bar()

    def _copy_ai_prompt(self):
        """Copia negli appunti il prompt standard per generare l'Excel TTS."""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(AI_SCRIPT_PROMPT)
            self.root.update_idletasks()
        except tk.TclError as exc:
            messagebox.showerror("Appunti non disponibili", str(exc))
            return
        for button in getattr(self, "_ai_prompt_buttons", []):
            try:
                button.configure(text="Prompt copiato ✓")
            except Exception:
                pass
        if self._ai_prompt_reset_after_id is not None:
            try:
                self.root.after_cancel(self._ai_prompt_reset_after_id)
            except Exception:
                pass
        try:
            self._ai_prompt_reset_after_id = self.root.after(
                1800, self._reset_ai_prompt_buttons
            )
        except Exception:
            self._ai_prompt_reset_after_id = None

    def _reset_ai_prompt_buttons(self):
        self._ai_prompt_reset_after_id = None
        for button in getattr(self, "_ai_prompt_buttons", []):
            try:
                button.configure(text="Copia prompt AI")
            except Exception:
                pass

    def _browse_pptx(self):
        legacy.App._browse_pptx(self)
        self._update_file_summaries()

    def _browse_xlsx(self):
        legacy.App._browse_xlsx(self)
        mode = "fix" if self.operation_mode_var.get() == "fix" else "generate"
        self._schedule_sheet_names_load(mode, immediate=True)
        self._update_file_summaries()

    def _browse_output(self):
        legacy.App._browse_output(self)
        self._update_file_summaries()

    def _generate_cancelled(self):
        self._set_progress_idle("Annullato")
        self.cancel_btn.configure(state="disabled")
        self._set_preview_enabled(True)
        self._is_busy = False
        self._log("\nOperazione annullata dall'utente.\n")
        self._update_action_bar()
        messagebox.showinfo("Annullato", "L'operazione è stata annullata. Il file originale non è stato modificato.")

    def _generate_done(self, out_path):
        self._set_progress_determinate(100, "Completato")
        self.cancel_btn.configure(state="disabled")
        self._set_preview_enabled(True)
        self._is_busy = False
        out = Path(out_path)
        is_fix = self.operation_mode_var.get() == "fix"
        is_video = not is_fix and self.output_mode_var.get() == "video"
        if is_fix:
            title = "PowerPoint riparato"
        elif is_video:
            title = "Video generato"
        else:
            title = "PowerPoint generato"
        report = out.with_name(f"{out.stem}_esecuzione.json")
        details = "Il file è stato verificato ed è pronto per l'utilizzo."
        try:
            data = json.loads(report.read_text(encoding="utf-8"))
            elapsed = data.get("elapsed_seconds") or data.get("duration_seconds")
            status = data.get("status", "completed")
            extra = []
            if elapsed:
                extra.append(f"Durata elaborazione: {slide_narrator.format_duration(float(elapsed))}")
            extra.append(f"Stato: {status}")
            details = " · ".join(extra)
        except Exception:
            pass
        self.result_title_var.set(title)
        self.result_path_var.set(str(out))
        self.result_details_var.set(details)
        self._show_page("result")

    def _generate_failed(self, err):
        self._set_progress_idle("Errore")
        self.cancel_btn.configure(state="disabled")
        self._set_preview_enabled(True)
        self._is_busy = False
        self._log(f"\n⚠ ERRORE: {err}\n")
        self._update_action_bar()
        if self.show_log_on_error_var.get():
            self._set_log_visible(True)
        action = "Fix" if self.operation_mode_var.get() == "fix" else "Generazione"
        messagebox.showerror("Errore", f"{action} fallito:\n{err}")

    def _on_close(self):
        self._save_preferences()
        self._on_root_destroy()
        try:
            if self._ui_poller_id is not None:
                self.root.after_cancel(self._ui_poller_id)
                self._ui_poller_id = None
        except Exception:
            pass
        legacy.App._on_close(self)

    def _refresh_fix_summary(self):
        if not hasattr(self, "fix_summary_vars"):
            return
        inp = self.fix_input_pptx_var.get().strip()
        out = self.fix_output_pptx_var.get().strip()
        result = self._last_fix_analysis or {}
        self.fix_summary_vars["pptx"].set(Path(inp).name if inp else "Non selezionato")
        if result:
            self.fix_summary_vars["analysis"].set(
                f"{result.get('total_slides', 0)} slide · "
                f"{len(result.get('valid', []))} audio validi · "
                f"{len(result.get('missing', []))} mancanti · "
                f"{len(result.get('invalid', {}))} da verificare"
            )
        else:
            self.fix_summary_vars["analysis"].set("Analisi non ancora eseguita")

        if self.fix_script_source_var.get() == "notes":
            scripts = "Note PowerPoint"
        else:
            xlsx = self.fix_input_xlsx_var.get().strip()
            sheet = self.fix_sheet_name_var.get().strip()
            column = self.fix_script_column_var.get().strip().upper() or "A"
            scripts = Path(xlsx).name if xlsx else "Excel non selezionato"
            if sheet:
                scripts += f" · foglio {sheet} · colonna {column}"
        self.fix_summary_vars["scripts"].set(scripts)
        self.fix_summary_vars["voice"].set(
            self._voice_id_to_display.get(self.fix_voice_var.get(), self.fix_voice_var.get())
            or "Non selezionata"
        )
        self.fix_summary_vars["output"].set(out or "Percorso non definito")

        actions = ["Completa gli audio mancanti"]
        if self.fix_regenerate_all_var.get():
            actions.append("rigenera tutte le slide")
        if self.fix_repair_invalid_var.get():
            actions.append("ripara audio anomali")
        if self.fix_normalize_autoplay_var.get():
            actions.append("uniforma autoplay")
        if self.fix_normalize_advance_var.get():
            actions.append("uniforma avanzamento")
        if self.fix_normalize_icon_var.get():
            actions.append("nasconde le icone audio")
        self.fix_summary_vars["actions"].set(" · ".join(actions))

        errors = self._validate_fix_project()
        self.fix_summary_vars["validation"].set(
            "Fix pronto" if not errors else "Da completare: " + " · ".join(errors)
        )

    # ------------------------------------------------------------ summary/ui
    def _set_progress_indeterminate(self, label_text: str = ""):
        self._show_footer_progress()
        try: self.progress.stop()
        except Exception: pass
        self.progress.configure(mode="indeterminate"); self.progress.start(10); self.progress_label.configure(text=label_text)

    def _set_progress_determinate(self, value: float, label_text: str):
        self._show_footer_progress()
        try: self.progress.stop()
        except Exception: pass
        self.progress.configure(mode="determinate", value=max(0, min(100, value))); self.progress_label.configure(text=label_text)

    def _set_progress_idle(self, label_text: str = "Pronto"):
        try:
            self.progress.stop(); self.progress.configure(mode="determinate", value=0); self.progress_label.configure(text=label_text)
        except Exception: pass
        if not getattr(self, "_is_busy", False): self._show_footer_idle()

    def _update_file_summaries(self):
        if self.wizard_step.get() == 4:
            self._refresh_summary()

    def _refresh_summary(self):
        pptx_path = self.generate_input_pptx_var.get().strip()
        script_source = self.generate_script_source_var.get()
        script = "Note PowerPoint" if script_source == "notes" else self.generate_input_xlsx_var.get().strip()
        self.summary_vars["pptx"].set(Path(pptx_path).name if pptx_path else "Non selezionato")
        self.summary_vars["scripts"].set(Path(script).name if script_source == "excel" and script else script)
        self.summary_vars["voice"].set(self._voice_id_to_display.get(self.voice_var.get(), self.voice_var.get()))
        output_label = "Video MP4" if self.output_mode_var.get() == "video" else "PowerPoint con audio"
        out = self.generate_output_pptx_var.get().strip()
        self.summary_vars["output"].set(f"{output_label} — {out or 'percorso non definito'}")
        if self.output_mode_var.get() == "video":
            options = f"{self.resolution_var.get()}, backend {self.render_backend_var.get()}"
            if self.subtitles_var.get(): options += ", sottotitoli"
            if self.transition_var.get(): options += f", transizione {self.transition_style_var.get()}"
        else:
            options = "Autoplay" if self.autoplay_var.get() else "Avvio manuale"
            if self.auto_advance_var.get(): options += ", avanzamento automatico"
        self.summary_vars["options"].set(options)
        validation = self._validate_generate_project()
        self.summary_vars["validation"].set(
            "Progetto completo" if not validation else "Da completare: " + " · ".join(validation)
        )
        if validation:
            self.summary_vars["preflight"].set("Il controllo tecnico partirà quando il progetto sarà completo.")
        else:
            key = (
                self.output_mode_var.get(), out or None, pptx_path or None,
                self.render_backend_var.get(),
            )
            if key == self._summary_preflight_key and self._summary_preflight_result is not None:
                self._apply_summary_preflight_result(self._summary_preflight_result, None)
            else:
                self._summary_preflight_key = key
                self._summary_preflight_result = None
                self._summary_preflight_token += 1
                token = self._summary_preflight_token
                self.summary_vars["preflight"].set("Controllo tecnico in corso…")

                def work():
                    try:
                        result = slide_narrator.preflight_system(
                            key[0], key[1], input_pptx=key[2], render_backend=key[3]
                        )
                        self._post_ui(self._summary_preflight_done, token, key, result, None)
                    except Exception as exc:
                        self._post_ui(self._summary_preflight_done, token, key, None, str(exc))

                threading.Thread(target=work, name="pptx-summary-preflight", daemon=True).start()
        self._update_action_bar()

    def _summary_preflight_done(self, token: int, key: tuple, result: dict | None, error: str | None):
        if token != self._summary_preflight_token or key != self._summary_preflight_key:
            return
        self._summary_preflight_result = result if error is None else None
        self._apply_summary_preflight_result(result, error)

    def _apply_summary_preflight_result(self, result: dict | None, error: str | None):
        if error:
            self.summary_vars["preflight"].set(f"Controllo non disponibile: {error}")
            return
        result = result or {}
        if result.get("ok") and not result.get("warnings"):
            text = "Ambiente pronto"
        elif result.get("ok"):
            text = "Avvisi: " + " · ".join(result.get("warnings", []))
        else:
            text = "Errori: " + " · ".join(result.get("errors", []))
        self.summary_vars["preflight"].set(text)

    def _toggle_log(self):
        self._set_log_visible(not self._log_visible)

    def _set_log_visible(self, visible: bool):
        self._log_visible = bool(visible)
        if visible:
            self.log_drawer.grid()
            self.log_toggle_btn.configure(text="Nascondi dettagli")
        else:
            self.log_drawer.grid_remove()
            self.log_toggle_btn.configure(text="Dettagli tecnici")

    def _toggle_theme_quick(self):
        current_dark = self._theme_mode == "dark" or (
            self._theme_mode == "system" and _system_prefers_dark()
        )
        self.theme_mode_var.set("light" if current_dark else "dark")
        self._on_theme_mode_change()

    def _on_theme_mode_change(self):
        self._theme_mode = self.theme_mode_var.get()
        if _BOOTSTRAP_AVAILABLE:
            try:
                self.style.theme_use(_theme_for_mode(self._theme_mode))
            except Exception as exc:
                messagebox.showerror("Tema", f"Impossibile cambiare tema: {exc}")
                return
        self._configure_custom_styles()
        if hasattr(self, "_brand_canvas"):
            self._brand_canvas.configure(bg=self._colors["sidebar"])
            self._draw_brand_badge()
        if hasattr(self, "_nav_rows"):
            self._refresh_sidebar_visuals()
        for page in self.pages.values():
            if isinstance(page, ScrollablePage):
                page.refresh_colors()
        self._save_preferences()

    def _save_preferences(self):
        try:
            SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
            data = {
                "theme_mode": self.theme_mode_var.get() if hasattr(self, "theme_mode_var") else self._theme_mode,
                "history_folder": self.history_folder_var.get() if hasattr(self, "history_folder_var") else str(Path.cwd()),
                "start_page": "generate",
                "show_log_on_error": bool(self.show_log_on_error_var.get()) if hasattr(self, "show_log_on_error_var") else True,
            }
            SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    # -------------------------------------------------------------- batch UI
    def _browse_batch_manifest(self):
        path = filedialog.askopenfilename(
            title="Seleziona manifest batch", filetypes=[("JSON", "*.json"), ("Tutti", "*.*")]
        )
        if path:
            self.batch_manifest_var.set(path)
            self._load_batch_manifest(path)

    def _load_batch_manifest(self, path: str):
        self._batch_jobs = []
        if hasattr(self, "batch_tree"):
            for item in self.batch_tree.get_children():
                self.batch_tree.delete(item)
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            jobs = data.get("jobs", []) if isinstance(data, dict) else []
            if not isinstance(jobs, list):
                raise ValueError("Il campo jobs deve essere un elenco.")
            for i, job in enumerate(jobs, 1):
                if not isinstance(job, dict):
                    continue
                self._batch_jobs.append(job)
                if hasattr(self, "batch_tree"):
                    self.batch_tree.insert("", "end", iid=f"job{i}", values=(
                        job.get("mode", "generate"),
                        Path(str(job.get("input_pptx", ""))).name,
                        Path(str(job.get("scripts_xlsx", job.get("input_xlsx", "")))).name,
                        Path(str(job.get("output_pptx", job.get("output", "")))).name,
                    ))
            self.batch_jobs_var.set(f"{len(self._batch_jobs)} lavori caricati")
            self.batch_result_var.set(f"Manifest pronto: {Path(path).name}")
        except Exception as exc:
            self._batch_jobs = []
            self.batch_jobs_var.set("Manifest non valido")
            self.batch_result_var.set(str(exc))
            messagebox.showerror("Manifest batch", str(exc))
        self._update_action_bar()

    def _update_voice_actions(self):
        selected = bool(getattr(self, "library_tree", None) and self.library_tree.selection())
        for name in ("library_use_btn", "library_preview_btn"):
            btn = getattr(self, name, None)
            if btn is not None:
                btn.configure(state="normal" if selected else "disabled")

    def _preview_library_voice(self):
        selection = self.library_tree.selection() if hasattr(self, "library_tree") else ()
        if not selection:
            return
        self.voice_var.set(selection[0])
        self._on_preview()

    def _update_history_actions(self):
        item = self._selected_history() if hasattr(self, "history_tree") else None
        if hasattr(self, "history_report_btn"):
            self.history_report_btn.configure(state="normal" if item else "disabled")
        output_ok = bool(item and item.get("output") and Path(item["output"]).exists())
        if hasattr(self, "history_output_btn"):
            self.history_output_btn.configure(state="normal" if output_ok else "disabled")

    @staticmethod
    def _find_libreoffice() -> str | None:
        for command in ("soffice", "libreoffice"):
            found = shutil.which(command)
            if found:
                return found
        if platform.system() == "Windows":
            candidates = [
                Path(os.environ.get("PROGRAMFILES", "")) / "LibreOffice" / "program" / "soffice.exe",
                Path(os.environ.get("PROGRAMFILES(X86)", "")) / "LibreOffice" / "program" / "soffice.exe",
            ]
            for candidate in candidates:
                if candidate.is_file():
                    return str(candidate)
        return None

    @staticmethod
    def _has_powerpoint() -> bool:
        if platform.system() != "Windows":
            return False
        try:
            import winreg
            for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                for key_path in (
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE",
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE",
                ):
                    try:
                        with winreg.OpenKey(hive, key_path):
                            return True
                    except OSError:
                        pass
        except Exception:
            pass
        return False

    def _refresh_tools_status(self):
        if not hasattr(self, "tool_status_vars"):
            return
        winget = shutil.which("winget") if platform.system() == "Windows" else None
        ffmpeg = shutil.which("ffmpeg")
        libreoffice = self._find_libreoffice()
        powerpoint = self._has_powerpoint()
        self.tool_status_vars["winget"].set("Disponibile" if winget else "Non disponibile")
        self.tool_status_vars["ffmpeg"].set(
            f"Installato — {ffmpeg}" if ffmpeg else "Non installato"
        )
        self.tool_status_vars["libreoffice"].set(
            f"Installato — {libreoffice}" if libreoffice else "Non installato"
        )
        self.tool_status_vars["powerpoint"].set(
            "Installato" if powerpoint else "Non rilevato"
        )
        if platform.system() != "Windows":
            note = "L’installazione automatica con WinGet è disponibile soltanto su Windows."
        elif not winget:
            note = "WinGet non è disponibile. Installa o aggiorna App Installer dal Microsoft Store."
        elif getattr(sys, "frozen", False):
            note = (
                "FFmpeg e LibreOffice possono essere installati qui. I motori vocali locali richiedono "
                "una nuova build dell’app e non possono essere aggiunti a questo EXE."
            )
        else:
            note = (
                "FFmpeg e LibreOffice possono essere installati qui. Per PocketTTS, Chatterbox o XTTS "
                "usa Installa_Slide_Narrator.bat e ricrea poi l’EXE."
            )
        self.tools_note_var.set(note)
        for key in ("ffmpeg", "libreoffice"):
            button = self.tool_install_buttons.get(key)
            if button is not None:
                installed = bool(ffmpeg) if key == "ffmpeg" else bool(libreoffice)
                button.configure(
                    state="disabled" if installed or not winget else "normal",
                    text="Installato" if installed else (
                        "Installa FFmpeg" if key == "ffmpeg" else "Installa LibreOffice"
                    ),
                )

    def _install_windows_tool(self, key: str):
        if platform.system() != "Windows":
            messagebox.showinfo("Installazione non disponibile", "Questa funzione usa WinGet ed è disponibile su Windows.")
            return
        winget = shutil.which("winget")
        if not winget:
            messagebox.showerror(
                "WinGet non disponibile",
                "Installa o aggiorna ‘App Installer’ dal Microsoft Store, quindi riapri Slide Narrator.",
            )
            return
        packages = {
            "ffmpeg": ("Gyan.FFmpeg", "FFmpeg"),
            "libreoffice": ("TheDocumentFoundation.LibreOffice", "LibreOffice"),
        }
        if key not in packages:
            return
        package_id, display_name = packages[key]
        if not messagebox.askyesno(
            f"Installare {display_name}?",
            f"Slide Narrator userà WinGet per scaricare e installare {display_name}.\n\nContinuare?",
        ):
            return
        button = self.tool_install_buttons.get(key)
        if button is not None:
            button.configure(state="disabled", text="Installazione…")
        self.tool_status_vars[key].set("Download e installazione in corso…")

        def work():
            command = [
                winget, "install", "--id", package_id, "-e", "--source", "winget",
                "--accept-package-agreements", "--accept-source-agreements",
            ]
            try:
                proc = subprocess.run(command, capture_output=True, text=True, timeout=1800)
                output = (proc.stdout or proc.stderr or "").strip()
                self._post_ui(self._tool_install_done, key, display_name, proc.returncode, output)
            except Exception as exc:
                self._post_ui(self._tool_install_done, key, display_name, -1, str(exc))

        threading.Thread(target=work, name=f"install-{key}", daemon=True).start()

    def _tool_install_done(self, key: str, display_name: str, returncode: int, output: str):
        self._refresh_tools_status()
        if returncode == 0:
            messagebox.showinfo(
                "Installazione completata",
                f"{display_name} è stato installato. Se non viene rilevato subito, riavvia Slide Narrator.",
            )
        else:
            detail = output[-1200:] if output else "Nessun dettaglio disponibile."
            messagebox.showerror(
                "Installazione non riuscita",
                f"WinGet non ha completato l’installazione di {display_name}.\n\n{detail}",
            )

    @staticmethod
    def _open_microsoft_365():
        webbrowser.open("https://www.microsoft.com/microsoft-365/powerpoint")

    def _run_diagnostics(self):
        """Controlla strumenti e backend senza bloccare la finestra."""
        self._diagnostics_token += 1
        token = self._diagnostics_token
        self.diagnostics_var.set("Diagnostica in corso…")
        if hasattr(self, "diagnostics_btn"):
            self.diagnostics_btn.configure(state="disabled", text="Controllo…")

        def work():
            try:
                result = slide_narrator.preflight_system("video", None, render_backend="auto")
                self._post_ui(self._diagnostics_done, token, result, None)
            except Exception as exc:
                self._post_ui(self._diagnostics_done, token, None, str(exc))

        threading.Thread(target=work, name="pptx-diagnostics", daemon=True).start()

    def _diagnostics_done(self, token: int, result: dict | None, error: str | None):
        if token != self._diagnostics_token:
            return
        if hasattr(self, "diagnostics_btn"):
            self.diagnostics_btn.configure(state="normal", text="Esegui diagnostica")
        if error:
            self.diagnostics_var.set(f"Diagnostica non riuscita: {error}")
            return
        result = result or {}
        parts = []
        if result.get("ok"):
            parts.append("Ambiente principale pronto.")
        if result.get("warnings"):
            parts.append("Avvisi: " + " · ".join(result["warnings"]))
        if result.get("errors"):
            parts.append("Errori: " + " · ".join(result["errors"]))
        self.diagnostics_var.set(" ".join(parts) or "Controllo completato.")

    def _on_run_batch(self):
        if self._is_busy:
            return
        manifest = self.batch_manifest_var.get().strip()
        if not manifest or not Path(manifest).exists():
            messagebox.showerror("Manifest mancante", "Seleziona un file JSON valido.")
            return
        self._is_busy = True
        self.generate_btn.configure(state="disabled", text="Batch in corso…")
        self.cancel_btn.configure(state="disabled")
        self._set_progress_indeterminate("Elaborazione batch in corso…")
        self._clear_log()
        self._set_log_visible(True)

        def work():
            redirector = legacy.StdoutRedirector(self.log_text)
            old_stdout, old_stderr = sys.stdout, sys.stderr
            sys.stdout = redirector
            sys.stderr = redirector
            try:
                summary = slide_narrator_batch.run_manifest(
                    manifest, force=self.batch_force_var.get(),
                    stop_on_error=self.batch_stop_on_error_var.get(),
                )
                self._post_ui(self._batch_done, summary)
            except Exception as exc:
                self._post_ui(self._batch_failed, str(exc))
            finally:
                sys.stdout, sys.stderr = old_stdout, old_stderr
        self._worker_thread = threading.Thread(target=work, daemon=True)
        self._worker_thread.start()

    def _batch_done(self, summary: str):
        self._is_busy = False
        self._set_progress_determinate(100, "Batch completato")
        self.batch_result_var.set(f"Riepilogo salvato in: {summary}")
        self.result_title_var.set("Elaborazione batch completata")
        self.result_path_var.set(summary)
        self.result_details_var.set("Il riepilogo JSON contiene l'esito di ogni lavoro.")
        self._show_page("result")

    def _batch_failed(self, error: str):
        self._is_busy = False
        self._set_progress_idle("Errore batch")
        self.batch_result_var.set(error)
        self._update_action_bar()
        messagebox.showerror("Errore batch", error)

    # ------------------------------------------------------------- voice page
    def _refresh_voice_library_page(self):
        if not hasattr(self, "library_tree"):
            return
        selected_id = self.voice_var.get()
        query = self.voice_search_var.get().strip().lower() if hasattr(self, "voice_search_var") else ""
        engine_filter = self.voice_engine_filter_var.get() if hasattr(self, "voice_engine_filter_var") else "Tutti"
        for item in self.library_tree.get_children():
            self.library_tree.delete(item)
        for entry in legacy.get_voice_entries():
            if query and query not in entry["name"].lower() and query not in entry["engine"].lower():
                continue
            if engine_filter != "Tutti" and entry["engine"] != engine_filter:
                continue
            iid = entry["voice_id"]
            self.library_tree.insert(
                "", "end", iid=iid,
                values=(entry["name"], entry["engine"],
                        "Voce locale" if entry["is_clone"] else "Microsoft", "Disponibile"),
            )
        if self.library_tree.exists(selected_id):
            self.library_tree.selection_set(selected_id)
            self.library_tree.see(selected_id)
        self._update_voice_actions()

    def _use_library_voice(self):
        selection = self.library_tree.selection()
        if not selection:
            messagebox.showinfo("Voce", "Seleziona una voce nella tabella.")
            return
        voice_id = selection[0]
        self.voice_var.set(voice_id)
        self.voice_combo_var.set(self._voice_id_to_display.get(voice_id, voice_id))
        self._update_pocket_quality_visibility()
        self._show_page("generate")
        self._show_wizard_step(2, force=True)

    # ------------------------------------------------------------ history UI
    def _browse_history_folder(self):
        folder = filedialog.askdirectory(title="Cartella dei risultati")
        if folder:
            self.history_folder_var.set(folder)
            self._scan_history()

    def _scan_history(self, silent: bool = False):
        """Ricerca i rapporti in un worker; Dropbox e cartelle grandi non congelano Tk."""
        if not hasattr(self, "history_tree"):
            return
        folder = Path(self.history_folder_var.get()).expanduser()
        if not folder.exists():
            if not silent:
                messagebox.showerror("Cartella", "La cartella selezionata non esiste.")
            return
        self._history_scan_token += 1
        token = self._history_scan_token
        if hasattr(self, "history_refresh_btn"):
            self.history_refresh_btn.configure(state="disabled", text="Ricerca…")
        if hasattr(self, "history_empty_label"):
            self.history_empty_label.configure(text="Ricerca delle elaborazioni in corso…")
            self.history_empty_label.grid(row=0, column=0)

        def work():
            rows = []
            error = None
            try:
                reports = sorted(
                    folder.rglob("*_esecuzione.json"),
                    key=lambda p: p.stat().st_mtime, reverse=True,
                )[:300]
                for report in reports:
                    try:
                        data = json.loads(report.read_text(encoding="utf-8"))
                        output = str(data.get("output") or data.get("output_pptx") or "")
                        if not output:
                            candidate = str(report).replace("_esecuzione.json", "")
                            for ext in (".pptx", ".mp4"):
                                if Path(candidate + ext).exists():
                                    output = candidate + ext
                                    break
                        date = datetime.fromtimestamp(report.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
                        rows.append({
                            "date": date,
                            "status": str(data.get("status", "—")),
                            "operation": str(data.get("mode") or data.get("operation") or "generazione"),
                            "display": output or report.stem,
                            "report": str(report),
                            "output": output,
                        })
                    except Exception:
                        continue
            except Exception as exc:
                error = str(exc)
            self._post_ui(self._history_scan_done, token, str(folder), rows, error, silent)

        threading.Thread(target=work, name="pptx-history-scan", daemon=True).start()

    def _history_scan_done(self, token: int, folder: str, rows: list[dict], error: str | None, silent: bool):
        if token != self._history_scan_token:
            return
        if hasattr(self, "history_refresh_btn"):
            self.history_refresh_btn.configure(state="normal", text="Aggiorna")
        if folder != str(Path(self.history_folder_var.get()).expanduser()):
            return
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        self._history_items.clear()
        for index, row in enumerate(rows):
            iid = f"h{index}"
            self.history_tree.insert(
                "", "end", iid=iid,
                values=(row["date"], row["status"], row["operation"], row["display"]),
            )
            self._history_items[iid] = {"report": row["report"], "output": row["output"]}
        if hasattr(self, "history_empty_label"):
            if rows:
                self.history_empty_label.grid_remove()
            else:
                text = "Nessuna elaborazione trovata\nCompleta il primo progetto o seleziona un'altra cartella."
                if error:
                    text = f"Ricerca non riuscita\n{error}"
                self.history_empty_label.configure(text=text)
                self.history_empty_label.grid(row=0, column=0)
        if error and not silent:
            messagebox.showerror("Cronologia", error)
        self._update_history_actions()
        self._save_preferences()

    def _selected_history(self):
        sel = self.history_tree.selection()
        return self._history_items.get(sel[0]) if sel else None

    def _open_selected_history_report(self):
        item = self._selected_history()
        if item:
            self._open_file(item["report"])

    def _open_selected_history_output(self):
        item = self._selected_history()
        if item and item.get("output") and Path(item["output"]).exists():
            self._open_file(item["output"])
        else:
            messagebox.showinfo("Risultato", "Il file di risultato non è stato trovato.")

    def _open_result_folder(self):
        path = self.result_path_var.get().strip()
        if not path:
            return
        folder = str(Path(path).resolve().parent)
        self._open_file(folder)


def main():
    _set_windows_app_user_model_id()
    _enable_windows_dpi_awareness()
    prefs = _load_preferences()
    theme = _theme_for_mode(str(prefs.get("theme_mode", "system")))
    if _BOOTSTRAP_AVAILABLE:
        root = tb.App(
            title=APP_NAME,
            theme=theme,
            light_theme="bootstrap-light",
            dark_theme="bootstrap-dark",
            high_dpi=True,
        )
    else:
        root = tk.Tk()

    # Evita che Windows mostri per un istante e poi memorizzi l'icona Tcl/Tk
    # predefinita prima che Slide Narrator abbia applicato la propria icona.
    root.withdraw()
    app = App(root)
    app._apply_window_identity()
    root.deiconify()
    try:
        root.after(120, app._apply_native_windows_icon)
    except Exception:
        pass
    root.mainloop()


if __name__ == "__main__":
    main()
