"""
Slide Narrator - Genera narrazioni TTS da script XLSX e le inserisce nelle slide PowerPoint

Uso:
    python slide_narrator.py input.pptx scripts.xlsx output.pptx
    python slide_narrator.py input.pptx scripts.xlsx output.pptx --voice it-IT-IsabellaNeural --rate +0%

Requirements:
    pip install edge-tts python-pptx openpyxl lxml mutagen

The XLSX file must have one script per row in column A:
    A1 = script for slide 1
    A2 = script for slide 2
    ...

Output:
    output.pptx           — presentation with embedded audio (autoplay, icon hidden)
    output_durate.txt     — duration report, one line per slide
    output_captions.json  — sentence-level word-boundary timings (per il SCORM Builder)
"""

import argparse
import asyncio
import base64
import copy
import csv
import json
import math
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from io import BytesIO
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import edge_tts
from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string
from pptx import Presentation
from pptx.util import Emu, Inches
from pptx.oxml.ns import qn
from lxml import etree
from mutagen import File as MutagenFile
from mutagen.mp3 import MP3

# Voci clonate (locali). Import opzionale: se i moduli non sono presenti il
# motore continua a funzionare con le sole voci Microsoft (edge-tts).
try:
    import voice_library
    import voice_clone
    _CLONE_AVAILABLE = True
except Exception:
    _CLONE_AVAILABLE = False

# Esportazione video (slide + audio -> MP4). Import opzionale: se manca, resta
# disponibile solo l'output PowerPoint.
try:
    import video_export
    _VIDEO_AVAILABLE = True
except Exception:
    _VIDEO_AVAILABLE = False


# ----------------------------------------------------------------------------
# Costanti
# ----------------------------------------------------------------------------
# Edge TTS WordBoundary events report offset/duration in 100-nanosecond ticks
# (Windows FILETIME convention). Divide by this to get milliseconds.
HUNDRED_NS_TO_MS = 10000

# OOXML measurement units. EMU = English Metric Unit, the OOXML universal unit.
EMU_PER_INCH = 914400
EMU_PER_PIXEL_96DPI = 9525

# Posizionamento dell'icona audio: la spostiamo di -1 inch (fuori dalla slide)
# così non è mai visibile né in editing né in slideshow.
ICON_OFFSCREEN_EMU = -EMU_PER_INCH
ICON_SIZE_INCHES = 0.4

# ffmpeg subprocess timeout (transcoding di un MP3 da pochi MB).
FFMPEG_TIMEOUT_S = 120

# Timeout per il consumo di un singolo Edge TTS stream. Senza, una connessione
# WebSocket appesa farebbe stallare l'intera generazione. Tipicamente 1 slide
# richiede 2-10s; 90s è una soglia generosa che copre rete lenta ma cattura
# i veri freeze.
EDGE_TTS_STREAM_TIMEOUT_S = 90

# Integrità audio: una singola richiesta TTS può terminare senza eccezione ma
# produrre un file vuoto/troncato (rete interrotta, rate limiting, cache locale
# incompleta). Ogni slide viene quindi tentata più volte e il file viene
# accettato solo dopo apertura con Mutagen e controllo della durata.
AUDIO_GENERATION_ATTEMPTS = 3
AUDIO_RECOVERY_PASSES = 1
AUDIO_RETRY_BASE_DELAY_S = 1.25
MIN_AUDIO_FILE_BYTES = 256
MIN_AUDIO_DURATION_S = 0.05
# Un file viene considerato interamente muto solo se il picco massimo è sotto
# questa soglia. -70 dBFS è prudente: non scarta normali pause o voci registrate
# a volume basso, ma intercetta file composti esclusivamente da silenzio digitale.
MAX_SILENCE_DBFS = -70.0
AUDIO_INSPECTION_TIMEOUT_S = 45
SUPPORTED_AUDIO_EXTENSIONS = {".mp3", ".wav", ".m4a", ".aac"}
SUPPORTED_AUDIO_MIME_TYPES = {
    ".mp3": "audio/mpeg",
    ".wav": "audio/wav",
    ".m4a": "audio/mp4",
    ".aac": "audio/aac",
}

# Anche l'incorporamento e il salvataggio vengono verificati: il primo controllo
# avviene nella struttura python-pptx, il secondo direttamente nel pacchetto ZIP
# OOXML finale, verificando shape, relationship e Media part.
AUDIO_EMBED_ATTEMPTS = 2
PPTX_SAVE_VERIFY_ATTEMPTS = 3

# Per lasciare intenzionalmente muta una slide bisogna dichiararlo nel foglio
# Excel. Una cella vuota non è più interpretata come scelta intenzionale: evita
# che una riga dimenticata generi inavvertitamente una slide senza audio.
SILENT_SCRIPT_MARKERS = {"[SENZA AUDIO]", "[NO AUDIO]"}

# Formato MP3 nativo prodotto da Edge TTS vs formato target dopo transcoding.
EDGE_TTS_NATIVE_SAMPLE_RATE_HZ = 24000
TARGET_SAMPLE_RATE_HZ = 44100
TARGET_BITRATE = "96k"

# Type alias per il callback di progresso (vedi process() per gli stage).
ProgressCallback = Optional[Callable[[dict], None]]


class OperationCancelled(RuntimeError):
    """Interruzione richiesta dall'utente durante una lavorazione."""


def _check_cancelled(cancel_event=None) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise OperationCancelled("Operazione annullata dall'utente.")


def _atomic_write_text(path: str | Path, text: str, encoding: str = "utf-8") -> None:
    """Scrive un sidecar senza lasciare file parziali."""
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=f".{p.name}.", suffix=".tmp", dir=p.parent)
    os.close(fd)
    try:
        Path(tmp).write_text(text, encoding=encoding)
        os.replace(tmp, p)
    finally:
        _remove_file_quietly(tmp)



# Available Italian voices in Edge TTS
ITALIAN_VOICES = {
    "isabella": "it-IT-IsabellaNeural",   # female, warm
    "elsa":     "it-IT-ElsaNeural",       # female, clear
    "diego":    "it-IT-DiegoNeural",      # male, professional
    "giuseppe": "it-IT-GiuseppeNeural",   # male, mature
}


def list_italian_edge_voices() -> list[dict]:
    """Restituisce le voci italiane attualmente esposte da Edge TTS.

    Se il servizio non è raggiungibile, restituisce l'elenco integrato così
    l'interfaccia resta utilizzabile offline.
    """
    async def _load():
        fn = getattr(edge_tts, "list_voices", None)
        if fn is None:
            return []
        voices = await fn()
        result = []
        for item in voices:
            locale = str(item.get("Locale", ""))
            short = str(item.get("ShortName", ""))
            if locale.lower().startswith("it-") and short:
                result.append({
                    "voice_id": short,
                    "name": item.get("FriendlyName") or short,
                    "locale": locale,
                    "gender": item.get("Gender", ""),
                })
        return sorted(result, key=lambda x: (x["gender"], x["name"]))
    try:
        dynamic = asyncio.run(_load())
        if dynamic:
            return dynamic
    except Exception as exc:
        print(f"   ! elenco voci online non disponibile: {exc}")
    return [
        {"voice_id": value, "name": key.title(), "locale": "it-IT", "gender": ""}
        for key, value in ITALIAN_VOICES.items()
    ]


def _pptx_declared_fonts(pptx_path: str) -> set[str]:
    """Restituisce i font realmente utili al rendering della presentazione.

    Evita di trattare come "usati" tutti i font supplementari per alfabeti
    internazionali elencati nel tema Office: quei valori sono fallback del
    tema e producevano decine di falsi avvisi. Vengono considerati i font
    espliciti del testo e i soli font latini principale/secondario del tema.
    """
    fonts: set[str] = set()
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    text_parts = (
        "ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/",
        "ppt/notesSlides/", "ppt/notesMasters/",
    )
    try:
        with zipfile.ZipFile(pptx_path, "r") as zin:
            for name in zin.namelist():
                if not name.endswith(".xml"):
                    continue
                try:
                    root = etree.fromstring(zin.read(name))
                except Exception:
                    continue
                if name.startswith(text_parts):
                    for tag in ("latin", "ea", "cs"):
                        for node in root.findall(f".//{{{a_ns}}}{tag}"):
                            typeface = (node.get("typeface") or "").strip()
                            if typeface and not typeface.startswith("+"):
                                fonts.add(typeface)
                elif name.startswith("ppt/theme/"):
                    ns = {"a": a_ns}
                    paths = (
                        ".//a:fontScheme/a:majorFont/a:latin",
                        ".//a:fontScheme/a:minorFont/a:latin",
                    )
                    for path in paths:
                        for node in root.xpath(path, namespaces=ns):
                            typeface = (node.get("typeface") or "").strip()
                            if typeface and not typeface.startswith("+"):
                                fonts.add(typeface)
    except Exception:
        pass
    return {f for f in fonts if f}


def _installed_font_names() -> set[str] | None:
    names: set[str] = set()
    if sys.platform.startswith("win"):
        try:
            import winreg
            for root, key in (
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            ):
                try:
                    with winreg.OpenKey(root, key) as handle:
                        index = 0
                        while True:
                            try:
                                value_name, _value, _typ = winreg.EnumValue(handle, index)
                            except OSError:
                                break
                            names.add(re.sub(r"\s*\([^)]*\)\s*$", "", value_name).strip())
                            index += 1
                except OSError:
                    pass
            return names
        except Exception:
            return None
    fc_list = shutil.which("fc-list")
    if not fc_list:
        return None
    try:
        proc = subprocess.run(
            [fc_list, "-f", "%{family}\n"], capture_output=True, text=True, timeout=20
        )
        if proc.returncode != 0:
            return None
        for line in proc.stdout.splitlines():
            for family in line.split(","):
                if family.strip():
                    names.add(family.strip())
        return names
    except Exception:
        return None


def preflight_system(
    output_mode: str = "pptx",
    output_path: str | None = None,
    input_pptx: str | None = None,
    render_backend: str = "auto",
) -> dict:
    """Controllo preventivo di programmi, spazio disco e font dichiarati."""
    checks = {"ok": True, "errors": [], "warnings": [], "tools": {}}
    for tool in ("ffmpeg", "ffprobe"):
        found = shutil.which(tool) or (_find_ffmpeg() if tool == "ffmpeg" else None)
        checks["tools"][tool] = found
    if not checks["tools"]["ffmpeg"]:
        checks["warnings"].append(
            "FFmpeg non trovato: controllo del silenzio, transcodifica e video non disponibili."
        )
    if output_mode == "video":
        backend = (render_backend or "auto").lower()
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        checks["tools"]["libreoffice"] = soffice
        has_powerpoint_backend = False
        if sys.platform.startswith("win"):
            try:
                import importlib.util
                has_powerpoint_backend = importlib.util.find_spec("win32com") is not None
            except Exception:
                pass
        checks["tools"]["powerpoint_backend"] = has_powerpoint_backend
        if backend == "libreoffice" and not soffice:
            checks["errors"].append("LibreOffice/soffice non trovato.")
        elif backend == "powerpoint" and not has_powerpoint_backend:
            checks["errors"].append(
                "Backend PowerPoint non disponibile: serve Windows, Microsoft PowerPoint e pywin32."
            )
        elif backend == "auto" and not soffice and not has_powerpoint_backend:
            checks["errors"].append("Nessun backend per renderizzare le slide.")
        if not checks["tools"]["ffmpeg"]:
            checks["errors"].append("FFmpeg non trovato per creare il video.")
        if input_pptx and Path(input_pptx).exists():
            declared = _pptx_declared_fonts(input_pptx)
            installed = _installed_font_names()
            checks["declared_fonts"] = sorted(declared)
            if installed is not None:
                normalized = {name.casefold() for name in installed}
                missing = sorted(font for font in declared if font.casefold() not in normalized)
                checks["missing_fonts"] = missing
                if missing:
                    preview = ", ".join(missing[:12]) + (", ..." if len(missing) > 12 else "")
                    checks["warnings"].append(
                        "Font del PowerPoint non individuati nel sistema; il video potrebbe avere "
                        "sostituzioni grafiche: " + preview
                    )
    if output_path:
        parent = Path(output_path).expanduser().resolve().parent
        parent.mkdir(parents=True, exist_ok=True)
        free = shutil.disk_usage(parent).free
        checks["free_disk_bytes"] = free
        if free < 500 * 1024 * 1024:
            checks["errors"].append("Spazio disco libero inferiore a 500 MB.")
        elif free < 2 * 1024 * 1024 * 1024:
            checks["warnings"].append("Spazio disco libero inferiore a 2 GB.")
    checks["ok"] = not checks["errors"]
    return checks


# Default number of audio syntheses running in parallel.
# Edge TTS streams each MP3 in real time over a WebSocket: the dominant cost
# is wall-clock waiting for bytes to arrive, not CPU. Running several syntheses
# concurrently from a single asyncio event loop gives a near-linear speedup
# without significant overhead, up to a point where Microsoft's free service
# starts rate-limiting (usually around 10+ concurrent connections).
# 6 is a safe sweet spot: ~5x speedup vs sequential, no observed throttling.
DEFAULT_CONCURRENCY = 6
MAX_CLONE_WORKERS = 10


# Margine aggiunto al tempo di avanzamento automatico delle slide: l'audio e
# l'avanzamento partono entrambi all'ingresso nella slide; un piccolo
# cuscinetto evita che la transizione scatti un istante prima che l'audio
# finisca davvero.
AUTO_ADVANCE_TAIL_MS = 300


# ----------------------------------------------------------------------------
# TTS generation
# ----------------------------------------------------------------------------
async def _synthesize(
    text: str, voice: str, rate: str, output_path: str,
    volume: str = "+0%", pitch: str = "+0Hz",
) -> None:
    communicate = edge_tts.Communicate(
        text=text, voice=voice, rate=rate, volume=volume, pitch=pitch
    )
    await communicate.save(output_path)


def synthesize_to_file(
    text: str, voice: str, rate: str, output_path: str,
    volume: str = "+0%", pitch: str = "+0Hz",
) -> None:
    """Sintetizza una singola anteprima con volume e tonalità configurabili."""
    asyncio.run(_synthesize(text, voice, rate, output_path, volume, pitch))


# ----- Parallel synthesis ---------------------------------------------------

# ----------------------------------------------------------------------------
# Transcodifica audio per compatibilità PowerPoint
# ----------------------------------------------------------------------------
# Edge TTS produce MP3 in formato "audio-24khz-48kbitrate-mono-mp3", cioè
# MPEG-2 Layer III @ 24kHz. Questo formato è perfettamente valido come MP3,
# ma PowerPoint (almeno alcune versioni recenti) lo considera "non standard"
# e mostra all'apertura l'avviso "PowerPoint ha rilevato un problema nel
# contenuto" — in fase di Ripristina propone di "Ottimizzare la
# compatibilità clip multimediali".
#
# Soluzione: dopo la sintesi, prima di incorporare nel pptx, riconvertiamo
# l'MP3 in MPEG-1 Layer III @ 44100 Hz mono usando ffmpeg. Questo è il
# formato che PowerPoint accetta nativamente senza avvisi.
#
# Se ffmpeg non è disponibile sul sistema, l'audio viene incorporato così
# com'è: PowerPoint potrebbe mostrare l'avviso ma il file rimane
# funzionante. In tal caso stampiamo un warning una volta sola.
_FFMPEG_PATH: str | None = None
_FFMPEG_WARNING_SHOWN = False


def _find_ffmpeg() -> str | None:
    """Cerca l'eseguibile ffmpeg sul sistema. Restituisce il path o None."""
    global _FFMPEG_PATH
    if _FFMPEG_PATH is not None:
        return _FFMPEG_PATH if _FFMPEG_PATH else None

    # 1. PATH di sistema
    found = shutil.which("ffmpeg")
    if found:
        _FFMPEG_PATH = found
        return found

    # 2. Path comuni su Windows
    if sys.platform.startswith('win'):
        candidates = [
            r"C:\ffmpeg\bin\ffmpeg.exe",
            r"C:\Program Files\ffmpeg\bin\ffmpeg.exe",
            r"C:\Program Files (x86)\ffmpeg\bin\ffmpeg.exe",
            os.path.expanduser(r"~\ffmpeg\bin\ffmpeg.exe"),
            os.path.expanduser(r"~\AppData\Local\Programs\ffmpeg\bin\ffmpeg.exe"),
        ]
        for c in candidates:
            if os.path.exists(c):
                _FFMPEG_PATH = c
                return c

    _FFMPEG_PATH = ""  # marca come "cercato e non trovato"
    return None


def _transcode_mp3_for_powerpoint(input_path: str, output_path: str) -> bool:
    """
    Riconverte input_path in MPEG-1 Layer III @ 44100 Hz mono 96kbps,
    formato accettato senza avvisi da PowerPoint, e scrive in output_path.
    Restituisce True se ok, False se ffmpeg non c'è o fallisce.
    """
    global _FFMPEG_WARNING_SHOWN
    ffmpeg = _find_ffmpeg()
    if not ffmpeg:
        if not _FFMPEG_WARNING_SHOWN:
            print(
                "   ! ffmpeg non trovato — gli audio non verranno riconvertiti.\n"
                "     PowerPoint potrebbe mostrare 'PowerPoint ha rilevato un\n"
                "     problema nel contenuto' all'apertura. Per evitarlo,\n"
                "     installa ffmpeg da https://www.gyan.dev/ffmpeg/builds/\n"
                "     (Windows) e aggiungilo al PATH, poi rilancia."
            )
            _FFMPEG_WARNING_SHOWN = True
        return False

    try:
        result = subprocess.run(
            [ffmpeg, "-y", "-loglevel", "error",
             "-i", input_path,
             "-ar", str(TARGET_SAMPLE_RATE_HZ), "-ac", "1",
             "-b:a", TARGET_BITRATE,
             "-codec:a", "libmp3lame",
             output_path],
            capture_output=True, timeout=FFMPEG_TIMEOUT_S,
        )
        if result.returncode != 0:
            err = result.stderr.decode(errors='replace')[:200]
            print(f"   ! ffmpeg ha restituito errore: {err}")
            return False
        return True
    except Exception as e:
        print(f"   ! Eccezione durante transcoding ffmpeg: {e}")
        return False


# ----------------------------------------------------------------------------
# Sintesi singola
# ----------------------------------------------------------------------------
async def _consume_stream(communicate, audio_out, words: list, cancel_event=None) -> None:
    """Consuma lo stream di una Communicate edge_tts: scrive i chunk audio
    su `audio_out` e accumula i WordBoundary in `words`.

    Estratto in funzione separata per poterla wrappare con asyncio.wait_for
    e fare timeout senza altri side-effect."""
    async for chunk in communicate.stream():
        _check_cancelled(cancel_event)
        ctype = chunk.get("type")
        if ctype == "audio":
            audio_out.write(chunk["data"])
        elif ctype in ("WordBoundary", "SentenceBoundary"):
            offset_ms = chunk.get("offset", 0) // HUNDRED_NS_TO_MS
            duration_ms = chunk.get("duration", 0) // HUNDRED_NS_TO_MS
            words.append({
                "text": chunk.get("text", ""),
                "offset_ms": int(offset_ms),
                "duration_ms": int(duration_ms),
            })


async def _synthesize_one(
    slide_num: int,
    text: str,
    voice: str,
    rate: str,
    output_path: str,
    volume: str,
    pitch: str,
    semaphore: asyncio.Semaphore,
    progress_state: dict,
    progress_callback: ProgressCallback,
    captured_timings: dict | None = None,
    transcode_audio: bool = False,
    cancel_event=None,
) -> tuple[int, bool, str]:
    """Sintetizza una slide con retry e controllo reale del file prodotto.

    Ogni tentativo scrive su un file temporaneo; solo un MP3 decodificabile e
    con durata positiva viene rinominato atomicamente nel percorso definitivo.
    In questo modo un tentativo interrotto non può essere scambiato per audio
    valido durante la successiva fase di incorporamento.
    """
    async with semaphore:
        t0_all = time.monotonic()
        last_error = "errore sconosciuto"
        ok = False
        final_words: list[dict] = []

        for attempt in range(1, AUDIO_GENERATION_ATTEMPTS + 1):
            _check_cancelled(cancel_event)
            t0 = time.monotonic()
            raw_tmp = output_path + ".part.mp3"
            transcoded_tmp = output_path + ".transcoded.part.mp3"
            _remove_file_quietly(raw_tmp)
            _remove_file_quietly(transcoded_tmp)
            _remove_file_quietly(output_path)
            words: list[dict] = []

            try:
                try:
                    communicate = edge_tts.Communicate(
                        text=text, voice=voice, rate=rate, volume=volume,
                        pitch=pitch, boundary="WordBoundary",
                    )
                except TypeError:
                    # Compatibilità con versioni/stub meno recenti.
                    communicate = edge_tts.Communicate(
                        text=text, voice=voice, rate=rate
                    )

                with open(raw_tmp, "wb") as audio_out:
                    try:
                        await asyncio.wait_for(
                            _consume_stream(communicate, audio_out, words, cancel_event),
                            timeout=EDGE_TTS_STREAM_TIMEOUT_S,
                        )
                    except asyncio.TimeoutError:
                        raise RuntimeError(
                            f"timeout dopo {EDGE_TTS_STREAM_TIMEOUT_S}s "
                            f"(server Edge TTS non risponde)"
                        ) from None

                valid, reason = validate_audio_file(raw_tmp)
                if not valid:
                    raise RuntimeError(reason)

                if transcode_audio:
                    loop = asyncio.get_running_loop()
                    transcoded = await loop.run_in_executor(
                        None, _transcode_mp3_for_powerpoint,
                        raw_tmp, transcoded_tmp,
                    )
                    if not transcoded:
                        raise RuntimeError("transcodifica ffmpeg fallita")
                    valid, reason = validate_audio_file(transcoded_tmp)
                    if not valid:
                        raise RuntimeError(
                            f"audio transcodificato non valido: {reason}"
                        )
                    os.replace(transcoded_tmp, output_path)
                    _remove_file_quietly(raw_tmp)
                else:
                    os.replace(raw_tmp, output_path)

                valid, reason = validate_audio_file(output_path)
                if not valid:
                    raise RuntimeError(
                        f"controllo finale del file audio fallito: {reason}"
                    )

                elapsed = time.monotonic() - t0
                retry_note = "" if attempt == 1 else f", tentativo {attempt}"
                print(
                    f"   slide {slide_num}: sintesi ok ({elapsed:.1f}s, "
                    f"{len(text)} caratteri, {len(words)} parole timed"
                    f"{retry_note})"
                )
                final_words = words
                ok = True
                last_error = ""
                break

            except OperationCancelled:
                _remove_file_quietly(raw_tmp)
                _remove_file_quietly(transcoded_tmp)
                _remove_file_quietly(output_path)
                raise
            except Exception as exc:
                last_error = str(exc)
                _remove_file_quietly(raw_tmp)
                _remove_file_quietly(transcoded_tmp)
                _remove_file_quietly(output_path)
                elapsed = time.monotonic() - t0
                if attempt < AUDIO_GENERATION_ATTEMPTS:
                    delay = AUDIO_RETRY_BASE_DELAY_S * attempt
                    print(
                        f"   slide {slide_num}: tentativo {attempt}/"
                        f"{AUDIO_GENERATION_ATTEMPTS} fallito dopo "
                        f"{elapsed:.1f}s — {exc}; nuovo tentativo"
                    )
                    await asyncio.sleep(delay)
                else:
                    print(
                        f"   slide {slide_num}: ERRORE sintesi dopo "
                        f"{time.monotonic() - t0_all:.1f}s e "
                        f"{AUDIO_GENERATION_ATTEMPTS} tentativi — {exc}"
                    )

        if ok and captured_timings is not None:
            captured_timings[slide_num] = {
                "text": text,
                "words": final_words,
            }

        progress_state["done"] += 1
        if progress_callback is not None:
            try:
                progress_callback({
                    "stage": "synthesis_progress",
                    "done": progress_state["done"],
                    "total": progress_state["total"],
                    "slide_num": slide_num,
                    "ok": ok,
                })
            except Exception:
                pass
        return (slide_num, ok, last_error)


async def _synthesize_many(
    tasks: list[tuple[int, str, str]],
    voice: str,
    rate: str,
    volume: str,
    pitch: str,
    concurrency: int,
    progress_callback: ProgressCallback = None,
    captured_timings: dict | None = None,
    transcode_audio: bool = False,
    cancel_event=None,
) -> list[tuple[int, bool, str]]:
    """
    Run all synthesis tasks concurrently, capped at `concurrency` in flight.
    `tasks` is a list of (slide_num, text, output_path).
    Returns the list of (slide_num, success, error) results.

    Se `captured_timings` viene fornito (dict vuoto), lo riempiamo con i
    word-level timings di ciascuna slide. Vedi _synthesize_one per i dettagli.
    """
    semaphore = asyncio.Semaphore(concurrency)
    progress_state = {"done": 0, "total": len(tasks)}
    coros = [
        _synthesize_one(
            slide_num, text, voice, rate, path, volume, pitch, semaphore,
            progress_state, progress_callback,
            captured_timings=captured_timings,
            transcode_audio=transcode_audio, cancel_event=cancel_event,
        )
        for slide_num, text, path in tasks
    ]
    return await asyncio.gather(*coros)


def synthesize_many_to_files(
    tasks: list[tuple[int, str, str]],
    voice: str,
    rate: str,
    volume: str = "+0%",
    pitch: str = "+0Hz",
    concurrency: int = DEFAULT_CONCURRENCY,
    progress_callback: ProgressCallback = None,
    captured_timings: dict | None = None,
    transcode_audio: bool = False,
    cancel_event=None,
) -> list[tuple[int, bool, str]]:
    """
    Synchronous wrapper. Runs ALL syntheses inside ONE asyncio event loop
    (instead of one event loop per slide as the old per-slide
    `synthesize_to_file()` would do). Avoids loop setup/teardown overhead.

    Se `captured_timings` è un dict, viene riempito con i word-level timings
    raccolti durante la sintesi (vedi _synthesize_one).
    """
    if not tasks:
        return []
    return asyncio.run(_synthesize_many(
        tasks, voice, rate, volume, pitch, concurrency, progress_callback,
        captured_timings=captured_timings,
        transcode_audio=transcode_audio, cancel_event=cancel_event,
    ))


def _audio_duration_seconds_generic(audio_path: str) -> float:
    """Legge la durata di MP3/WAV/M4A/AAC con Mutagen o ffprobe."""
    try:
        media = MutagenFile(audio_path)
        length = getattr(getattr(media, "info", None), "length", None)
        if length is not None:
            duration = float(length)
            if math.isfinite(duration):
                return duration
    except Exception:
        pass

    ffprobe = shutil.which("ffprobe")
    if not ffprobe:
        ffmpeg = _find_ffmpeg()
        if ffmpeg:
            sibling = Path(ffmpeg).with_name("ffprobe.exe" if sys.platform.startswith("win") else "ffprobe")
            if sibling.exists():
                ffprobe = str(sibling)
    if not ffprobe:
        raise ValueError("formato non riconosciuto da Mutagen e ffprobe non disponibile")
    proc = subprocess.run(
        [ffprobe, "-v", "error", "-select_streams", "a:0",
         "-show_entries", "format=duration", "-of", "default=nw=1:nk=1",
         audio_path],
        capture_output=True, timeout=AUDIO_INSPECTION_TIMEOUT_S,
    )
    if proc.returncode != 0:
        raise ValueError(proc.stderr.decode(errors="replace")[:250] or "ffprobe non riconosce l'audio")
    raw = proc.stdout.decode(errors="replace").strip().splitlines()
    if not raw:
        raise ValueError("durata audio non disponibile")
    duration = float(raw[-1])
    if not math.isfinite(duration):
        raise ValueError(f"durata non finita: {duration!r}")
    return duration


def _audio_max_volume_dbfs(audio_path: str) -> float | None:
    """Restituisce il picco massimo in dBFS tramite FFmpeg.

    None significa che FFmpeg non è disponibile: la validazione strutturale
    resta attiva, ma il controllo di silenzio non può essere eseguito.
    """
    ffmpeg = _find_ffmpeg()
    if not ffmpeg:
        return None
    try:
        proc = subprocess.run(
            [ffmpeg, "-nostdin", "-hide_banner", "-i", audio_path,
             "-map", "0:a:0", "-af", "volumedetect", "-f", "null", "-"],
            capture_output=True, timeout=AUDIO_INSPECTION_TIMEOUT_S,
        )
        stderr = proc.stderr.decode(errors="replace")
        match = re.search(r"max_volume:\s*(-?inf|[-+]?\d+(?:\.\d+)?)\s*dB", stderr, re.I)
        if not match:
            return None
        raw = match.group(1).lower()
        return float("-inf") if raw in {"-inf", "inf"} else float(raw)
    except Exception:
        return None


def inspect_audio_file(audio_path: str, require_signal: bool = True) -> tuple[bool, str, dict]:
    """Valida i formati audio supportati da PowerPoint e restituisce dettagli."""
    details: dict = {"path": audio_path}
    try:
        if not os.path.isfile(audio_path):
            return False, "file non creato", details
        size = os.path.getsize(audio_path)
        details["size_bytes"] = size
        if size < MIN_AUDIO_FILE_BYTES:
            return False, f"file troppo piccolo ({size} byte)", details
        ext = Path(audio_path).suffix.lower()
        details["extension"] = ext
        if ext and ext not in SUPPORTED_AUDIO_EXTENSIONS:
            return False, f"formato audio non supportato da PowerPoint ({ext})", details
        duration = _audio_duration_seconds_generic(audio_path)
        details["duration_s"] = duration
        if not math.isfinite(duration) or duration < MIN_AUDIO_DURATION_S:
            return False, f"durata non valida ({duration!r}s)", details
        max_db = _audio_max_volume_dbfs(audio_path) if require_signal else None
        details["max_volume_dbfs"] = max_db
        if require_signal and max_db is not None and (
            not math.isfinite(max_db) or max_db <= MAX_SILENCE_DBFS
        ):
            return False, (
                "audio completamente silenzioso "
                f"(picco {max_db if math.isfinite(max_db) else '-inf'} dBFS)"
            ), details
        return True, "", details
    except Exception as exc:
        return False, f"audio non leggibile: {type(exc).__name__}: {exc}", details


def validate_audio_file(audio_path: str) -> tuple[bool, str]:
    """Compatibilità API: controlla struttura, durata e assenza di silenzio totale."""
    ok, reason, _details = inspect_audio_file(audio_path, require_signal=True)
    return ok, reason

def _remove_file_quietly(path: str) -> None:
    try:
        if os.path.exists(path):
            os.unlink(path)
    except Exception:
        pass


def get_audio_duration_seconds(audio_path: str) -> float:
    """Restituisce la durata di un audio validato (MP3/WAV/M4A/AAC)."""
    ok, error, details = inspect_audio_file(audio_path, require_signal=True)
    if not ok:
        raise ValueError(f"Audio non valido '{audio_path}': {error}")
    return float(details["duration_s"])


def format_duration(seconds: float) -> str:
    """Format seconds as M:SS (e.g. 75.4 -> '1:15') or H:MM:SS for >= 1h
    (e.g. 3661 -> '1:01:01'). Negative input is clamped to 0."""
    total = max(0, int(round(seconds)))
    if total >= 3600:
        h, rem = divmod(total, 3600)
        m, s = divmod(rem, 60)
        return f"{h}:{m:02d}:{s:02d}"
    return f"{total // 60}:{total % 60:02d}"


# ----------------------------------------------------------------------------
# XLSX reading
# ----------------------------------------------------------------------------
def _normalise_header(value) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def read_scripts_from_xlsx(
    xlsx_path: str,
    sheet_name: str | None = None,
    column: str = "A",
    has_header: bool = False,
    slide_metadata: list[dict] | None = None,
) -> list[str]:
    """Legge gli script con controllo formule e mapping opzionale per ID slide.

    Compatibilità: senza opzioni usa il foglio attivo e la colonna A, una riga
    per slide. Se la prima riga contiene intestazioni riconoscibili (Script,
    Numero slide, ID slide), può associare gli script in modo stabile anche
    dopo un riordino delle slide.
    """
    try:
        col_idx = column_index_from_string(str(column).strip().upper())
    except Exception as exc:
        raise ValueError(f"Colonna Excel non valida: {column!r}") from exc

    wb_values = load_workbook(xlsx_path, data_only=True)
    wb_formulas = load_workbook(xlsx_path, data_only=False)
    try:
        if sheet_name:
            if sheet_name not in wb_values.sheetnames:
                raise ValueError(
                    f"Il foglio '{sheet_name}' non esiste. Fogli disponibili: "
                    + ", ".join(wb_values.sheetnames)
                )
            ws = wb_values[sheet_name]
            ws_formula = wb_formulas[sheet_name]
        else:
            ws = wb_values.active
            ws_formula = wb_formulas[ws.title]

        # Riconoscimento tabella avanzata tramite intestazioni.
        header_values = [ws_formula.cell(1, c).value for c in range(1, ws_formula.max_column + 1)]
        headers = {_normalise_header(v): i + 1 for i, v in enumerate(header_values) if v is not None}
        script_header = next((headers[k] for k in ("script", "testo", "testotts", "testoaudio") if k in headers), None)
        number_header = next((headers[k] for k in ("numeroslide", "slide", "nslide") if k in headers), None)
        id_header = next((headers[k] for k in ("idslide", "slideid") if k in headers), None)
        advanced = bool(script_header and (number_header or id_header))
        if advanced:
            has_header = True
            col_idx = script_header

        start_row = 2 if has_header else 1
        max_row = max(ws.max_row, ws_formula.max_row)
        missing_formula_cache: list[str] = []

        def read_cell(row: int, col: int) -> str:
            value = ws.cell(row, col).value
            raw = ws_formula.cell(row, col).value
            if value is None and isinstance(raw, str) and raw.startswith("="):
                missing_formula_cache.append(ws_formula.cell(row, col).coordinate)
                return ""
            return "" if value is None else str(value).strip()

        if advanced and slide_metadata is not None:
            by_num: dict[int, str] = {}
            by_id: dict[str, str] = {}
            for row in range(start_row, max_row + 1):
                text = read_cell(row, col_idx)
                if number_header:
                    raw_num = ws.cell(row, number_header).value
                    try:
                        if raw_num is not None:
                            by_num[int(raw_num)] = text
                    except (TypeError, ValueError):
                        pass
                if id_header:
                    raw_id = ws.cell(row, id_header).value
                    if raw_id is not None:
                        by_id[str(raw_id).strip()] = text
            scripts = []
            for meta in slide_metadata:
                slide_id = str(meta.get("slide_id", ""))
                slide_num = int(meta["slide_num"])
                scripts.append(by_id.get(slide_id, by_num.get(slide_num, "")))
        else:
            scripts = [read_cell(row, col_idx) for row in range(start_row, max_row + 1)]

        if missing_formula_cache:
            cells = ", ".join(missing_formula_cache[:15])
            if len(missing_formula_cache) > 15:
                cells += ", ..."
            raise ValueError(
                "Il file Excel contiene formule senza un risultato calcolato "
                f"memorizzato nelle celle {cells}. Apri il file in Excel, "
                "ricalcola e salva, oppure sostituisci le formule con testo."
            )
    finally:
        wb_values.close()
        wb_formulas.close()

    while scripts and not scripts[-1]:
        scripts.pop()
    return scripts


def get_pptx_slide_metadata(pptx_path: str) -> list[dict]:
    prs = Presentation(pptx_path)
    result = []
    for num, slide in enumerate(prs.slides, 1):
        title = ""
        try:
            if slide.shapes.title is not None:
                title = slide.shapes.title.text.strip()
        except Exception:
            pass
        result.append({"slide_num": num, "slide_id": str(slide.slide_id), "title": title})
    return result


def read_scripts_from_notes(pptx_path: str) -> list[str]:
    """Estrae il testo del corpo delle Note PowerPoint, una voce per slide."""
    scripts: list[str] = []
    pkg_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    with zipfile.ZipFile(pptx_path, 'r') as zin:
        names = set(zin.namelist())
        for slide_part in _ordered_slide_parts(zin):
            rels_part = posixpath.join(posixpath.dirname(slide_part), '_rels', posixpath.basename(slide_part) + '.rels')
            note_part = None
            if rels_part in names:
                root = etree.fromstring(zin.read(rels_part))
                for rel in root.findall(f'{{{pkg_ns}}}Relationship'):
                    if (rel.get('Type') or '').endswith('/notesSlide'):
                        note_part = _resolve_package_target(slide_part, rel.get('Target', ''))
                        break
            text_parts: list[str] = []
            if note_part and note_part in names:
                note_root = etree.fromstring(zin.read(note_part))
                for sp in note_root.findall(f'.//{{{p_ns}}}sp'):
                    ph = sp.find(f'.//{{{p_ns}}}ph')
                    ph_type = ph.get('type') if ph is not None else None
                    if ph_type in {"hdr", "ftr", "dt", "sldNum"}:
                        continue
                    chunks = [t.text or '' for t in sp.findall(f'.//{{{a_ns}}}t')]
                    joined = ' '.join(x.strip() for x in chunks if x.strip()).strip()
                    if joined:
                        text_parts.append(joined)
            scripts.append('\n'.join(text_parts).strip())
    while scripts and not scripts[-1]:
        scripts.pop()
    return scripts


def estimate_scripts(scripts: list[str], rate: str = "+0%") -> dict:
    """Stima parole, caratteri e durata del parlato prima della sintesi."""
    usable = [s for s in scripts if s and s.upper() not in SILENT_SCRIPT_MARKERS]
    words = sum(len(re.findall(r"\b\w+\b", s, flags=re.UNICODE)) for s in usable)
    chars = sum(len(s) for s in usable)
    try:
        pct = int(str(rate).strip().rstrip("%"))
    except Exception:
        pct = 0
    wpm = max(70.0, 150.0 * (1.0 + pct / 100.0))
    seconds = words / wpm * 60.0 if words else 0.0
    return {
        "slides_with_script": len(usable),
        "words": words,
        "characters": chars,
        "estimated_seconds": seconds,
        "estimated_duration": format_duration(seconds),
    }


def create_script_template(pptx_path: str, output_xlsx: str) -> str:
    """Crea un modello Excel associabile per numero e ID stabile della slide."""
    metadata = get_pptx_slide_metadata(pptx_path)
    try:
        analysis = analyze_pptx_audio(pptx_path)
    except Exception:
        analysis = {"details": {}}
    wb = Workbook()
    ws = wb.active
    ws.title = "Script"
    ws.append(["Numero slide", "ID slide", "Titolo", "Stato audio", "Script"])
    for meta in metadata:
        detail = analysis.get("details", {}).get(meta["slide_num"], {})
        status = detail.get("status", "non analizzato")
        ws.append([meta["slide_num"], meta["slide_id"], meta["title"], status, ""])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{max(1, len(metadata)+1)}"
    widths = {"A": 14, "B": 16, "C": 45, "D": 18, "E": 90}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    wb.save(output_xlsx)
    return output_xlsx


# ----------------------------------------------------------------------------
# Conversione word-timings → frasi con timing
# ----------------------------------------------------------------------------
def _split_into_sentences(text: str) -> list[str]:
    """
    Spezza il testo in frasi usando come separatori . ? ! seguiti da spazio
    (o fine stringa). Mantiene la punteggiatura. Per testi italiani con
    abbreviazioni comuni (es. "Dott.", "es.") la divisione è approssimativa
    ma sufficiente per i sottotitoli.
    """
    # Regex: match fino a un terminatore di frase + eventuali spazi
    # I terminatori sono . ? ! ; — ;
    parts = re.split(r'(?<=[.!?])\s+', text.strip())
    # Filtro le stringhe vuote
    return [p.strip() for p in parts if p.strip()]


def _normalize_word(w: str) -> str:
    """Rimuove punteggiatura per il matching parola↔frase."""
    return re.sub(r'[^\w\s]', '', w).lower().strip()


def _sentences_with_timing(text: str, words: list[dict]) -> list[dict]:
    """
    Dato il testo originale e la lista di parole con offset_ms da Edge TTS,
    restituisce una lista di {"text": ..., "start_ms": ..., "end_ms": ...}
    per ogni frase.

    Algoritmo:
      1. Divido il testo in frasi.
      2. Per ogni frase, conto quante parole "logiche" contiene.
      3. Cammino sulla lista di words consumando esattamente quel numero di
         parole. La start_ms è l'offset della prima parola, l'end_ms è
         l'offset+duration dell'ultima.
    """
    if not text or not words:
        return []

    sentences = _split_into_sentences(text)
    if not sentences:
        return []

    # Conta delle "parole" per ogni frase, usando la stessa euristica di
    # tokenizzazione di Edge TTS (split su whitespace + punctuation rimossa)
    def word_count(s):
        return len([w for w in re.split(r'\s+', s) if _normalize_word(w)])

    result = []
    word_idx = 0
    n_words = len(words)
    for sent in sentences:
        n = word_count(sent)
        if n == 0 or word_idx >= n_words:
            continue
        # Prendo n parole, ma non oltre la fine
        end_idx = min(word_idx + n, n_words)
        first = words[word_idx]
        last = words[end_idx - 1]
        result.append({
            "text": sent,
            "start_ms": int(first["offset_ms"]),
            "end_ms": int(last["offset_ms"] + last.get("duration_ms", 0)),
        })
        word_idx = end_idx

    # Se sono rimaste parole non assegnate (incongruenza tokenizzazione tra il
    # nostro splitter e quello di Edge TTS), le aggancio all'ultima frase
    if word_idx < n_words and result:
        last_word = words[-1]
        result[-1]["end_ms"] = int(last_word["offset_ms"] + last_word.get("duration_ms", 0))

    return result


# ----------------------------------------------------------------------------
# Audio embedding (autoplay + hidden icon)
# ----------------------------------------------------------------------------
def _build_autoplay_timing(shape_id: int) -> etree._Element:
    """
    Build a <p:timing> element that:
      - starts the audio automatically when the slide is shown (delay="0")
      - hides the audio icon when not playing (showWhenStopped="0")
      - keeps mainSeq EMPTY so the user's first click advances to the next
        slide instead of triggering a second playback.

    BUG FIX (doppio audio):
    -----------------------
    The earlier version of this function injected a media-call effect inside
    mainSeq:

        <p:cTn id="5" presetClass="mediacall" ...>
          <p:cmd type="call" cmd="playFrom(0.0)">
            <p:tgtEl><p:spTgt spid="..."/></p:tgtEl>

    That effect was triggered by the slide's first click event, which is the
    very same input the user uses to advance the slideshow. Result:

      1. Slide enters → <p:audio delay="0"> starts the audio (correct).
      2. User presses Space/click to advance → mainSeq fires the media-call
         and the SAME audio restarts from 0. The click is "eaten" by the
         animation step instead of advancing.
      3. User presses Space/click again → finally goes to the next slide.

    The audio thus appeared to play twice, with the second click required for
    advancing. Removing the media-call (mainSeq is now empty) fixes both
    issues: the audio plays once, and Space/click immediately advances.

    The "Play in Background" autoplay behaviour is fully provided by the
    <p:audio>/<p:cMediaNode> block alone with <p:cond delay="0"/>.
    """
    timing_xml = f'''<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst/>
            </p:cTn>
            <p:prevCondLst>
              <p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>
            </p:prevCondLst>
            <p:nextCondLst>
              <p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>
            </p:nextCondLst>
          </p:seq>
          <p:audio>
            <p:cMediaNode showWhenStopped="0">
              <p:cTn id="3" fill="hold" display="0">
                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
              </p:cTn>
              <p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl>
            </p:cMediaNode>
          </p:audio>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>'''
    return etree.fromstring(timing_xml)


def _build_auto_advance_transition(advance_ms: int) -> etree._Element:
    """
    Costruisce un <p:transition advTm="..."> senza effetto visivo: serve solo
    a far avanzare automaticamente la slide dopo `advance_ms` millisecondi.
    advClick="1" lascia comunque all'utente la possibilità di anticipare col
    clic. Combinato con l'autoplay dell'audio, produce una presentazione che
    si riproduce da sola, slide dopo slide.
    """
    xml = (
        '<p:transition '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'advClick="1" advTm="{int(advance_ms)}"/>'
    )
    return etree.fromstring(xml)


def _set_slide_auto_advance(slide, duration_s: float) -> None:
    """
    Imposta sulla slide l'avanzamento automatico a fine audio, pari alla
    durata dell'audio più un piccolo margine. Rispetta l'ordine richiesto
    dallo schema OOXML: <p:transition> va dopo cSld/clrMapOvr e PRIMA di
    <p:timing> (quello dell'autoplay).
    """
    sld = slide.element
    advance_ms = int(round(duration_s * 1000)) + AUTO_ADVANCE_TAIL_MS

    # Mantiene l'effetto visivo già configurato e aggiorna soltanto gli
    # attributi che governano l'avanzamento.
    transition = sld.find(qn('p:transition'))
    if transition is not None:
        transition.set('advClick', '1')
        transition.set('advTm', str(advance_ms))
        return
    transition = _build_auto_advance_transition(advance_ms)

    timing = sld.find(qn('p:timing'))
    if timing is not None:
        timing.addprevious(transition)
        return
    anchor = sld.find(qn('p:clrMapOvr'))
    if anchor is None:
        anchor = sld.find(qn('p:cSld'))
    if anchor is not None:
        anchor.addnext(transition)
    else:
        sld.append(transition)


def _convert_video_ref_to_audio_ref(movie_shape) -> None:
    """
    python-pptx's add_movie always inserts an <a:videoFile> element regardless
    of mime type. Convert it to <a:audioFile> so PowerPoint treats the media
    as audio — needed for the "Play in Background" style to apply correctly.

    Inoltre rimuove un <a:hlinkClick r:id="" action="ppaction://media"/> con
    r:id vuoto: python-pptx.add_movie lo inserisce sempre, ma PowerPoint
    (almeno alcune versioni) lo considera corrotto perché r:id="" non è un
    valore valido secondo lo schema OOXML — un r:id deve puntare a una
    relationship esistente. Risultato: PowerPoint apre il file mostrando
    "PowerPoint ha rilevato un problema nel contenuto" e su Ripristina rimuove
    l'audio.
    """
    nvPr = movie_shape.element.find('.//' + qn('p:nvPr'))

    # 1. Conversione videoFile → audioFile
    if nvPr is not None:
        video_file = nvPr.find(qn('a:videoFile'))
        if video_file is not None:
            audio_file = etree.SubElement(nvPr, qn('a:audioFile'))
            for k, v in video_file.attrib.items():
                audio_file.set(k, v)
            nvPr.remove(video_file)
            nvPr.insert(0, audio_file)

    # 2. Pulizia degli hlinkClick con r:id="" inseriti da python-pptx
    cNvPr = movie_shape.element.find('.//' + qn('p:cNvPr'))
    if cNvPr is not None:
        rId_attr = qn('r:id')
        for hlink in list(cNvPr.findall(qn('a:hlinkClick'))):
            rid = hlink.get(rId_attr)
            if rid == "" or rid is None:
                # Rimuovo SOLO l'hyperlink "vuoto". Se ce ne fossero altri
                # con r:id valido (mai capita questa situazione su audio
                # autogenerati, ma per sicurezza) li lascio stare.
                cNvPr.remove(hlink)


# OOXML relationship type URIs.
# python-pptx's add_movie() unconditionally creates a relationship of type
# "video" (this URI) regardless of the media's actual content type. After
# we convert <a:videoFile> to <a:audioFile> in the shape XML, the file ends
# up with an <a:audioFile r:link="rIdN"/> pointing to a relationship that's
# still tagged "video" in the .rels file. PowerPoint flags this as a corrupted
# file ("PowerPoint ha rilevato un problema nel contenuto") and on Repair it
# removes the audio timing — which is why audio embedded with the previous
# version of this script appeared to work after Repair but no longer played
# automatically.
_VIDEO_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
_AUDIO_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
_MEDIA_RELTYPE = "http://schemas.microsoft.com/office/2007/relationships/media"
_IMAGE_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
_PKG_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Audio file extensions edge-tts may produce / users may bring in. Only used
# as a heuristic to know whether a "video" relationship actually points to
# audio data and should be re-tagged.
_AUDIO_EXTENSIONS = {'mp3', 'wav', 'm4a', 'aac', 'ogg', 'wma'}


def _fix_audio_relationship_types(pptx_path: str) -> tuple[int, int]:
    """
    Post-process the saved .pptx:
    1. Retag relationships that point to audio media but are incorrectly
       labelled as "video" (a side effect of python-pptx's add_movie).
    2. Remove printerSettings parts and their relationships. Some PPTX files
       contain a /ppt/printerSettings/printerSettingsN.bin part that is a
       Windows DEVMODE blob specific to the printer driver of the machine
       that authored the file. When opened on another machine PowerPoint
       cannot decode the DEVMODE and shows the warning "PowerPoint ha
       rilevato un problema nel contenuto" — even though the audio is fine.
       These files are not needed at all (PowerPoint will recreate them on
       Save), so we strip them out.

    Returns (fixed_total, mainseq_fixes):
      - fixed_total: number of changes made overall (rels retagged +
        printer parts removed + Default Extension orfani + slide mainSeq
        riparate)
      - mainseq_fixes: count delle slide riparate per il childTnLst vuoto

    Operates at the zip level rather than via python-pptx because the
    relationship type is exposed as a read-only lazyproperty in the library
    and cannot be mutated from the public API. Editing the .pptx zip
    directly is robust against future python-pptx changes.
    """
    with open(pptx_path, 'rb') as f:
        original_bytes = f.read()

    fixed_total = 0
    mainseq_fixes = 0

    with BytesIO(original_bytes) as in_buf, BytesIO() as new_buffer:
        with zipfile.ZipFile(in_buf) as zin:
            # Pre-pass: find all printerSettings parts so we can strip them and
            # their rels references in the second pass
            printer_parts = {
                n for n in zin.namelist()
                if n.startswith('ppt/printerSettings/') and n.endswith('.bin')
            }

            with zipfile.ZipFile(new_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
                for name in zin.namelist():
                    # Skip printerSettings parts entirely
                    if name in printer_parts:
                        fixed_total += 1
                        continue

                    data = zin.read(name)

                    # Slide xml: ripara mainSeq.childTnLst vuoto. Necessario
                    # perché _build_autoplay_timing genera questo timing e
                    # PowerPoint moderno lo considera corrotto. Agisce qui,
                    # dopo Presentation.save(), così non può essere riscritto.
                    if (name.startswith('ppt/slides/slide')
                            and name.endswith('.xml')):
                        new_data, mainseq_fixed = _fix_empty_mainseq_in_slide_xml(data)
                        if mainseq_fixed:
                            mainseq_fixes += 1
                            fixed_total += 1
                            data = new_data

                    # Slide rels: video → audio retag
                    if (name.startswith('ppt/slides/_rels/')
                            and name.endswith('.xml.rels')):
                        new_data, n_fixed = _retag_video_rels_to_audio(data)
                        if n_fixed > 0:
                            fixed_total += n_fixed
                            data = new_data

                    # presentation.xml.rels: rimuovi i Relationship a printerSettings
                    if (name == 'ppt/_rels/presentation.xml.rels'
                            and printer_parts):
                        new_data, n_removed = _remove_printer_rels(data)
                        if n_removed > 0:
                            fixed_total += n_removed
                            data = new_data

                    # [Content_Types].xml: rimuovi <Default Extension="X"/>
                    # orfani, dove non esiste alcun file .X nel pacchetto.
                    #
                    # Caso reale: il template del PPTX di partenza dichiara
                    # <Default Extension="bin" ContentType="...printerSettings"/>
                    # ma noi abbiamo appena rimosso tutti i .bin. Il Default
                    # orfano rimane e PowerPoint mostra "PowerPoint non è
                    # riuscito a leggere parte del contenuto" all'apertura.
                    if name == '[Content_Types].xml':
                        try:
                            ct_root = etree.fromstring(data)
                            CT_NS = '{http://schemas.openxmlformats.org/package/2006/content-types}'
                            # Estensioni effettivamente presenti nel pacchetto
                            # finale (escludendo quelle già marcate per rimozione).
                            present_exts = set()
                            for n in zin.namelist():
                                if n in printer_parts:
                                    continue
                                tail = n.split('/')[-1]
                                if '.' in tail:
                                    present_exts.add(tail.rsplit('.', 1)[-1].lower())

                            removed_defaults = 0
                            for default in list(ct_root.findall(f'{CT_NS}Default')):
                                ext = (default.get('Extension') or '').lower()
                                if ext and ext not in present_exts:
                                    ct_root.remove(default)
                                    removed_defaults += 1
                                    fixed_total += 1

                            if removed_defaults > 0:
                                data = etree.tostring(
                                    ct_root, xml_declaration=True,
                                    encoding='UTF-8', standalone=True,
                                )
                        except etree.XMLSyntaxError:
                            pass  # [Content_Types].xml malformato — lascio com'è

                    info = zin.getinfo(name)
                    zout.writestr(info, data)

        if fixed_total > 0:
            payload = new_buffer.getvalue()
            with open(pptx_path, 'wb') as f:
                f.write(payload)

    return fixed_total, mainseq_fixes


def _remove_printer_rels(rels_xml_bytes: bytes) -> tuple[bytes, int]:
    """
    Rimuove dai presentation.xml.rels i <Relationship> di tipo
    printerSettings. Restituisce (new_xml, n_rimossi).
    """
    PRINTER_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"
    PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    try:
        root = etree.fromstring(rels_xml_bytes)
    except etree.XMLSyntaxError:
        return rels_xml_bytes, 0
    n = 0
    for rel in list(root.findall(f'{{{PKG_NS}}}Relationship')):
        if rel.get('Type') == PRINTER_RELTYPE:
            root.remove(rel)
            n += 1
    if n == 0:
        return rels_xml_bytes, 0
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                          standalone=True), n


def _fix_empty_mainseq_in_slide_xml(xml_bytes: bytes) -> tuple[bytes, bool]:
    """
    Sostituisce <p:childTnLst/> vuoto sotto un cTn nodeType="mainSeq"
    con <p:childTnLst><p:par><p:cTn id="3"/></p:par></p:childTnLst>,
    e rinumera l'eventuale audio cTn con id=3 a id=4 per evitare
    collisioni (il dummy par usa id=3, l'audio cMediaNode usa id=4).

    PowerPoint moderno considera <p:childTnLst/> vuoto in mainSeq una
    corruzione e mostra l'avviso "PowerPoint non è riuscito a leggere
    parte del contenuto" all'apertura. Applicato come post-processing
    a livello zip, così agisce DOPO Presentation.save() di python-pptx
    e non può essere sovrascritto.

    Restituisce (new_xml, modified).
    """
    NS_P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
    try:
        root = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError:
        return xml_bytes, False

    modified = False
    for ctn in root.iter(f'{NS_P}cTn'):
        if ctn.get('nodeType') != 'mainSeq':
            continue
        ctl = ctn.find(f'{NS_P}childTnLst')
        if ctl is not None and len(ctl) == 0:
            # 1. Rinumero audio cTn id=3 → id=4 per evitare collisione
            for audio in root.iter(f'{NS_P}audio'):
                for inner_ctn in audio.iter(f'{NS_P}cTn'):
                    if inner_ctn.get('id') == '3':
                        inner_ctn.set('id', '4')
            # 2. Aggiungo il placeholder par
            par = etree.SubElement(ctl, f'{NS_P}par')
            dummy = etree.SubElement(par, f'{NS_P}cTn')
            dummy.set('id', '3')
            modified = True
            break  # un solo mainSeq per slide

    if not modified:
        return xml_bytes, False
    return etree.tostring(root, xml_declaration=True,
                          encoding='UTF-8', standalone=True), True


def _retag_video_rels_to_audio(rels_xml_bytes: bytes) -> tuple[bytes, int]:
    """
    Return (new_xml_bytes, count_changed) for a single .rels file.

    Parses the rels XML, retags Type="...video" relationships pointing to
    audio files, and re-serialises. Uses lxml so attribute ordering is
    deterministic across runs.
    """
    try:
        root = etree.fromstring(rels_xml_bytes)
    except etree.XMLSyntaxError:
        # Malformed rels — leave it alone, the rest of the toolchain will
        # surface the error.
        return rels_xml_bytes, 0

    n = 0
    for rel in root.findall(f'{{{_PKG_RELS_NS}}}Relationship'):
        if rel.get('Type') != _VIDEO_RELTYPE:
            continue
        target = rel.get('Target', '')
        if '.' not in target:
            continue
        ext = target.rsplit('.', 1)[-1].lower().split('?')[0]
        if ext in _AUDIO_EXTENSIONS:
            rel.set('Type', _AUDIO_RELTYPE)
            n += 1

    if n == 0:
        return rels_xml_bytes, 0

    new_bytes = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    return new_bytes, n


def _resolve_package_target(source_part: str, target: str) -> str:
    """Risoluzione di un Target OOXML rispetto alla parte sorgente."""
    if target.startswith('/'):
        return target.lstrip('/')
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(source_part), target)
    )


def _ordered_slide_parts(zin: zipfile.ZipFile) -> list[str]:
    """Restituisce i nomi delle slide nell'ordine reale della presentazione."""
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    presentation = etree.fromstring(zin.read('ppt/presentation.xml'))
    rels = etree.fromstring(zin.read('ppt/_rels/presentation.xml.rels'))
    rel_by_id = {
        rel.get('Id'): rel.get('Target', '')
        for rel in rels.findall(f'{{{pkg_ns}}}Relationship')
    }
    result = []
    for sld_id in presentation.findall(
        f'.//{{{p_ns}}}sldIdLst/{{{p_ns}}}sldId'
    ):
        rel_id = sld_id.get(f'{{{r_ns}}}id')
        target = rel_by_id.get(rel_id, '')
        if target:
            result.append(_resolve_package_target('ppt/presentation.xml', target))
    return result


def verify_pptx_audio_package(
    pptx_path: str,
    expected_slide_numbers: set[int],
    autoplay: bool = True,
) -> dict[int, str]:
    """Verifica gli audio direttamente nel pacchetto OOXML salvato.

    Per ogni slide prevista controlla: shape ``a:audioFile``, relationship di
    tipo audio, Media part presente e MP3 decodificabile. Con autoplay attivo
    controlla inoltre che esista il nodo ``p:audio`` nel timing della slide.
    Restituisce ``{numero_slide: motivo}``; un dict vuoto significa verifica OK.
    """
    issues: dict[int, str] = {}
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

    try:
        with zipfile.ZipFile(pptx_path, 'r') as zin:
            bad_zip = zin.testzip()
            if bad_zip:
                return {n: f"pacchetto ZIP corrotto: {bad_zip}"
                        for n in expected_slide_numbers}
            slide_parts = _ordered_slide_parts(zin)
            names = set(zin.namelist())

            for slide_num in sorted(expected_slide_numbers):
                if slide_num < 1 or slide_num > len(slide_parts):
                    issues[slide_num] = "parte slide non trovata"
                    continue
                slide_part = slide_parts[slide_num - 1]
                if slide_part not in names:
                    issues[slide_num] = f"parte mancante: {slide_part}"
                    continue

                try:
                    slide_root = etree.fromstring(zin.read(slide_part))
                except Exception as exc:
                    issues[slide_num] = f"XML slide non leggibile: {exc}"
                    continue

                audio_nodes = slide_root.findall(f'.//{{{a_ns}}}audioFile')
                if len(audio_nodes) != 1:
                    issues[slide_num] = (
                        f"attesa 1 shape audio, trovate {len(audio_nodes)}"
                    )
                    continue

                rel_id = (
                    audio_nodes[0].get(f'{{{r_ns}}}link')
                    or audio_nodes[0].get(f'{{{r_ns}}}embed')
                )
                if not rel_id:
                    issues[slide_num] = "shape audio senza relationship"
                    continue

                rels_part = posixpath.join(
                    posixpath.dirname(slide_part), '_rels',
                    posixpath.basename(slide_part) + '.rels',
                )
                if rels_part not in names:
                    issues[slide_num] = "file relationships della slide mancante"
                    continue
                try:
                    rels_root = etree.fromstring(zin.read(rels_part))
                except Exception as exc:
                    issues[slide_num] = f"relationships non leggibili: {exc}"
                    continue

                rel = None
                for candidate in rels_root.findall(
                    f'{{{pkg_ns}}}Relationship'
                ):
                    if candidate.get('Id') == rel_id:
                        rel = candidate
                        break
                if rel is None:
                    issues[slide_num] = f"relationship {rel_id} non trovata"
                    continue
                if rel.get('Type') != _AUDIO_RELTYPE:
                    issues[slide_num] = (
                        "relationship media non classificata come audio"
                    )
                    continue

                media_part = _resolve_package_target(
                    slide_part, rel.get('Target', '')
                )
                if media_part not in names:
                    issues[slide_num] = f"Media part mancante: {media_part}"
                    continue
                media_bytes = zin.read(media_part)
                if len(media_bytes) < MIN_AUDIO_FILE_BYTES:
                    issues[slide_num] = (
                        f"Media part troppo piccola ({len(media_bytes)} byte)"
                    )
                    continue
                try:
                    duration = float(MP3(BytesIO(media_bytes)).info.length)
                    if (not math.isfinite(duration)
                            or duration < MIN_AUDIO_DURATION_S):
                        raise ValueError(f"durata {duration!r}s")
                except Exception as exc:
                    issues[slide_num] = f"Media part MP3 non valida: {exc}"
                    continue

                if autoplay and slide_root.find(f'.//{{{p_ns}}}audio') is None:
                    issues[slide_num] = "timing autoplay audio mancante"

    except Exception as exc:
        return {
            n: f"impossibile verificare il PPTX: {type(exc).__name__}: {exc}"
            for n in expected_slide_numbers
        }
    return issues


def _remove_existing_audio_shapes(slide) -> int:
    """
    Rimuove dalla slide tutte le shape <p:pic> che contengono <a:audioFile>.
    Restituisce il numero di shape rimosse.

    Necessario quando un PPTX ha shape audio "orfane" (es. dopo un fix
    manuale, o copiate da un altro PPTX): senza questa pulizia
    add_audio_to_slide aggiungerebbe una SECONDA shape audio sopra,
    ma il <p:timing> minimale ha un solo <p:spTgt> → uno dei due audio
    parte in autoplay, l'altro resta orfano.

    LIMITAZIONE NOTA: la rigenerazione di un PPTX già prodotto da Slide Narrator
    (cioè con MediaParts MP3 attive nel package) NON è supportata: python-pptx
    ha un bug in `MediaParts._find_by_sha1` che crasha con AttributeError
    "'Part' object has no attribute 'sha1'" quando si chiama add_movie su un
    package che già contiene media. Per cambiare voce, rigenerare sempre
    dall'INPUT originale, non dall'output di una run precedente.
    """
    sld = slide.element
    audio_file_tag = qn('a:audioFile')
    pic_tag = qn('p:pic')

    pics_to_remove = []
    for pic in sld.iter(pic_tag):
        if pic.find('.//' + audio_file_tag) is not None:
            pics_to_remove.append(pic)

    for pic in pics_to_remove:
        parent = pic.getparent()
        if parent is not None:
            parent.remove(pic)
    return len(pics_to_remove)


def _insert_timing_in_schema_order(slide_root: etree._Element, timing: etree._Element) -> None:
    """Inserisce p:timing prima di extLst, rispettando l'ordine dello schema."""
    ext_lst = slide_root.find(qn('p:extLst'))
    if ext_lst is not None:
        ext_lst.addprevious(timing)
    else:
        slide_root.append(timing)


def _remove_audio_timing_nodes(slide_root: etree._Element, shape_ids: set[str] | None = None) -> int:
    """Rimuove solo i nodi timing audio riferiti alle shape indicate."""
    removed = 0
    for audio in list(slide_root.findall('.//' + qn('p:audio'))):
        targets = {n.get('spid') for n in audio.findall('.//' + qn('p:spTgt'))}
        if shape_ids is None or not targets or targets & shape_ids:
            parent = audio.getparent()
            if parent is not None:
                parent.remove(audio)
                removed += 1
    return removed


def add_audio_to_slide(
    slide,
    audio_path: str,
    autoplay: bool = True,
    slide_num: int | None = None,
) -> None:
    """Inserisce un audio preservando animazioni e transizioni preesistenti."""
    sld = slide.element
    prefix = f"   slide {slide_num}: " if slide_num is not None else "   "

    # Salva l'albero timing originario prima che python-pptx.add_movie aggiunga
    # il proprio blocco media-call. La copia viene ripristinata e poi arricchita
    # con il solo nodo autoplay dell'audio nuovo.
    original_timing = sld.find(qn('p:timing'))
    original_timing_copy = copy.deepcopy(original_timing) if original_timing is not None else None

    # In una rigenerazione rimuove le shape audio precedenti e i relativi nodi
    # di timing, senza toccare le altre animazioni della slide.
    old_shape_ids = set()
    for audio_file in sld.findall('.//' + qn('a:audioFile')):
        pic = audio_file
        while pic is not None and pic.tag != qn('p:pic'):
            pic = pic.getparent()
        if pic is not None:
            c_nv_pr = pic.find('.//' + qn('p:cNvPr'))
            if c_nv_pr is not None and c_nv_pr.get('id'):
                old_shape_ids.add(c_nv_pr.get('id'))
    n_removed = _remove_existing_audio_shapes(slide)
    if n_removed:
        print(f"{prefix}! sostituiti {n_removed} audio esistenti; animazioni preservate")
    if original_timing_copy is not None and old_shape_ids:
        _remove_audio_timing_nodes(original_timing_copy, old_shape_ids)

    off_slide = Emu(ICON_OFFSCREEN_EMU)
    ext = Path(audio_path).suffix.lower()
    mime_type = SUPPORTED_AUDIO_MIME_TYPES.get(ext, 'audio/mpeg')
    movie_shape = slide.shapes.add_movie(
        audio_path,
        left=off_slide, top=off_slide,
        width=Inches(ICON_SIZE_INCHES), height=Inches(ICON_SIZE_INCHES),
        mime_type=mime_type,
    )
    _convert_video_ref_to_audio_ref(movie_shape)

    # Scarta esclusivamente il timing iniettato da add_movie e ripristina quello
    # originale. In questo modo animazioni, trigger e sequenze rimangono intatti.
    current_timing = sld.find(qn('p:timing'))
    if current_timing is not None:
        sld.remove(current_timing)
    if original_timing_copy is not None:
        _insert_timing_in_schema_order(sld, original_timing_copy)

    if autoplay:
        _append_autoplay_to_slide_xml(sld, movie_shape.shape_id)

def _timing_has_animations(timing_el) -> bool:
    """Determina se un <p:timing> contiene animazioni custom non banali.

    Un timing è "vuoto" se la sua mainSeq (o equivalente) non ha figli con
    contenuto reale: un childTnLst vuoto, oppure un childTnLst contenente
    solo un mainSeq vuoto è considerato non-content.
    """
    ctn_root = timing_el.find(
        qn('p:tnLst') + '/' + qn('p:par') + '/' + qn('p:cTn')
    )
    if ctn_root is None:
        return False
    child_tn_lst = ctn_root.find(qn('p:childTnLst'))
    if child_tn_lst is None or len(child_tn_lst) == 0:
        return False
    for child in child_tn_lst:
        # <p:audio> di per sé non è un'animazione utente: è il blocco autoplay.
        if child.tag.endswith('}audio'):
            continue
        # Per <p:seq> con mainSeq: vedo se mainSeq ha figli reali.
        seq_ctn = child.find(qn('p:cTn'))
        if seq_ctn is not None:
            seq_child_lst = seq_ctn.find(qn('p:childTnLst'))
            if seq_child_lst is not None and len(seq_child_lst) > 0:
                return True
        elif len(child) > 0:
            return True
    return False


# ----------------------------------------------------------------------------
# Main pipeline
# ----------------------------------------------------------------------------
def _count_existing_audio_shapes_in_pptx(prs) -> int:
    """Conta shape audio (a:audioFile) presenti nel PPTX di input.
    Usato come check difensivo: rigenerare un PPTX già processato non è
    supportato a causa di un bug python-pptx (vedi docstring di
    _remove_existing_audio_shapes)."""
    n = 0
    audio_file_tag = qn('a:audioFile')
    pic_tag = qn('p:pic')
    for slide in prs.slides:
        for pic in slide._element.iter(pic_tag):
            if pic.find('.//' + audio_file_tag) is not None:
                n += 1
    return n


def _load_inputs(
    input_pptx: str,
    scripts_xlsx: str | None,
    cb,
    *,
    script_source: str = "excel",
    sheet_name: str | None = None,
    script_column: str = "A",
    has_header: bool = False,
) -> tuple:
    """Carica il PowerPoint e gli script con controlli di copertura.

    ``script_source`` può essere ``excel`` oppure ``notes``. Per Excel sono
    supportati foglio, colonna, intestazione e il modello avanzato con numero o
    ID stabile della slide. Le formule prive di valore memorizzato vengono
    segnalate esplicitamente da :func:`read_scripts_from_xlsx`.
    """
    cb({"stage": "loading"})
    print(f"-> Caricamento {input_pptx}")
    prs = Presentation(input_pptx)
    n_slides = len(prs.slides)

    n_existing_audio = _count_existing_audio_shapes_in_pptx(prs)
    if n_existing_audio > 0:
        raise ValueError(
            f"Il PPTX di input contiene già {n_existing_audio} shape audio. "
            "Usare la modalità Completa/Ripara audio (Fix), che preserva o "
            "sostituisce selettivamente gli audio già incorporati."
        )

    source = (script_source or "excel").strip().lower()
    metadata = get_pptx_slide_metadata(input_pptx)
    if source in {"notes", "note", "note relatore", "note powerpoint", "speaker_notes"}:
        print("-> Lettura script dalle Note PowerPoint")
        scripts = read_scripts_from_notes(input_pptx)
        source_description = "Note PowerPoint"
    else:
        if not scripts_xlsx:
            raise ValueError("Selezionare il file Excel contenente gli script.")
        print(f"-> Lettura script da {scripts_xlsx}")
        scripts = read_scripts_from_xlsx(
            scripts_xlsx,
            sheet_name=sheet_name,
            column=script_column,
            has_header=has_header,
            slide_metadata=metadata,
        )
        source_description = (
            f"{scripts_xlsx}, foglio {sheet_name or 'attivo'}, "
            f"colonna {script_column}"
        )

    n_non_empty = sum(1 for text in scripts if text)
    if n_non_empty == 0:
        raise ValueError(
            f"Nessuno script trovato nelle {source_description}: il file di "
            "output sarebbe privo di narrazione."
        )
    n_scripts = len(scripts)
    print(f"   {n_slides} slide | {n_scripts} righe script")

    missing_script_slides = []
    for index in range(n_slides):
        text = scripts[index] if index < n_scripts else ""
        if not text:
            missing_script_slides.append(index + 1)
    if missing_script_slides:
        preview = ", ".join(map(str, missing_script_slides[:20]))
        if len(missing_script_slides) > 20:
            preview += ", ..."
        raise ValueError(
            "Manca lo script audio per le slide: " + preview + ". "
            "Inserire il testo nella sorgente selezionata; per una slide "
            "intenzionalmente muta scrivere [SENZA AUDIO] o [NO AUDIO]."
        )
    if n_scripts > n_slides:
        print(
            f"   ! attenzione: {n_scripts} righe ma solo {n_slides} slide — "
            "le righe in eccesso vengono ignorate"
        )
    estimate = estimate_scripts(scripts)
    print(
        f"   stima: {estimate['words']} parole, circa "
        f"{estimate['estimated_duration']} di parlato a velocità standard"
    )
    return prs, scripts, n_slides, n_scripts

def _build_slide_jobs(
    scripts: list[str], n_slides: int, n_scripts: int, tmp_dir: str,
) -> list[tuple[int, str, str] | None]:
    """Costruisce la lista (allineata per indice) dei job da sintetizzare.
    Per ogni slide: tupla (slide_num, text, audio_path) oppure None se la
    slide non ha script."""
    slide_jobs: list[tuple[int, str, str] | None] = []
    for i in range(n_slides):
        slide_num = i + 1
        text = scripts[i] if i < n_scripts else ""
        if text.upper() in SILENT_SCRIPT_MARKERS:
            print(f"   slide {slide_num}: marcata intenzionalmente senza audio")
            slide_jobs.append(None)
        else:
            audio_path = os.path.join(tmp_dir, f"slide_{slide_num:03d}.mp3")
            slide_jobs.append((slide_num, text, audio_path))
    return slide_jobs


def _run_synthesis_phase(
    synth_tasks: list[tuple[int, str, str]],
    voice: str,
    rate: str,
    concurrency: int,
    transcode_audio: bool,
    cb,
    volume: str = "+0%",
    pitch: str = "+0Hz",
    cancel_event=None,
) -> tuple[dict[int, str], dict[int, dict]]:
    """Sintesi parallela. Restituisce (synthesis_errors, captured_timings).
    captured_timings viene riempito per ogni slide con successo."""
    synthesis_errors: dict[int, str] = {}
    captured_timings: dict[int, dict] = {}
    if not synth_tasks:
        return synthesis_errors, captured_timings

    n_tasks = len(synth_tasks)
    actual_concurrency = min(concurrency, n_tasks)
    print(
        f"-> Sintesi audio in parallelo: {n_tasks} slide, "
        f"fino a {actual_concurrency} contemporanee"
    )
    cb({"stage": "synthesis_start", "total": n_tasks})
    t0 = time.monotonic()
    results = synthesize_many_to_files(
        synth_tasks, voice, rate, volume, pitch, actual_concurrency,
        progress_callback=cb,
        captured_timings=captured_timings,
        transcode_audio=transcode_audio, cancel_event=cancel_event,
    )
    elapsed = time.monotonic() - t0
    for slide_num, ok, err in results:
        if not ok:
            synthesis_errors[slide_num] = err
    n_ok = n_tasks - len(synthesis_errors)
    avg = elapsed / n_tasks if n_tasks else 0
    print(
        f"   sintesi completata in {format_duration(elapsed)} "
        f"({n_ok}/{n_tasks} ok, {avg:.1f}s media per slide)"
    )
    if synthesis_errors:
        print(f"   ! {len(synthesis_errors)} slide con errori: "
              f"{sorted(synthesis_errors.keys())}")
    cb({"stage": "synthesis_end"})
    return synthesis_errors, captured_timings


def _clone_pool_init(threads_per_worker: int) -> None:
    """Inizializzatore dei processi worker: limita i thread di calcolo PRIMA che
    torch/numpy vengano importati (nel worker succede al primo caricamento del
    modello). Con N worker su M core, ~M/N thread ciascuno evita che i processi
    si contendano gli stessi core."""
    import os as _os
    v = str(max(1, int(threads_per_worker)))
    for name in ("OMP_NUM_THREADS", "MKL_NUM_THREADS", "OPENBLAS_NUM_THREADS",
                 "NUMEXPR_NUM_THREADS", "VECLIB_MAXIMUM_THREADS"):
        _os.environ[name] = v


def _synthesize_clone_with_retries(
    cloned_voice, text: str, rate: str, output_path: str,
    pocket_variant: str, use_cache: bool, pocket_quantize: bool,
):
    """Sintesi locale atomica con retry e validazione MP3."""
    last_error = "errore sconosciuto"
    for attempt in range(1, AUDIO_GENERATION_ATTEMPTS + 1):
        temp_path = output_path + ".part.mp3"
        _remove_file_quietly(temp_path)
        _remove_file_quietly(output_path)
        try:
            result = voice_clone.synthesize_clone(
                cloned_voice, text, rate, temp_path,
                pocket_variant=pocket_variant,
                use_cache=use_cache,
                pocket_quantize=pocket_quantize,
            )
            valid, reason = validate_audio_file(temp_path)
            if not valid:
                raise RuntimeError(reason)
            os.replace(temp_path, output_path)
            valid, reason = validate_audio_file(output_path)
            if not valid:
                raise RuntimeError(f"controllo finale fallito: {reason}")
            return result, attempt
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            _remove_file_quietly(temp_path)
            _remove_file_quietly(output_path)
            if attempt < AUDIO_GENERATION_ATTEMPTS:
                time.sleep(AUDIO_RETRY_BASE_DELAY_S * attempt)
    raise RuntimeError(
        f"{last_error} (dopo {AUDIO_GENERATION_ATTEMPTS} tentativi)"
    )


def _clone_synth_worker(payload):
    """Sintesi di UNA slide dentro un processo worker. Ritorna una tupla
    picklable così il risultato può tornare al processo padre.
    payload = (slide_num, text, path, voice, rate, pocket_variant, use_cache,
               pocket_quantize)."""
    (slide_num, text, path, voice, rate, pocket_variant, use_cache,
     pocket_quantize) = payload
    try:
        result, attempt = _synthesize_clone_with_retries(
            voice, text, rate, path, pocket_variant, use_cache,
            pocket_quantize,
        )
        return (slide_num, True, result.sentences, None, attempt)
    except Exception as e:
        return (slide_num, False, None, f"{type(e).__name__}: {e}",
                AUDIO_GENERATION_ATTEMPTS)


def _prep_voice_for_workers(cloned_voice, pocket_variant: str):
    """Pre-accorcia una volta sola il riferimento nel processo PADRE, così i
    worker non lo ricalcolano in contemporanea (che creerebbe una race sul file
    temporaneo) e non caricano nulla di pesante. Restituisce una copia della
    voce con il reference_path già a norma; in caso di problemi torna l'originale."""
    try:
        import dataclasses
        engine = cloned_voice.engine
        variant = pocket_variant if engine == voice_library.ENGINE_POCKET else None
        backend = voice_clone.get_backend(engine, variant)
        capped = backend._reference_for(cloned_voice)  # solo misura + ffmpeg, no modello
        if capped and capped != cloned_voice.reference_path:
            return dataclasses.replace(cloned_voice, reference_path=capped)
    except Exception:
        pass
    return cloned_voice


def _run_synthesis_phase_clone_parallel(
    synth_tasks, cloned_voice, rate, cb, pocket_variant, workers,
    use_cache=True, pocket_quantize=False, worker_fn=None,
):
    """Come _run_synthesis_phase_clone ma distribuisce le slide su più processi.
    Restituisce (synthesis_errors, captured_timings) IDENTICI alla versione
    sequenziale, così il resto della pipeline non cambia. `worker_fn` è
    iniettabile per i test (di default usa _clone_synth_worker)."""
    import concurrent.futures as _cf

    synthesis_errors: dict[int, str] = {}
    captured_timings: dict[int, dict] = {}
    if not synth_tasks:
        return synthesis_errors, captured_timings

    n_tasks = len(synth_tasks)
    cores = os.cpu_count() or 4
    workers = max(1, min(int(workers), n_tasks, MAX_CLONE_WORKERS))
    threads_per_worker = max(1, cores // workers)
    text_by_num = {sn: text for (sn, text, _p) in synth_tasks}

    voice_for_workers = _prep_voice_for_workers(cloned_voice, pocket_variant)

    print(f"-> Sintesi audio (voce clonata '{cloned_voice.name}' / "
          f"{cloned_voice.engine_label}): {n_tasks} slide, {workers} in "
          f"parallelo ({threads_per_worker} thread per processo)")
    cb({"stage": "synthesis_start", "total": n_tasks})
    t0 = time.monotonic()

    fn = worker_fn or _clone_synth_worker
    payloads = [(sn, text, path, voice_for_workers, rate, pocket_variant,
                 use_cache, pocket_quantize)
                for (sn, text, path) in synth_tasks]

    done = 0
    with _cf.ProcessPoolExecutor(
        max_workers=workers,
        initializer=_clone_pool_init,
        initargs=(threads_per_worker,),
    ) as ex:
        futures = {ex.submit(fn, p): p[0] for p in payloads}
        for fut in _cf.as_completed(futures):
            sn = futures[fut]
            try:
                worker_result = fut.result()
                # Compatibilità con eventuali test/estensioni che usavano il
                # vecchio contratto a 4 elementi prima dell'aggiunta del numero
                # di tentativo.
                if len(worker_result) == 4:
                    r_sn, ok, sentences, error = worker_result
                    attempt = 1
                else:
                    r_sn, ok, sentences, error, attempt = worker_result
            except Exception as e:  # crash del worker (raro): lo tratto come errore slide
                r_sn, ok, sentences, error, attempt = (
                    sn, False, None, f"{type(e).__name__}: {e}",
                    AUDIO_GENERATION_ATTEMPTS,
                )
            if ok:
                captured_timings[r_sn] = {
                    "text": text_by_num.get(r_sn, ""),
                    "sentences": sentences,
                }
                retry_note = "" if attempt == 1 else f", tentativo {attempt}"
                print(f"   slide {r_sn}: sintesi ok "
                      f"({len(text_by_num.get(r_sn, ''))} caratteri, "
                      f"{len(sentences)} frasi{retry_note})")
            else:
                synthesis_errors[r_sn] = error
                print(f"   slide {r_sn}: ERRORE sintesi — {error}")
            done += 1
            cb({"stage": "synthesis_progress", "done": done, "total": n_tasks,
                "slide_num": r_sn, "ok": ok})

    elapsed = time.monotonic() - t0
    n_ok = n_tasks - len(synthesis_errors)
    avg = elapsed / n_tasks if n_tasks else 0
    print(f"   sintesi completata in {format_duration(elapsed)} "
          f"({n_ok}/{n_tasks} ok, {avg:.1f}s media per slide, {workers} processi)")
    if synthesis_errors:
        print(f"   ! {len(synthesis_errors)} slide con errori: "
              f"{sorted(synthesis_errors.keys())}")
    cb({"stage": "synthesis_end"})
    return synthesis_errors, captured_timings


def _run_synthesis_phase_clone(
    synth_tasks: list[tuple[int, str, str]],
    cloned_voice,
    rate: str,
    cb,
    pocket_variant: str = "italian",
    workers: int = 1,
    use_cache: bool = True,
    pocket_quantize: bool = False,
) -> tuple[dict[int, str], dict[int, dict]]:
    """
    Sintesi con una voce clonata (motore locale, CPU-bound). Con workers>1
    distribuisce le slide su più processi; con workers<=1 resta sequenziale.
    Restituisce (synthesis_errors, captured_timings) con la stessa forma della
    fase edge, così il resto della pipeline non cambia. Le voci sono già a
    livello di frase: captured_timings[n] = {"text","sentences"}.
    """
    if workers and workers > 1:
        return _run_synthesis_phase_clone_parallel(
            synth_tasks, cloned_voice, rate, cb, pocket_variant, workers,
            use_cache=use_cache, pocket_quantize=pocket_quantize)

    synthesis_errors: dict[int, str] = {}
    captured_timings: dict[int, dict] = {}
    if not synth_tasks:
        return synthesis_errors, captured_timings

    n_tasks = len(synth_tasks)
    print(f"-> Sintesi audio (voce clonata '{cloned_voice.name}' / "
          f"{cloned_voice.engine_label}): {n_tasks} slide, una alla volta")
    cb({"stage": "synthesis_start", "total": n_tasks})
    t0 = time.monotonic()
    done = 0
    for slide_num, text, path in synth_tasks:
        ts = time.monotonic()
        try:
            result, attempt = _synthesize_clone_with_retries(
                cloned_voice, text, rate, path, pocket_variant, use_cache,
                pocket_quantize,
            )
            captured_timings[slide_num] = {
                "text": text,
                "sentences": result.sentences,
            }
            elapsed = time.monotonic() - ts
            retry_note = "" if attempt == 1 else f", tentativo {attempt}"
            print(f"   slide {slide_num}: sintesi ok ({elapsed:.1f}s, "
                  f"{len(text)} caratteri, {len(result.sentences)} frasi"
                  f"{retry_note})")
            ok = True
        except Exception as e:
            synthesis_errors[slide_num] = str(e)
            print(f"   slide {slide_num}: ERRORE sintesi — {e}")
            ok = False
        done += 1
        cb({"stage": "synthesis_progress", "done": done, "total": n_tasks,
            "slide_num": slide_num, "ok": ok})

    elapsed = time.monotonic() - t0
    n_ok = n_tasks - len(synthesis_errors)
    avg = elapsed / n_tasks if n_tasks else 0
    print(f"   sintesi completata in {format_duration(elapsed)} "
          f"({n_ok}/{n_tasks} ok, {avg:.1f}s media per slide)")
    if synthesis_errors:
        print(f"   ! {len(synthesis_errors)} slide con errori: "
              f"{sorted(synthesis_errors.keys())}")
    cb({"stage": "synthesis_end"})
    return synthesis_errors, captured_timings


def _audit_synthesized_audio(
    slide_jobs: list[tuple[int, str, str] | None],
) -> dict[int, str]:
    """Restituisce le slide previste il cui MP3 manca o non è valido."""
    issues: dict[int, str] = {}
    for job in slide_jobs:
        if job is None:
            continue
        slide_num, _text, audio_path = job
        valid, reason = validate_audio_file(audio_path)
        if not valid:
            issues[slide_num] = reason
    return issues


def _audio_shape_count(slide) -> int:
    audio_file_tag = qn('a:audioFile')
    pic_tag = qn('p:pic')
    return sum(
        1 for pic in slide._element.iter(pic_tag)
        if pic.find('.//' + audio_file_tag) is not None
    )


def _run_embedding_phase(
    prs,
    slide_jobs: list[tuple[int, str, str] | None],
    synthesis_errors: dict[int, str],
    autoplay: bool,
    cb,
    auto_advance: bool = False,
    cancel_event=None,
) -> list[tuple[int, float | None]]:
    """Incorpora tutti gli audio e verifica una shape audio per ogni slide."""
    print("-> Incorporamento audio nelle slide con controllo di integrità")
    n_synth_tasks = sum(1 for j in slide_jobs if j is not None)
    n_to_embed = n_synth_tasks - len(synthesis_errors)
    durations: list[tuple[int, float | None]] = []
    embed_done = 0
    embedding_errors: dict[int, str] = {}

    for i, slide in enumerate(prs.slides):
        _check_cancelled(cancel_event)
        slide_num = i + 1
        job = slide_jobs[i]
        if job is None:
            durations.append((slide_num, None))
            continue
        if slide_num in synthesis_errors:
            embedding_errors[slide_num] = (
                "audio non disponibile: " + synthesis_errors[slide_num]
            )
            durations.append((slide_num, None))
            continue

        _, _, audio_path = job
        valid, reason = validate_audio_file(audio_path)
        if not valid:
            embedding_errors[slide_num] = reason
            durations.append((slide_num, None))
            continue

        duration = get_audio_duration_seconds(audio_path)
        durations.append((slide_num, duration))
        last_error = "errore sconosciuto"
        inserted = False
        for attempt in range(1, AUDIO_EMBED_ATTEMPTS + 1):
            try:
                print(
                    f"   slide {slide_num}: durata {format_duration(duration)}, "
                    f"incorporo audio"
                    + ("" if attempt == 1 else f" (tentativo {attempt})")
                )
                add_audio_to_slide(
                    slide, audio_path, autoplay=autoplay, slide_num=slide_num
                )
                count = _audio_shape_count(slide)
                if count != 1:
                    raise RuntimeError(
                        f"attese 1 shape audio, trovate {count}"
                    )
                if auto_advance and autoplay:
                    _set_slide_auto_advance(slide, duration)
                inserted = True
                break
            except Exception as exc:
                last_error = f"{type(exc).__name__}: {exc}"
                print(
                    f"   slide {slide_num}: inserimento fallito "
                    f"({attempt}/{AUDIO_EMBED_ATTEMPTS}) — {exc}"
                )

        if not inserted:
            embedding_errors[slide_num] = last_error
            continue

        embed_done += 1
        if n_to_embed > 0:
            cb({
                "stage": "embedding_progress",
                "done": embed_done,
                "total": n_to_embed,
            })

    if embedding_errors:
        details = "; ".join(
            f"slide {n}: {err}" for n, err in sorted(embedding_errors.items())
        )
        raise RuntimeError(
            "Impossibile incorporare tutti gli audio. " + details
        )
    return durations


def _captions_sidecar_path(pptx_path: str | Path) -> Path:
    p = Path(pptx_path)
    return p.with_name(f"{p.stem}_captions.json")


def _load_existing_captions(pptx_path: str | Path) -> dict | None:
    path = _captions_sidecar_path(pptx_path)
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, dict) and isinstance(data.get("slides", {}), dict):
            return data
    except Exception as exc:
        print(f"   ! sidecar caption esistente non leggibile: {exc}")
    return None


def _write_captions_json(
    output_pptx: str,
    voice: str,
    rate: str,
    captured_timings: dict[int, dict],
    existing_data: dict | None = None,
    replaced_slides: set[int] | None = None,
    known_audio_slides: set[int] | None = None,
) -> None:
    """Scrive sempre un sidecar atomico e, nel Fix, unisce i dati esistenti."""
    captions_json_path = _captions_sidecar_path(output_pptx)
    base_slides = {}
    if existing_data:
        base_slides = dict(existing_data.get("slides", {}))
    for slide_num in replaced_slides or set():
        base_slides.pop(str(slide_num), None)

    n_new = 0
    for slide_num, payload in captured_timings.items():
        if "sentences" in payload:
            sentences = payload["sentences"]
        else:
            sentences = _sentences_with_timing(
                payload["text"], payload.get("words", [])
            )
        if sentences:
            base_slides[str(slide_num)] = {
                "text": payload["text"],
                "sentences": sentences,
            }
            n_new += 1

    known = set(known_audio_slides or set())
    captioned = {int(k) for k in base_slides if str(k).isdigit()}
    missing_existing = sorted(known - captioned)
    captions_data = {
        "version": 2,
        "voice": voice,
        "rate": rate,
        "complete": not missing_existing,
        "slides_without_caption_data": missing_existing,
        "slides": {k: base_slides[k] for k in sorted(base_slides, key=lambda x: int(x) if str(x).isdigit() else 10**9)},
    }
    _atomic_write_text(
        captions_json_path,
        json.dumps(captions_data, ensure_ascii=False, indent=2),
    )
    if base_slides:
        print(
            f"-> Captions sincronizzate: {captions_json_path} "
            f"({len(base_slides)} slide, {n_new} aggiornate)"
        )
    else:
        print(f"-> Sidecar captions aggiornato (nessuna caption): {captions_json_path}")
    if missing_existing:
        print(
            "   ! caption non disponibili per audio preesistenti nelle slide: "
            + ", ".join(map(str, missing_existing))
        )


def _write_duration_report(
    output_pptx: str,
    report_path: str | None,
    durations: list[tuple[int, float | None]],
) -> str:
    """Salva il report .txt con le durate per slide. Restituisce il path."""
    out_p = Path(output_pptx)
    if report_path is None:
        report_path = str(out_p.with_name(f"{out_p.stem}_durate.txt"))

    total_seconds = sum(d for _, d in durations if d is not None)
    with open(report_path, "w", encoding="utf-8") as f:
        for slide_num, duration in durations:
            if duration is None:
                f.write(f"[{slide_num}] Slide: ---\n")
            else:
                f.write(f"[{slide_num}] Slide: {format_duration(duration)}\n")
        f.write(f"\nTotale: {format_duration(total_seconds)}\n")

    print(f"-> Report durate: {report_path}")
    print(f"   durata totale: {format_duration(total_seconds)}")
    return report_path


def _resolve_sentences(payload: dict) -> list[dict]:
    """Restituisce le frasi con timing per una slide, indipendentemente dal
    backend: le voci clonate le forniscono già ("sentences"), quelle Microsoft
    le ricavano dai WordBoundary ("words")."""
    if not payload:
        return []
    if "sentences" in payload:
        return payload["sentences"]
    return _sentences_with_timing(payload.get("text", ""), payload.get("words", []))


def _export_video(
    input_pptx: str,
    output_mp4: str,
    slide_jobs: list,
    synthesis_errors: dict[int, str],
    captured_timings: dict[int, dict],
    resolution: str,
    subtitles: bool,
    transition: bool,
    report_path: str | None,
    work_dir: str,
    cb,
    *,
    silent_slide_s: float = 3.0,
    transition_style: str = "fade",
    render_backend: str = "auto",
    cancel_event=None,
) -> str:
    """Stadio finale alternativo all'embedding: costruisce un video MP4 dalle
    slide renderizzate + l'audio già sintetizzato. Restituisce il path del
    report durate (stesso contratto del percorso pptx)."""
    slides: list[dict] = []
    durations: list[tuple[int, float | None]] = []
    for i, job in enumerate(slide_jobs):
        slide_num = i + 1
        if job is None or slide_num in synthesis_errors:
            slides.append({"audio": None, "duration_s": None, "sentences": []})
            durations.append((slide_num, None))
            continue
        _, _, audio_path = job
        dur = get_audio_duration_seconds(audio_path)
        sentences = _resolve_sentences(captured_timings.get(slide_num, {}))
        slides.append({"audio": audio_path, "duration_s": dur,
                       "sentences": sentences})
        durations.append((slide_num, dur))

    video_export.build_video(
        input_pptx, slides, output_mp4,
        resolution=resolution, subtitles=subtitles, transition=transition,
        work_dir=work_dir, progress_callback=cb,
        silent_slide_s=silent_slide_s,
        transition_style=transition_style,
        render_backend=render_backend,
        cancel_event=cancel_event,
    )
    report_path = _write_duration_report(output_mp4, report_path, durations)
    print("Fatto.")
    cb({"stage": "done"})
    return report_path



# ----------------------------------------------------------------------------
# Modalità FIX: completa esclusivamente gli audio mancanti in un PPTX esistente
# ----------------------------------------------------------------------------
# Icona PNG trasparente 1x1. La shape audio viene comunque collocata fuori
# dalla slide; l'immagine serve soltanto a soddisfare la struttura richiesta
# da PowerPoint per il media placeholder.
_FIX_AUDIO_ICON_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def _inspect_embedded_audio_bytes(media_bytes: bytes, media_part: str) -> tuple[bool, str, dict]:
    suffix = Path(media_part).suffix.lower() or ".bin"
    fd, tmp = tempfile.mkstemp(prefix=".slide_narrator_probe_", suffix=suffix)
    os.close(fd)
    try:
        Path(tmp).write_bytes(media_bytes)
        return inspect_audio_file(tmp, require_signal=True)
    finally:
        _remove_file_quietly(tmp)


def _parse_audio_slide_from_package(
    zin: zipfile.ZipFile,
    slide_part: str,
    names: set[str] | None = None,
) -> dict:
    """Analizza una singola slide senza modificarla.

    Stati restituiti:
      - valid: esiste esattamente un audio incorporato e decodificabile;
      - missing: nessuna shape audio e nessun timing audio orfano;
      - invalid: struttura ambigua, relazione/media mancante o audio corrotto.
    """
    # ``namelist()`` su presentazioni molto grandi è costoso se ripetuto per
    # ogni slide. L'analizzatore principale calcola l'insieme una sola volta
    # e lo passa qui; il fallback mantiene la compatibilità con le chiamate
    # dirette usate nei test e in eventuali integrazioni esterne.
    if names is None:
        names = set(zin.namelist())
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

    if slide_part not in names:
        return {"status": "invalid", "reason": f"parte mancante: {slide_part}"}
    try:
        slide_root = etree.fromstring(zin.read(slide_part))
    except Exception as exc:
        return {"status": "invalid", "reason": f"XML slide non leggibile: {exc}"}

    audio_nodes = slide_root.findall(f'.//{{{a_ns}}}audioFile')
    timing_audio_nodes = slide_root.findall(f'.//{{{p_ns}}}audio')
    if len(audio_nodes) == 0:
        if timing_audio_nodes:
            return {
                "status": "invalid",
                "reason": "timing audio presente ma shape audio mancante",
            }
        return {"status": "missing", "reason": "nessun audio incorporato"}
    if len(audio_nodes) != 1:
        return {
            "status": "invalid",
            "kind": "multiple",
            "audio_count": len(audio_nodes),
            "reason": f"rilevate {len(audio_nodes)} shape audio; selezionare la slide per sostituire la narrazione",
        }

    rel_id = (
        audio_nodes[0].get(f'{{{r_ns}}}link')
        or audio_nodes[0].get(f'{{{r_ns}}}embed')
    )
    if not rel_id:
        return {"status": "invalid", "reason": "shape audio senza relationship"}

    rels_part = posixpath.join(
        posixpath.dirname(slide_part), '_rels',
        posixpath.basename(slide_part) + '.rels',
    )
    if rels_part not in names:
        return {"status": "invalid", "reason": "relationships della slide mancanti"}
    try:
        rels_root = etree.fromstring(zin.read(rels_part))
    except Exception as exc:
        return {"status": "invalid", "reason": f"relationships non leggibili: {exc}"}

    rel = None
    for candidate in rels_root.findall(f'{{{pkg_ns}}}Relationship'):
        if candidate.get('Id') == rel_id:
            rel = candidate
            break
    if rel is None:
        return {"status": "invalid", "reason": f"relationship {rel_id} non trovata"}
    rel_type = rel.get('Type', '')
    if rel_type not in {_AUDIO_RELTYPE, _VIDEO_RELTYPE}:
        return {
            "status": "invalid",
            "reason": "relationship media non classificata come audio/video",
        }

    media_part = _resolve_package_target(slide_part, rel.get('Target', ''))
    if media_part not in names:
        return {"status": "invalid", "reason": f"Media part mancante: {media_part}"}
    media_bytes = zin.read(media_part)
    if len(media_bytes) < MIN_AUDIO_FILE_BYTES:
        return {
            "status": "invalid",
            "reason": f"Media part troppo piccola ({len(media_bytes)} byte)",
            "kind": "corrupt",
        }
    audio_ok, audio_reason, audio_details = _inspect_embedded_audio_bytes(
        media_bytes, media_part
    )
    if not audio_ok:
        return {
            "status": "invalid",
            "reason": f"audio non valido: {audio_reason}",
            "kind": "corrupt",
            "media_part": media_part,
            "format": Path(media_part).suffix.lower().lstrip('.'),
        }
    duration = float(audio_details["duration_s"])

    shape = audio_nodes[0].getparent()
    while shape is not None and shape.tag != f'{{{p_ns}}}pic':
        shape = shape.getparent()
    shape_id = None
    if shape is not None:
        c_nv_pr = shape.find(f'.//{{{p_ns}}}cNvPr')
        if c_nv_pr is not None:
            shape_id = c_nv_pr.get('id')
    autoplay = False
    if shape_id:
        autoplay = any(
            node.get('spid') == shape_id
            for node in slide_root.findall(f'.//{{{p_ns}}}spTgt')
        )
    transition = slide_root.find(f'{{{p_ns}}}transition')
    auto_advance = bool(transition is not None and transition.get('advTm'))
    icon_hidden = False
    if shape is not None:
        off = shape.find(f'.//{{{a_ns}}}off')
        if off is not None:
            try:
                icon_hidden = int(off.get('x', '0')) < 0 or int(off.get('y', '0')) < 0
            except ValueError:
                icon_hidden = False
    return {
        "status": "valid",
        "reason": "audio valido",
        "duration_s": duration,
        "media_part": media_part,
        "format": Path(media_part).suffix.lower().lstrip('.'),
        "mime_type": SUPPORTED_AUDIO_MIME_TYPES.get(Path(media_part).suffix.lower(), "audio/*"),
        "max_volume_dbfs": audio_details.get("max_volume_dbfs"),
        "autoplay": autoplay,
        "auto_advance": auto_advance,
        "icon_hidden": icon_hidden,
        "shape_id": shape_id,
        "relationship_type": rel_type,
        "audio_count": 1,
    }


def analyze_pptx_audio(
    pptx_path: str,
    progress_callback=None,
    cancel_event=None,
    *,
    deep_integrity: bool = False,
) -> dict:
    """Controlla tutte le slide di un PowerPoint già esistente.

    L'analisi può richiedere diversi secondi su presentazioni con centinaia di
    audio. ``progress_callback`` riceve eventi serializzabili con ``current``
    e ``total``; ``cancel_event`` può essere un ``threading.Event``.

    Per impostazione predefinita non viene eseguito ``ZipFile.testzip()``:
    quel controllo decomprime l'intero pacchetto e, su file molto grandi,
    raddoppia inutilmente il lavoro prima della vera analisi. Ogni parte letta
    viene comunque verificata da ``zipfile`` tramite CRC. Il controllo completo
    resta disponibile con ``deep_integrity=True``.
    """
    result = {
        "total_slides": 0,
        "valid": [],
        "missing": [],
        "invalid": {},
        "details": {},
    }
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zin:
            names = set(zin.namelist())
            required = {"[Content_Types].xml", "ppt/presentation.xml"}
            missing_required = sorted(required - names)
            if missing_required:
                raise ValueError(
                    "pacchetto PowerPoint incompleto: " + ", ".join(missing_required)
                )
            if deep_integrity:
                bad_zip = zin.testzip()
                if bad_zip:
                    raise ValueError(f"pacchetto ZIP corrotto: {bad_zip}")
            slide_parts = _ordered_slide_parts(zin)
            total = len(slide_parts)
            result["total_slides"] = total
            if progress_callback:
                progress_callback({
                    "stage": "analyze", "current": 0, "total": total,
                    "message": f"Avvio analisi di {total} slide",
                })
            for slide_num, slide_part in enumerate(slide_parts, 1):
                if cancel_event is not None and cancel_event.is_set():
                    raise OperationCancelled("Analisi annullata dall'utente")
                detail = _parse_audio_slide_from_package(zin, slide_part, names)
                result["details"][slide_num] = detail
                status = detail["status"]
                if status == "valid":
                    result["valid"].append(slide_num)
                elif status == "missing":
                    result["missing"].append(slide_num)
                else:
                    result["invalid"][slide_num] = detail.get("reason", "errore")
                if progress_callback:
                    progress_callback({
                        "stage": "analyze",
                        "current": slide_num,
                        "total": total,
                        "slide": slide_num,
                        "status": status,
                        "message": f"Analisi slide {slide_num} di {total}",
                    })
    except OperationCancelled:
        raise
    except Exception as exc:
        raise ValueError(
            f"Impossibile analizzare il PowerPoint: {type(exc).__name__}: {exc}"
        ) from exc
    return result


def _next_relationship_id(rels_root: etree._Element) -> str:
    used = {rel.get('Id') for rel in rels_root}
    n = 1
    while f"rId{n}" in used:
        n += 1
    return f"rId{n}"


def _next_package_part_name(names: set[str], preferred: str) -> str:
    if preferred not in names:
        names.add(preferred)
        return preferred
    stem, ext = posixpath.splitext(preferred)
    n = 2
    while f"{stem}_{n}{ext}" in names:
        n += 1
    chosen = f"{stem}_{n}{ext}"
    names.add(chosen)
    return chosen


def _ensure_default_content_type(
    root: etree._Element, extension: str, content_type: str,
) -> None:
    tag = f'{{{_CONTENT_TYPES_NS}}}Default'
    for item in root.findall(tag):
        if (item.get('Extension') or '').lower() == extension.lower():
            return
    node = etree.SubElement(root, tag)
    node.set('Extension', extension)
    node.set('ContentType', content_type)


def _build_audio_picture_shape_ooxml(
    shape_id: int,
    audio_rel_id: str,
    media_rel_id: str,
    image_rel_id: str,
    display_name: str,
) -> etree._Element:
    p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    pic = etree.Element(f'{{{p}}}pic')
    nv_pic_pr = etree.SubElement(pic, f'{{{p}}}nvPicPr')
    c_nv_pr = etree.SubElement(nv_pic_pr, f'{{{p}}}cNvPr')
    c_nv_pr.set('id', str(shape_id))
    c_nv_pr.set('name', display_name)
    c_nv_pic_pr = etree.SubElement(nv_pic_pr, f'{{{p}}}cNvPicPr')
    locks = etree.SubElement(c_nv_pic_pr, f'{{{a}}}picLocks')
    locks.set('noChangeAspect', '1')
    nv_pr = etree.SubElement(nv_pic_pr, f'{{{p}}}nvPr')
    audio_file = etree.SubElement(nv_pr, f'{{{a}}}audioFile')
    audio_file.set(f'{{{r}}}link', audio_rel_id)
    ext_lst = etree.SubElement(nv_pr, f'{{{p}}}extLst')
    ext = etree.SubElement(ext_lst, f'{{{p}}}ext')
    ext.set('uri', '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}')
    media = etree.SubElement(ext, f'{{{_P14_NS}}}media', nsmap={'p14': _P14_NS})
    media.set(f'{{{r}}}embed', media_rel_id)

    blip_fill = etree.SubElement(pic, f'{{{p}}}blipFill')
    blip = etree.SubElement(blip_fill, f'{{{a}}}blip')
    blip.set(f'{{{r}}}embed', image_rel_id)
    stretch = etree.SubElement(blip_fill, f'{{{a}}}stretch')
    etree.SubElement(stretch, f'{{{a}}}fillRect')

    sp_pr = etree.SubElement(pic, f'{{{p}}}spPr')
    xfrm = etree.SubElement(sp_pr, f'{{{a}}}xfrm')
    off = etree.SubElement(xfrm, f'{{{a}}}off')
    off.set('x', str(ICON_OFFSCREEN_EMU))
    off.set('y', str(ICON_OFFSCREEN_EMU))
    size = int(round(ICON_SIZE_INCHES * EMU_PER_INCH))
    ext_size = etree.SubElement(xfrm, f'{{{a}}}ext')
    ext_size.set('cx', str(size))
    ext_size.set('cy', str(size))
    geom = etree.SubElement(sp_pr, f'{{{a}}}prstGeom')
    geom.set('prst', 'rect')
    etree.SubElement(geom, f'{{{a}}}avLst')
    return pic


def _append_autoplay_to_slide_xml(
    slide_root: etree._Element, shape_id: int,
) -> None:
    """Aggiunge l'autoplay preservando le animazioni già presenti.

    Se la slide non ha un timing, usa il timing minimale già adottato dal
    motore. Se il timing esiste, aggiunge soltanto il nodo audio sotto il
    ``tmRoot`` anziché sostituire l'intero albero delle animazioni.
    """
    p_tag = lambda local: f'{{http://schemas.openxmlformats.org/presentationml/2006/main}}{local}'
    timing = slide_root.find(p_tag('timing'))
    if timing is None:
        new_timing = _build_autoplay_timing(shape_id)
        ext_lst = slide_root.find(p_tag('extLst'))
        if ext_lst is not None:
            ext_lst.addprevious(new_timing)
        else:
            slide_root.append(new_timing)
        return

    if timing.find(f'.//{p_tag("audio")}') is not None:
        raise ValueError(
            "la slide contiene già un nodo timing audio senza una shape valida"
        )
    tm_roots = [
        node for node in timing.findall(f'.//{p_tag("cTn")}')
        if node.get('nodeType') == 'tmRoot'
    ]
    if len(tm_roots) != 1:
        raise ValueError(
            "timing esistente non riconoscibile; impossibile aggiungere l'autoplay senza alterare le animazioni"
        )
    tm_root = tm_roots[0]
    child_list = tm_root.find(p_tag('childTnLst'))
    if child_list is None:
        child_list = etree.SubElement(tm_root, p_tag('childTnLst'))

    ids = []
    for node in timing.findall(f'.//{p_tag("cTn")}'):
        try:
            ids.append(int(node.get('id', '0')))
        except ValueError:
            pass
    ctn_id = max(ids or [0]) + 1
    audio = etree.SubElement(child_list, p_tag('audio'))
    media_node = etree.SubElement(audio, p_tag('cMediaNode'))
    media_node.set('showWhenStopped', '0')
    ctn = etree.SubElement(media_node, p_tag('cTn'))
    ctn.set('id', str(ctn_id))
    ctn.set('fill', 'hold')
    ctn.set('display', '0')
    st_cond = etree.SubElement(ctn, p_tag('stCondLst'))
    cond = etree.SubElement(st_cond, p_tag('cond'))
    cond.set('delay', '0')
    tgt_el = etree.SubElement(media_node, p_tag('tgtEl'))
    sp_tgt = etree.SubElement(tgt_el, p_tag('spTgt'))
    sp_tgt.set('spid', str(shape_id))


def _set_slide_auto_advance_xml(
    slide_root: etree._Element, duration_s: float,
) -> None:
    p_tag = lambda local: f'{{http://schemas.openxmlformats.org/presentationml/2006/main}}{local}'
    advance_ms = int(round(duration_s * 1000)) + AUTO_ADVANCE_TAIL_MS
    transition = slide_root.find(p_tag('transition'))
    if transition is not None:
        transition.set('advClick', '1')
        transition.set('advTm', str(advance_ms))
        return
    transition = _build_auto_advance_transition(advance_ms)
    timing = slide_root.find(p_tag('timing'))
    if timing is not None:
        timing.addprevious(transition)
        return
    anchor = slide_root.find(p_tag('clrMapOvr'))
    if anchor is None:
        anchor = slide_root.find(p_tag('cSld'))
    if anchor is not None:
        anchor.addnext(transition)
    else:
        slide_root.append(transition)



def _audio_shape_ids(slide_root: etree._Element) -> list[str]:
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    ids: list[str] = []
    for audio_file in slide_root.findall(f'.//{{{a_ns}}}audioFile'):
        pic = audio_file
        while pic is not None and pic.tag != f'{{{p_ns}}}pic':
            pic = pic.getparent()
        if pic is not None:
            c_nv_pr = pic.find(f'.//{{{p_ns}}}cNvPr')
            if c_nv_pr is not None and c_nv_pr.get('id'):
                ids.append(c_nv_pr.get('id'))
    return ids


def _remove_audio_from_slide_package(
    slide_root: etree._Element, rels_root: etree._Element,
) -> dict:
    """Rimuove shape, timing e relazioni audio di una slide, preservando il resto."""
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rel_ids: set[str] = set()
    shape_ids: set[str] = set()
    pics = []
    for audio_file in slide_root.findall(f'.//{{{a_ns}}}audioFile'):
        for attr in (f'{{{r_ns}}}link', f'{{{r_ns}}}embed'):
            if audio_file.get(attr):
                rel_ids.add(audio_file.get(attr))
        pic = audio_file
        while pic is not None and pic.tag != f'{{{p_ns}}}pic':
            pic = pic.getparent()
        if pic is not None and pic not in pics:
            pics.append(pic)
            c_nv_pr = pic.find(f'.//{{{p_ns}}}cNvPr')
            if c_nv_pr is not None and c_nv_pr.get('id'):
                shape_ids.add(c_nv_pr.get('id'))
            for node in pic.iter():
                for attr in (f'{{{r_ns}}}link', f'{{{r_ns}}}embed'):
                    if node.get(attr):
                        rel_ids.add(node.get(attr))
    for pic in pics:
        parent = pic.getparent()
        if parent is not None:
            parent.remove(pic)
    for audio in list(slide_root.findall(f'.//{{{p_ns}}}audio')):
        targets = {n.get('spid') for n in audio.findall(f'.//{{{p_ns}}}spTgt')}
        if not targets or targets & shape_ids:
            parent = audio.getparent()
            if parent is not None:
                parent.remove(audio)
    for rel in list(rels_root):
        if rel.get('Id') in rel_ids:
            rels_root.remove(rel)
    return {"removed_shapes": len(pics), "shape_ids": sorted(shape_ids), "relationship_ids": sorted(rel_ids)}


def _set_audio_autoplay_xml(slide_root: etree._Element, enabled: bool) -> None:
    ids = _audio_shape_ids(slide_root)
    _remove_audio_timing_nodes(slide_root, set(ids))
    if enabled:
        for shape_id in ids:
            _append_autoplay_to_slide_xml(slide_root, int(shape_id))


def _set_auto_advance_enabled_xml(
    slide_root: etree._Element, enabled: bool, duration_s: float | None = None,
) -> None:
    p_tag = lambda local: f'{{http://schemas.openxmlformats.org/presentationml/2006/main}}{local}'
    transition = slide_root.find(p_tag('transition'))
    if enabled:
        if duration_s is None:
            raise ValueError("durata audio non disponibile per normalizzare l'avanzamento")
        _set_slide_auto_advance_xml(slide_root, duration_s)
    elif transition is not None:
        transition.attrib.pop('advTm', None)
        transition.set('advClick', '1')


def _set_audio_icons_hidden_xml(slide_root: etree._Element, hidden: bool) -> None:
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    for audio_file in slide_root.findall(f'.//{{{a_ns}}}audioFile'):
        pic = audio_file
        while pic is not None and pic.tag != f'{{{p_ns}}}pic':
            pic = pic.getparent()
        if pic is None:
            continue
        off = pic.find(f'.//{{{a_ns}}}off')
        if off is not None:
            if hidden:
                off.set('x', str(ICON_OFFSCREEN_EMU))
                off.set('y', str(ICON_OFFSCREEN_EMU))
            else:
                off.set('x', '0')
                off.set('y', '0')
        for media_node in slide_root.findall(f'.//{{{p_ns}}}cMediaNode'):
            targets = {n.get('spid') for n in media_node.findall(f'.//{{{p_ns}}}spTgt')}
            c_nv_pr = pic.find(f'.//{{{p_ns}}}cNvPr')
            if c_nv_pr is not None and c_nv_pr.get('id') in targets:
                media_node.set('showWhenStopped', '0' if hidden else '1')

def _inject_audio_files_into_pptx(
    input_pptx: str,
    output_pptx: str,
    slide_jobs: list[tuple[int, str, str]],
    autoplay: bool,
    auto_advance: bool,
    cb,
    replace_existing_slides: set[int] | None = None,
    normalize_autoplay: bool | None = None,
    normalize_auto_advance: bool | None = None,
    normalize_icon_hidden: bool | None = None,
    analysis: dict | None = None,
    cancel_event=None,
) -> list[tuple[int, float]]:
    """Inserisce MP3 soltanto nelle slide indicate, lavorando a livello ZIP.

    Questo percorso non carica il file con python-pptx e quindi non ricrea né
    modifica gli audio già incorporati. Le parti originali vengono copiate
    byte-per-byte; cambiano soltanto le slide da completare, i relativi .rels,
    i nuovi media e i Content Types.
    """
    replace_existing_slides = set(replace_existing_slides or set())
    needs_normalization = any(
        value is not None for value in
        (normalize_autoplay, normalize_auto_advance, normalize_icon_hidden)
    )
    if not slide_jobs and not needs_normalization:
        shutil.copy2(input_pptx, output_pptx)
        return []

    audio_by_slide = {sn: path for sn, _text, path in slide_jobs}
    durations = [(sn, get_audio_duration_seconds(path)) for sn, path in audio_by_slide.items()]
    duration_by_slide = dict(durations)
    output_abs = os.path.abspath(output_pptx)
    output_dir = os.path.dirname(output_abs) or os.getcwd()
    os.makedirs(output_dir, exist_ok=True)
    fd, staging = tempfile.mkstemp(prefix='.slide_narrator_fix_', suffix='.pptx', dir=output_dir)
    os.close(fd)

    try:
        with zipfile.ZipFile(input_pptx, 'r') as zin:
            bad_zip = zin.testzip()
            if bad_zip:
                raise ValueError(f"PowerPoint corrotto: {bad_zip}")
            slide_parts = _ordered_slide_parts(zin)
            original_names = set(zin.namelist())
            planned_names = set(original_names)
            replacements: dict[str, bytes] = {}
            additions: dict[str, bytes] = {}

            icon_part = None
            if slide_jobs:
                icon_part = _next_package_part_name(
                    planned_names, 'ppt/media/slide_narrator_audio_icon.png'
                )
                additions[icon_part] = _FIX_AUDIO_ICON_PNG

            ct_root = etree.fromstring(zin.read('[Content_Types].xml'))
            if slide_jobs:
                _ensure_default_content_type(ct_root, 'mp3', 'audio/mpeg')
                _ensure_default_content_type(ct_root, 'png', 'image/png')
            replacements['[Content_Types].xml'] = etree.tostring(
                ct_root, xml_declaration=True, encoding='UTF-8', standalone=True
            )

            # Normalizzazione degli audio già presenti, slide per slide.
            if needs_normalization:
                for slide_num, slide_part in enumerate(slide_parts, 1):
                    _check_cancelled(cancel_event)
                    slide_root = etree.fromstring(zin.read(slide_part))
                    if not _audio_shape_ids(slide_root):
                        continue
                    rels_part = posixpath.join(
                        posixpath.dirname(slide_part), '_rels',
                        posixpath.basename(slide_part) + '.rels',
                    )
                    if normalize_autoplay is not None:
                        _set_audio_autoplay_xml(slide_root, normalize_autoplay)
                    if normalize_auto_advance is not None:
                        detail = (analysis or {}).get('details', {}).get(slide_num, {})
                        _set_auto_advance_enabled_xml(
                            slide_root, normalize_auto_advance, detail.get('duration_s')
                        )
                    if normalize_icon_hidden is not None:
                        _set_audio_icons_hidden_xml(slide_root, normalize_icon_hidden)
                    replacements[slide_part] = etree.tostring(
                        slide_root, xml_declaration=True, encoding='UTF-8', standalone=True
                    )

            total = len(slide_jobs)
            for done, (slide_num, _text, audio_path) in enumerate(slide_jobs, 1):
                _check_cancelled(cancel_event)
                if slide_num < 1 or slide_num > len(slide_parts):
                    raise ValueError(f"slide {slide_num} non trovata")
                slide_part = slide_parts[slide_num - 1]
                slide_root = etree.fromstring(replacements.get(slide_part, zin.read(slide_part)))
                p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
                sp_tree = slide_root.find(f'.//{{{p_ns}}}spTree')
                if sp_tree is None:
                    raise ValueError(f"slide {slide_num}: spTree non trovato")

                existing_audio = slide_root.findall(
                    './/{http://schemas.openxmlformats.org/drawingml/2006/main}audioFile'
                )
                if existing_audio and slide_num not in replace_existing_slides:
                    raise ValueError(
                        f"slide {slide_num}: contiene già audio; selezionarla esplicitamente per la rigenerazione"
                    )

                shape_ids = []
                for node in slide_root.findall(f'.//{{{p_ns}}}cNvPr'):
                    try:
                        shape_ids.append(int(node.get('id', '0')))
                    except ValueError:
                        pass
                shape_id = max(shape_ids or [1]) + 1

                rels_part = posixpath.join(
                    posixpath.dirname(slide_part), '_rels',
                    posixpath.basename(slide_part) + '.rels',
                )
                if rels_part in original_names:
                    rels_root = etree.fromstring(zin.read(rels_part))
                else:
                    rels_root = etree.Element(
                        f'{{{_PKG_RELS_NS}}}Relationships',
                        nsmap={None: _PKG_RELS_NS},
                    )

                if existing_audio and slide_num in replace_existing_slides:
                    removed = _remove_audio_from_slide_package(slide_root, rels_root)
                    print(
                        f"   slide {slide_num}: rimossi {removed['removed_shapes']} audio "
                        "prima della rigenerazione"
                    )

                media_rel_id = _next_relationship_id(rels_root)
                media_rel = etree.SubElement(
                    rels_root, f'{{{_PKG_RELS_NS}}}Relationship'
                )
                media_rel.set('Id', media_rel_id)
                media_rel.set('Type', _MEDIA_RELTYPE)

                audio_rel_id = _next_relationship_id(rels_root)
                audio_rel = etree.SubElement(
                    rels_root, f'{{{_PKG_RELS_NS}}}Relationship'
                )
                audio_rel.set('Id', audio_rel_id)
                audio_rel.set('Type', _AUDIO_RELTYPE)

                image_rel_id = _next_relationship_id(rels_root)
                image_rel = etree.SubElement(
                    rels_root, f'{{{_PKG_RELS_NS}}}Relationship'
                )
                image_rel.set('Id', image_rel_id)
                image_rel.set('Type', _IMAGE_RELTYPE)

                media_part = _next_package_part_name(
                    planned_names,
                    f'ppt/media/slide_narrator_fix_slide_{slide_num:04d}.mp3',
                )
                additions[media_part] = Path(audio_path).read_bytes()
                slide_dir = posixpath.dirname(slide_part)
                media_target = posixpath.relpath(media_part, slide_dir)
                image_target = posixpath.relpath(icon_part, slide_dir)
                media_rel.set('Target', media_target)
                audio_rel.set('Target', media_target)
                image_rel.set('Target', image_target)

                pic = _build_audio_picture_shape_ooxml(
                    shape_id, audio_rel_id, media_rel_id, image_rel_id,
                    f'Slide Narrator Fix slide {slide_num}.mp3',
                )
                sp_tree.append(pic)
                if autoplay:
                    _append_autoplay_to_slide_xml(slide_root, shape_id)
                if auto_advance and autoplay:
                    _set_slide_auto_advance_xml(
                        slide_root, duration_by_slide[slide_num]
                    )

                replacements[slide_part] = etree.tostring(
                    slide_root, xml_declaration=True, encoding='UTF-8', standalone=True
                )
                replacements[rels_part] = etree.tostring(
                    rels_root, xml_declaration=True, encoding='UTF-8', standalone=True
                )
                cb({"stage": "embedding_progress", "done": done, "total": total})

            with zipfile.ZipFile(staging, 'w', zipfile.ZIP_DEFLATED) as zout:
                for name in zin.namelist():
                    data = replacements.get(name, zin.read(name))
                    zout.writestr(zin.getinfo(name), data)
                # Alcuni PPTX non hanno un file .rels per una slide vuota.
                # In quel caso il nuovo .rels è in replacements ma non era
                # presente nell'archivio originale: va scritto esplicitamente.
                for name, data in replacements.items():
                    if name not in original_names:
                        zout.writestr(name, data)
                for name, data in additions.items():
                    if name not in original_names and name not in replacements:
                        zout.writestr(name, data)

        _fix_audio_relationship_types(staging)
        os.replace(staging, output_abs)
    finally:
        _remove_file_quietly(staging)
    return sorted(durations)


def _durations_from_pptx_analysis(analysis: dict) -> list[tuple[int, float | None]]:
    durations = []
    for slide_num in range(1, analysis.get('total_slides', 0) + 1):
        detail = analysis.get('details', {}).get(slide_num, {})
        duration = detail.get('duration_s') if detail.get('status') == 'valid' else None
        durations.append((slide_num, duration))
    return durations


def _rewrite_pptx_removing_audio(input_pptx: str, output_pptx: str, slides: set[int]) -> None:
    """Rimuove in modo selettivo audio, timing e relazioni da slide scelte."""
    with zipfile.ZipFile(input_pptx, "r") as zin:
        slide_parts = _ordered_slide_parts(zin)
        replacements: dict[str, bytes] = {}
        names = set(zin.namelist())
        for slide_num in sorted(slides):
            if slide_num < 1 or slide_num > len(slide_parts):
                raise ValueError(f"slide {slide_num} non trovata")
            slide_part = slide_parts[slide_num - 1]
            rels_part = posixpath.join(
                posixpath.dirname(slide_part), "_rels", posixpath.basename(slide_part) + ".rels"
            )
            root = etree.fromstring(zin.read(slide_part))
            if rels_part in names:
                rels = etree.fromstring(zin.read(rels_part))
            else:
                rels = etree.Element(f"{{{_PKG_RELS_NS}}}Relationships", nsmap={None: _PKG_RELS_NS})
            _remove_audio_from_slide_package(root, rels)
            replacements[slide_part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            replacements[rels_part] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(output_pptx, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                zout.writestr(zin.getinfo(name), replacements.get(name, zin.read(name)))
            for name, data in replacements.items():
                if name not in names:
                    zout.writestr(name, data)


def _write_execution_summary(output_path: str, payload: dict) -> str:
    """Scrive un rapporto JSON atomico accanto all'output."""
    out = Path(output_path)
    path = out.with_name(f"{out.stem}_esecuzione.json")
    data = {
        "version": 1,
        "created_at": datetime.now().astimezone().isoformat(),
        "output": str(out.resolve()),
        **payload,
    }
    _atomic_write_text(path, json.dumps(data, ensure_ascii=False, indent=2))
    return str(path)


def process_fix(
    input_pptx: str,
    scripts_xlsx: str | None,
    output_pptx: str,
    voice: str,
    rate: str,
    autoplay: bool = True,
    auto_advance: bool = False,
    report_path: str | None = None,
    concurrency: int = DEFAULT_CONCURRENCY,
    progress_callback: ProgressCallback = None,
    transcode_audio: bool = False,
    pocket_variant: str = "italian",
    clone_workers: int = 1,
    use_cache: bool = True,
    pocket_quantize: bool = False,
    *,
    volume: str = "+0%",
    pitch: str = "+0Hz",
    script_source: str = "excel",
    sheet_name: str | None = None,
    script_column: str = "A",
    has_header: bool = False,
    regenerate_slides: set[int] | list[int] | tuple[int, ...] | None = None,
    regenerate_all: bool = False,
    repair_invalid: bool = False,
    normalize_autoplay: bool | None = None,
    normalize_auto_advance: bool | None = None,
    normalize_icon_hidden: bool | None = None,
    cancel_event=None,
) -> str:
    """Completa, ripara o rigenera selettivamente gli audio di un PPTX.

    Per impostazione predefinita sintetizza solo le slide senza audio. Le slide
    già valide vengono preservate byte-per-byte; quelle corrotte o con più
    audio richiedono ``repair_invalid=True`` o una selezione esplicita in
    ``regenerate_slides``. La funzione può anche normalizzare autoplay,
    avanzamento e visibilità delle icone senza rigenerare l'audio.
    """
    cb = progress_callback if progress_callback is not None else (lambda e: None)
    _check_cancelled(cancel_event)
    if Path(input_pptx).resolve() == Path(output_pptx).resolve():
        raise ValueError("Il Fix deve essere salvato in un nuovo file: non sovrascrivere l'originale.")
    cb({"stage": "loading"})
    print(f"-> Modalità FIX PowerPoint: {input_pptx}")
    analysis = analyze_pptx_audio(input_pptx)
    total = analysis["total_slides"]
    print(
        f"   {total} slide | {len(analysis['valid'])} audio validi | "
        f"{len(analysis['missing'])} mancanti | {len(analysis['invalid'])} anomalie"
    )

    selected = {int(n) for n in (regenerate_slides or set())}
    out_of_range = sorted(n for n in selected if n < 1 or n > total)
    if out_of_range:
        raise ValueError("Slide selezionate non esistenti: " + ", ".join(map(str, out_of_range)))
    target = set(analysis["missing"])
    if regenerate_all:
        target.update(range(1, total + 1))
    target.update(selected)
    if repair_invalid:
        target.update(analysis["invalid"])

    unresolved_invalid = sorted(set(analysis["invalid"]) - target)
    if unresolved_invalid:
        details = "; ".join(
            f"slide {n}: {analysis['invalid'][n]}" for n in unresolved_invalid
        )
        raise ValueError(
            "È presente una struttura audio ambigua o danneggiata non selezionata per la riparazione. "
            "Attivare 'Ripara audio corrotti/anomali' oppure selezionare le slide. " + details
        )

    metadata = get_pptx_slide_metadata(input_pptx)
    source = (script_source or "excel").strip().lower()
    if source in {"notes", "note", "note relatore", "note powerpoint", "speaker_notes"}:
        print("-> Lettura script dalle Note PowerPoint")
        scripts = read_scripts_from_notes(input_pptx)
    else:
        if not scripts_xlsx:
            raise ValueError("Selezionare il file Excel degli script per il Fix.")
        print(f"-> Lettura script da {scripts_xlsx}")
        scripts = read_scripts_from_xlsx(
            scripts_xlsx, sheet_name=sheet_name, column=script_column,
            has_header=has_header, slide_metadata=metadata,
        )

    missing_scripts: list[int] = []
    intentionally_silent: list[int] = []
    planned: list[tuple[int, str]] = []
    for slide_num in sorted(target):
        text = scripts[slide_num - 1] if slide_num <= len(scripts) else ""
        if text.upper() in SILENT_SCRIPT_MARKERS:
            intentionally_silent.append(slide_num)
        elif not text:
            missing_scripts.append(slide_num)
        else:
            planned.append((slide_num, text))
    if missing_scripts:
        preview = ", ".join(map(str, missing_scripts[:30]))
        if len(missing_scripts) > 30:
            preview += ", ..."
        raise ValueError(
            "Manca lo script per le slide da completare/rigenerare: " + preview + ". "
            "Le slide non interessate possono avere la cella vuota; usare [SENZA AUDIO] "
            "per rimuovere o lasciare intenzionalmente la narrazione."
        )

    replace_existing = target & (set(analysis["valid"]) | set(analysis["invalid"]))
    # Una slide selezionata e marcata senza audio deve perdere l'audio esistente.
    remove_only_slides = set(intentionally_silent) & replace_existing
    cloned_voice = None
    if _CLONE_AVAILABLE and voice_library.is_clone_voice(voice):
        cloned_voice = voice_library.VoiceLibrary().get(voice)

    captured_timings: dict[int, dict] = {}
    with tempfile.TemporaryDirectory() as tmp:
        slide_jobs = [
            (slide_num, text, os.path.join(tmp, f"fix_slide_{slide_num:04d}.mp3"))
            for slide_num, text in planned
        ]
        if slide_jobs:
            if cloned_voice is not None:
                synthesis_errors, captured_timings = _run_synthesis_phase_clone(
                    slide_jobs, cloned_voice, rate, cb,
                    pocket_variant=pocket_variant, workers=clone_workers,
                    use_cache=use_cache, pocket_quantize=pocket_quantize,
                )
            else:
                synthesis_errors, captured_timings = _run_synthesis_phase(
                    slide_jobs, voice, rate, concurrency, transcode_audio, cb,
                    volume=volume, pitch=pitch, cancel_event=cancel_event,
                )
            audio_issues = _audit_synthesized_audio(slide_jobs)
            for slide_num, error in synthesis_errors.items():
                audio_issues.setdefault(slide_num, error)
            for _recovery_pass in range(AUDIO_RECOVERY_PASSES):
                _check_cancelled(cancel_event)
                if not audio_issues:
                    break
                recovery_tasks = [job for job in slide_jobs if job[0] in audio_issues]
                print("-> Recupero automatico Fix: slide " + str(sorted(audio_issues)))
                if cloned_voice is not None:
                    recovery_errors, recovery_timings = _run_synthesis_phase_clone(
                        recovery_tasks, cloned_voice, rate, cb,
                        pocket_variant=pocket_variant, workers=1,
                        use_cache=use_cache, pocket_quantize=pocket_quantize,
                    )
                else:
                    recovery_errors, recovery_timings = _run_synthesis_phase(
                        recovery_tasks, voice, rate, 1, transcode_audio, cb,
                        volume=volume, pitch=pitch, cancel_event=cancel_event,
                    )
                captured_timings.update(recovery_timings)
                audio_issues = _audit_synthesized_audio(slide_jobs)
                for slide_num, error in recovery_errors.items():
                    audio_issues.setdefault(slide_num, error)
            if audio_issues:
                raise RuntimeError(
                    "Fix interrotto: impossibile creare tutti gli audio. " + "; ".join(
                        f"slide {n}: {err}" for n, err in sorted(audio_issues.items())
                    )
                )

        _check_cancelled(cancel_event)
        cb({"stage": "saving"})
        # Per rimuovere una narrazione senza sostituirla usiamo una copia ZIP
        # preparatoria e poi la normale iniezione/normalizzazione.
        source_for_injection = input_pptx
        removal_tmp = None
        if remove_only_slides:
            fd, removal_tmp = tempfile.mkstemp(prefix=".slide_narrator_remove_", suffix=".pptx")
            os.close(fd)
            _rewrite_pptx_removing_audio(input_pptx, removal_tmp, remove_only_slides)
            source_for_injection = removal_tmp
            replace_existing -= remove_only_slides
        try:
            _inject_audio_files_into_pptx(
                source_for_injection, output_pptx, slide_jobs,
                autoplay=autoplay, auto_advance=auto_advance, cb=cb,
                replace_existing_slides=replace_existing,
                normalize_autoplay=normalize_autoplay,
                normalize_auto_advance=normalize_auto_advance,
                normalize_icon_hidden=normalize_icon_hidden,
                analysis=analysis, cancel_event=cancel_event,
            )
        finally:
            if removal_tmp:
                _remove_file_quietly(removal_tmp)

    _check_cancelled(cancel_event)
    final_analysis = analyze_pptx_audio(output_pptx)
    expected_new = {job[0] for job in slide_jobs}
    expected_valid = (set(analysis["valid"]) - target) | expected_new
    lost = sorted(expected_valid - set(final_analysis["valid"]))
    remaining_invalid = sorted(final_analysis["invalid"])
    if lost or remaining_invalid:
        details = []
        if lost:
            details.append("audio mancanti/non validi nelle slide " + ", ".join(map(str, lost)))
        if remaining_invalid:
            details.append("anomalie nelle slide " + ", ".join(map(str, remaining_invalid)))
        raise RuntimeError("Controllo finale Fix non superato: " + " | ".join(details))

    existing_captions = _load_existing_captions(input_pptx)
    known_audio = set(final_analysis["valid"])
    _write_captions_json(
        output_pptx, voice, rate, captured_timings,
        existing_data=existing_captions,
        replaced_slides=target,
        known_audio_slides=known_audio,
    )
    report_path = _write_duration_report(
        output_pptx, report_path, _durations_from_pptx_analysis(final_analysis)
    )
    _write_execution_summary(output_pptx, {
        "mode": "fix",
        "status": "completed",
        "input": str(Path(input_pptx).resolve()),
        "slides_total": total,
        "slides_generated": sorted(expected_new),
        "slides_removed_audio": sorted(remove_only_slides),
        "slides_preserved": sorted(set(analysis["valid"]) - target),
        "normalization": {
            "autoplay": normalize_autoplay,
            "auto_advance": normalize_auto_advance,
            "icon_hidden": normalize_icon_hidden,
        },
        "final_analysis": final_analysis,
    })
    print(
        f"-> Fix completato: {len(expected_new)} audio creati, "
        f"{len(remove_only_slides)} rimossi, "
        f"{len(set(analysis['valid']) - target)} preservati"
    )
    cb({"stage": "done"})
    return report_path

def process(
    input_pptx: str,
    scripts_xlsx: str | None,
    output_pptx: str,
    voice: str,
    rate: str,
    autoplay: bool = True,
    auto_advance: bool = False,
    report_path: str | None = None,
    concurrency: int = DEFAULT_CONCURRENCY,
    progress_callback: ProgressCallback = None,
    transcode_audio: bool = False,
    output_mode: str = "pptx",
    resolution: str = "1080p",
    subtitles: bool = False,
    transition: bool = False,
    pocket_variant: str = "italian",
    clone_workers: int = 1,
    use_cache: bool = True,
    pocket_quantize: bool = False,
    *,
    volume: str = "+0%",
    pitch: str = "+0Hz",
    script_source: str = "excel",
    sheet_name: str | None = None,
    script_column: str = "A",
    has_header: bool = False,
    silent_slide_s: float = 3.0,
    transition_style: str = "fade",
    render_backend: str = "auto",
    cancel_event=None,
) -> str:
    """Esegue la generazione completa PowerPoint o MP4 con verifica finale."""
    cb = progress_callback if progress_callback is not None else (lambda e: None)
    _check_cancelled(cancel_event)
    cloned_voice = None
    if _CLONE_AVAILABLE and voice_library.is_clone_voice(voice):
        cloned_voice = voice_library.VoiceLibrary().get(voice)

    prs, scripts, n_slides, n_scripts = _load_inputs(
        input_pptx, scripts_xlsx, cb, script_source=script_source,
        sheet_name=sheet_name, script_column=script_column,
        has_header=has_header,
    )

    with tempfile.TemporaryDirectory() as tmp:
        slide_jobs = _build_slide_jobs(scripts, n_slides, n_scripts, tmp)
        synth_tasks = [job for job in slide_jobs if job is not None]
        if cloned_voice is not None:
            synthesis_errors, captured_timings = _run_synthesis_phase_clone(
                synth_tasks, cloned_voice, rate, cb,
                pocket_variant=pocket_variant, workers=clone_workers,
                use_cache=use_cache, pocket_quantize=pocket_quantize,
            )
        else:
            synthesis_errors, captured_timings = _run_synthesis_phase(
                synth_tasks, voice, rate, concurrency, transcode_audio, cb,
                volume=volume, pitch=pitch, cancel_event=cancel_event,
            )

        audio_issues = _audit_synthesized_audio(slide_jobs)
        for slide_num, error in synthesis_errors.items():
            audio_issues.setdefault(slide_num, error)
        for _recovery_pass in range(1, AUDIO_RECOVERY_PASSES + 1):
            _check_cancelled(cancel_event)
            if not audio_issues:
                break
            recovery_tasks = [
                job for job in slide_jobs if job is not None and job[0] in audio_issues
            ]
            print("-> Recupero automatico audio: slide " + str(sorted(audio_issues)))
            if cloned_voice is not None:
                recovery_errors, recovery_timings = _run_synthesis_phase_clone(
                    recovery_tasks, cloned_voice, rate, cb,
                    pocket_variant=pocket_variant, workers=1,
                    use_cache=use_cache, pocket_quantize=pocket_quantize,
                )
            else:
                recovery_errors, recovery_timings = _run_synthesis_phase(
                    recovery_tasks, voice, rate, 1, transcode_audio, cb,
                    volume=volume, pitch=pitch, cancel_event=cancel_event,
                )
            captured_timings.update(recovery_timings)
            audio_issues = _audit_synthesized_audio(slide_jobs)
            for slide_num, error in recovery_errors.items():
                audio_issues.setdefault(slide_num, error)
        if audio_issues:
            raise RuntimeError(
                "Generazione interrotta: audio mancanti o non validi. " + "; ".join(
                    f"slide {n}: {err}" for n, err in sorted(audio_issues.items())
                )
            )
        print(f"-> Controllo audio superato: {len(synth_tasks)}/{len(synth_tasks)}")

        if output_mode == "video":
            if not _VIDEO_AVAILABLE:
                raise RuntimeError("Esportazione video non disponibile: modulo o dipendenze mancanti.")
            result = _export_video(
                input_pptx, output_pptx, slide_jobs, {}, captured_timings,
                resolution, subtitles, transition, report_path, tmp, cb,
                silent_slide_s=silent_slide_s,
                transition_style=transition_style,
                render_backend=render_backend,
                cancel_event=cancel_event,
            )
            _write_execution_summary(output_pptx, {
                "mode": "video", "status": "completed",
                "input": str(Path(input_pptx).resolve()),
                "slides_total": n_slides,
                "slides_with_audio": [j[0] for j in synth_tasks],
                "silent_slides": [i + 1 for i, j in enumerate(slide_jobs) if j is None],
                "resolution": resolution,
                "subtitles": subtitles,
                "transition": transition,
                "transition_style": transition_style,
                "render_backend": render_backend,
            })
            return result

        durations = _run_embedding_phase(
            prs, slide_jobs, {}, autoplay, cb,
            auto_advance=auto_advance, cancel_event=cancel_event,
        )
        expected_audio_slides = {job[0] for job in slide_jobs if job is not None}
        output_abs = os.path.abspath(output_pptx)
        output_dir = os.path.dirname(output_abs) or os.getcwd()
        os.makedirs(output_dir, exist_ok=True)
        fd, staging_pptx = tempfile.mkstemp(prefix=".slide_narrator_build_", suffix=".pptx", dir=output_dir)
        os.close(fd)
        package_issues: dict[int, str] = {}
        try:
            for save_attempt in range(1, PPTX_SAVE_VERIFY_ATTEMPTS + 1):
                _check_cancelled(cancel_event)
                cb({"stage": "saving"})
                print(f"-> Salvataggio e verifica ({save_attempt}/{PPTX_SAVE_VERIFY_ATTEMPTS})")
                prs.save(staging_pptx)
                fixed, mainseq_fixes = _fix_audio_relationship_types(staging_pptx)
                if fixed:
                    print(f"-> Parti OOXML corrette: {fixed}")
                if mainseq_fixes:
                    print(f"-> mainSeq riparati: {mainseq_fixes}")
                package_issues = verify_pptx_audio_package(
                    staging_pptx, expected_audio_slides, autoplay=autoplay
                )
                if not package_issues:
                    break
                if save_attempt < PPTX_SAVE_VERIFY_ATTEMPTS:
                    print("   ! reinserimento audio nelle slide " + str(sorted(package_issues)))
                    for slide_num in sorted(package_issues):
                        _check_cancelled(cancel_event)
                        job = slide_jobs[slide_num - 1]
                        if job is None:
                            continue
                        _sn, _text, audio_path = job
                        slide = prs.slides[slide_num - 1]
                        add_audio_to_slide(slide, audio_path, autoplay=autoplay, slide_num=slide_num)
                        if auto_advance and autoplay:
                            _set_slide_auto_advance(slide, get_audio_duration_seconds(audio_path))
            if package_issues:
                raise RuntimeError(
                    "Il PowerPoint non ha superato il controllo finale. " + "; ".join(
                        f"slide {n}: {err}" for n, err in sorted(package_issues.items())
                    )
                )
            _check_cancelled(cancel_event)
            os.replace(staging_pptx, output_abs)
            final_issues = verify_pptx_audio_package(
                output_abs, expected_audio_slides, autoplay=autoplay
            )
            if final_issues:
                raise RuntimeError(
                    "Controllo dopo il salvataggio non superato: " + "; ".join(
                        f"slide {n}: {err}" for n, err in sorted(final_issues.items())
                    )
                )
        finally:
            _remove_file_quietly(staging_pptx)

    _write_captions_json(
        output_pptx, voice, rate, captured_timings,
        known_audio_slides=expected_audio_slides,
    )
    report_path = _write_duration_report(output_pptx, report_path, durations)
    _write_execution_summary(output_pptx, {
        "mode": "pptx", "status": "completed",
        "input": str(Path(input_pptx).resolve()),
        "slides_total": n_slides,
        "slides_with_audio": sorted(expected_audio_slides),
        "silent_slides": [i + 1 for i, j in enumerate(slide_jobs) if j is None],
        "autoplay": autoplay,
        "auto_advance": auto_advance,
    })
    print("Fatto.")
    cb({"stage": "done"})
    return report_path


# ----------------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------------
def parse_args():
    p = argparse.ArgumentParser(
        description="Generate Italian TTS audio from XLSX and embed into PowerPoint slides."
    )
    p.add_argument("input_pptx", help="Input .pptx file")
    p.add_argument("scripts_xlsx", nargs="?", default=None, help="XLSX con gli script; omettere con --script-source notes")
    p.add_argument("output_pptx", help="Output .pptx with embedded audio")
    p.add_argument(
        "--voice",
        default="it-IT-IsabellaNeural",
        help="Edge TTS voice. Italian options: "
             + ", ".join(ITALIAN_VOICES.values())
             + " (default: it-IT-IsabellaNeural)",
    )
    p.add_argument(
        "--rate",
        default="+0%",
        help="Speech rate, e.g. '-10%%', '+0%%', '+15%%' (default: +0%%)",
    )
    p.add_argument(
        "--no-autoplay",
        action="store_true",
        help="Do not configure autoplay (audio will require a click to play)",
    )
    p.add_argument(
        "--auto-advance",
        action="store_true",
        help=("Avanza automaticamente alla slide successiva a fine audio "
              "(richiede l'autoplay attivo): la presentazione si riproduce "
              "da sola, slide dopo slide."),
    )
    p.add_argument(
        "--concurrency",
        type=int,
        default=DEFAULT_CONCURRENCY,
        help=f"Number of audio syntheses to run in parallel "
             f"(default: {DEFAULT_CONCURRENCY}). Lower if Edge TTS rate-limits "
             f"you, raise to gain speed at your own risk.",
    )
    p.add_argument(
        "--report",
        default=None,
        help="Path of the duration report .txt (default: <output>_durate.txt)",
    )
    p.add_argument(
        "--transcode-audio",
        action="store_true",
        help=(
            "Riconverti gli MP3 in MPEG-1 @ 44100Hz mono via ffmpeg. "
            "Migliora la compatibilita' con alcune versioni di PowerPoint, "
            "ma raddoppia la dimensione dei file audio. "
            "Default: disattivato (audio Edge TTS originali, MPEG-2 24kHz)."
        ),
    )
    p.add_argument(
        "--video",
        action="store_true",
        help="Genera un video MP4 (slide + audio) invece del pptx con audio.",
    )
    p.add_argument(
        "--resolution",
        choices=["720p", "1080p"],
        default="1080p",
        help="Risoluzione del video (default: 1080p). Solo con --video.",
    )
    p.add_argument(
        "--subtitles",
        action="store_true",
        help="Brucia i sottotitoli nel video. Solo con --video.",
    )
    p.add_argument(
        "--transition",
        action="store_true",
        help="Dissolvenza incrociata tra le slide. Solo con --video.",
    )
    p.add_argument(
        "--pocket-variant",
        choices=["italian", "italian_24l"],
        default="italian",
        help=("Qualità del modello PocketTTS per le voci clonate: "
              "'italian' (veloce, default) o 'italian_24l' (qualità alta, lenta)."),
    )
    p.add_argument(
        "--clone-workers",
        type=int,
        default=1,
        help=("Slide sintetizzate in parallelo con le voci clonate (1 = "
              "sequenziale). 2-10 può accelerare su CPU multi-core; ogni processo "
              "carica una copia del modello, quindi consuma più RAM."),
    )
    p.add_argument(
        "--no-cache",
        action="store_true",
        help=("Disattiva la cache: forza la rigenerazione di ogni slide anche se "
              "un audio identico è già stato prodotto in precedenza."),
    )
    p.add_argument(
        "--pocket-quantize",
        action="store_true",
        help=("Sperimentale: quantizzazione INT8 di PocketTTS (più veloce su CPU, "
              "qualità da verificare). Ignorato dagli altri motori."),
    )
    p.add_argument("--fix", action="store_true", help="Completa/ripara un PPTX già sonorizzato.")
    p.add_argument("--volume", default="+0%", help="Volume Edge TTS, es. +10%% o -20%%.")
    p.add_argument("--pitch", default="+0Hz", help="Tonalità Edge TTS, es. +20Hz o -10Hz.")
    p.add_argument("--script-source", choices=["excel", "notes"], default="excel")
    p.add_argument("--sheet", default=None, help="Nome del foglio Excel.")
    p.add_argument("--column", default="A", help="Colonna degli script.")
    p.add_argument("--header", action="store_true", help="La prima riga Excel è un'intestazione.")
    p.add_argument("--repair-invalid", action="store_true", help="Nel Fix sostituisce audio corrotti o multipli.")
    p.add_argument("--regenerate-all", action="store_true", help="Nel Fix rigenera tutte le narrazioni.")
    p.add_argument("--regenerate-slides", default="", help="Nel Fix: numeri slide separati da virgola.")
    p.add_argument("--normalize-autoplay", action="store_true")
    p.add_argument("--normalize-auto-advance", action="store_true")
    p.add_argument("--normalize-hidden-icons", action="store_true")
    p.add_argument("--silent-slide-seconds", type=float, default=3.0)
    p.add_argument("--render-backend", choices=["auto", "powerpoint", "libreoffice"], default="auto")
    p.add_argument(
        "--transition-style", default="fade",
        choices=["fade", "wipeleft", "wiperight", "slideleft", "slideright", "dissolve"],
    )
    return p.parse_args()


def _force_utf8_stdio():
    """Forza UTF-8 su stdout/stderr quando lanciati da console Windows.
    Senza questo, i print con caratteri come ↳, →, ⚠, ✓, è, à crashano
    sulla console Windows con codepage default (cp1252)."""
    if sys.platform.startswith('win'):
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
            sys.stderr.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass


def main():
    _force_utf8_stdio()
    args = parse_args()
    # Per le voci clonate (clone:motore:slug) non applico l'alias delle voci
    # Microsoft né tocco maiuscole/minuscole.
    if _CLONE_AVAILABLE and voice_library.is_clone_voice(args.voice):
        voice = args.voice
    else:
        voice = ITALIAN_VOICES.get(args.voice.lower(), args.voice)

    if not Path(args.input_pptx).exists():
        sys.exit(f"Input pptx not found: {args.input_pptx}")
    if args.script_source == "excel":
        if not args.scripts_xlsx or not Path(args.scripts_xlsx).exists():
            sys.exit("Specificare un file Excel esistente oppure usare --script-source notes")
    common = dict(
        input_pptx=args.input_pptx,
        scripts_xlsx=args.scripts_xlsx,
        output_pptx=args.output_pptx,
        voice=voice,
        rate=args.rate,
        autoplay=not args.no_autoplay,
        auto_advance=args.auto_advance,
        report_path=args.report,
        concurrency=max(1, args.concurrency),
        transcode_audio=args.transcode_audio,
        pocket_variant=args.pocket_variant,
        clone_workers=args.clone_workers,
        use_cache=not args.no_cache,
        pocket_quantize=args.pocket_quantize,
        volume=args.volume,
        pitch=args.pitch,
        script_source=args.script_source,
        sheet_name=args.sheet,
        script_column=args.column,
        has_header=args.header,
    )
    if args.fix:
        selected = {int(x.strip()) for x in args.regenerate_slides.split(",") if x.strip()}
        process_fix(
            **common,
            regenerate_slides=selected,
            regenerate_all=args.regenerate_all,
            repair_invalid=args.repair_invalid,
            normalize_autoplay=((not args.no_autoplay) if args.normalize_autoplay else None),
            normalize_auto_advance=(args.auto_advance if args.normalize_auto_advance else None),
            normalize_icon_hidden=(True if args.normalize_hidden_icons else None),
        )
    else:
        process(
            **common,
            output_mode=("video" if args.video else "pptx"),
            resolution=args.resolution,
            subtitles=args.subtitles,
            transition=args.transition,
            silent_slide_s=max(0.5, args.silent_slide_seconds),
            transition_style=args.transition_style,
            render_backend=args.render_backend,
        )



if __name__ == "__main__":
    main()
