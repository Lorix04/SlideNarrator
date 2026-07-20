"""
pptx_tts.py - Generate Italian TTS audio from XLSX scripts and embed into PowerPoint slides

Usage:
    python pptx_tts.py input.pptx scripts.xlsx output.pptx
    python pptx_tts.py input.pptx scripts.xlsx output.pptx --voice it-IT-IsabellaNeural --rate +0%

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
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Callable, Optional

import edge_tts
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Emu, Inches
from pptx.oxml.ns import qn
from lxml import etree
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

# Formato MP3 nativo prodotto da Edge TTS vs formato target dopo transcoding.
EDGE_TTS_NATIVE_SAMPLE_RATE_HZ = 24000
TARGET_SAMPLE_RATE_HZ = 44100
TARGET_BITRATE = "96k"

# Type alias per il callback di progresso (vedi process() per gli stage).
ProgressCallback = Optional[Callable[[dict], None]]


# Available Italian voices in Edge TTS
ITALIAN_VOICES = {
    "isabella": "it-IT-IsabellaNeural",   # female, warm
    "elsa":     "it-IT-ElsaNeural",       # female, clear
    "diego":    "it-IT-DiegoNeural",      # male, professional
    "giuseppe": "it-IT-GiuseppeNeural",   # male, mature
}


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
async def _synthesize(text: str, voice: str, rate: str, output_path: str) -> None:
    communicate = edge_tts.Communicate(text=text, voice=voice, rate=rate)
    await communicate.save(output_path)


def synthesize_to_file(text: str, voice: str, rate: str, output_path: str) -> None:
    """Synthesize a single text. Used by the GUI for the voice preview."""
    asyncio.run(_synthesize(text, voice, rate, output_path))


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
async def _consume_stream(communicate, audio_out, words: list) -> None:
    """Consuma lo stream di una Communicate edge_tts: scrive i chunk audio
    su `audio_out` e accumula i WordBoundary in `words`.

    Estratto in funzione separata per poterla wrappare con asyncio.wait_for
    e fare timeout senza altri side-effect."""
    async for chunk in communicate.stream():
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
    semaphore: asyncio.Semaphore,
    progress_state: dict,
    progress_callback: ProgressCallback,
    captured_timings: dict | None = None,
    transcode_audio: bool = False,
) -> tuple[int, bool, str]:
    """
    Synthesize one slide's audio, respecting the concurrency limit.
    Returns (slide_num, success, error_message_or_empty).
    Errors are caught and reported per-slide, so a single failed slide does
    not abort the rest of the batch.

    Edge TTS, oltre all'audio, espone uno stream di eventi WordBoundary che
    indicano per ogni parola il millisecondo di inizio nell'audio prodotto.
    Catturiamo questi eventi mentre sintetizziamo: se viene fornito il dict
    `captured_timings`, ci scriviamo dentro la lista delle parole con i loro
    offset, indicizzata per slide_num. Questi timings finiscono poi in un
    file JSON accanto al .pptx, usato dal SCORM Builder per i sottotitoli
    sincronizzati frase-per-frase.

    On completion (whether success or failure) increments
    progress_state["done"] and invokes progress_callback with a
    "synthesis_progress" event. asyncio is single-threaded so the increment
    and callback are safe without explicit locks.
    """
    async with semaphore:
        t0 = time.monotonic()
        try:
            # Edge TTS 7.x ha cambiato il default da WordBoundary a
            # SentenceBoundary. Per ottenere i timing parola-per-parola
            # (necessari per le captions sincronizzate del SCORM Builder)
            # serve passare esplicitamente boundary="WordBoundary".
            try:
                communicate = edge_tts.Communicate(
                    text=text, voice=voice, rate=rate,
                    boundary="WordBoundary",
                )
            except TypeError:
                # Fallback per edge-tts < 7.x che non accetta il parametro
                communicate = edge_tts.Communicate(text=text, voice=voice, rate=rate)
            words: list[dict] = []
            # Streaming mode: invece di .save() che butta via gli eventi non-audio,
            # iteriamo manualmente tenendo gli "audio" per scrivere il MP3 e
            # i "WordBoundary" per i timing.
            #
            # asyncio.wait_for protegge da WebSocket appesi: se Edge TTS smette
            # di mandare chunk per N secondi, la coroutine viene cancellata e
            # solo QUESTA slide fallisce (TimeoutError → catturato dal except
            # Exception generale qui sotto, marcato come errore di sintesi).
            with open(output_path, 'wb') as audio_out:
                try:
                    await asyncio.wait_for(
                        _consume_stream(communicate, audio_out, words),
                        timeout=EDGE_TTS_STREAM_TIMEOUT_S,
                    )
                except asyncio.TimeoutError:
                    raise RuntimeError(
                        f"timeout dopo {EDGE_TTS_STREAM_TIMEOUT_S}s "
                        f"(server Edge TTS non risponde)"
                    ) from None
            elapsed = time.monotonic() - t0

            # Transcodifica opzionale: se transcode_audio=True, riconverte
            # l'MP3 da MPEG-2 24kHz (formato Edge TTS) a MPEG-1 44100Hz
            # (formato preferito da PowerPoint). Costa ~0.5s e raddoppia
            # la dimensione del file. Default OFF per risparmiare spazio.
            #
            # IMPORTANTE: _transcode_mp3_for_powerpoint usa subprocess.run
            # che è bloccante. Eseguito direttamente qui bloccherebbe l'event
            # loop asyncio, serializzando tutte le altre sintesi in flight
            # ad ogni transcode (~0.5s × N slide). Usiamo run_in_executor
            # per spostare la chiamata su un thread del pool default, così
            # le altre coroutine continuano a streamare audio in parallelo.
            if transcode_audio:
                t_transcode = time.monotonic()
                transcoded_path = output_path + ".tmp.mp3"
                loop = asyncio.get_running_loop()
                ok_transcode = await loop.run_in_executor(
                    None, _transcode_mp3_for_powerpoint,
                    output_path, transcoded_path,
                )
                if ok_transcode:
                    os.replace(transcoded_path, output_path)
                    tc_elapsed = time.monotonic() - t_transcode
                    print(f"   slide {slide_num}: sintesi ok ({elapsed:.1f}s + "
                          f"transcode {tc_elapsed:.1f}s, "
                          f"{len(text)} caratteri, {len(words)} parole timed)")
                else:
                    if os.path.exists(transcoded_path):
                        try:
                            os.unlink(transcoded_path)
                        except Exception:
                            pass
                    print(f"   slide {slide_num}: sintesi ok ({elapsed:.1f}s, "
                          f"{len(text)} caratteri, {len(words)} parole timed) "
                          f"[transcode richiesto ma fallito]")
            else:
                print(f"   slide {slide_num}: sintesi ok ({elapsed:.1f}s, "
                      f"{len(text)} caratteri, {len(words)} parole timed)")
            if captured_timings is not None:
                captured_timings[slide_num] = {
                    "text": text,
                    "words": words,
                }
            ok, err = True, ""
        except Exception as e:
            elapsed = time.monotonic() - t0
            print(f"   slide {slide_num}: ERRORE sintesi dopo {elapsed:.1f}s — {e}")
            ok, err = False, str(e)

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
        return (slide_num, ok, err)


async def _synthesize_many(
    tasks: list[tuple[int, str, str]],
    voice: str,
    rate: str,
    concurrency: int,
    progress_callback: ProgressCallback = None,
    captured_timings: dict | None = None,
    transcode_audio: bool = False,
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
            slide_num, text, voice, rate, path, semaphore,
            progress_state, progress_callback,
            captured_timings=captured_timings,
            transcode_audio=transcode_audio,
        )
        for slide_num, text, path in tasks
    ]
    return await asyncio.gather(*coros)


def synthesize_many_to_files(
    tasks: list[tuple[int, str, str]],
    voice: str,
    rate: str,
    concurrency: int = DEFAULT_CONCURRENCY,
    progress_callback: ProgressCallback = None,
    captured_timings: dict | None = None,
    transcode_audio: bool = False,
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
        tasks, voice, rate, concurrency, progress_callback,
        captured_timings=captured_timings,
        transcode_audio=transcode_audio,
    ))


def get_audio_duration_seconds(mp3_path: str) -> float:
    """Return the duration of an MP3 file in seconds."""
    return float(MP3(mp3_path).info.length)


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
def read_scripts_from_xlsx(xlsx_path: str) -> list[str]:
    wb = load_workbook(xlsx_path, data_only=True)
    try:
        ws = wb.active
        scripts = []
        for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
            cell = row[0]
            scripts.append("" if cell is None else str(cell).strip())
    finally:
        wb.close()
    while scripts and not scripts[-1]:
        scripts.pop()
    return scripts


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

    # Rimuovo eventuali transizioni preesistenti, poi inserisco la nuova.
    for old in sld.findall(qn('p:transition')):
        sld.remove(old)
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


def _remove_existing_audio_shapes(slide) -> int:
    """
    Rimuove dalla slide tutte le shape <p:pic> che contengono <a:audioFile>.
    Restituisce il numero di shape rimosse.

    Necessario quando un PPTX ha shape audio "orfane" (es. dopo un fix
    manuale, o copiate da un altro PPTX): senza questa pulizia
    add_audio_to_slide aggiungerebbe una SECONDA shape audio sopra,
    ma il <p:timing> minimale ha un solo <p:spTgt> → uno dei due audio
    parte in autoplay, l'altro resta orfano.

    LIMITAZIONE NOTA: la rigenerazione di un PPTX già prodotto da PPTX TTS
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


def add_audio_to_slide(
    slide,
    audio_path: str,
    autoplay: bool = True,
    slide_num: int | None = None,
) -> None:
    """
    Insert an audio file into the slide, autoplay-configured by default.
    The audio icon is positioned outside the slide boundary so it never
    appears in the editing view nor during the slideshow.

    Se la slide ha già shape audio (PPTX rigenerato), le rimuove prima
    di aggiungere quella nuova, con warning nel log.

    Se la slide ha già un <p:timing> con animazioni custom non vuote, lo
    sovrascrive ma con warning nel log: l'utente sa che le animazioni
    della slide originale vengono perse (data loss esplicito invece che
    silenzioso).

    `slide_num` è solo per il logging — opzionale.
    """
    sld = slide.element
    prefix = f"   slide {slide_num}: " if slide_num is not None else "   "

    # Pulizia preventiva: se ci sono audio shape già presenti, le rimuovo.
    n_removed = _remove_existing_audio_shapes(slide)
    if n_removed > 0:
        print(f"{prefix}! rimossi {n_removed} audio esistenti "
              f"(PPTX rigenerato — sostituisco)")

    # Check del timing PRIMA di add_movie, perché python-pptx ne inietta
    # uno suo nel processo (con un media-call effect): se controllassimo
    # dopo, troveremmo sempre un timing "non vuoto" anche su slide vergini.
    pre_existing_timing = sld.find(qn('p:timing'))
    pre_existing_has_animations = (
        pre_existing_timing is not None
        and _timing_has_animations(pre_existing_timing)
    )

    # Position the icon way off the top-left corner of the slide so it is
    # out of the visible area (PowerPoint accepts negative offsets).
    off_slide = Emu(ICON_OFFSCREEN_EMU)
    movie_shape = slide.shapes.add_movie(
        audio_path,
        left=off_slide, top=off_slide,
        width=Inches(ICON_SIZE_INCHES), height=Inches(ICON_SIZE_INCHES),
        mime_type='audio/mpeg',
    )

    # Make PowerPoint treat the media as audio (smaller speaker icon, audio
    # styles applicable like "Play in Background").
    _convert_video_ref_to_audio_ref(movie_shape)

    if not autoplay:
        return

    shape_id = movie_shape.shape_id
    existing = sld.find(qn('p:timing'))
    if existing is not None:
        if pre_existing_has_animations:
            print(f"{prefix}! timing/animazioni esistenti sovrascritti")
        sld.remove(existing)
    sld.append(_build_autoplay_timing(shape_id))


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


def _load_inputs(input_pptx: str, scripts_xlsx: str, cb) -> tuple:
    """Carica pptx + xlsx. Restituisce (prs, scripts, n_slides, n_scripts).
    Solleva ValueError se l'XLSX non contiene script utili o se il PPTX
    di input è già stato processato (contiene shape audio)."""
    cb({"stage": "loading"})
    print(f"-> Caricamento {input_pptx}")
    prs = Presentation(input_pptx)
    n_slides = len(prs.slides)

    n_existing_audio = _count_existing_audio_shapes_in_pptx(prs)
    if n_existing_audio > 0:
        raise ValueError(
            f"Il PPTX di input contiene già {n_existing_audio} shape "
            f"audio. Sembra un PPTX già processato da PPTX TTS. "
            f"Per rigenerare con voce/script diversi, riparti dal PPTX "
            f"ORIGINALE senza audio. Rigenerare da un PPTX già "
            f"processato non è supportato (limitazione python-pptx)."
        )

    print(f"-> Lettura script da {scripts_xlsx}")
    scripts = read_scripts_from_xlsx(scripts_xlsx)
    # Errore esplicito se il file è vuoto o contiene solo whitespace: senza
    # questo check, process() produrrebbe un PPTX identico all'input senza
    # che l'utente se ne accorga.
    n_non_empty = sum(1 for s in scripts if s)
    if n_non_empty == 0:
        raise ValueError(
            f"Nessuno script trovato in {scripts_xlsx}: il PPTX di output "
            f"sarebbe identico all'input."
        )
    n_scripts = len(scripts)

    print(f"   {n_slides} slide | {n_scripts} script")
    if n_scripts < n_slides:
        print(f"   ! attenzione: solo {n_scripts} script per {n_slides} slide — le slide finali resteranno mute")
    elif n_scripts > n_slides:
        print(f"   ! attenzione: {n_scripts} script ma solo {n_slides} slide — gli script in eccesso vengono ignorati")

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
        if not text:
            print(f"   slide {slide_num}: nessuno script, saltata")
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
        synth_tasks, voice, rate, actual_concurrency,
        progress_callback=cb,
        captured_timings=captured_timings,
        transcode_audio=transcode_audio,
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


def _clone_synth_worker(payload):
    """Sintesi di UNA slide dentro un processo worker. Ritorna una tupla
    picklable così il risultato può tornare al processo padre.
    payload = (slide_num, text, path, voice, rate, pocket_variant, use_cache,
               pocket_quantize)."""
    (slide_num, text, path, voice, rate, pocket_variant, use_cache,
     pocket_quantize) = payload
    try:
        result = voice_clone.synthesize_clone(
            voice, text, rate, path, pocket_variant=pocket_variant,
            use_cache=use_cache, pocket_quantize=pocket_quantize)
        return (slide_num, True, result.sentences, None)
    except Exception as e:
        return (slide_num, False, None, f"{type(e).__name__}: {e}")


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
                r_sn, ok, sentences, error = fut.result()
            except Exception as e:  # crash del worker (raro): lo tratto come errore slide
                r_sn, ok, sentences, error = sn, False, None, f"{type(e).__name__}: {e}"
            if ok:
                captured_timings[r_sn] = {
                    "text": text_by_num.get(r_sn, ""),
                    "sentences": sentences,
                }
                print(f"   slide {r_sn}: sintesi ok "
                      f"({len(text_by_num.get(r_sn, ''))} caratteri, "
                      f"{len(sentences)} frasi)")
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
            result = voice_clone.synthesize_clone(
                cloned_voice, text, rate, path, pocket_variant=pocket_variant,
                use_cache=use_cache, pocket_quantize=pocket_quantize)
            captured_timings[slide_num] = {
                "text": text,
                "sentences": result.sentences,
            }
            elapsed = time.monotonic() - ts
            print(f"   slide {slide_num}: sintesi ok ({elapsed:.1f}s, "
                  f"{len(text)} caratteri, {len(result.sentences)} frasi)")
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


def _run_embedding_phase(
    prs,
    slide_jobs: list[tuple[int, str, str] | None],
    synthesis_errors: dict[int, str],
    autoplay: bool,
    cb,
    auto_advance: bool = False,
) -> list[tuple[int, float | None]]:
    """Embedding sequenziale degli audio nelle slide. Restituisce la lista
    (slide_num, duration|None) per il report durate."""
    print("-> Incorporamento audio nelle slide")
    n_synth_tasks = sum(1 for j in slide_jobs if j is not None)
    n_to_embed = n_synth_tasks - len(synthesis_errors)
    durations: list[tuple[int, float | None]] = []
    embed_done = 0
    for i, slide in enumerate(prs.slides):
        slide_num = i + 1
        job = slide_jobs[i]
        if job is None:
            durations.append((slide_num, None))
            continue
        if slide_num in synthesis_errors:
            print(f"   slide {slide_num}: salto incorporamento (sintesi fallita)")
            durations.append((slide_num, None))
            continue

        _, _, audio_path = job
        duration = get_audio_duration_seconds(audio_path)
        durations.append((slide_num, duration))
        print(f"   slide {slide_num}: durata {format_duration(duration)}, incorporo audio")
        add_audio_to_slide(slide, audio_path, autoplay=autoplay, slide_num=slide_num)
        # Avanzamento automatico a fine audio: ha senso solo con l'autoplay
        # attivo (la slide va avanti da sola quando l'audio finisce).
        if auto_advance and autoplay:
            _set_slide_auto_advance(slide, duration)
        embed_done += 1
        if n_to_embed > 0:
            cb({"stage": "embedding_progress",
                "done": embed_done, "total": n_to_embed})
    return durations


def _write_captions_json(
    output_pptx: str,
    voice: str,
    rate: str,
    captured_timings: dict[int, dict],
) -> None:
    """Salva <output>_captions.json con le frasi sincronizzate per slide.
    Struttura {slides: {"N": {text, sentences: [{text, start_ms, end_ms}]}}}
    — è il contratto col SCORM Builder downstream, non rinominare campi."""
    out_p = Path(output_pptx)
    captions_json_path = out_p.with_name(f"{out_p.stem}_captions.json")
    captions_data = {
        "version": 1,
        "voice": voice,
        "rate": rate,
        "slides": {},
    }
    n_with_captions = 0
    for slide_num, payload in captured_timings.items():
        # Voci clonate: i timing sono già a livello di frase. Voci Microsoft:
        # ricavo le frasi dai WordBoundary catturati durante la sintesi.
        if "sentences" in payload:
            sentences = payload["sentences"]
        else:
            sentences = _sentences_with_timing(
                payload["text"], payload.get("words", [])
            )
        if sentences:
            captions_data["slides"][str(slide_num)] = {
                "text": payload["text"],
                "sentences": sentences,
            }
            n_with_captions += 1
    if n_with_captions > 0:
        with open(captions_json_path, "w", encoding="utf-8") as f:
            json.dump(captions_data, f, ensure_ascii=False, indent=2)
        print(f"-> Captions sincronizzate: {captions_json_path} "
              f"({n_with_captions} slide)")


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
    )
    report_path = _write_duration_report(output_mp4, report_path, durations)
    print("Fatto.")
    cb({"stage": "done"})
    return report_path


def process(
    input_pptx: str,
    scripts_xlsx: str,
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
) -> str:
    """
    Run the full pipeline. Returns the path of the duration report file.

    Two-phase approach:
      1. PARALLEL SYNTHESIS — all MP3 files are generated concurrently
         (capped by `concurrency`) inside a single asyncio event loop.
         This is the dominant cost of the whole pipeline.
      2. SEQUENTIAL EMBEDDING — slides are walked in order and each MP3
         is embedded into the corresponding slide. This phase is fast
         (local I/O only) and order-dependent, so we keep it sequential.

    transcode_audio (default False):
        Se True, dopo la sintesi gli MP3 vengono riconvertiti in
        MPEG-1 @ 44100Hz mono 96kbps via ffmpeg. Migliora la
        compatibilità con alcune versioni di PowerPoint, ma raddoppia
        circa la dimensione dei file audio. Richiede ffmpeg installato.
        Lasciato a False, gli MP3 di Edge TTS (MPEG-2 @ 24kHz, ~50% più
        piccoli) vengono incorporati così come sono.

    progress_callback (optional): called with a dict {"stage": ..., ...}
    on each milestone. Stages:
        loading             — opening pptx
        synthesis_start     — total slides to synthesize
        synthesis_progress  — done / total / slide_num / ok
        synthesis_end       — synthesis phase finished
        embedding_progress  — done / total
        saving              — about to write output pptx
        done                — pipeline complete

    The callback may be invoked from the asyncio event loop thread (during
    synthesis) — GUI clients should marshal updates to their UI thread.
    """
    cb = progress_callback if progress_callback is not None else (lambda e: None)

    # Risolvo la voce: se è una voce clonata la carico dalla libreria; altrimenti
    # resta la stringa della voce Microsoft (edge-tts) e il flusso è quello solito.
    cloned_voice = None
    if _CLONE_AVAILABLE and voice_library.is_clone_voice(voice):
        cloned_voice = voice_library.VoiceLibrary().get(voice)

    prs, scripts, n_slides, n_scripts = _load_inputs(input_pptx, scripts_xlsx, cb)

    with tempfile.TemporaryDirectory() as tmp:
        slide_jobs = _build_slide_jobs(scripts, n_slides, n_scripts, tmp)
        synth_tasks = [job for job in slide_jobs if job is not None]

        if cloned_voice is not None:
            synthesis_errors, captured_timings = _run_synthesis_phase_clone(
                synth_tasks, cloned_voice, rate, cb, pocket_variant=pocket_variant,
                workers=clone_workers, use_cache=use_cache,
                pocket_quantize=pocket_quantize,
            )
        else:
            synthesis_errors, captured_timings = _run_synthesis_phase(
                synth_tasks, voice, rate, concurrency, transcode_audio, cb,
            )

        # Stadio finale: video oppure pptx. Per il video tutto avviene dentro
        # questo blocco temporaneo (servono gli audio appena prodotti).
        if output_mode == "video":
            if not _VIDEO_AVAILABLE:
                raise RuntimeError(
                    "Esportazione video non disponibile: manca il modulo "
                    "video_export (o le sue dipendenze)."
                )
            return _export_video(
                input_pptx, output_pptx, slide_jobs, synthesis_errors,
                captured_timings, resolution, subtitles, transition,
                report_path, tmp, cb,
            )

        durations = _run_embedding_phase(
            prs, slide_jobs, synthesis_errors, autoplay, cb,
            auto_advance=auto_advance,
        )

        cb({"stage": "saving"})
        print(f"-> Salvataggio {output_pptx}")
        prs.save(output_pptx)

    # Post-process: fix the "video" → "audio" relationship type issue caused
    # by python-pptx's add_movie. Without this, PowerPoint shows a corruption
    # warning on open ("PowerPoint ha rilevato un problema nel contenuto") and
    # strips the audio timing on Repair (the audio stops playing automatically).
    fixed, mainseq_fixes = _fix_audio_relationship_types(output_pptx)
    if fixed > 0:
        print(f"-> Relazioni audio ricontrassegnate: {fixed}")
    if mainseq_fixes > 0:
        print(f"-> mainSeq vuoti riparati: {mainseq_fixes}")

    _write_captions_json(output_pptx, voice, rate, captured_timings)
    report_path = _write_duration_report(output_pptx, report_path, durations)

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
    p.add_argument("scripts_xlsx", help="XLSX with scripts in column A (1 row per slide)")
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
    if not Path(args.scripts_xlsx).exists():
        sys.exit(f"Scripts xlsx not found: {args.scripts_xlsx}")

    process(
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
        output_mode=("video" if args.video else "pptx"),
        resolution=args.resolution,
        subtitles=args.subtitles,
        transition=args.transition,
        pocket_variant=args.pocket_variant,
        clone_workers=args.clone_workers,
        use_cache=not args.no_cache,
        pocket_quantize=args.pocket_quantize,
    )


if __name__ == "__main__":
    main()
