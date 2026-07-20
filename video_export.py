"""
video_export.py — Esporta le slide + audio TTS in un video MP4

Riusa l'audio già prodotto dal motore (una traccia per slide) e costruisce un
filmato in cui ogni slide resta a schermo per la durata del suo audio.

Pipeline:
  1. RENDER  — le slide vengono disegnate in immagini PNG. python-pptx non sa
     renderizzare, quindi si usa LibreOffice in headless: pptx -> pdf, poi
     ogni pagina del pdf -> PNG (con PyMuPDF; in mancanza, pdftoppm/poppler).
  2. CLIP    — per ogni slide un clip = immagine (per la durata del suo audio)
     + l'audio, scalata e centrata nella risoluzione scelta (ffmpeg).
  3. CONCAT  — i clip vengono concatenati in un unico MP4 (H.264 + AAC).
  4. SRT     — dai timing per frase (già calcolati) si genera un file .srt;
     opzionalmente i sottotitoli vengono "bruciati" nel video.

Le slide senza audio restano a schermo per DEFAULT_SILENT_SLIDE_S secondi.

Dipendenze: LibreOffice (rendering) + ffmpeg (montaggio) + PyMuPDF (pdf->png).
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Callable, Optional


# ----------------------------------------------------------------------------
# Costanti
# ----------------------------------------------------------------------------
VIDEO_FPS = 25
AUDIO_SAMPLE_RATE_HZ = 44100
AUDIO_BITRATE = "192k"

# Preset di codifica x264. Per immagini statiche un preset veloce non cambia la
# qualità percepita (il fotogramma non si muove) ma accelera molto la codifica.
VIDEO_PRESET = os.environ.get("PPTXTTS_VIDEO_PRESET", "veryfast")
VIDEO_CRF = os.environ.get("PPTXTTS_VIDEO_CRF", "23")

# Quanti clip creare in parallelo (ffmpeg indipendenti). None = auto in base
# ai core. Ogni ffmpeg è già multi-thread, quindi non conviene esagerare.
_CLIP_WORKERS_ENV = os.environ.get("PPTXTTS_CLIP_WORKERS", "").strip()
try:
    CLIP_WORKERS: Optional[int] = int(_CLIP_WORKERS_ENV) if _CLIP_WORKERS_ENV else None
except ValueError:
    CLIP_WORKERS = None

# Durata a schermo delle slide senza audio.
DEFAULT_SILENT_SLIDE_S = 3.0

# Durata della dissolvenza incrociata tra una slide e la successiva.
TRANSITION_DURATION_S = 0.5

# Risoluzioni offerte (16:9). Le slide vengono scalate e centrate (con bande
# nere se l'aspetto non combacia).
RESOLUTIONS = {
    "720p": (1280, 720),
    "1080p": (1920, 1080),
}
DEFAULT_RESOLUTION = "1080p"

# Timeout generosi per i sottoprocessi (conversione e codifica).
SOFFICE_TIMEOUT_S = int(os.environ.get("PPTXTTS_SOFFICE_TIMEOUT_S", "180"))
FFMPEG_TIMEOUT_S = int(os.environ.get("PPTXTTS_FFMPEG_TIMEOUT_S", "600"))
FFMPEG_LONG_TIMEOUT_MIN_S = int(
    os.environ.get("PPTXTTS_FFMPEG_LONG_TIMEOUT_MIN_S", "1800")
)
FFMPEG_LONG_TIMEOUT_FACTOR = float(
    os.environ.get("PPTXTTS_FFMPEG_LONG_TIMEOUT_FACTOR", "4.0")
)


ProgressCallback = Optional[Callable[[dict], None]]


class VideoExportError(Exception):
    """Errore durante l'esportazione video."""


# ----------------------------------------------------------------------------
# Utilità sottoprocessi
# ----------------------------------------------------------------------------
def _run(cmd: list[str], timeout: int, what: str) -> None:
    """Esegue un comando e solleva VideoExportError con stderr se fallisce."""
    try:
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise VideoExportError(f"{what}: timeout dopo {timeout}s.")
    except FileNotFoundError:
        raise VideoExportError(f"{what}: comando non trovato ({cmd[0]}).")
    if proc.returncode != 0:
        err = proc.stderr.decode(errors="replace")[:400]
        raise VideoExportError(f"{what} fallito: {err}")


def _long_timeout_for(duration_s: float) -> int:
    """Timeout per operazioni che processano tutta la timeline video."""
    duration_s = max(0.0, float(duration_s or 0.0))
    return max(
        FFMPEG_TIMEOUT_S,
        int(FFMPEG_LONG_TIMEOUT_MIN_S + duration_s * FFMPEG_LONG_TIMEOUT_FACTOR),
    )


def _replace_output(src: str | Path, dst: str | Path) -> None:
    """Sposta `src` su `dst`, sostituendo il file finale solo a operazione finita."""
    src_p = Path(src)
    dst_p = Path(dst)
    dst_p.parent.mkdir(parents=True, exist_ok=True)
    try:
        os.replace(src_p, dst_p)
    except OSError:
        if dst_p.exists():
            dst_p.unlink()
        shutil.move(str(src_p), str(dst_p))


def _sidecar_path(path: str | Path, suffix: str) -> Path:
    p = Path(path)
    return p.with_name(f"{p.stem}{suffix}{p.suffix}")


def _find_soffice() -> Optional[str]:
    """Cerca l'eseguibile di LibreOffice nel PATH e nei percorsi comuni Windows."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


# ----------------------------------------------------------------------------
# 1. Rendering delle slide in immagini
# ----------------------------------------------------------------------------
def pptx_to_pdf(pptx_path: str | Path, work_dir: str | Path) -> Path:
    """Converte il pptx in PDF con LibreOffice headless. Restituisce il path PDF."""
    soffice = _find_soffice()
    if not soffice:
        raise VideoExportError(
            "LibreOffice non trovato: serve per disegnare le slide del video. "
            "Installalo da https://www.libreoffice.org/ e riprova."
        )
    work_dir = Path(work_dir)
    # Profilo utente temporaneo: evita conflitti se un'altra istanza di
    # LibreOffice è già aperta.
    profile = (work_dir / "lo_profile").as_uri()
    _run(
        [soffice, "--headless", "--norestore",
         f"-env:UserInstallation={profile}",
         "--convert-to", "pdf", "--outdir", str(work_dir), str(pptx_path)],
        timeout=SOFFICE_TIMEOUT_S, what="Conversione pptx->pdf (LibreOffice)",
    )
    pdf_path = work_dir / (Path(pptx_path).stem + ".pdf")
    if not pdf_path.exists():
        raise VideoExportError("LibreOffice non ha prodotto il PDF atteso.")
    return pdf_path


def pdf_to_images(pdf_path: str | Path, work_dir: str | Path,
                  target_width: int) -> list[Path]:
    """Rasterizza ogni pagina del PDF in un PNG, largo circa target_width px.
    Usa PyMuPDF se disponibile, altrimenti pdftoppm (poppler)."""
    pdf_path = Path(pdf_path)
    work_dir = Path(work_dir)
    images: list[Path] = []

    try:
        import fitz  # PyMuPDF
    except Exception:
        fitz = None

    if fitz is not None:
        doc = fitz.open(str(pdf_path))
        try:
            for i, page in enumerate(doc):
                # zoom così che la larghezza renderizzata ~ target_width
                page_w = page.rect.width or target_width
                zoom = max(1.0, target_width / page_w)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                out = work_dir / f"slide_{i + 1:03d}.png"
                pix.save(str(out))
                images.append(out)
        finally:
            doc.close()
        return images

    # Fallback: pdftoppm
    if shutil.which("pdftoppm"):
        prefix = work_dir / "slide"
        _run(["pdftoppm", "-png", "-scale-to-x", str(target_width),
              "-scale-to-y", "-1", str(pdf_path), str(prefix)],
             timeout=FFMPEG_TIMEOUT_S, what="Rasterizzazione PDF (pdftoppm)")
        images = sorted(work_dir.glob("slide-*.png"))
        if images:
            return images

    raise VideoExportError(
        "Nessuno strumento per convertire il PDF in immagini. "
        "Installa PyMuPDF con:  pip install pymupdf"
    )


def render_slides_to_images(pptx_path: str | Path, work_dir: str | Path,
                            target_width: int) -> list[Path]:
    """pptx -> pdf -> una immagine PNG per slide (in ordine)."""
    pdf = pptx_to_pdf(pptx_path, work_dir)
    return pdf_to_images(pdf, work_dir, target_width)


# ----------------------------------------------------------------------------
# 2. Costruzione dei clip per-slide
# ----------------------------------------------------------------------------
def _scale_pad_filter(w: int, h: int) -> str:
    """Filtro video: scala mantenendo le proporzioni e centra con bande nere."""
    return (f"scale={w}:{h}:force_original_aspect_ratio=decrease,"
            f"pad={w}:{h}:(ow-iw)/2:(oh-ih)/2:color=black")


def _escape_filter_path(path: str | Path) -> str:
    """Escape minimale per path Windows dentro un filtro ffmpeg."""
    return (str(Path(path).resolve())
            .replace("\\", "\\\\")
            .replace(":", "\\:")
            .replace("'", "\\'"))


def _subtitles_filter(srt_path: str | Path) -> str:
    return f"subtitles='{_escape_filter_path(srt_path)}'"


def build_slide_clip(image: str | Path, audio: str | Path | None,
                     duration_s: float, out_clip: str | Path,
                     w: int, h: int,
                     subtitle_srt: str | Path | None = None) -> None:
    """Crea un clip video: l'immagine resta a schermo per la durata dell'audio
    (o per duration_s se la slide è senza audio, con silenzio)."""
    vf = _scale_pad_filter(w, h)
    if subtitle_srt:
        vf = f"{vf},{_subtitles_filter(subtitle_srt)}"
    base = ["ffmpeg", "-y", "-loglevel", "error",
            "-loop", "1", "-framerate", str(VIDEO_FPS), "-i", str(image)]
    if audio:
        cmd = base + [
            "-i", str(audio),
            "-vf", vf, "-c:v", "libx264", "-preset", VIDEO_PRESET,
            "-crf", VIDEO_CRF, "-tune", "stillimage",
            "-r", str(VIDEO_FPS), "-pix_fmt", "yuv420p",
            "-c:a", "aac", "-ar", str(AUDIO_SAMPLE_RATE_HZ), "-ac", "2",
            "-b:a", AUDIO_BITRATE,
            "-shortest", str(out_clip),
        ]
    else:
        cmd = base + [
            "-f", "lavfi", "-i",
            f"anullsrc=channel_layout=stereo:sample_rate={AUDIO_SAMPLE_RATE_HZ}",
            "-vf", vf, "-c:v", "libx264", "-preset", VIDEO_PRESET,
            "-crf", VIDEO_CRF, "-tune", "stillimage",
            "-r", str(VIDEO_FPS), "-pix_fmt", "yuv420p",
            "-c:a", "aac", "-ar", str(AUDIO_SAMPLE_RATE_HZ), "-ac", "2",
            "-b:a", AUDIO_BITRATE,
            "-t", f"{duration_s:.3f}", str(out_clip),
        ]
    _run(cmd, timeout=FFMPEG_TIMEOUT_S, what="Creazione clip slide (ffmpeg)")


def concat_clips(clips: list[str | Path], out_mp4: str | Path,
                 work_dir: str | Path) -> None:
    """Concatena i clip (stessi parametri di codifica) in un unico MP4."""
    if not clips:
        raise VideoExportError("Nessun clip da concatenare.")
    list_path = Path(work_dir) / "concat.txt"
    with open(list_path, "w", encoding="utf-8") as f:
        for c in clips:
            safe = str(Path(c).resolve()).replace("'", "'\\''")
            f.write(f"file '{safe}'\n")
    _run(["ffmpeg", "-y", "-loglevel", "error",
          "-f", "concat", "-safe", "0", "-i", str(list_path),
          "-c", "copy", "-movflags", "+faststart", str(out_mp4)],
         timeout=FFMPEG_TIMEOUT_S, what="Concatenazione clip (ffmpeg)")


def render_with_transitions(metas: list[dict], out_video: str | Path,
                            w: int, h: int, transition_s: float,
                            work_dir: str | Path) -> None:
    """Monta il video con dissolvenza incrociata tra le slide.

    Idea chiave: ogni slide riceve una piccola coda di `transition_s` secondi.
    La dissolvenza video avviene DENTRO quella coda, mentre l'audio (narrazione)
    resta una sequenza senza sovrapposizioni: quando una slide sfuma nella
    successiva, il parlato della prima è già finito e parte quello della seconda.

    `metas`: lista per-slide con {"image", "audio"(path|None), "c"(durata
    contenuto in s)}.
    """
    n = len(metas)
    if n < 2:
        raise VideoExportError("Le transizioni richiedono almeno 2 slide.")
    T = float(transition_s)
    work = Path(work_dir)
    total_duration = sum(float(m["c"]) for m in metas)

    # --- 1. Video con dissolvenze (xfade), senza audio --------------------
    # Ogni input è l'immagine tenuta a schermo per (contenuto + coda T).
    v_inputs: list[str] = []
    for m in metas:
        v_inputs += ["-loop", "1", "-t", f"{m['c'] + T:.3f}", "-i", str(m["image"])]

    fc: list[str] = []
    for i in range(n):
        fc.append(
            f"[{i}:v]scale={w}:{h}:force_original_aspect_ratio=decrease,"
            f"pad={w}:{h}:(ow-iw)/2:(oh-ih)/2,fps={VIDEO_FPS},"
            f"format=yuv420p,setsar=1[v{i}]"
        )
    # Catena xfade: l'offset di ogni transizione è la lunghezza accumulata - T.
    cum = metas[0]["c"] + T
    prev = "[v0]"
    for k in range(1, n):
        offset = cum - T
        out_label = "[vout]" if k == n - 1 else f"[x{k}]"
        fc.append(
            f"{prev}[v{k}]xfade=transition=fade:duration={T:.3f}:"
            f"offset={offset:.3f}{out_label}"
        )
        cum += (metas[k]["c"] + T) - T
        prev = out_label

    silent_video = work / "vx_silent.mp4"
    _run(["ffmpeg", "-y", "-loglevel", "error", *v_inputs,
          "-filter_complex", ";".join(fc), "-map", prev,
          "-c:v", "libx264", "-preset", VIDEO_PRESET, "-crf", VIDEO_CRF,
          "-r", str(VIDEO_FPS), "-pix_fmt", "yuv420p", str(silent_video)],
         timeout=_long_timeout_for(total_duration),
         what="Transizioni video (ffmpeg xfade)")

    # --- 2. Traccia audio: narrazioni in sequenza (no sovrapposizioni) ----
    a_inputs: list[str] = []
    for m in metas:
        if m["audio"]:
            a_inputs += ["-i", str(m["audio"])]
        else:
            a_inputs += ["-f", "lavfi", "-t", f"{m['c']:.3f}", "-i",
                         f"anullsrc=channel_layout=stereo:"
                         f"sample_rate={AUDIO_SAMPLE_RATE_HZ}"]
    af: list[str] = []
    fmt = (f"aresample={AUDIO_SAMPLE_RATE_HZ},"
           f"aformat=sample_fmts=fltp:channel_layouts=stereo")
    for i, m in enumerate(metas):
        if m["audio"]:
            af.append(f"[{i}:a]{fmt},apad,atrim=0:{m['c']:.3f},"
                      f"asetpts=PTS-STARTPTS[a{i}]")
        else:
            af.append(f"[{i}:a]{fmt},asetpts=PTS-STARTPTS[a{i}]")
    concat_in = "".join(f"[a{i}]" for i in range(n))
    af_str = ";".join(af) + f";{concat_in}concat=n={n}:v=0:a=1[aout]"

    audio_track = work / "vx_audio.m4a"
    _run(["ffmpeg", "-y", "-loglevel", "error", *a_inputs,
          "-filter_complex", af_str, "-map", "[aout]",
          "-c:a", "aac", "-b:a", AUDIO_BITRATE, str(audio_track)],
         timeout=_long_timeout_for(total_duration),
         what="Traccia audio (ffmpeg)")

    # --- 3. Unione video + audio (la coda muta finale viene rifilata) ------
    _run(["ffmpeg", "-y", "-loglevel", "error",
          "-i", str(silent_video), "-i", str(audio_track),
          "-map", "0:v:0", "-map", "1:a:0", "-c:v", "copy", "-c:a", "copy",
          "-shortest", "-movflags", "+faststart", str(out_video)],
         timeout=_long_timeout_for(total_duration),
         what="Unione video+audio (ffmpeg)")


# ----------------------------------------------------------------------------
# 4. Sottotitoli (SRT) dai timing per frase
# ----------------------------------------------------------------------------
def _ms_to_srt_time(ms: int) -> str:
    ms = max(0, int(ms))
    h, rem = divmod(ms, 3_600_000)
    m, rem = divmod(rem, 60_000)
    s, ms = divmod(rem, 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def generate_srt(slides: list[dict]) -> str:
    """
    Costruisce il testo SRT per l'intero video. `slides` è la lista (in ordine)
    dei dizionari per-slide con:
        - "clip_ms": durata del clip della slide nel video finale
        - "sentences": [{"text","start_ms","end_ms"}] relativi all'inizio della
          slide (così come prodotti dal motore)
    I timing di ogni frase vengono spostati dell'offset cumulativo della slide.
    """
    lines: list[str] = []
    idx = 1
    offset_ms = 0
    for s in slides:
        for sent in s.get("sentences", []) or []:
            start = offset_ms + int(sent["start_ms"])
            end = offset_ms + int(sent["end_ms"])
            lines.append(str(idx))
            lines.append(f"{_ms_to_srt_time(start)} --> {_ms_to_srt_time(end)}")
            lines.append(sent["text"])
            lines.append("")
            idx += 1
        offset_ms += int(s.get("clip_ms", 0))
    return "\n".join(lines)


def write_slide_srt(sentences: list[dict], srt_path: str | Path) -> Path | None:
    """Scrive un SRT temporaneo con timing relativi alla singola slide."""
    if not sentences:
        return None
    text = generate_srt([{"clip_ms": 0, "sentences": sentences}])
    if not text.strip():
        return None
    out = Path(srt_path)
    with open(out, "w", encoding="utf-8") as f:
        f.write(text)
    return out


def burn_subtitles(in_mp4: str | Path, srt_path: str | Path,
                   out_mp4: str | Path,
                   duration_s: float | None = None) -> None:
    """Brucia i sottotitoli SRT nel video (ri-codifica del solo video)."""
    out = Path(out_mp4)
    tmp = _sidecar_path(out, ".__subtmp__")
    if tmp.exists():
        tmp.unlink()
    try:
        _run(["ffmpeg", "-y", "-loglevel", "error", "-i", str(in_mp4),
              "-vf", _subtitles_filter(srt_path),
              "-c:v", "libx264", "-preset", VIDEO_PRESET, "-crf", VIDEO_CRF,
              "-c:a", "copy", "-movflags", "+faststart", str(tmp)],
             timeout=_long_timeout_for(duration_s or 0.0),
             what="Sottotitoli nel video (ffmpeg)")
        _replace_output(tmp, out)
    finally:
        if tmp.exists():
            tmp.unlink()


# ----------------------------------------------------------------------------
# Orchestrazione
# ----------------------------------------------------------------------------
def _resolve_clip_workers(n: int) -> int:
    """Numero di clip da creare in parallelo. Default: fino a 4, senza superare i
    core o il numero di slide (ogni ffmpeg è già multi-thread, meglio non
    esagerare per non saturare la CPU)."""
    if CLIP_WORKERS is not None:
        return max(1, min(int(CLIP_WORKERS), n))
    cores = os.cpu_count() or 2
    return max(1, min(cores, 4, n))


def build_clips_parallel(metas: list[dict], work: Path, w: int, h: int, cb):
    """Crea un clip per slide, in parallelo (ffmpeg indipendenti), e restituisce
    la lista ORDINATA dei clip. L'ordine finale è garantito dall'indice anche se
    i lavori terminano in ordine sparso."""
    import concurrent.futures as cf
    n = len(metas)
    clips = [work / f"clip_{i + 1:03d}.mp4" for i in range(n)]
    workers = _resolve_clip_workers(n)

    def _one(i: int) -> int:
        m = metas[i]
        build_slide_clip(
            m["image"], m["audio"], m["c"], clips[i], w, h,
            m.get("subtitle_srt"),
        )
        return i

    def _report(i: int, done: int) -> None:
        cb({"stage": "clip_progress", "done": done, "total": n})
        m = metas[i]
        print(f"   slide {i + 1}/{n}: clip creato "
              f"({'audio' if m['audio'] else 'muta'}, {m['c']:.1f}s)")

    done = 0
    if workers <= 1:
        for i in range(n):
            _one(i)
            done += 1
            _report(i, done)
        return clips

    print(f"   creazione clip in parallelo ({workers} alla volta)…")
    with cf.ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(_one, i): i for i in range(n)}
        for fut in cf.as_completed(futs):
            i = futs[fut]
            fut.result()  # propaga eventuali errori dal thread
            done += 1
            _report(i, done)
    return clips


def build_video(
    pptx_path: str | Path,
    slides: list[dict],
    output_mp4: str | Path,
    resolution: str = DEFAULT_RESOLUTION,
    subtitles: bool = False,
    transition: bool = False,
    work_dir: str | Path | None = None,
    progress_callback: ProgressCallback = None,
) -> str:
    """
    Costruisce il video MP4.

    `slides` è la lista (allineata alle slide del pptx, in ordine) di dizionari:
        {
          "audio": path mp3 | None,
          "duration_s": float | None,   # durata audio; ignorata se audio è None
          "sentences": [{"text","start_ms","end_ms"}],  # per i sottotitoli
        }

    Se `transition` è True, tra una slide e l'altra c'è una dissolvenza
    incrociata (il parlato resta comunque sequenziale, senza sovrapposizioni).

    Restituisce il path del video prodotto. Se `subtitles` è True, i sottotitoli
    vengono bruciati nel video; in ogni caso viene scritto un .srt a fianco.
    """
    cb = progress_callback if progress_callback is not None else (lambda e: None)
    if resolution not in RESOLUTIONS:
        resolution = DEFAULT_RESOLUTION
    w, h = RESOLUTIONS[resolution]
    output_mp4 = str(output_mp4)

    owns_tmp = work_dir is None
    tmp_ctx = tempfile.TemporaryDirectory() if owns_tmp else None
    work = Path(tmp_ctx.name) if owns_tmp else Path(work_dir)

    try:
        # 1. RENDER
        cb({"stage": "render_start"})
        print(f"-> Rendering slide in immagini ({w}x{h}) con LibreOffice…")
        images = render_slides_to_images(pptx_path, work, target_width=w)
        if not images:
            raise VideoExportError("Nessuna immagine prodotta dal rendering.")
        n = min(len(images), len(slides))
        if len(images) != len(slides):
            print(f"   ! attenzione: {len(images)} slide renderizzate vs "
                  f"{len(slides)} previste; uso le prime {n}.")
        print(f"   {len(images)} slide renderizzate.")

        # 2. Preparo i metadati per slide (durata contenuto + sottotitoli).
        metas: list[dict] = []
        srt_slides: list[dict] = []
        for i in range(n):
            s = slides[i]
            audio = s.get("audio")
            duration = (float(s.get("duration_s") or 0.0) if audio
                        else DEFAULT_SILENT_SLIDE_S)
            duration = max(0.1, duration)
            metas.append({"image": images[i], "audio": audio, "c": duration})
            srt_slides.append({
                "clip_ms": int(round(duration * 1000)),
                "sentences": s.get("sentences", []),
            })

        bare_video = work / "video_nosub.mp4"
        use_transitions = transition and n >= 2
        total_duration = sum(float(m["c"]) for m in metas)
        burn_subtitles_at_end = subtitles and use_transitions

        if use_transitions:
            cb({"stage": "muxing"})
            print(f"-> Montaggio del video con dissolvenze "
                  f"({TRANSITION_DURATION_S:.1f}s)…")
            render_with_transitions(metas, bare_video, w, h,
                                    TRANSITION_DURATION_S, work)
        else:
            if subtitles:
                wrote_any = False
                for i, (m, srt_slide) in enumerate(zip(metas, srt_slides)):
                    srt = write_slide_srt(
                        srt_slide.get("sentences", []),
                        work / f"slide_sub_{i + 1:03d}.srt",
                    )
                    if srt is not None:
                        m["subtitle_srt"] = srt
                        wrote_any = True
                if wrote_any:
                    print("-> Sottotitoli applicati durante la creazione dei clip.")
            # Un clip per slide (in parallelo), poi concatenazione (taglio netto).
            clips = build_clips_parallel(metas, work, w, h, cb)
            cb({"stage": "muxing"})
            print("-> Montaggio del video…")
            concat_clips(clips, bare_video, work)

        # 4. SRT (sempre) + eventuale burn-in
        srt_text = generate_srt(srt_slides)
        srt_path = Path(output_mp4).with_suffix(".srt")
        if srt_text.strip():
            tmp_srt = _sidecar_path(srt_path, ".__tmp__")
            with open(tmp_srt, "w", encoding="utf-8") as f:
                f.write(srt_text)
            _replace_output(tmp_srt, srt_path)
            print(f"-> Sottotitoli: {srt_path}")

        if burn_subtitles_at_end and srt_text.strip():
            cb({"stage": "subtitles"})
            print("-> Sottotitoli nel video…")
            try:
                burn_subtitles(bare_video, srt_path, output_mp4, total_duration)
            except VideoExportError as e:
                fallback = _sidecar_path(output_mp4, "_senza_sottotitoli")
                _replace_output(bare_video, fallback)
                raise VideoExportError(
                    f"{e}\n"
                    f"Video senza sottotitoli salvato qui: {fallback}\n"
                    f"Sottotitoli salvati qui: {srt_path}"
                ) from e
        else:
            _replace_output(bare_video, output_mp4)

        print(f"-> Video pronto: {output_mp4}")
        return output_mp4
    finally:
        if tmp_ctx is not None:
            tmp_ctx.cleanup()
