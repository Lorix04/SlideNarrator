"""
voice_library.py — Gestione delle voci clonate per slide_narrator

Questo è il "primo strato" della feature di voice cloning: si occupa SOLO di
archiviare e gestire le voci clonate dall'utente. Non sintetizza nulla e non
carica modelli TTS: quello è compito di voice_clone.py (lo strato successivo).

Cosa fa:
- mantiene una cartella `voices/` con una sottocartella per ogni voce salvata
- per ogni voce conserva il campione audio di riferimento (reference.wav,
  normalizzato a mono) più un meta.json con i metadati
- offre operazioni CRUD: aggiungi, elenca, leggi, rinomina, elimina

Schema dell'identità voce (lo stesso che userà il routing nel motore):
- voci Microsoft (edge-tts):  "it-IT-IsabellaNeural"   (NON gestite qui)
- voci clonate:               "clone:<motore>:<slug>"   es. "clone:pocket:giuseppe"

Struttura su disco:
    voices/
        giuseppe/
            reference.wav
            meta.json
        giuseppe_hd/
            reference.wav
            meta.json

Dipendenze: solo standard library. ffmpeg è opzionale ma raccomandato: se
presente, il campione viene riconvertito in WAV mono pulito a una frequenza
costante; altrimenti i file .wav vengono copiati così come sono (e i formati
non-wav richiedono ffmpeg).
"""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import unicodedata
import wave
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


# ----------------------------------------------------------------------------
# Costanti
# ----------------------------------------------------------------------------
# Motori di sintesi supportati per le voci clonate.
ENGINE_POCKET = "pocket"
ENGINE_CHATTERBOX = "chatterbox"
ENGINE_XTTS = "xtts"
SUPPORTED_ENGINES = (ENGINE_POCKET, ENGINE_CHATTERBOX, ENGINE_XTTS)

# Etichette leggibili (per la GUI / i badge).
ENGINE_LABELS = {
    ENGINE_POCKET: "PocketTTS",
    ENGINE_CHATTERBOX: "Chatterbox",
    ENGINE_XTTS: "XTTS v2 (non commerciale)",
}

# Prefisso che distingue una voce clonata da una voce Microsoft.
CLONE_PREFIX = "clone:"

# Formato di riferimento prodotto quando ffmpeg è disponibile: WAV PCM mono.
# 24 kHz è coerente con il resto del progetto (l'audio nativo di Edge TTS è a
# 24 kHz) ed è più che sufficiente per estrarre l'impronta vocale: i modelli
# di cloning ricampionano comunque internamente.
REFERENCE_SAMPLE_RATE_HZ = 24000
REFERENCE_FILENAME = "reference.wav"
META_FILENAME = "meta.json"

# Versione dello schema di meta.json, per compatibilità futura.
META_SCHEMA_VERSION = 1

# Durata del campione: sotto MIN è quasi certamente un errore dell'utente;
# sotto RECOMMENDED la clonazione funziona ma la resa cala. La libreria non
# blocca su RECOMMENDED (lascia decidere alla GUI), blocca solo sotto MIN.
MIN_REFERENCE_SECONDS = 1.0
RECOMMENDED_REFERENCE_SECONDS = 30.0

# Tetto massimo: oltre questa durata il campione viene troncato alla parte
# iniziale. Serve a proteggere i motori di cloning (soprattutto PocketTTS su
# CPU): un riferimento lungo va "digerito" a ogni sintesi e con campioni di
# molti minuti la RAM esplode e il PC si blocca. 60 secondi sono un tetto
# prudente (i motori usano comunque solo ~12-20s per l'impronta vocale).
MAX_REFERENCE_SECONDS = 60.0

# Posizione di default della libreria: cartella `voices/` accanto a questo file.
DEFAULT_LIBRARY_DIR = Path(__file__).resolve().parent / "voices"


# ----------------------------------------------------------------------------
# Eccezioni
# ----------------------------------------------------------------------------
class VoiceLibraryError(Exception):
    """Errore generico della libreria voci."""


class VoiceNotFoundError(VoiceLibraryError):
    """La voce richiesta non esiste nella libreria."""


# ----------------------------------------------------------------------------
# Identità voce: clone:<motore>:<slug>
# ----------------------------------------------------------------------------
def is_clone_voice(voice_id: str) -> bool:
    """True se l'id rappresenta una voce clonata (non una voce Microsoft)."""
    return isinstance(voice_id, str) and voice_id.startswith(CLONE_PREFIX)


def make_clone_id(engine: str, slug: str) -> str:
    """Costruisce l'id voce nel formato clone:<motore>:<slug>."""
    return f"{CLONE_PREFIX}{engine}:{slug}"


def parse_clone_id(voice_id: str) -> tuple[str, str]:
    """
    Scompone un id voce clonata in (motore, slug).
    Solleva ValueError se l'id non è nel formato atteso.
    """
    if not is_clone_voice(voice_id):
        raise ValueError(f"non è un id di voce clonata: {voice_id!r}")
    rest = voice_id[len(CLONE_PREFIX):]
    parts = rest.split(":", 1)
    if len(parts) != 2 or not parts[0] or not parts[1]:
        raise ValueError(f"id voce clonata malformato: {voice_id!r}")
    engine, slug = parts
    return engine, slug


def slugify(name: str) -> str:
    """
    Trasforma un nome leggibile in uno slug sicuro per il filesystem e per
    l'id voce: minuscolo, solo [a-z0-9_], senza accenti.
    Es. "Giuseppe (la mia voce)" -> "giuseppe_la_mia_voce".
    """
    # Rimuove gli accenti (à -> a) decomponendo in Unicode.
    normalized = unicodedata.normalize("NFKD", name)
    ascii_name = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_name = ascii_name.lower().strip()
    # Sostituisce qualsiasi sequenza non alfanumerica con un singolo underscore.
    slug = re.sub(r"[^a-z0-9]+", "_", ascii_name).strip("_")
    return slug or "voce"


# ----------------------------------------------------------------------------
# Modello dati
# ----------------------------------------------------------------------------
@dataclass
class ClonedVoice:
    """Una voce clonata salvata nella libreria."""
    slug: str            # identificativo stabile, sicuro per il filesystem
    name: str            # nome visualizzato (può contenere spazi, accenti...)
    engine: str          # uno tra SUPPORTED_ENGINES
    language: str        # codice lingua, es. "it"
    reference_path: str  # path assoluto al reference.wav
    duration_s: float    # durata del campione in secondi
    sample_rate: int     # frequenza di campionamento del reference.wav
    channels: int        # numero di canali del reference.wav (1 = mono)
    created_at: str      # data di creazione in ISO 8601 (UTC)

    @property
    def voice_id(self) -> str:
        """Id completo nel formato clone:<motore>:<slug>."""
        return make_clone_id(self.engine, self.slug)

    @property
    def engine_label(self) -> str:
        """Etichetta leggibile del motore (per la GUI)."""
        return ENGINE_LABELS.get(self.engine, self.engine)

    @property
    def is_short(self) -> bool:
        """True se il campione è più corto della durata raccomandata."""
        return self.duration_s < RECOMMENDED_REFERENCE_SECONDS

    def to_meta(self) -> dict:
        """Dizionario serializzabile per meta.json (senza il path assoluto,
        che viene ricostruito dalla cartella della voce)."""
        d = asdict(self)
        d.pop("reference_path", None)
        d["schema_version"] = META_SCHEMA_VERSION
        return d


# ----------------------------------------------------------------------------
# Utilità audio
# ----------------------------------------------------------------------------
def _ffmpeg_available() -> bool:
    return shutil.which("ffmpeg") is not None


def _wav_info(wav_path: Path) -> tuple[float, int, int]:
    """Restituisce (durata_s, sample_rate, channels) leggendo un file WAV con
    la sola standard library."""
    with wave.open(str(wav_path), "rb") as w:
        frames = w.getnframes()
        rate = w.getframerate()
        channels = w.getnchannels()
        duration = frames / float(rate) if rate else 0.0
    return duration, rate, channels


def _trim_wav_in_place(path: Path, max_seconds: float) -> None:
    """Tronca un WAV alla durata massima indicata, tenendo solo la parte
    iniziale. Legge solo i primi frame, quindi è veloce anche su file lunghi."""
    path = Path(path)
    with wave.open(str(path), "rb") as r:
        rate = r.getframerate()
        channels = r.getnchannels()
        sampwidth = r.getsampwidth()
        frames = r.readframes(int(max_seconds * rate))
    tmp = path.with_suffix(path.suffix + ".tmp")
    with wave.open(str(tmp), "wb") as w:
        w.setnchannels(channels)
        w.setsampwidth(sampwidth)
        w.setframerate(rate)
        w.writeframes(frames)
    tmp.replace(path)


def _prepare_reference(source_audio: Path, dest_wav: Path) -> tuple[float, int, int]:
    """
    Prepara il campione di riferimento in dest_wav e ne restituisce
    (durata_s, sample_rate, channels).

    Strategia:
      1. se ffmpeg è disponibile, riconverte qualsiasi formato in WAV PCM
         mono a REFERENCE_SAMPLE_RATE_HZ (pulito e uniforme);
      2. altrimenti, se la sorgente è già .wav, la copia così com'è;
      3. altrimenti (formato non-wav senza ffmpeg) solleva un errore chiaro.
    """
    if _ffmpeg_available():
        proc = subprocess.run(
            ["ffmpeg", "-y", "-loglevel", "error",
             "-i", str(source_audio),
             "-t", str(MAX_REFERENCE_SECONDS),
             "-ac", "1",
             "-ar", str(REFERENCE_SAMPLE_RATE_HZ),
             "-c:a", "pcm_s16le",
             str(dest_wav)],
            capture_output=True,
        )
        if proc.returncode != 0:
            err = proc.stderr.decode(errors="replace")[:300]
            raise VoiceLibraryError(f"conversione del campione fallita: {err}")
    else:
        if source_audio.suffix.lower() != ".wav":
            raise VoiceLibraryError(
                "ffmpeg non è disponibile: posso importare solo file .wav. "
                "Installa ffmpeg per usare anche mp3/m4a/altri formati."
            )
        shutil.copyfile(source_audio, dest_wav)

    if not dest_wav.exists() or dest_wav.stat().st_size == 0:
        raise VoiceLibraryError("il campione di riferimento risulta vuoto.")

    try:
        duration, rate, channels = _wav_info(dest_wav)
    except (wave.Error, EOFError) as e:
        raise VoiceLibraryError(f"il file audio non è un WAV valido: {e}") from e

    # Tetto di sicurezza: se il campione è più lungo del massimo (es. importato
    # senza ffmpeg, oppure ffmpeg non ha troncato) ne tengo solo la parte
    # iniziale. Evita che i motori di cloning saturino la RAM.
    if duration > MAX_REFERENCE_SECONDS + 0.05:
        _trim_wav_in_place(dest_wav, MAX_REFERENCE_SECONDS)
        duration, rate, channels = _wav_info(dest_wav)

    if duration < MIN_REFERENCE_SECONDS:
        raise VoiceLibraryError(
            f"il campione dura solo {duration:.1f}s: troppo poco per clonare "
            f"una voce. Fornisci un audio più lungo (ideale 15-30 secondi)."
        )
    return duration, rate, channels


# ----------------------------------------------------------------------------
# Libreria voci
# ----------------------------------------------------------------------------
class VoiceLibrary:
    """
    Gestisce la cartella `voices/` e le voci clonate al suo interno.

    La fonte di verità è il filesystem: ogni voce è una sottocartella che
    contiene reference.wav + meta.json. list_voices() scandisce la cartella e
    rilegge i meta.json, così non c'è un indice separato da tenere sincronizzato.
    """

    def __init__(self, root: Optional[str | Path] = None) -> None:
        self.root = Path(root) if root is not None else DEFAULT_LIBRARY_DIR
        self.root.mkdir(parents=True, exist_ok=True)

    # ----- percorsi interni -------------------------------------------------
    def _voice_dir(self, slug: str) -> Path:
        return self.root / slug

    def _meta_path(self, slug: str) -> Path:
        return self._voice_dir(slug) / META_FILENAME

    def _reference_path(self, slug: str) -> Path:
        return self._voice_dir(slug) / REFERENCE_FILENAME

    # ----- lettura ----------------------------------------------------------
    def _load_voice(self, slug: str) -> ClonedVoice:
        """Carica una voce dal suo meta.json. Solleva VoiceNotFoundError se
        manca o è illeggibile."""
        meta_path = self._meta_path(slug)
        if not meta_path.exists():
            raise VoiceNotFoundError(f"voce non trovata: {slug!r}")
        try:
            with open(meta_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            raise VoiceNotFoundError(
                f"meta.json illeggibile per {slug!r}: {e}"
            ) from e
        return ClonedVoice(
            slug=data["slug"],
            name=data["name"],
            engine=data["engine"],
            language=data.get("language", "it"),
            reference_path=str(self._reference_path(slug)),
            duration_s=float(data.get("duration_s", 0.0)),
            sample_rate=int(data.get("sample_rate", REFERENCE_SAMPLE_RATE_HZ)),
            channels=int(data.get("channels", 1)),
            created_at=data.get("created_at", ""),
        )

    def list_voices(self) -> list[ClonedVoice]:
        """Elenca tutte le voci salvate, ordinate per nome visualizzato."""
        voices: list[ClonedVoice] = []
        for child in sorted(self.root.iterdir()):
            if not child.is_dir():
                continue
            if not (child / META_FILENAME).exists():
                continue
            try:
                voices.append(self._load_voice(child.name))
            except VoiceNotFoundError:
                # Cartella corrotta: la salto senza far fallire tutto l'elenco.
                continue
        voices.sort(key=lambda v: v.name.lower())
        return voices

    def get(self, voice_id_or_slug: str) -> ClonedVoice:
        """
        Restituisce una voce dato il suo id completo (clone:motore:slug) oppure
        il solo slug. Solleva VoiceNotFoundError se non esiste.
        """
        if is_clone_voice(voice_id_or_slug):
            _, slug = parse_clone_id(voice_id_or_slug)
        else:
            slug = voice_id_or_slug
        return self._load_voice(slug)

    def exists(self, slug: str) -> bool:
        return self._meta_path(slug).exists()

    # ----- scrittura --------------------------------------------------------
    def _unique_slug(self, base_slug: str) -> str:
        """Garantisce uno slug non già in uso, aggiungendo _2, _3, ... se serve."""
        if not self.exists(base_slug):
            return base_slug
        i = 2
        while self.exists(f"{base_slug}_{i}"):
            i += 1
        return f"{base_slug}_{i}"

    def add_voice(
        self,
        name: str,
        engine: str,
        source_audio: str | Path,
        language: str = "it",
    ) -> ClonedVoice:
        """
        Aggiunge una nuova voce clonata alla libreria.

        - name:         nome visualizzato (può avere spazi/accenti)
        - engine:       uno tra SUPPORTED_ENGINES ("pocket" / "chatterbox")
        - source_audio: percorso del campione audio dell'utente (wav/mp3/...)
        - language:     codice lingua del campione (default "it")

        Copia/normalizza il campione in reference.wav, ne misura la durata e
        scrive meta.json. Restituisce la ClonedVoice creata.
        """
        name = (name or "").strip()
        if not name:
            raise VoiceLibraryError("il nome della voce non può essere vuoto.")
        if engine not in SUPPORTED_ENGINES:
            raise VoiceLibraryError(
                f"motore non supportato: {engine!r} "
                f"(ammessi: {', '.join(SUPPORTED_ENGINES)})."
            )
        source_path = Path(source_audio)
        if not source_path.exists():
            raise VoiceLibraryError(f"file audio non trovato: {source_path}")

        slug = self._unique_slug(slugify(name))
        voice_dir = self._voice_dir(slug)
        voice_dir.mkdir(parents=True, exist_ok=False)

        try:
            reference = self._reference_path(slug)
            duration, rate, channels = _prepare_reference(source_path, reference)
            voice = ClonedVoice(
                slug=slug,
                name=name,
                engine=engine,
                language=language,
                reference_path=str(reference),
                duration_s=round(duration, 2),
                sample_rate=rate,
                channels=channels,
                created_at=datetime.now(timezone.utc).isoformat(timespec="seconds"),
            )
            self._write_meta(voice)
        except Exception:
            # Pulizia: se qualcosa va storto a metà, non lascio cartelle monche.
            shutil.rmtree(voice_dir, ignore_errors=True)
            raise
        return voice

    def _write_meta(self, voice: ClonedVoice) -> None:
        with open(self._meta_path(voice.slug), "w", encoding="utf-8") as f:
            json.dump(voice.to_meta(), f, ensure_ascii=False, indent=2)

    def rename_voice(self, slug: str, new_name: str) -> ClonedVoice:
        """
        Cambia il nome visualizzato di una voce. Lo slug (e quindi l'id e i
        percorsi su disco) resta invariato, così i riferimenti già salvati
        altrove continuano a funzionare.
        """
        new_name = (new_name or "").strip()
        if not new_name:
            raise VoiceLibraryError("il nuovo nome non può essere vuoto.")
        voice = self._load_voice(slug)
        voice.name = new_name
        self._write_meta(voice)
        return voice

    def delete_voice(self, slug: str) -> None:
        """Elimina definitivamente una voce (cartella, campione e metadati)."""
        voice_dir = self._voice_dir(slug)
        if not voice_dir.exists():
            raise VoiceNotFoundError(f"voce non trovata: {slug!r}")
        shutil.rmtree(voice_dir)


# ----------------------------------------------------------------------------
# Mini CLI per test manuali (facoltativa)
#   python voice_library.py list
#   python voice_library.py add "Giuseppe" pocket /percorso/campione.wav
#   python voice_library.py delete giuseppe
# ----------------------------------------------------------------------------
def _main() -> None:
    import sys

    lib = VoiceLibrary()
    args = sys.argv[1:]
    cmd = args[0] if args else "list"

    if cmd == "list":
        voices = lib.list_voices()
        if not voices:
            print("Nessuna voce clonata salvata.")
            return
        for v in voices:
            short = "  (campione corto)" if v.is_short else ""
            print(f"{v.voice_id:<32} {v.name}  [{v.engine_label}]  "
                  f"{v.duration_s:.0f}s{short}")
    elif cmd == "add" and len(args) >= 4:
        name, engine, source = args[1], args[2], args[3]
        v = lib.add_voice(name=name, engine=engine, source_audio=source)
        print(f"Aggiunta: {v.voice_id}  ({v.duration_s:.0f}s @ {v.sample_rate} Hz)")
    elif cmd == "delete" and len(args) >= 2:
        lib.delete_voice(args[1])
        print(f"Eliminata: {args[1]}")
    else:
        print("Uso: list | add <nome> <pocket|chatterbox> <audio> | delete <slug>")


if __name__ == "__main__":
    _main()
