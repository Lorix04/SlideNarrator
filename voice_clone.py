"""
voice_clone.py — Backend di sintesi locale (voci clonate) per slide_narrator

Questo è il "secondo strato" della feature: trasforma una voce salvata da
voice_library in audio parlato, usando un modello di cloning che gira in
locale sulla CPU.

Due motori supportati, dietro la stessa astrazione:
- PocketTTS  (Kyutai)   — leggero, CPU-first, veloce. Default consigliato.
- Chatterbox (Resemble) — più espressivo, qualità più alta, ma lento su CPU.

--- Il punto chiave: i sottotitoli ---
Il motore Edge TTS originale fornisce i WordBoundary (timing parola-per-parola)
da cui slide_narrator ricava le frasi sincronizzate per il _captions.json del SCORM
Builder. I modelli di cloning NON danno i WordBoundary. Ma il contratto finale
del JSON è a livello di FRASE, non di parola: quindi qui sintetizziamo una
frase alla volta, misuriamo la durata di ogni clip e ricaviamo i timing frase
per frase per somma cumulativa. Niente forced-alignment, gira bene su CPU, e
produce esattamente la struttura attesa: [{"text","start_ms","end_ms"}].

--- Cosa è testato e cosa no ---
L'orchestrazione condivisa (split frasi, misura durate, timing, concatenazione,
velocità) è coperta da test con un backend finto. Gli adattatori PocketTTS e
Chatterbox sono stati allineati alle firme REALI lette dal codice sorgente
ufficiale dei due pacchetti (TTSModel di kyutai/pocket-tts e
ChatterboxMultilingualTTS di resemble-ai/chatterbox). Restano da provare con i
modelli effettivamente in esecuzione sul tuo PC, perché qui non posso
scaricarli.

--- Hugging Face / accesso ai modelli ---
Al primo avvio entrambi scaricano i pesi da Hugging Face. Per PocketTTS i pesi
CON voice cloning sono in un repo ad accesso riservato (kyutai/pocket-tts):
serve autenticarsi (huggingface-cli login) e ottenere l'accesso, altrimenti il
pacchetto ripiega sui pesi senza cloning e la clonazione fallisce con un errore
chiaro. Chatterbox usa la variabile d'ambiente HF_TOKEN se impostata.

Output: ogni slide diventa un MP3 a 44100 Hz mono (formato già gradito a
PowerPoint), quindi le voci clonate non hanno bisogno del transcoding ffmpeg
opzionale del motore.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import wave
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import voice_library as vl


# ----------------------------------------------------------------------------
# Costanti
# ----------------------------------------------------------------------------
# Formato dell'audio finale incorporato nelle slide.
OUTPUT_SAMPLE_RATE_HZ = 44100
OUTPUT_CHANNELS = 1
OUTPUT_BITRATE = "96k"

# atempo di ffmpeg accetta un fattore tra 0.5 e 2.0 in un singolo stadio.
# Lo slider della GUI (-50%..+50%) resta dentro questo intervallo.
ATEMPO_MIN = 0.5
ATEMPO_MAX = 2.0


# ----------------------------------------------------------------------------
# Risultato della sintesi
# ----------------------------------------------------------------------------
@dataclass
class SynthesisResult:
    """Esito della sintesi di una slide con una voce clonata."""
    audio_path: str          # path del file MP3 prodotto
    duration_s: float        # durata totale dell'audio (secondi)
    sentences: list[dict]    # [{"text","start_ms","end_ms"}] — per i sottotitoli


# ----------------------------------------------------------------------------
# Utilità di testo e tempo
# ----------------------------------------------------------------------------
def split_into_sentences(text: str) -> list[str]:
    """
    Spezza il testo in frasi su . ! ? seguiti da spazio (o fine stringa),
    mantenendo la punteggiatura. Stessa euristica usata da slide_narrator, così i
    sottotitoli restano coerenti tra voci Microsoft e voci clonate.
    Se non c'è alcun terminatore, restituisce l'intero testo come unica frase.
    """
    text = (text or "").strip()
    if not text:
        return []
    parts = re.split(r"(?<=[.!?])\s+", text)
    return [p.strip() for p in parts if p.strip()]


def parse_rate_factor(rate: str) -> float:
    """
    Converte la stringa di velocità in stile Edge TTS ("+0%", "+20%", "-10%")
    nel fattore di velocità per atempo. "+20%" -> 1.20 (più veloce),
    "-10%" -> 0.90 (più lento). Valori fuori range vengono limitati.
    Qualsiasi formato non riconosciuto vale 1.0 (nessuna variazione).
    """
    if not rate:
        return 1.0
    m = re.fullmatch(r"\s*([+-]?\d+)\s*%\s*", rate)
    if not m:
        return 1.0
    factor = 1.0 + int(m.group(1)) / 100.0
    return max(ATEMPO_MIN, min(ATEMPO_MAX, factor))


def _wav_duration_ms(wav_path: str | Path) -> int:
    """Durata di un file WAV in millisecondi (solo standard library)."""
    with wave.open(str(wav_path), "rb") as w:
        frames = w.getnframes()
        rate = w.getframerate()
    if not rate:
        return 0
    return int(round(frames / float(rate) * 1000))


def _concat_and_encode(
    clips: list[str],
    output_mp3: str,
    speed_factor: float = 1.0,
) -> None:
    """
    Concatena i clip WAV (in ordine) in un unico MP3 a 44100 Hz mono, con
    eventuale variazione di velocità (atempo). Usa il demuxer concat di ffmpeg.
    """
    if not clips:
        raise RuntimeError("nessun clip audio da concatenare.")

    with tempfile.TemporaryDirectory() as tmp:
        # File di lista per il demuxer concat. I path vanno tra apici singoli
        # e gli eventuali apici interni raddoppiati (sintassi di ffmpeg).
        list_path = Path(tmp) / "concat.txt"
        with open(list_path, "w", encoding="utf-8") as f:
            for c in clips:
                safe = str(Path(c).resolve()).replace("'", "'\\''")
                f.write(f"file '{safe}'\n")

        cmd = [
            "ffmpeg", "-y", "-loglevel", "error",
            "-f", "concat", "-safe", "0", "-i", str(list_path),
            "-ar", str(OUTPUT_SAMPLE_RATE_HZ), "-ac", str(OUTPUT_CHANNELS),
        ]
        if abs(speed_factor - 1.0) > 1e-3:
            cmd += ["-filter:a", f"atempo={speed_factor:.4f}"]
        cmd += ["-c:a", "libmp3lame", "-b:a", OUTPUT_BITRATE, str(output_mp3)]

        proc = subprocess.run(cmd, capture_output=True)
        if proc.returncode != 0:
            err = proc.stderr.decode(errors="replace")[:300]
            raise RuntimeError(f"concatenazione audio fallita: {err}")


# ----------------------------------------------------------------------------
# Astrazione di backend
# ----------------------------------------------------------------------------
class CloneBackend(ABC):
    """
    Base comune ai motori di cloning. Le sottoclassi implementano solo il
    caricamento del modello e la sintesi di UNA frase; tutta l'orchestrazione
    (split, misura, timing, concatenazione, velocità) è qui e identica per tutti.

    concurrency = 1: questi modelli sono CPU-bound, non ha senso lanciarne
    diversi in parallelo come si fa con Edge TTS (che è limitato dalla rete).
    """
    engine_key: str = ""
    sample_rate: int = 24000   # frequenza dell'audio prodotto dal modello
    concurrency: int = 1

    def __init__(self) -> None:
        self._loaded = False
        # cache dei riferimenti "capati": se un campione è troppo lungo ne uso
        # solo la parte iniziale (vedi _reference_for), calcolata una volta.
        self._capped_refs: dict[str, str] = {}

    # --- da implementare nelle sottoclassi ---------------------------------
    @abstractmethod
    def is_available(self) -> tuple[bool, str]:
        """(installato?, messaggio). Non deve caricare il modello: serve alla
        GUI per dire all'utente se il motore è pronto o cosa installare."""

    @abstractmethod
    def _load(self) -> None:
        """Carica il modello in memoria (operazione costosa, una sola volta)."""

    @abstractmethod
    def _synthesize_sentence(self, text: str, voice: "vl.ClonedVoice",
                             out_wav: str) -> None:
        """Sintetizza UNA frase con la voce data e scrive un WAV mono in out_wav."""

    # --- orchestrazione condivisa ------------------------------------------
    def load(self) -> None:
        """Carica il modello se non già fatto (lazy)."""
        if not self._loaded:
            self._load()
            self._loaded = True

    def _reference_for(self, voice: "vl.ClonedVoice") -> str:
        """Restituisce un campione di riferimento di durata sicura.

        Se il campione supera il tetto (globale o, se più stretto, quello del
        motore) ne produce una copia troncata alla parte iniziale e la riusa.
        Protegge i motori dal saturare la RAM con riferimenti lunghi e, per
        PocketTTS, accorcia il "prompt" così ogni frase si genera più in fretta
        (il modello usa comunque solo ~20 secondi di riferimento).
        """
        # Tetto effettivo: il minimo tra il globale e l'eventuale limite del
        # motore (reference_cap_seconds). PocketTTS usa un tetto più stretto.
        cap = vl.MAX_REFERENCE_SECONDS
        engine_cap = getattr(self, "reference_cap_seconds", None)
        if engine_cap is not None:
            cap = min(cap, engine_cap)

        slug = voice.slug
        cached = self._capped_refs.get(slug)
        if cached is not None:
            return cached
        ref = voice.reference_path
        chosen = ref
        try:
            dur_s = _wav_duration_ms(ref) / 1000.0
            if dur_s > cap + 0.05:
                capped = str(Path(tempfile.gettempdir()) /
                             f"slide_narrator_ref_{slug}_{int(cap)}s.wav")
                _trim_wav_to(ref, capped, cap)
                chosen = capped
        except Exception:
            chosen = ref  # nel dubbio uso l'originale
        self._capped_refs[slug] = chosen
        return chosen

    def synthesize(
        self,
        text: str,
        voice: "vl.ClonedVoice",
        rate: str,
        output_path: str,
    ) -> SynthesisResult:
        """
        Sintetizza l'intero testo di una slide e produce:
          - l'MP3 finale in output_path
          - i timing frase-per-frase per i sottotitoli

        Frasi sintetizzate una alla volta, durate misurate e accumulate; la
        velocità (rate) viene applicata in fase di concatenazione e i timing
        scalati di conseguenza, così restano coerenti con l'audio finale.
        """
        self.load()

        sentences = split_into_sentences(text)
        if not sentences:
            raise ValueError("testo vuoto: niente da sintetizzare.")

        with tempfile.TemporaryDirectory() as tmp:
            clips: list[str] = []
            timings: list[dict] = []
            cursor_ms = 0
            for i, sent in enumerate(sentences):
                clip = str(Path(tmp) / f"sent_{i:04d}.wav")
                self._synthesize_sentence(sent, voice, clip)
                dur_ms = _wav_duration_ms(clip)
                clips.append(clip)
                timings.append({
                    "text": sent,
                    "start_ms": cursor_ms,
                    "end_ms": cursor_ms + dur_ms,
                })
                cursor_ms += dur_ms

            factor = parse_rate_factor(rate)
            _concat_and_encode(clips, output_path, speed_factor=factor)

        # Se c'è variazione di velocità, l'audio finale dura cursor_ms / factor;
        # scalo i timing dello stesso fattore così puntano ai punti giusti.
        if abs(factor - 1.0) > 1e-3:
            inv = 1.0 / factor
            for t in timings:
                t["start_ms"] = int(round(t["start_ms"] * inv))
                t["end_ms"] = int(round(t["end_ms"] * inv))

        total_ms = timings[-1]["end_ms"] if timings else 0
        return SynthesisResult(
            audio_path=output_path,
            duration_s=round(total_ms / 1000.0, 3),
            sentences=timings,
        )


# ----------------------------------------------------------------------------
# Helper: scrittura WAV da array (numpy o torch) per gli adattatori dei modelli
# ----------------------------------------------------------------------------
def _trim_wav_to(src: str, dst: str, max_seconds: float) -> None:
    """Copia in dst solo i primi max_seconds del WAV src (legge i soli frame
    iniziali: rapido anche su file lunghi)."""
    with wave.open(str(src), "rb") as r:
        rate = r.getframerate()
        channels = r.getnchannels()
        sampwidth = r.getsampwidth()
        frames = r.readframes(int(max_seconds * rate))
    with wave.open(str(dst), "wb") as w:
        w.setnchannels(channels)
        w.setsampwidth(sampwidth)
        w.setframerate(rate)
        w.writeframes(frames)


def _write_wav_from_array(path: str, samples, sample_rate: int) -> None:
    """
    Scrive un WAV mono 16-bit da un array di campioni float in [-1, 1].
    Accetta numpy.ndarray o torch.Tensor (li porta su CPU e a numpy).
    """
    import numpy as np  # disponibile se è installato un modello TTS

    arr = samples
    # torch.Tensor -> numpy
    if hasattr(arr, "detach"):
        arr = arr.detach()
    if hasattr(arr, "cpu"):
        arr = arr.cpu()
    if hasattr(arr, "numpy"):
        arr = arr.numpy()
    arr = np.asarray(arr, dtype=np.float32).reshape(-1)

    # clip e conversione a int16
    arr = np.clip(arr, -1.0, 1.0)
    int16 = (arr * 32767.0).astype(np.int16)

    with wave.open(path, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(int(sample_rate))
        w.writeframes(int16.tobytes())


def _concat_audio(segments: list):
    """Concatena una lista di segmenti audio (numpy o torch) lungo il tempo,
    restituendo un unico array numpy float32 1-D. Usato per riunire i sotto-chunk
    di una stessa frase in un solo clip continuo."""
    import numpy as np
    flat = []
    for seg in segments:
        a = seg
        if hasattr(a, "detach"): a = a.detach()
        if hasattr(a, "cpu"): a = a.cpu()
        if hasattr(a, "numpy"): a = a.numpy()
        flat.append(np.asarray(a, dtype=np.float32).reshape(-1))
    if not flat:
        return np.zeros(0, dtype=np.float32)
    return np.concatenate(flat)


# ----------------------------------------------------------------------------
# Adattatore PocketTTS (Kyutai)
# ----------------------------------------------------------------------------
class PocketBackend(CloneBackend):
    """
    Adattatore per PocketTTS. CPU-first, veloce (~6x tempo reale), licenza MIT.

    API di riferimento (pacchetto `pocket-tts`):
        from pocket_tts import TTSModel
        model = TTSModel.load_model()                       # variante lingua
        state = model.get_state_for_audio_prompt(ref_wav)   # impronta voce
        audio = model.generate_audio(state, text)           # campioni audio

    Per l'italiano Kyutai fornisce la variante "italian_24l" (qualità più alta,
    più lenta). Il nome del modello è configurabile qui sotto.
    """
    engine_key = vl.ENGINE_POCKET
    sample_rate = 24000  # impostata da model.sample_rate dopo il load

    # PocketTTS usa solo ~20s di riferimento per l'impronta vocale: accorcio il
    # campione a 24s così ogni frase si genera più in fretta (il "prompt" audio
    # viene ri-attraversato a ogni sintesi). Non intacca la qualità.
    reference_cap_seconds = 24.0

    # PocketTTS avvisa "generation may skip words" se un chunk supera 50 token e
    # non ha punteggiatura interna su cui spezzare. Tengo un margine sotto il 50.
    _POCKET_MAX_TOKENS = 48

    # Modello di lingua da caricare. Per l'italiano PocketTTS offre:
    #   "italian"      -> modello standard, più veloce (consigliato su CPU)
    #   "italian_24l"  -> variante "24 layer", qualità più alta ma più lenta
    language_model = "italian"

    # Quantizzazione INT8 dinamica (sperimentale, spenta di default): può
    # velocizzare la generazione su CPU a scapito di un po' di qualità audio.
    quantize = False

    def __init__(self) -> None:
        super().__init__()
        self._model = None
        # cache delle impronte voce (state) per non ricalcolarle a ogni frase
        self._states: dict[str, object] = {}

    def is_available(self) -> tuple[bool, str]:
        try:
            import pocket_tts  # noqa: F401
            return True, "PocketTTS pronto."
        except Exception:
            return False, ("PocketTTS non installato. "
                           "Installa con:  pip install pocket-tts")

    def _load(self) -> None:
        from pocket_tts import TTSModel
        # load_model accetta 'language' (verificato sul sorgente: i valori
        # ammessi includono "italian" e "italian_24l").
        self._model = TTSModel.load_model(language=self.language_model)
        # La frequenza reale dell'output è esposta dal modello.
        self.sample_rate = self._model.sample_rate
        # I pesi CON voice cloning stanno in un repo HF ad accesso riservato
        # (kyutai/pocket-tts). Se il download fallisce (token mancante o accesso
        # non concesso) il pacchetto ripiega sui pesi SENZA cloning e imposta
        # has_voice_cloning=False: in quel caso clonare una voce è impossibile.
        # Lo segnalo subito con un messaggio chiaro invece di fallire più avanti.
        if not getattr(self._model, "has_voice_cloning", True):
            raise RuntimeError(
                "PocketTTS è stato caricato SENZA supporto al voice cloning "
                "(pesi riservati non scaricati). Autenticati su Hugging Face e "
                "ottieni l'accesso al modello 'kyutai/pocket-tts' "
                "(huggingface-cli login + accettazione delle condizioni)."
            )
        self._apply_quantization()

    def _apply_quantization(self) -> None:
        """Se `quantize` è attivo, applica la quantizzazione dinamica INT8 ai
        layer lineari del transformer (flow_lm). È sperimentale: se non è
        applicabile o torch non la supporta, il modello resta invariato (nessun
        errore). Va provata dall'utente confrontando qualità e velocità."""
        if not getattr(self, "quantize", False):
            return
        try:
            import torch
            target = getattr(self._model, "flow_lm", None)
            if target is not None:
                self._model.flow_lm = torch.quantization.quantize_dynamic(
                    target, {torch.nn.Linear}, dtype=torch.qint8)
                print("   [PocketTTS] quantizzazione INT8 attiva (sperimentale).")
        except Exception as e:
            print(f"   [PocketTTS] quantizzazione non applicata ({type(e).__name__}); "
                  f"uso il modello normale.")

    def _state_for(self, voice: "vl.ClonedVoice"):
        """Impronta vocale per la voce, calcolata una volta e poi riusata."""
        key = voice.slug
        if key not in self._states:
            # get_state_for_audio_prompt accetta Path | str | Tensor; passo un
            # Path così non viene scambiato per il NOME di una voce predefinita.
            # _reference_for garantisce una durata sicura (campioni lunghi
            # verrebbero troncati) per non saturare la RAM.
            self._states[key] = self._model.get_state_for_audio_prompt(
                Path(self._reference_for(voice))
            )
        return self._states[key]

    def _tokenizer(self):
        """Tokenizer interno di PocketTTS (per contare i token come fa il modello).
        Se non accessibile, torna None e si usa una stima."""
        try:
            return self._model.flow_lm.conditioner.tokenizer
        except Exception:
            return None

    def _count_tokens(self, tok, text: str) -> int:
        if tok is not None:
            try:
                return len(tok(text).tokens[0].tolist())
            except Exception:
                pass
        # stima di ripiego: ~1.4 token per parola (subword, italiano) + margine
        return int(len(text.split()) * 1.4) + 1

    def _pocket_chunks(self, text: str) -> list[str]:
        """Spezza il testo in chunk entro il limite di token di PocketTTS.

        PocketTTS da solo divide già su .!?,;: ma NON spezza un tratto lungo
        privo di punteggiatura: lì avvisa e può saltare parole. Qui replico la
        divisione sulla punteggiatura e aggiungo il taglio sui confini di parola
        come ultima risorsa, così nessun chunk supera il limite.
        """
        tok = self._tokenizer()
        limit = self._POCKET_MAX_TOKENS
        # 1) segmenti sulla punteggiatura (mantenendola)
        segments = [s.strip() for s in re.split(r"(?<=[.!?,;:])\s+", text) if s.strip()]
        if not segments:
            segments = [text.strip()]
        # 2) i segmenti ancora troppo lunghi (senza punteggiatura) si spezzano
        #    sui confini di parola
        pieces: list[str] = []
        for seg in segments:
            if self._count_tokens(tok, seg) <= limit:
                pieces.append(seg)
                continue
            cur: list[str] = []
            for w in seg.split():
                cur.append(w)
                if self._count_tokens(tok, " ".join(cur)) > limit:
                    cur.pop()
                    if cur:
                        pieces.append(" ".join(cur))
                    cur = [w]
            if cur:
                pieces.append(" ".join(cur))
        # 3) impacchetto i pezzi in chunk entro il limite
        chunks: list[str] = []
        cur_txt = ""
        cur_n = 0
        for p in pieces:
            n = self._count_tokens(tok, p)
            if cur_txt and cur_n + n > limit:
                chunks.append(cur_txt)
                cur_txt, cur_n = p, n
            else:
                cur_txt = (cur_txt + " " + p).strip() if cur_txt else p
                cur_n += n
        if cur_txt:
            chunks.append(cur_txt)
        return chunks or [text.strip()]

    def _synthesize_sentence(self, text: str, voice: "vl.ClonedVoice",
                             out_wav: str) -> None:
        state = self._state_for(voice)
        chunks = self._pocket_chunks(text)
        # generate_audio(model_state, text) restituisce un torch.Tensor (PCM).
        # Con copy_state=True (default) lo state in cache non viene alterato.
        if len(chunks) == 1:
            audio = self._model.generate_audio(state, chunks[0])
        else:
            # frase lunga divisa in più chunk: sintetizzo ciascuno e li unisco
            # in un unico clip continuo.
            audio = _concat_audio(
                [self._model.generate_audio(state, ch) for ch in chunks]
            )
        _write_wav_from_array(out_wav, audio, self.sample_rate)


# ----------------------------------------------------------------------------
# Adattatore Chatterbox (Resemble AI) — variante Multilingual per l'italiano
# ----------------------------------------------------------------------------
class ChatterboxBackend(CloneBackend):
    """
    Adattatore per Chatterbox Multilingual. Qualità più alta ed espressiva,
    licenza MIT, ma su CPU è lento: adatto a pochi pezzi di pregio o a batch
    lasciati girare. Per l'italiano serve la variante Multilingual (la Turbo
    è solo inglese).

    API di riferimento (pacchetto `chatterbox-tts`):
        from chatterbox.mtl_tts import ChatterboxMultilingualTTS
        model = ChatterboxMultilingualTTS.from_pretrained(device="cpu")
        wav = model.generate(text, language_id="it",
                             audio_prompt_path=ref_wav)     # tensor audio
        # model.sr = frequenza di campionamento dell'output
    """
    engine_key = vl.ENGINE_CHATTERBOX
    sample_rate = 24000  # impostata da model.sr dopo il load

    def __init__(self) -> None:
        super().__init__()
        self._model = None
        # slug della voce per cui i "conditionals" sono già preparati nel
        # modello, così non rileggo il campione di riferimento a ogni frase.
        self._prepared_voice: Optional[str] = None

    def is_available(self) -> tuple[bool, str]:
        try:
            import chatterbox  # noqa: F401
            return True, "Chatterbox pronto."
        except Exception:
            return False, ("Chatterbox non installato. "
                           "Installa con:  pip install chatterbox-tts")

    def _load(self) -> None:
        from chatterbox.mtl_tts import ChatterboxMultilingualTTS
        self._model = ChatterboxMultilingualTTS.from_pretrained(device="cpu")
        self._prepared_voice = None
        # La frequenza reale dell'output è esposta dal modello come .sr
        sr = getattr(self._model, "sr", None)
        if isinstance(sr, int) and sr > 0:
            self.sample_rate = sr

    def _synthesize_sentence(self, text: str, voice: "vl.ClonedVoice",
                             out_wav: str) -> None:
        # Preparo i "conditionals" dal campione UNA volta per voce: le frasi
        # successive li riusano senza rileggere il file di riferimento.
        if self._prepared_voice != voice.slug:
            self._model.prepare_conditionals(self._reference_for(voice))
            self._prepared_voice = voice.slug
        # generate(text, language_id, ...) -> torch.Tensor di forma (1, N).
        # "it" è il codice italiano (verificato su SUPPORTED_LANGUAGES).
        wav = self._model.generate(text, language_id=(voice.language or "it"))
        _write_wav_from_array(out_wav, wav, self.sample_rate)


# ----------------------------------------------------------------------------
# Adattatore XTTS v2 (Coqui / Idiap) — SOLO USO NON COMMERCIALE
# ----------------------------------------------------------------------------
class XttsBackend(CloneBackend):
    """
    Adattatore per XTTS v2. Qualità di clonazione tra le migliori in locale e
    italiano nativo, ma:
      * LICENZA NON COMMERCIALE (Coqui Public Model License). Va usato solo per
        materiale personale/non venduto: per i corsi a pagamento resta PocketTTS.
      * su CPU è LENTO (indicativamente ~10-30s a frase): adatto a batch lasciati
        girare, non all'iterazione veloce.
      * tende a imporre una cadenza/intonazione propria più che quella esatta del
        campione — compromesso noto del modello.

    API di riferimento (pacchetto `coqui-tts`, importabile come `TTS`):
        from TTS.api import TTS
        tts = TTS("tts_models/multilingual/multi-dataset/xtts_v2").to("cpu")
        wav = tts.tts(text=..., speaker_wav=ref, language="it")   # lista di float
    """
    engine_key = vl.ENGINE_XTTS
    sample_rate = 24000  # XTTS v2 produce audio a 24 kHz

    def __init__(self) -> None:
        super().__init__()
        self._model = None

    def is_available(self) -> tuple[bool, str]:
        # Uso find_spec per NON importare il pacchetto (pesante, trascina torch)
        # durante il semplice controllo di disponibilità della GUI.
        import importlib.util
        if importlib.util.find_spec("TTS") is not None:
            return True, "XTTS v2 pronto (solo uso non commerciale)."
        return False, ("XTTS non installato. Installa con:  pip install coqui-tts"
                       "  (modello a licenza NON commerciale).")

    def _load(self) -> None:
        import os
        # XTTS chiede di accettare la licenza CPML al primo caricamento: senza
        # questa variabile il loader resta bloccato in attesa di un "y" a video.
        # L'utente ha scelto esplicitamente XTTS per uso NON commerciale.
        os.environ.setdefault("COQUI_TOS_AGREED", "1")

        # --- Compatibilità transformers 5.x --------------------------------
        # coqui-tts (XTTS) importa `isin_mps_friendly` da transformers.pytorch_utils,
        # funzione RIMOSSA in transformers 5.x. Qui è installata la 5.2.0 (la
        # pretende chatterbox-tts), quindi l'import di XTTS fallirebbe. Reinietto
        # una versione equivalente PRIMA di importare qualsiasi modulo di TTS,
        # senza toccare transformers (così Chatterbox resta compatibile).
        try:
            import transformers.pytorch_utils as _tpu
            if not hasattr(_tpu, "isin_mps_friendly"):
                import torch as _t

                def isin_mps_friendly(elements, test_elements):
                    # Su MPS torch.isin non è supportato: emulazione; altrove
                    # (CPU/CUDA, il tuo caso) è direttamente torch.isin.
                    if getattr(elements, "device", None) is not None and \
                            elements.device.type == "mps":
                        test_elements = test_elements.to(elements.device)
                        return elements.unsqueeze(-1).eq(test_elements).any(dim=-1)
                    return _t.isin(elements, test_elements)

                _tpu.isin_mps_friendly = isin_mps_friendly
        except Exception:
            # Se transformers cambia ancora, l'errore vero riemergerà sotto e
            # verrà mostrato nel log: non lo nascondo con un fallback silenzioso.
            pass

        # --- Evito la dipendenza da torchcodec -----------------------------
        # Da PyTorch 2.9 torchaudio delega l'I/O audio a torchcodec, che su
        # Windows richiede le librerie SHARED di FFmpeg (le DLL) ed è fragile.
        # A XTTS torchaudio serve solo per CARICARE il wav di riferimento: lo
        # sostituisco con soundfile (già installato), che legge il nostro wav
        # mono senza toccare torchcodec. La sintesi la scriviamo noi a parte.
        try:
            import torchaudio
            import soundfile as _sf
            import numpy as _np
            import torch as _t

            def _load_via_soundfile(filepath, *args, **kwargs):
                data, sr = _sf.read(str(filepath), dtype="float32", always_2d=True)
                # soundfile: (campioni, canali) -> torchaudio: (canali, campioni)
                wav = _t.from_numpy(_np.ascontiguousarray(data.T))
                return wav, sr

            torchaudio.load = _load_via_soundfile
        except Exception:
            # Se torchaudio non è importabile o XTTS usa un'altra via, l'errore
            # originale riemergerà e sarà visibile nel log.
            pass

        # --- Supero il "gate" torchcodec di coqui-tts ----------------------
        # TTS/__init__.py, con torch>=2.9, si rifiuta di importarsi se torchcodec
        # non è installato: `if not is_torchcodec_available(): raise`. Quella
        # verifica controlla solo se il PACCHETTO esiste, non lo usa. Visto che il
        # caricamento audio l'ho già dirottato su soundfile, torchcodec non serve:
        # dico alla funzione di transformers che è disponibile, così l'import di
        # TTS passa senza installare torchcodec né le DLL "shared" di FFmpeg.
        try:
            import transformers.utils.import_utils as _tiu
            _tiu.is_torchcodec_available = lambda: True
        except Exception:
            pass

        # PyTorch recente (>=2.6) usa weights_only=True come default in torch.load
        # e rifiuta i checkpoint XTTS (che contengono oggetti di config, non solo
        # pesi). Il modello XTTS è quello UFFICIALE scelto dall'utente, quindi
        # durante SOLO questo caricamento ripristino weights_only=False. (coqui
        # registra già i propri "safe globals" all'import del pacchetto.)
        import torch

        _orig_load = torch.load
        def _load_full(*args, **kwargs):
            kwargs.setdefault("weights_only", False)
            return _orig_load(*args, **kwargs)

        from TTS.api import TTS
        torch.load = _load_full
        try:
            self._model = TTS("tts_models/multilingual/multi-dataset/xtts_v2").to("cpu")
        finally:
            torch.load = _orig_load

    def _synthesize_sentence(self, text: str, voice: "vl.ClonedVoice",
                             out_wav: str) -> None:
        # Le frasi sono già divise a monte (sotto il limite di 250 caratteri di
        # XTTS), quindi disattivo lo split interno del modello. _reference_for
        # garantisce un campione di durata sicura.
        wav = self._model.tts(
            text=text,
            speaker_wav=self._reference_for(voice),
            language=(voice.language or "it"),
            split_sentences=False,
        )
        _write_wav_from_array(out_wav, wav, self.sample_rate)


# ----------------------------------------------------------------------------
# Registro / factory dei backend
# ----------------------------------------------------------------------------
_BACKEND_CLASSES = {
    vl.ENGINE_POCKET: PocketBackend,
    vl.ENGINE_CHATTERBOX: ChatterboxBackend,
    vl.ENGINE_XTTS: XttsBackend,
}

# I backend vengono creati una sola volta e riusati (tengono il modello in RAM).
_BACKEND_CACHE: dict[str, CloneBackend] = {}

# Varianti del modello italiano di PocketTTS.
POCKET_VARIANT_FAST = "italian"       # standard, più veloce (default su CPU)
POCKET_VARIANT_HQ = "italian_24l"     # "24 layer", qualità più alta ma più lenta
POCKET_VARIANTS = (POCKET_VARIANT_FAST, POCKET_VARIANT_HQ)


def get_backend(engine_key: str, variant: str | None = None,
                quantize: bool = False) -> CloneBackend:
    """Restituisce (creando al bisogno) il backend per il motore richiesto.

    Per PocketTTS `variant` seleziona il modello di lingua ("italian" oppure
    "italian_24l") e `quantize` attiva la quantizzazione INT8 (sperimentale):
    ogni combinazione è un'istanza distinta in cache, così si può passare
    dall'una all'altra senza che si sovrascrivano il modello caricato.
    """
    if engine_key not in _BACKEND_CLASSES:
        raise ValueError(
            f"motore di cloning sconosciuto: {engine_key!r} "
            f"(ammessi: {', '.join(_BACKEND_CLASSES)})."
        )
    # Variante e quantizzazione hanno senso solo per PocketTTS.
    if engine_key == vl.ENGINE_POCKET:
        if variant not in POCKET_VARIANTS:
            variant = None
        cache_key = engine_key
        if variant is not None:
            cache_key += f":{variant}"
        if quantize:
            cache_key += ":q"
    else:
        cache_key = engine_key
        variant = None
        quantize = False
    if cache_key not in _BACKEND_CACHE:
        inst = _BACKEND_CLASSES[engine_key]()
        if variant is not None:
            inst.language_model = variant
        if quantize:
            inst.quantize = True
        _BACKEND_CACHE[cache_key] = inst
    return _BACKEND_CACHE[cache_key]


def available_backends() -> dict[str, tuple[bool, str]]:
    """Per la GUI: {motore: (installato?, messaggio)} per ciascun backend."""
    out: dict[str, tuple[bool, str]] = {}
    for key, cls in _BACKEND_CLASSES.items():
        try:
            out[key] = cls().is_available()
        except Exception as e:
            out[key] = (False, f"errore nel controllo: {e}")
    return out


def synthesize_clone(
    voice: "vl.ClonedVoice",
    text: str,
    rate: str,
    output_path: str,
    pocket_variant: str | None = None,
    use_cache: bool = True,
    pocket_quantize: bool = False,
) -> SynthesisResult:
    """
    Punto d'ingresso ad alto livello: sceglie il backend dalla voce e sintetizza.
    Sarà chiamato dal motore (slide_narrator) per le slide con voce clonata.

    `pocket_variant` seleziona la qualità per le voci PocketTTS ("italian" o
    "italian_24l"); `pocket_quantize` attiva la quantizzazione INT8 sperimentale.
    Entrambi vengono ignorati per gli altri motori.

    Con `use_cache` (default True) l'audio di una slide identica (stesso motore,
    variante, quantizzazione, voce, velocità e testo) viene riusato da disco
    invece di essere rigenerato: enorme risparmio quando si rigenera un corso
    dopo piccole modifiche. La cache si invalida da sé se cambia uno di quei
    fattori.
    """
    is_pocket = voice.engine == vl.ENGINE_POCKET
    variant = pocket_variant if is_pocket else None
    quantize = bool(pocket_quantize) if is_pocket else False

    key = None
    if use_cache:
        try:
            key = _cache_key(voice, variant, rate, text, quantize)
            hit = _cache_lookup(key, output_path)
            if hit is not None:
                return hit
        except Exception:
            key = None  # in caso di problemi con la cache, sintetizzo e basta

    backend = get_backend(voice.engine, variant, quantize)
    result = backend.synthesize(text, voice, rate, output_path)

    if use_cache and key is not None:
        try:
            _cache_store(key, output_path, result)
        except Exception:
            pass  # un errore nel salvare in cache non deve far fallire la sintesi
    return result


# ----------------------------------------------------------------------------
# Cache dei risultati: riusa l'audio di slide identiche tra un run e l'altro
# ----------------------------------------------------------------------------
_CACHE_VERSION = "v1"          # cambiare se muta la logica di sintesi/formato
_reference_hash_memo: dict[str, str] = {}


def get_cache_dir() -> Path:
    """Cartella della cache. Sovrascrivibile con la variabile d'ambiente
    SLIDENARRATOR_CACHE_DIR; di default una cartella nascosta nella home utente."""
    env = os.environ.get("SLIDENARRATOR_CACHE_DIR")
    base = Path(env) if env else (Path.home() / ".slide_narrator_cache")
    base.mkdir(parents=True, exist_ok=True)
    return base


def _reference_hash(path: str) -> str:
    """SHA-256 del file di riferimento, memoizzato per (path, dimensione, mtime)
    così non si rilegge il wav a ogni slide ma si ricalcola se il file cambia."""
    try:
        st = os.stat(path)
        memo_key = f"{path}:{st.st_size}:{int(st.st_mtime)}"
    except OSError:
        memo_key = path
    cached = _reference_hash_memo.get(memo_key)
    if cached is not None:
        return cached
    h = hashlib.sha256()
    try:
        with open(path, "rb") as f:
            for block in iter(lambda: f.read(1 << 20), b""):
                h.update(block)
        digest = h.hexdigest()
    except OSError:
        digest = "no-ref"
    _reference_hash_memo[memo_key] = digest
    return digest


def _cache_key(voice: "vl.ClonedVoice", variant: str | None,
               rate: str, text: str, quantize: bool = False) -> str:
    """Chiave che dipende da TUTTO ciò che determina l'audio prodotto."""
    material = "\x00".join([
        _CACHE_VERSION,
        voice.engine or "",
        variant or "",
        "q1" if quantize else "q0",
        _reference_hash(voice.reference_path),
        rate or "",
        text or "",
    ])
    return hashlib.sha256(material.encode("utf-8")).hexdigest()


def _cache_lookup(key: str, output_path: str) -> Optional[SynthesisResult]:
    """Se in cache c'è l'audio per questa chiave, lo copia in output_path e
    ricostruisce il SynthesisResult (timing inclusi). Altrimenti None."""
    d = get_cache_dir()
    mp3 = d / f"{key}.mp3"
    meta = d / f"{key}.json"
    if not (mp3.exists() and meta.exists()):
        return None
    try:
        info = json.loads(meta.read_text(encoding="utf-8"))
    except Exception:
        return None
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(mp3, output_path)
    return SynthesisResult(
        audio_path=output_path,
        duration_s=float(info.get("duration_s", 0.0)),
        sentences=list(info.get("sentences", [])),
    )


def _cache_store(key: str, output_path: str, result: SynthesisResult) -> None:
    """Salva in cache l'audio prodotto e i suoi metadati (scrittura atomica)."""
    d = get_cache_dir()
    mp3 = d / f"{key}.mp3"
    meta = d / f"{key}.json"
    tmp_mp3 = d / f"{key}.mp3.tmp"
    tmp_meta = d / f"{key}.json.tmp"
    shutil.copyfile(output_path, tmp_mp3)
    tmp_meta.write_text(json.dumps({
        "duration_s": result.duration_s,
        "sentences": result.sentences,
    }), encoding="utf-8")
    os.replace(tmp_mp3, mp3)     # rename atomico: niente entry a metà
    os.replace(tmp_meta, meta)


def cleanup_cache(max_age_days: int = 30, max_bytes: int = 5 * 1024**3) -> dict:
    """Rimuove entry vecchie e limita la cache alla dimensione indicata."""
    d = get_cache_dir()
    now = time.time()
    removed_files = 0
    removed_bytes = 0
    entries = []
    for mp3 in d.glob("*.mp3"):
        key = mp3.stem
        meta = d / f"{key}.json"
        try:
            st = mp3.stat()
        except OSError:
            continue
        entries.append((st.st_mtime, st.st_size, mp3, meta))
    cutoff = now - max(1, max_age_days) * 86400
    kept = []
    for mtime, size, mp3, meta in entries:
        if mtime < cutoff:
            for file in (mp3, meta):
                try:
                    removed_bytes += file.stat().st_size
                    file.unlink(); removed_files += 1
                except OSError:
                    pass
        else:
            kept.append((mtime, size, mp3, meta))
    total = sum(size + (meta.stat().st_size if meta.exists() else 0)
                for _, size, _, meta in kept)
    for mtime, size, mp3, meta in sorted(kept):
        if total <= max_bytes:
            break
        entry_size = size + (meta.stat().st_size if meta.exists() else 0)
        for file in (mp3, meta):
            try:
                removed_bytes += file.stat().st_size
                file.unlink(); removed_files += 1
            except OSError:
                pass
        total -= entry_size
    for tmp in d.glob("*.tmp"):
        try:
            if tmp.stat().st_mtime < now - 86400:
                removed_bytes += tmp.stat().st_size
                tmp.unlink(); removed_files += 1
        except OSError:
            pass
    return {"removed_files": removed_files, "removed_bytes": removed_bytes,
            "remaining_bytes": max(0, total)}


def clear_cache() -> int:
    """Svuota la cache. Restituisce il numero di clip audio rimosse."""
    d = get_cache_dir()
    removed = 0
    for f in d.glob("*.mp3"):
        try:
            f.unlink(); removed += 1
        except OSError:
            pass
    for f in list(d.glob("*.json")) + list(d.glob("*.tmp")):
        try:
            f.unlink()
        except OSError:
            pass
    return removed
