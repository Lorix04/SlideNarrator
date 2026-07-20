#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sonda empirica del limite di token "per chunk" di PocketTTS.

Il software di PocketTTS spezza il testo in chunk da massimo 50 token, ma quel 50
è un default prudente scritto nel codice, non un limite documentato del modello.
Questo strumento verifica, sul campo, fino a quanti token il MODELLO regge
davvero prima di iniziare a "saltare parole", e confronta le due varianti
italiane:  'italian' (Veloce) e 'italian_24l' (Qualità alta).

COME FUNZIONA
  1. costruisce frasi SENZA punteggiatura, di lunghezza crescente misurata in
     token con il tokenizer del modello;
  2. le passa al modello come CHUNK UNICO -> usa max_tokens molto alto, così il
     software NON le spezza e il modello è costretto a generarle tutte insieme;
  3. misura quanta parte dell'input viene effettivamente pronunciata.

MISURA DELLA RESA (due metodi, sceglie il migliore disponibile)
  - Whisper (se installato: 'faster-whisper' oppure 'openai-whisper'):
    trascrive l'audio e calcola la frazione di parole ritrovate. È il metodo
    diretto e affidabile.
  - Euristica sulla DURATA (nessuna dipendenza extra): stabilisce un ritmo
    (secondi per token) sui chunk piccoli — che sicuramente sono resi bene — e
    poi segnala quando l'audio prodotto è troppo corto rispetto all'atteso,
    segno che il modello ha tagliato dei token.

USO
  python test_limite_token.py --voice "campione.wav"
  python test_limite_token.py --voice alba --languages italian italian_24l
  python test_limite_token.py --voice "campione.wav" --min 20 --max 130 --step 10

Suggerimento: mettilo nella cartella dell'app (accanto a voice_clone.py) e
lancialo con la stessa .venv, così usa il pocket-tts già installato.
"""
from __future__ import annotations

import argparse
import difflib
import os
import re
import statistics
import shutil
import subprocess
import sys
import tempfile
import wave
from pathlib import Path

# Parole italiane comuni, senza punteggiatura: servono solo a "riempire" frasi
# di una data lunghezza in token. Il senso non conta, conta la lunghezza.
WORD_POOL = (
    "la casa nuova sopra la collina guarda il mare calmo mentre il vento porta "
    "un odore di pioggia lontana e le persone camminano piano lungo la strada "
    "stretta parlando di cose semplici come il lavoro la famiglia i giorni "
    "passati insieme e quelli ancora da vivere con calma e con fiducia nel "
    "futuro che arriva sempre un poco alla volta come le onde sulla sabbia "
    "bianca sotto un cielo grande e sereno pieno di luce e di silenzio gentile "
    "ordine tempo numero parola voce suono ombra luce porta finestra tavolo "
    "sedia libro pagina inizio fine mezzo strada ponte fiume monte valle bosco"
).split()


# ----------------------------------------------------------------------------
# Utilità indipendenti dal modello (verificabili senza PocketTTS)
# ----------------------------------------------------------------------------
def count_tokens(tokenizer, text: str) -> int:
    """Numero di token secondo il tokenizer del modello."""
    return len(tokenizer(text).tokens[0].tolist())


def build_text(tokenizer, target_tokens: int) -> tuple[str, list[str], int]:
    """Costruisce una frase senza punteggiatura di ~target_tokens token.
    Ritorna (testo, elenco_parole, token_effettivi)."""
    words: list[str] = []
    i = 0
    while True:
        words.append(WORD_POOL[i % len(WORD_POOL)])
        i += 1
        if i % 4 == 0 or len(words) >= 4:  # controlla ogni tanto per non rallentare
            n = count_tokens(tokenizer, " ".join(words))
            if n >= target_tokens:
                break
        if len(words) > target_tokens * 4 + 20:  # guardia di sicurezza
            break
    text = " ".join(words)
    return text, words, count_tokens(tokenizer, text)


_word_re = re.compile(r"[a-zàèéìòóù]+", re.IGNORECASE)


def normalize_words(text: str) -> list[str]:
    return _word_re.findall(text.lower())


def retention(expected_words: list[str], transcript: str) -> float:
    """Frazione delle parole attese effettivamente presenti nell'audio (0..1),
    calcolata allineando le due sequenze (robusto a errori/ordine dell'ASR)."""
    exp = [w.lower() for w in expected_words]
    hyp = normalize_words(transcript)
    if not exp:
        return 0.0
    sm = difflib.SequenceMatcher(None, exp, hyp, autojunk=False)
    matched = sum(block.size for block in sm.get_matching_blocks())
    return matched / len(exp)


def wav_duration_s(path: str | Path) -> float:
    with wave.open(str(path), "rb") as w:
        return w.getnframes() / float(w.getframerate() or 1)


def analyze_duration(rows: list[dict]) -> int | None:
    """Euristica: ritmo (s/token) dai chunk piccoli, poi primo punto in cui la
    durata reale scende sotto l'85% di quella attesa -> il modello ha tagliato."""
    small = [r for r in rows if r["tokens"] <= 30 and r["dur"] > 0]
    if len(small) < 2:
        small = rows[:2]
    ratios = [r["dur"] / r["tokens"] for r in small if r["tokens"]]
    if not ratios:
        return None
    base = statistics.median(ratios)
    for r in rows:
        expected = base * r["tokens"]
        if expected > 0 and r["dur"] < 0.85 * expected:
            return r["tokens"]
    return None


def analyze_retention(rows: list[dict]) -> int | None:
    """Con ASR: primo punto in cui la resa scende di oltre 10 punti sotto la
    resa di riferimento (media dei chunk piccoli)."""
    small = [r for r in rows if r["tokens"] <= 30 and r["ret"] is not None]
    if len(small) < 2:
        small = [r for r in rows if r["ret"] is not None][:2]
    if not small:
        return None
    base = statistics.mean(r["ret"] for r in small)
    for r in rows:
        if r["ret"] is not None and r["ret"] < base - 0.10:
            return r["tokens"]
    return None


# ----------------------------------------------------------------------------
# ASR opzionale (Whisper): misura diretta del "salto parole"
# ----------------------------------------------------------------------------
def get_asr():
    """Ritorna una funzione trascrivi(wav_path)->str, o None se Whisper manca."""
    try:
        from faster_whisper import WhisperModel
        model = WhisperModel("small", device="cpu", compute_type="int8")
        def _tx(path):
            segments, _ = model.transcribe(str(path), language="it")
            return " ".join(s.text for s in segments)
        print("ASR: uso faster-whisper (modello 'small').")
        return _tx
    except Exception:
        pass
    try:
        import whisper
        model = whisper.load_model("small")
        def _tx(path):
            return model.transcribe(str(path), language="it").get("text", "")
        print("ASR: uso openai-whisper (modello 'small').")
        return _tx
    except Exception:
        pass
    print("ASR non disponibile (installa 'faster-whisper' per la misura diretta). "
          "Uso l'euristica sulla durata.")
    return None


# ----------------------------------------------------------------------------
# Sonda vera e propria (richiede pocket-tts installato)
# ----------------------------------------------------------------------------
def prepare_voice(voice: str) -> str:
    """PocketTTS legge meglio i file WAV. Se `voice` è un file audio non-wav
    (es. .mp3, .m4a), lo converte in un wav mono temporaneo con ffmpeg. Le voci
    predefinite (nomi come 'giovanni') e gli URL hf:// restano invariati."""
    if voice.startswith("hf://") or voice.startswith("http"):
        return voice
    p = Path(voice)
    if not p.exists():
        return voice  # probabilmente il nome di una voce predefinita
    if p.suffix.lower() == ".wav":
        return str(p)
    if not shutil.which("ffmpeg"):
        print(f"(attenzione: {p.suffix} non convertito, ffmpeg non trovato)")
        return str(p)
    out = Path(tempfile.gettempdir()) / (p.stem + "_prova.wav")
    try:
        subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", str(p),
                        "-ac", "1", "-ar", "24000", str(out)], check=True)
        print(f"Voce '{p.name}' convertita in wav: {out}")
        return str(out)
    except Exception as e:
        print(f"(conversione voce fallita: {e}; provo col file originale)")
        return str(p)


def probe_variant(language: str, voice: str, points: list[int], asr, out_dir: Path):
    from pocket_tts import TTSModel

    print(f"\n=== Variante: {language} ===")
    print(f"Carico il modello ({language})…")
    model = TTSModel.load_model(language=language)
    tokenizer = model.flow_lm.conditioner.tokenizer
    print(f"Estraggo lo stato dalla voce: {voice}")
    state = model.get_state_for_audio_prompt(voice)

    rows: list[dict] = []
    print(f"\n{'token':>6} | {'durata(s)':>9} | {'attesi':>6} | "
          f"{'resa':>6} | esito")
    print("-" * 52)
    for target in points:
        text, words, ntok = build_text(tokenizer, target)
        # max_tokens ENORME -> il software non spezza: chunk unico oltre il 50.
        audio = model.generate_audio(state, text, max_tokens=1_000_000)
        wav_path = out_dir / f"{language}_{ntok:03d}tok.wav"
        _save_wav(wav_path, audio, model.sample_rate)
        dur = wav_duration_s(wav_path)

        ret = None
        if asr is not None:
            try:
                ret = retention(words, asr(wav_path))
            except Exception as e:
                print(f"   (ASR fallito su {ntok} token: {e})")
        rows.append({"tokens": ntok, "dur": dur, "words": len(words), "ret": ret})
        ret_str = f"{ret*100:5.0f}%" if ret is not None else "  -  "
        print(f"{ntok:>6} | {dur:>9.1f} | {len(words):>6} | {ret_str:>6} | "
              f"{wav_path.name}")

    limit = analyze_retention(rows) if asr is not None else analyze_duration(rows)
    if limit is not None:
        print(f"\n>> {language}: primo calo intorno ai {limit} token "
              f"(oltre questo punto il modello inizia a tagliare).")
    else:
        print(f"\n>> {language}: nessun calo rilevato fino a {points[-1]} token "
              f"(regge oltre l'intervallo provato).")
    return rows, limit


def _save_wav(path: Path, audio, sample_rate: int) -> None:
    import numpy as np
    a = audio
    if hasattr(a, "detach"):
        a = a.detach()
    if hasattr(a, "cpu"):
        a = a.cpu()
    if hasattr(a, "numpy"):
        a = a.numpy()
    a = np.asarray(a, dtype=np.float32).reshape(-1)
    a = np.clip(a, -1.0, 1.0)
    pcm = (a * 32767.0).astype(np.int16)
    with wave.open(str(path), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(int(sample_rate))
        w.writeframes(pcm.tobytes())


def main():
    p = argparse.ArgumentParser(
        description="Sonda il limite reale di token per chunk di PocketTTS.")
    p.add_argument("--voice", default="alba",
                   help="Voce: un file .wav di riferimento oppure una voce "
                        "predefinita (es. 'alba'). Default: alba.")
    p.add_argument("--languages", nargs="+", default=["italian", "italian_24l"],
                   help="Varianti da provare. Default: italian italian_24l.")
    p.add_argument("--min", type=int, default=20, help="Token minimi (default 20).")
    p.add_argument("--max", type=int, default=130, help="Token massimi (default 130).")
    p.add_argument("--step", type=int, default=10, help="Passo (default 10).")
    p.add_argument("--outdir", default="_prova_token",
                   help="Cartella per i wav prodotti (default: _prova_token).")
    p.add_argument("--no-asr", action="store_true",
                   help="Non usare Whisper anche se disponibile (solo durata).")
    args = p.parse_args()

    # Cartella di output robusta: percorso assoluto + os.makedirs, con ripiego
    # su una cartella temporanea se la creazione fallisce (capita su Windows con
    # percorsi dentro Documenti/OneDrive).
    out_dir = Path(args.outdir)
    if not out_dir.is_absolute():
        out_dir = Path.cwd() / out_dir
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError:
        out_dir = Path(tempfile.mkdtemp(prefix="prova_token_"))
        print(f"(non riesco a creare '{args.outdir}' qui; uso {out_dir})")
    points = list(range(args.min, args.max + 1, args.step))

    voice = prepare_voice(args.voice)

    asr = None if args.no_asr else get_asr()

    summary = {}
    for lang in args.languages:
        try:
            _, limit = probe_variant(lang, voice, points, asr, out_dir)
            summary[lang] = limit
        except Exception as e:
            print(f"\n!! Errore con la variante {lang}: {type(e).__name__}: {e}")
            summary[lang] = "errore"

    print("\n" + "=" * 52)
    print("RIEPILOGO (limite empirico di token per chunk)")
    for lang, lim in summary.items():
        if lim == "errore":
            print(f"  {lang:<14}: errore")
        elif lim is None:
            print(f"  {lang:<14}: nessun calo fino a {points[-1]} token")
        else:
            print(f"  {lang:<14}: cala intorno a {lim} token")
    print("Nota: i wav prodotti sono in", out_dir,
          "- ascoltali per confermare a orecchio.")


if __name__ == "__main__":
    sys.exit(main())
