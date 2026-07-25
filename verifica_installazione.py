from __future__ import annotations
import importlib
import contextlib
import io
import json
import os
import shutil
import sys
from pathlib import Path

BASE = ["edge_tts", "ttkbootstrap", "pptx", "openpyxl", "lxml", "mutagen", "fitz"]
OPTIONAL = ["pygame", "pocket_tts", "chatterbox", "TTS", "sounddevice"]
result = {
    "python": sys.version,
    "base": {},
    "optional": {},
    "tools": {"ffmpeg": shutil.which("ffmpeg"), "ffprobe": shutil.which("ffprobe"),
              "libreoffice": shutil.which("soffice") or shutil.which("libreoffice")},
    "ok": True,
}
for name in BASE:
    try:
        mod = importlib.import_module(name)
        result["base"][name] = getattr(mod, "__version__", "ok")
    except Exception as exc:
        result["base"][name] = f"ERRORE: {exc}"
        result["ok"] = False
for name in OPTIONAL:
    try:
        # I backend opzionali possono scrivere avvisi su stderr (per esempio
        # Transformers quando PyTorch non e' installato). Li catturiamo per
        # evitare che l'installer PowerShell li scambi per errori fatali.
        optional_stderr = io.StringIO()
        with contextlib.redirect_stderr(optional_stderr):
            mod = importlib.import_module(name)
        result["optional"][name] = getattr(mod, "__version__", "installato")
    except Exception as exc:
        result["optional"][name] = f"non disponibile: {exc}"
try:
    import slide_narrator
    result["project_import"] = "ok"
except Exception as exc:
    result["project_import"] = f"ERRORE: {exc}"
    result["ok"] = False
out = Path(os.environ.get("SLIDENARRATOR_VERIFY_OUTPUT", "")).expanduser() if os.environ.get("SLIDENARRATOR_VERIFY_OUTPUT") else Path(__file__).with_name("verifica_installazione.json")
out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(result, ensure_ascii=False, indent=2))
raise SystemExit(0 if result["ok"] else 1)
