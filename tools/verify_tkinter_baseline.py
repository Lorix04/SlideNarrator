#!/usr/bin/env python3
"""Verifica l'integrità della baseline Tkinter di SlideNarrator."""
from __future__ import annotations

import hashlib
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SUMS_FILE = ROOT / "baseline" / "TKINTER_BASELINE_SHA256SUMS.txt"
REQUIRED = (
    "slide_narrator.py",
    "slide_narrator_gui.py",
    "slide_narrator_gui_legacy.py",
    "slide_narrator_batch.py",
    "requirements.txt",
    "README.md",
    "assets/slide_narrator_icon.png",
    "assets/slide_narrator_icon.ico",
    "docs/TKINTER_BASELINE.md",
    "docs/FUNCTIONAL_INVENTORY_TKINTER.md",
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    errors: list[str] = []

    for relative in REQUIRED:
        if not (ROOT / relative).is_file():
            errors.append(f"File obbligatorio mancante: {relative}")

    if not SUMS_FILE.is_file():
        errors.append(f"Manifest hash mancante: {SUMS_FILE.relative_to(ROOT)}")
    else:
        for line_number, raw in enumerate(SUMS_FILE.read_text(encoding="utf-8").splitlines(), 1):
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            try:
                expected, relative = line.split("  ", 1)
            except ValueError:
                errors.append(f"Riga hash non valida {line_number}: {raw}")
                continue
            path = ROOT / relative
            if not path.is_file():
                errors.append(f"File del manifest mancante: {relative}")
                continue
            actual = sha256(path)
            if actual.lower() != expected.lower():
                errors.append(
                    f"Hash differente: {relative}\n"
                    f"  previsto: {expected}\n"
                    f"  attuale:  {actual}"
                )

    if errors:
        print("Baseline Tkinter NON valida:\n")
        for error in errors:
            print(f"- {error}")
        return 1

    print("Baseline Tkinter valida.")
    print(f"File verificati: {sum(1 for l in SUMS_FILE.read_text(encoding='utf-8').splitlines() if l and not l.startswith('#'))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
