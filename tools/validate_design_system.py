#!/usr/bin/env python3
"""Valida documentazione, token e contrasto di SlideNarrator Adaptive."""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOKEN_FILE = ROOT / "design" / "tokens" / "slidenarrator.tokens.json"
REQUIRED_DOCS = [
    "docs/DESIGN_SYSTEM.md",
    "docs/UI_COMPONENTS.md",
    "docs/RESPONSIVE_LAYOUT.md",
    "docs/ACCESSIBILITY.md",
    "docs/ICONOGRAPHY.md",
    "docs/IMPLEMENTATION_CHECKLIST.md",
    "docs/REFERENCES.md",
]
HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$")


def _linear(value: int) -> float:
    channel = value / 255.0
    return channel / 12.92 if channel <= 0.04045 else ((channel + 0.055) / 1.055) ** 2.4


def luminance(color: str) -> float:
    value = color.lstrip("#")[:6]
    red, green, blue = (int(value[index:index + 2], 16) for index in (0, 2, 4))
    return 0.2126 * _linear(red) + 0.7152 * _linear(green) + 0.0722 * _linear(blue)


def contrast_ratio(foreground: str, background: str) -> float:
    high, low = sorted((luminance(foreground), luminance(background)), reverse=True)
    return (high + 0.05) / (low + 0.05)


def fail(message: str, errors: list[str]) -> None:
    errors.append(message)
    print(f"[ERRORE] {message}")


def main() -> int:
    errors: list[str] = []

    if not TOKEN_FILE.is_file():
        fail(f"File token mancante: {TOKEN_FILE}", errors)
        return 1

    try:
        data = json.loads(TOKEN_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"Token JSON non valido: {exc}", errors)
        return 1

    for relative in REQUIRED_DOCS:
        if not (ROOT / relative).is_file():
            fail(f"Documento mancante: {relative}", errors)

    if data.get("meta", {}).get("name") != "SlideNarrator Adaptive":
        fail("Nome design system non valido", errors)

    dimensions = data.get("dimensions", {})
    if dimensions.get("minimumTargetSize", 0) < 36:
        fail("minimumTargetSize deve essere almeno 36", errors)
    if dimensions.get("controlHeight", 0) < 40:
        fail("controlHeight deve essere almeno 40", errors)
    if dimensions.get("focusRingWidth", 0) < 2:
        fail("focusRingWidth deve essere almeno 2", errors)
    if dimensions.get("minimumWindowWidth", 0) < 800:
        fail("minimumWindowWidth deve essere almeno 800", errors)

    breakpoints = data.get("breakpoints", {})
    ordered = [breakpoints.get(key, -1) for key in ("narrow", "compact", "standard", "wide")]
    if ordered != sorted(ordered) or len(set(ordered)) != len(ordered):
        fail("Breakpoint non strettamente crescenti", errors)

    themes = data.get("themes", {})
    for theme_name in ("light", "dark"):
        colors = themes.get(theme_name, {}).get("colors", {})
        if not colors:
            fail(f"Tema mancante: {theme_name}", errors)
            continue
        for token, value in colors.items():
            if not isinstance(value, str) or not HEX_RE.fullmatch(value):
                fail(f"Colore non valido: {theme_name}.{token}={value!r}", errors)

    print("\nVerifica contrasto:")
    for check in data.get("contrastChecks", []):
        theme = check["theme"]
        foreground_name = check["foreground"]
        background_name = check["background"]
        minimum = float(check["minimum"])
        colors = themes[theme]["colors"]
        ratio = contrast_ratio(colors[foreground_name], colors[background_name])
        print(f"  {theme:5} {foreground_name:14} / {background_name:14}: {ratio:.2f}:1 (min {minimum:.1f})")
        if ratio + 1e-9 < minimum:
            fail(f"Contrasto insufficiente: {theme}.{foreground_name}/{background_name} = {ratio:.2f}", errors)

    if errors:
        print(f"\nValidazione fallita: {len(errors)} errore/i.")
        return 1

    print("\nSlideNarrator Adaptive valido.")
    print(f"Documenti verificati: {len(REQUIRED_DOCS)}")
    print(f"Controlli contrasto: {len(data.get('contrastChecks', []))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
