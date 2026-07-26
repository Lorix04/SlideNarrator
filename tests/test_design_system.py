from __future__ import annotations

import json
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOKENS = json.loads((ROOT / "design" / "tokens" / "slidenarrator.tokens.json").read_text(encoding="utf-8"))


def linear(value: int) -> float:
    channel = value / 255.0
    return channel / 12.92 if channel <= 0.04045 else ((channel + 0.055) / 1.055) ** 2.4


def luminance(color: str) -> float:
    color = color.lstrip("#")[:6]
    red, green, blue = (int(color[index:index + 2], 16) for index in (0, 2, 4))
    return 0.2126 * linear(red) + 0.7152 * linear(green) + 0.0722 * linear(blue)


def contrast(a: str, b: str) -> float:
    high, low = sorted((luminance(a), luminance(b)), reverse=True)
    return (high + 0.05) / (low + 0.05)


class DesignSystemTests(unittest.TestCase):
    def test_name_and_version(self) -> None:
        self.assertEqual(TOKENS["meta"]["name"], "SlideNarrator Adaptive")
        self.assertRegex(TOKENS["meta"]["version"], r"^\d+\.\d+\.\d+$")

    def test_required_themes(self) -> None:
        self.assertIn("light", TOKENS["themes"])
        self.assertIn("dark", TOKENS["themes"])

    def test_target_and_controls(self) -> None:
        self.assertGreaterEqual(TOKENS["dimensions"]["minimumTargetSize"], 36)
        self.assertGreaterEqual(TOKENS["dimensions"]["controlHeight"], 40)
        self.assertGreaterEqual(TOKENS["dimensions"]["navigationItemHeight"], 44)

    def test_breakpoints_are_ordered(self) -> None:
        values = [TOKENS["breakpoints"][key] for key in ("narrow", "compact", "standard", "wide")]
        self.assertEqual(values, sorted(values))
        self.assertEqual(len(values), len(set(values)))

    def test_contrast_pairs(self) -> None:
        for check in TOKENS["contrastChecks"]:
            colors = TOKENS["themes"][check["theme"]]["colors"]
            ratio = contrast(colors[check["foreground"]], colors[check["background"]])
            self.assertGreaterEqual(ratio, check["minimum"], msg=str(check))

    def test_official_assets_exist(self) -> None:
        for name in (
            "slide_narrator_icon.png",
            "slide_narrator_icon_sidebar.png",
            "slide_narrator_icon.ico",
            "slide_narrator_logo.png",
        ):
            self.assertTrue((ROOT / "assets" / name).is_file(), name)


if __name__ == "__main__":
    unittest.main()
