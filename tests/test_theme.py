"""Check semantic token contrast, including the increased-contrast overrides."""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def tokens(path):
    return dict(re.findall(r'--([\w-]+):\s*([^;]+);', path.read_text()))


def resolve(values, name):
    value = values[name]
    if value.startswith('var(--'):
        return resolve(values, value[6:-1])
    return value


def luminance(color):
    rgb = [int(color[index:index + 2], 16) / 255 for index in (1, 3, 5)]
    return sum((v / 12.92 if v <= .04045 else ((v + .055) / 1.055) ** 2.4) * w
               for v, w in zip(rgb, [.2126, .7152, .0722]))


class ThemeContrastTests(unittest.TestCase):
    def check_palette(self, values):
        pairs = [
            ('color-foreground', 'color-surface', 4.5),
            ('color-muted', 'color-surface', 4.5),
            ('color-muted', 'color-surface-white', 4.5),
            ('color-on-brand', 'color-action', 4.5),
            ('color-on-brand', 'color-action-hover', 4.5),
            ('color-brand-dark', 'color-surface', 4.5),
            ('color-control-line', 'color-surface', 3),
            ('color-success-ink', 'color-success-surface', 4.5),
            ('color-warning-ink', 'color-warning-surface', 4.5),
            ('color-error', 'color-error-surface', 4.5),
            ('color-info-ink', 'color-info-surface', 4.5),
        ]
        for foreground, background, minimum in pairs:
            with self.subTest(foreground=foreground, background=background):
                low, high = sorted(luminance(resolve(values, name)) for name in (foreground, background))
                self.assertGreaterEqual((high + .05) / (low + .05), minimum)

    def test_default_palette(self):
        self.check_palette(tokens(ROOT / 'assets/tokens.css'))

    def test_increased_contrast_palette(self):
        self.check_palette({**tokens(ROOT / 'assets/tokens.css'), **tokens(ROOT / 'assets/high-contrast.css')})
