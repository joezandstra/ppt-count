"""Regenerate the add-in icons in assets/.

The PNGs are committed, so you only need this if you want to change the icon.
Requires Python 3 with Pillow:  python3 scripts/make-icons.py
"""

from pathlib import Path

from PIL import Image, ImageDraw

SIZES = [16, 20, 24, 32, 40, 48, 64, 80, 128]
SUPERSAMPLE = 8  # draw large, then downscale for smooth edges

BACKGROUND = (37, 99, 235, 255)  # blue-600
PILL = (255, 255, 255, 255)
ACCENT = (253, 186, 116, 255)  # orange-300: the "counted" word

# Rows of "words" as (start, end) fractions of the usable width.
ROWS = [
    [(0.00, 0.38), (0.50, 1.00)],
    [(0.00, 0.58), (0.70, 1.00)],
    [(0.00, 0.30), (0.42, 0.74)],
]
ACCENT_WORD = (1, 0)  # row, word index drawn in the accent colour

OUT_DIR = Path(__file__).resolve().parent.parent / "assets"


def draw_icon(size: int) -> Image.Image:
    big = size * SUPERSAMPLE
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    d.rounded_rectangle([0, 0, big - 1, big - 1], radius=big * 0.22, fill=BACKGROUND)

    # Small icons get fewer, chunkier rows so they stay legible.
    rows = ROWS[:2] if size <= 20 else ROWS
    margin = big * (0.20 if size <= 20 else 0.19)
    usable = big - 2 * margin
    pill_h = usable * (0.26 if len(rows) == 2 else 0.17)
    gap = (usable - pill_h * len(rows)) / (len(rows) - 1)

    for r, words in enumerate(rows):
        top = margin + r * (pill_h + gap)
        for w, (a, b) in enumerate(words):
            left = margin + a * usable
            right = margin + b * usable
            colour = ACCENT if (r, w) == ACCENT_WORD else PILL
            d.rounded_rectangle([left, top, right, top + pill_h], radius=pill_h / 2, fill=colour)

    return img.resize((size, size), Image.LANCZOS)


def main() -> None:
    OUT_DIR.mkdir(exist_ok=True)
    for size in SIZES:
        path = OUT_DIR / f"icon-{size}.png"
        draw_icon(size).save(path, optimize=True)
        print(f"wrote {path.relative_to(OUT_DIR.parent)}")


if __name__ == "__main__":
    main()
