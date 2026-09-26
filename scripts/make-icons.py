"""Build the add-in icons in assets/ from the master artwork.

design/icon-master.png is the full-size icon: a warm orange-to-pink gradient circle
with a white tally mark (made with Higgsfield, GPT Image 2.5, then cut out of its
white background). Sizes of 32 px and up are scaled straight from it. At 16-24 px the
scaled tally blurs into a smudge, so for those sizes the gradient is scaled without
the tally and the tally is redrawn on the pixel grid.

The PNGs are committed, so you only need this to change the icon.
Requires Python 3 with Pillow:  python3 scripts/make-icons.py
"""

from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
MASTER = ROOT / "design" / "icon-master.png"
OUT_DIR = ROOT / "assets"

SIZES = [16, 20, 24, 32, 40, 48, 64, 80, 128, 256]  # 256 is for the website, not the manifest
REDRAWN = {
    # size: (stroke width, gap between strokes, stroke top, stroke height), in pixels
    16: (1, 1, 4, 8),
    20: (2, 1, 5, 10),
    24: (2, 2, 6, 12),
}
SUPERSAMPLE = 8


def scaled(master: Image.Image, size: int) -> Image.Image:
    # Premultiplied alpha, so the transparent corners don't bleed into the edge colour.
    return master.convert("RGBa").resize((size, size), Image.LANCZOS).convert("RGBA")


def gradient_only(master: Image.Image, size: int) -> Image.Image:
    """The master scaled down with the white tally left out of every pixel's average."""
    step = master.width // size
    src = master.load()
    out = Image.new("RGBA", (size, size))
    dst = out.load()
    for ty in range(size):
        for tx in range(size):
            r = g = b = n = coverage = 0
            for y in range(ty * step, (ty + 1) * step, 2):
                for x in range(tx * step, (tx + 1) * step, 2):
                    pr, pg, pb, pa = src[x, y]
                    coverage += pa
                    if pa > 128 and min(pr, pg, pb) < 215:  # inside the circle and not the white tally
                        r, g, b, n = r + pr, g + pg, b + pb, n + 1
            samples = ((step + 1) // 2) ** 2
            alpha = round(coverage / samples)
            if n == 0:  # a block entirely covered by the tally: borrow the neighbour's colour below
                dst[tx, ty] = (0, 0, 0, alpha)
            else:
                dst[tx, ty] = (round(r / n), round(g / n), round(b / n), alpha)
    # Fill any tally-only blocks from their neighbours.
    for ty in range(size):
        for tx in range(size):
            if dst[tx, ty][:3] == (0, 0, 0) and dst[tx, ty][3] > 0:
                near = [dst[x, y] for x, y in ((tx - 1, ty), (tx + 1, ty), (tx, ty - 1), (tx, ty + 1)) if 0 <= x < size and 0 <= y < size and dst[x, y][:3] != (0, 0, 0)]
                if near:
                    dst[tx, ty] = tuple(round(sum(c[i] for c in near) / len(near)) for i in range(3)) + (dst[tx, ty][3],)
    return out


def tally(size: int) -> Image.Image:
    """Four vertical strokes on whole pixels, crossed by a diagonal, drawn in white."""
    width, gap, top, height = REDRAWN[size]
    span = 4 * width + 3 * gap
    left = (size - span) // 2
    s = SUPERSAMPLE
    layer = Image.new("L", (size * s, size * s), 0)
    draw = ImageDraw.Draw(layer)
    for i in range(4):
        x = left + i * (width + gap)
        draw.rectangle([x * s, top * s, (x + width) * s - 1, (top + height) * s - 1], fill=255)
    # Diagonal from lower left to upper right, a little past the outer strokes.
    x0, y0 = (left - 1) * s, (top + height * 0.72) * s
    x1, y1 = (left + span + 1) * s, (top + height * 0.28) * s
    draw.line([x0, y0, x1, y1], fill=255, width=width * s)
    return layer.resize((size, size), Image.BOX)


def redrawn(master: Image.Image, size: int) -> Image.Image:
    icon = gradient_only(master, size)
    circle = icon.getchannel("A")
    icon.paste(Image.new("RGBA", (size, size), (255, 255, 255, 255)), (0, 0), tally(size))
    icon.putalpha(circle)  # keep the circle's own edge: the tally never paints outside it
    return icon


def main() -> None:
    master = Image.open(MASTER).convert("RGBA")
    OUT_DIR.mkdir(exist_ok=True)
    for size in SIZES:
        icon = redrawn(master, size) if size in REDRAWN else scaled(master, size)
        path = OUT_DIR / f"icon-{size}.png"
        icon.save(path, optimize=True)
        print(f"wrote {path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
