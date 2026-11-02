#!/usr/bin/env python3
"""Generate windows/app.ico (and a PNG preview) for the frozen Windows build.

Run from anywhere:  python3 tools/make_icon.py
Requires Pillow.  The shape is drawn at 1024 px and downsampled, so the small
sizes in the .ico stay clean.
"""
import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
OUT_ICO = os.path.join(REPO, 'windows', 'app.ico')
OUT_PNG = os.path.join(REPO, 'windows', 'app.png')

S = 1024
TOP = (37, 99, 235)      # #2563EB
BOTTOM = (29, 78, 216)   # #1D4ED8
WHITE = (255, 255, 255, 255)


def rounded_mask(size, radius):
    m = Image.new('L', (size, size), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle((0, 0, size - 1, size - 1), radius=radius, fill=255)
    return m


def main():
    # vertical gradient
    grad = Image.new('RGB', (1, S))
    for y in range(S):
        t = y / (S - 1)
        grad.putpixel((0, y), tuple(round(TOP[i] + (BOTTOM[i] - TOP[i]) * t) for i in range(3)))
    bg = grad.resize((S, S))

    layer = Image.new('RGBA', (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    # arrow: shaft + head
    d.rounded_rectangle((466, 214, 558, 566), radius=18, fill=WHITE)
    d.polygon([(336, 522), (688, 522), (512, 744)], fill=WHITE)

    # baseline "tray" the arrow presses down onto
    d.rounded_rectangle((262, 806, 762, 866), radius=30, fill=(255, 255, 255, 215))

    bg = Image.alpha_composite(bg.convert('RGBA'), layer)
    bg.putalpha(rounded_mask(S, 210))

    icon = bg.resize((256, 256), Image.LANCZOS)
    os.makedirs(os.path.dirname(OUT_ICO), exist_ok=True)
    icon.save(OUT_ICO, sizes=[(16, 16), (24, 24), (32, 32), (48, 48),
                              (64, 64), (128, 128), (256, 256)])
    icon.save(OUT_PNG)
    print('wrote', OUT_ICO, os.path.getsize(OUT_ICO), 'bytes')
    print('wrote', OUT_PNG, os.path.getsize(OUT_PNG), 'bytes')


if __name__ == '__main__':
    main()
