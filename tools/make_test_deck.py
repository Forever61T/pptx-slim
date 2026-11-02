#!/usr/bin/env python3
"""Build a deliberately oversized .pptx so CI can prove the frozen exe shrinks it.

The deck is one slide with a big noisy photo and a text slide, which is what
makes a real conference deck heavy: one 4.5-inch-wide photo that was saved at
12 megapixels.

    python tools/make_test_deck.py out.pptx [target_mb]

Needs python-pptx and Pillow -- build-time only, never needed by users.
"""

import io
import os
import sys

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt


def noisy_jpeg(width, height, quality=92):
    """Random noise: JPEG cannot compress it away, so the file stays heavy."""
    raw = os.urandom(width * height * 3)
    img = Image.frombytes('RGB', (width, height), raw)
    buf = io.BytesIO()
    img.save(buf, 'JPEG', quality=quality)
    return buf.getvalue()


def main(argv=None):
    argv = sys.argv if argv is None else argv
    out = argv[1] if len(argv) > 1 else 'deck.pptx'
    target_mb = float(argv[2]) if len(argv) > 2 else 8.0

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # One phone-photo-sized image, displayed 4.5 in wide -- the 12 MP trap.
    slide = prs.slides.add_slide(blank)
    picture = noisy_jpeg(4032, 3024)
    slide.shapes.add_picture(io.BytesIO(picture), Inches(1.5), Inches(1.6),
                             width=Inches(4.5))

    legend = slide.shapes.add_textbox(Inches(6.6), Inches(2.0), Inches(6.0), Inches(2.5))
    frame = legend.text_frame
    frame.word_wrap = True
    frame.text = 'pptx-slim test deck'
    frame.paragraphs[0].font.size = Pt(32)
    frame.add_paragraph().text = 'A 4032x3024 photo shown 4.5 in wide.'
    frame.add_paragraph().text = 'Target: %.1f MB' % target_mb

    second = prs.slides.add_slide(blank)
    box = second.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.5), Inches(5.5))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = 'Slide two'
    tf.paragraphs[0].font.size = Pt(40)
    tf.add_paragraph().text = 'Layout, text and animation must survive compression.'

    prs.save(out)
    size = os.path.getsize(out)
    print('%s  %d bytes (%.2f MB)' % (out, size, size / 1048576.0))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
