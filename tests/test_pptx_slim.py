"""End-to-end regression tests for pptx-slim.

The suite runs the real command line on ``_test/sample.pptx`` and checks the
result the way a recipient would: the file opens in python-pptx, it is no
bigger than the requested target, and no image carries more pixels than its
on-slide size can show at 200 DPI.

Run from the repository root with::

    python -m pytest -v

Real presentations are gitignored (``*.pptx``), so CI starts without a sample.
When ``_test/sample.pptx`` is missing the session fixture builds one there.
The generated deck is deliberately image-heavy: at the 3 MB target used here
the engine cannot stop at its 220 DPI first round, so the 200 DPI assertion
is exercised against an actually downscaled deck rather than a lucky pass.
"""

import hashlib
import io
import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

REPO_ROOT = Path(__file__).resolve().parents[1]
SAMPLE = REPO_ROOT / "_test" / "sample.pptx"

# A realistic target for the ~20 MB sample: it forces the engine past its
# 220 DPI first round, which is what makes the 200 DPI check meaningful.
TARGET_MB = 3.0
SIZE_TOLERANCE = 0.02          # accepted overshoot over the requested target
MAX_DPI = 200
EMU_PER_INCH = 914400

# Enough noisy photos that the 220 DPI round alone exceeds the 3 MB target.
GENERATED_SLIDES = 24


def _noisy_jpeg(width=1200, height=900, quality=92):
    """Random noise: JPEG cannot compress it away, so the deck stays heavy."""
    raw = os.urandom(width * height * 3)
    image = Image.frombytes("RGB", (width, height), raw)
    buf = io.BytesIO()
    image.save(buf, "JPEG", quality=quality)
    return buf.getvalue()


def _generate_sample(path):
    """Build a representative deck at *path* for CI, where no sample exists."""
    from pptx import Presentation as NewPresentation
    from pptx.util import Inches

    deck = NewPresentation()
    deck.slide_width = Inches(13.333)
    deck.slide_height = Inches(7.5)
    blank = deck.slide_layouts[6]

    for index in range(GENERATED_SLIDES):
        slide = deck.slides.add_slide(blank)
        slide.shapes.add_picture(
            io.BytesIO(_noisy_jpeg()), Inches(1.5), Inches(1.6), width=Inches(4.5))
        text = slide.shapes.add_textbox(Inches(6.6), Inches(2.0), Inches(6.0), Inches(1.5))
        text.text_frame.word_wrap = True
        text.text_frame.text = "pptx-slim regression sample %d" % index

    path.parent.mkdir(parents=True, exist_ok=True)
    deck.save(str(path))


def _pictures(shapes):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _pictures(shape.shapes)
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            yield shape


@pytest.fixture(scope="session")
def sample_deck():
    """The deck under test, generated at the expected path when absent."""
    if not SAMPLE.is_file():
        _generate_sample(SAMPLE)
    return SAMPLE


@pytest.fixture(scope="session")
def compressed_output(sample_deck, tmp_path_factory):
    """Compress the sample once through the public CLI; all tests share it."""
    out_dir = tmp_path_factory.mktemp("compressed")
    out_path = out_dir / "sample (compressed).pptx"
    env = dict(os.environ, PPTX_SLIM_LANG="en")
    completed = subprocess.run(
        [sys.executable, str(REPO_ROOT / "pptx_slim.py"), str(sample_deck),
         "%g" % TARGET_MB, "--out", str(out_path)],
        cwd=str(REPO_ROOT), env=env, capture_output=True, text=True, timeout=600)
    assert completed.returncode == 0, (
        "pptx_slim.py failed with exit code %d\n--- stdout ---\n%s\n--- stderr ---\n%s"
        % (completed.returncode, completed.stdout, completed.stderr))
    assert out_path.is_file(), "pptx_slim.py exited 0 but wrote no output"
    return out_path


def test_output_opens_in_python_pptx(compressed_output, sample_deck):
    """The compressed deck is a valid presentation, not just a smaller zip."""
    with sample_deck.open("rb") as handle:
        source = Presentation(handle)
    with compressed_output.open("rb") as handle:
        result = Presentation(handle)

    assert len(result.slides) == len(source.slides)
    assert len(result.slides) > 0
    for slide in result.slides:
        assert len(slide.shapes) > 0


def test_output_is_within_target(compressed_output):
    """The output honours the requested target, with 2% headroom for encoders."""
    limit = int(TARGET_MB * 1024 * 1024 * (1 + SIZE_TOLERANCE))
    size = compressed_output.stat().st_size
    assert size > 0
    assert size <= limit, "output is %d bytes, target allows %d" % (size, limit)


def test_images_fit_their_display_size_at_200_dpi(compressed_output):
    """Every embedded image has at most display_size_inches x 200 pixels wide."""
    with compressed_output.open("rb") as handle:
        deck = Presentation(handle)
    slide_width_in = deck.slide_width / EMU_PER_INCH

    # The engine guarantees the cap at the largest place an image appears,
    # so measure the same way: the widest occurrence of each image part.
    display_width = {}
    parts = list(deck.slides)
    for master in deck.slide_masters:
        parts.append(master)
        parts.extend(master.slide_layouts)
    for part in parts:
        for picture in _pictures(part.shapes):
            width_in = picture.width / EMU_PER_INCH
            key = picture.image.sha1
            display_width[key] = max(display_width.get(key, 0.0), width_in)

    with zipfile.ZipFile(compressed_output) as package:
        media = [name for name in package.namelist()
                 if name.startswith("ppt/media/") and not name.endswith("/")]
        assert media, "compressed deck has no images to check"
        for name in media:
            blob = package.read(name)
            pixel_width = Image.open(io.BytesIO(blob)).size[0]
            width_in = display_width.get(hashlib.sha1(blob).hexdigest(), slide_width_in)
            allowed = round(width_in * MAX_DPI)
            assert pixel_width <= allowed, (
                "%s is %d px wide but %.2f in at %d DPI allows %d px"
                % (name, pixel_width, width_in, MAX_DPI, allowed))
