# pptx-slim

**Shrink a `.pptx` under a size limit — without touching the layout, the text, or a single frame of animation.**

Conference portals, grant systems and university submission forms all cap uploads at 5 MB, 10 MB, 20 MB. A deck full of phone photos blows straight past that. Most "PPT compressors" either re-encode everything blindly and leave you with a mushy mess, or quietly flatten your animated GIFs into a still frame.

`pptx-slim` does neither. It works out how many inches each image *actually occupies on the slide*, derives how many pixels that really needs, and spends its byte budget accordingly.

> A 4032×3024 phone photo displayed 4.5 inches wide does not need 12 megapixels. At 200 DPI it needs 900×675 — about 3% of the pixels. That is where your 20 MB went.

[中文说明 →](README.zh-CN.md)

---

## Download

**macOS — no Python, no terminal.** Grab the drag-and-drop app from
[**Releases**](https://github.com/Forever61T/pptx-slim/releases/latest):
drop a `.pptx` on the icon, type the target size in MB, done.

> **First launch only:** macOS blocks unsigned apps. Right-click the app → **Open** → **Open** again.
> Or once from a terminal: `xattr -cr /Applications/PPT瘦身.app`.
> Every free app without a paid Apple certificate hits this — it is not a fault in the app.

**Windows.** Install [Python](https://www.python.org/downloads/) and tick *Add python.exe to PATH*,
then drag your `.pptx` onto `windows\run.bat`. It installs Pillow for you if it is missing.

**Any platform, command line:**

```bash
pip install git+https://github.com/Forever61T/pptx-slim.git
```

Or just download `pptx_slim.py` and run it. The only dependency is Pillow:

```bash
python3 -m pip install --user Pillow
```

## Use

```bash
pptx-slim deck.pptx          # squeeze under 5 MB (the default)
pptx-slim deck.pptx 20       # squeeze under 20 MB
pptx-slim deck.pptx 5 --inplace        # overwrite the original
pptx-slim deck.pptx 5 --out small.pptx # choose the output path
pptx-slim deck.pptx 5 --lang zh        # force Chinese output
```

Without `--out` or `--inplace` the result lands next to the original as `deck (compressed).pptx`. **Your original is never touched.**

```
Original          22.01 MB
  non-image parts 36.0 KB
  keep as-is      2.50 MB
  shrinkable      19.47 MB  (8 images)
Target 5.00 MB -> budget for those 8 images: 2.46 MB

  220 dpi / max quality 90  ->  8.31 MB
  180 dpi / max quality 86  ->  5.72 MB
  150 dpi / max quality 80  ->  4.07 MB   OK

Image                   Before     After  Ratio   Q      New size
----------------------------------------------------------------------
image1.jpeg           2.90 MB    305 KB    11%  88     1200x900
image2.jpeg           4.38 MB    313 KB     7%  88     1200x900
image4.png            3.40 MB    116 KB     3%  90     1500x828
image9.png            3.14 MB    269 KB     9%  88     1400x730
image6.png            1.86 MB    189 KB    10%  90     1200x1079
image3.png            1.05 MB    116 KB    11%  90     1300x975
image5.png           1025 KB    167 KB    16%  84     1400x800
image8.png            857 KB    133 KB    16%  90     1041x693

Kept as-is: image1.gif (1.40 MB, animated, frames preserved),
            image2.gif (1.10 MB, animated, frames preserved)

Done: 22.01 MB -> 4.07 MB  (18%, saved 17.94 MB). Structure check passed.
```

## What it refuses to touch, and why

This is the part that matters. Shrinking these would make the file smaller *and wrong*:

| Kept as-is | Why |
|---|---|
| **Animated GIFs** | Re-encoding with Pillow reliably makes them *bigger* (measured: 162% and 167% of the original) and any real reduction costs frames. Your animation survives intact. |
| **Transparent images** | Flattening to JPEG turns every transparent pixel into a white box — logos on coloured slides get a visible rectangle behind them. Only *real* transparency counts: an alpha channel that is uniformly opaque is treated as opaque and compressed. |
| **Tiny images (< 50 KB)** | Nothing meaningful to win, everything to lose. |

If the target is unreachable once those are set aside, the tool says so and hands you the smallest version it produced, rather than silently degrading something it should not.

## The safety net

Renaming `image4.png` to `image4.jpg` breaks every relationship file that pointed at it, and a part extension that `[Content_Types].xml` does not declare makes PowerPoint open with *"PowerPoint found a problem with content… repair?"* Most quick scripts get this wrong.

Every output is checked before it is handed to you:

- all XML parts are well-formed
- every internal relationship resolves to a part that exists
- every part's extension is declared in `[Content_Types].xml` — and when the source declares only `jpg`, the tool uses `jpg` rather than inventing `jpeg`
- renamed media is rewritten across every `.rels` and slide reference

If any check fails, the run **aborts instead of writing a broken file**.

## Platforms

The engine is pure Python — standard library plus Pillow, with no macOS-specific
calls in it. It runs anywhere Python 3.8+ and Pillow do: macOS, Windows, Linux.

The drag-and-drop wrappers are per-platform, because dragging a file onto an icon
is the one thing that is not portable:

| Platform | Wrapper | Notes |
|---|---|---|
| macOS | `macos/build.sh` | Builds a `.app` droplet with the engine embedded |
| Windows | `windows/run.bat` | Drag a `.pptx` onto the `.bat`, or double-click and paste a path |
| Linux | — | Use the command line; `python3 pptx_slim.py deck.pptx 5` |

### macOS

```bash
macos/build.sh                    # builds "PPT Slimmer.app"
APP_NAME="PPT瘦身" macos/build.sh  # ...or a localised name
```

Drag a `.pptx` onto the icon, type the target size in MB, and it reports the
result. The engine is embedded in the app bundle, so the app works on its own.
It is signed ad-hoc, so the first launch needs Control-click → Open (or
`xattr -cr /Applications/PPT瘦身.app`).

### Windows

Drag a `.pptx` onto `windows\run.bat`. It finds your Python (`python` or the
`py` launcher), installs Pillow for you if it is missing, and asks for a target
size. Requires Python from [python.org](https://www.python.org/downloads/) with
**Add python.exe to PATH** ticked during install.

### Packaging a release

`ditto -c -k` writes UTF-8 filenames without setting the zip UTF-8 flag, so a
bundle named `PPT瘦身.app` arrives as `PPTτÿªΦ║½.app` in 7-Zip, WinRAR and
Windows Explorer. Use the bundled packer instead — it sets the flag and still
preserves symlinks and permission bits:

```bash
python3 tools/zip_app.py "PPT瘦身.app" release.zip
```

(macOS's own `/usr/bin/unzip` is Info-ZIP 6.00 and shows mojibake either way —
that is a limitation of that ancient tool, not of the archive.)

## Options

| Flag | Meaning |
|---|---|
| `target_mb` | Target size in MB (default `5`) |
| `--out PATH` | Write the result here |
| `--inplace` | Overwrite the original file |
| `--lang en\|zh` | Force output language (otherwise from `LANG`/`LC_ALL`) |
| `--quiet` | Suppress the report |
| `--version` | Print version |

## How it works

1. **Measure.** Parse every slide, layout, master and notes page. For each picture, read the `a:ext` transform (in EMU, 914400 per inch) to get the size it is drawn at — not its pixel size.
2. **Budget.** Subtracting everything that must stay (non-image parts, GIFs, transparent art), divide the remaining byte budget among the shrinkable images in proportion to their *on-slide area*. A full-bleed background gets more bytes than a 1-inch thumbnail.
3. **Compress.** For each image: downscale to the pixels its display size actually needs at the current DPI setting, then binary-search JPEG quality downward until it fits its share.
4. **Escalate.** If the whole deck is still over target, drop the DPI and try again — a lower resolution at high quality beats a high resolution at quality 40.
5. **Verify and pack.** Rewrite relationships, re-check the package, rebuild the zip with `[Content_Types].xml` first.

## Limitations

- **`.pptx` only.** The legacy binary `.ppt` format is not supported.
- **Embedded video and audio are left alone.** They are often the real bulk in a modern deck; re-encoding video is a different problem with different trade-offs, and silently degrading someone's recording would be worse than being over budget.
- **Fonts are not subset.** Embedded fonts are kept whole.
- **Raster images only.** EMF/WMF vector graphics are not converted (doing so usually makes them worse).
- Requires Python 3.7+ and Pillow.

## Contributing

Issues and pull requests welcome — especially real-world decks that come out wrong. If a file trips the structure check, that is a bug worth reporting, and the checker's message tells you exactly which part is at fault.

## License

[MIT](LICENSE)
