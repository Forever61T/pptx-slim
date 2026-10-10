#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptx-slim — shrink a .pptx under a target size without touching layout, text or animation.

The core idea: work out how many inches each image actually occupies on the slide,
then derive how many pixels it really needs. Everything beyond that is waste —
a 4032x3024 phone photo shown 4.5 inches wide does not need 12 megapixels.

Usage:
    python3 pptx_slim.py deck.pptx [TARGET_MB] [--out PATH] [--inplace] [--lang en|zh]

Default target is 5 MB. The result is written next to the original as
"<name> (compressed).pptx" unless you pass --out or --inplace.
"""
import argparse
import io
import os
import re
import shutil
import sys
import tempfile
import zipfile

__version__ = "1.0.1"

try:
    from PIL import Image
except ImportError:                                        # pragma: no cover
    sys.exit('Pillow is required.  Install it with:  python3 -m pip install --user Pillow')

import xml.etree.ElementTree as ET

LANCZOS = getattr(getattr(Image, 'Resampling', Image), 'LANCZOS')

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XA, XP, XR = f'{{{A}}}', f'{{{P}}}', f'{{{R}}}'
EMU_IN = 914400.0

# Pixels to keep per inch of on-slide size. 200 is already far beyond what a
# projector or a printer can resolve.
ROUNDS = [(220, 90), (180, 86), (150, 80), (125, 74), (100, 66), (80, 58)]
MIN_QUALITY = 45
SMALL_ENOUGH = 50_000          # below this, an image is not worth touching


# ---------------------------------------------------------------- i18n

def detect_lang():
    """Pick a language from the environment.

    LC_ALL outranks LANG, but "C" and "POSIX" carry no language information —
    plenty of shells export LC_ALL=C.UTF-8, and treating that as a vote for
    English would ignore a user who deliberately set LANG=zh_CN.UTF-8.
    """
    for var in ('PPTX_SLIM_LANG', 'LC_ALL', 'LC_MESSAGES', 'LANG'):
        v = (os.environ.get(var) or '').strip().lower()
        if not v or v == 'c' or v == 'posix' or v.startswith(('c.', 'posix.')):
            continue
        return 'zh' if v.startswith(('zh', 'chinese')) else 'en'
    return 'en'


LANG = detect_lang()

M = {
    'no_media': (
        'No images in this pptx — nothing to shrink. The size probably comes from '
        'embedded video or fonts.',
        '这个 pptx 里没有图片，压不动。体积大概来自内嵌视频或字体。'),
    'no_shrinkable': (
        'Nothing shrinkable: every image is either animated or has real transparency.',
        '没有可压缩的图片——剩下的都是动图或带真透明通道的素材。'),
    'already_small': ('Already {size} — under target, nothing to do.',
                      '已经只有 {size}，没超过目标，不用压。'),
    'not_found': ('File not found: {path}', '找不到文件：{path}'),
    'not_pptx': ('Not a .pptx file (the legacy .ppt format is not supported): {path}',
                 '这不是 .pptx 文件（旧版 .ppt 格式不支持）：{path}'),
    'bad_target': ('Target must be greater than 0 MB (got {got}).',
                   '目标必须大于 0 MB（你填的是 {got}）。'),
    'bad_zip': ('Cannot read this as a .pptx — the file may be incomplete, '
                'password-protected, or not a PowerPoint file at all:\n  {path}',
                '读不出 pptx 结构 —— 文件可能没下载完整、被加密、'
                '或者根本不是 PowerPoint 文件：\n  {path}'),
    'no_write': ('No permission to write here:\n  {path}\n'
                 'Choose another location with --out.',
                 '没有写入权限：\n  {path}\n可以用 --out 换一个能写的位置。'),
    'hdr_original': ('Original          {size}', '原始文件          {size}'),
    'hdr_overhead': ('  non-image parts {size}', '  非图片部分      {size}'),
    'hdr_keep': ('  keep as-is      {size}', '  必须保留的图片  {size}'),
    'hdr_shrink': ('  shrinkable      {size}  ({n} images)', '  可压缩的图片    {size}  共 {n} 张'),
    'hdr_budget': ('Target {target} -> budget for those {n} images: {budget}',
                   '目标 {target} → 这 {n} 张图的预算 {budget}'),
    'round': ('  {dpi:>3} dpi / max quality {q:>2}  ->  {size}{ok}',
              '  档位 {dpi:>3} dpi / 质量上限 {q:>2}  →  {size}{ok}'),
    'round_ok': ('   OK', '   ✓ 达标'),
    'tbl_image': ('Image', '图片'),
    'tbl_before': ('Before', '原大小'),
    'tbl_after': ('After', '压缩后'),
    'tbl_ratio': ('Ratio', '比例'),
    'tbl_quality': ('Q', '质量'),
    'tbl_size': ('New size', '新尺寸'),
    'kept': ('Kept as-is: {items}', '必须保留：{items}'),
    'skipped': ('skip', '跳过'),
    'reason_gif': ('animated, frames preserved', '动图，保留动画'),
    'reason_tiny': ('small enough', '太小，不值得动'),
    'reason_alpha': ('has transparency', '带透明通道，保留'),
    'reason_unreadable': ('unreadable, left alone', '无法识别，保留'),
    'verify_fail': ('\nStructure check FAILED:', '\n⚠ 结构体检发现问题：'),
    'verify_abort': ('Aborting rather than hand you a file PowerPoint will want to repair. '
                     'Please report this.',
                     '为避免交出一个打不开的文件，已中止。请把这个问题反馈一下。'),
    'done_ok': ('Done: {before} -> {after}  ({pct}%, saved {saved}). Structure check passed.',
                '完成：{before} → {after}（{pct}%，省掉 {saved}），结构体检通过。'),
    'done_partial': ('Best effort: {before} -> {after}, still above the {target} target.',
                     '已尽力：{before} → {after}，仍未压到 {target} 以内。'),
    'done_partial_why': ('What is left is almost certainly animated GIFs, transparent art or '
                         'embedded video. Shrinking those costs frames or reveals white boxes, '
                         'so they were left alone.',
                         '剩下的体积多半来自动图、透明素材或内嵌视频。动它们会掉帧或露白底，所以没碰。'),
    'output': ('Output: {path}', '输出：{path}'),
    'desc': ('Shrink a .pptx under a target size without touching layout or animation.',
             'PPTX 瘦身：压到指定大小，版式和动画都不动。'),
}

# Fill in after --lang is parsed
_TRANSLATIONS = {}


def t(key, **kw):
    pair = M[key]
    s = pair[1] if LANG == 'zh' else pair[0]
    return s.format(**kw) if kw else s


def hsize(n):
    if n < 1024:
        return f'{n} B'
    if n < 1024 * 1024:
        return f'{n / 1024:.1f} KB'
    return f'{n / 1024 / 1024:.2f} MB'


# ---------------------------------------------------------------- parsing

def part_rels(z, names, part):
    """{rId: target} for one part."""
    d, b = os.path.split(part)
    rels = f'{d}/_rels/{b}.rels'
    out = {}
    if rels in names:
        for rel in ET.fromstring(z.read(rels)):
            out[rel.get('Id')] = rel.get('Target')
    return out


def usages(z, names, part):
    """[(media basename, width_in, height_in), ...] for one part."""
    try:
        root = ET.fromstring(z.read(part))
    except Exception:
        return []
    rels = part_rels(z, names, part)
    found = []
    for tag in (XP + 'pic', XP + 'sp'):
        for el in root.iter(tag):
            blip = el.find(f'.//{XA}blip')
            ext = el.find(f'.//{XP}spPr/{XA}xfrm/{XA}ext')
            if blip is None or ext is None:
                continue
            tgt = rels.get(blip.get(XR + 'embed'))
            if not tgt:
                continue
            try:
                found.append((os.path.basename(tgt),
                              int(ext.get('cx')) / EMU_IN,
                              int(ext.get('cy')) / EMU_IN))
            except (TypeError, ValueError):
                pass
    return found


def slide_size_in(z, names):
    """Slide size in inches; falls back to 13.333x7.5 (16:9)."""
    p = 'ppt/presentation.xml'
    if p in names:
        try:
            sz = ET.fromstring(z.read(p)).find(f'{XP}sldSz')
            if sz is not None:
                return int(sz.get('cx')) / EMU_IN, int(sz.get('cy')) / EMU_IN
        except Exception:
            pass
    return 13.333, 7.5


def survey(z, names):
    """Largest on-slide size each media file appears at."""
    sw, sh = slide_size_in(z, names)
    best = {}
    for part in names:
        if not (part.startswith(('ppt/slides/slide', 'ppt/slideLayouts/slideLayout',
                                 'ppt/slideMasters/slideMaster', 'ppt/notesSlides/notesSlide'))
                and part.endswith('.xml')):
            continue
        for media, w, h in usages(z, names, part):
            cur = best.get(media)
            if cur is None or w * h > cur[0] * cur[1]:
                best[media] = (w, h)
    # Media we never saw placed: assume full-bleed, the conservative choice.
    for part in names:
        if part.startswith('ppt/media/'):
            best.setdefault(os.path.basename(part), (sw, sh))
    return best


# --------------------------------------------- staying valid when renaming

def pick_jpeg_ext(z):
    """Pick a JPEG extension [Content_Types].xml already declares.

    OPC requires every part extension to have a Default or Override entry.
    Sources often declare only 'jpg'; renaming a png to .jpeg then makes
    PowerPoint offer to repair the file.
    """
    try:
        root = ET.fromstring(z.read('[Content_Types].xml'))
    except Exception:
        return '.jpeg', True
    declared = {e.get('Extension', '').lower() for e in root if e.tag.endswith('}Default')}
    for ext in ('jpeg', 'jpg', 'jpe'):
        if ext in declared:
            return '.' + ext, False
    return '.jpeg', True


def add_content_type_default(path, ext):
    """Add a Default entry by string insertion, to avoid ET mangling namespaces."""
    try:
        text = open(path, 'r', encoding='utf-8').read()
    except OSError:
        return False
    if re.search(rf'Extension="{re.escape(ext)}"', text, re.I):
        return False
    new = re.sub(r'(<Types\b[^>]*>)',
                 rf'\1<Default Extension="{ext}" ContentType="image/jpeg"/>',
                 text, count=1)
    if new == text:
        return False
    open(path, 'w', encoding='utf-8').write(new)
    return True


# ---------------------------------------------------------------- encoding

def encode_jpeg(im, quality):
    buf = io.BytesIO()
    im.save(buf, 'JPEG', quality=quality, optimize=True, subsampling=2)
    return buf.getvalue()


def has_alpha(im):
    if im.mode in ('RGBA', 'LA'):
        return True
    return im.mode == 'P' and 'transparency' in im.info


def alpha_is_real(im):
    """True when the alpha channel actually carries information."""
    if not has_alpha(im):
        return False
    if im.mode == 'P':
        im = im.convert('RGBA')
    lo, _ = im.convert('RGBA').getchannel('A').getextrema()
    return lo < 255


def fit(im, cap_px):
    w, h = im.size
    if max(w, h) <= cap_px:
        return im
    s = cap_px / max(w, h)
    return im.resize((max(1, round(w * s)), max(1, round(h * s))),
                     LANCZOS, reducing_gap=3.0)


def compress_one(path, cap_px, budget, qmax):
    """Encode one image into `budget` bytes. Returns (bytes, quality, (w, h)) or None."""
    im = Image.open(path)
    if alpha_is_real(im):
        return None, None, None
    rgb = fit(im.convert('RGB'), cap_px)
    data, q = None, None
    for q in range(qmax, MIN_QUALITY - 1, -4):
        data = encode_jpeg(rgb, q)
        if len(data) <= budget:
            break
    return data, q, rgb.size


# ---------------------------------------------------------------- packaging

def repack(workdir, out_path):
    ct = os.path.join(workdir, '[Content_Types].xml')
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(ct, '[Content_Types].xml')          # must come first
        for root, dirs, files in os.walk(workdir):
            dirs.sort()
            for f in sorted(files):
                full = os.path.join(root, f)
                rel = os.path.relpath(full, workdir).replace(os.sep, '/')
                if rel == '[Content_Types].xml':
                    continue
                z.write(full, rel)


def rewrite_refs(workdir, renames):
    """Keep every .rels / .xml reference in sync after a rename."""
    if not renames:
        return 0
    touched = 0
    for root, dirs, files in os.walk(workdir):
        for f in files:
            if not f.endswith(('.rels', '.xml')):
                continue
            p = os.path.join(root, f)
            try:
                text = open(p, 'r', encoding='utf-8').read()
            except (UnicodeDecodeError, OSError):
                continue
            new = text
            for old, rep in renames.items():
                new = new.replace(f'media/{old}', f'media/{rep}')
            if new != text:
                open(p, 'w', encoding='utf-8').write(new)
                touched += 1
    return touched


def verify(path):
    """XML well-formedness, relationship integrity, Content_Types coverage."""
    problems = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        for n in names:
            if n.endswith(('.xml', '.rels')):
                try:
                    ET.fromstring(z.read(n))
                except Exception as e:
                    problems.append(f'broken XML {n}: {e}')
        for n in names:
            if not n.endswith('.rels'):
                continue
            base = os.path.dirname(os.path.dirname(n))
            try:
                root = ET.fromstring(z.read(n))
            except Exception:
                continue
            for rel in root:
                if rel.get('TargetMode') == 'External':
                    continue
                tgt = os.path.normpath(
                    os.path.join(base, rel.get('Target', ''))).replace('\\', '/')
                if tgt not in names:
                    problems.append(f'dangling reference {n} -> {rel.get("Target")}')
        try:
            ct = ET.fromstring(z.read('[Content_Types].xml'))
            declared = {e.get('Extension', '').lower() for e in ct if e.tag.endswith('}Default')}
            overrides = {e.get('PartName', '') for e in ct if e.tag.endswith('}Override')}
            missing = set()
            for n in names:
                ext = os.path.splitext(n)[1][1:].lower()
                if ext and ext not in declared and ('/' + n) not in overrides:
                    missing.add(ext)
            for ext in sorted(missing):
                problems.append(f'[Content_Types].xml does not declare .{ext}')
        except Exception as e:
            problems.append(f'cannot parse [Content_Types].xml: {e}')
    return problems


# ---------------------------------------------------------------- main flow

def slim(src, target_bytes, out_path=None, verbose=True, inplace=False):
    """Shrink `src` under `target_bytes`. Returns (path, before, after, problems)."""
    src = os.path.abspath(os.path.expanduser(src))
    if not os.path.isfile(src):
        raise SystemExit(t('not_found', path=src))
    if not src.lower().endswith('.pptx'):
        raise SystemExit(t('not_pptx', path=src))
    if target_bytes <= 0:
        # 0 or a negative target used to run the whole pipeline and then
        # announce "still above 0 B" — reject it instead of pretending.
        raise SystemExit(t('bad_target', got=f'{target_bytes / (1024 * 1024):g}'))

    orig_size = os.path.getsize(src)
    if orig_size <= target_bytes:
        if verbose:
            print(t('already_small', size=hsize(orig_size)))
        # Nothing to do, but never hand back a path that does not exist.
        return src, orig_size, orig_size, []

    if out_path is None:
        base, ext = os.path.splitext(src)
        out_path = base + ' (compressed)' + ext

    try:
        z = zipfile.ZipFile(src)
    except (zipfile.BadZipFile, OSError):
        # A truncated download, a password-protected file and a renamed .docx
        # all land here. Letting BadZipFile escape gives the user a traceback.
        raise SystemExit(t('bad_zip', path=src))
    names = set(z.namelist())
    media = sorted(n for n in names if n.startswith('ppt/media/') and not n.endswith('/'))
    if not media:
        z.close()
        raise SystemExit(t('no_media'))

    display = survey(z, names)
    jpeg_ext, need_ct_patch = pick_jpeg_ext(z)

    # How many zip bytes go to everything that is not media.
    media_zip = sum(zi.compress_size for zi in z.infolist()
                    if zi.filename.startswith('ppt/media/'))
    overhead = orig_size - media_zip

    keep, shrink = {}, []
    for m in media:
        name = os.path.basename(m)
        size = z.getinfo(m).file_size
        if name.lower().endswith('.gif'):
            keep[name] = (size, t('reason_gif'))
            continue
        if size < SMALL_ENOUGH:
            keep[name] = (size, t('reason_tiny'))
            continue
        try:
            probe = Image.open(io.BytesIO(z.read(m)))
        except Exception:
            keep[name] = (size, t('reason_unreadable'))
            continue
        if alpha_is_real(probe):
            keep[name] = (size, t('reason_alpha'))
            continue
        w, h = display.get(name, (13.333, 7.5))
        shrink.append({'part': m, 'name': name, 'size': size, 'win': w, 'hin': h})

    if not shrink:
        z.close()
        raise SystemExit(t('no_shrinkable'))

    keep_bytes = sum(v[0] for v in keep.values())
    avail = target_bytes - overhead - keep_bytes
    if verbose:
        print(t('hdr_original', size=hsize(orig_size)))
        print(t('hdr_overhead', size=hsize(overhead)))
        print(t('hdr_keep', size=hsize(keep_bytes)))
        print(t('hdr_shrink', size=hsize(sum(s['size'] for s in shrink)), n=len(shrink)))
        print(t('hdr_budget', target=hsize(target_bytes), n=len(shrink),
                budget=hsize(max(avail, 0))))
        print()

    # Split the budget by on-slide area: bigger on screen, more bytes.
    areas = [max(s['win'] * s['hin'], 0.05) for s in shrink]
    total_area = sum(areas)
    for s, a in zip(shrink, areas):
        s['budget'] = max(40_000, int(max(avail, 0) * a / total_area))

    tmp = tempfile.mkdtemp(prefix='pptxslim_')
    try:
        best = None
        for dpi, qmax in ROUNDS:
            work = os.path.join(tmp, f'd{dpi}')
            with zipfile.ZipFile(src) as zz:
                zz.extractall(work)
            if need_ct_patch:
                add_content_type_default(os.path.join(work, '[Content_Types].xml'),
                                         jpeg_ext.lstrip('.'))

            renames, results = {}, []
            for s in shrink:
                path = os.path.join(work, s['part'].replace('/', os.sep))
                cap = max(240, int(max(s['win'], s['hin']) * dpi))
                data, q, dims = compress_one(path, cap, s['budget'], qmax)
                if data is None:
                    results.append((s['name'], s['size'], s['size'], None, t('skipped')))
                    continue
                newname = re.sub(r'\.(png|bmp|tif|tiff)$', jpeg_ext, s['name'])
                newpath = os.path.join(os.path.dirname(path), newname)
                with open(newpath, 'wb') as f:
                    f.write(data)
                if newpath != path:
                    os.remove(path)
                    renames[s['name']] = newname
                results.append((s['name'], s['size'], len(data), q,
                                f'{dims[0]}x{dims[1]}'))

            rewrite_refs(work, renames)
            cand = os.path.join(tmp, f'out_{dpi}.pptx')
            repack(work, cand)
            size = os.path.getsize(cand)
            if verbose:
                print(t('round', dpi=dpi, q=qmax, size=hsize(size),
                        ok=t('round_ok') if size <= target_bytes else ''))
            if best is None or size < best[1]:
                best = (cand, size, results, dpi, qmax)
            if size <= target_bytes:
                break

        cand, size, results, dpi, qmax = best
        if verbose:
            print()
            print(f"{t('tbl_image'):<20}{t('tbl_before'):>11}{t('tbl_after'):>11}"
                  f"{t('tbl_ratio'):>7}{t('tbl_quality'):>4}  {t('tbl_size'):>12}")
            print('-' * 74)
            for name, o, n, q, dim in results:
                print(f'{name:<20}{hsize(o):>11}{hsize(n):>11}{n / o * 100:>6.0f}%'
                      f'{(str(q) if q else "-"):>4}  {dim:>12}')
            if keep:
                print()
                print(t('kept', items=', '.join(f'{k} ({hsize(v[0])}, {v[1]})'
                                                for k, v in sorted(keep.items()))))

        problems = verify(cand)
        if problems:
            print(t('verify_fail'))
            for p in problems[:10]:
                print('   ', p)
            raise SystemExit(t('verify_abort'))

        final_out = src if inplace else out_path
        if os.path.abspath(final_out) != os.path.abspath(cand):
            try:
                shutil.copyfile(cand, final_out)
            except OSError:
                # Read-only folder, a locked file, a full disk — anything the
                # OS refuses. Say which path and stop, do not traceback.
                raise SystemExit(t('no_write', path=final_out))
        if verbose:
            print()
            if size <= target_bytes:
                print(t('done_ok', before=hsize(orig_size), after=hsize(size),
                        pct=f'{size / orig_size * 100:.0f}',
                        saved=hsize(orig_size - size)))
            else:
                print(t('done_partial', before=hsize(orig_size), after=hsize(size),
                        target=hsize(target_bytes)))
                print('  ' + t('done_partial_why'))
            print(t('output', path=final_out))
        return final_out, orig_size, size, problems
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        try:
            z.close()
        except Exception:
            pass


def main(argv=None):
    global LANG
    ap = argparse.ArgumentParser(
        prog='pptx-slim', description=M['desc'][1 if LANG == 'zh' else 0])
    ap.add_argument('input', help='the .pptx to shrink')
    ap.add_argument('target_mb', nargs='?', type=float, default=5.0,
                    help='target size in MB (default: 5)')
    ap.add_argument('--out', help='write the result here')
    ap.add_argument('--inplace', action='store_true', help='overwrite the original')
    ap.add_argument('--lang', choices=['en', 'zh'], help='force output language')
    ap.add_argument('--quiet', action='store_true')
    ap.add_argument('--version', action='version', version=f'pptx-slim {__version__}')
    a = ap.parse_args(argv)

    if a.lang:
        LANG = a.lang

    slim(a.input, int(a.target_mb * 1024 * 1024), a.out,
         verbose=not a.quiet, inplace=a.inplace)


if __name__ == '__main__':
    main()
