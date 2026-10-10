#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pack a macOS .app (or any folder) into a zip that extracts correctly everywhere.

Why not `ditto -c -k`? Because it writes UTF-8 filenames *without* setting the
UTF-8 flag bit in the zip entry. macOS's own Archive Utility assumes UTF-8 and
copes, but 7-Zip, WinRAR, Windows Explorer and most in-app unzippers honour the
missing flag and fall back to CP437 — so `PPT瘦身.app` arrives as
`PPTτÿªΦ║½.app`. Python's zipfile sets the flag correctly.

What this keeps from ditto: symlinks stay symlinks, and every permission bit
survives via the Unix external attributes. Both matter — the code signature
covers `Contents/MacOS/droplet` and its executable bit.

Usage:
    python3 tools/zip_app.py "PPT瘦身.app" out.zip
    python3 tools/zip_app.py "dist-folder" out.zip
"""
import os
import stat
import sys
import time
import zipfile

UNIX = 3          # ZipInfo.create_system: 3 == Unix, so external_attr is honoured


def _zi(arc, st, is_dir=False):
    zi = zipfile.ZipInfo(arc + '/' if is_dir else arc)
    zi.create_system = UNIX
    zi.external_attr = ((stat.S_IFDIR if is_dir else stat.S_IFREG) | (st.st_mode & 0o7777)) << 16
    zi.compress_type = zipfile.ZIP_DEFLATED
    # Without this every entry is stamped 1980-01-01, which looks broken.
    zi.date_time = time.localtime(st.st_mtime)[:6]
    return zi


def add(z, path, arc):
    st = os.lstat(path)
    if stat.S_ISLNK(st.st_mode):
        zi = zipfile.ZipInfo(arc)
        zi.create_system = UNIX
        zi.external_attr = (stat.S_IFLNK | 0o777) << 16
        zi.compress_type = zipfile.ZIP_STORED
        zi.date_time = time.localtime(st.st_mtime)[:6]
        z.writestr(zi, os.readlink(path))
    elif stat.S_ISDIR(st.st_mode):
        z.writestr(_zi(arc, st, is_dir=True), b'')
        for name in sorted(os.listdir(path)):
            add(z, os.path.join(path, name), f'{arc}/{name}')
    else:
        with open(path, 'rb') as f:
            z.writestr(_zi(arc, st), f.read())


def pack(src, dst):
    src = src.rstrip('/')
    if not os.path.exists(src):
        raise SystemExit(f'not found: {src}')
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        add(z, src, os.path.basename(src))
    return dst


def main():
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    dst = pack(sys.argv[1], sys.argv[2])
    with zipfile.ZipFile(dst) as z:
        infos = z.infolist()
        flagged = sum(1 for i in infos if i.flag_bits & 0x800)
        links = sum(1 for i in infos if stat.S_ISLNK(i.external_attr >> 16))
    print(f'{dst}: {len(infos)} entries, {flagged} UTF-8 flagged, {links} symlink(s)')


if __name__ == '__main__':
    main()
