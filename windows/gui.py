#!/usr/bin/env python3
"""Windows front-end for pptx-slim.

Three ways to run it:

    PPT瘦身.exe                     open the window
    PPT瘦身.exe 报告.pptx            drag files onto the exe -> list is pre-filled
    PPT瘦身.exe --cli 报告.pptx 5    headless: compress, exit code 0/1 (used by CI)
    PPT瘦身.exe --self-test          headless: verify the frozen build (used by CI)

The windowed build has no console, so main() never relies on stdout, and any
stray print() from the engine goes to a buffer instead of crashing.
"""
import io
import os
import queue
import sys
import threading
import traceback

# PyInstaller --windowed leaves sys.stdout / sys.stderr as None. The engine
# prints its report, so give it somewhere to go before anything else runs.
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

HERE = os.path.dirname(os.path.abspath(__file__))
APP_ID = 'Forever61T.pptx-slim'


# ------------------------------------------------------------------ engine

def load_engine():
    """Import pptx_slim from the repo (windows/..) or the frozen bundle."""
    for p in (HERE, os.path.dirname(HERE)):
        if p and p not in sys.path:
            sys.path.insert(0, p)
    import pptx_slim
    return pptx_slim


def ui_lang():
    forced = (os.environ.get('PPTX_SLIM_LANG') or '').strip().lower()
    if forced.startswith('zh'):
        return 'zh'
    if forced.startswith('en'):
        return 'en'
    if sys.platform == 'win32':
        try:
            import ctypes
            lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
            if (lang_id & 0x3FF) == 0x04:      # LANG_CHINESE
                return 'zh'
        except Exception:
            pass
    return 'en'


L = {
    'zh': {
        'title': 'PPT 瘦身',
        'subtitle': '把 .pptx 压到指定大小 —— 版式、文字、动画一帧都不动',
        'files': '要压缩的文件',
        'add': '选择文件…',
        'remove': '移除选中',
        'clear': '清空',
        'target': '压到',
        'mb': 'MB 以内',
        'target_hint': '提交系统给多少就填多少，常见 5 或 10',
        'start': '开始压缩',
        'open': '打开输出文件夹',
        'log': '压缩过程',
        'ready': '就绪',
        'no_files': '先选一个 .pptx 文件。',
        'only_pptx': '只支持 .pptx。旧版 .ppt 请先在 PowerPoint 里「另存为」成 .pptx。',
        'bad_mb': '目标大小要填一个大于 0 的数字，比如 5。',
        'running': '正在压缩 {i}/{n}…',
        'dropped': '已载入 {n} 个文件 —— 确认下面的大小，然后点「开始压缩」。',
        'done_title': '完成',
        'done_ok': '全部搞定，共 {n} 个文件。\n\n结果就在原文件旁边，文件名带 (compressed)。',
        'done_partial': '{ok} 个完成，{bad} 个没压成。\n\n{first}',
        'failed': '没压成',
        'output': '输出：{path}',
        'already': '本来就小于目标，不用压。',
        'cancelled': '已取消。',
        'engine_missing': '找不到压缩引擎（pptx_slim.py）。安装包可能不完整，请重新下载。',
        'open_none': '还没有生成任何文件。',
        'detail': '详细过程看下面的日志。',
    },
    'en': {
        'title': 'PPT Slim',
        'subtitle': 'Shrink a .pptx under a target size — layout, text and animation untouched',
        'files': 'Files',
        'add': 'Choose files…',
        'remove': 'Remove',
        'clear': 'Clear',
        'target': 'Shrink to',
        'mb': 'MB or less',
        'target_hint': 'Use the limit your submission system gives you — usually 5 or 10',
        'start': 'Compress',
        'open': 'Open output folder',
        'log': 'Report',
        'ready': 'Ready',
        'no_files': 'Pick a .pptx file first.',
        'only_pptx': 'Only .pptx is supported. Save legacy .ppt as .pptx in PowerPoint first.',
        'bad_mb': 'Enter a target greater than 0 MB, e.g. 5.',
        'running': 'Compressing {i}/{n}…',
        'dropped': '{n} file(s) loaded — check the target size, then press Compress.',
        'done_title': 'Done',
        'done_ok': 'All done — {n} file(s).\n\nThe result sits next to the original, named "(compressed)".',
        'done_partial': '{ok} done, {bad} failed.\n\n{first}',
        'failed': 'failed',
        'output': 'Output: {path}',
        'already': 'Already under target — nothing to do.',
        'cancelled': 'Cancelled.',
        'engine_missing': 'Cannot find the compression engine (pptx_slim.py). The package looks incomplete — please download it again.',
        'open_none': 'Nothing has been written yet.',
        'detail': 'See the log below for details.',
    },
}


def ui_font():
    if sys.platform == 'win32':
        return 'Microsoft YaHei UI'
    if sys.platform == 'darwin':
        return 'PingFang SC'
    return 'DejaVu Sans'


def mono_font():
    if sys.platform == 'win32':
        return 'Consolas'
    if sys.platform == 'darwin':
        return 'Menlo'
    return 'DejaVu Sans Mono'


def resource(name):
    """Path to a bundled file (PyInstaller onefile or the repo)."""
    import sys as _sys
    base = getattr(_sys, '_MEIPASS', None)
    for d in (base, HERE):
        if d:
            p = os.path.join(d, name)
            if os.path.isfile(p):
                return p
    return None


class _Writer(io.TextIOBase):
    """A stdout replacement that forwards the engine's report to the GUI."""

    def __init__(self, sink):
        self._sink = sink

    def write(self, s):
        if s:
            self._sink(s)
        return len(s)

    def flush(self):
        pass

    def isatty(self):
        return False


class App:
    def __init__(self, root, files=(), scale=1.0):
        import tkinter as tk
        from tkinter import messagebox

        self.tk = tk
        self.messagebox = messagebox
        self.root = root
        self.lang = ui_lang()
        self.T = lambda key, **kw: L[self.lang][key].format(**kw)

        try:
            self.engine = load_engine()
            self.engine.LANG = self.lang
        except Exception:
            self.messagebox.showerror('pptx-slim', self.T('engine_missing'))
            raise SystemExit(1)

        self.files = []
        self.outputs = []
        self.failures = []
        self.running = False
        self.scale = max(1.0, scale)
        self.q = queue.Queue()

        self._build()
        for f in files:
            self._add(f)
        if self.files:
            self.status.config(text=self.T('dropped', n=len(self.files)))
        self.root.after(80, self._drain)

    # ------------------------------------------------------------ widgets
    def _build(self):
        tk = self.tk
        font = ui_font()
        bg = '#F6F7F9'
        ink = '#111827'
        muted = '#6B7280'

        self.root.title(self.T('title'))
        self.root.configure(bg=bg)
        # Grown by the DPI scale, but never past the screen: a 200% laptop still
        # has to fit the window on screen.
        s = self.scale
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = min(int(780 * s), int(sw * 0.95))
        h = min(int(680 * s), int(sh * 0.88))
        self.root.geometry('%dx%d' % (w, h))
        self.root.minsize(min(int(680 * s), int(sw * 0.80)),
                          min(int(600 * s), int(sh * 0.70)))

        ico = resource('app.ico')
        if ico:
            try:
                self.root.iconbitmap(ico)
            except Exception:
                pass

        # header ---------------------------------------------------------
        head = tk.Frame(self.root, bg=bg)
        head.pack(fill='x', padx=20, pady=(18, 8))
        tk.Label(head, text=self.T('title'), bg=bg, fg=ink,
                 font=(font, 20, 'bold')).pack(anchor='w')
        tk.Label(head, text=self.T('subtitle'), bg=bg, fg=muted,
                 font=(font, 10)).pack(anchor='w', pady=(4, 0))

        # file list ------------------------------------------------------
        box = tk.LabelFrame(self.root, text=self.T('files'), bg=bg, fg=ink,
                            font=(font, 10, 'bold'), bd=1, relief='solid',
                            labelanchor='nw', padx=10, pady=10)
        box.pack(fill='both', expand=False, padx=20, pady=(10, 0))

        inner = tk.Frame(box, bg=bg)
        inner.pack(fill='both', expand=True)
        self.lb = tk.Listbox(inner, height=4, activestyle='none', bd=1,
                             relief='solid', highlightthickness=0,
                             selectmode='extended', font=(font, 10),
                             bg='white', fg=ink, selectbackground='#DBEAFE',
                             selectforeground=ink)
        scroll = tk.Scrollbar(inner, orient='vertical', command=self.lb.yview)
        self.lb.configure(yscrollcommand=scroll.set)
        self.lb.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        row = tk.Frame(box, bg=bg)
        row.pack(fill='x', pady=(8, 0))
        self._btn(row, self.T('add'), self._choose, primary=False).pack(side='left')
        self._btn(row, self.T('remove'), self._remove, primary=False).pack(side='left', padx=(8, 0))
        self._btn(row, self.T('clear'), self._clear, primary=False).pack(side='left', padx=(8, 0))

        # target + action -----------------------------------------------
        ctrl = tk.Frame(self.root, bg=bg)
        ctrl.pack(fill='x', padx=20, pady=(14, 0))

        tk.Label(ctrl, text=self.T('target'), bg=bg, fg=ink,
                 font=(font, 11)).pack(side='left')
        self.mb = tk.StringVar(value='5')
        entry = tk.Entry(ctrl, textvariable=self.mb, width=5, justify='center',
                         font=(font, 12), bd=1, relief='solid',
                         highlightthickness=0)
        entry.pack(side='left', padx=6)
        tk.Label(ctrl, text=self.T('mb'), bg=bg, fg=ink,
                 font=(font, 11)).pack(side='left')
        tk.Label(ctrl, text='  ' + self.T('target_hint'), bg=bg, fg=muted,
                 font=(font, 9)).pack(side='left')

        self.btn_start = self._btn(ctrl, self.T('start'), self._start, primary=True)
        self.btn_start.pack(side='right')
        self.btn_open = self._btn(ctrl, self.T('open'), self._open_folder, primary=False)
        self.btn_open.pack(side='right', padx=(0, 8))
        self.btn_open.configure(state='disabled')

        # log ------------------------------------------------------------
        logbox = tk.LabelFrame(self.root, text=self.T('log'), bg=bg, fg=ink,
                               font=(font, 10, 'bold'), bd=1, relief='solid',
                               labelanchor='nw', padx=8, pady=8)
        logbox.pack(fill='both', expand=True, padx=20, pady=(14, 8))

        self.log = tk.Text(logbox, height=14, wrap='none', state='disabled',
                           font=(mono_font(), 9), bg='#0F172A', fg='#DCE6F5',
                           insertbackground='white', relief='flat',
                           padx=10, pady=8)
        logscroll = tk.Scrollbar(logbox, orient='vertical', command=self.log.yview)
        self.log.configure(yscrollcommand=logscroll.set)
        self.log.pack(side='left', fill='both', expand=True)
        logscroll.pack(side='right', fill='y')

        # status ---------------------------------------------------------
        self.status = tk.Label(self.root, text=self.T('ready'), anchor='w',
                               bg=bg, fg=muted, font=(font, 9))
        self.status.pack(fill='x', padx=20, pady=(0, 12))

    def _btn(self, parent, text, cmd, primary):
        tk = self.tk
        font = ui_font()
        if primary:
            b = tk.Button(parent, text=text, command=cmd, relief='flat', bd=0,
                          bg='#2563EB', fg='white', activebackground='#1D4ED8',
                          activeforeground='white', padx=18, pady=7,
                          font=(font, 10, 'bold'), cursor='hand2')
        else:
            b = tk.Button(parent, text=text, command=cmd, relief='solid', bd=1,
                          bg='white', fg='#111827', activebackground='#F3F4F6',
                          padx=12, pady=6, font=(font, 10), cursor='hand2')
        return b

    # ------------------------------------------------------------ helpers
    def _append(self, text):
        self.log.configure(state='normal')
        self.log.insert('end', text)
        self.log.see('end')
        self.log.configure(state='disabled')

    def _add(self, path):
        path = os.path.abspath(path)
        if path.lower().endswith('.pptx') and os.path.isfile(path) \
                and path not in self.files:
            self.files.append(path)
        self._refresh_list()

    def _refresh_list(self):
        self.lb.delete(0, 'end')
        for p in self.files:
            self.lb.insert('end', '  ' + p)

    def _choose(self):
        from tkinter import filedialog
        picked = filedialog.askopenfilenames(
            title=self.T('add'), filetypes=[('PowerPoint', '*.pptx')])
        for p in picked:
            self._add(p)
        if self.files:
            self.status.config(text=self.T('ready'))

    def _remove(self):
        for i in reversed(self.lb.curselection()):
            del self.files[i]
        self._refresh_list()

    def _clear(self):
        self.files = []
        self._refresh_list()

    def _open_folder(self):
        folder = None
        for out in reversed(self.outputs):
            if out and os.path.isfile(out):
                folder = os.path.dirname(out)
                break
        if not folder:
            self.messagebox.showinfo(self.T('title'), self.T('open_none'))
            return
        try:
            if sys.platform == 'win32':
                os.startfile(folder)
            elif sys.platform == 'darwin':
                import subprocess
                subprocess.Popen(['open', folder])
            else:
                import subprocess
                subprocess.Popen(['xdg-open', folder])
        except Exception:
            self.messagebox.showinfo(self.T('title'), folder)

    # ------------------------------------------------------------- running
    def _start(self):
        if self.running:
            return
        if not self.files:
            self.messagebox.showwarning(self.T('title'), self.T('no_files'))
            return
        bad = [p for p in self.files if not p.lower().endswith('.pptx')]
        if bad:
            self.messagebox.showwarning(self.T('title'), self.T('only_pptx'))
            return
        try:
            mb = float(str(self.mb.get()).replace(',', '.').strip())
            if not (0 < mb < 100000):
                raise ValueError
        except ValueError:
            self.messagebox.showwarning(self.T('title'), self.T('bad_mb'))
            return

        self.running = True
        self.outputs = []
        self.failures = []
        self.btn_start.configure(state='disabled')
        self.btn_open.configure(state='disabled')
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')
        self._append(self.T('title') + '  v' + getattr(self.engine, '__version__', '') + '\n')
        self.status.config(text=self.T('running', i=0, n=len(self.files)))

        files = list(self.files)
        threading.Thread(target=self._run, args=(files, mb), daemon=True).start()

    def _run(self, files, mb):
        total = len(files)
        for i, path in enumerate(files, 1):
            self.q.put(('progress', (i, total)))
            self.q.put(('log', '\n=== [%d/%d] %s ===\n' % (i, total, os.path.basename(path))))
            old = sys.stdout
            sys.stdout = _Writer(lambda s: self.q.put(('log', s)))
            try:
                out, before, after, _problems = self.engine.slim(
                    path, int(mb * 1024 * 1024), verbose=True)
                self.outputs.append(out)
                self.q.put(('log', '\n' + self.T('output', path=out) + '\n'))
            except SystemExit as e:
                msg = str(e.code) if e.code not in (None, 0) else self.T('cancelled')
                self.q.put(('log', '\n' + msg + '\n'))
                self.failures.append((path, msg))
            except Exception:
                tb = traceback.format_exc()
                self.q.put(('log', '\n' + tb + '\n'))
                self.failures.append((path, tb.strip().splitlines()[-1]))
            finally:
                sys.stdout = old
        self.q.put(('all_done', None))

    def _drain(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == 'log':
                    self._append(payload)
                elif kind == 'progress':
                    i, n = payload
                    self.status.config(text=self.T('running', i=i, n=n))
                elif kind == 'all_done':
                    self._finish()
        except queue.Empty:
            pass
        self.root.after(80, self._drain)

    def _finish(self):
        self.running = False
        self.btn_start.configure(state='normal')
        if self.outputs:
            self.btn_open.configure(state='normal')
            self.status.config(text=self.T('ready'))

        if self.failures:
            first = '\n'.join(m.replace('\n', ' ')[:160] for _p, m in self.failures[:3])
            self.messagebox.showwarning(
                self.T('title'),
                self.T('done_partial', ok=len(self.outputs), bad=len(self.failures),
                     first=first))
        else:
            self.messagebox.showinfo(
                self.T('title'), self.T('done_ok', n=len(self.outputs)))


# ---------------------------------------------------------------- headless

def _log_path(name='pptx-slim-cli.log'):
    import tempfile
    return os.path.join(tempfile.gettempdir(), name)


def _write_log(text, name='pptx-slim-cli.log'):
    try:
        with open(_log_path(name), 'w', encoding='utf-8') as f:
            f.write(text + '\n')
    except OSError:
        pass


def _cli(argv):
    """argv: [src, [target_mb]] (already past --cli) -> exit code."""
    if not argv:
        _write_log('no input path')
        return 2
    src = argv[0]
    try:
        mb = float(argv[1]) if len(argv) > 1 else 5.0
    except ValueError:
        _write_log('bad target: %r' % argv[1:2])
        return 2
    try:
        engine = load_engine()
        out, before, after, problems = engine.slim(
            src, int(mb * 1024 * 1024), verbose=False)
        _write_log('ok %s %d %d %d %s' % (out, before, after, len(problems), problems))
        return 0
    except SystemExit as e:
        _write_log('refused: %s' % (e.code,))
        return 1
    except Exception:
        _write_log(traceback.format_exc().strip())
        return 2


def _self_test():
    """Verify the frozen build can actually do its job on this machine."""
    lines = []
    code = 0
    try:
        import PIL
        from PIL import Image
        buf = io.BytesIO()
        Image.new('RGB', (48, 32), (37, 99, 235)).save(buf, 'JPEG')
        img = Image.open(io.BytesIO(buf.getvalue()))
        lines.append('pillow %s ok (%s)' % (PIL.__version__, img.size))

        engine = load_engine()
        lines.append('engine %s ok' % engine.__version__)

        import tkinter
        tkinter.Tcl()
        lines.append('tk ok')

        # a real zip round-trip, the way the engine does it on a pptx
        import zipfile
        zpath = _log_path('pptx-slim-selftest.zip')
        with zipfile.ZipFile(zpath, 'w') as z:
            z.writestr('[Content_Types].xml', '<Types/>')
        with zipfile.ZipFile(zpath) as z:
            assert z.read('[Content_Types].xml') == b'<Types/>'
        os.remove(zpath)
        lines.append('zip ok')
    except Exception:
        lines.append(traceback.format_exc().strip())
        code = 1
    _write_log('\n'.join(lines), 'pptx-slim-selftest.log')
    return code


def main():
    argv = sys.argv[1:]

    if sys.platform == 'win32':
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
        except Exception:
            pass
        try:
            import ctypes
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    if argv and argv[0] == '--self-test':
        return _self_test()
    if argv and argv[0] == '--cli':
        return _cli(argv[1:])

    import tkinter as tk
    root = tk.Tk()
    scale = 1.0
    try:
        dpi = float(root.winfo_fpixels('1i') or 96.0)
        root.tk.call('tk', 'scaling', dpi / 72.0)
        scale = dpi / 96.0
    except Exception:
        pass
    files = [a for a in argv if a.lower().endswith('.pptx')]
    App(root, files, scale=scale)
    root.mainloop()
    return 0


if __name__ == '__main__':
    sys.exit(main())
