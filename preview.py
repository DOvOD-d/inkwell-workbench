"""后台 Word/WPS 预览渲染：生成临时 Word → 导出 PDF → 取页图。

Word/WPS 只在后台线程里通过 COM 操作，界面线程只负责把 PDF 页转成 QPixmap。
"""
from __future__ import annotations

import io
import queue
import shutil
import subprocess
import tempfile
from pathlib import Path

import pypdfium2 as pdfium
from PySide6.QtCore import QThread, Signal

import engine

try:
    import pythoncom
    import win32com.client as win32
except Exception:  # pragma: no cover - reported at runtime
    pythoncom = None
    win32 = None


class PreviewWorker(QThread):
    rendered = Signal(int, bytes, int)      # token, pdf bytes, page count
    failed = Signal(int, str, bool)         # token, message, office unavailable
    status = Signal(str)

    def __init__(self, work_dir, parent=None):
        super().__init__(parent)
        self.work_dir = Path(work_dir)
        self._queue = queue.Queue(maxsize=1)
        self._stopping = False
        self._office = None
        self._office_kind = ''
        self._keeper = None
        self._no_office = ''

    def submit(self, token, template, config, values):
        if self._stopping:
            return
        if not self.isRunning():
            self.start()
        try:
            while True:
                self._queue.get_nowait()
        except queue.Empty:
            pass
        try:
            self._queue.put_nowait((token, Path(template), config, values))
        except queue.Full:
            pass

    def stop(self):
        self._stopping = True
        try:
            while True:
                self._queue.get_nowait()
        except queue.Empty:
            pass
        try:
            self._queue.put_nowait(None)
        except queue.Full:
            pass

    def _start_office(self):
        if self._office is not None:
            try:
                self._office.Version
                return self._office
            except Exception:
                self._drop_office()
        if pythoncom is None or win32 is None:
            self._no_office = '运行环境缺少 pywin32，无法调用 Word/WPS 实时预览'
            raise RuntimeError(self._no_office)
        self.status.emit('正在启动 Word 渲染（首次约几秒）…')
        last = None
        for progid in ('Word.Application', 'KWPS.Application'):
            try:
                app = win32.DispatchEx(progid)
                try:
                    app.Visible = False
                except Exception:
                    pass
                try:
                    app.DisplayAlerts = 0
                except Exception:
                    pass
                try:
                    app.ScreenUpdating = False
                except Exception:
                    pass
                try:
                    app.AutomationSecurity = 3
                except Exception:
                    pass
                self._office = app
                self._office_kind = progid
                try:
                    self._keeper = app.Documents.Add()
                except Exception:
                    self._keeper = None
                return app
            except Exception as exc:
                last = exc
        raise RuntimeError(f'未能启动 Word/WPS：{last}')

    def _drop_office(self):
        app, self._office = self._office, None
        self._keeper = None
        if app is not None:
            try:
                app.Quit(0)
            except Exception:
                try:
                    app.Quit()
                except Exception:
                    pass

    def _export_pdf(self, docx_path, pdf_path):
        last = None
        for attempt in range(2):
            try:
                app = self._start_office()
            except Exception as exc:
                last = exc
                if attempt == 0:
                    self._drop_office()
                    continue
                break
            self.status.emit('正在排版预览…')
            doc = None
            try:
                try:
                    doc = app.Documents.Open(str(docx_path), ReadOnly=True,
                                             AddToRecentFiles=False, Visible=False)
                except Exception:
                    doc = app.Documents.Open(str(docx_path))
                try:
                    doc.ExportAsFixedFormat(str(pdf_path), 17)
                except Exception:
                    doc.SaveAs(str(pdf_path), 17)
                if not pdf_path.is_file() or not pdf_path.stat().st_size:
                    raise RuntimeError('Word 没有生成 PDF')
                return pdf_path.read_bytes()
            except Exception as exc:
                last = exc
                if doc is not None:
                    try:
                        doc.Close(0)
                    except Exception:
                        pass
                    doc = None
                if attempt == 0:
                    self._drop_office()
                    continue
                break
            finally:
                if doc is not None:
                    try:
                        doc.Close(0)
                    except Exception:
                        pass
        fallback = self._export_pdf_soffice(docx_path, pdf_path)
        if fallback:
            return fallback
        raise RuntimeError(f'Word 导出 PDF 失败：{last}')

    @staticmethod
    def _soffice_path():
        for name in ('soffice', 'soffice.exe'):
            found = shutil.which(name)
            if found:
                return found
        for base in (r'C:\Program Files\LibreOffice\program\soffice.exe',
                     r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'):
            if Path(base).is_file():
                return base
        return ''

    def _export_pdf_soffice(self, docx_path, pdf_path):
        exe = self._soffice_path()
        if not exe:
            return None
        self.status.emit('Word/WPS 不可用，正在用 LibreOffice 渲染…')
        try:
            subprocess.run(
                [exe, '--headless', '--convert-to', 'pdf',
                 '--outdir', str(pdf_path.parent), str(docx_path)],
                check=False, timeout=180, stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL)
        except Exception:
            return None
        produced = pdf_path.parent / (docx_path.stem + '.pdf')
        if not produced.is_file() or not produced.stat().st_size:
            return None
        if produced != pdf_path:
            shutil.copy2(produced, pdf_path)
        return pdf_path.read_bytes()

    def _render_job(self, token, template, config, values):
        self.work_dir.mkdir(parents=True, exist_ok=True)
        work = Path(tempfile.mkdtemp(prefix='wtt-preview-', dir=self.work_dir))
        docx_path = work / 'preview.docx'
        pdf_path = work / 'preview.pdf'
        try:
            engine.render_word(template, docx_path, config, values, validate=False,
                               keep_empty_bookmarks=True)
            pdf_bytes = self._export_pdf(docx_path, pdf_path)
            count = len(pdfium.PdfDocument(pdf_bytes))
            if count <= 0:
                raise RuntimeError('预览 PDF 没有页面')
            self.rendered.emit(token, pdf_bytes, count)
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def run(self):
        com_ready = False
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
                com_ready = True
            except Exception:
                com_ready = False
        try:
            while True:
                job = self._queue.get()
                if job is None:
                    break
                token, template, config, values = job
                if self._no_office:
                    self.failed.emit(token, self._no_office, True)
                    continue
                try:
                    self._render_job(token, template, config, values)
                except Exception as exc:
                    msg = str(exc)
                    if 'pywin32' in msg or '运行环境缺少' in msg:
                        self._no_office = msg
                        self.failed.emit(token, msg, True)
                    else:
                        self.failed.emit(token, msg, False)
        finally:
            self._keeper = None
            self._drop_office()
            if com_ready and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


def page_count(pdf_bytes):
    """PDF 页数；解析失败时返回 0，让界面如实提示而不是崩。"""
    if not pdf_bytes:
        return 0
    try:
        doc = pdfium.PdfDocument(pdf_bytes)
    except Exception:
        return 0
    try:
        return len(doc)
    finally:
        doc.close()


def render_page(pdf_bytes, page_index, scale=2.0):
    """把 PDF 的某一页渲染成 PNG 字节，供界面线程加载。"""
    doc = pdfium.PdfDocument(pdf_bytes)
    try:
        if len(doc) == 0:
            raise ValueError('PDF 没有页面')
        index = max(0, min(int(page_index), len(doc) - 1))
        bitmap = doc[index].render(scale=float(scale))
        image = bitmap.to_pil()
        buffer = io.BytesIO()
        image.save(buffer, 'PNG')
        return buffer.getvalue()
    finally:
        doc.close()


def export_pdf_sync(docx_path, pdf_path, status=None):
    """Synchronous DOCX→PDF export for background recognition.

    Reuses the same Word/WPS/LibreOffice fallback as the live preview, but
    initializes COM on the calling worker thread.
    """
    worker = PreviewWorker(Path(pdf_path).parent)
    com_ready = False
    try:
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
                com_ready = True
            except Exception:
                com_ready = False
        if status is not None:
            worker.status.connect(status)
        return worker._export_pdf(Path(docx_path), Path(pdf_path))
    finally:
        worker._drop_office()
        if com_ready and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
