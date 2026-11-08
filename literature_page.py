# -*- coding: utf-8 -*-
"""文献检索页：e2sci 主入口、平台跳转、OpenAlex/Crossref 题录与 GB/T 复制。

放在独立模块里，避免把联网逻辑塞进原来的模板识别/排版代码。
布局按 UI 设计稿：左栏固定宽度放检索词、e2sci 主入口和平台跳转，右栏题录表撑满可视区。
"""
from __future__ import annotations

import re

from PySide6.QtCore import Qt, QThread, Signal, QUrl, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QFrame, QGridLayout, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QPushButton, QScrollArea, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

import feedback
import literature

LEFT_WIDTH = 430
MIN_BODY_WIDTH = 940      # 再窄就出横向滚动条，而不是把两栏挤成一列一个字


def label(text, role=None):
    w = QLabel(text)
    # 标题类标签保持单行：被挤窄时不能竖成一列一个字
    w.setWordWrap(role not in ('title', 'section', 'tag'))
    if role:
        w.setObjectName(role)
    return w


def button(text, callback=None, primary=False):
    w = QPushButton(text)
    if primary:
        w.setObjectName('primary')
    if callback:
        w.clicked.connect(callback)
    return w


def set_state(widget, **properties):
    """改 QSS 属性选择器用的属性后必须重新 polish，否则 [state="ok"] 不会即时生效。"""
    changed = False
    for key, value in properties.items():
        if widget.property(key) != value:
            widget.setProperty(key, value)
            changed = True
    if changed:
        widget.style().unpolish(widget)
        widget.style().polish(widget)


class PlatformTile(QFrame):
    """平台入口：按钮下面标一行跳转方式，区分 URL 预填和需要手动粘贴。"""

    def __init__(self, name, mode, callback):
        super().__init__()
        self.setObjectName('tile')
        self.setMinimumHeight(44)
        box = QVBoxLayout(self)
        box.setContentsMargins(6, 0, 6, 3)
        box.setSpacing(0)
        btn = QPushButton(name)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setToolTip(name + '：' + mode)
        btn.clicked.connect(callback)
        desc = QLabel(mode)
        desc.setObjectName('tiledesc')
        desc.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        box.addWidget(btn)
        box.addWidget(desc)


class NetWorker(QThread):
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, fn, parent=None):
        super().__init__(parent)
        self.fn = fn

    def run(self):
        try:
            self.done.emit(self.fn())
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))


class LiteraturePage(QWidget):
    def __init__(self, store, workspace):
        super().__init__()
        self.store = store
        self.workspace = workspace
        self.rows = []
        self._net_worker = None
        self._search_worker = None
        self._net_checked = False
        self._drop_results = False
        self.build_ui()

    def build_ui(self):
        self.setObjectName('glasspage')
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        host.setObjectName('glasspage')
        host.setMinimumWidth(MIN_BODY_WIDTH)
        lay = QVBoxLayout(host)
        lay.setContentsMargins(28, 20, 28, 18)
        lay.setSpacing(10)

        head = QHBoxLayout()
        head.setSpacing(8)
        head.addWidget(label('文献检索', 'title'))
        head.addWidget(label('联网', 'badge'))
        self.net_badge = label('未检测', 'net')
        set_state(self.net_badge, state='idle')
        head.addWidget(self.net_badge)
        head.addStretch()
        self.net_btn = button('联网检测', self.check_network)
        self.net_btn.setObjectName('tool')
        head.addWidget(self.net_btn)
        lay.addLayout(head)
        lay.addWidget(label('e2sci 检索入口 · 知网/万方/Google Scholar 等平台跳转 · '
                             'OpenAlex/Crossref 题录 · 只取公开元数据，不下载全文。', 'subtitle'))

        body = QHBoxLayout()
        body.setSpacing(12)
        lay.addLayout(body, 1)

        # ------------------------------------------------------------------ 左栏
        side = QWidget()
        side.setFixedWidth(LEFT_WIDTH)
        left = QVBoxLayout(side)
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(12)
        body.addWidget(side)

        query_card = QFrame()
        query_card.setObjectName('card')
        qv = QVBoxLayout(query_card)
        qv.setContentsMargins(16, 14, 16, 14)
        qv.setSpacing(8)
        qv.addWidget(label('检索词', 'section'))
        self.query = QLineEdit()
        self.query.setPlaceholderText('输入题目、关键词或研究问题')
        self.query.returnPressed.connect(self.open_e2sci)
        qv.addWidget(self.query)
        row = QHBoxLayout()
        row.setSpacing(8)
        for text, callback in (('用当前模板标题', self.fill_from_template),
                               ('精简检索词', self.simplify_query),
                               ('复制检索词', self.copy_query)):
            item = button(text, callback)
            item.setObjectName('tool')
            row.addWidget(item, 1)
        qv.addLayout(row)
        qv.addWidget(label('e2sci 是对话式检索：打开后把检索词粘贴到输入框，点 Search；'
                           '其他平台会尽量用 URL 预填检索词。', 'muted'))
        left.addWidget(query_card)

        es = QFrame()
        es.setObjectName('accentcard')
        eh = QVBoxLayout(es)
        eh.setContentsMargins(18, 16, 18, 16)
        eh.setSpacing(5)
        title = QHBoxLayout()
        title.setSpacing(8)
        title.addWidget(label('e2sci · 元空 AI Science', 'section'))
        title.addWidget(label('主入口', 'tag'))
        title.addStretch()
        eh.addLayout(title)
        eh.addWidget(label('AI for Science 平台：调研可信科研文献、阅读与整理。'
                           '入口已绑定你的 partner_uuid 与 local_login。', 'muted'))
        eh.addWidget(label('partner_uuid：' + literature.E2SCI_PARTNER_UUID, 'stylemeta'))
        eh.addSpacing(6)
        eh.addWidget(button('打开 e2sci 检索', self.open_e2sci, primary=True))
        eh.addWidget(label('检索词会同时复制到剪贴板', 'muted'))
        left.addWidget(es)

        plat_card = QFrame()
        plat_card.setObjectName('card')
        pv = QVBoxLayout(plat_card)
        pv.setContentsMargins(16, 14, 16, 14)
        pv.setSpacing(6)
        pv.addWidget(label('其他平台', 'section'))
        pv.addWidget(label('点平台按钮打开系统浏览器；知网/万方不抓全文，只跳转到检索页。', 'muted'))
        grid = QGridLayout()
        grid.setSpacing(8)
        others = [p for p in literature.PLATFORMS if p['id'] != 'e2sci']
        for index, plat in enumerate(others):
            mode = 'URL 预填' if plat.get('query') else '需粘贴'
            tile = PlatformTile(plat['name'], mode,
                                lambda _=False, pid=plat['id']: self.open_platform(pid))
            grid.addWidget(tile, index // 3, index % 3)
        for column in range(3):
            grid.setColumnStretch(column, 1)
        pv.addLayout(grid)
        left.addWidget(plat_card)
        left.addStretch()

        # ------------------------------------------------------------------ 右栏
        meta = QFrame()
        meta.setObjectName('card')
        body.addWidget(meta, 1)
        mv = QVBoxLayout(meta)
        mv.setContentsMargins(16, 14, 16, 14)
        mv.setSpacing(8)
        top = QHBoxLayout()
        top.setSpacing(8)
        top.addWidget(label('在线题录检索', 'section'))
        top.addStretch()
        tip = label('OpenAlex / Crossref 公开接口，无需 Key', 'aihint')
        tip.setWordWrap(False)
        tip.setToolTip('OpenAlex 限流时自动回退 Crossref，不需要额外配置。')
        tip.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        top.addWidget(tip)
        mv.addLayout(top)

        bar = QHBoxLayout()
        bar.setSpacing(6)
        bar.addWidget(label('来源', 'muted'))
        self.source = QComboBox()
        self.source.addItem('OpenAlex', 'openalex')
        self.source.addItem('Crossref', 'crossref')
        self.source.setMinimumWidth(104)
        bar.addWidget(self.source)
        bar.addWidget(label('条数', 'muted'))
        self.count = QComboBox()
        for value in (10, 20, 30, 50):
            self.count.addItem(str(value), value)
        self.count.setCurrentIndex(1)
        self.count.setMinimumWidth(72)
        bar.addWidget(self.count)
        self.year = QLineEdit()
        self.year.setPlaceholderText('起始年，可空')
        self.year.setFixedWidth(110)
        bar.addWidget(self.year)
        bar.addStretch()
        self.search_btn = button('联网检索题录', self.search_metadata, primary=True)
        bar.addWidget(self.search_btn)
        mv.addLayout(bar)

        self.task = feedback.TaskBar('停止等待')
        self.task.cancelled.connect(self.cancel_search)
        mv.addWidget(self.task)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(['年', '题名', '作者', '来源', 'DOI'])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setMinimumHeight(240)
        self.table.setWordWrap(False)
        self.table.setTextElideMode(Qt.TextElideMode.ElideRight)
        self.table.itemDoubleClicked.connect(lambda *_: self.open_doi())
        self.table.itemSelectionChanged.connect(self.refresh_selection)
        self.table.currentCellChanged.connect(lambda *_: self.refresh_selection())
        mv.addWidget(self.table, 1)
        self.apply_column_widths()

        status = QHBoxLayout()
        status.setSpacing(8)
        self.search_status = label('还没有检索。', 'muted')
        status.addWidget(self.search_status, 1)
        self.select_badge = label('未选中', 'aihint')
        self.select_badge.setVisible(False)
        status.addWidget(self.select_badge)
        mv.addLayout(status)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        actions.addStretch()
        actions.addWidget(button('复制选中 GB/T', self.copy_selected))
        actions.addWidget(button('复制全部', self.copy_all))
        actions.addWidget(button('打开 DOI', self.open_doi))
        mv.addLayout(actions)

        scroll.setWidget(host)
        root.addWidget(scroll)

    def apply_column_widths(self):
        self.table.setColumnWidth(0, 52)
        self.table.setColumnWidth(2, 150)
        self.table.setColumnWidth(3, 100)
        self.table.setColumnWidth(4, 150)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)

    # ------------------------------------------------------------------ helpers
    def current_query(self) -> str:
        return literature.clean_query(self.query.text())

    def fill_from_template(self, _=None, silent=False):
        config = getattr(self.workspace.fill, 'current', None)
        title = config.get('title') if isinstance(config, dict) else ''
        if not title:
            title = self.workspace.search.text().strip()
        if title:
            self.query.setText(title)
            self.search_status.setText('已填入当前模板标题：' + title[:60])
        elif not silent:
            self.search_status.setText('当前没有选中模板，请手动输入检索词。')

    def simplify_query(self):
        text = self.current_query()
        if not text:
            self.search_status.setText('先填写检索词。')
            return
        simplified = literature.simplify(text)
        if simplified:
            self.query.setText(simplified)
            self.search_status.setText('已精简为：' + simplified)

    def copy_query(self):
        text = self.current_query()
        if not text:
            self.search_status.setText('先填写检索词。')
            return
        QApplication.clipboard().setText(text)
        self.search_status.setText('检索词已复制，可直接粘贴到 e2sci 或其他平台。')
        self.workspace.notify('检索词已复制，粘贴到 e2sci 或平台搜索框即可', 'ok', 2200)

    def open_e2sci(self):
        text = self.current_query()
        if text:
            QApplication.clipboard().setText(text)
        QDesktopServices.openUrl(QUrl(literature.E2SCI_URL))
        self.search_status.setText('已打开 e2sci' + ('；检索词已复制，粘贴到输入框后点 Search。' if text else '。'))

    def open_platform(self, pid: str):
        item = next((p for p in literature.PLATFORMS if p['id'] == pid), None)
        if not item:
            return
        text = self.current_query()
        if pid == 'e2sci':
            self.open_e2sci()
            return
        if not text:
            self.search_status.setText('先填写检索词，再打开 ' + item['name'] + '。')
            return
        QDesktopServices.openUrl(QUrl(literature.platform_url(pid, text)))
        self.search_status.setText('已打开 ' + item['name'] + '：' + text)

    # ------------------------------------------------------------------ network
    def on_show(self):
        if not self.query.text().strip():
            self.fill_from_template(silent=True)
        if not self._net_checked:
            self._net_checked = True
            QTimer.singleShot(250, self.check_network)

    def check_network(self):
        if self._net_worker is not None and self._net_worker.isRunning():
            return
        self.net_badge.setText('检测中…')
        set_state(self.net_badge, state='idle')
        self.net_btn.setEnabled(False)
        worker = NetWorker(literature.check_online, self)
        self._net_worker = worker
        worker.done.connect(self._net_done)
        worker.failed.connect(self._net_failed)
        worker.finished.connect(lambda: self.net_btn.setEnabled(True))
        worker.start()

    def _net_done(self, result):
        ok, message = result if isinstance(result, tuple) else (False, str(result))
        self.net_badge.setText('联网正常' if ok else '无法联网')
        set_state(self.net_badge, state='ok' if ok else 'bad')
        self.net_badge.setToolTip(message)
        self.search_status.setText(message)
        self._net_worker = None

    def _net_failed(self, message):
        self.net_badge.setText('检测失败')
        set_state(self.net_badge, state='bad')
        self.net_badge.setToolTip(message)
        self.search_status.setText('联网检测失败：' + message)
        self._net_worker = None

    def search_metadata(self):
        text = self.current_query()
        if not text:
            self.search_status.setText('先填写检索词。')
            return
        if self._search_worker is not None and self._search_worker.isRunning():
            return
        source = self.source.currentData()
        rows = int(self.count.currentData())
        raw_year = re.sub(r"\D", "", self.year.text() or "")
        year = int(raw_year) if raw_year else None
        self.search_btn.setEnabled(False)
        self.search_btn.setText('检索中…')
        self._drop_results = False
        self.search_status.setText('正在连接 ' + ('OpenAlex' if source == 'openalex' else 'Crossref') + '…')
        self.task.start('正在连接 ' + ('OpenAlex' if source == 'openalex' else 'Crossref')
                        + '，公开接口一般几秒返回…')
        worker = NetWorker(lambda: literature.search(source, text, rows, year), self)
        self._search_worker = worker
        worker.done.connect(self._search_done)
        worker.failed.connect(self._search_failed)
        worker.finished.connect(self._search_finished)
        worker.start()

    def cancel_search(self):
        """联网请求是一次性的，取消就是不再等：结果回来直接丢弃，不写任何文件。"""
        self._drop_results = True
        self.search_status.setText('已停止等待：这次检索的结果不会被采用。')

    def _search_finished(self):
        self.search_btn.setEnabled(True)
        self.search_btn.setText('联网检索题录')
        self.task.finish()
        self._search_worker = None

    def _search_done(self, rows):
        if self._drop_results:
            self.search_status.setText('已取消：本次检索到的题录没有采用，表里仍是上一次的结果。')
            self._drop_results = False
            return
        self.rows = list(rows or [])
        self.table.setRowCount(0)
        for index, ref in enumerate(self.rows):
            self.table.insertRow(index)
            values = [ref.get('year', ''), ref.get('title', ''), ref.get('authors', ''),
                      ref.get('source', ''), ref.get('doi', '')]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or ''))
                if column in (1, 2, 4):
                    item.setToolTip(str(value or ''))
                self.table.setItem(index, column, item)
            self.table.item(index, 0).setData(Qt.ItemDataRole.UserRole, ref)
        self.apply_column_widths()
        source = self.rows[0].get('source', '') if self.rows else ''
        self.search_status.setText(f'找到 {len(self.rows)} 条题录' + (f'（{source}）' if source else '。'))
        self.refresh_selection()

    def _search_failed(self, message):
        self.search_status.setText('检索失败：' + message)

    # ------------------------------------------------------------------ table
    def selected_refs(self):
        rows = sorted({index.row() for index in self.table.selectedIndexes()})
        if not rows and self.table.currentRow() >= 0:
            rows = [self.table.currentRow()]
        return [self.rows[i] for i in rows if 0 <= i < len(self.rows)]

    def refresh_selection(self, *_):
        count = len(self.selected_refs())
        total = len(self.rows)
        if not total:
            self.select_badge.setText('未选中')
        elif count:
            self.select_badge.setText(f'已选中 {count} / {total} 条')
        else:
            self.select_badge.setText(f'共 {total} 条，未选中')
        self.select_badge.setVisible(bool(total))

    def copy_selected(self):
        refs = self.selected_refs()
        if not refs:
            self.search_status.setText('先在表里选中要复制的行（可以按住 Ctrl 多选）。')
            return
        QApplication.clipboard().setText(literature.format_gbt_list(refs))
        self.search_status.setText(f'已复制选中的 {len(refs)} 条 GB/T 7714 题录。')
        self.workspace.notify(f'已复制 {len(refs)} 条 GB/T 7714 题录', 'ok', 2200)

    def copy_all(self):
        if not self.rows:
            self.search_status.setText('还没有可复制的题录。')
            return
        QApplication.clipboard().setText(literature.format_gbt_list(self.rows))
        self.search_status.setText(f'已复制全部 {len(self.rows)} 条题录。')
        self.workspace.notify(f'已复制全部 {len(self.rows)} 条题录', 'ok', 2200)

    def open_doi(self):
        refs = self.selected_refs()
        if not refs:
            self.search_status.setText('先选中一行题录。')
            return
        ref = refs[0]
        url = ref.get('url') or (('https://doi.org/' + ref['doi']) if ref.get('doi') else '')
        if not url:
            self.search_status.setText('这条题录没有 DOI / URL。')
            return
        QDesktopServices.openUrl(QUrl(url))
        self.search_status.setText('已打开：' + url)
