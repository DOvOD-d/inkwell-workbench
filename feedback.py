# -*- coding: utf-8 -*-
"""长任务进度行与右下角轻提示：两处都要用，所以单独成模块，避免 literature_page 反过来 import main。"""
from __future__ import annotations

from PySide6.QtCore import QEvent, QTimer, Signal
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QProgressBar, QPushButton


class Cancelled(RuntimeError):
    """用户取消：在下一个进度点抛出，调用方把结果丢弃即可，已经落盘的文件不动。"""


class TaskBar(QFrame):
    """一行进度：说明 + 进度条 + 取消。没有任务时整行隐藏，不占版面。

    取消是协作式的：工作循环轮询 is_cancelled()，做到当前这一项就收手，
    已经写出去的文件不回滚，剩下的不写——调用方负责把这句话如实说清楚。
    """

    cancelled = Signal()

    def __init__(self, cancel_text='取消', parent=None):
        super().__init__(parent)
        self.setObjectName('taskbar')
        self._cancel_requested = False
        self.hide()
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(10)
        self.info = QLabel('')
        self.info.setObjectName('tasktext')
        lay.addWidget(self.info, 1)
        self.bar = QProgressBar()
        self.bar.setTextVisible(False)
        self.bar.setFixedWidth(180)
        lay.addWidget(self.bar)
        self.cancel_btn = QPushButton(cancel_text)
        self.cancel_btn.setObjectName('tool')
        self.cancel_btn.clicked.connect(self._on_cancel)
        lay.addWidget(self.cancel_btn)

    def _on_cancel(self):
        self._cancel_requested = True
        self.cancel_btn.setEnabled(False)
        self.info.setText('正在取消：做完当前这一项就停下，已生成的文件不会删除。')
        self.cancelled.emit()

    def is_cancelled(self) -> bool:
        return self._cancel_requested

    def start(self, text, total=0):
        self._cancel_requested = False
        self.cancel_btn.setEnabled(True)
        self.info.setText(text)
        self.bar.setRange(0, max(0, int(total)))
        self.bar.setValue(0)
        if not total:
            self.bar.setRange(0, 0)          # 不确定进度：来回滚动
        self.show()

    def step(self, done, total=None, text=''):
        if total is not None:
            self.bar.setRange(0, max(1, int(total)))
            self.bar.setMaximum(max(self.bar.maximum(), int(total)))
        self.bar.setValue(int(done))
        if text:
            self.info.setText(text)

    def finish(self):
        self.hide()
        self.bar.setRange(0, 0)
        self._cancel_requested = False


class Toast(QFrame):
    """右下角浮层，默认 1.8 秒自动消失，点「关闭」立即收掉。用来替代“必须点确定”的成功弹框。"""

    def __init__(self, parent):
        super().__init__(parent)
        self.setObjectName('toast')
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 9, 14, 9)
        lay.setSpacing(8)
        self.label = QLabel('')
        self.label.setObjectName('toasttext')
        self.label.setWordWrap(True)
        self.label.setMaximumWidth(520)
        lay.addWidget(self.label)
        self.dismiss = QPushButton('关闭')
        self.dismiss.setObjectName('mini')
        self.dismiss.clicked.connect(self.hide)
        lay.addWidget(self.dismiss)
        self.hide()
        if parent is not None:
            parent.installEventFilter(self)

    def eventFilter(self, watched, event):
        if watched is self.parent() and event.type() in (QEvent.Type.Resize, QEvent.Type.Show):
            self._place()
        return False

    def _place(self):
        host = self.parent()
        if host is None:
            return
        self.adjustSize()
        self.move(host.width() - self.width() - 24, host.height() - self.height() - 24)

    def say(self, text, kind='info', ms=1800):
        self.setProperty('kind', kind)
        self.style().unpolish(self)
        self.style().polish(self)
        self.label.setText(str(text).strip() or '完成')
        self.adjustSize()
        self._place()
        self.show()
        self.raise_()
        self._timer.start(max(600, int(ms)))
