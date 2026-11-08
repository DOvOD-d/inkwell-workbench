# -*- coding: utf-8 -*-
"""启动动效页：流体磨砂玻璃背景 + 泼墨标志，默认 1.2 秒自动关闭，点一下或按任意键提前进主界面。"""
from __future__ import annotations

from PySide6.QtCore import Qt, QTimer, QEventLoop, Signal, QByteArray
from PySide6.QtGui import QColor, QFont, QImage, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication, QFrame, QLabel, QStackedLayout, QVBoxLayout, QWidget,
)

import glassfx

DURATION_MS = 1200


def _mark(path, size):
    """QPixmap 直接读 SVG 要 imageformats 插件，这里用 QtSvg 自己画。"""
    try:
        from PySide6.QtSvg import QSvgRenderer
        renderer = QSvgRenderer(QByteArray(path.read_bytes()))
    except Exception:
        return None
    if not renderer.isValid():
        return None
    image = QImage(size, size, QImage.Format.Format_ARGB32)
    image.fill(QColor(0, 0, 0, 0))
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    renderer.render(painter)
    painter.end()
    return QPixmap.fromImage(image)


class SplashWindow(QWidget):
    finished = Signal()

    def __init__(self, assets_dir, theme='light', version=''):
        super().__init__(None, Qt.WindowType.FramelessWindowHint
                         | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowTitle('Word 模板工作台')
        self.setFixedSize(560, 320)
        dark = theme == 'dark'
        layers = QStackedLayout(self)
        layers.setStackingMode(QStackedLayout.StackingMode.StackAll)
        layers.setContentsMargins(0, 0, 0, 0)
        self.backdrop = glassfx.GlassBackdrop()
        self.backdrop.set_dark(dark)
        layers.addWidget(self.backdrop)
        layer = QWidget()
        # 内容层不能自己刷背景，否则会把底下的流体玻璃盖掉
        layer.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)
        layers.addWidget(layer)
        layers.setCurrentWidget(layer)     # StackAll 下当前层会被抬高，内容层必须在最上
        outer = QVBoxLayout(layer)
        outer.setContentsMargins(34, 26, 34, 26)
        panel = QFrame()
        panel.setObjectName('splashpanel')
        panel.setStyleSheet(
            'QFrame#splashpanel { background: %s; border: 1px solid %s;'
            ' border-radius: 14px; }'
            % ('rgba(15,17,21,0.55)' if dark else 'rgba(255,255,255,0.62)',
               '#1e2530' if dark else '#d7e3ef'))
        outer.addWidget(panel)
        box = QVBoxLayout(panel)
        box.setContentsMargins(30, 24, 30, 24)
        box.setSpacing(9)
        mark = QLabel()
        name = 'logo-ink-white.svg' if dark else 'logo-ink.svg'
        pixmap = _mark(assets_dir / name, 72)
        if pixmap is not None:
            mark.setPixmap(pixmap)
        mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        box.addWidget(mark)
        title = QLabel('Word 模板工作台')
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont('Microsoft YaHei UI', 19, QFont.Weight.Bold))
        title.setStyleSheet('color: %s; background: transparent;'
                            % ('#eef4fb' if dark else '#0d2740'))
        box.addWidget(title)
        sub = QLabel(f'v{version} · 本机运行，文档内容不出这台电脑' if version
                     else '本机运行，文档内容不出这台电脑')
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setStyleSheet('color: %s; font-size: 12px; background: transparent;'
                          % ('#b9cfe4' if dark else '#2b4a68'))
        box.addWidget(sub)

    def center_on_screen(self):
        screen = QApplication.primaryScreen()
        if screen is None:
            return
        area = screen.availableGeometry()
        self.move(area.center().x() - self.width() // 2,
                  area.center().y() - self.height() // 2)

    def mousePressEvent(self, event):
        self.finish()

    def keyPressEvent(self, event):
        self.finish()

    def finish(self):
        if self.isVisible():
            self.finished.emit()


def show(assets_dir, theme='light', version='', ms=DURATION_MS):
    """弹启动页并跑一个短事件循环；返回一个 callable 用于确保关掉它。"""
    splash = SplashWindow(assets_dir, theme, version)
    splash.center_on_screen()
    splash.show()
    QApplication.processEvents()
    loop = QEventLoop()
    splash.finished.connect(loop.quit)
    timer = QTimer()
    timer.setSingleShot(True)
    timer.timeout.connect(loop.quit)
    timer.start(max(200, int(ms)))
    loop.exec()
    timer.stop()
    splash.hide()
    splash.deleteLater()
    QApplication.processEvents()
    return splash
