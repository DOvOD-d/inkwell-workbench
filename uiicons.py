# -*- coding: utf-8 -*-
"""界面小图标：按主题色现场渲染的矢量线条图标。

不存一堆 svg 文件是因为图标要跟主题走：白天 #5b6472、黑夜 #a8b3c4，
同一份路径描一遍就行，省得每个图标出两套深浅资源。
注册表记下用图标的控件，切主题时统一换色，不用每个页面自己管。
"""
from __future__ import annotations

import weakref

from PySide6.QtCore import QByteArray, QSize
from PySide6.QtGui import QColor, QIcon, QImage, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer

STROKE_WIDTH = 1.7

# @ 是描边颜色的占位符，渲染前替换成主题色
SHAPES = {
    'moon': '<path d="M15.6 3.4a8.2 8.2 0 1 0 5 12.7 6.7 6.7 0 0 1-5-12.7z"/>',
    'sun': ('<circle cx="12" cy="12" r="3.9"/>'
            '<path d="M12 2.6v2.5M12 18.9v2.5M2.6 12h2.5M18.9 12h2.5'
            'M5.4 5.4l1.8 1.8M16.8 16.8l1.8 1.8M18.6 5.4l-1.8 1.8M7.2 16.8l-1.8 1.8"/>'),
    'sparkle': ('<path d="M11.4 3.6l1.9 5 5 1.9-5 1.9-1.9 5-1.9-5-5-1.9 5-1.9z"/>'
                '<path d="M18.4 15.4l.8 2.1 2.1.8-2.1.8-.8 2.1-.8-2.1-2.1-.8 2.1-.8z"/>'),
    'help': ('<circle cx="12" cy="12" r="8.6"/>'
             '<path d="M9.5 9.4a2.6 2.6 0 0 1 5.1.7c0 1.7-2.5 2-2.5 3.7"/>'
             '<circle cx="12" cy="17.2" r="1" fill="@" stroke="none"/>'),
    'up': '<polyline points="6.5,14.8 12,8.8 17.5,14.8"/>',
    'down': '<polyline points="6.5,9.2 12,15.2 17.5,9.2"/>',
    'zoom-in': ('<circle cx="10.6" cy="10.6" r="6.4"/><path d="M15.3 15.3l4.6 4.6"/>'
                '<path d="M10.6 7.9v5.4M7.9 10.6h5.4"/>'),
    'zoom-out': '<circle cx="10.6" cy="10.6" r="6.4"/><path d="M15.3 15.3l4.6 4.6"/>'
                '<path d="M7.9 10.6h5.4"/>',
    'close': '<path d="M7 7l10 10M17 7L7 17"/>',
    'undo': ('<path d="M4.8 9.6h9.7a4.6 4.6 0 0 1 0 9.2h-4.1"/>'
             '<polyline points="8.4,5.6 4.4,9.6 8.4,13.6"/>'),
    'redo': ('<path d="M19.2 9.6H9.5a4.6 4.6 0 0 0 0 9.2h4.1"/>'
             '<polyline points="15.6,5.6 19.6,9.6 15.6,13.6"/>'),
}

# 角色 → 主题令牌名：按钮用弱色，强调位置用主色
ROLES = {'muted': 'text_muted', 'text': 'text', 'strong': 'strong', 'accent': 'accent'}

_template = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" '
             'width="{px}" height="{px}"><g fill="none" stroke="{color}" '
             'stroke-width="{sw}" stroke-linecap="round" stroke-linejoin="round">'
             '{body}</g></svg>')

_registry = []       # [(弱引用(控件), 图标名, 角色, 像素边长)]
_theme = {}


def _colour(theme, role):
    key = ROLES.get(role, 'text_muted')
    return theme.get(key) or '#5b6472'


def _live_entries():
    """取出还活着的按钮；控件已销毁的条目顺手清掉，避免切主题时碰到悬空引用。"""
    alive = []
    for ref, name, role, px in _registry:
        widget = ref()
        if widget is not None:
            alive.append((widget, name, role, px))
    _registry[:] = [(weakref.ref(widget), name, role, px)
                    for widget, name, role, px in alive]
    return alive


def pixmap(name, color, px=18):
    body = SHAPES.get(name)
    if not body:
        return None
    svg = _template.format(px=px, color=color, sw=STROKE_WIDTH, body=body.replace('@', color))
    renderer = QSvgRenderer(QByteArray(svg.encode('utf-8')))
    if not renderer.isValid():
        return None
    image = QImage(px, px, QImage.Format.Format_ARGB32)
    image.fill(QColor(0, 0, 0, 0))
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    renderer.render(painter)
    painter.end()
    return QPixmap.fromImage(image)


def icon(theme, name, role='muted', px=18):
    """给图标挑几个尺寸，一次渲好，缩放到高 DPI 也不会糊。"""
    color = _colour(theme, role)
    out = QIcon()
    for size in {max(12, px - 2), px, px + 6, px * 2}:
        shot = pixmap(name, color, size)
        if shot is not None:
            out.addPixmap(shot)
    return out


def set_theme(theme):
    """主题变了：记下配色并把注册过的按钮全部换色。"""
    global _theme
    _theme = theme or {}
    for widget, name, role, px in _live_entries():
        try:
            widget.setIcon(icon(_theme, name, role, px))
        except RuntimeError:
            continue          # C++ 侧已销毁，跳过


def rename(widget, name):
    """同一个按钮换图形（白天/黑夜这类），角色和尺寸跟着原来的走。"""
    entries = _live_entries()
    for index, (entry_widget, _old, role, px) in enumerate(entries):
        if entry_widget is widget:
            _registry[index] = (weakref.ref(widget), name, role, px)
            try:
                widget.setIcon(icon(_theme or {}, name, role, px))
            except RuntimeError:
                pass
            return


def apply(widget, name, tip, role='muted', px=18, side=30):
    """图标按钮：渲得出来就用图标，渲不出来（没有 QtSvg）就留着原文字，不至于空白。"""
    widget.setToolTip(tip)
    candidate = icon(_theme or {}, name, role, px)
    if candidate.isNull():
        widget.setText(tip)      # 渲不出来就退回文字，不留空按钮
        return widget
    widget.setIcon(candidate)
    widget.setIconSize(QSize(px, px))
    widget.setProperty('iconic', True)
    widget.setFixedSize(side, side)
    _registry.append((weakref.ref(widget), name, role, px))
    return widget
