"""Word 模板工作台 v5.15.0 —— 借 OfficeCLI 的做法补齐设置能力：单位/颜色自由输入与范围校验、分页与底纹与回车继承、名单批量填写、生成后 OOXML 自检、错误提示带下一步。原 v5.14.0 识别按四个科研项目模板校正：{{记号}}优先、重复行做成表格格填写、同名基础信息归并、表头与公章类噪声丢弃。原 v5.13.0 在 v5.12.0 基础上新增联网“文献检索”页：e2sci 主入口、知网/万方/Google Scholar 等平台跳转、OpenAlex/Crossref 题录与 GB/T 复制。

顶栏：品牌 + 全模板搜索 + 白天/黑夜主题切换 + AI 设置 + 使用说明。
左侧导航四项：首页（模板卡片墙，点卡片直接进填写页） / 填写生成（直接拖入 docx 或选模板） / 模板制作（自动识别 + 智能打标） / 自动排版（整篇文档重排）。
模板库不再是导航页：首页「模板库」卡片或模板卡片上的入口打开同一个弹窗，导入/重命名/导出/删除/体检功能不变。
自动排版页补齐段前段后间距与颜色档，样式库可「新建样式」把自己的参数存成常用样式。
AI 支持本机 Ollama、任意 OpenAI 兼容服务，以及内置离线小模型。
识别归并：同一份表里反复出现的基础信息（学生/姓名/学生姓名、院系/学院/所在学院、班次/班级）并成一个栏目，填一次写到所有位置；规范/说明材料里的参考文献示例行不再被当成栏目。
数据目录与旧版一致：%LOCALAPPDATA%\\LocalTools\\WordTemplateTool，草稿、模板与字号设置直接沿用。
"""
from __future__ import annotations

import copy
import json
import re
import shutil
import sys
import tempfile
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QThread, QStandardPaths, QUrl, Signal, QByteArray
from PySide6.QtGui import (QColor, QFont, QIcon, QImage, QPainter, QPixmap, QDesktopServices,
                           QKeySequence, QShortcut)
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QLabel,
    QPushButton, QLineEdit, QPlainTextEdit, QComboBox, QCheckBox, QListWidget,
    QListWidgetItem, QTabWidget, QScrollArea, QFrame, QFileDialog, QMessageBox,
    QDialog, QDialogButtonBox, QFormLayout, QTableWidget, QTableWidgetItem,
    QHeaderView, QSplitter, QAbstractItemView, QStackedWidget, QStackedLayout, QDoubleSpinBox,
    QSpinBox, QRadioButton, QButtonGroup, QGridLayout,
)

import ai
import batchfill
import blockfill
import bookmarkfill
import directfill
import embedded_ai
import exampleparse
import feedback
import glassfx
import layout
import literature
import literature_page
import ooxmlcheck
import pipeline
import preview
import reqparse
import smartparse
import splash
import textparse
import units
import uiicons
import visionparse
from engine import read_json, write_json, render_word, scan_fields, validate_values

ASSETS = Path(__file__).resolve().parent / 'assets'
VERSION = '1.0.5'
# 自动排版和“删说明搬提示”不再暴露开关，始终执行。
REQ_OPTIONS = {'apply_format': True, 'remove_notes': True}
CN_SIZES = [('初号', 42), ('小初', 36), ('一号', 26), ('小一', 24), ('二号', 22), ('小二', 18),
            ('三号', 16), ('小三', 15), ('四号', 14), ('小四', 12), ('五号', 10.5), ('小五', 9)]

THEMES = {
    'light': {
        'name': '白天', 'bg': '#f0f2f5', 'surface': '#ffffff', 'surface_alt': '#f9fafb',
        'border': '#e5e7eb', 'border_soft': '#f3f4f6', 'text': '#1f2937',
        'text_muted': '#5b6472', 'text_faint': '#8b93a1', 'strong': '#111827',
        'accent': '#2563eb', 'accent_hover': '#1d4ed8', 'accent_soft': '#eff6ff',
        'on_accent': '#ffffff', 'danger': '#b42318', 'danger_soft': '#fef3f2',
        'sidebar_bg': '#fafbfc', 'sidebar_fg': '#374151', 'sidebar_muted': '#8c96a3',
        'nav_fg': '#4b5563', 'nav_active_bg': '#eff6ff', 'nav_active_fg': '#2563eb',
        'nav_hover_bg': '#f3f4f6', 'list_bg': '#ffffff', 'list_fg': '#374151',
        'list_active_bg': '#eff6ff', 'list_active_fg': '#1d4ed8', 'list_hover_bg': '#f3f4f6',
        'input_bg': '#ffffff', 'input_border': '#d1d5db', 'header_bg': '#f9fafb',
        'header_fg': '#374151', 'grid': '#eef2f7', 'sel_bg': '#dbeafe', 'sel_fg': '#1e3a8a',
        'scroll_bg': '#f0f2f5', 'scroll_fg': '#cbd5e1', 'chip_bg': '#f0f9ff',
        'chip_border': '#bae6fd', 'chip_fg': '#0369a1', 'topbar_bg': '#ffffff',
        'panel_bg': '#ffffff', 'tooltip_bg': '#1f2937', 'tooltip_fg': '#f9fafb',
        'menu_bg': '#ffffff', 'menu_sel': '#eff6ff', 'disabled_fg': '#9ca3af',
        'disabled_bg': '#f3f4f6', 'brand_bg': '#2563eb', 'paper_bg': '#e5e7eb',
    },
    'dark': {
        'name': '黑夜', 'bg': '#0f1115', 'surface': '#171a21', 'surface_alt': '#1b2029',
        'border': '#262b36', 'border_soft': '#1f242e', 'text': '#e6eaf2',
        'text_muted': '#9aa4b2', 'text_faint': '#6c7684', 'strong': '#f5f7fa',
        'accent': '#3b82f6', 'accent_hover': '#60a5fa', 'accent_soft': '#10233f',
        'on_accent': '#ffffff', 'danger': '#f87171', 'danger_soft': '#3b1d1d',
        'sidebar_bg': '#0b0d11', 'sidebar_fg': '#dfe6f0', 'sidebar_muted': '#7b8798',
        'nav_fg': '#c7d0dd', 'nav_active_bg': '#10233f', 'nav_active_fg': '#60a5fa',
        'nav_hover_bg': '#161a22', 'list_bg': '#171a21', 'list_fg': '#cbd5e1',
        'list_active_bg': '#10233f', 'list_active_fg': '#60a5fa', 'list_hover_bg': '#1b2029',
        'input_bg': '#11141b', 'input_border': '#2b3140', 'header_bg': '#1b2029',
        'header_fg': '#b6c0cf', 'grid': '#262b36', 'sel_bg': '#1e3a5f', 'sel_fg': '#e6eaf2',
        'scroll_bg': '#0f1115', 'scroll_fg': '#39404d', 'chip_bg': '#0f2436',
        'chip_border': '#1e4a6b', 'chip_fg': '#7dd3fc', 'topbar_bg': '#12151b',
        'panel_bg': '#171a21', 'tooltip_bg': '#262b36', 'tooltip_fg': '#e6eaf2',
        'menu_bg': '#171a21', 'menu_sel': '#10233f', 'disabled_fg': '#5b6472',
        'disabled_bg': '#1b2029', 'brand_bg': '#3b82f6', 'paper_bg': '#0b0d11',
    },
}


def _rgba(hex_color, alpha):
    """QSS 里带透明度的颜色只能写 rgba()，主题令牌存的是 #RRGGBB。"""
    value = (hex_color or '#000000').lstrip('#')
    if len(value) == 8:
        value = value[2:]
    return f'rgba({int(value[0:2], 16)}, {int(value[2:4], 16)}, {int(value[4:6], 16)}, {alpha})'


def svg_pixmap(path, size):
    """QPixmap 直接读 SVG 依赖 imageformats 插件，这里用 QtSvg 自己画，稳。"""
    try:
        renderer = QSvgRenderer(QByteArray(Path(path).read_bytes()))
    except OSError:
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


def svg_icon(path, sizes=(16, 24, 32, 48, 64, 128, 256)):
    icon = QIcon()
    for size in sizes:
        pixmap = svg_pixmap(path, size)
        if pixmap is not None:
            icon.addPixmap(pixmap)
    return icon if not icon.isNull() else None


def build_style(t):
    # 背景是一层静态玻璃，顶栏和侧栏改成半透明才透得出玻璃质感
    topbar_bg = _rgba(t['topbar_bg'], 0.72)
    sidebar_bg = _rgba(t['sidebar_bg'], 0.70)
    return f'''
QWidget {{ font-family: "Microsoft YaHei UI", "Segoe UI", "PingFang SC", sans-serif;
          font-size: 13px; color: {t['text']}; background: {t['bg']}; }}
QMainWindow, QDialog {{ background: {t['bg']}; }}
QToolTip {{ background: {t['tooltip_bg']}; color: {t['tooltip_fg']};
            border: 1px solid {t['border']}; padding: 6px 8px; border-radius: 5px; }}

QFrame#topbar {{ background: {topbar_bg}; border: none;
                 border-bottom: 1px solid {t['border']}; }}
QFrame#topbar QLabel {{ background: transparent; color: {t['text']}; }}
QLabel#brand {{ font-size: 15px; font-weight: 600; color: {t['strong']}; background: transparent; }}
QLabel#logo {{ background: {t['brand_bg']}; color: {t['on_accent']}; font-size: 15px;
               font-weight: 700; border-radius: 6px; padding: 6px 9px; }}
QLabel#versionchip {{ color: {t['text_muted']}; font-size: 11px; background: {t['surface_alt']};
                      border: 1px solid {t['border']}; border-radius: 10px; padding: 3px 9px; }}

QFrame#sidebar {{ background: {sidebar_bg}; border: none;
                  border-right: 1px solid {t['border']}; }}
QWidget#glasslayer {{ background: transparent; }}
QWidget#glasspage {{ background: transparent; }}
QStackedWidget {{ background: transparent; }}
QFrame#sidebar QLabel {{ color: {t['sidebar_fg']}; background: transparent; }}
QLabel#sidehint {{ color: {t['sidebar_muted']}; font-size: 12px; }}
QListWidget#nav {{ border: none; background: transparent; outline: none; }}
QListWidget#nav::item {{ color: {t['nav_fg']}; padding: 10px 14px 10px 11px; margin: 2px 8px;
                         border-radius: 6px; border-left: 3px solid transparent; }}
QListWidget#nav::item:selected {{ background: {t['nav_active_bg']};
                                  color: {t['nav_active_fg']}; font-weight: 600;
                                  border-left: 3px solid {t['accent']}; }}
QListWidget#nav::item:hover {{ background: {t['nav_hover_bg']}; color: {t['strong']}; }}

QLabel#title {{ font-size: 21px; font-weight: 700; color: {t['strong']}; background: transparent; }}
QLabel#subtitle {{ font-size: 13px; color: {t['text_muted']}; background: transparent; }}
QLabel#muted {{ color: {t['text_muted']}; background: transparent; }}
QLabel#section {{ font-size: 15px; font-weight: 600; color: {t['strong']}; background: transparent; }}
QLabel#badge {{ color: {t['text']}; background: {t['surface_alt']};
                border: 1px solid {t['border']}; padding: 4px 10px; border-radius: 10px;
                font-size: 12px; }}
QLabel#chip {{ color: {t['text_muted']}; background: {t['surface_alt']};
               border: 1px solid {t['border']}; border-radius: 6px; padding: 5px 10px;
               font-size: 12px; }}
QLabel#aihint {{ color: {t['text_muted']}; font-size: 12px; background: transparent; }}
QLabel#tag {{ font-size: 12px; color: {t['accent_hover']}; background: {t['surface']};
              border: 1px solid {t['chip_border']}; border-radius: 10px; padding: 3px 9px; }}
QLabel#net {{ font-size: 12px; padding: 4px 10px; border-radius: 10px;
              color: {t['text']}; background: {t['surface_alt']};
              border: 1px solid {t['border']}; }}
QLabel#net[state="ok"] {{ color: {t['chip_fg']}; background: {t['chip_bg']};
                          border: 1px solid {t['chip_border']}; }}
QLabel#net[state="bad"] {{ color: {t['danger']}; background: {t['danger_soft']};
                           border: 1px solid {t['danger']}; }}
QDialog QLabel#dialogtitle {{ font-size: 18px; font-weight: 700; color: {t['strong']}; }}

QFrame#paper, QFrame#footer, QFrame#card, QFrame#panel {{
    background: {t['surface']}; border: 1px solid {t['border']}; border-radius: 8px; }}
QFrame#panel {{ border-radius: 0px; border-left: none; border-right: none; border-top: none; }}
QFrame#cardhover {{ background: {t['surface']}; border: 1px solid {t['border']}; border-radius: 8px; }}
QFrame#groupcard {{ background: {t['surface_alt']}; border: 1px solid {t['border_soft']};
                    border-radius: 8px; }}
QFrame#accentcard {{ background: {t['accent_soft']}; border: 1px solid {t['accent']};
                     border-radius: 8px; }}
QFrame#tile {{ background: {t['surface']}; border: 1px solid {t['input_border']};
               border-radius: 7px; }}
QFrame#tile:hover {{ background: {t['surface_alt']}; border-color: {t['accent']}; }}
QFrame#tile QPushButton {{ background: transparent; border: none; color: {t['text']};
                           font-size: 12px; padding: 5px 2px 0 2px; }}
QFrame#tile QPushButton:hover {{ background: transparent; color: {t['accent']}; }}
QFrame#tile QPushButton:pressed {{ background: transparent; color: {t['accent_hover']}; }}
QLabel#tiledesc {{ font-size: 10px; color: {t['chip_fg']}; background: transparent; }}
QLabel#paneltitle {{ font-size: 13px; font-weight: 600; color: {t['strong']};
                     border-bottom: 1px solid {t['border_soft']}; padding: 13px 16px;
                     background: {t['surface']}; }}
QFrame#previewarea {{ background: {t['paper_bg']}; border: none; border-radius: 8px; }}
QWidget#formpage {{ background: {t['surface']}; }}
QLabel[over="true"] {{ color: {t['danger']}; font-weight: 600; }}
QFrame#cardhover:hover {{ border: 1px solid {t['accent']}; background: {t['accent_soft']}; }}
QFrame#cardhover QPushButton#mini {{ background: {t['surface_alt']};
                                    border: 1px solid {t['border']}; }}
QLabel#hero {{ font-size: 24px; font-weight: 700; color: {t['strong']}; background: transparent; }}
QLabel#cardmeta {{ font-size: 12px; color: {t['text_muted']}; background: transparent; }}
QFrame#stylecard {{ background: {t['surface_alt']}; border: 1px solid {t['border_soft']};
                      border-radius: 8px; }}
QFrame#stylecard:hover {{ border-color: {t['border']}; background: {t['surface']}; }}
QFrame#stylecard[selected="true"] {{ background: {t['accent_soft']};
                                     border: 1px solid {t['accent']}; }}
QLabel#stylename {{ font-size: 13px; font-weight: 600; color: {t['strong']}; background: transparent; }}
QLabel#stylemeta {{ font-size: 11px; color: {t['text_muted']}; background: transparent; }}

QLineEdit, QPlainTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
    background: {t['surface_alt']}; border: 1px solid transparent; border-radius: 7px;
    padding: 8px 10px; color: {t['text']}; selection-background-color: {t['sel_bg']};
    selection-color: {t['sel_fg']}; }}
QLineEdit:hover, QPlainTextEdit:hover, QComboBox:hover {{ border-color: {t['border']}; }}
QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus {{
    border: 1px solid {t['accent']}; background: {t['surface']}; }}
QLineEdit:disabled, QPlainTextEdit:disabled, QComboBox:disabled {{
    background: {t['disabled_bg']}; color: {t['disabled_fg']}; }}
QLineEdit#search {{ background: {t['surface_alt']}; border-radius: 17px; padding: 7px 14px; }}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox QAbstractItemView {{ background: {t['menu_bg']}; border: 1px solid {t['border']};
    selection-background-color: {t['menu_sel']}; selection-color: {t['text']}; outline: none; }}

QPushButton {{ background: {t['surface']}; border: 1px solid {t['input_border']};
               border-radius: 7px; padding: 8px 14px; color: {t['text']}; }}
QPushButton:hover {{ background: {t['surface_alt']}; border-color: {t['input_border']};
                     color: {t['text']}; }}
QPushButton:focus {{ border: 1px solid {t['accent']}; }}
QPushButton:pressed {{ background: {t['accent_soft']}; color: {t['accent']}; }}
QPushButton:disabled {{ color: {t['disabled_fg']}; background: {t['disabled_bg']};
                        border-color: {t['border']}; }}
QPushButton#primary {{ background: {t['accent']}; color: {t['on_accent']}; border: none;
                       font-weight: 600; padding: 10px 22px; }}
QPushButton#primary:hover {{ background: {t['accent_hover']}; color: {t['on_accent']}; }}
QPushButton#primary:disabled {{ background: {t['border']}; color: {t['disabled_fg']}; }}
QPushButton#danger {{ color: {t['danger']}; border-color: {t['danger']}; }}
QPushButton#danger:hover {{ background: {t['danger_soft']}; color: {t['danger']}; }}
QPushButton#ghost {{ background: transparent; border: 1px solid {t['border']};
                     color: {t['text_muted']}; padding: 6px 10px; font-size: 12px; }}
QPushButton#ghost:hover {{ background: {t['surface_alt']}; color: {t['text']}; }}
QPushButton#mini {{ padding: 5px 10px; font-size: 12px; color: {t['text_muted']};
                    background: transparent; border: 1px solid transparent; }}
QPushButton#mini:hover {{ background: {t['surface_alt']}; color: {t['text']};
                          border-color: {t['border']}; }}
QPushButton#tool {{ padding: 6px 10px; border-radius: 6px; background: transparent;
                    border: 1px solid {t['border']}; color: {t['text_muted']}; }}
QPushButton#tool:hover {{ background: {t['surface_alt']}; color: {t['text']}; }}
QPushButton#tool:checked {{ background: {t['accent_soft']}; color: {t['accent']};
                            border-color: {t['accent']}; }}
QPushButton#tool[iconic="true"] {{ padding: 0; }}
QPushButton#tool[iconic="true"]:disabled {{ background: transparent;
                                            border-color: {t['border']}; }}

QTabWidget::pane {{ border: none; background: {t['surface']}; }}
QTabBar::tab {{ padding: 10px 16px; color: {t['text_muted']}; background: transparent;
                border: none; border-bottom: 3px solid transparent; }}
QTabBar::tab:selected {{ color: {t['accent']}; border-bottom: 3px solid {t['accent']};
                         font-weight: 600; }}
QTabBar::tab:hover {{ color: {t['strong']}; }}

QScrollArea {{ background: transparent; border: none; }}
QListWidget {{ border: 1px solid {t['border']}; background: {t['list_bg']};
               border-radius: 8px; outline: none; padding: 4px; }}
QListWidget::item {{ color: {t['list_fg']}; padding: 10px 12px; margin: 2px; border-radius: 6px; }}
QListWidget::item:selected {{ background: {t['list_active_bg']}; color: {t['list_active_fg']};
                              font-weight: 600; }}
QListWidget::item:hover {{ background: {t['list_hover_bg']}; }}
QListWidget#plain {{ border: none; background: transparent; }}

QTableWidget {{ border: 1px solid {t['border']}; background: {t['surface']};
                gridline-color: {t['grid']}; border-radius: 6px;
                alternate-background-color: {t['surface_alt']};
                selection-background-color: {t['sel_bg']}; selection-color: {t['sel_fg']}; }}
QTableWidget::item {{ padding: 6px 8px; }}
QHeaderView::section {{ background: {t['header_bg']}; color: {t['header_fg']};
                        padding: 9px 8px; border: none; border-right: 1px solid {t['grid']};
                        font-weight: 600; }}
QTableView {{ background: {t['surface']}; border: 1px solid {t['border']}; }}

QCheckBox, QRadioButton {{ background: transparent; spacing: 8px; color: {t['text']}; }}
QCheckBox::indicator, QRadioButton::indicator {{ width: 16px; height: 16px;
    border: 1px solid {t['input_border']}; background: {t['input_bg']}; border-radius: 4px; }}
QRadioButton::indicator {{ border-radius: 9px; }}
QCheckBox::indicator:hover, QRadioButton::indicator:hover {{ border-color: {t['accent']}; }}
QCheckBox::indicator:checked, QRadioButton::indicator:checked {{ background: {t['accent']};
    border-color: {t['accent']}; }}
QGroupBox {{ border: 1px solid {t['border']}; border-radius: 8px; margin-top: 10px;
             padding: 12px; background: {t['surface']}; }}
QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 5px; color: {t['text_muted']}; }}

QScrollBar:vertical {{ width: 8px; background: {t['scroll_bg']}; border: none; }}
QScrollBar::handle:vertical {{ background: {t['scroll_fg']}; border-radius: 4px; min-height: 28px; }}
QScrollBar:horizontal {{ height: 8px; background: {t['scroll_bg']}; border: none; }}
QScrollBar::handle:horizontal {{ background: {t['scroll_fg']}; border-radius: 4px; min-width: 28px; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}

QSplitter::handle {{ background: {t['border']}; }}
QMenu {{ background: {t['menu_bg']}; border: 1px solid {t['border']}; border-radius: 6px;
         padding: 4px; }}
QMenu::item {{ padding: 7px 22px; border-radius: 4px; }}
QMenu::item:selected {{ background: {t['menu_sel']}; color: {t['accent']}; }}
QProgressBar {{ border: none; background: {t['border_soft']}; border-radius: 4px; height: 6px; }}
QProgressBar::chunk {{ background: {t['accent']}; border-radius: 4px; }}

QFrame#taskbar {{ background: {t['surface_alt']}; border: 1px solid {t['border']};
                  border-radius: 7px; }}
QLabel#tasktext {{ font-size: 12px; color: {t['text_muted']}; background: transparent; }}
QFrame#toast {{ background: {t['tooltip_bg']}; border: 1px solid {t['border']};
                border-left: 3px solid {t['accent']}; border-radius: 8px; }}
QFrame#toast[kind="ok"] {{ border-left: 3px solid {t['chip_fg']}; }}
QFrame#toast[kind="bad"] {{ border-left: 3px solid {t['danger']}; }}
QLabel#toasttext {{ font-size: 12px; color: {t['tooltip_fg']}; background: transparent; }}
QFrame#toast QPushButton#mini {{ color: {t['tooltip_fg']}; }}
QFrame#toast QPushButton#mini:hover {{ background: {t['surface']}; color: {t['text']}; }}
QFrame#cardhover:focus {{ border: 1px solid {t['accent']}; }}
QWidget[dropping="true"] {{ background: {t['accent_soft']}; }}
'''


THEME_KEY = 'ui_theme.json'


def load_theme_name(path):
    try:
        name = str(read_json(Path(path)).get('theme') or '').strip().lower()
    except Exception:
        name = ''
    return name if name in THEMES else 'light'


def save_theme_name(path, name):
    try:
        write_json(Path(path), {'theme': name if name in THEMES else 'light'})
    except OSError:
        pass


def apply_theme(app, name):
    name = name if name in THEMES else 'light'
    app.setStyleSheet(build_style(THEMES[name]))
    uiicons.set_theme(THEMES[name])
    return name

MULTI_HINT = re.compile(r'(内容|报告|分工|评语|分析|方案|总结|说明|设计|实施|保障|计划|心得|体会)')
DEFAULT_GROUP = '填写内容'


def is_multiline_label(text):
    # 表单为了对齐常在字间加空格（评    语）；字段 id 还可能转成下划线，需要先去掉再判断。
    return bool(MULTI_HINT.search(re.sub(r'[\s_]+', '', str(text or ''))))


def format_size_name(pt):
    if not pt:
        return '跟随模板'
    for name, value in CN_SIZES:
        if abs(float(pt) - value) < 0.01:
            return f'{name} {value:g}pt'
    return f'{float(pt):g}pt'


ALIGN_TEXT = {'center': '居中', 'justify': '两端对齐', 'right': '右对齐', 'left': '左对齐'}


def describe_rules(rules):
    """把 reqparse 解析出的格式规则翻成一句人话，用于左侧规则面板。"""
    rules = rules or {}
    bits = []
    if rules.get('font_east'):
        bits.append('中文 ' + str(rules['font_east']))
    if rules.get('font_ascii'):
        bits.append('西文 ' + str(rules['font_ascii']))
    if rules.get('size_pt'):
        bits.append('字号 ' + format_size_name(rules['size_pt']))
    if 'bold' in rules:
        bits.append('加粗' if rules['bold'] else '不加粗')
    if rules.get('italic'):
        bits.append('斜体')
    line = rules.get('line')
    if line:
        kind, value = line.get('type'), line.get('value')
        bits.append(f'行距 {value:g} 倍' if kind == 'multiple' else f'行距 {value:g} 磅')
    if rules.get('align'):
        bits.append(ALIGN_TEXT.get(rules['align'], rules['align']))
    if rules.get('valign'):
        bits.append('单元格垂直居中')
    if rules.get('first_line_chars'):
        bits.append(f'首行缩进 {rules["first_line_chars"]:g} 字符')
    if rules.get('hanging_chars'):
        bits.append(f'悬挂缩进 {rules["hanging_chars"]:g} 字符')
    if rules.get('blank_lines'):
        bits.append(f'空 {rules["blank_lines"]:g} 行')
    return ' · '.join(bits)


def describe_content_rule(rule, labels=None):
    labels = labels or {}
    name = labels.get(rule.get('field_key')) or rule.get('label') or '未指定栏目'
    content = rule.get('content') or {}
    bits = []
    if content.get('min_chars') and content.get('max_chars'):
        bits.append(f'字数 {content["min_chars"]}–{content["max_chars"]} 字')
    elif content.get('min_chars'):
        bits.append(f'不少于 {content["min_chars"]} 字')
    elif content.get('max_chars'):
        bits.append(f'不超过 {content["max_chars"]} 字')
    if content.get('min_refs'):
        bits.append(f'参考文献不少于 {content["min_refs"]} 条')
    if content.get('min_records'):
        bits.append(f'不少于 {content["min_records"]} 条记录')
    if content.get('score_min'):
        bits.append(f'分数不低于 {content["score_min"]} 分')
    if content.get('max_years'):
        bits.append(f'仅限 {content["max_years"]} 年以内')
    if content.get('ending'):
        bits.append(f'以“{content["ending"]}”结尾')
    if content.get('must_handwrite'):
        bits.append('需手写/手签')
    if content.get('keep_together'):
        bits.append('不得跨页')
    if content.get('title_case'):
        bits.append('英文实词首字母大写')
    text = ' · '.join(bits) or str(rule.get('text') or '').strip()
    return f'{name}：{text}'


def as_bool(value):
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in ('true', 'yes', '1', '是', '必填', '必选')
    return bool(value)


def coerce_pt(value):
    if isinstance(value, (int, float)):
        size = float(value)
    else:
        text = str(value or '').strip().lower().replace('pt', '').replace('磅', '').strip()
        for name, pt in CN_SIZES:
            if text == name:
                return float(pt)
        match = re.search(r'\d+(?:\.\d+)?', text)
        if not match:
            return None
        size = float(match.group())
    return round(size, 1) if 8 <= size <= 42 else None


def render_format(raw, config):
    """把界面保存的格式偏好转换为 engine 认识的 format 结构。"""
    raw = raw or {}
    out = {}
    for key in ('default_size_pt', 'field_size_pt', 'field_bold', 'field_font'):
        value = raw.get(key)
        if value:
            out[key] = copy.deepcopy(value)
    body_size = raw.get('body_size_pt')
    if body_size:
        if raw.get('apply_to_all'):
            out['default_size_pt'] = body_size
        else:
            sizes = dict(out.get('field_size_pt') or {})
            for field in config.get('fields', []):
                if field.get('type') == 'multiline':
                    sizes[field['key']] = body_size
            if sizes:
                out['field_size_pt'] = sizes
    return out or None


def format_summary(raw):
    raw = raw or {}
    bits = []
    body = raw.get('body_size_pt')
    bits.append('正文 ' + format_size_name(body) if body else '正文跟随模板')
    if body and raw.get('apply_to_all'):
        bits.append('全部填写内容')
    if raw.get('default_size_pt'):
        bits.append('默认 ' + format_size_name(raw.get('default_size_pt')))
    bold = raw.get('field_bold') or []
    if bold:
        bits.append(f'{len(bold)} 项加粗')
    fonts = raw.get('field_font') or {}
    if fonts:
        bits.append(f'{len(fonts)} 项字体')
    sizes = raw.get('field_size_pt') or {}
    if sizes:
        bits.append(f'{len(sizes)} 项单独字号')
    return ' · '.join(bits)


class AIWorker(QThread):
    done = Signal(object)
    failed = Signal(str)
    status = Signal(str)

    def __init__(self, fn, parent=None, with_status=False):
        super().__init__(parent)
        self.fn = fn
        self.with_status = with_status

    def run(self):
        try:
            result = self.fn(self.status.emit) if self.with_status else self.fn()
            self.done.emit(result)
        except Exception as exc:
            self.failed.emit(str(exc))


def label(text, role=None):
    w = QLabel(text)
    w.setWordWrap(True)
    if role:
        w.setObjectName(role)
    return w


def button(text, callback=None, primary=False, danger=False):
    w = QPushButton(text)
    if primary:
        w.setObjectName('primary')
    if danger:
        w.setObjectName('danger')
    if callback:
        w.clicked.connect(callback)
    return w


def icon_button(name, tip, callback=None, role='muted', px=18, side=30):
    """小按钮一律用图标：文字只留在悬浮提示里，横条上不再被它们挤。"""
    w = QPushButton()
    w.setObjectName('tool')
    uiicons.apply(w, name, tip, role, px, side)
    if callback:
        w.clicked.connect(callback)
    return w


def open_in_explorer(path):
    QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))


def mark_over(widget, over):
    """超字数提示：用主题色而不是硬编码红色，切主题后仍然看得清。"""
    wanted = 'true' if over else 'false'
    if widget.property('over') == wanted:
        return
    widget.setProperty('over', wanted)
    widget.style().unpolish(widget)
    widget.style().polish(widget)


def docx_paths(mime):
    """拖放里只认本机的 .docx；.doc 得先在 Word 里另存，不然 python-docx 读不了。"""
    if not mime.hasUrls():
        return []
    return [u.toLocalFile() for u in mime.urls()
            if u.isLocalFile() and u.toLocalFile().lower().endswith('.docx')]


def mark_dropping(widget, on):
    """拖入文档时整页轻微高亮，松手或拖出立刻恢复，让人知道放手会发生什么。"""
    wanted = 'true' if on else 'false'
    if widget.property('dropping') == wanted:
        return
    widget.setProperty('dropping', wanted)
    widget.style().unpolish(widget)
    widget.style().polish(widget)


class DropToOpen:
    """几个页面共用 .docx 拖入口：高亮反馈、只认 .docx、拖错的即时提示都保持一致。"""

    def init_drops(self, on_drop):
        self.setAcceptDrops(True)
        self._drop_target = on_drop

    def dragEnterEvent(self, event):
        if docx_paths(event.mimeData()):
            mark_dropping(self, True)
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if docx_paths(event.mimeData()):
            event.acceptProposedAction()

    def dragLeaveEvent(self, event):
        mark_dropping(self, False)

    def dropEvent(self, event):
        mark_dropping(self, False)
        paths = docx_paths(event.mimeData())
        if not paths:
            if event.mimeData().hasUrls():
                self.workspace.notify('只能拖 .docx 文件；.doc 请先在 Word 里另存为 .docx',
                                      'bad', 3200)
            event.ignore()
            return
        event.acceptProposedAction()
        self._drop_target(paths[0])


def describe_document(path):
    """失败时如实说明文档里有什么，让用户知道下一步该往哪走。"""
    try:
        from docx import Document as _Document
        doc = _Document(str(path))
    except Exception:
        return ''
    body = doc.element.body
    from docx.oxml.ns import qn
    parts = []
    parts.append(f'{len(doc.tables)} 张表'
                 + (f'（第一张 {len(doc.tables[0].rows)} 行）' if doc.tables else ''))
    parts.append(f'{len(doc.paragraphs)} 个正文段落')
    for label, tag in (('内容控件', 'w:sdt'), ('表单域', 'w:fldChar'),
                       ('旧版花括号占位符', None)):
        if tag is None:
            count = sum(1 for paragraph in doc.paragraphs if '{{' in paragraph.text)
            count += sum(1 for table in doc.tables for row in table.rows
                         for cell in row.cells if '{{' in cell.text)
        else:
            count = len(body.findall('.//' + qn(tag)))
        if count:
            parts.append(f'{label} {count} 处')
    return '、'.join(parts)


def error_advice(exc, path=None):
    """把异常翻成“下一步怎么办”，并列出工具此刻在这份文档里看到的东西。"""
    text = str(exc)
    lines = []
    lowered = text.lower()
    if 'ip_blocked' in lowered or '被封禁' in text or 'autoban' in lowered:
        lines.append('下一步：这个服务把你的出口 IP 临时封了（常见于短时间内请求太密）。'
                     '先停几分钟到一小时再试；仍不行就按它给的申诉途径处理，'
                     '或改用服务商的官方地址。')
    elif 'no access to model' in lowered or '没有权限访问模型' in text:
        lines.append('下一步：Key 本身被接受了，但这把 token 不允许使用你填的这个模型名。'
                     '到服务商后台看这把 token 勾了哪些模型，把模型名改成其中一个；'
                     '想连 DeepSeek 官方就用 platform.deepseek.com 生成的 Key，'
                     '协议选 OpenAI 兼容、地址 https://api.deepseek.com、模型 deepseek-chat。')
    elif '1010' in text or 'cloudflare' in lowered or 'just a moment' in lowered:
        lines.append('下一步：这个服务在 Cloudflare 后面按浏览器指纹拦请求。'
                     '程序已经带浏览器标识并换第二种客户端重试过；仍被拦就换官方地址'
                     '或本机代理，再不行就换一家服务。')
    elif 'authentication' in lowered or ('invalid' in lowered and 'api key' in lowered):
        lines.append('下一步：Key 没被接受。确认这把 Key 是这家服务商的（中转站的 Key '
                     '不能拿去官方站用，反之也一样），开头结尾有没有多余空格，'
                     '以及地址里的 /v1 有没有多写或漏写。')
    elif 'permission denied' in lowered or 'being used' in lowered \
            or 'access is denied' in lowered or '拒绝' in text or '占用' in text:
        lines.append('下一步：这个文件可能正被 Word 打开或没有写入权限——先关掉 Word，'
                     '或换一个保存位置再试一次。')
    elif 'zip' in lowered or '损坏' in text or 'not a zip' in lowered:
        lines.append('下一步：这份 .docx 可能已损坏，或是把旧 .doc 直接改了扩展名。'
                     '在 Word 里“另存为 .docx”后再拖进来。')
    elif '没有识别到可填写栏目' in text or '规范/说明材料' in text:
        lines.append('下一步：这类文档通常确实没得填。若只是想统一字体行距，'
                     '去「自动排版」页按规则重排整篇；若里面其实有空格要填，'
                     '可以在「模板制作」里勾“把空白格都当填写位”再识别一次。')
    elif '模板字段与配置不一致' in text:
        lines.append('下一步：这份模板是旧版录入的。到「模板制作」重新识别并录入一次，'
                     '或者在填写页直接拖原文档“打开 Word 直接填写”。')
    elif 'Word' in text and ('预览' in text or 'pdf' in lowered):
        lines.append('下一步：整页预览需要本机 Word/WPS。没装也能应用和导出，只是看不到效果图。')
    if path is not None:
        facts = describe_document(path)
        if facts:
            lines.append('这份文档里现在能看到：' + facts + '。')
    return '\n\n'.join(lines)


def backfill_requirements(store, config):
    """老模板没存过“模板自己写明的要求”：按模板文件补一次离线解析，并写回 config.json。"""
    current = config.get('req') or {}
    has_rules = bool(current.get('notes') or current.get('field_rules')
                     or current.get('container_rules'))
    if has_rules and current.get('version') == reqparse.REQ_VERSION:
        return None
    try:
        path = Path(config['_dir']) / config['file']
    except (KeyError, TypeError):
        return None
    ranges = current.get('field_ranges') or {}
    plan = {'fields': [{'key': f['key'], 'label': f.get('label'), 'type': f.get('type'),
                        'bookmark_range': ranges.get(f['key'])}
                       for f in config.get('fields', [])]}
    try:
        spec = reqparse.build(str(path), plan)
    except Exception:
        return None
    req = reqparse.to_config(spec, plan)
    if not (req.get('notes') or req.get('container_rules') or req.get('field_rules')):
        return None
    config['req'] = req
    config['req_summary'] = reqparse.summarize(spec)
    if config.get('_config_path'):
        data = {k: v for k, v in config.items() if not str(k).startswith('_')}
        try:
            write_json(Path(config['_config_path']), data)
        except OSError:
            pass
    return req


def _selected_grids(plan, fields):
    """只保留勾选到的表格格：整行都空就删行，列数按最宽的一行收缩。"""
    grids = (plan or {}).get('grids') or []
    if not grids:
        return []
    chosen = {field.get('key') for field in fields}
    result = []
    for grid in grids:
        rows = [[key for key in row if key in chosen] for row in grid.get('rows') or []]
        rows = [row for row in rows if any(row)]
        if not rows:
            continue
        width = max(len(row) for row in rows)
        rows = [row + [''] * (width - len(row)) for row in rows]
        columns = (grid.get('columns') or [])[:width] or ['列'] * width
        result.append({'group': grid.get('group') or '表格',
                       'title': grid.get('title') or '表格',
                       'columns': columns, 'rows': rows})
    return result


def upgrade_field_keys(config, store=None):
    """老模板配置迁移：把“学生姓名2 / 专业3 / 学生”这类并回同一个栏目，书签与模板文件都不动。

    只有真的多处共用一个值时才改 key；改到的 key 会同步迁移 req 规则和本机草稿，
    不会出现“改了名、草稿没了”。
    """
    fields = config.get('fields') or []
    if not fields:
        return False
    before = [field.get('key') for field in fields]
    trial = copy.deepcopy(fields)
    if not smartparse.group_by_label(trial):
        return False
    # 带 {{记号}} 的老模板：键名必须和文档里的记号一一对上，改了就生成不了，跳过
    try:
        tokens = set(scan_fields(Path(config['_dir']) / config['file']))
    except (KeyError, OSError, ValueError):
        tokens = set()
    if tokens and tokens - {field.get('key') for field in trial}:
        return False
    for field, updated in zip(fields, trial):
        field['key'] = updated.get('key')
        field['label'] = updated.get('label')
        field['type'] = updated.get('type')
        field['slots'] = updated.get('slots')
    remap = {old: new for old, new in zip(before, [f.get('key') for f in fields])
             if old and new and old != new}
    if not remap:
        return True
    req = config.get('req') or {}
    for name in ('field_rules', 'field_ranges', 'hints'):
        table = req.get(name) or {}
        if not table:
            continue
        merged = {}
        for key, value in table.items():
            target = remap.get(key, key)
            if target in merged and isinstance(value, dict) and isinstance(merged[target], dict):
                merged[target].update(value)
            else:
                merged[target] = value
        req[name] = merged
    for item in (req.get('notes') or []) + (req.get('content_rules') or []):
        if item.get('field_key') in remap:
            item['field_key'] = remap[item['field_key']]
    if req:
        config['req'] = req
    if store is not None:
        draft = store.root / 'drafts' / (str(config.get('id')) + '.json')
        if draft.is_file():
            try:
                data = read_json(draft)
                values = data.get('values') or {}
                moved = {}
                for key, value in values.items():
                    moved.setdefault(remap.get(key, key), value)
                data['values'] = moved
                write_json(draft, data)
            except (OSError, ValueError):
                pass
    if str(config.get('id', '')).startswith('builtin-'):
        return True          # 内置模板只在本机内存里并键，不写回程序目录
    path = config.get('_config_path')
    if path:
        data = {k: v for k, v in config.items() if not str(k).startswith('_')}
        try:
            write_json(Path(path), data)
        except OSError:
            pass
    return True


def mark_selected(widget, selected):
    wanted = 'true' if selected else 'false'
    if widget.property('selected') == wanted:
        return
    widget.setProperty('selected', wanted)
    widget.style().unpolish(widget)
    widget.style().polish(widget)


class Store:
    def __init__(self, root):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)
        self.custom = self.root / 'templates'
        self.custom.mkdir(exist_ok=True)
        self.formats = self.root / 'formats'
        self.formats.mkdir(exist_ok=True)
        self.ai_config = self.root / 'ai_config.json'
        self.warnings = []

    def load_format(self, template_id):
        try:
            data = read_json(self.formats / (template_id + '.json'))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def save_format(self, template_id, fmt):
        write_json(self.formats / (template_id + '.json'), fmt or {})

    def drop_format(self, template_id):
        path = self.formats / (template_id + '.json')
        if path.is_file():
            path.unlink()

    def templates(self):
        configs = []
        for path in sorted(ASSETS.glob('builtin-*.json')) + sorted(self.custom.glob('*/config.json')):
            try:
                config = read_json(path)
                config['_dir'] = str(path.parent)
                config['_config_path'] = str(path)
                if not (path.parent / config['file']).is_file():
                    raise ValueError('模板文件不存在')
                configs.append(config)
            except (OSError, ValueError, KeyError) as exc:
                self.warnings.append(f'{path}：{exc}')
        return configs

    def load(self, config):
        values = {f['key']: f.get('default', '') for f in config['fields']}
        values.update({f['key']: '__original__' for f in config.get('images', [])})
        path = self.root / 'drafts' / (config['id'] + '.json')
        if path.exists():
            try:
                values.update(read_json(path)['values'])
            except (OSError, ValueError, KeyError) as exc:
                self.warnings.append(f'草稿无法读取，已保留原文件：{exc}')
        return values

    def save(self, config, values):
        write_json(self.root / 'drafts' / (config['id'] + '.json'),
                   {'template': config['id'], 'saved_at': datetime.now().isoformat(), 'values': values})


# ============================================================ ① 填写生成页
class FillPage(DropToOpen, QWidget):
    def __init__(self, store, workspace):
        super().__init__()
        self.store, self.workspace = store, workspace
        self.configs = []
        self.current = None
        self.editors, self.grids, self.images = {}, [], {}
        self.dirty = False
        self.building = False
        self.last_output = None
        self.preview_page = 0
        self.zoom = 1.0
        self.combo_map = []
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(700)
        self.timer.timeout.connect(self.save_draft)
        self.live_pdf = b''
        self.live_page_count = 0
        self.live_page_index = -1
        self.live_pixmap = QPixmap()
        self.preview_token = 0
        self.baseline_token = 0
        self.page_baseline = None
        self._baseline_cache = {}       # (模板路径, mtime) -> 页数，切回同一模板不再重渲
        self._overflow_max = 0
        self.live_available = True
        self.direct_config = None
        self._direct_worker = None
        self.init_drops(self.start_direct)
        self.preview_worker = preview.PreviewWorker(store.root / 'preview', self)
        self.preview_worker.rendered.connect(self.on_live_rendered)
        self.preview_worker.failed.connect(self.on_live_failed)
        self.preview_worker.status.connect(self.on_live_status)
        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(800)
        self.preview_timer.timeout.connect(self.request_live_preview)
        self.build_ui()
        self.reload()

    def build_ui(self):
        shell = QVBoxLayout(self)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)
        shell.addWidget(self.build_toolbar())
        main = QVBoxLayout()
        main.setContentsMargins(20, 14, 20, 10)
        main.setSpacing(10)
        shell.addLayout(main, 1)
        heading = QHBoxLayout()
        heading.setSpacing(8)
        self.title = label('', 'title')
        heading.addWidget(self.title, 1)
        self.field_count = label('', 'badge')
        heading.addWidget(self.field_count)
        main.addLayout(heading)
        self.description = label('', 'muted')
        self.direct_status = label('把 .docx 拖到本页面即可识别填写位', 'chip')
        self.direct_status.setWordWrap(False)
        self.direct_status.setMaximumWidth(430)
        meta = QHBoxLayout()
        meta.setSpacing(10)
        meta.addWidget(self.description, 1)
        meta.addWidget(self.direct_status, 0, Qt.AlignmentFlag.AlignRight)
        main.addLayout(meta)
        self.list = QListWidget()
        self.list.hide()
        self.list.currentRowChanged.connect(self.select_template)
        self.empty_search = label('没有匹配的模板', 'muted')
        self.empty_search.hide()
        main.addWidget(self.empty_search)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.addWidget(self.build_rules_panel())
        self.preview_panel = self.build_preview_panel()
        self.splitter.addWidget(self.preview_panel)
        self.splitter.addWidget(self.build_form_panel())
        self.splitter.setSizes([292, 600, 470])
        self.splitter.splitterMoved.connect(lambda *_: self.render_preview())
        main.addWidget(self.splitter, 1)
        footer = QFrame()
        footer.setObjectName('footer')
        foot = QHBoxLayout(footer)
        foot.setContentsMargins(15, 12, 15, 12)
        st = QVBoxLayout()
        self.progress = label('')
        self.save_status = label('填写后自动保存草稿', 'muted')
        self.format_status = label('', 'muted')
        st.addWidget(self.progress)
        st.addWidget(self.save_status)
        st.addWidget(self.format_status)
        foot.addLayout(st, 1)
        self.save_template_button = button('存入模板库', self.save_direct_template)
        self.save_template_button.hide()
        foot.addWidget(self.save_template_button)
        foot.addWidget(button('保存草稿', self.save_draft))
        self.open_button = button('打开生成文件', self.open_output)
        self.open_button.setEnabled(False)
        foot.addWidget(self.open_button)
        main.addWidget(footer)

    def build_toolbar(self):
        bar = QFrame()
        bar.setObjectName('panel')
        bar.setFixedHeight(54)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(20, 9, 20, 9)
        lay.setSpacing(10)
        self.combo = QComboBox()
        self.combo.setMinimumWidth(250)
        self.combo.setToolTip('切换要填写的模板；也可以直接在下面搜索。')
        self.combo.currentIndexChanged.connect(self.combo_changed)
        lay.addWidget(self.combo)
        self.search = QLineEdit()
        self.search.setObjectName('search')
        self.search.setPlaceholderText('筛选模板名称')
        self.search.setClearButtonEnabled(True)
        self.search.setFixedWidth(190)
        self.search.textChanged.connect(self.filter_templates)
        lay.addWidget(self.search)
        self.open_doc_btn = button('打开 Word 直接填写', self.open_document_dialog)
        self.open_doc_btn.setToolTip('选一个 .docx，自动识别填写位后直接填，不需要先做模板。')
        lay.addWidget(self.open_doc_btn)
        self.layout_btn = button('整篇排版', self.send_to_layout)
        self.layout_btn.setToolTip('把当前模板带到“自动排版”页，按它自己的要求重排整篇文档。')
        lay.addWidget(self.layout_btn)
        self._req_worker = None
        lay.addStretch(1)
        self.preview_toggle = QCheckBox('显示预览')
        self.preview_toggle.setChecked(True)
        self.preview_toggle.toggled.connect(self.toggle_preview)
        lay.addWidget(self.preview_toggle)
        self.generate_button = button('生成 Word', self.generate, primary=True)
        lay.addWidget(self.generate_button)
        self.batch_btn = button('批量填写', self.open_batch)
        self.batch_btn.setToolTip('一份名单生成多份 Word：CSV / TXT / JSON / Excel 都行。')
        lay.addWidget(self.batch_btn)
        return bar

    def open_batch(self):
        BatchDialog(self.workspace, self.current, self).exec()

    def build_rules_panel(self):
        panel = QFrame()
        panel.setObjectName('panel')
        panel.setMinimumWidth(252)
        box = QVBoxLayout(panel)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(0)
        box.addWidget(label('排版规则 · 模板要求', 'paneltitle'))
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        host.setObjectName('formpage')
        self.rules_layout = QVBoxLayout(host)
        self.rules_layout.setContentsMargins(14, 12, 14, 12)
        self.rules_layout.setSpacing(7)
        scroll.setWidget(host)
        box.addWidget(scroll, 1)
        tail = QVBoxLayout()
        tail.setContentsMargins(12, 8, 12, 10)
        tail.setSpacing(6)
        box.addLayout(tail)
        self.show_rules()
        return panel

    def build_preview_panel(self):
        panel = QFrame()
        panel.setObjectName('panel')
        panel.setMinimumWidth(300)
        box = QVBoxLayout(panel)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(0)
        head = QHBoxLayout()
        head.setContentsMargins(14, 10, 12, 4)
        head.setSpacing(8)
        head.addWidget(label('文档预览', 'section'))
        self.live_toggle = QCheckBox('实时渲染')
        self.live_toggle.setChecked(True)
        self.live_toggle.setToolTip('使用本机 Word/WPS 按当前填写内容渲染预览')
        self.live_toggle.toggled.connect(self.toggle_live)
        head.addWidget(self.live_toggle)
        self.retry_preview_btn = button('重试', lambda: self.schedule_live_preview(0))
        self.retry_preview_btn.setObjectName('mini')
        self.retry_preview_btn.hide()
        head.addWidget(self.retry_preview_btn)
        head.addStretch(1)
        box.addLayout(head)
        self.preview_status = label('准备预览…', 'muted')
        self.preview_status.setWordWrap(True)
        status = QHBoxLayout()
        status.setContentsMargins(14, 0, 14, 6)
        status.addWidget(self.preview_status, 1)
        box.addLayout(status)
        area = QFrame()
        area.setObjectName('previewarea')
        al = QVBoxLayout(area)
        al.setContentsMargins(12, 12, 12, 12)
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(True)
        self.preview_image = QLabel()
        self.preview_image.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.preview_image.setMinimumWidth(1)
        self.preview_image.setWordWrap(True)
        self.preview_scroll.setWidget(self.preview_image)
        al.addWidget(self.preview_scroll)
        box.addWidget(area, 1)
        bar = QHBoxLayout()
        bar.setContentsMargins(14, 6, 14, 10)
        bar.setSpacing(6)
        self.fit_width = QCheckBox('适应宽度')
        self.fit_width.setChecked(True)
        self.fit_width.toggled.connect(self.zoom_mode_changed)
        bar.addWidget(self.fit_width)
        bar.addStretch(1)
        self.prev = icon_button('up', '上一页', lambda: self.move_page(-1))
        self.next = icon_button('down', '下一页', lambda: self.move_page(1))
        self.page_number = label('')
        self.page_number.setAlignment(Qt.AlignmentFlag.AlignCenter)
        bar.addWidget(self.prev)
        bar.addWidget(self.page_number)
        bar.addWidget(self.next)
        bar.addSpacing(8)
        zoom_out = icon_button('zoom-out', '缩小预览', self.zoom_out)
        self.zoom_label = label('适应宽度', 'muted')
        self.zoom_label.setMinimumWidth(46)
        self.zoom_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        zoom_in = icon_button('zoom-in', '放大预览', self.zoom_in)
        bar.addWidget(zoom_out)
        bar.addWidget(self.zoom_label)
        bar.addWidget(zoom_in)
        box.addLayout(bar)
        return panel

    def build_form_panel(self):
        panel = QFrame()
        panel.setObjectName('panel')
        panel.setMinimumWidth(340)
        box = QVBoxLayout(panel)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(0)
        box.addWidget(label('填写内容', 'paneltitle'))
        inner = QWidget()
        inner.setObjectName('formpage')
        il = QVBoxLayout(inner)
        il.setContentsMargins(12, 8, 12, 12)
        self.tabs = QTabWidget()
        self.tabs.setUsesScrollButtons(True)
        il.addWidget(self.tabs)
        box.addWidget(inner, 1)
        return panel

    def clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
                widget.deleteLater()

    def show_rules(self):
        layout = self.rules_layout
        self.clear_layout(layout)
        config = self.current
        if not config:
            layout.addWidget(label('还没有选择模板：选一个模板，或点“打开 Word 直接填写”。', 'muted'))
            layout.addStretch(1)
            return
        req = config.get('req') or {}
        labels = {f['key']: (f.get('label') or f['key']) for f in config['fields']}
        field_rules = req.get('field_rules') or {}
        container_rules = req.get('container_rules') or []
        notes = req.get('notes') or []
        content_rules = req.get('content_rules') or []
        layout.addWidget(label('本次生成会照做的规则', 'section'))
        if not (field_rules or container_rules or notes or content_rules):
            layout.addWidget(label('这个模板没有读到格式要求（没有彩色/底纹说明，也没有行距、缩进这类句子）。'
                                   '生成时只填内容，不改版式。', 'muted'))
        else:
            chips = QHBoxLayout()
            chips.setSpacing(6)
            for text in (f'说明 {len(notes)} 条', f'格式规则 {len(field_rules) + len(container_rules)} 条',
                         f'内容要求 {len(content_rules)} 条'):
                chips.addWidget(label(text, 'chip'))
            chips.addStretch()
            layout.addLayout(chips)
            for key, rules in list(field_rules.items())[:10]:
                text = describe_rules(rules)
                if text:
                    layout.addWidget(label('· ' + labels.get(key, key) + '：' + text, 'muted'))
            def where(entry):
                container = entry.get('container') or 'body'
                if container == 'cell':
                    return ('表%d 行%d 列%d' % (int(entry.get('table') or 0) + 1,
                                              int(entry.get('row') or 0) + 1,
                                              int(entry.get('col') or 0) + 1))
                if container == 'header':
                    return '页眉'
                if container == 'footer':
                    return '页脚'
                return '正文第%d段' % (int(entry.get('p') or 0) + 1)

            grouped = {}
            for entry in container_rules:
                text = describe_rules(entry.get('rules') or {})
                if text:
                    grouped.setdefault(text, []).append(where(entry))
            for text, places in list(grouped.items())[:8]:
                extra = f'（等 {len(places)} 处）' if len(places) > 1 else ''
                layout.addWidget(label(f'· {places[0]}：{text}{extra}', 'muted'))
        hints = req.get('hints') or {}
        general = req.get('general_hints') or []
        if hints or general:
            layout.addWidget(label('从说明里搬来的填写提示', 'section'))
            layout.addWidget(label('原来写在模板里的彩色/底纹说明会被删掉，内容搬到这里和每个填写栏下方。', 'muted'))
            for item in general[:8]:
                layout.addWidget(label('· ' + str(item.get('text') or '') +
                                       (f"（{item['where']}）" if item.get('where') else ''), 'muted'))
            for key, texts in list(hints.items())[:8]:
                for text in (texts or [])[:2]:
                    layout.addWidget(label('· ' + labels.get(key, key) + '：' + str(text), 'muted'))
        if content_rules:
            layout.addWidget(label('内容要求（只提醒，不拦截）', 'section'))
            for rule in content_rules[:8]:
                layout.addWidget(label('· ' + describe_content_rule(rule, labels), 'muted'))
        layout.addStretch(1)

    def send_to_layout(self):
        """把当前模板交给自动排版页：源文件 + 已解析的要求一起带过去。"""
        if not self.current:
            return
        self.save_draft()
        self.workspace.layout_page.load_template(self.current)
        self.workspace.nav.setCurrentRow(Workspace.LAYOUT)

    def backfill_requirements_async(self, config):
        """老模板补读要求：读到后刷新左栏规则并重渲染预览（说明会被删掉）。"""
        if self._req_worker is not None and self._req_worker.isRunning():
            return
        target_id = config.get('id')
        self.save_status.setText('正在补读模板要求（离线，几秒）…')
        worker = AIWorker(lambda: backfill_requirements(self.store, config), self)
        self._req_worker = worker

        def done(result):
            self._req_worker = None
            if not result or not self.current or self.current.get('id') != target_id:
                return
            self.current['req'] = result
            self.current['req_options'] = dict(REQ_OPTIONS)
            self.show_rules()
            self.apply_editor_sizes()
            self.save_status.setText('已补读模板要求：' + str(self.current.get('req_summary') or ''))
            self.schedule_live_preview(0)

        worker.done.connect(done)
        worker.failed.connect(lambda _message: setattr(self, '_req_worker', None))
        worker.start()

    def sync_combo(self):
        current = self.configs.index(self.current) if self.current in self.configs else -1
        self.combo.blockSignals(True)
        self.combo.clear()
        self.combo_map = []
        for index, config in enumerate(self.configs):
            if self.list.item(index) is not None and self.list.item(index).isHidden():
                continue
            star = '★ ' if config.get('_direct') else ''
            self.combo.addItem(f'{star}{config["title"]}')
            self.combo_map.append(index)
        if current in self.combo_map:
            self.combo.setCurrentIndex(self.combo_map.index(current))
        self.combo.blockSignals(False)

    def combo_changed(self, position):
        if position < 0 or position >= len(self.combo_map):
            return
        self.list.setCurrentRow(self.combo_map[position])

    def zoom_in(self):
        self.set_zoom(self.zoom * 1.25)

    def zoom_out(self):
        self.set_zoom(self.zoom / 1.25)

    def set_zoom(self, factor):
        self.zoom = min(3.0, max(0.5, factor))
        self.fit_width.setChecked(False)
        self.update_zoom_label()
        self.render_preview()

    def zoom_mode_changed(self, _checked):
        self.update_zoom_label()
        self.render_preview()

    def update_zoom_label(self):
        self.zoom_label.setText('适应宽度' if self.fit_width.isChecked()
                                else f'{int(round(self.zoom * 100))}%')

    def reload(self, keep_current=True):
        keep_id = self.current['id'] if (keep_current and self.current) else None
        self.configs = self.store.templates()
        if self.direct_config:
            self.configs.insert(0, self.direct_config)
        self.list.blockSignals(True)
        self.list.clear()
        index = 0
        for i, config in enumerate(self.configs):
            star = '★ ' if config.get('_direct') else ''
            item = QListWidgetItem(f'{i + 1:02d}   {star}{config["title"]}')
            item.setToolTip(config['title'])
            self.list.addItem(item)
            if keep_id and config['id'] == keep_id:
                index = i
        self.list.blockSignals(False)
        self.search.setPlaceholderText(f'筛选模板名称（共 {len(self.configs)} 个）')
        self.current = None
        self.list.setCurrentRow(index)
        if keep_id and not any(c['id'] == keep_id for c in self.configs):
            self.building = True
            self.tabs.clear()
            self.editors, self.grids, self.images = {}, [], {}
            self.title.setText('模板已被删除')
            self.description.setText('')
            self.current = None
            self.building = False
            self.progress.setText('')
            self.format_status.setText('')
            self.preview_token += 1
            self.preview_timer.stop()
            self.live_pdf = b''
            self.live_page_count = 0
            self.live_page_index = -1
            self.live_pixmap = QPixmap()
            self.preview_status.setText('实时预览已停止')
        self.filter_templates(self.search.text())

    def open_document_dialog(self):
        path, _ = QFileDialog.getOpenFileName(
            self, '打开 Word 直接填写', '', 'Word 文档 (*.docx)')
        if path:
            self.start_direct(path)

    def start_direct(self, path):
        if self._direct_worker is not None:
            return
        source = Path(path)
        if source.suffix.lower() != '.docx':
            QMessageBox.information(
                self, '需要 .docx 文件',
                '只支持 .docx。旧版 .doc 请先用 Word 另存为 .docx 再打开。')
            return
        if self.current and not self.save_draft():
            return
        text_cfg = self.workspace._text_cfg()
        vision_cfg = self.workspace._vision_cfg()
        out_dir = self.store.root / 'direct' / ('d' + uuid.uuid4().hex[:10])
        self.direct_status.setText(f'正在识别：{source.name}…')
        self.open_doc_btn.setEnabled(False)
        if text_cfg or vision_cfg:
            worker = AIWorker(
                lambda emit: directfill.prepare(str(source), out_dir, text_cfg=text_cfg,
                                                vision_cfg=vision_cfg, progress=emit),
                self, with_status=True)
            worker.status.connect(self.direct_status.setText)
        else:
            worker = AIWorker(lambda: directfill.prepare(str(source), out_dir), self)
        self._direct_worker = worker
        worker.done.connect(
            lambda result, p=source, d=out_dir: self.direct_ready(result, p, d))
        worker.failed.connect(
            lambda message, p=source: self.direct_failed(message, p))
        worker.finished.connect(self._direct_finished)
        worker.finished.connect(worker.deleteLater)
        worker.start()


    def direct_ready(self, result, source, out_dir):
        plan, config = result
        config['id'] = 'direct-' + uuid.uuid4().hex[:12]
        config['_dir'] = str(out_dir)
        config['_direct'] = True
        config['_source'] = str(source)
        self.direct_config = config
        text = plan.get('text') or {}
        vision = plan.get('vision') or {}
        extra = ''
        if text.get('used'):
            extra += f" · 文本模型补识别 {text.get('added', 0)} 项"
        if text.get('error'):
            extra += ' · 文本模型未完成，仅结构识别'
        if vision.get('checked'):
            extra += f' · 视觉复核 {vision.get("pages", 0)} 页'
        self.direct_status.setText(
            f'当前：{source.name} · {len(plan["fields"])} 个填写项{extra}')
        self.reload(keep_current=False)
        self.workspace.nav.setCurrentRow(Workspace.FILL)
        self.save_status.setText('文档已识别，填写后直接生成；也可点“存入模板库”长期使用')
        if text.get('error'):
            self.save_status.setText('文本模型未完成，已保留结构识别结果；可继续填写或重试')


    def direct_failed(self, message, source):
        self.direct_status.setText('识别失败，可换一份文档或到“模板制作”页检查')
        QMessageBox.warning(
            self, '没有识别到可填写项',
            f'{source.name}\n\n{message}\n\n'
            '如果这份文档确实需要填写，请把文件发给开发者补充识别规则。\n'
            '原文件没有被修改。')


    def _direct_finished(self):
        self._direct_worker = None
        self.open_doc_btn.setEnabled(True)


    def save_direct_template(self):
        if not self.current or not self.current.get('_direct'):
            return
        tid = 'custom-' + uuid.uuid4().hex[:12]
        folder = self.store.custom / tid
        try:
            folder.mkdir(parents=True, exist_ok=True)
            shutil.copy2(Path(self.current['_dir']) / self.current['file'],
                         folder / 'template.docx')
            config = {k: v for k, v in self.current.items() if not k.startswith('_')}
            config['id'] = tid
            write_json(folder / 'config.json', config)
        except OSError as exc:
            QMessageBox.warning(self, '存入模板库失败', str(exc))
            return
        self.direct_config = None
        self.workspace.refresh_all()
        self.workspace.go_fill(tid)
        QMessageBox.information(
            self, '已存入模板库',
            f'模板「{config["title"]}」已保存，共 {len(config["fields"])} 个填写项。')


    def filter_templates(self, text):
        visible = 0
        for i, config in enumerate(self.configs):
            show = text.strip().casefold() in config['title'].casefold()
            self.list.item(i).setHidden(not show)
            visible += show
        self.sync_combo()
        self.empty_search.setVisible(not visible)


    def select_template(self, index):
        if index < 0 or index >= len(self.configs):
            return
        config = self.configs[index]
        if self.current and config['id'] == self.current['id']:
            return
        if self.current and not self.save_draft():
            self.list.blockSignals(True)
            self.list.setCurrentRow(self.configs.index(self.current))
            self.list.blockSignals(False)
            self.sync_combo()
            return
        self.timer.stop()
        self.preview_timer.stop()
        self.preview_token += 1
        self.baseline_token -= 1
        self.page_baseline = None
        self._overflow_max = 0
        self.live_pdf = b''
        self.live_page_count = 0
        self.live_page_index = -1
        self.live_pixmap = QPixmap()
        self.building = True
        self.current = config
        upgrade_field_keys(config, self.store)
        values = self.store.load(config)
        self.title.setText(config['title'])
        self.description.setText(config['description'])
        slots = len(config['fields'])
        columns = len({field['key'] for field in config['fields']})
        self.field_count.setText(f"{columns} 个栏目" + (f" · 共 {slots} 处位置" if slots > columns else ''))
        self.show_rules()
        self.sync_combo()
        req_now = config.get('req') or {}
        if (req_now.get('version') != reqparse.REQ_VERSION
                or not (req_now.get('notes') or req_now.get('field_rules')
                        or req_now.get('container_rules'))):
            self.backfill_requirements_async(config)
        self.editors, self.grids, self.images = {}, [], {}
        while self.tabs.count():
            page = self.tabs.widget(0)
            self.tabs.removeTab(0)
            page.deleteLater()
        groups = list(dict.fromkeys(f['group'] for f in config['fields']))
        for g in dict.fromkeys(i.get('group', '图片') for i in config.get('images', [])):
            if g not in groups:
                groups.append(g)
        for group in groups:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            content = QWidget()
            content.setObjectName('formpage')
            layout = QVBoxLayout(content)
            layout.setContentsMargins(15, 20, 15, 15)
            layout.setSpacing(13)
            for field in config['fields']:
                if field['group'] == group and not field.get('grid'):
                    if field['key'] in self.editors:
                        continue     # 同名栏目共用一个输入框，不重复摆两个
                    self.make_editor(field, values.get(field['key'], ''), layout)
                    slots = sum(1 for other in config['fields'] if other['key'] == field['key'])
                    if slots > 1:
                        layout.addWidget(label(
                            f'这一项在模板里出现 {slots} 处，填一次会同时写到所有位置。', 'muted'))
            for grid in config.get('grids', []):
                if grid['group'] != group:
                    continue
                layout.addWidget(label(grid['title'], 'section'))
                layout.addWidget(label('双击单元格填写，使用 Tab 切换。行数与原始模板一致。', 'muted'))
                table = QTableWidget(len(grid['rows']), len(grid['columns']))
                table.setHorizontalHeaderLabels(grid['columns'])
                table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
                table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
                table.verticalHeader().setDefaultSectionSize(66)
                table.setMinimumHeight(110 + 66 * len(grid['rows']))
                for r, row in enumerate(grid['rows']):
                    for c, key in enumerate(row):
                        table.setItem(r, c, QTableWidgetItem(values.get(key, '')))
                table.itemChanged.connect(self.changed)
                layout.addWidget(table)
                self.grids.append((grid, table))
            for image in config.get('images', []):
                if image['group'] == group:
                    layout.addWidget(label(image['label'], 'section'))
                    row = QHBoxLayout()
                    state = label('')
                    self.images[image['key']] = (values.get(image['key'], '__original__'), state)
                    row.addWidget(state, 1)
                    row.addWidget(button('更换图片', lambda _, f=image: self.choose_image(f)))
                    row.addWidget(button('保留原图', lambda _, k=image['key']: self.set_image(k, '__original__')))
                    row.addWidget(button('不使用图片', lambda _, k=image['key']: self.set_image(k, '')))
                    layout.addLayout(row)
                    layout.addWidget(label('更换图片按原图区域显示，建议使用接近原图的宽高比例。', 'muted'))
                    self.update_image_label(image['key'])
            layout.addStretch()
            scroll.setWidget(content)
            self.tabs.addTab(scroll, group)
        self.building = False
        self.dirty = False
        self.last_output = None
        self.open_button.setEnabled(False)
        self.save_status.setText('填写后自动保存草稿')
        self.update_format_status()
        self.save_template_button.setVisible(bool(config.get('_direct')))
        self.preview_page = 0
        self.preview_status.setText('正在准备实时预览…' if self.live_toggle.isChecked()
                                    else '实时渲染已关闭')
        self.update_progress()
        QTimer.singleShot(0, self.capture_page_baseline)
        QTimer.singleShot(0, self.render_preview)
        self.schedule_live_preview(900)


    def make_editor(self, field, value, layout):
        head = QHBoxLayout()
        head.setSpacing(8)
        head.addWidget(label(field['label'] + ('  *' if field.get('required') else '')), 1)
        kind = field.get('type', 'text')
        if kind == 'multiline':
            ai_btn = button('AI 生成', lambda: self.generate_field(field))
            ai_btn.setObjectName('mini')
            ai_btn.setToolTip('根据模板和已填内容，用 AI 起草这段正文')
            head.addWidget(ai_btn)
        layout.addLayout(head)
        if kind == 'multiline':
            editor = QPlainTextEdit()
            editor.setMinimumHeight(155)
            editor.setMaximumHeight(230)
            editor.setPlainText(value)
            editor.setPlaceholderText('在这里填写正文，可换行…')
            counter = label('', 'muted')

            def update_count():
                count, maximum = len(editor.toPlainText()), field.get('max_length', 0)
                counter.setText(f'{count} / {maximum} 字' if maximum else f'{count} 字')
                mark_over(counter, bool(maximum and count > maximum))
            editor.textChanged.connect(update_count)
            editor.textChanged.connect(self.changed)
            update_count()
            layout.addWidget(editor)
            layout.addWidget(counter)
        elif kind == 'select':
            editor = QComboBox()
            for option in field['options']:
                editor.addItem(option['label'], option['value'])
            idx = editor.findData(value)
            if idx >= 0:
                editor.setCurrentIndex(idx)
            editor.currentIndexChanged.connect(self.changed)
            layout.addWidget(editor)
        else:
            editor = QLineEdit(value)
            editor.setPlaceholderText('请输入' + field['label'])
            editor.textChanged.connect(self.changed)
            layout.addWidget(editor)
        editor.setAccessibleName(field['label'])
        self._style_editor(editor, field)
        self.editors[field['key']] = editor
        if field.get('hint'):
            layout.addWidget(label(field['hint'], 'muted'))

    def _style_editor(self, editor, field):
        """输入框字号跟随“模板要求/手动格式”，并在悬浮提示里写清会照做的规则。"""
        config = self.current or {}
        rule = ((config.get('req') or {}).get('field_rules') or {}).get(field['key']) or {}
        applied = {}
        if config.get('id'):
            applied = render_format(self.store.load_format(config['id']), config) or {}
        manual = ((applied.get('field_size_pt') or {}).get(field['key'])
                  or applied.get('default_size_pt'))
        size = manual or rule.get('size_pt')
        if size:
            try:
                size = float(size)
            except (TypeError, ValueError):
                size = None
        if size:
            font = editor.font()
            font.setPointSizeF(size)
            editor.setFont(font)
            if isinstance(editor, QPlainTextEdit):
                editor.setMinimumHeight(min(225, max(155, int(70 + 13 * size))))
        text = describe_rules(rule)
        if text:
            editor.setToolTip('模板要求：' + text + '（生成时照此排版）')

    def apply_editor_sizes(self):
        """补读模板要求后，把已有输入框的字号重新对齐一次。"""
        for field in (self.current or {}).get('fields', []):
            editor = self.editors.get(field['key'])
            if editor is not None:
                self._style_editor(editor, field)


    def values(self):
        data = {}
        for key, editor in self.editors.items():
            if isinstance(editor, QPlainTextEdit):
                data[key] = editor.toPlainText()
            elif isinstance(editor, QComboBox):
                data[key] = editor.currentData()
            else:
                data[key] = editor.text()
        for grid, table in self.grids:
            for r, row in enumerate(grid['rows']):
                for c, key in enumerate(row):
                    item = table.item(r, c)
                    data[key] = item.text() if item else ''
        data.update({key: value[0] for key, value in self.images.items()})
        return data


    def changed(self, *_):
        if self.building:
            return
        self.dirty = True
        self.save_status.setText('正在保存…')
        self.update_progress()
        self.timer.start()
        self.schedule_live_preview()


    def update_progress(self):
        if not self.current:
            self.progress.setText('')
            return
        values = self.values()
        fields = self.current['fields']
        count = sum(bool(str(values.get(f['key'], '')).strip()) for f in fields)
        required = sum(bool(f.get('required')) for f in fields)
        self.progress.setText(f'已填写 {count} / {len(fields)} 项'
                              + (f'   ·   * 为必填项，共 {required} 项' if required else ''))


    def save_draft(self, quiet=True):
        focus = QApplication.focusWidget()
        if focus and any(table.isAncestorOf(focus) for _, table in self.grids):
            self.generate_button.setFocus()
            QApplication.processEvents()
        self.timer.stop()
        if not self.current or not self.dirty:
            return True
        try:
            self.store.save(self.current, self.values())
        except OSError as exc:
            self.save_status.setText('草稿保存失败，请检查本机存储空间')
            QMessageBox.warning(self, '草稿未保存', f'填写内容仍在窗口中，请勿关闭程序。\n{exc}')
            return False
        self.dirty = False
        self.save_status.setText('草稿已保存到本机')
        if not quiet:
            self.workspace.notify('草稿已保存到本机', 'ok', 1500)
        return True


    def open_output(self):
        if self.last_output and self.last_output.is_file():
            if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_output))):
                QMessageBox.information(self, '无法打开', '请安装 Word 或 WPS 后打开 .docx 文件。')


    def _show_preview_pixmap(self, pixmap):
        base = max(160, self.preview_scroll.viewport().width() - 16)
        width = base if self.fit_width.isChecked() else max(160, int(base * self.zoom))
        self.preview_image.setMinimumWidth(width)
        self.preview_image.setPixmap(
            pixmap.scaledToWidth(width, Qt.TransformationMode.SmoothTransformation))


    def _live_page_pixmap(self):
        if not self.live_pdf:
            return None
        if self.live_page_index != self.preview_page or self.live_pixmap.isNull():
            png = preview.render_page(self.live_pdf, self.preview_page, scale=2.0)
            pixmap = QPixmap()
            if not pixmap.loadFromData(png, 'PNG'):
                return None
            self.live_pixmap = pixmap
            self.live_page_index = self.preview_page
        return self.live_pixmap


    def render_preview(self):
        if not self.current:
            self.preview_image.setPixmap(QPixmap())
            self.preview_image.setText('')
            return
        if self.live_toggle.isChecked() and self.live_pdf:
            count = max(1, self.live_page_count)
            self.preview_page = max(0, min(self.preview_page, count - 1))
            self.prev.setEnabled(count > 1 and self.preview_page > 0)
            self.next.setEnabled(count > 1 and self.preview_page < count - 1)
            self.page_number.setText(f'{self.preview_page + 1} / {count}')
            try:
                pixmap = self._live_page_pixmap()
            except Exception as exc:
                pixmap = None
                self.preview_status.setText('实时预览页渲染失败：' + str(exc)[:80])
            if pixmap is not None:
                self._show_preview_pixmap(pixmap)
                return
        pages = self.current.get('previews', [])
        if pages:
            self.preview_page = max(0, min(self.preview_page, len(pages) - 1))
        self.prev.setEnabled(bool(pages) and self.preview_page > 0)
        self.next.setEnabled(bool(pages) and self.preview_page < len(pages) - 1)
        self.page_number.setText(f'{self.preview_page + 1} / {len(pages)}' if pages else '')
        if not pages:
            self.preview_image.setPixmap(QPixmap())
            if self.live_toggle.isChecked():
                self.preview_image.setText('实时预览准备中…\n首次启动 Word 约需几秒。')
            else:
                self.preview_image.setText('此模板没有参考预览。\n打开“实时渲染”或生成后用 Word/WPS 查看。')
            return
        pixmap = QPixmap(str(Path(self.current['_dir']) / pages[self.preview_page]))
        self._show_preview_pixmap(pixmap)


    def move_page(self, offset):
        live = self.live_toggle.isChecked() and bool(self.live_pdf)
        total = self.live_page_count if live else len(self.current.get('previews', [])) \
            if self.current else 0
        if total <= 0:
            return
        self.preview_page = max(0, min(self.preview_page + offset, total - 1))
        self.render_preview()


    def toggle_preview(self, visible):
        self.preview_panel.setVisible(visible)
        if visible:
            if self.page_baseline is None:
                QTimer.singleShot(0, self.capture_page_baseline)
            QTimer.singleShot(0, self.render_preview)
            self.schedule_live_preview(0)
        else:
            self.preview_timer.stop()


    def toggle_live(self, checked):
        if checked:
            self.retry_preview_btn.hide()
            self.preview_status.setText('准备实时渲染…')
            if self.page_baseline is None:
                QTimer.singleShot(0, self.capture_page_baseline)
            self.schedule_live_preview(0)
            self.render_preview()
        else:
            self.preview_timer.stop()
            self.preview_status.setText('实时渲染已关闭')
            self.render_preview()


    def schedule_live_preview(self, delay=800):
        if not self.current or not self.live_toggle.isChecked():
            return
        if self.preview_panel.isHidden() or not self.live_available:
            return
        self.preview_timer.start(delay)


    def request_live_preview(self):
        if not self.current or not self.live_toggle.isChecked() or not self.live_available:
            return
        if self.preview_panel.isHidden():
            return
        try:
            config = copy.deepcopy(self.current)
            applied = render_format(self.store.load_format(config['id']), config)
            if applied:
                config['format'] = applied
            if config.get('req'):
                config['req_options'] = dict(REQ_OPTIONS)
            values = dict(self.values())
            self.preview_token += 1
            token = self.preview_token
            self.preview_status.setText('正在渲染实时预览…')
            self.preview_worker.submit(token,
                                       Path(self.current['_dir']) / self.current['file'],
                                       config, values)
        except Exception as exc:
            self.preview_status.setText('实时预览准备失败：' + str(exc)[:80])


    def _template_file(self):
        return Path(self.current['_dir']) / self.current['file']

    def _baseline_key(self):
        if not self.current:
            return None
        try:
            path = self._template_file()
            return (str(path.resolve()), path.stat().st_mtime_ns)
        except (OSError, KeyError):
            return None

    def capture_page_baseline(self):
        """用“全空填写值”渲一次，拿到模板自身的页数；输入后再变多就知道超页了。"""
        if not self.current or not self.live_toggle.isChecked():
            return
        if not self.live_available or self.preview_panel.isHidden():
            return
        key = self._baseline_key()
        if key is not None and key in self._baseline_cache:
            self.page_baseline = self._baseline_cache[key]
            return
        try:
            config = copy.deepcopy(self.current)
            applied = render_format(self.store.load_format(config['id']), config)
            if applied:
                config['format'] = applied
            if config.get('req'):
                config['req_options'] = dict(REQ_OPTIONS)
            self.baseline_token -= 1
            self.preview_worker.submit(self.baseline_token, self._template_file(),
                                       config, {})
        except Exception:
            self.page_baseline = None

    def check_page_overflow(self):
        """输入内容把文档撑到模板原本页数之外时，浮层提醒一次（同一个新高度只提醒一次）。"""
        base = self.page_baseline
        count = self.live_page_count
        if not base or not count or count <= base:
            return
        if count <= self._overflow_max:
            return
        self._overflow_max = count
        self.workspace.notify(
            f'内容已超出模板原有 {base} 页（现在 {count} 页），可能挤到下一页；'
            '生成前请检查排版。', 'info', 4600)

    def on_live_status(self, text):
        if self.live_toggle.isChecked():
            self.preview_status.setText(text)


    def on_live_rendered(self, token, pdf_bytes, count):
        if token == self.baseline_token:
            self.page_baseline = int(count)     # 空值渲染：模板自身页数
            key = self._baseline_key()
            if key is not None:
                self._baseline_cache[key] = self.page_baseline
            return
        if token != self.preview_token:
            return
        self.live_pdf = bytes(pdf_bytes)
        self.live_page_count = int(count)
        self.live_page_index = -1
        self.live_pixmap = QPixmap()
        if self.preview_page >= self.live_page_count:
            self.preview_page = max(0, self.live_page_count - 1)
        self.preview_status.setText(f'实时预览 · {self.live_page_count} 页')
        self.preview_status.setToolTip('')
        self.retry_preview_btn.hide()
        if self.page_baseline is None:
            self.page_baseline = self.live_page_count   # 基线没拿到时用首帧兜底
        else:
            self.check_page_overflow()
        self.render_preview()


    def on_live_failed(self, token, message, no_office):
        if token != self.preview_token:
            return
        if no_office:
            self.live_available = False
            self.live_toggle.blockSignals(True)
            self.live_toggle.setChecked(False)
            self.live_toggle.blockSignals(False)
            self.live_toggle.setEnabled(False)
            self.retry_preview_btn.hide()
            self.preview_status.setText('未检测到 Word/WPS，已回退原始模板预览')
        else:
            self.retry_preview_btn.show()
            self.preview_status.setText('实时预览失败：' + str(message)[:200])
        self.preview_status.setToolTip(str(message))
        self.render_preview()


    def shutdown_preview(self):
        self.preview_timer.stop()
        worker = self.preview_worker
        if worker is not None and worker.isRunning():
            worker.stop()
            if not worker.wait(15000):
                worker.terminate()
                worker.wait(2000)


    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'preview_image'):
            self.render_preview()


    def choose_image(self, field):
        path, _ = QFileDialog.getOpenFileName(self, '选择替换图片', '', '图片 (*.png *.jpg *.jpeg *.bmp)')
        if not path:
            return
        image = QImage(path)
        if image.isNull():
            QMessageBox.warning(self, '无法读取图片', '请选择有效的 PNG、JPG 或 BMP 图片。')
            return
        folder = self.store.root / 'images'
        fmt = (field.get('fmt') or '.png').lstrip('.').lower()
        save_fmt = {'jpeg': 'JPG', 'jpg': 'JPG', 'png': 'PNG',
                    'gif': 'GIF', 'bmp': 'BMP'}.get(fmt, 'PNG')
        ext = '.jpg' if save_fmt == 'JPG' else '.' + save_fmt.lower()
        try:
            folder.mkdir(exist_ok=True)
            target = folder / (uuid.uuid4().hex + ext)
            if not image.save(str(target), save_fmt):
                raise OSError('无法保存图片副本')
        except OSError as exc:
            QMessageBox.warning(self, '图片未保存', str(exc))
            return
        self.set_image(field['key'], str(target))


    def set_image(self, key, value):
        _, state = self.images[key]
        self.images[key] = (value, state)
        self.update_image_label(key)
        self.changed()


    def update_image_label(self, key):
        value, state = self.images[key]
        state.setText('保留模板原图' if value == '__original__' else '不输出此图' if not value else '已选择新图片')


    def generate_field(self, field):
        editor = self.editors.get(field['key'])
        if not self.current or not isinstance(editor, QPlainTextEdit):
            return
        if not self.workspace.ensure_ai(self):
            return
        dialog = BodyGenerateDialog(field, self.current, self.values(),
                                    self.workspace.ai_client, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        text = dialog.result_text().strip()
        if not text:
            return
        if dialog.append_mode:
            current = editor.toPlainText().rstrip()
            editor.setPlainText((current + '\n' + text) if current else text)
        else:
            editor.setPlainText(text)
        self.changed()


    def update_format_status(self):
        if not self.current:
            self.format_status.setText('')
            return
        raw = self.store.load_format(self.current['id'])
        text = format_summary(raw)
        self.format_status.setText('格式：' + text)
        self.format_status.setToolTip('生成 Word 时只作用于填写内容，模板固定文字保持原样。')


# ============================================================ ② 自动识别页


    def generate(self):
        self.generate_button.setFocus()
        QApplication.processEvents()
        if not self.current:
            return
        if not self.save_draft():
            return
        values = self.values()
        missing = [f.get('label') or f['key'] for f in self.current['fields']
                   if f.get('required') and not str(values.get(f['key'], '')).strip()]
        if missing:
            preview = '、'.join(missing[:8]) + (' 等' if len(missing) > 8 else '')
            answer = QMessageBox.question(
                self, '有必填项为空',
                '以下必填项还是空的：\n\n' + preview + '\n\n仍然生成 Word 吗？',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes)
            if answer != QMessageBox.StandardButton.Yes:
                return
        try:
            validate_values(self.current, values, allow_missing=True)
        except ValueError as exc:
            QMessageBox.information(self, '还有没填的内容', str(exc))
            return
        name = values.get('项目名称', '').strip() or self.current['title']
        safe = re.sub(r'[<>:"/\\|?*]', '_', name)[:55].strip(' .') or '文档'
        filename = f'{self.current["title"]}_{safe}_{datetime.now():%Y%m%d_%H%M%S}.docx'
        folder = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        path, _ = QFileDialog.getSaveFileName(self, '导出生成的 Word', str(Path(folder) / filename),
                                              'Word 文档 (*.docx)')
        if not path:
            return
        if not path.lower().endswith('.docx'):
            path += '.docx'
        target = Path(path).resolve()
        protected = [ASSETS.resolve(), self.store.root.resolve()]
        if any(target.is_relative_to(root) for root in protected):
            QMessageBox.warning(self, '请选择其他位置', '请把生成的文件存到文档或其他文件夹。')
            return
        self.generate_button.setEnabled(False)
        QApplication.processEvents()
        config = dict(self.current)
        applied = render_format(self.store.load_format(self.current['id']), config)
        if applied:
            config['format'] = applied
        if config.get('req'):
            config['req_options'] = dict(REQ_OPTIONS)
        report = {}
        try:
            skipped = render_word(Path(self.current['_dir']) / self.current['file'],
                                  target, config, values, allow_missing=True,
                                  keep_empty_bookmarks=True, report=report)
        except Exception as exc:
            QMessageBox.warning(self, '生成失败',
                                '请确认保存位置可写、同名文件未在 Word 中打开。\n\n' + str(exc)
                                + ('\n\n' + error_advice(exc, target)
                                   if error_advice(exc, target) else ''))
        else:
            self.last_output = target
            self.open_button.setEnabled(True)
            self.save_status.setText('Word 已生成，可以打开检查排版')
            message = 'Word 已保存：\n' + str(target)
            req = report.get('req') or {}
            bits = []
            if req.get('removed'):
                bits.append('删除说明 %d 条' % req['removed'])
            if req.get('paragraphs'):
                bits.append('套用段落 %d 段' % req['paragraphs'])
            if req.get('blank_lines'):
                bits.append('空行 %d 处' % req['blank_lines'])
            if bits:
                message += '\n\n按模板要求：' + ' · '.join(bits)
            if self.live_page_count:
                message += '\n当前预览 %d 页' % self.live_page_count
            warnings = reqparse.check_content(
                config.get('req') or {}, values,
                {f['key']: f.get('label') for f in config['fields']})
            if warnings:
                message += '\n\n内容要求（只提醒）：\n' + '\n'.join(warnings[:6])
            if skipped:
                message += ('\n\n以下字段的字号/字体未套用（需重新识别该模板）：\n'
                            + '、'.join(sorted(skipped)[:10]))
            try:
                issues = ooxmlcheck.check(target)
            except Exception:
                issues = []
            errors = [issue for issue in issues if issue['level'] == ooxmlcheck.LEVEL_ERROR]
            message += '\n\n' + self.structure_summary(target, issues)
            if not warnings and not skipped and not errors:
                # 一切正常就用轻提示，不再要求点确定；有任何需要过目的内容才弹框
                self.workspace.notify(f'已生成 {target.name} · 结构自检通过', 'ok', 3200)
            else:
                QMessageBox.information(self, '生成完成', message)
        finally:
            self.generate_button.setEnabled(True)

    def structure_summary(self, path, issues=None):
        """生成后自检：结构非法要如实说，别让 Word 的“文件需要修复”当惊喜。"""
        if issues is None:
            try:
                issues = ooxmlcheck.check(path)
            except Exception as exc:
                return f'自检没跑成（不影响已生成的文件）：{str(exc)[:80]}'
        errors = [issue for issue in issues if issue['level'] == ooxmlcheck.LEVEL_ERROR]
        if errors:
            return (f'自检发现 {len(errors)} 处结构问题，Word 可能提示修复：\n'
                    + '\n'.join(f'· {issue["where"]}：{issue["message"]}' for issue in errors[:4]))
        if issues:
            return (f'自检通过，另有 {len(issues)} 条提醒：\n'
                    + '\n'.join(f'· {issue["where"]}：{issue["message"]}' for issue in issues[:3]))
        return '自检通过：书签配对、表格骨架、样式引用、取值范围都符合 Word 的要求。'

class MarkPage(QWidget):
    # 补名进度在工作线程里产生，只能靠信号送回界面线程更新进度行。
    name_progress = Signal(int, int)

    def __init__(self, store, workspace):
        super().__init__()
        self.store, self.workspace = store, workspace
        self.source = None
        self.candidates = []
        self.uncertain = []
        self.quiet = False
        self.recon = None
        self._ai_worker = None
        self._ai_source = ''
        self._vision_worker = None
        self._vision_source = ''
        self._name_worker = None
        self._name_meta = None
        self._cancel = False
        self._normalized_from = ''
        self.example_plan = None
        self.form_plan = None
        self._example_source = ''
        self._example_description = ''
        self._name_from_recon = False
        self._naming_summary = ''
        self._skip_block_naming = False
        self.build_ui()

    def build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 18, 24, 14)
        root.setSpacing(10)
        head = QHBoxLayout()
        head.addWidget(label('自动识别 / 制作模板', 'title'))
        head.addStretch()
        head.addWidget(label('原文件永不修改', 'badge'))
        root.addLayout(head)
        root.addWidget(label('选一份 .docx（参考模板、空白表单、多表格或普通公文都行）→ 点“开始识别” → '
                             '核对/改名 → 勾选授权 → 录入模板库。会读取表格标签、段落下划线、括号提示、'
                             '内容控件和旧式表单域，不需要在 Word 里加任何占位符；同时读模板自己写明的'
                             '字体、字号、行距、缩进和内容要求，生成时照做并把说明搬成填写提示。'
                             '原文、示例答案、颜色、字体、缩进、行高和批注框全部保留，只加入不可见填写位。'
                             '规范说明类文档没有可填位时会拒绝录入。', 'muted'))
        step1 = QFrame()
        step1.setObjectName('card')
        r1 = QVBoxLayout(step1)
        r1.setContentsMargins(14, 10, 14, 10)
        r1.setSpacing(8)
        row1 = QHBoxLayout()
        row1.setSpacing(10)
        row1.addWidget(label('① 文档'))
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText('选择需要变成可填模板的 .docx…')
        row1.addWidget(self.file_edit, 1)
        self.browse_btn = button('浏览…', self.pick_file)
        row1.addWidget(self.browse_btn)
        self.recon_btn = button('开始识别', self.run_recon, primary=True)
        self.recon_btn.setEnabled(False)
        self.recon_btn.setToolTip('自动识别表格栏目、段落填空、下划线、括号提示、内容控件和表单域；'
                                  '识别后可改名并录入')
        row1.addWidget(self.recon_btn)
        r1.addLayout(row1)
        row2 = QHBoxLayout()
        row2.setSpacing(10)
        self.name_btn = button('小模型补全空标记名称', lambda: self.auto_name(from_recon=True))
        self.name_btn.setObjectName('mini')
        self.name_btn.setEnabled(False)
        row2.addWidget(self.name_btn)
        row2.addWidget(label('识别时会自动补读模板要求并把说明搬成填写提示', 'muted'))
        row2.addStretch(1)
        r1.addLayout(row2)
        root.addWidget(step1)
        step2 = QFrame()
        step2.setObjectName('card')
        r2 = QVBoxLayout(step2)
        r2.setContentsMargins(14, 10, 14, 10)
        r2.setSpacing(6)
        bar = QHBoxLayout()
        bar.addWidget(label('② 核对 / 改名（双击字段名可编辑；勾选＝允许写入；相同名称共用一次填写）'))
        bar.addStretch()
        bar.addWidget(button('全选', lambda: self.check_all(True)))
        bar.addWidget(button('取消全选', lambda: self.check_all(False)))
        bar.addWidget(button('仅勾高置信', self.check_high))
        self.diff_btn = button('用两份已填写样例比对', self.run_diff)
        self.diff_btn.hide()
        bar.addWidget(self.diff_btn)
        self.ai_btn = button('AI 识别分类', self.run_ai_classify)
        self.ai_btn.setObjectName('mini')
        self.ai_btn.setEnabled(False)
        self.ai_btn.setToolTip('让 AI 规范化字段名、判断输入方式，并自动分组')
        bar.addWidget(self.ai_btn)
        r2.addLayout(bar)
        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(
            ['授权', '填写项目名称', '分组', '输入方式', '必填', '位置', '置信', '提示 / 上下文'])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 46)
        self.table.setColumnWidth(2, 102)
        self.table.setColumnWidth(3, 108)
        self.table.setColumnWidth(4, 52)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        r2.addWidget(self.table, 1)
        self.recon_info = label('尚未识别。', 'muted')
        r2.addWidget(self.recon_info)
        self.task = feedback.TaskBar()
        self.task.cancelled.connect(self.cancel_task)
        self.name_progress.connect(lambda done, total: self.task.step(done, total))
        r2.addWidget(self.task)
        root.addWidget(step2, 1)
        step3 = QFrame()
        step3.setObjectName('card')
        r3 = QHBoxLayout(step3)
        r3.setContentsMargins(14, 10, 14, 10)
        r3.setSpacing(10)
        r3.addWidget(label('③ 模板名称 / 录入'))
        self.name_edit = QLineEdit()
        r3.addWidget(self.name_edit, 1)
        self.category_edit = QLineEdit()
        self.category_edit.setPlaceholderText('分类，如 科研项目')
        self.category_edit.setFixedWidth(132)
        r3.addWidget(self.category_edit)
        self.register_btn = button('确认改名并录入模板库', self.register, primary=True)
        self.register_btn.setEnabled(False)
        r3.addWidget(self.register_btn)
        self.save_btn = button('仅保存填写位副本', self.save_only)
        self.save_btn.setEnabled(False)
        r3.addWidget(self.save_btn)
        root.addWidget(step3)

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择 Word 文档', '', 'Word 文档 (*.docx)')
        if path:
            self.load(path)

    def load(self, path):
        self.source = Path(path)
        self.file_edit.setText(str(self.source))
        self.name_edit.setText(self.source.stem[:60])
        self.category_edit.clear()
        self.recon = None
        self.candidates = []
        self.uncertain = []
        self.table.setRowCount(0)
        self.recon_info.setText('尚未识别。')
        self.ai_btn.setEnabled(False)
        self.recon_btn.setEnabled(True)
        self.name_btn.setEnabled(True)
        self.register_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.example_plan = None
        self.form_plan = None
        self._example_source = ''
        self._example_description = ''
        self._name_meta = None
        self._normalized_from = ''
        self._name_from_recon = False
        self._naming_summary = ''
        self._skip_block_naming = False

    def inform(self, title, text):
        if not self.quiet:
            QMessageBox.information(self, title, text)

    def done(self, text, ms=2600):
        """成功类反馈走右下角浮层，不再要求点确定；需要过目的内容仍用 inform。"""
        if not self.quiet:
            self.workspace.notify(text, 'ok', ms)

    def warn(self, title, text):
        if not self.quiet:
            QMessageBox.warning(self, title, text)

    def run_recon(self):
        naming_summary = self._naming_summary
        self._naming_summary = ''
        self.example_plan = None
        self.form_plan = None
        self._example_source = ''
        self._example_description = ''
        if not self.source or self._vision_worker is not None:
            return
        self.recon_btn.setEnabled(False)
        self.recon_info.setText('正在识别：表格、段落填空、下划线、内容控件和表单域…')
        QApplication.processEvents()
        structural_error = ''
        try:
            plan = smartparse.analyze(str(self.source))
        except Exception as exc:
            self.table.setRowCount(0)
            self.register_btn.setEnabled(False)
            self.save_btn.setEnabled(False)
            self.ai_btn.setEnabled(False)
            self.recon_info.setText('未发现可填写项。')
            advice = error_advice(exc, self.source)
            self.warn('没有识别到可填写项',
                      f'{exc}\n\n原文件没有被修改。' + (f'\n\n{advice}' if advice else ''))
            return
        finally:
            self.recon_btn.setEnabled(True)
        self._apply_plan(plan, 'smart', auto_ai=False, naming_summary=naming_summary)
        if structural_error:
            self.recon_info.setText(self.recon_info.text() + ' 结构识别未发现栏目，正在交给文本模型和模板说明复核。')
        self._start_enhance(plan)

    def _auto_classify_after_recon(self):
        if self.quiet:
            return
        ok, _ = self.workspace.ai_client.usable()
        if ok:
            QTimer.singleShot(0, self.run_ai_classify)

    def _start_enhance(self, plan):
        use_format = True
        remove_notes = True
        text_cfg = self.workspace._text_cfg()
        vision_cfg = self.workspace._vision_cfg()
        self._vision_source = str(self.source)
        self._set_name_busy(True)
        steps = ' + '.join(([('文本模型找位')] if text_cfg else [])
                           + ([('读模板要求')] if use_format else [])
                           + ([('视觉复核')] if vision_cfg else []))
        self.recon_info.setText(f'{steps}：正在通读文档结构…')
        self._cancel = False
        self.task.start(steps + '：正在通读文档结构…')
        source = self._vision_source

        def job(emit):
            def report(text):
                if self._cancel:
                    raise feedback.Cancelled('已取消：本次增强识别的结果没有应用。')
                emit(text)

            current = plan
            if text_cfg:
                current = textparse.enhance(source, current, text_cfg, progress=report)
            if use_format:
                current = reqparse.enhance(source, current, progress=report)
                current['req_options'] = {'apply_format': True,
                                          'remove_notes': remove_notes}
            if vision_cfg:
                current = visionparse.verify(source, current, vision_cfg, progress=report)
            return current

        worker = AIWorker(job, self, with_status=True)
        self._vision_worker = worker
        worker.status.connect(self._vision_status)
        worker.done.connect(self.apply_enhanced)
        worker.failed.connect(self.enhance_failed)
        worker.finished.connect(self._vision_finished)
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def _vision_status(self, text):
        self.recon_info.setText(text)

    def apply_enhanced(self, merged):
        if self._vision_source != str(self.source):
            self.warn('结果已丢弃', '识别返回前文档已变化，本次结果没有应用。')
            return
        if not merged.get('fields'):
            self.table.setRowCount(0)
            self.candidates = []
            self.register_btn.setEnabled(False)
            self.save_btn.setEnabled(False)
            self.ai_btn.setEnabled(False)
            self.recon_info.setText('结构、文本模型和要求识别都没有找到可填写栏目。')
            self.warn('没有识别到可填写项',
                      '结构识别、文本模型和模板说明都没有找到明确填写位。\n'
                      '这份文档可能只是规范/说明材料。\n原文件没有被修改。')
            return
        self._apply_plan(merged, 'smart', auto_ai=True,
                         naming_summary=self._naming_summary)
        lines = []
        text = merged.get('text') or {}
        if text.get('used'):
            lines.append(f'文本模型（{text.get("model") or "文本模型"}）：校正 '
                         f'{text.get("matched", 0)} 项、补识别 {text.get("added", 0)} 项'
                         + (f'、{len(text.get("rejected") or [])} 项未采纳'
                            if text.get('rejected') else ''))
        elif text.get('error'):
            lines.append(f'文本模型未完成：{str(text["error"])[:120]}')
        vision = merged.get('vision') or {}
        if vision.get('checked'):
            lines.append(f'视觉复核 {vision.get("pages", 0)} 页：确认 '
                         f'{vision.get("confirmed", 0)} 项'
                         + (f'，改正 {vision.get("fixed", 0)} 项' if vision.get('fixed') else ''))
        if merged.get('req_summary'):
            lines.append('模板要求：' + merged['req_summary'])
        if merged.get('req_error'):
            lines.append(f'要求识别未完成：{str(merged["req_error"])[:120]}')
        if lines:
            self.inform('识别完成', '\n'.join(lines)
                        + '\n\n请核对字段名与勾选，然后录入模板库。')
        self.ask_questions(merged.get('req_questions') or [])

    def cancel_task(self):
        self._cancel = True

    def cancel_flag(self):
        return self._cancel

    def enhance_failed(self, message):
        message = str(message)
        if message.startswith('已取消'):
            self.recon_info.setText(message + '（结构识别的结果仍在，原文件未改动。）')
            self.workspace.notify(message, 'info', 2600)
            return
        self.recon_info.setText(
            self.recon_info.text() + f' · 增强识别失败：{message[:120]}；已保留结构识别结果。')
        self.warn('增强识别未完成', f'{message}\n\n已保留结构识别结果，原文件未修改。')
        self._auto_classify_after_recon()

    def ask_questions(self, questions):
        if not questions or self.quiet:
            return
        dialog = QuestionsDialog(questions, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.recon_info.setText(self.recon_info.text()
                                    + ' · 待确认问题未回答，按默认选项处理。')
            return
        plan = self.form_plan or self.example_plan
        if plan is not None:
            reqparse.refresh_plan_req(plan, dialog.answers)
            self.recon_info.setText(
                self.recon_info.text()
                + f' · 已按你的选择重算要求（{len(dialog.answers)} 项决定）。')

    def _vision_finished(self):
        self._vision_worker = None
        self.task.finish()
        self._reset_mark_buttons()

    def _apply_plan(self, plan, kind='example', auto_ai=True, naming_summary=''):
        self.example_plan = plan if kind == 'example' else None
        self.form_plan = plan if kind in ('form', 'smart') else None
        self._example_source = str(self.source)
        self._example_description = plan.get('description', '')
        self._apply_kind = kind
        self.recon = None
        self.candidates = []
        self.uncertain = []
        self.table.setRowCount(0)
        used = set()
        for field in plan['fields']:
            f = dict(field)
            base = re.sub(r'[^\w\u4e00-\u9fff]+', '_',
                          (f.get('key') or f.get('label') or 'field')).strip()[:24].strip('_') or 'field'
            fid, n = base, 2
            while fid in used:
                fid = f'{base}_{n}'
                n += 1
            used.add(fid)
            f['fid'] = fid
            f['example_id'] = f.get('plan_id')
            self.candidates.append(f)
            row = self.table.rowCount()
            self.table.insertRow(row)
            check = QTableWidgetItem()
            check.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            check.setCheckState(Qt.CheckState.Checked)
            self.table.setItem(row, 0, check)
            name = QTableWidgetItem(fid)
            name.setFlags(Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled)
            self.table.setItem(row, 1, name)
            self.table.setItem(row, 2, QTableWidgetItem(f.get('group') or DEFAULT_GROUP))
            combo = QComboBox()
            combo.addItem('单行文字', 'text')
            combo.addItem('多行正文', 'multiline')
            combo.setCurrentIndex(1 if f.get('type') == 'multiline' else 0)
            self.table.setCellWidget(row, 3, combo)
            req = QTableWidgetItem()
            req.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            req.setCheckState(Qt.CheckState.Checked if f.get('required') else Qt.CheckState.Unchecked)
            self.table.setItem(row, 4, req)
            self.table.setItem(row, 5, QTableWidgetItem(f.get('where') or ''))
            if f.get('source') == '文本模型':
                conf_text = '文本'
            elif f.get('text_checked'):
                conf_text = '结构+文本'
            else:
                conf_text = {'example': '示例值', 'form': '表单', 'smart': '自动'}[kind]
            conf = QTableWidgetItem(conf_text)
            conf.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self.table.setItem(row, 6, conf)
            original = str(f.get('original_content') or '')
            if f.get('source') == '文本模型':
                tip = '文本模型定位 · ' + (original[:18] or '空白')
            else:
                tip = ((f'示例正文 {len(original)} 字' if kind == 'example' else
                        (f'栏目正文 {len(original)} 字' if original else '空白栏目'))
                       if f.get('mode') == 'section'
                       else (f'示例值 {original[:18]}' if kind == 'example'
                             else (f'原值 {original[:18]}' if original else '空白值')))
            tip_item = QTableWidgetItem(tip)
            tip_item.setToolTip(str(f.get('context') or ''))
            self.table.setItem(row, 7, tip_item)
        title = (plan.get('title') or self.source.stem)[:60]
        self.name_edit.setText(title)
        self.category_edit.setText(plan.get('category') or '')
        vision = plan.get('vision') or {}
        text = plan.get('text') or {}
        what = {'example': '示例模板', 'form': '表单结构（空白大段/多表格）',
                'smart': ('自动识别 + 文本模型' if text.get('used')
                          else '自动识别（无占位符）')}[kind]
        extra = ''
        if text.get('used'):
            extra = (f' 文本模型通读结构：校正 {text.get("matched", 0)} 项、'
                     f'补识别 {text.get("added", 0)} 项'
                     + (f'、{len(text.get("rejected") or [])} 项未采纳'
                        if text.get('rejected') else '') + '。')
        elif text.get('error'):
            extra = f' 文本模型未完成：{str(text["error"])[:90]}；已保留结构识别结果。'
        if vision.get('checked'):
            extra += (f' 视觉复核 {vision.get("pages", 0)} 页：确认 '
                      f'{vision.get("confirmed", 0)} 项'
                      + (f'、改正 {vision.get("fixed", 0)} 项' if vision.get('fixed') else '')
                      + '。')
        if plan.get('req_summary'):
            extra += ' ' + plan['req_summary'] + '；说明文字将搬成填写栏下方的提示。'
        if plan.get('req_error'):
            extra += f' 要求识别未完成：{str(plan["req_error"])[:80]}。'
        self.recon_info.setText(
            f'识别完成：{what}，找到 {len(plan["fields"])} 个填写项，已全部勾选。'
            '原文档文字、示例答案、颜色和版式全部保留，只加入不可见填写位；'
            '生成 Word 时才替换成填写内容。可直接双击字段名改名，'
            '确认后点“确认改名并录入模板库”。'
            + extra + naming_summary)
        self.register_btn.setEnabled(bool(plan['fields']))
        self.save_btn.setEnabled(bool(plan['fields']))
        self.ai_btn.setEnabled(bool(plan['fields']))
        self._reset_mark_buttons()
        if not auto_ai or self.quiet:
            return
        ok, _ = self.workspace.ai_client.usable()
        if ok:
            QTimer.singleShot(0, self.run_ai_classify)

    def run_diff(self):
        self.example_plan = None
        self.form_plan = None
        self._example_source = ''
        path_a, _ = QFileDialog.getOpenFileName(self, '选择已填写样例 1', '', 'Word 文档 (*.docx)')
        if not path_a:
            return
        path_b, _ = QFileDialog.getOpenFileName(self, '选择已填写样例 2（结构相同、内容不同）', '',
                                                'Word 文档 (*.docx)')
        if not path_b:
            return
        if not self.recon or Path(self.recon.get('file', '')).resolve() != Path(path_a).resolve():
            self.recon = None
        try:
            data = pipeline.phase_diff(path_a, path_b)
        except Exception as exc:
            self.warn('比对失败', str(exc))
            return
        if not data['structure_same']:
            self.warn('段落数不同',
                      f'两份样例段落数不一致（{data["paragraphs_a"]} vs {data["paragraphs_b"]}），'
                      '差异比对要求结构相同。')
            return
        if not data['fields']:
            self.inform('没有差异', '两份样例文字内容一致，没有可提取的变量。')
            return
        if not self.candidates:
            self.source = Path(path_a)
            self.file_edit.setText(str(self.source))
            self.name_edit.setText(self.source.stem[:60])
        used = {re.sub(r'[^\w\u4e00-\u9fff]+', '_',
                       self.table.item(r, 1).text() if self.table.item(r, 1) else '')
                for r in range(self.table.rowCount())}
        added = 0
        for f in data['fields']:
            fid = f['field_id']
            n = 2
            while fid in used:
                fid = f"{f['field_id']}_{n}"
                n += 1
            used.add(fid)
            f['fid'] = fid
            self.candidates.append(f)
            row = self.table.rowCount()
            self.table.insertRow(row)
            cell0 = QTableWidgetItem()
            cell0.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            cell0.setCheckState(Qt.CheckState.Checked)
            self.table.setItem(row, 0, cell0)
            name = QTableWidgetItem(fid)
            name.setFlags(Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled)
            self.table.setItem(row, 1, name)
            self.table.setItem(row, 2, QTableWidgetItem('补充项'))
            combo = QComboBox()
            combo.addItem('单行文字', 'text')
            combo.addItem('多行正文', 'multiline')
            combo.setCurrentIndex(1 if is_multiline_label(fid) else 0)
            self.table.setCellWidget(row, 3, combo)
            req = QTableWidgetItem()
            req.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            req.setCheckState(Qt.CheckState.Unchecked)
            self.table.setItem(row, 4, req)
            self.table.setItem(row, 5, QTableWidgetItem(f"样例差异 @ 段落 {f['location']['p'] + 1}"))
            c = QTableWidgetItem('样例差异')
            c.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self.table.setItem(row, 6, c)
            self.table.setItem(row, 7, QTableWidgetItem(
                f"A≈{f['original_content'][:14]} / B≈{f['alt_content'][:14]}"))
            added += 1
        self.register_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        self.ai_btn.setEnabled(True)
        self.recon_info.setText(f'差异比对新增 {added} 个候选（两份样例写法不同处）。'
                                '勾选＝授权将该段替换为填写位。')

    def _candidate_context(self, candidate):
        if candidate.get('context'):
            return str(candidate['context'])
        loc = candidate.get('location') or {}
        if self.recon and loc.get('kind') == 'paragraph':
            p = loc.get('p', 0)
            texts = []
            for i in range(max(0, p - 1), min(len(self.recon['paragraphs']), p + 2)):
                if i != p and self.recon['paragraphs'][i]['text'].strip():
                    texts.append(self.recon['paragraphs'][i]['text'][:60])
            if texts:
                return ' / '.join(texts)
        if self.recon and loc.get('kind') == 'cell':
            table = next((t for t in self.recon['tables'] if t['index'] == loc.get('table')), None)
            cell = next((c for c in (table or {}).get('cells', [])
                         if c['row'] == loc.get('row') and c['col'] == loc.get('col')), None)
            if table and cell:
                label = (pipeline._left_label(table, cell, table['index'])
                         or pipeline._up_label(table, cell, table['index'])
                         or pipeline._row_label(table, cell) or '')
                return (label + ' ' + cell.get('text', '')).strip()
        return str(candidate.get('original_content') or '')

    def run_ai_classify(self):
        if not self.candidates or not self.source or self._ai_worker is not None:
            return
        if not self.workspace.ensure_ai(self):
            return
        try:
            recon = self.recon or pipeline.phase_recon(str(self.source))
        except Exception as exc:
            self.warn('读取文档失败', str(exc))
            return
        self.recon = recon
        outline = [(p['index'], p['text'][:60]) for p in recon['paragraphs'] if p['text'].strip()][:70]
        for table in recon['tables']:
            for cell in table['cells']:
                if cell.get('text', '').strip() and len(outline) < 90:
                    outline.append((f"表{table['index'] + 1}行{cell['row'] + 1}",
                                    cell['text'][:50]))
        candidates = []
        for i, field in enumerate(self.candidates):
            name_item = self.table.item(i, 1)
            where_item = self.table.item(i, 5)
            group_item = self.table.item(i, 2)
            combo = self.table.cellWidget(i, 3)
            candidates.append({
                'where': where_item.text() if where_item else '',
                'label': name_item.text() if name_item else field.get('fid', ''),
                'type': combo.currentData() if combo else field.get('type', 'text'),
                'marked': field.get('source') == '标记',
                'group': group_item.text() if group_item else DEFAULT_GROUP,
                'context': self._candidate_context(field),
            })
        prompt = ai.field_prompt(outline, candidates, self.name_edit.text().strip())
        client = self.workspace.ai_client
        self._ai_source = str(self.source)
        self.ai_btn.setEnabled(False)
        self.ai_btn.setText('AI 识别中…')
        self.recon_btn.setEnabled(False)
        self.browse_btn.setEnabled(False)
        self.diff_btn.setEnabled(False)
        self.register_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.table.setEnabled(False)
        worker = AIWorker(lambda: client.ask_json(ai.FIELD_SYSTEM, prompt, 'fields'), self)
        self._ai_worker = worker
        worker.done.connect(self.apply_ai_classify)
        worker.failed.connect(self.ai_classify_failed)
        worker.finished.connect(self._ai_finished)
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def _ai_finished(self):
        self._ai_worker = None
        self.ai_btn.setText('AI 识别分类')
        self.ai_btn.setEnabled(bool(self.candidates))
        self.recon_btn.setEnabled(bool(self.source))
        self.browse_btn.setEnabled(True)
        self.diff_btn.setEnabled(True)
        self.register_btn.setEnabled(bool(self.candidates))
        self.save_btn.setEnabled(bool(self.candidates))
        self.table.setEnabled(True)

    def ai_classify_failed(self, message):
        self.warn('AI 识别失败', f'{message}\n\n可在顶栏「AI 设置」检查服务和模型。')

    def _set_name_busy(self, busy, text=''):
        for widget in (self.browse_btn, self.recon_btn, self.diff_btn, self.ai_btn,
                       self.register_btn, self.save_btn, self.name_btn):
            widget.setEnabled(not busy)
        self.table.setEnabled(not busy)

    def _reset_mark_buttons(self):
        if (self._ai_worker is not None or self._name_worker is not None
                or self._vision_worker is not None):
            return
        self.browse_btn.setEnabled(True)
        self.recon_btn.setEnabled(bool(self.source))
        self.name_btn.setEnabled(bool(self.source))
        self.diff_btn.setEnabled(True)
        self.ai_btn.setEnabled(bool(self.candidates))
        self.register_btn.setEnabled(bool(self.candidates))
        self.save_btn.setEnabled(bool(self.candidates))
        self.table.setEnabled(True)

    def auto_name(self, from_recon=False):
        if not self.source or self._name_worker is not None or self._ai_worker is not None:
            return
        self._name_from_recon = from_recon
        try:
            blocks = blockfill.find_blocks(str(self.source))
        except Exception as exc:
            self.warn('读取文档失败', str(exc))
            return
        if not blocks:
            if from_recon:
                self._skip_block_naming = True
                try:
                    self.run_recon()
                finally:
                    self._skip_block_naming = False
            else:
                self.inform('没有可补全的空标记',
                            '没有发现需要补全的空白区域。\n'
                            '如果字段名已经写在文档里，直接在表格中检查即可；'
                            '如需规范已有字段名，可点“AI 识别分类”。')
            return
        existing = blockfill.existing_names(str(self.source))
        self._name_meta = {'blocks': blocks, 'existing': existing,
                           'source': str(self.source)}
        assets = embedded_ai.find_assets()
        if not assets:
            if from_recon:
                self._normalize_blocks(blocks, {}, False, ['未找到内置模型'], from_recon=True)
                return
            answer = QMessageBox.question(
                self, '内置模型未找到',
                embedded_ai.missing_hint() + '\n\n是否先用上下文规则自动命名？',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes)
            if answer != QMessageBox.StandardButton.Yes:
                return
            self._normalize_blocks(blocks, {}, False, ['未找到内置模型'])
            return
        items, used = blockfill.model_items(blocks, existing)
        if not items:
            self._normalize_blocks(blocks, {}, False, [], from_recon=from_recon)
            return
        self._set_name_busy(True)
        self.recon_info.setText(
            ('识别中：正在补全字段名…' if from_recon
             else '正在用内置小模型补全字段名…')
            + '首次加载模型可能需要几十秒。')
        client = embedded_ai.shared(self.workspace.store.root / 'embedded_ai.log')
        title = self.name_edit.text().strip()
        self._cancel = False
        self.task.start(f'内置小模型补全字段名：共 {len(items)} 处，每 6 处一批…')
        worker = AIWorker(lambda: client.suggest_names(
            items, title, used,
            progress=lambda done, total: self.name_progress.emit(int(done), int(total)),
            should_stop=self.cancel_flag), self)
        self._name_worker = worker
        worker.done.connect(self._block_names_ready)
        worker.failed.connect(self._block_names_failed)
        worker.finished.connect(self._name_finished)
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def _block_names_ready(self, result):
        meta = self._name_meta or {}
        if not meta or meta.get('source') != str(self.source):
            self.warn('结果已丢弃', '文档已变化，本次补全结果没有应用。')
            return
        suggestions, errors = result
        self._normalize_blocks(meta['blocks'], suggestions, True, errors,
                               from_recon=self._name_from_recon)

    def _block_names_failed(self, message):
        meta = self._name_meta or {}
        if not meta:
            return
        from_recon = self._name_from_recon
        self._normalize_blocks(meta['blocks'], {}, False, [message], silent=True,
                               from_recon=from_recon)
        if not from_recon:
            QMessageBox.warning(self, '内置模型未完成',
                                f'{message}\n\n已改用上下文规则生成字段名，请在表格中检查。')

    def _normalize_blocks(self, blocks, suggestions, model_used, errors, silent=False,
                          from_recon=False):
        original = self.source
        if not original:
            return
        try:
            existing = (self._name_meta or {}).get('existing', [])
            blocks = blockfill.resolve_names(blocks, suggestions, existing)
            out_dir = self.workspace.store.root / 'auto_named'
            out_dir.mkdir(parents=True, exist_ok=True)
            safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', original.stem)[:50]
            out = out_dir / f'{safe}_{datetime.now():%Y%m%d_%H%M%S}_已补全.docx'
            result = blockfill.normalize(str(original), str(out), blocks)
        except Exception as exc:
            self._reset_mark_buttons()
            self.warn('补全失败', f'{exc}\n\n原文件未修改。')
            return
        self._name_meta = None
        self._normalized_from = str(original)
        self.source = Path(out)
        self.example_plan = None
        self.form_plan = None
        self._example_source = ''
        self.file_edit.setText(str(out))
        self.recon = None
        self.candidates = []
        self.uncertain = []
        self.table.setRowCount(0)
        names = result['fields']
        model_count = sum(1 for b in blocks if b.get('name_source') == 'model')
        info = f' · 自动补名 {len(names)} 个'
        if model_count:
            info += f'（模型 {model_count} 个）'
        if result['skipped']:
            info += f' · 截图空白 {result["skipped"]} 处'
        self._naming_summary = info
        self._name_from_recon = False
        self._skip_block_naming = True
        try:
            self.run_recon()
        finally:
            self._skip_block_naming = False
        self._reset_mark_buttons()
        if silent or from_recon:
            return
        lines = [f'已补全 {len(names)} 个字段名。']
        if result['skipped']:
            lines.append(f'{result["skipped"]} 处截图区域保留空白，可在生成的 Word 里手动粘贴。')
        if model_used and model_count:
            lines.append(f'其中 {model_count} 个由内置小模型根据上下文命名。')
        elif not model_used:
            lines.append('本次使用上下文规则命名（内置模型未使用）。')
        if errors:
            lines.append('部分区域回退到规则命名。')
        lines.append('原文件未修改；当前使用自动生成的临时副本。\n请检查字段名，勾选后录入模板库。')
        self.inform('字段名补全完成', '\n'.join(lines))

    def _name_finished(self):
        self._name_worker = None
        if self._ai_worker is None:
            self._reset_mark_buttons()
        if self._vision_worker is None:
            # 补名结束后可能已经接力启动了视觉复核，那一步还在用同一行进度
            self.task.finish()

    def apply_ai_classify(self, data):
        if self._ai_source != str(self.source):
            self.warn('结果已丢弃', 'AI 返回前文档或识别结果已变化，本次分类结果没有应用。')
            return
        items = data.get('fields') if isinstance(data, dict) else None
        if not isinstance(items, list):
            self.warn('AI 返回异常', '没有拿到字段分类结果，请重试或换用其他模型。')
            return
        used = {self.table.item(r, 1).text().strip() for r in range(self.table.rowCount())}
        groups, renamed, applied = set(), 0, 0
        for item in items:
            try:
                i = int(item.get('i'))
            except (TypeError, ValueError):
                continue
            if not (0 <= i < len(self.candidates) and i < self.table.rowCount()):
                continue
            field = self.candidates[i]
            combo = self.table.cellWidget(i, 3)
            kind = item.get('type')
            if combo and kind in ('text', 'multiline'):
                combo.setCurrentIndex(1 if kind == 'multiline' else 0)
            group = str(item.get('group') or '').strip()[:12] or DEFAULT_GROUP
            if self.table.item(i, 2):
                self.table.item(i, 2).setText(group)
            groups.add(group)
            label = str(item.get('label') or '').strip()
            if label and field.get('source') != '标记' and self.table.item(i, 1):
                fid = re.sub(r'[^\w\u4e00-\u9fff]+', '_', label)[:24].strip('_') or f'field{i + 1}'
                old = self.table.item(i, 1).text().strip()
                used.discard(old)
                base, n = fid, 2
                while fid in used:
                    fid = f'{base}_{n}'
                    n += 1
                used.add(fid)
                self.table.item(i, 1).setText(fid)
                self.candidates[i]['fid'] = fid
                self.candidates[i]['display_fid'] = fid
                self.candidates[i]['label'] = label
                renamed += 1
            applied += 1
        title = str(data.get('template_title') or '').strip()
        stem = self.source.stem[:60] if self.source else ''
        if title and self.name_edit.text().strip() in ('', stem):
            self.name_edit.setText(title[:60])
        category = str(data.get('category') or '').strip()
        if category and not self.category_edit.text().strip():
            self.category_edit.setText(category[:12])
        self.recon_info.setText(
            self.recon_info.text() + f' · AI 已归类 {applied} 项到 {len(groups)} 个分组')
        self.inform('AI 识别分类完成',
                    f'已处理 {applied} 个候选，更新 {renamed} 个字段名，归入 {len(groups)} 个分组。\n'
                    '已有明确结构的字段名保持不变。请核对/改名后勾选，再点“确认改名并录入模板库”。')

    def check_all(self, value):
        for row in range(self.table.rowCount()):
            self.table.item(row, 0).setCheckState(
                Qt.CheckState.Checked if value else Qt.CheckState.Unchecked)

    def check_high(self):
        for row in range(self.table.rowCount()):
            state = Qt.CheckState.Checked if self.candidates[row]['confidence'] in ('高', '示例') \
                else Qt.CheckState.Unchecked
            self.table.item(row, 0).setCheckState(state)

    def checked_fields(self):
        out, names = [], {}
        for row in range(self.table.rowCount()):
            if self.table.item(row, 0).checkState() != Qt.CheckState.Checked:
                continue
            raw = self.table.item(row, 1).text().strip()
            fid = re.sub(r'[^\w\u4e00-\u9fff]+', '_', raw)[:24].strip('_')
            if not fid:
                fid = f'field{row + 1}'
            self.table.item(row, 1).setText(fid)
            f = dict(self.candidates[row])
            previous = str(f.get('display_fid') or f.get('fid') or '')
            if raw != previous:
                f['label'] = (raw or fid)[:40]
            f['display_fid'] = fid
            f['fid'] = fid
            f['type'] = self.table.cellWidget(row, 3).currentData()
            f['required'] = self.table.item(row, 4).checkState() == Qt.CheckState.Checked
            group = self.table.item(row, 2).text().strip() if self.table.item(row, 2) else ''
            f['group'] = group[:12] or DEFAULT_GROUP
            names[fid] = names.get(fid, 0) + 1
            out.append(f)
        dup = {k for k, v in names.items() if v > 1}
        return out, dup

    def generate(self, tmp):
        fields, dup = self.checked_fields()
        if not fields:
            self.inform('没有授权字段', '请先勾选至少一个填写位置。')
            return None
        plan_mode = (bool(self.example_plan) or bool(self.form_plan)) \
            and self._example_source == str(self.source)
        active_plan = self.example_plan or self.form_plan
        bookmark_mode = plan_mode and all(
            f.get('control') or (f.get('bookmark_range') and f.get('bookmark'))
            for f in fields)
        if plan_mode:
            marked = tmp / 'template.docx'
            selected = {f['example_id']: f['fid'] for f in fields if f.get('example_id')}
            try:
                if bookmark_mode:
                    built = bookmarkfill.build(str(self.source), str(marked),
                                               active_plan, selected)
                else:
                    built = exampleparse.build(str(self.source), str(marked),
                                               active_plan, selected)
            except Exception as exc:
                self.warn('模板生成失败', f'{exc}\n\n原文件未修改。')
                return None
            if bookmark_mode:
                names = set(bookmarkfill.scan_bookmarks(marked))
                declared = {f['bookmark'] for f in fields if f.get('bookmark')}
                keys = {n for n in names if n in declared or n.startswith('WTT5_')}
            else:
                keys = set(built['markers'])
                declared = {f['fid'] for f in fields}
            mismatch = (declared - keys) if bookmark_mode else (keys ^ declared)
            if mismatch:
                detail = []
                if declared - keys:
                    detail.append('缺少：' + '、'.join(sorted(declared - keys)))
                if not bookmark_mode and keys - declared:
                    detail.append('多出：' + '、'.join(sorted(keys - declared)))
                self.warn('字段校验未通过',
                          '生成文件中的字段与勾选项不一致：\n' + '\n'.join(detail)
                          + '\n\n原文件未改动，未注册。')
                return None
            res = {'markers': sorted(keys),
                   'mode': 'bookmark' if bookmark_mode else ('example' if self.example_plan else 'form')}
            if bookmark_mode:
                res['bookmarks'] = sorted(keys)
        else:
            whitelist = {'fields': [{'field_id': f['fid'], 'label': f['fid'],
                                     'location': f['location'],
                                     'replace_para': f.get('source') == '样例差异'} for f in fields]}
            wpath = tmp / 'whitelist.json'
            wpath.write_text(json.dumps(whitelist, ensure_ascii=False, indent=2), encoding='utf-8')
            marked, log_path = tmp / 'template.docx', tmp / 'fill_log.json'
            pipeline.phase_fill(str(self.source), str(wpath), None, 'mark', str(marked), str(log_path))
            res = pipeline.phase_verify(str(self.source), str(marked), str(log_path),
                                        str(tmp / 'verify.json'))
            if res['blocked_count']:
                detail = '\n'.join('· ' + b for b in res['blocked'][:6])
                self.warn('验收存在阻断项', f'结构白名单校验未通过：\n{detail}\n\n原文件未改动，未注册。')
                return None
        merged, order, images = {}, [], []
        for f in fields:
            if f['location']['kind'] == 'image':
                loc = f['location']
                images.append({'key': f['fid'], 'label': f.get('label') or f['fid'], 'group': '图片',
                               'part': loc['part'], 'relationship': loc['relationship'],
                               'media': loc['media'], 'fmt': Path(loc['media']).suffix.lower()})
                continue
            display = str(f.get('label') or f['fid']).strip() or f['fid']
            # 同名（含同义写法）的栏目共用一个 key：配置里每个位置仍是一条记录（各自书签），
            # 但填写界面只出一个输入框，生成时同一个值写到所有位置。
            slot = f.get('key') or f['fid']
            item = {'key': slot, 'label': display, 'type': f['type'],
                    'required': f['required'], 'group': f.get('group') or DEFAULT_GROUP,
                    'max_length': 0, 'default': ''}
            if bookmark_mode and f.get('bookmark'):
                item['bookmark'] = f['bookmark']
            if bookmark_mode and f.get('control'):
                item['control'] = f['control']
            if f.get('choices'):
                item['choices'] = f['choices']
            if f.get('options'):
                item['options'] = f['options']
            if f.get('grid'):
                item['grid'] = f['grid']
            if slot not in merged:
                merged[slot] = []
                order.append(slot)
            merged[slot].append(item)
        title = self.name_edit.text().strip() or self.source.stem
        category = self.category_edit.text().strip() or '自定义模板'
        description = (self._example_description if plan_mode and self._example_description
                       else '由自动识别生成：仅勾选位置可填，其余内容按原模板输出。')
        config = dict(id='', title=title, category=category,
                      description=description,
                      file='template.docx',
                      fields=[item for slot in order for item in merged[slot]],
                      grids=_selected_grids(active_plan, fields), images=images)
        if plan_mode and (active_plan or {}).get('req'):
            req = copy.deepcopy(active_plan['req'])
            remap = {f.get('key'): (f.get('key') or f['fid']) for f in fields if f.get('key')}
            for name in ('field_rules', 'field_ranges', 'hints'):
                req[name] = {remap[k]: value for k, value in (req.get(name) or {}).items()
                             if k in remap}
            for item in (req.get('notes') or []) + (req.get('content_rules') or []):
                if item.get('field_key') in remap:
                    item['field_key'] = remap[item['field_key']]
            config['req'] = req
            config['req_options'] = dict((active_plan or {}).get('req_options')
                                         or {'apply_format': True, 'remove_notes': True})
            config['req_summary'] = (active_plan or {}).get('req_summary') or ''
        if dup:
            names = '、'.join(sorted(dup))
            self.inform('同名合并', f'“{names}”出现在多处，已合并为同一填写项。')
        return {'marked': marked, 'config': config, 'fields': fields, 'res': res}

    def register(self):
        if not self.name_edit.text().strip():
            self.inform('缺少模板名称', '请填写第③步的模板名称。')
            return
        tid = 'custom-' + uuid.uuid4().hex[:12]
        with tempfile.TemporaryDirectory(prefix='workbench-') as td:
            try:
                result = self.generate(Path(td))
            except Exception as exc:
                self.warn('生成填写位失败', f'{exc}\n\n原文件未被修改。')
                return
            if not result:
                return
            result['config']['id'] = tid
            folder = self.store.custom / tid
            try:
                folder.mkdir(parents=True, exist_ok=True)
                shutil.copy2(result['marked'], folder / 'template.docx')
                write_json(folder / 'config.json', result['config'])
            except OSError as exc:
                self.warn('注册失败', f'无法写入模板库：\n{exc}\n\n可用“仅保存填写位副本”导出后手动导入。')
                return
        self.workspace.refresh_all()
        self.workspace.go_fill(tid)
        self.done(f'已录入模板库：{result["config"]["title"]} · '
                  f'{len(result["config"]["fields"])} 个填写字段，已切到填写页', 3200)

    def save_only(self):
        base = Path(self._normalized_from) if self._normalized_from else self.source
        default = str(base.parent / f'{base.stem}_填写位.docx')
        path, _ = QFileDialog.getSaveFileName(self, '保存填写位副本', default, 'Word 文档 (*.docx)')
        if not path:
            return
        with tempfile.TemporaryDirectory(prefix='workbench-') as td:
            try:
                result = self.generate(Path(td))
            except Exception as exc:
                self.warn('生成填写位失败', f'{exc}\n\n原文件未被修改。')
                return
            if not result:
                return
            try:
                shutil.copy2(result['marked'], path)
            except OSError as exc:
                self.warn('保存失败', str(exc))
                return
        self.done(f'填写位副本已保存（原文件未改动）：{Path(path).name}', 3200)


# ============================================================ ③ 模板库管理页
class QuestionsDialog(QDialog):
    """Every uncertain requirement becomes one explicit decision for the user."""

    def __init__(self, questions, parent=None):
        super().__init__(parent)
        self.setWindowTitle('待确认的模板要求')
        self.resize(720, 640)
        self.answers = {}
        self.groups = []
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 14)
        root.setSpacing(9)
        root.addWidget(label('待确认的模板要求', 'dialogtitle'))
        root.addWidget(label(f'识别到 {len(questions)} 处需要你决定的要求。默认项已选中，'
                             '改好后点“按我的选择继续”；取消则全部按默认处理。', 'aihint'))
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        box = QWidget()
        form = QVBoxLayout(box)
        form.setSpacing(9)
        for question in questions:
            card = QFrame()
            card.setObjectName('card')
            cl = QVBoxLayout(card)
            cl.setContentsMargins(12, 10, 12, 10)
            cl.setSpacing(5)
            cl.addWidget(label(question.get('text') or '', 'section'))
            cl.addWidget(label('位置：' + (question.get('where') or '未知'), 'muted'))
            group = QButtonGroup(card)
            for index, option in enumerate(question.get('options') or []):
                radio = QRadioButton(option)
                if index == question.get('default', 0):
                    radio.setChecked(True)
                group.addButton(radio, index)
                cl.addWidget(radio)
            self.groups.append((question, group))
            form.addWidget(card)
        form.addStretch(1)
        scroll.setWidget(box)
        root.addWidget(scroll, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                                   | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText('按我的选择继续')
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText('全部按默认')
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def accept(self):
        self.answers = {}
        for question, group in self.groups:
            checked = group.checkedButton()
            if checked is not None:
                self.answers[question['id']] = group.id(checked)
        super().accept()


class LibraryPage(QWidget):
    def __init__(self, store, workspace):
        super().__init__()
        self.store, self.workspace = store, workspace
        self.rows = []
        self.query = ''
        self.build_ui()

    def build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 18, 24, 14)
        root.setSpacing(10)
        head = QHBoxLayout()
        head.addWidget(label('模板库管理', 'title'))
        head.addStretch()
        self.count_label = label('', 'badge')
        head.addWidget(self.count_label)
        root.addLayout(head)
        root.addWidget(label('导入任意 .docx 会自动识别填写位；重命名/删除只影响自定义模板，内置模板始终可用。', 'muted'))
        self.search = QLineEdit()
        self.search.setObjectName('search')
        self.search.setPlaceholderText('筛选模板：输入名称 / 分类 / 标识')
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self.apply_filter)
        root.addWidget(self.search)
        topbar = QHBoxLayout()
        topbar.addWidget(button('导入 Word 文档…', self.import_template, primary=True))
        topbar.addWidget(button('重命名', self.rename))
        topbar.addWidget(button('导出模板包…', self.export_zip))
        topbar.addWidget(button('打开数据目录', lambda: open_in_explorer(self.store.root)))
        topbar.addStretch()
        self.delete_btn = button('删除模板', self.delete_selected, danger=True)
        topbar.addWidget(self.delete_btn)
        root.addLayout(topbar)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(['模板名称', '分类', '字段数', '来源', '草稿状态', '标识'])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(2, 70)
        self.table.setColumnWidth(3, 110)
        self.table.setColumnWidth(4, 130)
        self.table.setColumnWidth(5, 300)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.itemSelectionChanged.connect(self.update_buttons)
        root.addWidget(self.table, 1)
        self.hint = label('双击任意行可直接去填写。', 'muted')
        root.addWidget(self.hint)
        self.table.itemDoubleClicked.connect(self._open_fill)

    def reload(self):
        configs = self.store.templates()
        self.rows = configs
        drafts = self.store.root / 'drafts'
        self.table.setRowCount(0)
        for config in configs:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(config['title']))
            self.table.item(row, 0).setData(Qt.ItemDataRole.UserRole, config['id'])
            self.table.setItem(row, 1, QTableWidgetItem(config.get('category') or '未分类'))
            self.table.setItem(row, 2, QTableWidgetItem(str(len(config['fields']))))
            builtin = config['id'].startswith('builtin-')
            self.table.setItem(row, 3, QTableWidgetItem('内置' if builtin else '自定义'))
            draft = drafts / (config['id'] + '.json')
            if draft.is_file():
                try:
                    saved = read_json(draft).get('saved_at', '')[:16].replace('T', ' ')
                    state = f'草稿 {saved}'
                except Exception:
                    state = '草稿异常'
            else:
                state = '无草稿'
            self.table.setItem(row, 4, QTableWidgetItem(state))
            self.table.setItem(row, 5, QTableWidgetItem(config['id']))
        custom = sum(1 for c in configs if not c['id'].startswith('builtin-'))
        self.count_label.setText(f'{len(configs)} 个模板 · 自定义 {custom}')
        self.update_buttons()
        self.apply_filter()

    def matches(self, row_index):
        if not self.query:
            return True
        blob = ' '.join((self.table.item(row_index, 0).text(),
                         self.table.item(row_index, 1).text(),
                         self.table.item(row_index, 5).text())).lower()
        return self.query.lower() in blob

    def apply_filter(self):
        self.query = (self.search.text() or '').strip()
        shown = 0
        for row_index in range(self.table.rowCount()):
            visible = self.matches(row_index)
            self.table.setRowHidden(row_index, not visible)
            shown += 1 if visible else 0
        if self.query:
            self.count_label.setText(f'匹配 {shown} / {self.table.rowCount()} 个模板')
            self.hint.setText(f'筛选“{self.query}”：{shown} 个匹配，双击任意行去填写。')
        else:
            self.hint.setText('双击任意行可直接去填写。')

    def set_filter(self, text):
        text = (text or '').strip()
        if (self.search.text() or '').strip() == text:
            return
        self.search.blockSignals(True)
        self.search.setText(text)
        self.search.blockSignals(False)
        self.apply_filter()

    def select(self, template_id):
        for row_index in range(self.table.rowCount()):
            if self.table.item(row_index, 0).data(Qt.ItemDataRole.UserRole) == template_id:
                self.table.setCurrentCell(row_index, 0)
                self.table.selectRow(row_index)
                self.table.scrollToItem(self.table.item(row_index, 0))
                return

    def current_config(self):
        items = self.table.selectedItems()
        if not items:
            return None
        tid = self.table.item(items[0].row(), 0).data(Qt.ItemDataRole.UserRole)
        return next((c for c in self.rows if c['id'] == tid), None)

    def update_buttons(self):
        config = self.current_config()
        editable = bool(config) and not config['id'].startswith('builtin-')
        self.delete_btn.setEnabled(editable)

    def _open_fill(self, item):
        self.workspace.go_fill(self.table.item(item.row(), 0).data(Qt.ItemDataRole.UserRole))

    def import_template(self):
        path, _ = QFileDialog.getOpenFileName(self, '导入 Word 文档（自动识别填写位）', '',
                                              'Word 文档 (*.docx)')
        if not path:
            return
        tid = 'custom-' + uuid.uuid4().hex[:12]
        directory = self.store.custom / tid
        try:
            plan, config = directfill.prepare(path, directory)
            config['id'] = tid
            write_json(directory / 'config.json', config)
        except Exception as exc:
            shutil.rmtree(directory, ignore_errors=True)
            self.workspace.warn_box(f'导入失败：{exc}')
            return
        self.workspace.refresh_all()
        QMessageBox.information(
            self, '导入完成',
            f'模板「{config["title"]}」已加入模板库，自动识别 {len(plan["fields"])} 个填写项。')

    def rename(self):
        config = self.current_config()
        if not config or config['id'].startswith('builtin-'):
            return
        text, ok = QInputDialog_text(self, '重命名模板', '新名称：', config['title'])
        if not ok or not text.strip() or text.strip() == config['title']:
            return
        cfg_path = Path(config['_config_path'])
        data = read_json(cfg_path)
        data['title'] = text.strip()
        write_json(cfg_path, data)
        self.workspace.refresh_all()

    def delete_selected(self):
        config = self.current_config()
        if not config or config['id'].startswith('builtin-'):
            return
        answer = QMessageBox.question(
            self, '删除模板',
            f'确认删除模板「{config["title"]}」？\n\n'
            '模板文件与它的草稿会从本机数据目录移除，不可恢复。\n'
            '已生成过的 Word 文档不受影响。',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            shutil.rmtree(Path(config['_dir']))
            draft = self.store.root / 'drafts' / (config['id'] + '.json')
            if draft.is_file():
                draft.unlink()
            self.store.drop_format(config['id'])
        except OSError as exc:
            QMessageBox.warning(self, '删除失败', str(exc))
            return
        self.workspace.refresh_all()

    def export_zip(self):
        config = self.current_config()
        if not config:
            return
        default = str(Path.home() / 'Desktop' / f'{config["title"]}.zip')
        path, _ = QFileDialog.getSaveFileName(self, '导出模板包', default, 'ZIP 压缩包 (*.zip)')
        if not path:
            return
        src = Path(config['_dir'])
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
            if config['id'].startswith('builtin-'):
                payload = {k: v for k, v in config.items() if not k.startswith('_')}
                payload['id'] = 'custom-' + uuid.uuid4().hex[:12]
                payload['file'] = 'template.docx'
                zf.writestr('config.json', json.dumps(payload, ensure_ascii=False, indent=2))
                zf.write(src / config['file'], 'template.docx')
                for page in payload.get('previews', []):
                    if (src / page).is_file():
                        zf.write(src / page, page)
            else:
                for member in sorted(src.iterdir()):
                    zf.write(member, member.name)
        QMessageBox.information(self, '已导出',
                                f'模板包已保存：\n{path}\n\n'
                                '迁移方法：关闭程序后，把 zip 解压为数据目录 templates 下的一个文件夹即可。')


class LibraryDialog(QDialog):
    """模板库不再占左侧导航，只从首页卡片进来；管理功能一个不少。"""

    def __init__(self, workspace, parent=None):
        super().__init__(parent)
        self.workspace = workspace
        self.setWindowTitle('模板库')
        self.resize(1080, 720)
        box = QVBoxLayout(self)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(0)
        box.addWidget(workspace.library, 1)
        footer = QHBoxLayout()
        footer.setContentsMargins(24, 0, 24, 14)
        footer.setSpacing(8)
        footer.addWidget(button('结构体检（在模板制作页）', self.go_inspect))
        footer.addStretch()
        footer.addWidget(button('关闭', self.close))
        box.addLayout(footer)

    def select(self, template_id):
        self.workspace.library.select(template_id)

    def go_inspect(self):
        self.close()
        self.workspace.nav.setCurrentRow(Workspace.STUDIO)
        self.workspace.studio.tabs.setCurrentWidget(self.workspace.studio.inspect)


def QInputDialog_text(parent, title, prompt, initial):
    from PySide6.QtWidgets import QInputDialog
    text, ok = QInputDialog.getText(parent, title, prompt, QLineEdit.EchoMode.Normal, initial)
    return text, ok


class ImportDialog(QDialog):
    def __init__(self, keys, title, parent):
        super().__init__(parent)
        self.setWindowTitle('配置新模板')
        self.resize(650, 510)
        layout = QVBoxLayout(self)
        layout.addWidget(label('设置模板名称和填写方式。相同名称的标记会使用同一个填写值。', 'muted'))
        self.name = QLineEdit(title)
        self.name.setMaxLength(80)
        form = QFormLayout()
        form.addRow('模板名称', self.name)
        layout.addLayout(form)
        self.table = QTableWidget(len(keys), 3)
        self.table.setHorizontalHeaderLabels(['填写项目名称', '输入方式', '必填'])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(2, 60)
        self.keys = keys
        for i, key in enumerate(keys):
            self.table.setItem(i, 0, QTableWidgetItem(key))
            combo = QComboBox()
            combo.addItem('单行文字', 'text')
            combo.addItem('多行正文', 'multiline')
            if is_multiline_label(key):
                combo.setCurrentIndex(1)
            self.table.setCellWidget(i, 1, combo)
            check = QTableWidgetItem()
            check.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            check.setCheckState(Qt.CheckState.Unchecked)
            self.table.setItem(i, 2, check)
        layout.addWidget(self.table)
        actions = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                                   | QDialogButtonBox.StandardButton.Cancel)
        actions.button(QDialogButtonBox.StandardButton.Ok).setText('导入模板')
        actions.button(QDialogButtonBox.StandardButton.Cancel).setText('取消')
        actions.accepted.connect(self.accept)
        actions.rejected.connect(self.reject)
        layout.addWidget(actions)

    def fields(self):
        return [dict(key=key, label=self.table.item(i, 0).text().strip() or key,
                     type=self.table.cellWidget(i, 1).currentData(),
                     required=self.table.item(i, 2).checkState() == Qt.CheckState.Checked,
                     group='填写内容', max_length=0, default='') for i, key in enumerate(self.keys)]


# ============================================================ ④ 结构体检页
class InspectPage(QWidget):
    def __init__(self, store, workspace):
        super().__init__()
        self.workspace = workspace
        self.last = None
        self.build_ui()

    def build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 18, 24, 14)
        root.setSpacing(10)
        head = QHBoxLayout()
        head.addWidget(label('结构体检（只读）', 'title'))
        head.addStretch()
        root.addLayout(head)
        root.addWidget(label('不改写任何文件：统计段落、表格、合并单元格、页眉页脚、内容控件和旧版占位符，'
                             '并解释导入/生成会遇到的结构问题。新文档不需要任何占位符。', 'muted'))
        bar = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText('选择要体检的 .docx…')
        bar.addWidget(self.file_edit, 1)
        bar.addWidget(button('浏览…', self.pick))
        self.run_btn = button('开始体检', self.run, primary=True)
        self.run_btn.setEnabled(False)
        bar.addWidget(self.run_btn)
        bar.addWidget(button('保存报告…', self.save_report))
        root.addLayout(bar)
        self.out = QPlainTextEdit()
        self.out.setReadOnly(True)
        self.out.setPlaceholderText('体检结果将显示在这里。')
        root.addWidget(self.out, 1)

    def pick(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择 Word 文档', '', 'Word 文档 (*.docx)')
        if path:
            self.file_edit.setText(path)
            self.run_btn.setEnabled(True)

    def run(self):
        path = Path(self.file_edit.text())
        try:
            recon = pipeline.phase_recon(str(path))
        except Exception as exc:
            self.out.setPlainText(f'体检失败：{exc}')
            return
        self.last = recon
        lines = [f'文件：{path}', '']
        lines.append(f'段落 {recon["paragraph_count"]} · 表格 {recon["table_count"]} · '
                     f'节 {recon["section_count"]} · 页眉页脚 {"有" if recon["has_header_footer"] else "无"}')
        lines.append(f'内容控件(w:sdt) {recon["content_controls"]} · 表单域(w:fldChar) {recon["form_fields"]}')
        if recon['content_controls'] or recon['form_fields']:
            lines.append('提示：内容控件/表单域会被自动识别为填写位，可直接在工具里填写并生成。')
        lines.append('')
        lines.append('—— 表格 ——')
        for t in recon['tables']:
            lines.append(f'表{t["index"] + 1}：{t["rows"]}行×{t["cols"]}列，唯一单元格 {t["unique_tc"]}，'
                         f'合并区 {len(t["merged_regions"])}')
            for mr in t['merged_regions'][:10]:
                lines.append(f'   合并 anchor=行{mr["anchor"][0] + 1}列{mr["anchor"][1] + 1} '
                             f'共享 {len(mr["occurrences"])} 格 vMerge={mr["v_merge"]} '
                             f'gridSpan={mr["grid_span"]}')
        colon = [p for p in recon['paragraphs'] if p['ends_with_colon']]
        empties = [p for p in recon['paragraphs'] if p['empty']]
        fmt_only = [p for p in recon['paragraphs'] if p['format_only_runs']]
        lines.append('')
        lines.append('—— 段落 ——')
        lines.append(f'冒号标签段 {len(colon)} 个；空段 {len(empties)} 个；仅格式无文字 run 的段 {len(fmt_only)} 个')
        for p in colon[:12]:
            lines.append(f'   P{p["index"] + 1}: {p["text"][:40]}')
        lines.append('')
        lines.append(f'—— 旧版占位符（兼容读取） {len(recon["markers"])} 处 ——')
        for m in recon['markers']:
            loc = m['target']
            where = (f"表{loc.get('table', 0) + 1} 行{loc.get('row', 0) + 1} 列{loc.get('col', 0) + 1}"
                     if loc['kind'] == 'cell' else f"段落 {loc['p'] + 1}")
            flags = []
            if m['cross_run']:
                flags.append('跨run')
            if m['runs'] > 2:
                flags.append(f"{m['runs']}run")
            lines.append(f'   [旧占位符] {m["key"]} @ {where} ' + ' '.join(flags))
        if not recon['markers']:
            lines.append('   （无）——新模板不需要占位符，可直接到「自动识别」页识别并录入。')
        lines.append('')
        lines.append('—— OOXML 自检（结构合法性与 Word 取值范围）——')
        try:
            issues = ooxmlcheck.check(path)
        except Exception as exc:
            lines.append(f'   自检没跑成：{str(exc)[:120]}')
        else:
            lines.extend(f'   {line}' for line in ooxmlcheck.summarize(issues))
        self.out.setPlainText('\n'.join(lines))

    def save_report(self):
        if not self.last:
            QMessageBox.information(self, '无报告', '请先完成一次体检。')
            return
        default = str(Path(self.last['file']).parent / (Path(self.last['file']).stem + '_体检报告.json'))
        path, _ = QFileDialog.getSaveFileName(self, '保存体检报告 JSON', default, 'JSON (*.json)')
        if path:
            Path(path).write_text(json.dumps(self.last, ensure_ascii=False, indent=2), encoding='utf-8')
            self.workspace.notify('体检报告已保存：' + Path(path).name, 'ok', 3000)


# ============================================================ 使用说明
class HelpDialog(QDialog):
    def __init__(self, store, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Word 模板工作台 · 使用说明')
        self.resize(760, 640)
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 16)
        root.addWidget(label('使用说明', 'dialogtitle'))
        text = QPlainTextEdit()
        text.setReadOnly(True)
        text.setPlainText(
            '① 填写生成\n'
            '两种方式：\n'
            '  · 直接拖入任意 .docx，自动识别填写位后立即填写，不需要先做模板；需要长期使用时点“存入模板库”。\n'
            '  · 用填写页顶部的下拉框切换已保存的模板，旁边可以按名称筛选。\n'
            '填写页分三栏：左边“排版规则 · 模板要求”列出这次生成会照做的格式、搬走的说明和内容要求；'
            '中间“文档预览”看整页效果，底部可翻页和缩放；右边“填写内容”是表单。'
            '按分组填写，默认不强制全部填写；没填的字段保留模板原内容或原示例。'
            '顶部不再有“视觉识别”开关：识别改由“结构 + 文本模型”完成，'
            '不再把整页截图丢给模型。文本模型配置可用时会自动参与，未配置时只做结构识别。'
            '需要某项必填可在模板制作页勾选。表格双击单元格；'
            '草稿自动保存（Ctrl+S 也可）；'
            '“生成 Word”另存 .docx 后可“打开生成文件”核对版式。生成文档是普通 Word，可以继续编辑。'
            '浮动文本框里的填写说明会先读取、套用到它指向的栏目，然后连文本框一起删掉；'
            '无论模板用的是新版 wps:txbx 还是旧式 VML 文本框，生成/导入后的文件里一律不保留；'
            '突出显示（黄色等标记）也会全部清掉。'
            '模板要求里的字号会同步到右侧输入框（例如“小四”就按 12pt 显示），'
            '生成时再把字体、字号、行距、缩进真正套到填写内容上。'
            '“（一）选题的意义：”这类标签后面填多行内容时，该段会固定为左对齐，'
            '不会因为两端对齐把标签字符拉开。\n\n'
            '② 模板制作（自动识别 + 文本模型复核 + 按要求自动排版 + 结构体检）\n'
            '选原始 .docx → 点“开始识别”。自动读取：\n'
            '  （1）表格：标签/值单元格、整行大段、合并单元格、多表格文档；\n'
            '  （2）段落：冒号标签后的空白、成排下划线、括号占位、封面填空、日期；\n'
            '  （3）结构化填写位：Word 内容控件、旧式表单域、页眉页脚；\n'
            '  （4）已有内容的模板：示例答案/示例正文按栏目范围识别。\n'
            '识别完成后，若 AI 设置里的文本模型可用，程序会把整份文档的结构清单'
            '（段落、表格、空单元格、下划线、占位符）交给文本模型通读，'
            '补上结构漏掉的栏目、校正字段名和单行/多行类型；模型给出的位置必须能在原文里对上，'
            '对不上的直接丢弃。文本模型未配置时只做结构识别，功能不受影响。\n'
            '程序还会自动读取模板自己写明的要求（如“中文字全部采用宋体、小四号字、加粗”'
            '“行间距：固定值20磅”“悬挂缩进2字符”“此处空一行”“参考文献不少于10篇”），'
            '在生成的 Word 里照做：字体字号加粗对齐套到填写内容，行距缩进空行套到对应栏目段落。'
            '这一步完全读文档自己的文字，不联网也能用。'
            '要求不要求彩色或底纹：普通黑字里写的“中文用宋体、小四号字”也会读取，'
            '但这类黑字只套格式、不删原文；彩色或带底纹的说明仍按原逻辑搬走删除。'
            '填写页右侧输入框的字号也会跟着规则变，所见即所得。\n'
            '彩色/带底纹的说明文字会从生成的 Word 中删掉，改为显示在填写页对应栏位的下方（“模板要求：……”），'
            '填写时照样看得到。示例正文只删紧跟其后的彩色说明部分，示例本身保留；'
            '标题里的“（请注意此处有变化）”这类括注也只删括注。此行为始终开启，无需勾选。\n'
            '要求互相矛盾、或某条要求不知道管哪一段时，会弹窗逐条问你，按你的选择重算。\n'
            '“小模型补全空标记名称”是 v3.2 的老流程：文档里已有空的 {{ }} 区域时，'
            '用内置离线小模型按上下文补字段名，再自动识别录入；不联网。\n'
            '不需要在 Word 里写任何占位符。识别只在原文档中加入不可见的 Word 书签或定位到现有控件，'
            '原文、示例答案、说明、签名/日期、颜色、字体、缩进、行高和批注框全部保留，页数与原件一致。'
            '生成 Word 时才替换勾选的填写位。\n'
            '识别结果会列在表格里并默认勾选；双击“填写项目名称”即可改名，确认后点“确认改名并录入模板库”。'
            '原文件永不修改。\n'
            '结构体检：只读扫描表格/合并区/冒号标签/内容控件/表单域，导入前定位结构问题。\n\n'
            '③ 模板库（首页卡片入口，弹窗管理）\n'
            '模板库不再占左侧导航：首页的「模板库」卡片、以及每张模板卡片右下角的“模板库”，'
            '打开的是同一个弹窗，功能不变——“导入 Word 文档”自动识别并加入模板库，'
            '可重命名、删除自定义模板、导出/迁移模板包（zip）、打开数据目录，还能一键跳到结构体检。'
            '内置模板不可删除；删除自定义模板时，其草稿和字号设置一并移除。\n\n'
            '④ AI 识别与分类（模板制作页）\n'
            '先点“开始识别”得到候选，文本 AI 再规范字段名、判断单行/多行并分组，'
            '模板名与分类也会自动建议。识别时程序会自动读取模板自己写明的字体、字号、行距、缩进和'
            '内容要求，并把说明整理成填写提示——这一步始终开启，无需勾选。'
            'AI 只处理你点过识别的本地文档，点击按钮才会联网或调用本机模型，'
            '原 Word 文件本身不会上传。\n\n'
            '⑤ AI 生成正文（填写页）\n'
            '每个多行正文字段右上角有“AI 生成”。可填写写作要求、选择目标字数，并可选择在现有内容后追加。'
            'AI 会带上模板说明和其他已填字段作为上下文；生成结果先预览，确认后才写入填写框，超模板字数会提示。\n\n'
            '⑥ 正文字号与格式\n'
            '填写页不再有“格式 / 字号”按钮和“手动设置”区块：字号、字体、行距、缩进直接读模板自己'
            '写明的要求（例如“中文字全部采用宋体、小四号字、加粗”“行间距固定值 20 磅”），'
            '生成时照做，说明同时搬成填写提示。想手工设定，用「自动排版」页：正文 / 标题 1-3 / '
            '表格内文字 / 页眉页脚分档给参数，还可设段前段后间距与颜色，并可把这一档存成自己的样式。'
            '格式只作用于填写内容，模板固定文字保持原样；老模板遗留的手动格式仍然生效并按模板存在本机。\n\n'
            '⑦ AI 服务设置（顶栏“AI 设置”）\n'
            '分“文本模型”和“视觉模型”两页，可分别使用不同服务商。\n'
            'API 协议可选：本机 Ollama、OpenAI 兼容 API（豆包/DeepSeek/通义千问/任意兼容服务）、'
            'Anthropic Messages（claude 或走这类中转站）以及「内置小模型」。\n'
            '选 Anthropic Messages 时：API 地址填到 /v1（例如 https://api.xxx.com/v1），'
            '程序请求 /v1/messages，带 x-api-key 与 anthropic-version 头，system 单独成字段、'
            'max_tokens 必填；模型名写中转站实际接受的 ID，“检测并获取模型”会请求 GET /v1/models，'
            '列不出来就手填再用“测试对话”确认。视觉页选同一协议时，图片按 base64 image block 送出，'
            '需要模型支持图片输入；与文本同一个服务时，视觉的地址和密钥都可以留空复用上面的配置。\n'
            '文本模型支持本机 Ollama、豆包（火山方舟）、DeepSeek、通义千问、任意 OpenAI 兼容服务，'
            '以及「内置小模型」——随包附带的离线 Qwen3-0.6B（llama.cpp 纯 CPU），'
            '选它不要 Key、不联网、不上传任何文档内容，可做字段名规范化、分组建议和短正文起草；'
            '但 0.6B 语义能力有限，长正文和复杂分类建议用在线大模型或本机 Ollama 大档，'
            '首次调用需加载模型、可能等几十秒。\n'
            '想完全免费离线：装本机 Ollama 后拉 qwen3:4b-instruct（文本）+ qwen3-vl:4b（视觉），'
            '两页都选“本机 Ollama”即可。服务没跑时不用手动开：工具在调用 AI 前会自己找到 '
            'ollama.exe 并把它拉起，侧栏会显示“本机 Ollama：运行中 / 正在自动启动”。'
            '文本模型请用非思考版：Ollama 上的 qwen3 思考版会把正文写成“首先用户要求我……”'
            '这类独白，选错时 AI 设置里会直接提醒。'
            '注意模型目录不要放在含中文的用户名路径下'
            '（把环境变量 OLLAMA_MODELS 指到如 D:\\ollama\\models 这样的纯英文路径）。'
            '视觉模型默认使用开源权重的 Qwen3-VL，预置硅基流动、阿里云百炼、OpenRouter 和本机 Ollama；'
            '也可以用任意 OpenAI 兼容视觉服务。它现在只做“少量截图复核”：'
            '只有当文本模型和结构识别的栏目名/类型对不上时，才把涉及的那 1–3 页截图发过去核对；'
            '不再整页整页地识别，也不发送原始 .docx。API Key 仅保存在本机数据目录。\n'
            '识别填写位时，文本模型会收到整份文档的文字与表格结构（不含页面图片）；'
            '模型返回的栏目必须能在原文里定位到，否则会被丢弃。不需要模型时清空文本模型即可退回纯结构识别。\n'
            '配置后先点“测试对话/测试视觉”验证，再保存。生成慢或超时可调大超时时间；长正文可调大输出上限。\n\n'
            '⑧ 实时预览\n'
            '填写页中间“文档预览”会用本机 Word/WPS 按当前填写内容渲染整页版式，'
            '停止输入约 1 秒后自动刷新；首次启动 Word 约需几秒，之后通常 1～2 秒更新。'
            '底部可翻页、放大缩小或勾选“适应宽度”；可取消“实时渲染”回退原始模板参考图。'
            '预览会先记录模板本身的页数；当你填进去的内容把文档撑到下一页时，'
            '右下角会浮层提醒一次，不会打断输入。'
            '本功能需要本机安装 Word 或 WPS，不需要联网。\n\n'
            '⑨ 兼容与图片\n'
            '旧版已经录入的模板（含旧式占位符或图片替代文字）仍可正常填写和生成；新文档不再需要任何标记。'
            '图片字段沿用原图位置和尺寸；旧模板里的图片替换继续可用。'
            '旧 .doc 请先在 Word 中另存为 .docx。\n\n'
            '⑩ 界面、快捷键与拖放\n'
            '顶栏：模板搜索（回车跳到模板库）、白天/黑夜主题切换、AI 设置、使用说明。'
            '左侧导航五项：首页（模板卡片墙，点整张卡片就去填写）、填写生成、模板制作、自动排版、文献检索；'
            '模板库改由首页卡片以弹窗打开，功能不减。主题选择记在本机数据目录里，下次打开仍然是上次的配色；'
            '两套主题只改界面配色，不影响生成出来的 Word 文档。\n'
            '快捷键：Alt+1…Alt+5 切页；Ctrl+F 聚焦顶栏搜索（Esc 先清空关键词，已经空了就交还焦点）；'
            'Ctrl+S 保存草稿；Ctrl+Enter 在填写页直接生成；F1 打开使用说明；弹窗用 Esc 关闭。'
            '首页卡片墙可以纯键盘走：Tab 进入，方向键在卡片间移动，回车打开。\n'
            '拖放：首页、填写生成、模板制作、自动排版四个页面都接收 .docx，拖入时整页高亮，'
            '放手就开始识别或载入；拖进来的不是 .docx 会立刻说明怎么办。\n'
            '长任务（批量生成、联网检索、增强识别、应用排版）会出现一行进度与“取消”，'
            '取消只停在未来那一份/那一批上，已经写出的文件不回滚，原文件始终不动。\n'
            '成功类反馈（生成完成、草稿已保存、已录入模板库、题录已复制）走右下角浮层，几秒自动消失；'
            '需要过目的内容（必填为空、结构自检问题、内容要求提醒）仍然弹窗。\n'
            '背景：五个页面都铺同一张静态玻璃背景（顶栏与左侧导航半透明，透出蓝色流体磨砂玻璃质感），'
            '由 GPU 着色器只渲一次后缓存成图片，切页和滚动不再重算，所以不占显卡；'
            '本机没有可用 OpenGL 时自动改用等效静态渐变，不会黑屏。'
            '只有启动页那 1.2 秒是逐帧动效（点一下或按任意键可提前进入），'
            '命令行加 --no-splash 可跳过启动页。\n'
            '顶栏与预览条上的小按钮改成图标，鼠标停一下就能看到它是什么。\n'
            '图标与标志已换成泼墨版：顶栏左侧、启动页中央、以及 exe/任务栏应用图标（assets/icon.ico）。'
            '旧图标保留在 assets/icon-legacy.svg 与 icon-legacy.ico。\n\n'
            '⑪ 自动排版工作台（整篇文档重排）\n'
            '选一个 .docx（也可以直接拖进来），左侧按档设置排版规则：'
            '页面设置（纸张 A3/A4/A5/B5/Letter、横向、上下左右页边距，写到文档每一节）、'
            '正文、一级/二级/三级标题、表格内文字、页眉页脚。'
            '每档自带“作用范围”：不改 / 整篇文档 / 仅填写内容'
            '（“仅填写内容”只动本工具填写位书签里的文字，纯排版文档请用“整篇文档”）。'
            '标题按 Word 样式名或大纲级别判定，认不出来的按正文处理。\n'
            '右侧样式库是预设参数：点“用到左栏”就把这套参数灌进对应档，仍可继续改。\n'
            '点“应用排版到文档”后在中栏看整页效果（可翻页、缩放），'
            '确认后再点“导出 Word”另存新文件；原文件始终不会被改动。'
            '撤销/重做在历次排版结果之间切换，回到第 0 步即未排版状态。\n'
            '整页预览需要本机 Word/WPS 或 LibreOffice；没有也能应用并导出，只是看不到效果图。\n\n'
            f'本机数据目录：{store.root}\n'
            '与旧版共用同一数据目录：已导入模板、草稿和字号设置直接沿用。'
            'EXE 内置运行环境和首批模板，不需要安装 Python。')
        root.addWidget(text, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.button(QDialogButtonBox.StandardButton.Close).setText('关闭')
        buttons.rejected.connect(self.accept)
        root.addWidget(buttons)


class AISettingsDialog(QDialog):
    def __init__(self, cfg, parent=None):
        super().__init__(parent)
        self.setWindowTitle('AI 设置')
        self.resize(720, 690)
        self.cfg = dict(cfg or ai.default_config())
        self._worker = None
        self._loading = True
        self._status_widget = None
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 16)
        root.setSpacing(10)
        root.addWidget(label('AI 服务设置', 'dialogtitle'))
        root.addWidget(label('文本模型负责识别填写位、字段分类和正文生成；'
                             '视觉模型只在少数栏目有分歧时看截图复核。'
                             '所有 AI 调用都需要你点击按钮才会发生，API Key 只保存在本机。', 'aihint'))
        self.tabs = QTabWidget()

        text_tab = QWidget()
        text_root = QVBoxLayout(text_tab)
        text_root.setContentsMargins(4, 8, 4, 4)
        text_root.setSpacing(9)
        form = QFormLayout()
        form.setSpacing(9)
        self.preset = QComboBox()
        self.preset.addItem('豆包（火山方舟）', 'doubao')
        self.preset.addItem('DeepSeek', 'deepseek')
        self.preset.addItem('通义千问（DashScope）', 'qwen')
        self.preset.addItem('本机 Ollama', 'ollama')
        self.preset.addItem('Anthropic 中转站（claude）', 'anthropic')
        self.preset.addItem('内置小模型（离线，随包附带）', 'embedded')
        self.preset.addItem('自定义 / 当前配置', 'custom')
        form.addRow('服务预设', self.preset)
        self.backend = QComboBox()
        self.backend.addItem('本机 Ollama（离线）', 'ollama')
        self.backend.addItem('OpenAI 兼容 API（云端/局域网）', 'openai')
        self.backend.addItem('Anthropic Messages（claude / 中转站）', 'anthropic')
        self.backend.addItem('内置 Qwen3-0.6B（离线，质量有限）', 'embedded')
        idx = self.backend.findData(self.cfg.get('backend', 'ollama'))
        self.backend.setCurrentIndex(max(0, idx))
        form.addRow('API 协议 / 服务类型', self.backend)
        self.base = QLineEdit(self.cfg.get('base_url') or ai.OLLAMA_DEFAULT)
        self.base.setPlaceholderText('Ollama: http://127.0.0.1:11434；'
                                     'Anthropic: 填到 /v1，例如 https://api.xxx.com/v1')
        form.addRow('API 地址', self.base)
        model_row = QWidget()
        mr = QHBoxLayout(model_row)
        mr.setContentsMargins(0, 0, 0, 0)
        mr.setSpacing(8)
        self.model = QComboBox()
        self.model.setEditable(True)
        current_model = (self.cfg.get('model') or '').strip()
        if current_model:
            self.model.addItem(current_model)
        self.model.setCurrentText(current_model)
        self.models_btn = button('检测并获取模型', self.fetch_models)
        self.test_btn = button('测试对话', self.test_chat)
        mr.addWidget(self.model, 1)
        mr.addWidget(self.models_btn)
        mr.addWidget(self.test_btn)
        form.addRow('模型', model_row)
        self.model_warn = label('', 'aihint')
        self.model_warn.setWordWrap(True)
        form.addRow('', self.model_warn)
        self.model.currentTextChanged.connect(lambda *_: self.update_model_warning())
        self.key = QLineEdit(self.cfg.get('api_key') or '')
        self.key.setEchoMode(QLineEdit.EchoMode.Password)
        self.key.setPlaceholderText('云端服务通常需要；Ollama 留空。仅保存在本机数据目录')
        form.addRow('API 密钥', self.key)
        self.temperature = QDoubleSpinBox()
        self.temperature.setRange(0.0, 1.5)
        self.temperature.setSingleStep(0.1)
        self.temperature.setDecimals(1)
        self.temperature.setValue(float(self.cfg.get('temperature', 0.2) or 0.2))
        form.addRow('温度（越低越稳定）', self.temperature)
        self.timeout = QSpinBox()
        self.timeout.setRange(10, 600)
        self.timeout.setSuffix(' 秒')
        self.timeout.setValue(int(self.cfg.get('timeout', 120) or 120))
        form.addRow('超时时间', self.timeout)
        self.max_tokens = QSpinBox()
        self.max_tokens.setRange(256, 16384)
        self.max_tokens.setSingleStep(256)
        self.max_tokens.setValue(int(self.cfg.get('max_tokens', 2048) or 2048))
        form.addRow('单次输出上限', self.max_tokens)
        text_root.addLayout(form)
        self.status = label('', 'aihint')
        text_root.addWidget(self.status)
        self.hint = label('', 'aihint')
        self.hint.setWordWrap(True)
        text_root.addWidget(self.hint)
        text_root.addStretch(1)
        self.tabs.addTab(text_tab, '文本模型')

        vision_tab = QWidget()
        vision_root = QVBoxLayout(vision_tab)
        vision_root.setContentsMargins(4, 8, 4, 4)
        vision_root.setSpacing(9)
        vision_root.addWidget(label(
            '视觉模型只在栏目名有分歧时复核：只发送涉及争议栏目的 1–3 页截图，不发送原始 Word 文件，'
            '也不再用截图做整篇识别。默认使用开源权重的 Qwen3-VL，'
            '推荐在硅基流动或阿里云百炼申请对应服务的 Key。', 'aihint'))
        vform = QFormLayout()
        vform.setSpacing(9)
        self.vpreset = QComboBox()
        for label_text, key, _backend, _base, _model in ai.VISION_PRESETS:
            self.vpreset.addItem(label_text, key)
        vform.addRow('视觉服务预设', self.vpreset)
        self.vbackend = QComboBox()
        self.vbackend.addItem('OpenAI 兼容 API（联网）', 'openai')
        self.vbackend.addItem('Anthropic Messages（claude 视觉 / 中转站）', 'anthropic')
        self.vbackend.addItem('本机 Ollama（离线）', 'ollama')
        vidx = self.vbackend.findData(self.cfg.get('vision_backend', 'openai'))
        self.vbackend.setCurrentIndex(max(0, vidx))
        vform.addRow('API 协议 / 服务类型', self.vbackend)
        self.vbase = QLineEdit(self.cfg.get('vision_base_url') or ai.SILICONFLOW_DEFAULT)
        self.vbase.setPlaceholderText('例如 https://api.siliconflow.cn/v1；与文本同一服务时可留空')
        vform.addRow('API 地址', self.vbase)
        vmodel_row = QWidget()
        vmr = QHBoxLayout(vmodel_row)
        vmr.setContentsMargins(0, 0, 0, 0)
        vmr.setSpacing(8)
        self.vmodel = QComboBox()
        self.vmodel.setEditable(True)
        vcurrent = (self.cfg.get('vision_model') or ai.QWEN_VL_MODEL).strip()
        if vcurrent:
            self.vmodel.addItem(vcurrent)
        self.vmodel.setCurrentText(vcurrent)
        self.vmodels_btn = button('检测模型', self.fetch_vision_models)
        self.vtest_btn = button('测试视觉', self.test_vision)
        vmr.addWidget(self.vmodel, 1)
        vmr.addWidget(self.vmodels_btn)
        vmr.addWidget(self.vtest_btn)
        vform.addRow('视觉模型', vmodel_row)
        self.vkey = QLineEdit(self.cfg.get('vision_api_key') or '')
        self.vkey.setEchoMode(QLineEdit.EchoMode.Password)
        self.vkey.setPlaceholderText('与文本模型同一个服务时留空即可复用上面的密钥')
        vform.addRow('视觉 API 密钥', self.vkey)
        self.vmax_pages = QSpinBox()
        self.vmax_pages.setRange(1, 30)
        self.vmax_pages.setSuffix(' 页')
        self.vmax_pages.setValue(int(self.cfg.get('vision_max_pages', 12) or 12))
        vform.addRow('最多识别页数', self.vmax_pages)
        vision_root.addLayout(vform)
        self.vstatus = label('', 'aihint')
        vision_root.addWidget(self.vstatus)
        self.vhint = label('', 'aihint')
        self.vhint.setWordWrap(True)
        vision_root.addWidget(self.vhint)
        vision_root.addStretch(1)
        self.tabs.addTab(vision_tab, '视觉模型（开源 VLM）')
        root.addWidget(self.tabs, 1)

        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                                        | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setText('保存')
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText('取消')
        self.buttons.accepted.connect(self.save)
        self.buttons.rejected.connect(self.reject)
        root.addWidget(self.buttons)
        self.backend.currentIndexChanged.connect(self.backend_changed)
        self.backend_changed()
        base_lower = self.base.text().strip().lower()
        if self.backend.currentData() == 'embedded':
            preset = 'embedded'
        elif self.backend.currentData() == 'ollama':
            preset = 'ollama'
        elif 'ark.cn-beijing.volces.com' in base_lower:
            preset = 'doubao'
        elif 'deepseek' in base_lower:
            preset = 'deepseek'
        elif 'dashscope' in base_lower:
            preset = 'qwen'
        else:
            preset = 'custom'
        self.preset.blockSignals(True)
        self.preset.setCurrentIndex(max(0, self.preset.findData(preset)))
        self.preset.blockSignals(False)
        self.vpreset.blockSignals(True)
        self.vpreset.setCurrentIndex(
            max(0, self.vpreset.findData(self._detect_vision_preset())))
        self.vpreset.blockSignals(False)
        self._loading = False
        self.update_model_warning()
        self.preset.currentIndexChanged.connect(self.preset_changed)
        self.vpreset.currentIndexChanged.connect(self.vision_preset_changed)
        self.vbackend.currentIndexChanged.connect(self.vision_backend_changed)
        self.update_preset_hint()
        self.vision_backend_changed()
        self._status_widget = self.status

    def _detect_vision_preset(self):
        base = (self.cfg.get('vision_base_url') or '').lower()
        backend = self.cfg.get('vision_backend') or 'openai'
        if backend == 'ollama':
            return 'ollama'
        if 'siliconflow' in base:
            return 'siliconflow'
        if 'dashscope' in base:
            return 'dashscope'
        if 'openrouter' in base:
            return 'openrouter'
        return 'custom'

    def current_cfg(self):
        backend = self.backend.currentData()
        if backend == 'embedded':
            default = ''
        elif backend == 'ollama':
            default = ai.OLLAMA_DEFAULT
        elif backend == ai.ANTHROPIC_BACKEND:
            default = ai.ANTHROPIC_DEFAULT
        else:
            default = ai.OPENAI_DEFAULT
        cfg = dict(self.cfg)
        cfg.update({'backend': backend,
                    'base_url': self.base.text().strip() or default,
                    'model': self.model.currentText().strip(),
                    'api_key': self.key.text().strip(),
                    'temperature': round(self.temperature.value(), 2),
                    'timeout': int(self.timeout.value()),
                    'max_tokens': int(self.max_tokens.value()),
                    'vision_backend': self.vbackend.currentData(),
                    'vision_base_url': self.vbase.text().strip(),
                    'vision_model': self.vmodel.currentText().strip(),
                    'vision_api_key': self.vkey.text().strip(),
                    'vision_max_pages': int(self.vmax_pages.value())})
        return cfg

    def backend_changed(self, *_):
        backend = self.backend.currentData()
        base = self.base.text().strip()
        if backend == 'embedded':
            self.base.setText('')
            self.model.setCurrentText('')
            for widget in (self.base, self.key, self.model):
                widget.setEnabled(False)
            if hasattr(self, 'hint'):
                self.update_preset_hint()
            return
        for widget in (self.base, self.key, self.model):
            widget.setEnabled(True)
        if backend == 'ollama':
            if base in ('', ai.OPENAI_DEFAULT, ai.ANTHROPIC_DEFAULT):
                self.base.setText(ai.OLLAMA_DEFAULT)
            self.key.setEnabled(False)
        else:
            wanted = (ai.ANTHROPIC_DEFAULT if backend == ai.ANTHROPIC_BACKEND
                      else ai.OPENAI_DEFAULT)
            stale = ai.OPENAI_DEFAULT if wanted == ai.ANTHROPIC_DEFAULT else ai.ANTHROPIC_DEFAULT
            if base in ('', ai.OLLAMA_DEFAULT, stale):
                self.base.setText(wanted)
            self.key.setEnabled(True)
        if hasattr(self, 'hint'):
            self.update_preset_hint()

    def preset_changed(self, *_):
        if self._loading:
            return
        preset = self.preset.currentData()
        if preset == 'doubao':
            backend, base, model = 'openai', ai.DOUBAO_DEFAULT, ai.DOUBAO_MODEL
        elif preset == 'deepseek':
            backend, base, model = 'openai', ai.OPENAI_DEFAULT, 'deepseek-chat'
        elif preset == 'qwen':
            backend, base, model = 'openai', 'https://dashscope.aliyuncs.com/compatible-mode/v1', 'qwen-plus'
        elif preset == 'ollama':
            backend, base, model = 'ollama', ai.OLLAMA_DEFAULT, ''
        elif preset == 'anthropic':
            backend, base, model = ai.ANTHROPIC_BACKEND, ai.ANTHROPIC_DEFAULT, ''
        elif preset == 'embedded':
            backend, base, model = 'embedded', '', ''
        else:
            self.update_preset_hint()
            return
        idx = self.backend.findData(backend)
        self.backend.setCurrentIndex(max(0, idx))
        self.base.setText(base)
        self.model.setCurrentText(model)
        self.backend_changed()
        self.update_preset_hint()

    def update_model_warning(self):
        """实测踩过的坑：Ollama 上的 qwen3 思考版会把正文写成思考独白，提前提醒。"""
        model = (self.model.currentText() or '').strip().lower()
        thinking = (re.match(r'qwen3(?::|[-/][0-9])', model) or 'r1' in model
                    or 'thinking' in model)
        if self.backend.currentData() == 'ollama' and thinking and 'instruct' not in model:
            self.model_warn.setText(
                '注意：这是思考型模型。Ollama 上 qwen3 思考版会把正文写成“首先用户要求我……”'
                '这类独白，正文生成和 JSON 解析都会受影响。建议改用非思考版，'
                '例如 qwen3:4b-instruct 或 qwen2.5:7b。')
            self.model_warn.show()
        else:
            self.model_warn.clear()
            self.model_warn.hide()

    def update_preset_hint(self):
        preset = self.preset.currentData()
        if self.backend.currentData() == ai.ANTHROPIC_BACKEND:
            self.hint.setText(
                'Anthropic Messages：地址填到 /v1（例如 https://api.xxx.com/v1），'
                '程序请求 /v1/messages，带 x-api-key 与 anthropic-version 头，'
                'system 单独成字段、max_tokens 必填；较新的 claude 若拒绝 temperature，'
                '程序会自动去掉该参数重试一次。“检测并获取模型”走 GET /v1/models，'
                '列不出来就手填中转站接受的模型 ID，再用“测试对话”确认连通。'
                '密钥只写进本机数据目录的 ai_config.json，不上传到任何其他地方。')
            return
        if preset == 'doubao':
            text = ('豆包：地址已填火山方舟 OpenAI 兼容端点。API Key 在火山方舟控制台创建；'
                    '模型填模型 ID（如 doubao-seed-2-1-pro-260915）或推理接入点 ID（ep-…）。'
                    '方舟通常不提供模型列表，请直接用“测试对话”验证。')
        elif preset == 'deepseek':
            text = 'DeepSeek：模型 deepseek-chat，到 DeepSeek 开放平台创建 API Key 后填到上面。'
        elif preset == 'qwen':
            text = '通义千问：使用 DashScope OpenAI 兼容地址，模型 qwen-plus；填写百炼平台 API Key。'
        elif preset == 'ollama':
            text = ('本机 Ollama：地址 127.0.0.1:11434，无需 API Key；服务没跑时工具会自己把它拉起来。'
                    '正文生成和结构识别建议用非思考版（qwen3:4b-instruct 或 qwen2.5:7b），'
                    '视觉复核再配 qwen3-vl:4b。')
        elif preset == 'embedded':
            text = ('内置小模型：随包附带的 Qwen3-0.6B（llama.cpp 纯 CPU 推理），不要 Key、不联网、'
                    '不上传任何文档内容。可以拿来做字段名规范化、分组建议和短正文起草，'
                    '但 0.6B 语义能力有限，长正文和复杂分类请改用在线大模型或本机 Ollama 大档。'
                    '首次调用要加载模型，可能等几十秒。')
        else:
            text = '自定义 OpenAI 兼容服务：填写 base_url、模型名和 Key；本机局域网服务可留空 Key。'
        self.hint.setText(text)

    def vision_preset_changed(self, *_):
        if self._loading:
            return
        key = self.vpreset.currentData()
        entry = next((item for item in ai.VISION_PRESETS if item[1] == key), None)
        if entry is None:
            self.update_vision_hint()
            return
        _label, _key, backend, base, model = entry
        idx = self.vbackend.findData(backend)
        self.vbackend.setCurrentIndex(max(0, idx))
        self.vbase.setText(base)
        self.vmodel.setCurrentText(model)
        self.vision_backend_changed()
        self.update_vision_hint()

    def vision_backend_changed(self, *_):
        backend = self.vbackend.currentData()
        base = self.vbase.text().strip()
        if backend == 'ollama':
            if base in ('', ai.SILICONFLOW_DEFAULT, ai.ANTHROPIC_DEFAULT):
                self.vbase.setText(ai.OLLAMA_DEFAULT)
            self.vkey.setEnabled(False)
        elif backend == ai.ANTHROPIC_BACKEND:
            self.vkey.setEnabled(True)
            if base in ('', ai.SILICONFLOW_DEFAULT, ai.OLLAMA_DEFAULT, ai.OPENAI_DEFAULT):
                # 同一个中转站：地址直接跟上面文本配置对齐，Key 留空即复用
                self.vbase.setText(self.base.text().strip())
        else:
            if base == ai.OLLAMA_DEFAULT:
                self.vbase.setText(ai.SILICONFLOW_DEFAULT)
            self.vkey.setEnabled(True)
        if hasattr(self, 'vhint'):
            self.update_vision_hint()

    def update_vision_hint(self):
        key = self.vpreset.currentData()
        if key == 'siliconflow':
            text = ('硅基流动：模型 Qwen/Qwen3-VL-32B-Instruct（Qwen3-VL 开源权重）。'
                    '到 siliconflow.cn 创建 API Key 后填入上方“视觉 API Key”。')
        elif key == 'dashscope':
            text = ('阿里云百炼：模型 qwen3-vl-plus，OpenAI 兼容地址已填。'
                    '到百炼控制台创建 API Key；它是 Qwen3-VL 的托管服务。')
        elif key == 'openrouter':
            text = ('OpenRouter：模型 qwen/qwen3-vl-235b-a22b-instruct，需在 openrouter.ai 创建 Key。'
                    '如账号没有该模型，可在模型框手填其他 Qwen3-VL 模型 ID。')
        elif key == 'ollama':
            text = ('本机 Ollama：先安装 Ollama 并执行 ollama pull qwen3-vl:8b，'
                    '保持 http://127.0.0.1:11434 运行；无需 API Key。')
        elif self.vbackend.currentData() == ai.ANTHROPIC_BACKEND:
            text = ('Anthropic Messages：地址填到 /v1（例如 https://api.xxx.com/v1），'
                    '程序会请求 /v1/messages，图片按 base64 image block 送出。'
                    '模型框里要选支持图片输入的模型；与文本用同一个中转站时，'
                    '地址和密钥都可以留空复用上面的配置。')
        else:
            text = ('自定义视觉服务：填写 OpenAI 兼容地址、模型 ID 和 Key；'
                    '本机局域网服务可留空 Key。模型需支持图片输入。')
        same = self.vbackend.currentData() == self.backend.currentData()
        vbase = self.vbase.text().strip().rstrip('/').lower()
        shared = same and (not vbase
                           or vbase == self.base.text().strip().rstrip('/').lower())
        text += (' 视觉地址与密钥留空时，直接复用上面文本模型的地址和密钥（同一个服务不必填两遍）。'
                 if shared else ' 视觉服务与文本服务不同时，请单独填写视觉地址与密钥。')
        self.vhint.setText(text)

    def _set_busy(self, busy, text='', status_widget=None):
        if status_widget is not None:
            self._status_widget = status_widget
        for widget in (self.models_btn, self.test_btn,
                       self.vmodels_btn, self.vtest_btn):
            widget.setEnabled(not busy)
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(not busy)
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setEnabled(not busy)
        (self._status_widget or self.status).setText(text)

    def fetch_models(self):
        if self._worker:
            return
        client = ai.Client(self.current_cfg())
        ok, msg = client.usable()
        if not ok:
            self.status.setText(msg)
            return
        self._worker = AIWorker(client.list_models, self)
        self._worker.done.connect(self.models_ready)
        self._worker.failed.connect(self.models_failed)
        self._worker.finished.connect(self.worker_finished)
        self._worker.finished.connect(self._worker.deleteLater)
        self._set_busy(True, '正在连接服务并获取模型列表…', self.status)
        self._worker.start()

    def test_chat(self):
        if self._worker:
            return
        client = ai.Client(self.current_cfg())
        ok, msg = client.usable()
        if not ok:
            self.status.setText(msg)
            return
        self._worker = AIWorker(
            lambda: client.chat([{'role': 'user', 'content': '只回复 OK'}], max_tokens=16), self)
        self._worker.done.connect(self.chat_ready)
        self._worker.failed.connect(self.chat_failed)
        self._worker.finished.connect(self.worker_finished)
        self._worker.finished.connect(self._worker.deleteLater)
        self._set_busy(True, '正在测试对话…', self.status)
        self._worker.start()

    def fetch_vision_models(self):
        if self._worker:
            return
        client = ai.vision_client(self.current_cfg())
        ok, msg = client.usable()
        if not ok:
            self.vstatus.setText(msg)
            return
        self._worker = AIWorker(client.list_models, self)
        self._worker.done.connect(self.vision_models_ready)
        self._worker.failed.connect(self.vision_models_failed)
        self._worker.finished.connect(self.worker_finished)
        self._worker.finished.connect(self._worker.deleteLater)
        self._set_busy(True, '正在连接视觉服务并获取模型列表…', self.vstatus)
        self._worker.start()

    def test_vision(self):
        if self._worker:
            return
        client = ai.vision_client(self.current_cfg())
        ok, msg = client.usable()
        if not ok:
            self.vstatus.setText(msg)
            return
        try:
            from PIL import Image
            import io
            buffer = io.BytesIO()
            Image.new('RGB', (480, 300), 'white').save(buffer, 'PNG')
            image = buffer.getvalue()
        except Exception:
            image = b''
        self._worker = AIWorker(
            lambda: client.chat_vision(
                [{'role': 'system', 'content': '你只需要确认能看见图片。'},
                 {'role': 'user', 'content': '这是一张白底测试图，只回复 OK'}],
                [image] if image else [], max_tokens=32), self)
        self._worker.done.connect(self.vision_ready)
        self._worker.failed.connect(self.vision_failed)
        self._worker.finished.connect(self.worker_finished)
        self._worker.finished.connect(self._worker.deleteLater)
        self._set_busy(True, '正在测试视觉模型（发送一张空白测试图）…', self.vstatus)
        self._worker.start()

    def vision_models_ready(self, models):
        current = self.vmodel.currentText().strip()
        self.vmodel.clear()
        self.vmodel.addItems(models)
        if current and current not in models:
            self.vmodel.addItem(current)
        target = current
        if not target and models:
            target = next((m for m in models if 'qwen3-vl' in m.lower()), '') \
                or next((m for m in models if 'vl' in m.lower()), '') or models[0]
        if target:
            self.vmodel.setCurrentText(target)
        self.vstatus.setText(f'连接成功，发现 {len(models)} 个模型。')

    def vision_models_failed(self, message):
        message = str(message)
        advice = error_advice(message)
        self.vstatus.setText(f'{message}。可直接手填模型 ID 并点“测试视觉”。'
                             + (('\n' + advice) if advice else ''))

    def vision_ready(self, text):
        text = (text or '').strip()
        self.vstatus.setText('视觉连接成功：' + (text[:60] or '模型已响应'))

    def vision_failed(self, message):
        message = str(message)
        advice = error_advice(message)
        self.vstatus.setText('视觉测试失败：' + message[:180]
                             + (('\n' + advice) if advice else ''))

    def chat_ready(self, text):
        text = (text or '').strip()
        self.status.setText('连接成功：' + (text[:60] or '服务已响应'))

    def chat_failed(self, message):
        message = str(message)
        advice = error_advice(message)
        self.status.setText('测试失败：' + message[:160]
                            + (('\n' + advice) if advice else ''))

    def models_ready(self, models):
        current = self.model.currentText().strip()
        self.model.clear()
        self.model.addItems(models)
        if current and current not in models:
            self.model.addItem(current)
        target = current
        if not target and models:
            target = next((m for m in models if 'qwen' in m.lower()), '') \
                or next((m for m in models if 'deepseek' in m.lower()), '') or models[0]
        if target:
            self.model.setCurrentText(target)
        self.status.setText(f'连接成功，发现 {len(models)} 个模型。'
                            + ('' if models else '可先在命令行执行 ollama pull qwen2.5:7b。'))

    def models_failed(self, message):
        message = str(message)
        advice = error_advice(message)
        self.status.setText(f'{message}。部分服务不提供模型列表，可直接用“测试对话”。'
                            + (('\n' + advice) if advice else ''))

    def worker_finished(self):
        widget = self._status_widget or self.status
        text = widget.text()
        self._worker = None
        self._set_busy(False, text)

    def save(self):
        if self._worker:
            QMessageBox.information(self, '请稍候', '正在检测服务，请等检测结束再保存。')
            return
        cfg = self.current_cfg()
        text_ok, text_msg = ai.Client(cfg).usable()
        vision_ok, vision_msg = ai.vision_usable(cfg)
        if not text_ok and not vision_ok:
            QMessageBox.information(
                self, '配置不完整',
                f'文本模型：{text_msg}\n视觉模型：{vision_msg}\n\n'
                '至少配置一个模型后才能保存。')
            return
        self.cfg = cfg
        self.accept()

    def reject(self):
        if self._worker:
            QMessageBox.information(self, '请稍候', '正在检测服务，请等检测结束后再关闭。')
            return
        super().reject()


class BodyGenerateDialog(QDialog):
    def __init__(self, field, config, values, client, parent=None):
        super().__init__(parent)
        self.field = field
        self.config = config
        self.values = values
        self.client = client
        self._worker = None
        self.append_mode = False
        self.setWindowTitle('AI 生成正文')
        self.resize(760, 640)
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 16)
        root.setSpacing(9)
        root.addWidget(label('AI 生成正文', 'dialogtitle'))
        root.addWidget(label(f'模板：{config["title"]}　字段：{field["label"]}', 'aihint'))
        self.requirement = QLineEdit()
        self.requirement.setPlaceholderText('可选，例如：突出技术路线与创新点，语气正式，分三点写')
        root.addWidget(label('写作要求'))
        root.addWidget(self.requirement)
        options = QHBoxLayout()
        options.addWidget(label('目标字数'))
        self.target = QSpinBox()
        max_length = int(field.get('max_length') or 0)
        self.target.setRange(100, max(100, max_length or 5000))
        self.target.setSingleStep(100)
        self.target.setValue(min(600, max(100, max_length or 5000)))
        options.addWidget(self.target)
        existing = str(values.get(field['key'], '') or '').strip()
        self.append = QCheckBox('在现有内容后追加')
        self.append.setEnabled(bool(existing))
        if not existing:
            self.append.setToolTip('该字段当前为空')
        options.addWidget(self.append)
        options.addStretch()
        self.generate_btn = button('生成初稿', self.start, primary=True)
        options.addWidget(self.generate_btn)
        root.addLayout(options)
        self.preview = QPlainTextEdit()
        self.preview.setPlaceholderText('AI 生成的内容会先出现在这里，确认或修改后再应用到填写框。')
        self.preview.setMinimumHeight(240)
        self.preview.textChanged.connect(self.update_count)
        root.addWidget(self.preview, 1)
        self.counter = label('', 'aihint')
        self.status = label('', 'aihint')
        root.addWidget(self.counter)
        root.addWidget(self.status)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                                        | QDialogButtonBox.StandardButton.Cancel)
        self.apply_btn = self.buttons.button(QDialogButtonBox.StandardButton.Ok)
        self.apply_btn.setText('应用到填写框')
        self.apply_btn.setEnabled(False)
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setText('取消')
        self.buttons.accepted.connect(self.apply)
        self.buttons.rejected.connect(self.reject)
        root.addWidget(self.buttons)

    def update_count(self):
        count = len(self.preview.toPlainText())
        maximum = int(self.field.get('max_length') or 0)
        self.counter.setText(f'{count} 字' + (f' / 模板上限 {maximum} 字' if maximum else ''))
        mark_over(self.counter, bool(maximum and count > maximum))
        if maximum and count > maximum:
            self.counter.setText(self.counter.text() + '  超出上限，应用前请精简')

    def _set_busy(self, busy):
        self.generate_btn.setEnabled(not busy)
        self.generate_btn.setText('生成中…' if busy else '重新生成')
        self.apply_btn.setEnabled(not busy and bool(self.preview.toPlainText().strip()))
        self.buttons.button(QDialogButtonBox.StandardButton.Cancel).setEnabled(not busy)

    def start(self):
        if self._worker:
            return
        existing = str(self.values.get(self.field['key'], '') or '').strip()
        append = self.append.isChecked() and bool(existing)
        target = int(self.target.value())
        max_length = int(self.field.get('max_length') or 0)
        if max_length:
            target = min(target, max_length)
        prompt = ai.body_prompt(self.config['title'], self.config.get('description', ''),
                                self.config.get('fields', []), self.values, self.field['label'],
                                self.requirement.text().strip(), target, max_length,
                                existing, append)
        messages = [{'role': 'system', 'content': ai.BODY_SYSTEM},
                    {'role': 'user', 'content': prompt}]
        limit = min(max(1024, target * 3), 8192)
        client = self.client
        self._worker = AIWorker(
            lambda: client.chat(messages, temperature=0.55, max_tokens=limit), self)
        self._worker.done.connect(self.generated)
        self._worker.failed.connect(self.failed)
        self._worker.finished.connect(self.worker_finished)
        self._worker.finished.connect(self._worker.deleteLater)
        self._set_busy(True)
        self.status.setText('正在生成，通常 10–60 秒；本机小模型或长正文会更久…')
        self._worker.start()

    def generated(self, text):
        text = re.sub(r'^```[a-zA-Z]*\s*|\s*```$', '', (text or '').strip())
        if not text:
            self.failed('模型返回了空内容，请换用其他模型或调大输出上限。')
            return
        self.preview.setPlainText(text)
        self.update_count()
        self.apply_btn.setEnabled(True)
        self.status.setText('生成完成。可直接修改上面的内容，确认后点“应用到填写框”。')

    def failed(self, message):
        self.status.setText('生成失败：' + message)
        QMessageBox.warning(self, 'AI 生成失败', message + '\n\n可在顶栏「AI 设置」检查服务和模型。')

    def worker_finished(self):
        self._worker = None
        self._set_busy(False)

    def result_text(self):
        return self.preview.toPlainText()

    def apply(self):
        if self._worker:
            return
        text = self.preview.toPlainText().strip()
        if not text:
            QMessageBox.information(self, '还没有内容', '请先生成或粘贴一段正文。')
            return
        maximum = int(self.field.get('max_length') or 0)
        if maximum and len(text) > maximum:
            answer = QMessageBox.question(
                self, '超出字数上限',
                f'当前 {len(text)} 字，超过模板上限 {maximum} 字。仍要写入吗？',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        self.append_mode = self.append.isChecked()
        self.accept()

    def reject(self):
        if self._worker:
            QMessageBox.information(self, '请稍候', 'AI 正在生成，请等生成结束再关闭。')
            return
        super().reject()


class BatchDialog(QDialog):
    """一个名单 → N 份 Word：模板只做一次，数据表负责批量。"""

    def __init__(self, workspace, config=None, parent=None):
        super().__init__(parent)
        self.workspace = workspace
        self.store = workspace.store
        self.config = config
        self.headers = []
        self.records = []
        self.combos = []
        self.setWindowTitle('批量填写：一个名单生成多份 Word')
        self.resize(900, 660)
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(10)
        root.addWidget(label('选一个模板和一份名单（CSV / TXT / JSON / Excel），'
                             '每一行生成一份 Word。列名和栏目名对上就自动填，对不上的可以手动指定。'
                             '模板文件和名单都不会被改动。', 'muted'))

        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        self.template_combo = QComboBox()
        for item in self.workspace.fill.configs:
            if item.get('fields'):
                self.template_combo.addItem(item['title'], item['id'])
        if config is not None:
            index = self.template_combo.findData(config['id'])
            if index >= 0:
                self.template_combo.setCurrentIndex(index)
        self.template_combo.currentIndexChanged.connect(self.template_changed)
        grid.addWidget(label('模板', 'strong'), 0, 0)
        grid.addWidget(self.template_combo, 0, 1)
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText('名单：.csv / .txt / .json / .xlsx，第一行是列名…')
        grid.addWidget(label('名单', 'strong'), 1, 0)
        grid.addWidget(self.file_edit, 1, 1)
        grid.addWidget(mini_button('浏览…', self.pick_file), 1, 2)
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText('默认在名单同级的“批量填写”文件夹…')
        grid.addWidget(label('输出到', 'strong'), 2, 0)
        grid.addWidget(self.out_edit, 2, 1)
        grid.addWidget(mini_button('选目录…', self.pick_dir), 2, 2)
        grid.setColumnStretch(1, 1)
        root.addLayout(grid)

        head = QHBoxLayout()
        self.read_btn = button('读取名单', self.read_table, primary=True)
        self.run_btn = button('开始批量生成', self.run_batch)
        self.run_btn.setEnabled(False)
        head.addWidget(self.read_btn)
        head.addWidget(self.run_btn)
        head.addStretch()
        self.status = label('', 'muted')
        head.addWidget(self.status)
        root.addLayout(head)

        self.task = feedback.TaskBar()
        root.addWidget(self.task)

        self.map = QTableWidget(0, 2)
        self.map.setHorizontalHeaderLabels(['名单里的列名', '对应栏目'])
        self.map.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.map.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.map.verticalHeader().setVisible(False)
        self.map.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.map.setMaximumHeight(230)
        root.addWidget(self.map)

        self.out = QPlainTextEdit()
        self.out.setReadOnly(True)
        self.out.setPlaceholderText('读取名单后，这里显示列名与栏目的对应情况和生成结果。')
        root.addWidget(self.out, 1)
        footer = QHBoxLayout()
        self.open_btn = button('打开输出文件夹', self.open_folder)
        self.open_btn.setEnabled(False)
        footer.addWidget(self.open_btn)
        footer.addStretch()
        footer.addWidget(button('关闭', self.reject))
        root.addLayout(footer)

    # ---------------------------------------------------------------- 选择
    def current_config(self):
        template_id = self.template_combo.currentData()
        for item in self.workspace.fill.configs:
            if item['id'] == template_id:
                return item
        return None

    def template_changed(self, _index):
        self.config = self.current_config()
        if self.headers and self.config:
            self.build_map()

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择名单', '',
                                              '名单 (*.csv *.txt *.json *.xlsx);;所有文件 (*)')
        if path:
            self.file_edit.setText(path)
            if not self.out_edit.text().strip():
                self.out_edit.setText(str(Path(path).parent / '批量填写'))

    def pick_dir(self):
        path = QFileDialog.getExistingDirectory(self, '输出到哪个文件夹',
                                                self.out_edit.text().strip() or str(Path.home()))
        if path:
            self.out_edit.setText(path)

    # ---------------------------------------------------------------- 名单
    def read_table(self):
        config = self.current_config()
        path = self.file_edit.text().strip()
        if not config:
            self.inform('还没有模板', '先在上方选一个模板（模板要先识别出栏目）。')
            return
        if not path or not Path(path).is_file():
            self.inform('缺少名单', '请选择名单文件（.csv / .txt / .json / .xlsx）。')
            return
        self.read_btn.setEnabled(False)
        self.status.setText('正在读名单…')
        QApplication.processEvents()
        try:
            self.headers, self.records = batchfill.read_table(path)
        except Exception as exc:
            self.read_btn.setEnabled(True)
            self.status.setText('')
            self.inform('名单读不出来', f'{type(exc).__name__}：{exc}')
            return
        self.read_btn.setEnabled(True)
        if not self.headers:
            self.status.setText('名单是空的')
            self.out.setPlainText('没读到列名。第一行请写列名，例如：学生姓名,学号,中文题目')
            return
        self.build_map()
        self.status.setText(f'{len(self.records)} 行 · {len(self.headers)} 列')
        self.run_btn.setEnabled(bool(self.records))

    def build_map(self):
        mapping = batchfill.match_columns(self.headers, self.config['fields'])
        self.map.setRowCount(len(self.headers))
        self.combos = []
        for row, header in enumerate(self.headers):
            self.map.setItem(row, 0, QTableWidgetItem(header))
            combo = QComboBox()
            combo.addItem('（不填这一列）', '')
            for field in self.config['fields']:
                combo.addItem(f"{field.get('label') or field['key']}"
                              f"（{field.get('type', 'text')}）", field['key'])
            target = mapping.get(header, '')
            index = combo.findData(target)
            combo.setCurrentIndex(index if index >= 0 else 0)
            self.map.setCellWidget(row, 1, combo)
            self.combos.append((header, combo))
        lines = ['列名 → 栏目：']
        for header, combo in self.combos:
            lines.append(f'  · {header} → {combo.currentText()}')
        lines.append(f'共 {len(self.records)} 行数据。')
        self.out.setPlainText('\n'.join(lines))

    def assignments(self):
        return {header: combo.currentData() for header, combo in self.combos
                if combo.currentData()}

    # ---------------------------------------------------------------- 生成
    def run_batch(self):
        config = self.current_config()
        path = self.file_edit.text().strip()
        if not config or not self.records:
            return
        out_dir = self.out_edit.text().strip() or str(Path(path).parent / '批量填写')
        mapping = self.assignments()
        if not mapping:
            self.inform('没有对应关系', '至少要有一列对应到某个栏目，否则每份都是空白模板。')
            return
        name_column = next((header for header in self.headers
                            if '姓名' in header or '名称' in header), self.headers[0])
        self.run_btn.setEnabled(False)
        total = len(self.records)
        self.status.setText(f'正在生成 {total} 份…')
        self.task.start(f'正在生成 {total} 份…', total)
        done = [0]

        def tick(text):
            done[0] += 1
            self.status.setText(text)
            self.task.step(done[0], total, text)
            QApplication.processEvents()

        try:
            report = batchfill.run(Path(config['_dir']) / config['file'], out_dir, config,
                                   self.headers, self.records, mapping,
                                   name_column=name_column, progress=tick,
                                   should_stop=self.task.is_cancelled)
        except Exception as exc:
            self.run_btn.setEnabled(True)
            self.task.finish()
            self.status.setText('')
            self.inform('批量生成失败', str(exc)
                        + ('\n\n' + error_advice(exc, Path(config['_dir']) / config['file'])
                           if error_advice(exc, Path(config['_dir']) / config['file']) else ''))
            return
        self.run_btn.setEnabled(True)
        self.task.finish()
        self.last_output = out_dir
        self.open_btn.setEnabled(True)
        lines = [f'成功 {report["done"]} 份 → {out_dir}', '']
        if report.get('cancelled'):
            lines.insert(0, f'已取消：剩余 {report["cancelled"]} 份没有生成，'
                            f'已经写好的 {report["done"]} 份保留在输出目录。')
        for problem in report['problems']:
            lines.append('提示：' + problem)
        for number in report['skipped']:
            lines.append(f'第 {number} 行整行是空的，已跳过')
        for name, message in report['failed']:
            lines.append(f'失败：{name} —— {message}')
        if report['files']:
            lines.append('')
            lines.append('生成的文件：')
            lines.extend('  ' + Path(file).name for file in report['files'][:20])
            if len(report['files']) > 20:
                lines.append(f'  …另有 {len(report["files"]) - 20} 份')
        self.out.setPlainText('\n'.join(lines))
        tail = f'完成：{report["done"]} 份，失败 {len(report["failed"])} 份'
        if report.get('cancelled'):
            tail += f'，已取消 {report["cancelled"]} 份'
        self.status.setText(tail)

    last_output = ''

    def open_folder(self):
        if self.last_output:
            open_in_explorer(Path(self.last_output))

    def inform(self, title, text):
        QMessageBox.information(self, title, text)


class StudioPage(DropToOpen, QWidget):
    def __init__(self, store, workspace):
        super().__init__()
        self.workspace = workspace
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 18, 24, 14)
        root.setSpacing(10)
        root.addWidget(label('模板制作', 'title'))
        root.addWidget(label('“开始识别”自动读取表格栏目、段落填空、下划线、括号提示、内容控件和表单域，'
                             '不需要任何占位符。程序还会自动读取模板自己写明的字体、字号、行距、缩进和'
                             '内容要求，生成时照做，并把说明搬成填写提示——这一步始终开启，无需勾选。'
                             '“小模型补全空标记名称”保留 v3.2 流程。规范/说明文档没有可填位时会拒绝录入。'
                             '识别后可改名并录入模板库；“结构体检”页签用于只读排查问题。'
                             '两个页签都不会改动原文件，除非你明确录入或另存。', 'muted'))
        self.tabs = QTabWidget()
        self.mark = MarkPage(store, workspace)
        self.inspect = InspectPage(store, workspace)
        self.mark.layout().setContentsMargins(0, 8, 0, 0)
        self.inspect.layout().setContentsMargins(0, 8, 0, 0)
        self.tabs.addTab(self.mark, '自动识别')
        self.tabs.addTab(self.inspect, '结构体检')
        root.addWidget(self.tabs, 1)
        self.init_drops(self.drop_document)

    def drop_document(self, path):
        """拖进来的文档是要打标的，先切回“自动识别”再载入。"""
        self.tabs.setCurrentWidget(self.mark)
        self.mark.load(path)


# ============================================================ 首页
class ClickCard(QFrame):
    """整张卡片都能点的容器：卡里的小按钮自己吃掉点击，空白处走 on_click。

    也支持键盘：Tab 走进卡片墙，方向键按几何位置找最近的一张，回车打开。
    """

    def __init__(self, on_click=None, parent=None):
        super().__init__(parent)
        self.on_click = on_click
        self.setObjectName('cardhover')
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

    def mouseReleaseEvent(self, event):
        point = event.position().toPoint() if hasattr(event, 'position') else event.pos()
        if (self.on_click and event.button() == Qt.MouseButton.LeftButton
                and self.rect().contains(point)):
            self.on_click()
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Space):
            if self.on_click:
                self.on_click()
            return
        step = {Qt.Key.Key_Left: (-1, 0), Qt.Key.Key_Right: (1, 0),
                Qt.Key.Key_Up: (0, -1), Qt.Key.Key_Down: (0, 1)}.get(event.key())
        if step:
            self.move_focus(*step)
            return
        super().keyPressEvent(event)

    def move_focus(self, dx, dy):
        """不依赖列数：在该方向上找主轴距离最近、副轴偏移最小的一张卡片。"""
        host = self.parent()
        if host is None:
            return
        mine = self.mapTo(host, self.rect().center())
        best, best_score = None, None
        for card in host.findChildren(ClickCard):
            if card is self or not card.isVisible():
                continue
            other = card.mapTo(host, card.rect().center())
            vx, vy = other.x() - mine.x(), other.y() - mine.y()
            if (dx and vx * dx <= 0) or (dy and vy * dy <= 0):
                continue
            along = abs(vx) if dx else abs(vy)
            across = abs(vy) if dx else abs(vx)
            score = along + across * 2.5
            if best_score is None or score < best_score:
                best, best_score = card, score
        if best is not None:
            best.setFocus()


class HomePage(DropToOpen, QWidget):
    COLUMNS = 3
    MAX_CARDS = 30

    def __init__(self, store, workspace):
        super().__init__()
        self.store, self.workspace = store, workspace
        self.configs = []
        self.query = ''
        self._signature = None
        self.build_ui()
        self.init_drops(self.drop_document)
        self.reload()

    def drop_document(self, path):
        """首页写着“把 Word 文档丢进来”，那就真的能丢：直接送去填写页识别。"""
        self.workspace.nav.setCurrentRow(Workspace.FILL)
        self.workspace.fill.start_direct(path)

    def build_ui(self):
        self.setObjectName('glasspage')
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 20, 28, 16)
        root.setSpacing(14)
        root.addWidget(label('把 Word 文档丢进来，剩下的交给工作台', 'hero'))
        root.addWidget(label('自动识别填写位 · 读取模板自己写明的字体字号行距与内容要求并照做 · '
                             '生成的是新文件，原文件不改动 · 所有内容都留在本机。', 'subtitle'))
        chips = QHBoxLayout()
        chips.setSpacing(8)
        self.chips = []
        for _ in range(4):
            chip = label('', 'chip')
            chips.addWidget(chip)
            self.chips.append(chip)
        chips.addStretch()
        root.addLayout(chips)

        actions = QGridLayout()
        actions.setSpacing(12)
        entries = (
            ('直接填写', '拖入或选择 .docx，识别填写位后生成新文件', '去填写', self.go_direct),
            ('制作模板', '自动识别 + 智能打标，注册进模板库反复复用', '去打标', self.go_mark),
            ('模板库', '导入、重命名、导出、删除自定义模板', '去管理', self.go_library),
            ('整篇排版', '按纸张/字体/行距/标题规则重排整篇文档', '去排版', self.go_layout))
        for index, (text, tip, name, callback) in enumerate(entries):
            card = ClickCard(callback)
            box = QVBoxLayout(card)
            box.setContentsMargins(16, 14, 16, 14)
            box.setSpacing(6)
            box.addWidget(label(text, 'section'))
            box.addWidget(label(tip, 'cardmeta'))
            tail = QHBoxLayout()
            tail.addStretch()
            tail.addWidget(mini_button(name, callback))
            box.addLayout(tail)
            if index == 0:
                actions.addWidget(card, 0, 0, 1, 2)   # 主操作占两列
            else:
                actions.addWidget(card, 1, index - 1)
        for column in range(3):
            actions.setColumnStretch(column, 1)
        root.addLayout(actions)

        head = QHBoxLayout()
        head.setSpacing(8)
        head.addWidget(label('可用模板', 'section'))
        head.addStretch()
        self.count = label('', 'badge')
        head.addWidget(self.count)
        root.addLayout(head)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        host.setObjectName('glasspage')
        self.wall = QGridLayout(host)
        self.wall.setContentsMargins(0, 0, 0, 0)
        self.wall.setSpacing(12)
        for column in range(self.COLUMNS):
            self.wall.setColumnStretch(column, 1)
        scroll.setWidget(host)
        root.addWidget(scroll, 1)

    def reload(self):
        self.configs = self.store.templates()
        # 切页会调到这里：模板没变就别重建整面卡片墙，这是切页最贵的开销
        signature = [(config['id'], config['title'], config.get('category') or '',
                      len(config['fields'])) for config in self.configs]
        if signature != self._signature:
            self._signature = signature
            self.render_wall()
        cfg = ai.load_config(self.store.ai_config)
        texts = [f'版本 v{VERSION}',
                 '主题：' + THEMES[self.workspace.theme]['name'],
                 '文本 AI：已配置' if ai.Client(cfg).usable()[0] else '文本 AI：未配置',
                 '视觉 AI：已配置' if ai.vision_usable(cfg)[0] else '视觉 AI：未配置']
        for index, chip in enumerate(self.chips):
            chip.setText(texts[index] if index < len(texts) else '')
            chip.setVisible(index < len(texts))

    def matches(self, config):
        if not self.query:
            return True
        blob = ' '.join((config['title'], config.get('category') or '', config['id'])).lower()
        return self.query.lower() in blob

    def render_wall(self):
        while self.wall.count():
            item = self.wall.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
                widget.deleteLater()
        for row in range(self.wall.rowCount()):
            self.wall.setRowStretch(row, 0)
        shown = [config for config in self.configs if self.matches(config)]
        for number, config in enumerate(shown[:self.MAX_CARDS]):
            row, column = divmod(number, self.COLUMNS)
            self.wall.addWidget(self.template_card(config), row, column)
        if not shown:
            self.wall.addWidget(label('没有匹配的模板。可以直接填写页拖入一个 .docx，'
                                      '或到模板库里导入。', 'muted'), 0, 0, 1, self.COLUMNS)
        self.wall.setRowStretch(self.wall.rowCount(), 1)
        extra = '' if len(shown) <= self.MAX_CARDS else f'（仅显示前 {self.MAX_CARDS} 个）'
        self.count.setText(f'{len(shown)} / {len(self.configs)} 个模板{extra}')

    def template_card(self, config):
        card = ClickCard(lambda: self.workspace.go_fill(config['id']))
        box = QVBoxLayout(card)
        box.setContentsMargins(15, 13, 15, 13)
        box.setSpacing(5)
        box.addWidget(label(config['title'], 'section'))
        box.addWidget(label(' · '.join((
            config.get('category') or '未分类',
            f"{len(config['fields'])} 个填写位",
            '内置' if config['id'].startswith('builtin-') else '自定义')), 'cardmeta'))
        tail = QHBoxLayout()
        tail.setSpacing(6)
        tail.addWidget(mini_button('去填写', lambda: self.workspace.go_fill(config['id'])))
        tail.addWidget(mini_button('模板库',
                                   lambda: self.workspace.inspect_template(config['id'])))
        tail.addStretch()
        box.addLayout(tail)
        return card

    def set_filter(self, text):
        self.query = (text or '').strip()
        self.render_wall()

    def go_direct(self):
        self.workspace.nav.setCurrentRow(Workspace.FILL)
        self.workspace.fill.open_document_dialog()

    def go_mark(self):
        self.workspace.nav.setCurrentRow(Workspace.STUDIO)

    def go_library(self):
        self.workspace.open_library()

    def go_layout(self):
        self.workspace.nav.setCurrentRow(Workspace.LAYOUT)


def mini_button(text, callback=None):
    widget = button(text, callback)
    widget.setObjectName('mini')
    return widget


# ============================================================ ④ 自动排版页
CJK_FONTS = ['宋体', '黑体', '楷体', '仿宋', '华文中宋', '微软雅黑', '等线', '隶书', '小标宋体']
LATIN_FONTS = ['Times New Roman', 'Cambria', 'Arial', 'Calibri', 'Verdana']
NO_CHANGE = '不改'
ALIGN_BUTTONS = (('不', '不改对齐方式', None), ('左', '左对齐', 'left'),
                 ('中', '居中', 'center'), ('两', '两端对齐', 'justify'),
                 ('右', '右对齐', 'right'))
BOLD_CHOICES = (('不改', None), ('加粗', True), ('不加粗', False))
SPACE_CHOICES = (('不改', None), ('0 磅', 0.0), ('6 磅', 6.0), ('12 磅', 12.0), ('18 磅', 18.0))
COLOR_CHOICES = (('不改', None), ('黑色', '000000'), ('蓝色', '0000FF'),
                 ('深灰', '595959'), ('红色', 'FF0000'))
SHADING_CHOICES = (('不改', None), ('浅灰', 'D9D9D9'), ('浅黄', 'FFF2CC'),
                   ('浅蓝', 'DDEBF7'), ('清除底纹', ''))
HIGHLIGHT_CHOICES = (('不改', None), ('黄', 'yellow'), ('绿', 'green'),
                     ('青', 'cyan'), ('清除高亮', 'none'))
PAGINATION_CHOICES = (('不改', None), ('标题与下段同页', ('keep_next', True)),
                      ('段中不分页', ('keep_lines', True)),
                      ('孤行控制开', ('widow_control', True)),
                      ('孤行控制关', ('widow_control', False)),
                      ('同页+不分页', ('keep_next_keep_lines', True)))
MARK_CHOICES = (('不改', None), ('回车后跟随本档', 'follow'), ('回车后黑色不加粗', 'plain'))
SCOPE_CHOICES = (('不改', 'off'), ('整篇文档', 'all'), ('仅填写内容', 'filled'))


def pick_by_data(combo, value):
    """按 userData 选中下拉项，找不到就退回第一项（“不改”）。"""
    for index in range(combo.count()):
        if combo.itemData(index) == value:
            combo.setCurrentIndex(index)
            return
    combo.setCurrentIndex(0)


def narrow_combo(combo):
    """让下拉框按面板宽度收缩，而不是被最长条目撑出横向滚动条。"""
    combo.setMinimumContentsLength(4)
    combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
    return combo


def editable_combo(choices, sample='', tip=''):
    """档位下拉 + 自由输入框合一个：下拉给常用档位，也能直接输入 22磅 / 1.5倍 / FF0000。"""
    widget = QComboBox()
    widget.setEditable(True)
    for text, value in choices:
        widget.addItem(text, value)
    line = widget.lineEdit()
    if line is not None:
        line.setPlaceholderText(sample)
    if tip:
        widget.setToolTip(tip)
    return narrow_combo(widget)


def parse_indent(text):
    """缩进按“字符”给：Word 的首行缩进在中文文档里就是几个字宽。"""
    value = str(text or '').strip()
    match = re.match(r'^(-?\d+(?:\.\d+)?)\s*(字符|字|chars?)?$', value, re.IGNORECASE)
    if match:
        return float(match.group(1))
    raise units.UnitError('缩进请按字符填，例如 2 字符、1.5、0（不缩进）')


ALIGN_NAMES = {'left': '左对齐', 'center': '居中', 'right': '右对齐', 'justify': '两端对齐'}


def style_summary(rules):
    """自定义样式的一行摘要（内置样式自带 summary，不走这里）。"""
    bits = []
    if rules.get('font_east'):
        bits.append(str(rules['font_east']))
    if rules.get('font_ascii'):
        bits.append(str(rules['font_ascii']))
    size = rules.get('size_pt')
    if size:
        bits.append(next((name for name, value in CN_SIZES if float(value) == float(size)),
                         f'{size:g}pt'))
    line = rules.get('line') or {}
    if line:
        bits.append(f'{line.get("value"):g}倍行距' if line.get('type') == 'multiple'
                    else f'{line.get("value"):g}磅'
                         + ('固定' if line.get('type') == 'exact' else '最小'))
    if rules.get('first_line_chars'):
        bits.append(f'首行缩进{rules["first_line_chars"]:g}字符')
    if rules.get('space_before') is not None:
        bits.append(f'段前{rules["space_before"]:g}磅')
    if rules.get('space_after') is not None:
        bits.append(f'段后{rules["space_after"]:g}磅')
    if rules.get('align'):
        bits.append(ALIGN_NAMES.get(rules['align'], rules['align']))
    if rules.get('bold') is True:
        bits.append('加粗')
    elif rules.get('bold') is False:
        bits.append('不加粗')
    if rules.get('color'):
        bits.append('指定颜色')
    return ' / '.join(bits) or '跟随模板'


class RuleGroup(QFrame):
    """一档排版参数：正文 / 标题1-3 / 表格内文字 / 页眉页脚，各自带作用范围。"""

    def __init__(self, key, title, parent=None):
        super().__init__(parent)
        self.key = key
        self.setObjectName('groupcard')
        box = QVBoxLayout(self)
        box.setContentsMargins(10, 10, 10, 12)
        box.setSpacing(8)
        head = QHBoxLayout()
        head.setSpacing(8)
        head.addWidget(label(title, 'section'), 1)
        self.scope = narrow_combo(QComboBox())
        for text, value in SCOPE_CHOICES:
            self.scope.addItem(text, value)
        self.scope.setToolTip('“仅填写内容”只改本工具填写位书签里的文字；纯排版文档请选“整篇文档”。')
        self.scope.setFixedWidth(96)
        head.addWidget(self.scope)
        box.addLayout(head)
        row = QHBoxLayout()
        row.setSpacing(6)
        row.addWidget(label('中文', 'muted'))
        self.east = self._combo([NO_CHANGE] + CJK_FONTS)
        row.addWidget(self.east, 1)
        row.addWidget(label('西文', 'muted'))
        self.ascii = self._combo([NO_CHANGE] + LATIN_FONTS)
        row.addWidget(self.ascii, 1)
        box.addLayout(row)
        row2 = QHBoxLayout()
        row2.setSpacing(6)
        row2.addWidget(label('字号', 'muted'))
        self.size = editable_combo([(NO_CHANGE, None)] +
                                   [(f'{name} {value:g}pt', float(value)) for name, value in CN_SIZES],
                                   '小四 / 12pt / 10.5',
                                   '选中档位或直接输入：小四、五号、12pt、10.5')
        row2.addWidget(self.size, 1)
        row2.addWidget(label('行距', 'muted'))
        self.line = editable_combo([(NO_CHANGE, None)] + list(layout.LINE_CHOICES),
                                   '1.5倍 / 20磅 / 固定值20磅',
                                   '可以输入 1.5倍、150%、20磅、固定值20磅、至少12磅')
        row2.addWidget(self.line, 1)
        box.addLayout(row2)
        row3 = QHBoxLayout()
        row3.setSpacing(6)
        row3.addWidget(label('缩进', 'muted'))
        self.indent = editable_combo([(NO_CHANGE, None), ('不缩进', 0.0), ('2 字符', 2.0),
                                      ('4 字符', 4.0)], '2 字符 / 1cm',
                                     '可以输入 2 字符、1.5、1cm、0.5英寸')
        row3.addWidget(self.indent, 1)
        row3.addWidget(label('加粗', 'muted'))
        self.bold = QComboBox()
        for text, value in BOLD_CHOICES:
            self.bold.addItem(text, value)
        self.bold.setFixedWidth(78)
        row3.addWidget(self.bold)
        row3.addStretch(1)
        box.addLayout(row3)
        row4 = QHBoxLayout()
        row4.setSpacing(4)
        row4.addWidget(label('对齐', 'muted'))
        self.align_group = QButtonGroup(self)
        self.align_group.setExclusive(True)
        self.align_buttons = {}
        for index, (text, tip, value) in enumerate(ALIGN_BUTTONS):
            widget = button(text)
            widget.setObjectName('tool')
            widget.setCheckable(True)
            widget.setToolTip(tip)
            widget.setFixedWidth(38)
            self.align_group.addButton(widget, index)
            self.align_buttons[value] = index
            row4.addWidget(widget)
        row4.addStretch(1)
        box.addLayout(row4)
        row5 = QHBoxLayout()
        row5.setSpacing(6)
        row5.addWidget(label('段前', 'muted'))
        self.before = editable_combo(list(SPACE_CHOICES), '6磅 / 0.5cm',
                                     '可以输入 6磅、0.5cm、12')
        row5.addWidget(self.before, 1)
        row5.addWidget(label('段后', 'muted'))
        self.after = editable_combo(list(SPACE_CHOICES), '0磅 / 12磅',
                                    '可以输入 12磅、0.5cm、0')
        row5.addWidget(self.after, 1)
        row5.addWidget(label('颜色', 'muted'))
        self.color = editable_combo(list(COLOR_CHOICES), 'FF0000 / 深灰',
                                    '可以输入 FF0000、#C00000、深灰、蓝色、rgb(0,112,192)、自动')
        row5.addWidget(self.color, 1)
        box.addLayout(row5)
        row6 = QHBoxLayout()
        row6.setSpacing(6)
        row6.addWidget(label('底纹', 'muted'))
        self.shading = editable_combo(list(SHADING_CHOICES), 'D9D9D9',
                                      '段落底纹：可选档位或直接输十六进制色号，清除底纹填“清除底纹”')
        row6.addWidget(self.shading, 1)
        row6.addWidget(label('高亮', 'muted'))
        self.highlight = editable_combo(list(HIGHLIGHT_CHOICES), 'yellow',
                                        '文字高亮色：黄/绿/青/品红 或 FF0000；“清除高亮”去掉原有高亮')
        row6.addWidget(self.highlight, 1)
        row6.addWidget(label('回车', 'muted'))
        self.mark = editable_combo(list(MARK_CHOICES), '',
                                   '回车后新打出来的字继承什么：跟随本档，或回到黑色不加粗')
        row6.addWidget(self.mark, 1)
        box.addLayout(row6)
        row7 = QHBoxLayout()
        row7.setSpacing(6)
        row7.addWidget(label('分页', 'muted'))
        self.pagination = editable_combo(list(PAGINATION_CHOICES), '',
                                         '标题常用“标题与下段同页”，长段落用“段中不分页”')
        row7.addWidget(self.pagination, 1)
        row7.addStretch(1)
        box.addLayout(row7)
        self.parsers = {'size': units.parse_font_size, 'line': units.parse_spacing,
                        'before': units.parse_spacing_before,
                        'after': units.parse_spacing_before,
                        'color': units.parse_color, 'indent': parse_indent,
                        'shading': units.parse_color, 'highlight': units.parse_highlight}
        self.scope.currentIndexChanged.connect(lambda *_: self.visibility_changed())
        self.visibility_changed()

    def _combo(self, texts):
        combo = QComboBox()
        for text in texts:
            combo.addItem(text, None if text == NO_CHANGE else text)
        combo.setEditable(True)
        return narrow_combo(combo)

    def value_of(self, key, combo, name):
        """下拉选中就用档位值；手输的现场解析，解析不了就红框说明并当“不改”。"""
        text = (combo.currentText() or '').strip()
        for index in range(combo.count()):
            if (combo.itemText(index) or '').strip() == text:
                return combo.itemData(index)
        if not text:
            return None
        parser = self.parsers.get(key)
        if parser is None:
            return text
        try:
            return parser(text)
        except units.UnitError as exc:
            mark_over(combo, True)
            combo.setToolTip(str(exc))
            return None

    def _font_value(self, combo):
        text = (combo.currentText() or '').strip()
        if text == NO_CHANGE:
            return None
        for index in range(combo.count()):
            if combo.itemText(index) == text:
                return combo.itemData(index)
        return text or None

    def visibility_changed(self):
        off = self.scope.currentData() == 'off'
        for child in (self.east, self.ascii, self.size, self.line, self.indent, self.bold,
                      self.before, self.after, self.color, self.shading, self.highlight,
                      self.mark, self.pagination):
            child.setEnabled(not off)
        for widget in self.align_group.buttons():
            widget.setEnabled(not off)

    def _set_combo(self, combo, value):
        """把规则值放回下拉：能对上档位就选档位，对不上就原样填进输入框。"""
        if value is None:
            combo.setCurrentIndex(0)
            return
        for index in range(combo.count()):
            if combo.itemData(index) == value:
                combo.setCurrentIndex(index)
                return
        combo.setCurrentText(str(value))

    def set_rules(self, rules):
        rules = dict(rules or {})
        east = rules.get('font_east')
        if east and re.match(r'\s*[A-Za-z]', str(east)):
            rules.setdefault('font_ascii', east)
            rules['font_east'] = None
        pick_by_data(self.east, rules.get('font_east'))
        pick_by_data(self.ascii, rules.get('font_ascii'))
        self._set_combo(self.size, rules.get('size_pt'))
        self._set_combo(self.indent, float(rules['first_line_chars'])
                        if rules.get('first_line_chars') is not None else None)
        line = rules.get('line')
        matched = None
        for index in range(self.line.count()):
            value = self.line.itemData(index)
            if value and line and (value.get('type') == line.get('type')
                                   and abs(float(value.get('value') or 0)
                                           - float(line.get('value') or 0)) < 0.01):
                matched = index
                break
        self.line.setCurrentIndex(0 if matched is None else matched)
        for value, index in self.align_buttons.items():
            self.align_group.buttons()[index].setChecked(
                value == (rules.get('align') or None))
        bold = rules.get('bold')
        pick_by_data(self.bold, bold if bold in (True, False) else None)
        self._set_combo(self.before, rules.get('space_before'))
        self._set_combo(self.after, rules.get('space_after'))
        self._set_combo(self.color, rules.get('color'))
        self._set_combo(self.shading, rules.get('shading_fill'))
        self._set_combo(self.highlight, rules.get('highlight'))
        self._set_combo(self.pagination, self._pagination_value(rules))

    @staticmethod
    def _pagination_value(rules):
        if rules.get('keep_next') and rules.get('keep_lines'):
            return ('keep_next_keep_lines', True)
        for key in ('keep_next', 'keep_lines', 'widow_control'):
            if rules.get(key) is not None:
                return (key, rules[key])
        return None

    def set_scope(self, scope):
        for index in range(self.scope.count()):
            if self.scope.itemData(index) == scope:
                self.scope.setCurrentIndex(index)
                return
        self.scope.setCurrentIndex(0)

    def _choice(self, combo):
        """可编辑下拉：按当前文字找档位，找不到就当没选（setCurrentText 不一定会挪 index）。"""
        text = (combo.currentText() or '').strip()
        for index in range(combo.count()):
            if (combo.itemText(index) or '').strip() == text:
                return combo.itemData(index)
        return None

    def state(self):
        if self.scope.currentData() == 'off':
            return {'key': self.key, 'scope': 'off', 'rules': {}}
        for combo in (self.size, self.line, self.indent, self.before, self.after, self.color,
                      self.shading, self.highlight):
            mark_over(combo, False)
        rules = {'font_east': self._font_value(self.east),
                 'font_ascii': self._font_value(self.ascii),
                 'size_pt': self.value_of('size', self.size, '字号'),
                 'line': self.value_of('line', self.line, '行距'),
                 'first_line_chars': self.value_of('indent', self.indent, '缩进'),
                 'align': self._align_value(),
                 'bold': self.bold.currentData(),
                 'space_before': self.value_of('before', self.before, '段前间距'),
                 'space_after': self.value_of('after', self.after, '段后间距'),
                 'color': self.value_of('color', self.color, '颜色'),
                 'shading_fill': self.value_of('shading', self.shading, '底纹'),
                 'highlight': self.value_of('highlight', self.highlight, '高亮')}
        mark = self._choice(self.mark)
        if mark == 'follow':
            rules['mark'] = {key: rules[key] for key in
                             ('font_east', 'font_ascii', 'size_pt', 'bold')
                             if rules.get(key) is not None}
        elif mark == 'plain':
            rules['mark'] = {'color': '000000', 'bold': False}
        pagination = self._choice(self.pagination)
        if pagination:
            name, value = pagination
            if name == 'keep_next_keep_lines':
                rules['keep_next'] = rules['keep_lines'] = True
            else:
                rules[name] = value
        return {'key': self.key, 'scope': self.scope.currentData(),
                'rules': {k: v for k, v in rules.items() if v is not None}}

    def _align_value(self):
        checked = self.align_group.checkedButton()
        if checked is None:
            return None
        for value, index in self.align_buttons.items():
            if self.align_group.buttons()[index] is checked:
                return value
        return None


class PageGroup(QFrame):
    """页面设置：纸张、横纵向、四边页边距（cm），会写到文档所有分节。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('groupcard')
        box = QVBoxLayout(self)
        box.setContentsMargins(10, 10, 10, 12)
        box.setSpacing(8)
        box.addWidget(label('页面设置', 'section'))
        row = QHBoxLayout()
        row.setSpacing(6)
        row.addWidget(label('纸张', 'muted'))
        self.paper = narrow_combo(QComboBox())
        for name in layout.PAPERS:
            self.paper.addItem(name, name)
        self.paper.addItem('自定义', '自定义')
        row.addWidget(self.paper, 1)
        self.landscape = QCheckBox('横向')
        row.addWidget(self.landscape)
        box.addLayout(row)
        preset_row = QHBoxLayout()
        preset_row.setSpacing(6)
        preset_row.addWidget(label('边距', 'muted'))
        self.preset = narrow_combo(QComboBox())
        self.preset.addItem('自定义', None)
        for name in layout.MARGIN_PRESETS:
            self.preset.addItem(name, name)
        preset_row.addWidget(self.preset, 1)
        box.addLayout(preset_row)
        grid = QGridLayout()
        grid.setSpacing(6)
        self.spins = {}
        for position, (side, text) in enumerate((('top', '上'), ('bottom', '下'),
                                                 ('left', '左'), ('right', '右'))):
            spin = QDoubleSpinBox()
            spin.setRange(0.0, 30.0)
            spin.setSingleStep(0.1)
            spin.setDecimals(2)
            spin.setSuffix(' cm')
            spin.setMaximumWidth(96)
            spin.setValue(2.54)
            spin.valueChanged.connect(lambda _value, key=side: self._margin_touched(key))
            cell = QHBoxLayout()
            cell.setSpacing(4)
            cell.addWidget(label(text, 'muted'))
            cell.addWidget(spin, 1)
            grid.addLayout(cell, position // 2, position % 2)
            self.spins[side] = spin
        box.addLayout(grid)
        box.addWidget(label('纸张与页边距改的是文档每一节的版式，页数可能随之变化。', 'muted'))
        self.preset.currentIndexChanged.connect(self._use_preset)

    def _use_preset(self, _index):
        name = self.preset.currentData()
        if not name:
            return
        for side, value in layout.MARGIN_PRESETS[name].items():
            self.spins[side].setValue(value)

    def _margin_touched(self, _side):
        self.preset.setCurrentIndex(0)

    def state(self):
        return {'paper': self.paper.currentData(),
                'landscape': self.landscape.isChecked(),
                'margins': {side: round(spin.value(), 2) for side, spin in self.spins.items()}}

    def set_state(self, page):
        page = page or {}
        pick_by_data(self.paper, page.get('paper'))
        self.landscape.setChecked(bool(page.get('landscape')))
        margins = page.get('margins') or {}
        matched = None
        for name, preset in layout.MARGIN_PRESETS.items():
            if all(abs(float(margins.get(side) or -1) - preset[side]) < 0.05
                   for side in self.spins):
                matched = name
                break
        for side, spin in self.spins.items():
            if side in margins:
                spin.setValue(float(margins[side]))
        pick_by_data(self.preset, matched)


class LayoutPage(DropToOpen, QWidget):
    """自动排版工作台：左排版规则、中整页预览、右样式库，应用后导出新文件。"""

    def __init__(self, store, workspace):
        super().__init__()
        self.store, self.workspace = store, workspace
        self.source = None
        self.output = None
        self.pdf = b''
        self.page_count = 0
        self.page_index = 0
        self.zoom = 1.0
        self.history = []
        self.pointer = -1
        self._worker = None
        self._template_signature = None
        self._preview_worker = None
        self._think_worker = None
        self.heading_levels = {}
        self.template_config = None
        self._then_export = False
        self._cancel = False
        self.work_dir = store.root / 'layout'
        self.groups = {}
        self.init_drops(self.set_document)
        self.build_ui()

    # ---------- 界面 ----------
    def build_ui(self):
        shell = QVBoxLayout(self)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)
        shell.addWidget(self.build_toolbar())
        body = QHBoxLayout()
        body.setContentsMargins(20, 12, 20, 12)
        body.setSpacing(12)
        shell.addLayout(body, 1)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.addWidget(self.build_rules_panel())
        self.splitter.addWidget(self.build_preview_panel())
        self.splitter.addWidget(self.build_style_panel())
        self.splitter.setSizes([352, 570, 226])
        self.splitter.splitterMoved.connect(lambda *_: self.render_preview())
        body.addWidget(self.splitter, 1)

    def build_toolbar(self):
        bar = QFrame()
        bar.setObjectName('panel')
        bar.setFixedHeight(54)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(20, 9, 20, 9)
        lay.setSpacing(8)
        self.template_combo = narrow_combo(QComboBox())
        self.template_combo.setMinimumWidth(190)
        self.template_combo.setToolTip('从模板库选一个模板，直接在这里整篇排版（会带上它读到过的要求）')
        self.template_combo.currentIndexChanged.connect(self.template_selected)
        lay.addWidget(self.template_combo)
        lay.addWidget(button('打开文档…', self.choose_document))
        self.file_chip = label('还没有选择文档（也可以把 .docx 直接拖进来）', 'chip')
        self.file_chip.setWordWrap(False)
        self.file_chip.setMinimumWidth(220)
        lay.addWidget(self.file_chip, 1)
        self.clear_btn = icon_button('close', '清除当前文档', self.clear_document, side=32)
        self.clear_btn.setEnabled(False)
        lay.addWidget(self.clear_btn)
        self.undo_btn = icon_button('undo', '撤销上一步排版', self.undo)
        self.undo_btn.setEnabled(False)
        lay.addWidget(self.undo_btn)
        self.redo_btn = icon_button('redo', '重做被撤销的排版', self.redo)
        self.redo_btn.setEnabled(False)
        lay.addWidget(self.redo_btn)
        self.export_btn = button('导出 Word', self.export_document)
        lay.addWidget(self.export_btn)
        return bar

    def build_rules_panel(self):
        panel = QFrame()
        panel.setObjectName('panel')
        panel.setMinimumWidth(316)
        box = QVBoxLayout(panel)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(0)
        box.addWidget(label('排版规则设置', 'paneltitle'))
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        host.setObjectName('formpage')
        inner = QVBoxLayout(host)
        inner.setContentsMargins(10, 10, 10, 10)
        inner.setSpacing(10)
        self.page_group = PageGroup()
        inner.addWidget(self.page_group)
        for key in layout.BLOCK_KEYS:
            group = RuleGroup(key, layout.BLOCK_NAMES[key])
            self.groups[key] = group
            inner.addWidget(group)
        inner.addStretch(1)
        scroll.setWidget(host)
        box.addWidget(scroll, 1)
        tail = QVBoxLayout()
        tail.setContentsMargins(12, 8, 12, 12)
        tail.setSpacing(6)
        self.think_btn = button('AI 识别标题层级（思考模型）', self.infer_headings_now)
        self.think_btn.setToolTip('用思考模型（默认本机 qwen3:4b）判断哪些段落是几级标题，'
                                  '补上“手动加粗放大、没有 Word 样式”的假标题。')
        tail.addWidget(self.think_btn)
        self.apply_btn = button('应用排版到文档', self.apply_now, primary=True)
        tail.addWidget(self.apply_btn)
        self.task = feedback.TaskBar()
        self.task.cancelled.connect(self.cancel_apply)
        tail.addWidget(self.task)
        self.status_label = label('选好文档、勾好规则，点“应用排版到文档”。原文件不会被改动。', 'muted')
        tail.addWidget(self.status_label)
        box.addLayout(tail)
        return panel

    def build_preview_panel(self):
        panel = QFrame()
        panel.setObjectName('panel')
        panel.setMinimumWidth(300)
        box = QVBoxLayout(panel)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(0)
        head = QHBoxLayout()
        head.setContentsMargins(14, 10, 12, 4)
        head.setSpacing(8)
        head.addWidget(label('整页预览', 'section'))
        head.addStretch(1)
        self.page_label = label('', 'muted')
        head.addWidget(self.page_label)
        box.addLayout(head)
        area = QFrame()
        area.setObjectName('previewarea')
        al = QVBoxLayout(area)
        al.setContentsMargins(12, 12, 12, 12)
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(True)
        self.preview_image = QLabel('选一个 .docx，应用排版后在这里看整页效果。')
        self.preview_image.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.preview_image.setMinimumWidth(1)
        self.preview_image.setWordWrap(True)
        self.preview_scroll.setWidget(self.preview_image)
        al.addWidget(self.preview_scroll)
        box.addWidget(area, 1)
        bar = QHBoxLayout()
        bar.setContentsMargins(14, 6, 14, 10)
        bar.setSpacing(6)
        self.fit_width = QCheckBox('适应宽度')
        self.fit_width.setChecked(True)
        self.fit_width.toggled.connect(lambda *_: self.render_preview())
        bar.addWidget(self.fit_width)
        bar.addStretch(1)
        self.prev = icon_button('up', '上一页', lambda: self.move_page(-1))
        self.next = icon_button('down', '下一页', lambda: self.move_page(1))
        self.prev.setEnabled(False)
        self.next.setEnabled(False)
        zoom_out = icon_button('zoom-out', '缩小预览', lambda: self.set_zoom(self.zoom / 1.25))
        zoom_in = icon_button('zoom-in', '放大预览', lambda: self.set_zoom(self.zoom * 1.25))
        self.zoom_label = label('适应宽度', 'muted')
        self.zoom_label.setMinimumWidth(46)
        self.zoom_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        bar.addWidget(self.prev)
        bar.addWidget(self.next)
        bar.addSpacing(8)
        bar.addWidget(zoom_out)
        bar.addWidget(self.zoom_label)
        bar.addWidget(zoom_in)
        box.addLayout(bar)
        return panel

    def build_style_panel(self):
        panel = QFrame()
        panel.setObjectName('panel')
        panel.setMinimumWidth(200)
        box = QVBoxLayout(panel)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(0)
        box.addWidget(label('样式库', 'paneltitle'))
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        host = QWidget()
        host.setObjectName('formpage')
        inner = QVBoxLayout(host)
        inner.setContentsMargins(12, 12, 12, 12)
        inner.setSpacing(8)
        self.style_host = host
        self.style_list = inner
        self.style_cards = {}
        self.render_style_cards()
        inner.addStretch(1)
        scroll.setWidget(host)
        box.addWidget(scroll, 1)
        tail = QVBoxLayout()
        tail.setContentsMargins(12, 6, 12, 12)
        tail.addWidget(button('新建样式', self.new_style))
        tail.addWidget(label('点一张卡就把参数灌进左栏对应档，参数仍可继续改；'
                             '“新建样式”把左栏当前这一档的参数存成自己的样式。', 'muted'))
        box.addLayout(tail)
        return panel

    def presets(self):
        return list(layout.STYLE_PRESETS) + self.custom_styles()

    def custom_styles(self):
        path = self.store.root / 'styles_custom.json'
        if not path.is_file():
            return []
        try:
            data = read_json(path)
        except (OSError, ValueError):
            return []
        return data if isinstance(data, list) else []

    def render_style_cards(self):
        while self.style_list.count():
            item = self.style_list.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
                widget.deleteLater()
        self.style_cards = {}
        for preset in self.presets():
            card = QFrame()
            card.setObjectName('stylecard')
            card.setCursor(Qt.CursorShape.PointingHandCursor)
            row = QVBoxLayout(card)
            row.setContentsMargins(12, 10, 12, 10)
            row.setSpacing(4)
            row.addWidget(label(preset['name'], 'stylename'))
            row.addWidget(label(preset.get('summary') or style_summary(preset.get('rules') or {}),
                                'stylemeta'))
            actions = QHBoxLayout()
            actions.setSpacing(6)
            actions.addWidget(mini_button('用到左栏', lambda _checked=False, pid=preset['id']:
                                          self.use_preset(pid)))
            if preset.get('custom'):
                actions.addWidget(mini_button('删除', lambda _checked=False, pid=preset['id']:
                                              self.delete_style(pid)))
            actions.addStretch()
            row.addLayout(actions)
            self.style_list.addWidget(card)
            self.style_cards[preset['id']] = card

    # ---------- 文档 ----------
    def choose_document(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择要排版的 Word 文档', '',
                                              'Word 文档 (*.docx)')
        if path:
            self.set_document(path)

    def clear_document(self):
        self.set_document(None)

    def set_document(self, path):
        if path:
            source = Path(path)
            if source.suffix.lower() != '.docx':
                QMessageBox.information(self, '需要 .docx 文件',
                                        '只支持 .docx。旧版 .doc 请先用 Word 另存为 .docx。')
                return
            if not source.is_file():
                self.workspace.warn_box(f'文件不存在：\n{source}')
                return
        self.source = Path(path) if path else None
        self.output = None
        self.pdf = b''
        self.page_count = 0
        self.page_index = 0
        self.history = [{'state': None, 'output': str(self.source) if self.source else None,
                         'pdf': b'', 'pages': 0}]
        self.pointer = 0
        self.heading_levels = {}
        self.template_config = None
        self.file_chip.setText(self.source.name if self.source
                               else '还没有选择文档（也可以把 .docx 直接拖进来）')
        self.clear_btn.setEnabled(bool(self.source))
        self.update_history_buttons()
        self.render_preview()
        if self.source:
            self.load_page_setup()
            QTimer.singleShot(0, self.preview_source)
        self.status_label.setText('原文件不会被改动；应用排版会在工作目录生成一份新文件。'
                                  if self.source else '选好文档、勾好规则，点“应用排版到文档”。')

    def preview_source(self):
        """载入文档后先渲染原文，面板里立刻有画面（还没套用任何规则）。"""
        if not self.source:
            return
        if self._preview_worker is not None and self._preview_worker.isRunning():
            return
        source = self.source
        work = self.work_dir / (source.stem + '_原文.pdf')
        self.status_label.setText('正在渲染原文预览…')
        worker = AIWorker(lambda: preview.export_pdf_sync(source, work), self)
        self._preview_worker = worker

        def done(result):
            self._preview_worker = None
            if not result:
                self.status_label.setText('原文预览为空，请确认文档里有内容。')
                return
            if self.pdf and self.output:
                return          # 已经应用过排版，别用原文覆盖
            self.pdf = bytes(result)
            self.page_count = preview.page_count(self.pdf)
            self.page_index = 0
            self.render_preview()
            self.status_label.setText('预览：文档原文（尚未套用规则）；点“应用排版到文档”看成排后的效果。')

        worker.done.connect(done)
        worker.failed.connect(
            lambda message: (setattr(self, '_preview_worker', None),
                             self.status_label.setText('原文预览需要本机 Word/WPS：' + str(message)[:110])))
        worker.start()

    def template_selected(self, index):
        template_id = self.template_combo.itemData(index)
        if not template_id:
            return
        config = next((c for c in self.store.templates() if c['id'] == template_id), None)
        if config:
            self.load_template(config)

    def load_template(self, config):
        """从填写页/模板库把模板带进来：源文件 + 它已解析的要求，一起用于排版与删说明。"""
        if not config:
            return
        self.template_config = dict(config)
        self.set_document(Path(config['_dir']) / config['file'])
        self.template_config = dict(config)
        index = self.template_combo.findData(config.get('id'))
        if index > 0:
            self.template_combo.blockSignals(True)
            self.template_combo.setCurrentIndex(index)
            self.template_combo.blockSignals(False)
        self.file_chip.setText(config.get('title') or self.file_chip.text())
        rules = len((config.get('req') or {}).get('notes') or [])
        self.status_label.setText(
            f'已载入模板「{config.get("title", "")}」，读到说明 {rules} 条；'
            '点“应用排版到文档”会先删说明再把格式落到文档。' if rules else
            f'已载入模板「{config.get("title", "")}」；点“应用排版到文档”时会补读模板要求。')

    def take_document(self, path):
        """其它页面把文档交过来：切到本页并载入。"""
        self.set_document(path)
        self.workspace.nav.setCurrentRow(Workspace.LAYOUT)

    def load_page_setup(self):
        try:
            from docx import Document
            page = layout.read_page(Document(str(self.source)))
        except Exception as exc:
            self.status_label.setText('读不到页面设置：' + str(exc)[:80])
            return
        self.page_group.set_state(page)

    # ---------- 规则状态 ----------
    def collect_state(self):
        return {'page': self.page_group.state(),
                'blocks': [group.state() for group in self.groups.values()]}

    def apply_state(self, state):
        self.page_group.set_state((state or {}).get('page'))
        for key, group in self.groups.items():
            block = next((b for b in (state or {}).get('blocks') or [] if b['key'] == key), None)
            group.set_scope((block or {}).get('scope') or 'off')
            group.set_rules((block or {}).get('rules') or {})

    def use_preset(self, preset_id):
        preset = next((p for p in self.presets() if p['id'] == preset_id), None)
        key = (preset or {}).get('block')
        if key not in self.groups:
            return
        group = self.groups[key]
        group.set_scope('all')
        group.set_rules(dict(preset.get('rules') or {}))
        for pid, card in self.style_cards.items():
            mark_selected(card, pid == preset_id)
        self.status_label.setText(f'已把「{preset["name"]}」的参数灌进左栏“{layout.BLOCK_NAMES[key]}”'
                                  f'档，可以继续改，然后点“应用排版到文档”。')

    def new_style(self):
        """把左栏当前这一档的参数存成自己的样式，写在本机数据目录。"""
        from PySide6.QtWidgets import QInputDialog
        keys = [key for key in layout.BLOCK_KEYS if key in self.groups]
        names = [layout.BLOCK_NAMES[key] for key in keys]
        filled = next((index for index, key in enumerate(keys)
                       if self.groups[key].state()['rules']), 0)
        choice, ok = QInputDialog.getItem(self, '新建样式', '取哪一档的当前参数：',
                                          names, filled, False)
        if not ok:
            return
        key = keys[names.index(choice)]
        rules = self.groups[key].state()['rules']
        if not rules:
            QMessageBox.information(self, '这一档还没参数',
                                    f'“{choice}”档还是“不改”，先设几个参数再存样式。')
            return
        name, ok = QInputDialog.getText(self, '新建样式', '样式名称：',
                                        QLineEdit.EchoMode.Normal, f'我的{choice}')
        if not ok or not name.strip():
            return
        items = self.custom_styles()
        preset = {'id': 'custom-' + uuid.uuid4().hex[:8], 'name': name.strip(), 'block': key,
                  'summary': style_summary(rules), 'rules': rules, 'custom': True}
        items.append(preset)
        write_json(self.store.root / 'styles_custom.json', items)
        self.render_style_cards()
        self.use_preset(preset['id'])

    def delete_style(self, preset_id):
        items = self.custom_styles()
        target = next((p for p in items if p['id'] == preset_id), None)
        if target is None:
            return
        answer = QMessageBox.question(
            self, '删除样式', f'删除自定义样式「{target["name"]}」？\n内置样式不受影响。',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if answer != QMessageBox.StandardButton.Yes:
            return
        write_json(self.store.root / 'styles_custom.json',
                   [p for p in items if p['id'] != preset_id])
        self.render_style_cards()
        self.status_label.setText(f'已删除样式「{target["name"]}」。')

    # ---------- AI 标题层级 ----------
    def infer_headings_now(self):
        if not self.source:
            QMessageBox.information(self, '还没有文档', '先点“打开文档…”选一个 .docx。')
            return
        if self._think_worker is not None and self._think_worker.isRunning():
            return
        try:
            from docx import Document
            items = layout.paragraph_outline(Document(str(self.source)))
        except Exception as exc:
            self.status_label.setText('读取段落失败：' + str(exc)[:120])
            return
        if not items:
            self.status_label.setText('文档里没有可分析的正文段落。')
            return
        cfg = ai.load_config(self.store.ai_config)
        title = Path(self.source).stem
        self.think_btn.setEnabled(False)
        self.status_label.setText(
            f'思考模型正在判断 {len(items)} 个段落的标题层级（本机 qwen3:4b 首次约 30～90 秒）…')

        def job():
            return ai.infer_headings(ai.think_client(cfg), items, title)

        worker = AIWorker(job, self)
        self._think_worker = worker
        worker.done.connect(self.headings_ready)
        worker.failed.connect(self.headings_failed)
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def headings_ready(self, result):
        levels, note = result if isinstance(result, tuple) else (result or {}, '')
        self._think_worker = None
        self.think_btn.setEnabled(True)
        self.heading_levels = dict(levels or {})
        if not self.heading_levels:
            self.status_label.setText('思考模型没找到明确标题。可直接用样式判定，或手动在 Word 里设标题样式。')
            return
        counts = {1: 0, 2: 0, 3: 0}
        for level in self.heading_levels.values():
            counts[int(level)] = counts.get(int(level), 0) + 1
        self.status_label.setText(
            f'思考模型判定：一级 {counts.get(1, 0)} 个 · 二级 {counts.get(2, 0)} 个 · '
            f'三级 {counts.get(3, 0)} 个' + (f'（{note}）' if note else '')
            + '。已自动重新应用排版。')
        if any(block.get('scope') != 'off' for block in self.collect_state()['blocks']):
            self.apply_now()

    def headings_failed(self, message):
        self._think_worker = None
        self.think_btn.setEnabled(True)
        self.status_label.setText('标题层级识别失败：' + str(message)[:160])

    # ---------- 应用 / 撤销 ----------
    def apply_now(self, then_export=False):
        if not self.source:
            QMessageBox.information(self, '还没有文档', '先点“打开文档…”选一个 .docx。')
            return
        if self._worker is not None and self._worker.isRunning():
            return
        state = self.collect_state()
        active = [b for b in state['blocks'] if b['scope'] != 'off']
        if not active:
            self.status_label.setText('所有档都是“不改”：至少勾一档，或点右侧样式库的卡。')
            return
        self._then_export = then_export
        self._cancel = False
        self.set_busy(True, '正在排版并渲染整页预览…')
        self.task.start('正在排版并渲染整页预览…')
        step = len(self.history)
        source = self.source
        out = self.work_dir / f'{source.stem}_排版_{step}.docx'

        def job():
            spec = options = None
            if self.template_config:
                if not (self.template_config.get('req') or {}).get('notes'):
                    backfill_requirements(self.store, self.template_config)
                spec = self.template_config.get('req')
                options = dict(REQ_OPTIONS)
            try:
                report = layout.apply_to_file(source, out, state['blocks'], state['page'],
                                              self.heading_levels, spec=spec, options=options,
                                              should_stop=lambda: self._cancel)
            except layout.Cancelled as exc:
                return {'cancelled': str(exc)}
            if self._cancel:
                Path(out).unlink(missing_ok=True)
                return {'cancelled': '已取消：本次排版没有写入，原文件未改动。'}
            self.task.step(1, 2, '排版完成，正在渲染整页预览…')
            try:
                report['pdf'] = preview.export_pdf_sync(out, Path(out).with_suffix('.pdf'))
            except Exception as exc:
                report['pdf'] = None
                report['preview_error'] = str(exc)[:120]
            return {'report': report, 'state': state}

        self._worker = AIWorker(job, self)
        self._worker.done.connect(self.apply_done)
        self._worker.failed.connect(self.apply_failed)
        self._worker.start()

    def cancel_apply(self):
        self._cancel = True

    def apply_done(self, payload):
        self.task.finish()
        if payload.get('cancelled'):
            self._then_export = False
            self.set_busy(False)
            self.status_label.setText(payload['cancelled'])
            self.workspace.notify(payload['cancelled'], 'info', 2600)
            return
        report = payload['report']
        entry = {'state': payload['state'], 'output': report['output'],
                 'pdf': report.get('pdf') or b'', 'pages': preview.page_count(report.get('pdf'))}
        self.history = self.history[:self.pointer + 1] + [entry]
        self.pointer = len(self.history) - 1
        self.show_entry(entry)
        self.set_busy(False)
        bits = [f'段落 {report.get("paragraphs", 0)}', f'文字 {report.get("runs", 0)}',
                f'行距 {report.get("spacing", 0)}', f'缩进 {report.get("indent", 0)}',
                f'对齐 {report.get("align", 0)}']
        if report.get('sections'):
            bits.append(f'分节 {report.get("sections")}')
        if report.get('removed'):
            bits.append(f'删说明 {report.get("removed")}')
        if report.get('blank_lines'):
            bits.append(f'空行 {report.get("blank_lines")}')
        text = '已应用：' + ' · '.join(bits)
        if report.get('pagination'):
            text += f' · 分页 {report["pagination"]}'
        if report.get('shading'):
            text += f' · 底纹 {report["shading"]}'
        if report.get('mark'):
            text += f' · 回车继承 {report["mark"]}'
        if report.get('highlight'):
            text += f' · 高亮 {report["highlight"]}'
        try:
            issues = ooxmlcheck.check(report['output'])
            errors = [issue for issue in issues
                      if issue['level'] == ooxmlcheck.LEVEL_ERROR]
            text += ('\n自检：' + ('通过' if not issues else
                                  f'{len(errors)} 处结构错误 / {len(issues) - len(errors)} 条提醒'))
            if errors:
                text += f'｜{errors[0]["where"]}：{errors[0]["message"]}'
        except Exception:
            pass
        if report.get('skipped'):
            text += '\n跳过：' + '；'.join(report['skipped'])
        if report.get('preview_error'):
            text += '\n预览需要本机 Word/WPS：' + report['preview_error']
        self.status_label.setText(text)
        if self._then_export:
            self._then_export = False
            self.export_document()

    def apply_failed(self, message):
        self._then_export = False
        self.task.finish()
        self.set_busy(False)
        self.status_label.setText('排版失败：' + message[:160])
        advice = error_advice(message, self.source)
        QMessageBox.warning(self, '排版失败',
                            message + ('\n\n' + advice if advice else ''))

    def show_entry(self, entry):
        self.output = Path(entry['output']) if entry.get('output') else None
        self.pdf = entry.get('pdf') or b''
        self.page_count = entry.get('pages') or 0
        self.page_index = min(self.page_index, max(0, self.page_count - 1))
        if entry.get('state'):
            self.apply_state(entry['state'])
        self.update_history_buttons()
        self.render_preview()

    def set_busy(self, busy, text=''):
        self.apply_btn.setEnabled(not busy)
        self.export_btn.setEnabled(not busy and bool(self.source))
        self.undo_btn.setEnabled(not busy and self.pointer > 0)
        self.redo_btn.setEnabled(not busy and self.pointer < len(self.history) - 1)
        if busy:
            self.status_label.setText(text)

    def update_history_buttons(self):
        self.undo_btn.setEnabled(self.pointer > 0)
        self.redo_btn.setEnabled(self.pointer < len(self.history) - 1)

    def undo(self):
        self.step_history(-1)

    def redo(self):
        self.step_history(1)

    def step_history(self, offset):
        target = self.pointer + offset
        if target < 0 or target >= len(self.history):
            return
        self.pointer = target
        self.show_entry(self.history[target])
        self.status_label.setText('已回到未排版状态（导出会得到原样副本）。' if target == 0
                                  else f'已切换到第 {target} 次排版结果。')

    # ---------- 导出与预览 ----------
    def export_document(self):
        if not self.source:
            QMessageBox.information(self, '还没有文档', '先选一个 .docx。')
            return
        if not self.output or not Path(self.output).is_file():
            self.apply_now(then_export=True)
            return
        suggested = Path(self.source).with_name(self.source.stem + '_已排版.docx')
        path, _ = QFileDialog.getSaveFileName(self, '导出排版后的 Word 文档', str(suggested),
                                              'Word 文档 (*.docx)')
        if not path:
            return
        try:
            shutil.copyfile(str(self.output), path)
        except OSError as exc:
            QMessageBox.warning(self, '导出失败', str(exc))
            return
        self.status_label.setText('已导出：' + str(path))
        QMessageBox.information(self, '导出完成',
                                f'排版后的文件已保存到：\n{path}\n\n原文件没有被改动。')

    def set_zoom(self, factor):
        self.zoom = min(3.0, max(0.5, factor))
        self.fit_width.setChecked(False)
        self.render_preview()

    def move_page(self, offset):
        if self.page_count <= 0:
            return
        self.page_index = max(0, min(self.page_index + offset, self.page_count - 1))
        self.render_preview()

    def render_preview(self):
        if not self.pdf:
            self.preview_image.setPixmap(QPixmap())
            self.preview_image.setText('还没有可预览的排版结果。\n'
                                       '点“应用排版到文档”后，这里会显示整页效果。'
                                       if self.source else
                                       '选一个 .docx，应用排版后在这里看整页效果。')
            self.page_label.setText('')
            self.prev.setEnabled(False)
            self.next.setEnabled(False)
            self.zoom_label.setText('适应宽度')
            return
        try:
            png = preview.render_page(self.pdf, self.page_index, scale=2.0)
        except Exception as exc:
            self.preview_image.setText('预览页渲染失败：' + str(exc)[:80])
            return
        pixmap = QPixmap()
        if not pixmap.loadFromData(png, 'PNG'):
            self.preview_image.setText('预览页渲染失败。')
            return
        base = max(160, self.preview_scroll.viewport().width() - 16)
        width = base if self.fit_width.isChecked() else max(160, int(base * self.zoom))
        self.preview_image.setMinimumWidth(width)
        self.preview_image.setPixmap(
            pixmap.scaledToWidth(width, Qt.TransformationMode.SmoothTransformation))
        self.page_label.setText(f'第 {self.page_index + 1} / {self.page_count} 页')
        self.prev.setEnabled(self.page_index > 0)
        self.next.setEnabled(self.page_index < self.page_count - 1)
        self.zoom_label.setText('适应宽度' if self.fit_width.isChecked()
                                else f'{int(round(self.zoom * 100))}%')

    def reload(self):
        self.export_btn.setEnabled(bool(self.source))
        templates = self.store.templates()
        # 切页会走到这里：模板没变就不要重建下拉框，重建会丢焦点还会有一次闪烁
        signature = [(config['id'], config['title']) for config in templates]
        if signature != self._template_signature:
            self._template_signature = signature
            self.template_combo.blockSignals(True)
            self.template_combo.clear()
            self.template_combo.addItem('从模板库选择…', None)
            for config in templates:
                self.template_combo.addItem(config['title'], config['id'])
            self.template_combo.blockSignals(False)
        if self.template_config:
            index = self.template_combo.findData(self.template_config.get('id'))
            if index > 0:
                self.template_combo.blockSignals(True)
                self.template_combo.setCurrentIndex(index)
                self.template_combo.blockSignals(False)

    def shutdown(self):
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait(3000)


# ============================================================ 主窗体
class Workspace(QMainWindow):
    PAGES = ('首页', '填写生成', '模板制作', '自动排版', '文献检索')
    HOME, FILL, STUDIO, LAYOUT, LIT = range(5)

    def __init__(self, store):
        super().__init__()
        self.store = store
        self.theme = load_theme_name(store.root / THEME_KEY)
        self._ollama_worker = None
        self._ollama_tried = False
        self.ollama_state = ''
        self.ai_client = ai.Client(ai.load_config(store.ai_config))
        self.embedded = embedded_ai.shared(store.root / 'embedded_ai.log')
        self.setWindowTitle('Word 模板工作台')
        icon = svg_icon(ASSETS / 'icon.svg')
        if icon is not None:
            self.setWindowIcon(icon)
        else:
            self.setWindowIcon(QIcon(str(ASSETS / 'icon.svg')))
        self.setMinimumSize(1080, 680)
        screen = QApplication.primaryScreen().availableGeometry()
        self.resize(min(1420, screen.width() - 80), min(920, screen.height() - 80))
        central = QWidget()
        self.setCentralWidget(central)
        layers = QStackedLayout(central)
        layers.setStackingMode(QStackedLayout.StackingMode.StackAll)
        layers.setContentsMargins(0, 0, 0, 0)
        # 背景层在最底下，内容层透明处透出静态玻璃；着色器只渲一次，切页不再碰 GPU。
        self.backdrop = glassfx.StaticGlassBackdrop()
        layers.addWidget(self.backdrop)
        content = QWidget()
        content.setObjectName('glasslayer')
        layers.addWidget(content)
        # StackAll 会把“当前页”抬到最上层，默认是第一个（背景层），必须显式设成内容层
        layers.setCurrentWidget(content)
        shell = QVBoxLayout(content)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)
        shell.addWidget(self.build_topbar())
        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)
        shell.addLayout(body)
        side = QFrame()
        side.setObjectName('sidebar')
        side.setFixedWidth(210)
        left = QVBoxLayout(side)
        left.setContentsMargins(16, 18, 16, 16)
        left.addWidget(label('导航', 'sidehint'))
        left.addSpacing(6)
        self.nav = QListWidget()
        self.nav.setObjectName('nav')
        self.nav.setToolTip('Alt+1 … Alt+5 可以直接切页。')
        for name in self.PAGES:
            self.nav.addItem(QListWidgetItem(name))
        self.nav.currentRowChanged.connect(self.switch_page)
        left.addWidget(self.nav, 1)
        self.ai_status = label('', 'sidehint')
        self.ai_status.setWordWrap(True)
        left.addWidget(self.ai_status)
        left.addSpacing(4)
        left.addWidget(label(f'桌面联网版 · v{VERSION}', 'sidehint'))
        body.addWidget(side)
        self.stack = QStackedWidget()
        self.home = HomePage(store, self)
        self.fill = FillPage(store, self)
        self.studio = StudioPage(store, self)
        self.mark = self.studio.mark
        self.inspect = self.studio.inspect
        self.library = LibraryPage(store, self)     # 不在导航里，由首页卡片弹窗承载
        self.library_dialog = None
        self.layout_page = LayoutPage(store, self)
        self.literature = literature_page.LiteraturePage(store, self)
        for page in (self.home, self.fill, self.studio, self.layout_page, self.literature):
            self.stack.addWidget(page)
        body.addWidget(self.stack, 1)
        # 样式表在 main() 里创建窗口前已经设好，这里再设一遍会让所有控件重新 polish
        # （实测要好几秒），所以只把背景和图标配色对齐到本机主题。
        self.backdrop.set_dark(self.theme == 'dark')
        uiicons.set_theme(THEMES[self.theme])
        self.nav.setCurrentRow(self.HOME)
        self.refresh_ai_status()
        QShortcut(QKeySequence.StandardKey.Save, self,
                  activated=lambda: self.fill.save_draft(quiet=False))
        self.toast = feedback.Toast(self)
        self.build_shortcuts()
        if store.warnings:
            QTimer.singleShot(300, self.show_warnings)

    def build_shortcuts(self):
        """键盘入口：切页、聚焦搜索、直接生成、使用说明。都在窗口级注册，任何页面都能按。"""
        QShortcut(QKeySequence.StandardKey.Find, self, activated=self.focus_search)
        QShortcut(QKeySequence.StandardKey.HelpContents, self,
                  activated=lambda: self.show_help(self))
        QShortcut(QKeySequence('Ctrl+Return'), self, activated=self.generate_now)
        for index in range(len(self.PAGES)):
            QShortcut(QKeySequence(f'Alt+{index + 1}'), self,
                      activated=lambda _=False, n=index: self.nav.setCurrentRow(n))

    def focus_search(self):
        self.search.setFocus()
        self.search.selectAll()

    def blur_search(self):
        """搜索框里的 Esc：先清空关键词，已经空了才把焦点交还界面。"""
        if self.search.text():
            self.search.clear()
        else:
            self.search.clearFocus()

    def generate_now(self):
        if self.stack.currentWidget() is not self.fill:
            return
        self.fill.generate()

    def notify(self, text, kind='info', ms=1800):
        """右下角轻提示：成功类反馈不再用必须点确定的弹框。"""
        self.toast.say(text, kind, ms)

    def build_topbar(self):
        bar = QFrame()
        bar.setObjectName('topbar')
        bar.setFixedHeight(56)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(18, 8, 18, 8)
        lay.setSpacing(10)
        logo = QLabel('W')
        logo.setObjectName('logo')
        mark = svg_pixmap(ASSETS / 'logo-ink-white.svg', 22)
        if mark is not None:
            logo.setText('')
            logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
            logo.setPixmap(mark)
        lay.addWidget(logo)
        lay.addWidget(label('Word 模板工作台', 'brand'))
        lay.addWidget(label(f'v{VERSION}', 'versionchip'))
        lay.addStretch(1)
        self.search = QLineEdit()
        self.search.setObjectName('search')
        self.search.setPlaceholderText('搜索模板名称 / 分类 / 标识（Ctrl+F）')
        self.search.setClearButtonEnabled(True)
        self.search.setFixedWidth(300)
        self.search.textChanged.connect(self.filter_templates)
        self.search.returnPressed.connect(self.open_library)
        QShortcut(QKeySequence(Qt.Key.Key_Escape), self.search,
                  activated=self.blur_search).setContext(
            Qt.ShortcutContext.WidgetWithChildrenShortcut)
        lay.addWidget(self.search)
        self.theme_btn = icon_button('sun' if self.theme == 'dark' else 'moon',
                                     '切换黑夜/白天配色，选择会记住在本机。',
                                     self.toggle_theme, side=34)
        lay.addWidget(self.theme_btn)
        self.ai_button = icon_button('sparkle', 'AI 设置：服务地址、密钥与模型',
                                     lambda: self.open_ai_settings(self), side=34)
        lay.addWidget(self.ai_button)
        self.help_button = icon_button('help', '使用说明',
                                       lambda: self.show_help(self), side=34)
        lay.addWidget(self.help_button)
        return bar

    def set_theme(self, name):
        self.theme = apply_theme(QApplication.instance(), name)
        save_theme_name(self.store.root / THEME_KEY, self.theme)
        uiicons.rename(self.theme_btn, 'sun' if self.theme == 'dark' else 'moon')
        self.theme_btn.setToolTip('切换到白天配色，选择会记住在本机。' if self.theme == 'dark'
                                  else '切换到黑夜配色，选择会记住在本机。')
        if hasattr(self, 'backdrop'):
            self.backdrop.set_dark(self.theme == 'dark')
        if hasattr(self, 'home'):
            self.home.reload()

    def toggle_theme(self):
        self.set_theme('dark' if self.theme == 'light' else 'light')

    def filter_templates(self, text):
        query = (text or '').strip()
        self.home.set_filter(query)
        self.library.set_filter(query)

    def inspect_template(self, template_id):
        self.open_library(template_id)

    def open_library(self, template_id=None):
        """模板库只在首页有入口：弹窗承载完整管理功能，可定位到某个模板。"""
        if self.library_dialog is None:
            self.library_dialog = LibraryDialog(self, self)
        self.library.reload()
        self.library_dialog.show()
        self.library_dialog.raise_()
        self.library_dialog.activateWindow()
        if template_id:
            self.library_dialog.select(template_id)
        return self.library_dialog

    def _text_cfg(self):
        cfg = ai.load_config(self.store.ai_config)
        return cfg if ai.Client(cfg).usable()[0] else None

    def _vision_cfg(self):
        """视觉只用来复核有疑问的栏目，不再拿整页截图做识别。"""
        cfg = ai.load_config(self.store.ai_config)
        return cfg if ai.vision_usable(cfg)[0] else None

    def refresh_ai_status(self):
        cfg = ai.load_config(self.store.ai_config)
        self.ai_client = ai.Client(cfg)
        ok, message = self.ai_client.usable()
        vok, vmessage = ai.vision_usable(cfg)
        lines = ['文本 AI：' + (self.ai_client.describe() if ok else '未配置')]
        if vok:
            lines.append('视觉 AI：' + (ai.vision_config(cfg).get('model') or ''))
        else:
            lines.append('视觉 AI：未配置')
        label = self.embedded.label()
        if label:
            lines.append('内置补全：' + label)
        extra = self.ollama_status_line(cfg)
        if extra:
            lines.append(extra)
        if not ok and not vok and not label:
            lines.append('点顶栏「AI 设置」接入模型')
        self.ai_status.setToolTip('\n'.join(lines + ([message] if not ok else [])
                                             + ([vmessage] if not vok else [])))
        self.ai_status.setText('\n'.join(lines))

    def uses_local_ollama(self, cfg):
        return ai.points_to_local_ollama(cfg)

    def ollama_status_line(self, cfg):
        """选了本机 Ollama 但服务没跑：后台把它拉起来，状态栏如实说明。"""
        if not self.uses_local_ollama(cfg):
            return ''
        if ai.ollama_running(cfg.get('base_url') or cfg.get('vision_base_url')):
            self.ollama_state = '运行中'
            return '本机 Ollama：运行中'
        if self._ollama_worker is not None and self._ollama_worker.isRunning():
            self.ollama_state = '正在启动'
            return '本机 Ollama：正在自动启动…'
        if not self._ollama_tried:
            self._ollama_tried = True
            self.ollama_state = '自动启动中'
            worker = AIWorker(ai.start_ollama, self)
            self._ollama_worker = worker
            worker.done.connect(self.ollama_autostart_done)
            worker.failed.connect(lambda _message: self.refresh_ai_status())
            worker.start()
            return '本机 Ollama：未运行，正在自动启动…'
        self.ollama_state = '未运行'
        return '本机 Ollama：未运行（用到 AI 时会自动拉起）'

    def ollama_autostart_done(self, result):
        self._ollama_worker = None
        ok, message = result if isinstance(result, tuple) else (False, str(result))
        self.ollama_state = '运行中' if ok else '未运行'
        if not ok and message:
            self.store.warnings.append('本机 Ollama 自动启动失败：' + message)
        self.refresh_ai_status()

    def open_ai_settings(self, parent=None):
        dialog = AISettingsDialog(ai.load_config(self.store.ai_config), parent or self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                ai.save_config(self.store.ai_config, dialog.cfg)
            except OSError as exc:
                QMessageBox.warning(parent or self, 'AI 设置未保存', f'无法写入本机数据目录：\n{exc}')
                return False
            self.refresh_ai_status()
            return True
        return False

    def ensure_ai(self, parent):
        self.ai_client = ai.Client(ai.load_config(self.store.ai_config))
        ok, message = self.ai_client.usable()
        if ok:
            return True
        answer = QMessageBox.question(
            parent, '需要先配置 AI',
            f'{message}。\n\n现在打开「AI 设置」吗？',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes)
        if answer != QMessageBox.StandardButton.Yes:
            return False
        if not self.open_ai_settings(parent):
            return False
        return self.ai_client.usable()[0]

    def show_help(self, parent=None):
        HelpDialog(self.store, parent or self).exec()

    def switch_page(self, row):
        if row == self.HOME:
            self.home.reload()
        elif row == self.LAYOUT:
            self.layout_page.reload()
        elif row == self.LIT:
            self.literature.on_show()
        self.stack.setCurrentIndex(row)

    def refresh_all(self):
        self.fill.reload()
        self.library.reload()
        self.home.reload()

    def go_fill(self, template_id):
        """首页点卡片进来：切到填写页并选中这个模板；被搜索筛掉就先清空筛选。"""
        if self.library_dialog is not None:
            self.library_dialog.close()
        self.nav.setCurrentRow(self.FILL)
        for attempt in (0, 1):
            for index, config in enumerate(self.fill.configs):
                if config['id'] != template_id:
                    continue
                item = self.fill.list.item(index)
                if item is not None and item.isHidden():
                    self.fill.search.clear()
                self.fill.list.setCurrentRow(index)
                return True
            if attempt == 0:
                self.fill.reload()
        self.inspect_template(template_id)   # 模板刚被删掉之类的原因，退而看详情
        return False

    def warn_box(self, text):
        QMessageBox.warning(self, '模板需要先配置', text)

    def show_warnings(self):
        warnings = self.store.warnings[:]
        self.store.warnings.clear()
        QMessageBox.warning(self, '部分本机数据需要检查', '\n'.join(warnings))

    def closeEvent(self, event):
        if self.fill._direct_worker and self.fill._direct_worker.isRunning():
            QMessageBox.information(self, '正在识别文档', '请等识别完成后再退出。')
            event.ignore()
            return
        if self.mark._ai_worker and self.mark._ai_worker.isRunning():
            QMessageBox.information(self, 'AI 识别进行中', '请等 AI 识别分类结束后再退出。')
            event.ignore()
            return
        if self.mark._name_worker and self.mark._name_worker.isRunning():
            QMessageBox.information(self, '内置模型补全中', '请等字段名补全结束后再退出。')
            event.ignore()
            return
        if self.layout_page._worker and self.layout_page._worker.isRunning():
            QMessageBox.information(self, '正在排版', '请等“应用排版到文档”结束后再退出。')
            event.ignore()
            return
        self.fill.generate_button.setFocus()
        QApplication.processEvents()
        if self.fill.save_draft():
            self.fill.shutdown_preview()
            self.layout_page.shutdown()
            self.embedded.stop()
            event.accept()
        else:
            event.ignore()


def main():
    argv = sys.argv[:]
    page_arg = ''
    if '--page' in argv:
        i = argv.index('--page')
        page_arg = argv[i + 1] if len(argv) > i + 1 else ''
        argv = argv[:i] + argv[i + 2:]
    test = '--smoke-test' in argv
    i = argv.index('--smoke-test') if test else -1
    root = Path(argv[i + 1]) if test else Path(
        QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppLocalDataLocation))
    sample = argv[i + 2] if test and len(argv) >= i + 3 else None
    no_splash = '--no-splash' in argv
    if no_splash:
        argv.remove('--no-splash')
    if test:
        import os
        os.environ['QT_QPA_PLATFORM'] = 'offscreen'
        QMessageBox.information = staticmethod(lambda *a, **k: None)
        QMessageBox.warning = staticmethod(lambda *a, **k: None)
        QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.StandardButton.No)
    app = QApplication(argv[:i] + argv[i + 3:] if test else argv)
    app.setOrganizationName('LocalTools')
    app.setApplicationName('WordTemplateTool')
    app.setStyle('Fusion')
    app.setFont(QFont('Microsoft YaHei UI', 10))
    theme_name = load_theme_name(root / THEME_KEY)
    if not test and not no_splash:
        splash.show(ASSETS, theme_name, VERSION)
    app.setStyleSheet(build_style(THEMES[theme_name]))
    uiicons.set_theme(THEMES[theme_name])
    try:
        win = Workspace(Store(root))
    except Exception as exc:
        QMessageBox.critical(None, '无法启动', f'无法读取本机模板或草稿目录：\n{exc}')
        return 1
    win.show()
    if page_arg and not test:
        pages = {
            'home': Workspace.HOME, '首页': Workspace.HOME,
            'fill': Workspace.FILL, '填写': Workspace.FILL, '填写生成': Workspace.FILL,
            'studio': Workspace.STUDIO, '模板': Workspace.STUDIO, '模板制作': Workspace.STUDIO,
            'layout': Workspace.LAYOUT, '排版': Workspace.LAYOUT, '自动排版': Workspace.LAYOUT,
            'literature': Workspace.LIT, '文献': Workspace.LIT, '文献检索': Workspace.LIT,
        }
        index = pages.get(page_arg.strip().lower())
        if index is not None:
            win.nav.setCurrentRow(index)
    if test:
        def smoke():
            try:
                root.mkdir(exist_ok=True, parents=True)
                win.fill.preview_timer.stop()
                win.fill.live_toggle.blockSignals(True)
                win.fill.live_toggle.setChecked(False)
                win.fill.live_toggle.blockSignals(False)
                win.nav.setCurrentRow(Workspace.HOME)
                win.grab().save(str(root / 'ui-home-light.png'))
                win.set_theme('dark')
                win.grab().save(str(root / 'ui-home-dark.png'))
                win.set_theme('light')
                win.literature._net_checked = True
                win.nav.setCurrentRow(Workspace.LIT)
                QApplication.processEvents()
                win.grab().save(str(root / 'ui-literature-light.png'))
                win.set_theme('dark')
                win.nav.setCurrentRow(Workspace.LIT)
                QApplication.processEvents()
                win.grab().save(str(root / 'ui-literature-dark.png'))
                win.set_theme('light')
                win.nav.setCurrentRow(Workspace.FILL)
                made = []
                for k, config in enumerate(win.fill.configs):
                    win.fill.list.setCurrentRow(k)
                    QApplication.processEvents()
                    win.grab().save(str(root / f'ui-{k + 1}.png'))
                    values = win.fill.values()
                    for field in config['fields']:
                        if field.get('required') and not str(values.get(field['key'], '')).strip():
                            values[field['key']] = '测试项目' if field['key'] == '项目名称' else '测试负责人'
                    out = root / (config['id'] + '.docx')
                    render_word(Path(config['_dir']) / config['file'], out, config, values)
                    win.fill.store.save(config, values)
                    assert win.fill.store.load(config) == values
                    made.append(out.name)
                marked_ok, example_ok = None, None
                if sample and Path(sample).is_file():
                    win.nav.setCurrentRow(Workspace.STUDIO)
                    win.mark.load(sample)
                    win.mark.quiet = True
                    win.mark.run_recon()
                    if win.mark._name_worker is not None:
                        import time
                        end = time.time() + 120
                        while time.time() < end and win.mark._name_worker is not None:
                            QApplication.processEvents()
                            time.sleep(0.05)
                    if win.mark._vision_worker is not None:
                        import time
                        end = time.time() + 180
                        while time.time() < end and win.mark._vision_worker is not None:
                            QApplication.processEvents()
                            time.sleep(0.05)
                    example_ok = {'fields': len(win.mark.candidates),
                                  'plan': bool(win.mark.example_plan or win.mark.form_plan),
                                  'vision': (win.mark.form_plan or {}).get('vision'),
                                  'req': (win.mark.form_plan or {}).get('req_summary'),
                                  'kind': ('example' if win.mark.example_plan else
                                           'form' if win.mark.form_plan else 'normal')}
                    win.mark.check_all(True)
                    win.register_btn_state = win.mark.register_btn.isEnabled()
                    win.mark.register()
                    new_folders = [p for p in win.store.custom.glob('custom-*') if p.is_dir()]
                    if new_folders:
                        cfg = read_json(new_folders[-1] / 'config.json')
                        has_bookmarks = any(f.get('bookmark') for f in cfg['fields'])
                        if has_bookmarks:
                            keys = {n for n in bookmarkfill.scan_bookmarks(
                                new_folders[-1] / 'template.docx') if n.startswith('WTT')}
                            declared = {f['bookmark'] for f in cfg['fields'] if f.get('bookmark')}
                        else:
                            keys = scan_fields(new_folders[-1] / 'template.docx')
                            declared = {f['key'] for f in cfg['fields']}
                            declared |= {i['key'] for i in cfg.get('images', [])}
                        marked_ok = {'fields': len(cfg['fields']),
                                     'match': declared <= set(keys),
                                     'bookmarks': has_bookmarks,
                                     'example': example_ok}
                layout_ok = None
                if sample and Path(sample).is_file():
                    import time
                    win.nav.setCurrentRow(Workspace.LAYOUT)
                    win.layout_page.set_document(sample)
                    win.layout_page.page_group.paper.setCurrentText('A4')
                    win.layout_page.use_preset('body')
                    win.layout_page.apply_now()
                    end = time.time() + 300
                    worker = win.layout_page._worker
                    while worker is None and time.time() < end:
                        QApplication.processEvents()
                        time.sleep(0.05)
                        worker = win.layout_page._worker
                    if worker is not None:
                        worker.wait(300000)
                    QApplication.processEvents()
                    pending = bool(worker and worker.isRunning())
                    entry = (win.layout_page.history[win.layout_page.pointer]
                             if win.layout_page.history else {})
                    applied = None
                    if entry.get('output') and Path(entry['output']).is_file():
                        from docx import Document
                        applied = layout.read_page(Document(entry['output']))
                    win.layout_page.undo()
                    undone = win.layout_page.pointer
                    win.layout_page.redo()
                    layout_ok = {'output': Path(entry.get('output') or '').name,
                                 'pages': entry.get('pages', 0),
                                 'page': applied,
                                 'pending': pending,
                                 'status': win.layout_page.status_label.text()[:120],
                                 'undo_pointer': undone,
                                 'redo_pointer': win.layout_page.pointer}
                embedded = embedded_ai.find_assets()
                live_preview = None
                if os.environ.get('WTT_PREVIEW_TEST'):
                    import time
                    win.nav.setCurrentRow(Workspace.FILL)
                    win.fill.live_available = True
                    win.fill.live_toggle.blockSignals(True)
                    win.fill.live_toggle.setChecked(True)
                    win.fill.live_toggle.blockSignals(False)
                    win.fill.schedule_live_preview(0)
                    end = time.time() + 60
                    while time.time() < end and not win.fill.live_pdf \
                            and not win.fill.preview_worker._no_office:
                        QApplication.processEvents()
                        time.sleep(0.05)
                    live_preview = {'pages': win.fill.live_page_count,
                                    'no_office': win.fill.preview_worker._no_office,
                                    'status': win.fill.preview_status.text()}
                win.fill.shutdown_preview()
                (root / 'result.json').write_text(json.dumps(
                    {'ok': True, 'templates': len(win.fill.configs), 'generated': made,
                     'mark': marked_ok, 'version': VERSION, 'theme': win.theme,
                     'layout': layout_ok,
                     'pages': list(Workspace.PAGES), 'home_cards': win.home.wall.count(),
                     'literature': {'platforms': len(literature.PLATFORMS),
                                    'e2sci': literature.E2SCI_URL,
                                    'gbt': literature.format_gbt({'title': '测试文献', 'authors': '张三; 李四',
                                                                  'journal': '测试期刊', 'year': '2025',
                                                                  'volume': '1', 'issue': '2', 'pages': '3-4',
                                                                  'type': 'article'})},
                     'embedded': embedded['label'] if embedded else '',
                     'live_preview': live_preview},
                    ensure_ascii=False), encoding='utf-8')
                app.exit(0)
            except Exception:
                import traceback
                (root / 'error.txt').write_text(traceback.format_exc(), encoding='utf-8')
                app.exit(2)
        QTimer.singleShot(400, smoke)
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
