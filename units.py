# -*- coding: utf-8 -*-
"""单位、字号、行距、颜色的统一解析与 Word 合法范围校验。

做法参照 OfficeCLI：属性接受多种写法（cm / pt / 倍 / 百分比 / 中文色名…），
落到 OOXML 之前统一换算并钳制到 Word 自己允许的区间；越界就报中文错误，
而不是写出一份 Word 打不开或提示“文件已损坏”的文件。
"""
from __future__ import annotations

import re

TWIPS_PER_INCH = 1440.0
TWIPS_PER_CM = TWIPS_PER_INCH / 2.54
TWIPS_PER_MM = TWIPS_PER_CM / 10.0
TWIPS_PER_PT = 20.0
TWIPS_PER_PX = TWIPS_PER_INCH / 96.0

# Word 页面尺寸合法区间（twips）：约 0.4cm–55.9cm，EcmaSpec 上限 31680。
PAGE_MIN_TWIPS = 240
PAGE_MAX_TWIPS = 31680
MARGIN_MIN_TWIPS = 0
MARGIN_MAX_TWIPS = 31680
SPACING_MAX_TWIPS = 31680        # 段前/段后 Word 上限 1584 磅
LINE_MAX_TWIPS = 86400
SIZE_MIN_PT = 1.0
SIZE_MAX_PT = 400.0

CN_SIZES = [('初号', 42.0), ('小初', 36.0), ('一号', 26.0), ('小一', 24.0),
            ('二号', 22.0), ('小二', 18.0), ('三号', 16.0), ('小三', 15.0),
            ('四号', 14.0), ('小四', 12.0), ('五号', 10.5), ('小五', 9.0),
            ('六号', 7.5), ('小六', 6.5), ('七号', 5.5), ('八号', 5.0)]
SIZE_NAME_TO_PT = dict(CN_SIZES)
PT_TO_SIZE_NAME = {}
for _name, _pt in CN_SIZES:
    PT_TO_SIZE_NAME.setdefault(_pt, _name)

COLOR_NAMES = {
    '黑': '000000', '黑色': '000000', '白': 'FFFFFF', '白色': 'FFFFFF',
    '红': 'FF0000', '红色': 'FF0000', '纯红': 'FF0000', '深红': 'C00000',
    '蓝': '0000FF', '蓝色': '0000FF', '深蓝': '1F4E79', '浅蓝': 'DDEBF7',
    '绿': '00B050', '绿色': '00B050', '深绿': '006100', '黄': 'FFFF00',
    '黄色': 'FFFF00', '橙': 'FFC000', '橙色': 'FFC000', '紫': '7030A0',
    '粉': 'FF99CC', '灰': '808080', '灰色': '808080', '深灰': '595959',
    '浅灰': 'D9D9D9', '金黄': 'FFD700', '自动': '', '无': '', '默认': '',
    'red': 'FF0000', 'black': '000000', 'white': 'FFFFFF', 'blue': '0000FF',
    'green': '00B050', 'yellow': 'FFFF00', 'gray': '808080', 'grey': '808080',
    'orange': 'FFC000', 'purple': '7030A0', 'cyan': '00B0F0',
}
THEME_COLORS = {'accent1': '4472C4', 'accent2': 'ED7D31', 'accent3': 'A9D18E',
                'accent4': 'FFC000', 'accent5': '5B9BD5', 'accent6': '70AD47',
                'dark1': '000000', 'light1': 'FFFFFF', 'hyperlink': '0563C1'}
HIGHLIGHT_COLORS = {'yellow': 'FFFF00', 'green': '00FF00', 'cyan': '00FFFF',
                    'magenta': 'FF00FF', 'blue': '0000FF', 'red': 'FF0000',
                    'darkyellow': '808000', 'darkblue': '000080',
                    '黑': '000000', '黄': 'FFFF00', '绿': '00FF00', '蓝': '000080',
                    '红': 'FF0000', '无': 'none', '清除': 'none'}

UNIT_TO_TWIPS = {'pt': TWIPS_PER_PT, '磅': TWIPS_PER_PT, 'p': TWIPS_PER_PT,
                 'cm': TWIPS_PER_CM, '厘米': TWIPS_PER_CM, '公分': TWIPS_PER_CM,
                 'mm': TWIPS_PER_MM, '毫米': TWIPS_PER_MM,
                 'in': TWIPS_PER_INCH, 'inch': TWIPS_PER_INCH, '英寸': TWIPS_PER_INCH,
                 '"': TWIPS_PER_INCH, 'px': TWIPS_PER_PX, '像素': TWIPS_PER_PX,
                 'twips': 1.0, '缇': 1.0, '': 1.0}
LENGTH_RE = re.compile(r'^\s*([-+]?\d+(?:\.\d+)?)\s*'
                       r'(pt|磅|p|cm|厘米|公分|mm|毫米|in|inch|英寸|"|px|像素|twips|缇)?\s*$',
                       re.IGNORECASE)


class UnitError(ValueError):
    """单位/取值不合法：消息本身就是给用户看的中文说明。"""


def _num(text):
    try:
        return float(text)
    except (TypeError, ValueError):
        raise UnitError(f'“{text}”不是数字')


def parse_length(value, unit='pt', name='数值'):
    """任意写法 → twips。裸数字按 unit 解释（默认磅），不猜。"""
    if value is None:
        raise UnitError(f'{name} 不能为空')
    if isinstance(value, (int, float)):
        number, suffix = float(value), unit
    else:
        match = LENGTH_RE.match(str(value))
        if not match:
            raise UnitError(f'{name}“{value}”看不懂，可以写 18pt、2cm、0.5英寸、12磅')
        number = float(match.group(1))
        suffix = (match.group(2) or '').lower() or unit
    factor = UNIT_TO_TWIPS.get(suffix, UNIT_TO_TWIPS.get(suffix.rstrip('.')))
    if factor is None:
        raise UnitError(f'{name} 不支持单位“{suffix}”，可用：pt/磅、cm/厘米、mm/毫米、in/英寸、px')
    return int(round(number * factor))


def format_length(twips, unit='pt'):
    factor = UNIT_TO_TWIPS.get(unit, TWIPS_PER_PT)
    value = twips / factor
    return f'{value:g}{unit}'


def clamp(twips, name, low, high, display='pt'):
    if twips < low or twips > high:
        raise UnitError(f'{name} 超出 Word 允许范围'
                        f'（{format_length(low, display)}–{format_length(high, display)}），'
                        f'当前 {format_length(twips, display)}')
    return int(twips)


def parse_page_dim(value, name='纸张尺寸'):
    return clamp(parse_length(value, unit='mm', name=name), name,
                 PAGE_MIN_TWIPS, PAGE_MAX_TWIPS, display='cm')


def parse_margin(value, name='页边距'):
    return clamp(parse_length(value, unit='cm', name=name), name,
                 MARGIN_MIN_TWIPS, MARGIN_MAX_TWIPS, display='cm')


def parse_spacing_before(value, name='段前间距'):
    """段前/段后：磅、厘米或 twips，负数一律拒绝。"""
    if isinstance(value, (int, float)):
        value = f'{value}pt'
    return clamp(parse_length(value, unit='pt', name=name), name,
                 0, SPACING_MAX_TWIPS, display='pt')


LINE_WORD = {'单倍': ('multiple', 1.0), '单倍行距': ('multiple', 1.0), '1倍': ('multiple', 1.0),
             '1.5倍': ('multiple', 1.5), '2倍': ('multiple', 2.0),
             '1.5倍行距': ('multiple', 1.5), '2倍行距': ('multiple', 2.0)}
SPACING_RE = re.compile(r'^\s*(固定值?|至少|最小)?\s*'
                        r'(\d+(?:\.\d+)?)\s*'
                        r'(pt|磅|倍行距|倍|%|x|cm|毫米|mm|英寸|in)?\s*$', re.IGNORECASE)


def parse_spacing(value, name='行距'):
    """→ {'type': exact|at_least|multiple, 'value': float}，与 formatapply.set_spacing 一致。"""
    if isinstance(value, dict) and value.get('type'):
        return value
    if isinstance(value, (int, float)):
        number = float(value)
        return {'type': 'multiple' if number < 4 else 'exact', 'value': number}
    text = str(value or '').strip()
    if not text:
        raise UnitError(f'{name} 不能为空')
    if text in LINE_WORD:
        kind, number = LINE_WORD[text]
        return {'type': kind, 'value': number}
    for word, (kind, number) in LINE_WORD.items():
        if text == f'{word}行距':
            return {'type': kind, 'value': number}
    match = SPACING_RE.match(text)
    if not match:
        raise UnitError(f'{name}“{value}”看不懂，可以写 1.5倍、150%、20磅、固定值20磅、至少12磅')
    kind_word, number_text, unit = match.group(1), float(match.group(2)), (match.group(3) or '').lower()
    if unit in ('倍', '倍行距', 'x'):
        kind, number = 'multiple', number_text
    elif unit == '%':
        kind, number = 'multiple', number_text / 100.0
    elif unit in ('', 'pt', '磅'):
        kind = 'at_least' if kind_word and kind_word.startswith(('至少', '最小')) else 'exact'
        if not kind_word:
            kind = 'multiple' if number_text <= 4 and kind_word is None else 'exact'
        number = number_text
    else:
        twips = parse_length(f'{number_text}{unit}', name=name)
        kind = 'at_least' if kind_word and kind_word.startswith(('至少', '最小')) else 'exact'
        number = twips / TWIPS_PER_PT
    if kind == 'multiple' and not (0.1 <= number <= 20):
        raise UnitError(f'{name} 倍数要在 0.1–20 之间，当前 {number:g}')
    if kind in ('exact', 'at_least') and not (1 <= number <= 1584):
        raise UnitError(f'{name} 磅值要在 1–1584 之间，当前 {number:g} 磅')
    return {'type': kind, 'value': round(number, 3)}


CN_DIGITS = {'一': 1.0, '二': 2.0, '两': 2.0, '三': 3.0, '四': 4.0, '五': 5.0,
             '六': 6.0, '七': 7.0, '八': 8.0, '九': 9.0, '十': 10.0, '十二': 12.0,
             '十五': 15.0, '十六': 16.0, '十八': 18.0, '二十': 20.0}


def parse_font_size(value, name='字号'):
    """→ 磅值。接受“小四”“12pt”“12”“十二磅”。"""
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        text = str(value or '').strip()
        if not text:
            raise UnitError(f'{name} 不能为空')
        lowered = text.lower().rstrip('pt').strip() if text.lower().endswith('pt') else text
        if lowered in SIZE_NAME_TO_PT:
            return SIZE_NAME_TO_PT[lowered]
        for cn, pt in SIZE_NAME_TO_PT.items():
            if cn in text and not re.search(r'\d', text):
                return pt
        match = re.match(r'^\s*(\d+(?:\.\d+)?)\s*(pt|磅)?\s*$', text, re.IGNORECASE)
        if match:
            number = float(match.group(1))
        elif text.rstrip('磅') in CN_DIGITS:
            number = CN_DIGITS[text.rstrip('磅')]
        else:
            raise UnitError(f'{name}“{value}”看不懂，可以写 小四、12pt、10.5、五号')
    if not (SIZE_MIN_PT <= number <= SIZE_MAX_PT):
        raise UnitError(f'{name} 要在 {SIZE_MIN_PT:g}–{SIZE_MAX_PT:g} 磅之间，当前 {number:g}')
    return number


def size_label(pt):
    """磅值 → “12pt（小四）”，没有对应中文字号就只写磅值。"""
    name = PT_TO_SIZE_NAME.get(float(pt))
    return f'{pt:g}pt' + (f'（{name}）' if name else '')


COLOR_HEX_RE = re.compile(r'^#?([0-9A-Fa-f]{6})$')
RGB_RE = re.compile(r'^rgb\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$', re.IGNORECASE)


def parse_color(value, name='颜色'):
    """→ 6 位十六进制（不带 #）；“自动/无/默认”返回空串表示清除颜色。"""
    text = str(value or '').strip()
    if not text:
        raise UnitError(f'{name} 不能为空')
    match = COLOR_HEX_RE.match(text)
    if match:
        return match.group(1).upper()
    match = RGB_RE.match(text)
    if match:
        parts = [int(g) for g in match.groups()]
        if any(part < 0 or part > 255 for part in parts):
            raise UnitError(f'{name} 的 rgb() 分量要在 0–255 之间')
        return ''.join(f'{part:02X}' for part in parts)
    key = text.lower()
    if key in THEME_COLORS:
        return THEME_COLORS[key]
    if text in COLOR_NAMES:
        return COLOR_NAMES[text]
    if key in COLOR_NAMES:
        return COLOR_NAMES[key]
    raise UnitError(f'{name}“{value}”看不懂，可以写 FF0000、#C00000、深灰、蓝色、rgb(0,112,192)、accent1、自动')


def parse_highlight(value):
    text = str(value or '').strip()
    key = text.lower()
    if text in HIGHLIGHT_COLORS:
        return HIGHLIGHT_COLORS[text]
    if key in HIGHLIGHT_COLORS:
        return HIGHLIGHT_COLORS[key]
    hexed = COLOR_HEX_RE.match(text)
    if hexed:
        return hexed.group(1).upper()
    raise UnitError(f'高亮色“{value}”看不懂，可用：黄/绿/蓝/红/青/品红 或 FF0000、清除')


def describe_value(value):
    """把用户输入原样规范化成一句人话，用于界面回显和报告。"""
    text = str(value or '').strip()
    if not text:
        return ''
    try:
        return size_label(parse_font_size(text))
    except UnitError:
        pass
    if re.search(r'(cm|mm|in|inch|px|厘米|毫米|英寸|像素|")$', text, re.IGNORECASE):
        try:
            return f'{parse_length(text) / TWIPS_PER_CM:.2f}cm'
        except UnitError:
            pass
    try:
        spacing = parse_spacing(text)
        if spacing['type'] == 'multiple':
            return f"{spacing['value']:g} 倍行距"
        return f"{'固定值' if spacing['type'] == 'exact' else '最小值'} {spacing['value']:g} 磅"
    except UnitError:
        pass
    try:
        return parse_color(text) or '自动'
    except UnitError:
        pass
    return text
