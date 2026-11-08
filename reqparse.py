"""Read the requirements a template states about itself and turn them into rules.

School and office forms usually carry their own formatting instructions written in
coloured text next to the field they govern::

    中文字全部采用“宋体、小四号字、加粗”。英文、数字采用“Times new Roman小四号”，加粗。
    行间距：固定值20磅 …  悬挂缩进2字符 …  参考文献不少于10篇 …  （不少于100字）

This module finds those sentences and parses them **deterministically** (a model can
never invent a font name here), then anchors each rule to the field or paragraph it
governs so the fill step can apply the format and move the sentence out of the
document and into the field's hint.  Requirements are read from the document's own
OOXML text — no page screenshots are used.
"""
from __future__ import annotations

import copy
import re

from docx import Document
from docx.oxml.ns import qn

import bookmarkfill
import exampleparse as ep
import visionparse

SIZE_MAP = {'小初': 36, '初号': 42, '小一': 24, '一号': 26, '小二': 18, '二号': 22,
            '小三': 15, '三号': 16, '小四': 12, '四号': 14, '小五': 9, '五号': 10.5,
            '小六': 6.5, '六号': 7.5}

CN_FONTS = ('宋体', '黑体', '楷体', '仿宋', '华文中宋', '微软雅黑', '等线', '幼圆',
            '隶书', '方正小标宋简体', '华文楷体', '华文仿宋', '新宋体')
LATIN_FONT_RE = re.compile(r'(times\s*new\s*roman|arial|calibri|verdana|courier\s*new|'
                           r'times|garamond|georgia)', re.I)

# A sentence only becomes a deletable requirement when it both reads like guidance
# and yields something concrete to apply or check.
# 要求解析版本：旧模板的 req 没有这个版本号，加载时会离线重读一次（补上黑字要求）
REQ_VERSION = 2

REQ_LEAD_RE = re.compile(
    r'(?:中文字|中文文字|汉字|英文|数字|字母|阿拉伯数字|行距|行间距|对齐方式|对齐|首行缩进|'
    r'悬挂缩进|左缩进|右缩进|此处|本栏|注意|格式参考|参考格式|不少于|至少|不得|不要|不能|'
    r'应当|应该|要求|手签|签字|整页设计|每页都要|空\s*[一二两三四五\d]+\s*[行格]|双面打印|'
    r'装订|涂改|居中|加粗|不加粗|倾斜|小四|五号|六号|宋体|黑体|楷体|仿宋|Times|代码加汉字|'
    r'结束语|二级标题|三级标题|编号|撑满|跨页|跳页|仅供参考|自行填写|直接使用|字数|评语|'
    r'文献|篇|字）|字\)|按实|如实|只填|其中一种|建议|以内)', re.I)
STRONG_GUIDE_RE = re.compile(
    r'(?:此处|请注意|仅供参考|自行填写|直接使用|不得|不要|不能|应当|按照|要求|格式参考|'
    r'保持一致|一致$|整页设计|双面打印|左侧装订|涂改|手签|按实|如实|其中一种|建议|以内|只填)')
CN_NUM = {'一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7,
          '八': 8, '九': 9, '十': 10}

LINE_SPACING_RE = re.compile(
    r'行(?:距|间距)[^。，,；;。）)]{0,6}?'
    r'(?:固定值[，,]?\s*(\d+(?:\.\d+)?)\s*磅|(\d+(?:\.\d+)?)\s*磅|单倍|1\.5\s*倍|'
    r'1\s*倍|2\s*倍|多倍|最小值[，,]?\s*(\d+(?:\.\d+)?)\s*磅)')
FIRST_INDENT_RE = re.compile(r'首行缩进\s*(\d+(?:\.\d+)?|[一二两三四五])\s*(?:字符|字|格)')
HANGING_RE = re.compile(r'悬挂缩进\s*(\d+(?:\.\d+)?|[一二两三四五])\s*(?:字符|字|格)')
BLANK_LINES_RE = re.compile(r'空\s*([一二两三四五]|\d+)\s*[行格]')
PT_SIZE_RE = re.compile(r'(\d+(?:\.\d+)?)\s*(?:磅|pt)\s*字')
RANGE_CHARS_RE = re.compile(r'(\d+)\s*[—\-~至到]\s*(\d+)\s*字')
MIN_CHARS_RE = re.compile(r'不少于\s*(\d+)\s*字')
MIN_REFS_RE = re.compile(r'(?:不少于|至少)\s*(\d+)\s*篇')
MIN_RECORDS_RE = re.compile(r'(?:至少要有|不少于)\s*(\d+)\s*次')
SCORE_MIN_RE = re.compile(r'(?:不能低于|不低于|最低)\s*(\d+)\s*分')
ENDING_RE = re.compile(r'结束语为\s*[“"\'「]([^”"\'」]+)[”"\'」]')


def _num(value):
    value = str(value)
    return float(value) if value.replace('.', '', 1).isdigit() else CN_NUM.get(value, 0)


def _find_size(text):
    compact = re.sub(r'\s+', '', text)
    best = None
    for name, pt in SIZE_MAP.items():
        index = compact.find(name)
        if index < 0:
            continue
        if best is None or index < best[0]:
            best = (index, name, pt)
    if best:
        return best[2]
    match = PT_SIZE_RE.search(text)
    if match:
        size = float(match.group(1))
        if 6 <= size <= 72:
            return size
    return None


def _find_fonts(text):
    """Return (east_asian_font, ascii_font) implied by the sentence."""
    east = ascii_font = ''
    latin = LATIN_FONT_RE.search(text)
    if latin:
        ascii_font = re.sub(r'\s+', ' ', latin.group(1)).title()
        if ascii_font.lower() == 'times new roman':
            ascii_font = 'Times New Roman'
    for font in sorted(CN_FONTS, key=len, reverse=True):
        if font in text:
            east = font
            break
    if not east and not ascii_font:
        return '', ''
    # “英文、数字采用X” only ever names the Latin font; “中文字采用X” names CJK.
    latin_only = re.search(r'(?:英文|数字|字母)[^。；;]{0,24}?(Times New Roman|Arial|Calibri|'
                           r'Verdana|Courier New|Times)', text, re.I)
    if latin_only and not re.search(r'(?:中文字|中文文字|汉字)[^。；;]{0,18}?(宋体|黑体|楷体|仿宋)',
                                    text):
        return '', ascii_font
    return east, ascii_font


def _find_line_spacing(text):
    match = LINE_SPACING_RE.search(text)
    if not match:
        return None
    exact, plain, minimum = match.group(1), match.group(2), match.group(3)
    if exact:
        return {'type': 'exact', 'value': float(exact)}
    if minimum:
        return {'type': 'at_least', 'value': float(minimum)}
    if plain:
        return {'type': 'exact', 'value': float(plain)}
    token = match.group(0)
    if '单倍' in token or re.search(r'1\s*倍', token):
        return {'type': 'multiple', 'value': 1.0}
    if '1.5' in token:
        return {'type': 'multiple', 'value': 1.5}
    if re.search(r'2\s*倍', token):
        return {'type': 'multiple', 'value': 2.0}
    return {'type': 'multiple', 'value': 1.0}


def parse_requirement(text):
    """Map one Chinese requirement sentence onto concrete, applied-locally rules."""
    text = text or ''
    rules = {}
    spacing = _find_line_spacing(text)
    if spacing:
        rules['line'] = spacing
    size = _find_size(re.sub(r'行(?:距|间距)[^。，,；;]*', '', text))
    if size:
        rules['size_pt'] = size
    east, latin = _find_fonts(text)
    if east:
        rules['font_east'] = east
    if latin:
        rules['font_ascii'] = latin
    if '不加粗' in text:
        rules['bold'] = False
    elif re.search(r'加粗|粗体', text):
        rules['bold'] = True
    if re.search(r'倾斜|斜体', text):
        rules['italic'] = True
    if re.search(r'上下左右居中', text):
        rules['align'] = 'center'
        rules['valign'] = 'center'
    elif re.search(r'居中', text):
        rules['align'] = 'center'
    elif re.search(r'两端对齐', text):
        rules['align'] = 'justify'
    elif re.search(r'右对齐', text):
        rules['align'] = 'right'
    elif re.search(r'左对齐', text):
        rules['align'] = 'left'
    match = FIRST_INDENT_RE.search(text)
    if match:
        rules['first_line_chars'] = _num(match.group(1))
    match = HANGING_RE.search(text)
    if match:
        rules['hanging_chars'] = _num(match.group(1))
    match = BLANK_LINES_RE.search(text)
    if match:
        rules['blank_lines'] = _num(match.group(1))
    content = {}
    match = RANGE_CHARS_RE.search(text)
    if match:
        content['min_chars'] = int(match.group(1))
        content['max_chars'] = int(match.group(2))
    else:
        match = MIN_CHARS_RE.search(text)
        if match:
            content['min_chars'] = int(match.group(1))
    match = MIN_REFS_RE.search(text)
    if match:
        content['min_refs'] = int(match.group(1))
    match = MIN_RECORDS_RE.search(text)
    if match:
        content['min_records'] = int(match.group(1))
    match = SCORE_MIN_RE.search(text)
    if match:
        content['score_min'] = int(match.group(1))
    match = re.search(r'(\d+)\s*年以内', text)
    if match:
        content['max_years'] = int(match.group(1))
    match = ENDING_RE.search(text)
    if match:
        content['ending'] = match.group(1)
    if re.search(r'手签|手写', text):
        content['must_handwrite'] = True
    if re.search(r'不得跨页|不要跳页|不能跨页', text):
        content['keep_together'] = True
    if re.search(r'实词首字母大写', text):
        content['title_case'] = True
    if re.search(r'中文数字编号', text):
        content['cn_numbering'] = True
    if re.search(r'二级标题', text):
        content['heading_level'] = 2
    if re.search(r'双面打印|左侧装订', text):
        content['print_note'] = True
    if content:
        rules['content'] = content
    return rules


def has_concrete_rule(rules):
    return any(key in rules for key in
               ('size_pt', 'font_east', 'font_ascii', 'bold', 'italic', 'align',
                'valign', 'line', 'first_line_chars', 'hanging_chars', 'blank_lines'))


def is_requirement_text(text):
    text = (text or '').strip()
    if len(text) < 5 or not REQ_LEAD_RE.search(text):
        return False
    return bool(has_concrete_rule(parse_requirement(text)) or STRONG_GUIDE_RE.search(text))


def _run_mark(run):
    """Classify a run: guidance is highlighted/red, sample answers are blue/green."""
    rpr = run._element.find(qn('w:rPr'))
    highlight = shading = color = ''
    if rpr is not None:
        node = rpr.find(qn('w:highlight'))
        if node is not None:
            highlight = (node.get(qn('w:val')) or '').strip().lower()
        node = rpr.find(qn('w:shd'))
        if node is not None:
            shading = (node.get(qn('w:fill')) or '').strip().lower()
        color = ep._run_color(run)
    if highlight and highlight != 'none':
        return 'note'
    if shading and shading not in ('auto', 'ffffff', 'none'):
        return 'note'
    if color:
        return 'red' if ep._is_red(color) else (
            'sample' if ep._is_example_color(color) else 'note')
    return 'plain'


def _run_spans(paragraph):
    """[(start, end, mark)] for every run that carries text."""
    spans = []
    cursor = 0
    for run in paragraph.runs:
        text = run.text or ''
        if text:
            spans.append((cursor, cursor + len(text), _run_mark(run)))
            cursor += len(text)
    return spans


def _mark_ratio(spans, start, end, marks):
    total = matched = 0
    for s, e, mark in spans:
        overlap = min(e, end) - max(s, start)
        if overlap <= 0:
            continue
        total += overlap
        if mark in marks:
            matched += overlap
    return (matched / total) if total else 0.0


def _mark_at(spans, index):
    for s, e, mark in spans:
        if s <= index < e:
            return mark
    return 'plain'


def _span_start_ok(text, spans, start):
    """A partial instruction must begin at a clause edge or where the marking changes.

    Without this, ``主要研究内容与要求（请注意此处有变化）`` would be read as an
    instruction starting at ``要求`` and deleting it would mangle the heading.
    """
    if start == 0:
        return True
    before = text[start - 1]
    if before in '。！？；;：:，,、 \u3000\t（）()[""''「」':
        return True
    return _mark_at(spans, start) != _mark_at(spans, start - 1)


def _region_start(spans, index):
    """First character of the contiguous marked block that holds ``index``.

    A keyword sitting in ordinary black text (``主要研究内容与要求``) must not drag the
    whole run along, or deleting the “instruction” would cut the heading in half.
    """
    for start, end, mark in spans:
        if start <= index < end:
            if mark not in ('note', 'red'):
                return index
            while start > 0 and _mark_at(spans, start - 1) in ('note', 'red'):
                start -= 1
            return start
    return index


# Instructions written in plain black text are only treated as guidance when the whole
# short paragraph is clearly an imperative, so a sample answer can never be eaten.
STRONG_IMPERATIVE_RE = re.compile(
    r'(?:双面打印|左侧装订|不要减少栏目|不得减少栏目|不能改变内容|保持“?整页设计|'
    r'不得涂改|不能涂改|必须本人手签|此页只填写|下不赘述)')


def _guidance_end(spans, start, limit):
    """标出的说明到哪里结束：连续（或仅隔标点）的标记段末尾，遇到长段未标记正文即停。"""
    end = start
    for begin, stop, mark in spans:
        if stop <= start:
            continue
        if begin > end and end > start and begin - end > 3:
            break
        if mark in ('note', 'red'):
            end = max(end, stop)
        elif end > start and begin >= end:
            break
    return end if end > start else limit


def note_spans(paragraph):
    """Find guidance sentences inside one paragraph, whole or partial.

    A paragraph such as ``上海养老服务…研究中文字全部采用“宋体、小四号字、加粗”。`` mixes the
    blue sample answer with the highlighted instruction; only the instruction tail is
    returned so the sample survives.
    """
    text = ''.join(run.text or '' for run in paragraph.runs)
    if not text.strip() or not REQ_LEAD_RE.search(text):
        return []
    spans = _run_spans(paragraph)
    stripped = text.strip()
    imperative = STRONG_IMPERATIVE_RE.search(stripped)
    if imperative and len(stripped) <= 160 and not stripped.endswith(('：', ':')):
        return [(0, len(text), text, True,
                 _mark_ratio(spans, 0, len(text), ('note', 'red')) or 1.0)]
    results = []
    for match in REQ_LEAD_RE.finditer(text):
        start = _region_start(spans, match.start())
        # Take the opening bracket with the guidance, e.g. “（此处空一行”.
        while start > 0 and text[start - 1] in '（([':
            start -= 1
        if any(begin <= start < end for begin, end, _t, _w, _c in results):
            continue
        candidate = text[start:]
        if not is_requirement_text(candidate):
            continue
        marked_end = _guidance_end(spans, start, len(text))
        note_ratio = _mark_ratio(spans, start, marked_end, ('note', 'red'))
        whole = start == 0 or not text[:start].strip()
        if not whole and not _span_start_ok(text, spans, start):
            continue
        if note_ratio < 0.5 and not (whole and STRONG_GUIDE_RE.search(candidate)):
            continue
        end = marked_end
        tail = text.rstrip()
        # “（此处填写应与…一致。如：…）”这类括号整体是说明：标记只盖前半句，删除要连括号一起
        if text[start] in '（(' and tail.endswith(('）', ')')) and end < len(tail):
            end = len(tail)
        results.append((start, end, text[start:end], whole, note_ratio))
    return results


def _norm_range(rng):
    value = dict(rng or {})
    value.setdefault('container', 'cell')
    value.setdefault('part', 'body')
    return value


def _field_index(plan):
    index = []
    for field in (plan or {}).get('fields', []):
        rng = _norm_range(field.get('bookmark_range')) if field.get('bookmark_range') else None
        loc = field.get('location') or {}
        index.append({
            'key': field.get('key'), 'label': field.get('label') or field.get('key'),
            'range': rng, 'location': loc,
            'cell': (rng or {}).get('table'), 'has_cell': (rng or {}).get('container') == 'cell',
        })
    return index


def _same_cell(a, b):
    return (a.get('table'), a.get('row'), a.get('col')) == (b.get('table'), b.get('row'), b.get('col'))


def _match_field(fields, note):
    """Which fill slot does this guidance belong to?"""
    cell = (note.get('table'), note.get('row'), note.get('col'))
    if note['container'] == 'cell':
        exact = [f for f in fields if f['range'] and f['range'].get('container') == 'cell'
                 and _same_cell(f['range'], {'table': cell[0], 'row': cell[1], 'col': cell[2]})]
        if exact:
            return exact[0], 'same_cell'
        # The note sits in a label cell: the value cell of the same row owns it.
        row = [f for f in fields if f['range'] and f['range'].get('container') == 'cell'
               and (f['range'].get('table'), f['range'].get('row')) == (cell[0], cell[1])]
        if len(row) == 1:
            return row[0], 'same_row'
        if row:
            return None, 'row_ambiguous'
    else:
        p = note['p']
        inside = [f for f in fields if f['range'] and f['range'].get('container') == note['container']
                  and f['range'].get('start_p', 0) <= p <= f['range'].get('end_p', 0)]
        if len(inside) == 1:
            return inside[0], 'inside_range'
        above = [f for f in fields if f['range'] and f['range'].get('container') == note['container']
                 and f['range'].get('end_p', -1) < p]
        if above:
            best = max(above, key=lambda f: f['range'].get('end_p', 0))
            if p - best['range'].get('end_p', 0) <= 3:
                return best, 'section_above'
    label = _nearby_label(note)
    if label:
        for field in fields:
            if visionparse._norm_label(field.get('label')) == visionparse._norm_label(label):
                return field, 'label_match'
    return None, 'unmatched'


def _nearby_label(note):
    text = (note.get('text') or '').strip()
    if not text:
        return ''
    match = re.match(r'^\s*([\u4e00-\u9fff（）()A-Za-z/、·]{2,20}?)[:：]', text)
    return visionparse._clean_label(match.group(1)) if match else ''


def _box_text(box):
    lines = []
    for p in box.iter(qn('w:p')):
        lines.append(''.join(t.text or '' for t in p.iter(qn('w:t'))))
    return '\n'.join(lines).strip()


def _anchor_index(doc):
    """Map every paragraph element to its (container, part, table, row, col, p) spot."""
    index, alive = {}, []
    def remember(paragraph, kind, part, ti, ri, ci, position):
        element = paragraph._p
        alive.append(element)
        index[id(element)] = (kind, part, ti, ri, ci, position)
    for i, paragraph in enumerate(doc.paragraphs):
        remember(paragraph, 'body', 'body', None, None, None, i)
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            seen = set()
            for ci, cell in enumerate(row.cells):
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                for pi, paragraph in enumerate(cell.paragraphs):
                    remember(paragraph, 'cell', 'body', ti, ri, ci, pi)
    for kind, name, root in bookmarkfill.story_parts(doc):
        if kind == 'body':
            continue
        for pi, paragraph in enumerate(bookmarkfill._paragraphs_of_root(root)):
            remember(paragraph, kind, name, None, None, None, pi)
    return index, alive


SHAPE_TAGS = {qn('w:pict'), qn('w:drawing'),
              '{urn:schemas-microsoft-com:vml}shape',
              '{urn:schemas-microsoft-com:vml}rect',
              '{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}wsp'}


def _removable_box(box):
    """The innermost shape that holds one text box (Choice and Fallback each have one)."""
    node = box.getparent()
    picked = None
    while node is not None and node.tag != qn('w:p'):
        if node.tag in SHAPE_TAGS and picked is None:
            picked = node
        node = node.getparent()
    return picked


TEXTBOX_GUIDE_RE = re.compile(
    r'(?:此处|本栏|本框|本处|填写|请|按|如实|照抄|直接抄|参考|不得|不要|不能|必须|应|'
    r'选择|勾选|下拉|单选|多选|二选一|留空|不填|无需|勿填|手写|打印|'
    r'字数|不少于|不超过|以内|格式|要求|注意|'
    r'空\s*[一二两三四五六七八九十\d]+\s*[格行]|宋体|黑体|楷体|仿宋|微软雅黑|'
    r'初号|小初|一号|小一|二号|小二|三号|小三|四号|小四|五号|小五|加粗|不加粗|'
    r'行距|行间距|缩进|居中|对齐|字号|Times)')


def _textbox_looks_like_guidance(text):
    """文本框天然是批注：短提示（如“空2格”“填写专业班级…”）也算说明。

    上限放到 300 字：任务书里的“论文类型可选项 + 填表写法”这类整段说明虽然长，
    也确实是写给填表人看的，不能因为超长就漏掉。
    """
    text = (text or '').strip()
    if len(text) < 2 or len(text) > 300:
        return False
    if is_requirement_text(text):
        return True
    return bool(TEXTBOX_GUIDE_RE.search(text))


def _textbox_notes(doc, fields):
    """Guidance written inside floating text boxes (the arrowed callouts)."""
    index, alive = _anchor_index(doc)
    notes = []
    seen = set()
    for _kind, part, root in bookmarkfill.story_parts(doc):
        for box in root.iter(qn('w:txbxContent')):
            host = _removable_box(box)
            if host is None:
                continue
            alive.append(host)
            text = _box_text(box)
            if not text or len(text) < 2:
                continue
            anchor = box.getparent()
            while anchor is not None and anchor.tag != qn('w:p'):
                anchor = anchor.getparent()
            location = index.get(id(anchor)) if anchor is not None else None
            if not _textbox_looks_like_guidance(text):
                continue
            # Two identical callouts are one requirement; every copy gets deleted later.
            if text in seen:
                continue
            seen.add(text)
            note = {
                'source': '文本框', 'container': (location or ('body', 'body'))[0],
                'part': (location or ('body', 'body'))[1],
                'table': (location or (None,) * 6)[2], 'row': (location or (None,) * 6)[3],
                'col': (location or (None,) * 6)[4], 'p': (location or (None,) * 6)[5],
                'start': 0, 'end': 0, 'text': text, 'paragraph_text': text,
                'whole': True, 'marked': 1.0, 'rules': parse_requirement(text),
                'textbox': text,
            }
            field, how = _match_field(fields, note)
            note['field_key'] = field['key'] if field else ''
            note['field_label'] = field['label'] if field else ''
            note['anchor'] = ('textbox_' + how) if field else 'textbox_unmatched'
            note['applies_to'] = 'field' if field else 'document'
            notes.append(note)
    return notes


def _plain_requirement_span(paragraph):
    """黑色正文里写明的格式要求：读规则但不删原文。

    很多模板把“中文用宋体、小四号字”直接写在黑字里，没有高亮也不标红。
    只在句子短、能解析出硬格式规则、且确实对得上某个栏目时才认；
    这类说明只套格式，不删除文字（删了会伤到模板正文）。
    """
    text = ''.join(run.text or '' for run in paragraph.runs).strip()
    if not (5 <= len(text) <= 120) or not is_requirement_text(text):
        return None
    rules = parse_requirement(text)
    if not any(name in rules for name in ('size_pt', 'font_east', 'font_ascii', 'bold',
                                          'line', 'first_line_chars', 'hanging_chars')):
        return None
    bracketed = re.fullmatch(r'[（(【\[][^）)】\]]{2,116}[）)】\]]', text)
    starts_like_lead = bool(REQ_LEAD_RE.match(text)) or bool(
        re.match(r'^(?:中文|英文|数字|字母|汉字|西文|正文|标题)', text))
    if not bracketed and not starts_like_lead:
        return None
    return 0, len(text), text, True, 0.0


def extract_local(path, plan):
    """Coloured requirement sentences already present in the OOXML."""
    doc = Document(str(path))
    fields = _field_index(plan)
    notes = []
    for kind, part, paras, ti, ri, ci, cell in visionparse._containers(doc):
        for p, paragraph in enumerate(paras):
            spans = note_spans(paragraph)
            if not spans:
                plain = _plain_requirement_span(paragraph)
                spans = [plain] if plain else []
            for start, end, text, whole, marked in spans:
                rules = parse_requirement(text)
                note = {
                    'source': '结构', 'container': kind, 'part': part,
                    'table': ti, 'row': ri, 'col': ci, 'p': p,
                    'start': start, 'end': end, 'text': text,
                    'paragraph_text': ''.join(r.text or '' for r in paragraph.runs),
                    'whole': whole, 'marked': round(marked, 2),
                    'rules': rules,
                }
                field, how = _match_field(fields, note)
                if marked <= 0 and field is None:
                    continue        # 黑字句子没对上栏目：不当要求，避免误套/误删
                note['field_key'] = field['key'] if field else ''
                note['field_label'] = field['label'] if field else ''
                note['anchor'] = how
                note['plain'] = marked <= 0
                note['applies_to'] = ('field' if field and how in ('same_cell', 'inside_range',
                                                                   'label_match', 'section_above')
                                      else 'cell' if kind == 'cell' else 'paragraph')
                notes.append(note)
    try:
        notes += _textbox_notes(doc, fields)
    except Exception:
        pass
    return notes


PARA_KEYS = ('line', 'first_line_chars', 'hanging_chars', 'align', 'valign', 'blank_lines')
RUN_KEYS = ('size_pt', 'font_east', 'font_ascii', 'bold', 'italic', 'align')


def _rule_text(name, value):
    if name == 'size_pt':
        return f'{value:g} 磅'
    if name == 'line':
        if isinstance(value, dict):
            return (f'固定值 {value.get("value"):g} 磅' if value.get('type') == 'exact'
                    else f'{value.get("value"):g} 倍行距')
        return str(value)
    return {'font_east': '中文字体', 'font_ascii': '西文字体', 'bold': '加粗',
            'italic': '倾斜', 'align': '对齐', 'valign': '垂直对齐',
            'first_line_chars': '首行缩进', 'hanging_chars': '悬挂缩进',
            'blank_lines': '空行'}.get(name, name) + (f'={value}' if not
                                                    str(value).isalpha() else '')


def _merge_rules(notes, questions):
    """Collapse every note that governs one target into a single rule set."""
    merged = {}
    for index, note in enumerate(notes):
        key = note.get('field_key') or ''
        target = ('field', key) if key else (note['container'],
                                             note.get('table'), note.get('row'),
                                             note.get('col'), note.get('p'))
        bucket = merged.setdefault(target, {'rules': {}, 'notes': [], 'note': note,
                                            'index': index})
        bucket['notes'].append(note['text'])
        for name, value in (note['rules'] or {}).items():
            if name == 'content':
                continue
            if name in bucket['rules'] and bucket['rules'][name] != value:
                questions.append({
                    'id': f'conflict-{key or note["container"]}-{index}-{name}',
                    'kind': 'conflict',
                    'text': f'同一个位置给了两条不同要求：{name} 用 {_rule_text(name, bucket["rules"][name])} '
                            f'还是 {_rule_text(name, value)}？\n· {note["text"][:60]}',
                    'where': _where(note),
                    'options': [f'用先出现的 {_rule_text(name, bucket["rules"][name])}',
                                f'用后出现的 {_rule_text(name, value)}', '这条要求跳过不套'],
                    'default': 0, 'note_index': index,
                    'override': {'target': key or ('container', index), 'rule': name,
                                 'values': [bucket['rules'][name], value]},
                })
                continue
            bucket['rules'].setdefault(name, value)
    return merged


def gather(path, plan, progress=None):
    """模板要求只从文档自己的文字里读：说明、红字、括号要求都在 OOXML 里。"""
    return extract_local(path, plan)


def assemble(notes, plan, answers=None):
    """Turn gathered notes into rules, hints and questions; ``answers`` decides open ones."""
    answers = answers or {}
    questions = []
    kept, scopes = [], {}
    for index, note in enumerate(notes):
        drop = False
        for question in questions:
            if question.get('note_index') != index:
                continue
            choice = answers.get(question['id'])
            if choice is None:
                continue
            if question['kind'] == 'scope':
                if choice == question.get('skip'):
                    drop = True
                elif choice == 0:
                    scopes[index] = 'following'
            elif choice == question.get('skip'):
                drop = True
        if not drop:
            if scopes.get(index):
                note['scope'] = scopes[index]
            kept.append(note)
    notes = kept
    hints, general_hints, seen = {}, [], set()
    for note in notes:
        key = (note.get('field_key'), note['text'])
        if key in seen:
            continue
        seen.add(key)
        text = note['text'].strip()
        if note['field_key']:
            hints.setdefault(note['field_key'], []).append(text)
        else:
            general_hints.append({'text': text, 'where': _where(note)})
    merged = _merge_rules(notes, questions)
    for question in questions:
        if question['kind'] != 'scope':
            continue
        # body-scoped notes need an explicit range decision before they can be applied
        index = question['note_index']
        if index < len(notes) and answers.get(question['id']) == 0:
            notes[index]['scope'] = 'following'
    field_rules, container_rules, cells = {}, [], {}
    for position, (target, bucket) in enumerate(merged.items()):
        rules = bucket['rules']
        if not rules:
            continue
        kind = target[0]
        if kind == 'field':
            field_rules[target[1]] = rules
        elif kind == 'cell':
            _k, table, row, col, _p = target
            cell = cells.setdefault((table, row, col),
                                    {'container': 'cell', 'part': 'body', 'table': table,
                                     'row': row, 'col': col, 'rules': {}, 'exclude_p': []})
            for name, value in rules.items():
                cell['rules'].setdefault(name, value)
            cell['exclude_p'].append(target[4])
        elif kind in ('body', 'header', 'footer'):
            _k, _table, _row, _col, p = target
            note = bucket['note']
            entry = {'container': kind, 'part': note.get('part') or 'body', 'p': p,
                     'rules': rules, 'scope': note.get('scope') or 'self'}
            if has_concrete_rule(rules):
                questions.append({
                    'id': f'scope-{bucket["index"]}', 'kind': 'scope',
                    'text': f'这条要求管哪一段？\n· {note["text"][:60]}',
                    'where': _where(note),
                    'options': ['管它下面的正文（最多 6 段）', '只管它所在的那一段',
                                '跳过，不套版式'],
                    'default': 1, 'note_index': bucket['index'], 'skip': 2,
                    'entry_index': position,
                })
            container_rules.append(entry)
    for cell in cells.values():
        cell['exclude_p'] = sorted(set(cell['exclude_p']))
        container_rules.append(cell)
    for question in questions:
        if question['kind'] != 'conflict':
            continue
        choice = answers.get(question['id'])
        if choice is None:
            continue
        target, name, values = (question['override']['target'],
                                question['override']['rule'],
                                question['override']['values'])
        bucket = field_rules if isinstance(target, str) else None
        if bucket is not None and target in bucket:
            if choice == 2:
                bucket[target].pop(name, None)
            else:
                bucket[target][name] = values[max(0, min(int(choice), 1))]
    content_rules = []
    for note in notes:
        content = (note['rules'] or {}).get('content')
        if content:
            content_rules.append({'field_key': note.get('field_key') or '',
                                  'label': note.get('field_label') or _where(note),
                                  'text': note['text'], 'content': content})
    return {
        'notes': notes, 'hints': hints, 'general_hints': general_hints,
        'questions': questions, 'field_rules': field_rules,
        'container_rules': container_rules, 'content_rules': content_rules,
        'format_rules': [n for n in notes if has_concrete_rule(n['rules'])],
        '_notes': notes, '_plan': plan,
    }


def build(path, plan, progress=None, answers=None):
    """Return rules, per-field hints, deletions and open questions for one document."""
    return assemble(gather(path, plan, progress=progress), plan, answers)


def reanswer(spec, answers):
    """Re-assemble an in-memory spec after the user answered the open questions."""
    if not spec or '_notes' not in spec:
        return spec
    return assemble(spec['_notes'], spec['_plan'], answers)


def _where(note):
    if note['container'] == 'cell':
        return f'表{int(note["table"] or 0) + 1} 行{int(note["row"] or 0) + 1} ' \
               f'列{int(note["col"] or 0) + 1}'
    if note['container'] in ('header', 'footer'):
        return '页眉' if note['container'] == 'header' else '页脚'
    return f'正文 第{int(note["p"] or 0) + 1}段'


def summarize(spec):
    bits = [f'说明 {len(spec.get("notes") or [])} 条',
            f'格式规则 {len(spec.get("format_rules") or [])} 条',
            f'内容要求 {len(spec.get("content_rules") or [])} 条']
    if spec.get('questions'):
        bits.append(f'待确认 {len(spec["questions"])} 条')
    return ' · '.join(bits)


def to_config(spec, plan, answers=None):
    """The serialisable half of a spec: what a stored template needs at fill time."""
    ranges = {f.get('key'): f.get('bookmark_range') for f in (plan or {}).get('fields', [])
              if f.get('key') and f.get('bookmark_range')}
    return {
        'version': REQ_VERSION,
        'field_rules': spec.get('field_rules') or {},
        'field_ranges': ranges,
        'container_rules': [
            {k: v for k, v in entry.items() if k not in ('note',)}
            for entry in (spec.get('container_rules') or [])],
        'notes': [{k: v for k, v in note.items() if k != 'paragraph_text'}
                  for note in (spec.get('notes') or [])],
        'content_rules': spec.get('content_rules') or [],
        'hints': spec.get('hints') or {},
        'general_hints': spec.get('general_hints') or [],
        'answers': dict(answers or {}),
    }


def public_questions(spec):
    """Questions stripped of internal objects, ready for a dialog."""
    return [{'id': q.get('id'), 'kind': q.get('kind'), 'text': q.get('text'),
             'where': q.get('where'), 'options': list(q.get('options') or []),
             'default': q.get('default', 0)} for q in (spec.get('questions') or [])]


def refresh_plan_req(plan, answers):
    """Re-assemble after the user answered; keeps the plan's serialised spec in sync."""
    spec = reanswer((plan or {}).get('_req_spec') or {}, answers)
    if not spec:
        return plan
    plan['req'] = to_config(spec, spec.get('_plan') or plan, answers)
    plan['req_questions'] = public_questions(spec)
    plan['req_summary'] = summarize(spec)
    plan['_req_spec'] = spec
    return plan


def enhance(source, plan, progress=None, answers=None):
    plan = copy.deepcopy(plan)
    try:
        spec = build(source, plan, progress=progress, answers=answers)
    except Exception as exc:
        plan['req_error'] = str(exc)
        return plan
    plan['req'] = to_config(spec, plan, answers)
    plan['req_questions'] = public_questions(spec)
    plan['req_summary'] = summarize(spec)
    plan['_req_spec'] = spec          # in-memory only, lets the UI re-answer
    return plan
