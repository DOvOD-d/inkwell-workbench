"""手动排版面板的落地层：页面设置 + 正文/标题/表格/页眉页脚按档套用。

分工：:mod:`formatapply` 执行“模板自己写明的要求”（书签定位、删说明），本模块执行
“用户在排版规则里手动选的参数”，可以作用到整篇文档，包括纸张与页边距。
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Mm
from docx.enum.section import WD_ORIENT

import cleanup
import formatapply
import ooxmlcheck
import units

# 纸张按纵向毫米尺寸给出；横向时交换宽高。
PAPERS = {
    'A3': (297.0, 420.0),
    'A4': (210.0, 297.0),
    'A5': (148.0, 210.0),
    'B5': (176.0, 250.0),
    'Letter': (215.9, 279.4),
}
MARGIN_PRESETS = {
    '普通': {'top': 2.54, 'bottom': 2.54, 'left': 3.18, 'right': 3.18},
    '窄': {'top': 1.27, 'bottom': 1.27, 'left': 1.27, 'right': 1.27},
    '宽': {'top': 2.54, 'bottom': 2.54, 'left': 4.0, 'right': 4.0},
    '学校表格': {'top': 2.0, 'bottom': 2.0, 'left': 2.5, 'right': 2.0},
}
SIDE_NAMES = {'top': '上边距', 'bottom': '下边距', 'left': '左边距', 'right': '右边距'}
BLOCK_KEYS = ('body', 'heading1', 'heading2', 'heading3', 'table', 'header_footer')
BLOCK_NAMES = {
    'body': '正文',
    'heading1': '一级标题',
    'heading2': '二级标题',
    'heading3': '三级标题',
    'table': '表格内文字',
    'header_footer': '页眉页脚',
}
SCOPE_NAMES = {'off': '不改', 'all': '整篇文档', 'filled': '仅填写内容'}
RULE_FIELDS = ('font_east', 'font_ascii', 'size_pt', 'bold', 'color', 'line',
               'first_line_chars', 'hanging_chars', 'space_before', 'space_after', 'align',
               'keep_next', 'keep_lines', 'widow_control', 'shading_fill', 'highlight', 'mark')
HEAD_STYLE_RE = re.compile(r'^(?:heading|标题)\s*([1-6])$', re.IGNORECASE)


def _mm(value):
    """python-docx 长度对象 → 毫米，用于和 PAPERS 表比对。"""
    return round(value / 36000.0, 1) if value is not None else None


def read_page(doc):
    """读回第一分节的纸张与页边距，供界面显示初始值。"""
    section = doc.sections[0]
    width, height = _mm(section.page_width), _mm(section.page_height)
    landscape = bool(width and height and width > height)
    long_side, short_side = (max(width, height), min(width, height)) if width and height else (0, 0)
    paper = '自定义'
    for name, (short, long) in PAPERS.items():
        if abs(short - short_side) <= 2.0 and abs(long - long_side) <= 2.0:
            paper = name
            break
    margins = {'top': round(section.top_margin / 360000.0, 2),
               'bottom': round(section.bottom_margin / 360000.0, 2),
               'left': round(section.left_margin / 360000.0, 2),
               'right': round(section.right_margin / 360000.0, 2)}
    return {'paper': paper, 'landscape': landscape, 'margins': margins,
            'size_mm': [short_side, long_side],
            'sections': len(doc.sections)}


def set_page(doc, paper='A4', landscape=False, margins=None):
    """把纸张与页边距写到所有分节。返回改动的分节数。

    纸张和页边距先按 Word 自己允许的区间校验，越界直接报中文错误，
    免得写出一份 Word 打不开或提示“文件已损坏”的文档。
    """
    if paper in PAPERS:
        short, long = PAPERS[paper]
        width, height = (long, short) if landscape else (short, long)
    else:
        current = read_page(doc)
        width, height = current['size_mm']
        if landscape != current['landscape']:
            width, height = height, width
    width = units.parse_page_dim(f'{width}mm', '纸张宽度') / units.TWIPS_PER_MM
    height = units.parse_page_dim(f'{height}mm', '纸张高度') / units.TWIPS_PER_MM
    margins = margins or MARGIN_PRESETS['普通']
    checked = {}
    for side, value in margins.items():
        if value in (None, ''):
            continue
        checked[side] = units.parse_margin(f'{value}cm', SIDE_NAMES.get(side, '页边距')) \
            / units.TWIPS_PER_CM
    changed = 0
    for section in doc.sections:
        section.page_width = Mm(width)
        section.page_height = Mm(height)
        section.orientation = (WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT)
        for side, value in checked.items():
            setattr(section, side + '_margin', Cm(value))
        changed += 1
    return changed


def heading_level(paragraph):
    """标题级别：优先 Word 样式名与大纲级别，认不出当作正文（返回 0）。"""
    ppr = paragraph._p.find(qn('w:pPr'))
    if ppr is None:
        return 0
    style = ppr.find(qn('w:pStyle'))
    if style is not None:
        value = (style.get(qn('w:val')) or '').strip()
        match = HEAD_STYLE_RE.match(value)
        if match:
            return int(match.group(1))
        if value in ('1', '2', '3', '4', '5', '6'):
            # 中文模板里 “标题 1” 的 styleId 常常就是裸数字
            name = _style_name(paragraph.part, value) or ''
            if re.search(r'标题|heading|title', name, re.IGNORECASE):
                return int(value)
    outline = ppr.find(qn('w:outlineLvl'))
    if outline is not None:
        try:
            level = int(outline.get(qn('w:val')))
        except (TypeError, ValueError):
            return 0
        return level + 1 if 0 <= level <= 8 else 0
    return 0


def _style_name(part, style_id):
    try:
        from docx.oxml.ns import qn as _qn
        styles = part.document.styles.element
        for style in styles.findall(_qn('w:style')):
            if style.get(_qn('w:styleId')) == style_id:
                name = style.find(_qn('w:name'))
                return name.get(_qn('w:val')) if name is not None else ''
    except Exception:
        return None
    return None


def paragraph_outline(doc, max_items=120):
    """给思考模型看的段落清单：编号、样式名、字号、加粗、对齐。"""
    items = []
    for index, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue
        size = None
        bold = False
        for run in paragraph.runs[:3]:
            if run.font.size is not None:
                size = round(run.font.size.pt, 1)
            if run.font.bold:
                bold = True
        align = None
        try:
            if paragraph.alignment is not None:
                align = str(paragraph.alignment).split('.')[-1].split(' ')[0].lower()
        except Exception:
            align = None
        items.append({'i': index, 'text': text,
                      'style': paragraph.style.name if paragraph.style is not None else '',
                      'size': size, 'bold': bold, 'align': align})
        if len(items) >= max_items:
            break
    return items


def _body_paragraphs(doc):
    return [p for p in doc.paragraphs if p.text.strip()]


def _table_paragraphs(doc):
    found = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                found.extend(p for p in cell.paragraphs if p.text.strip())
    return found


def _header_footer_paragraphs(doc):
    found = []
    for section in doc.sections:
        for part in (section.header, section.footer, section.first_page_header,
                     section.first_page_footer, section.even_page_header,
                     section.even_page_footer):
            try:
                found.extend(p for p in part.paragraphs if p.text.strip())
            except (AttributeError, NotImplementedError):
                continue
    return found


def _paragraphs_for(doc, key, scope, levels=None):
    """按档取段落；scope='filled' 时只保留落在填写位书签里的那些。

    levels: {doc.paragraphs 下标: 级别}，来自思考模型的判定；Word 样式仍优先。
    """
    levels = levels or {}

    def level_of(index, paragraph):
        return levels.get(index) or heading_level(paragraph)

    if key == 'table':
        pool = _table_paragraphs(doc)
    elif key == 'header_footer':
        pool = _header_footer_paragraphs(doc)
    elif key.startswith('heading'):
        want = int(key[-1])
        pool = [p for i, p in enumerate(doc.paragraphs)
                if p.text.strip() and level_of(i, p) == want]
    else:
        pool = [p for i, p in enumerate(doc.paragraphs)
                if p.text.strip() and not level_of(i, p)]
    if scope != 'filled':
        return pool, False
    filled = [p for p in pool if formatapply._has_bookmark(p)]
    return filled, bool(pool) and not filled


def rules_present(rules):
    return any(rules.get(field) not in (None, '', {}, []) or
               (field == 'bold' and 'bold' in rules) for field in RULE_FIELDS)


def clear_indent(paragraph):
    """“不缩进”要真的清掉段落原有的首行/悬挂缩进，而不是只写一个新值。"""
    ppr = paragraph._p.find(qn('w:pPr'))
    ind = ppr.find(qn('w:ind')) if ppr is not None else None
    if ind is None:
        return False
    cleared = formatapply.fold_indent_aliases(ind)
    for attr in ('firstLineChars', 'firstLine', 'hangingChars', 'hanging'):
        if qn('w:' + attr) in ind.attrib:
            del ind.attrib[qn('w:' + attr)]
            cleared = True
    return cleared


def _apply_paragraph(paragraph, rules, report):
    line = rules.get('line')
    before, after = rules.get('space_before'), rules.get('space_after')
    if line or before is not None or after is not None:
        formatapply.set_spacing(paragraph, line, before, after)
        report['spacing'] += 1
    indent = rules.get('first_line_chars')
    if indent:
        formatapply.set_indent(paragraph, first_chars=float(indent))
        report['indent'] += 1
    elif indent == 0 and clear_indent(paragraph):
        report['indent'] += 1
    if rules.get('align'):
        formatapply.set_align(paragraph, rules['align'])
        report['align'] += 1
    if any(rules.get(key) is not None for key in ('keep_next', 'keep_lines', 'widow_control')):
        formatapply.set_pagination(paragraph, rules.get('keep_next'), rules.get('keep_lines'),
                                   rules.get('widow_control'))
        report['pagination'] = report.get('pagination', 0) + 1
    ppr = paragraph._p.get_or_add_pPr()
    if rules.get('shading_fill') is not None:
        formatapply.set_shading(ppr, rules['shading_fill'])
        report['shading'] = report.get('shading', 0) + 1
    if rules.get('mark'):
        formatapply.apply_mark_rules(paragraph, rules['mark'])
        report['mark'] = report.get('mark', 0) + 1
    run_rules = {k: v for k, v in rules.items()
                 if k in ('font_east', 'font_ascii', 'size_pt', 'bold', 'italic', 'color')}
    for run in paragraph.runs:
        if run_rules and formatapply.apply_run_rules(run._r, run_rules,
                                                     keep_color=not run_rules.get('color')):
            report['runs'] += 1
        rpr = run._r.find(qn('w:rPr'))
        if rpr is not None and rules.get('shading_fill') is not None:
            formatapply.set_shading(rpr, rules['shading_fill'])
    if rules.get('highlight'):
        for run in paragraph.runs:
            formatapply.set_highlight(run._r, rules['highlight'])
            report['highlight'] = report.get('highlight', 0) + 1
    report['paragraphs'] += 1


def apply_layout(doc, blocks, page=None, heading_levels=None):
    """blocks: [{'key','scope','rules'}]；page: None 或 {'paper','landscape','margins'}。

    返回报告：每档命中多少段、跳过了什么原因，供界面如实说明。
    """
    report = {'sections': 0, 'paragraphs': 0, 'runs': 0, 'spacing': 0, 'indent': 0,
              'align': 0, 'blocks': {}, 'skipped': []}
    if page:
        report['sections'] = set_page(doc, page.get('paper') or 'A4',
                                      bool(page.get('landscape')), page.get('margins'))
    for block in blocks or []:
        key = block.get('key')
        scope = block.get('scope') or 'off'
        rules = block.get('rules') or {}
        if key not in BLOCK_KEYS or scope == 'off' or not rules_present(rules):
            continue
        pool, empty_filled = _paragraphs_for(doc, key, scope, heading_levels)
        if not pool:
            if scope == 'filled':
                reason = ('文档里没有填写位书签，这一档没改动（可改用“整篇文档”）' if empty_filled
                          else '没有落在填写位里的内容，这一档没改动')
            else:
                reason = '文档里没有该档段落'
            report['skipped'].append(BLOCK_NAMES[key] + '：' + reason)
            continue
        for paragraph in pool:
            _apply_paragraph(paragraph, rules, report)
        report['blocks'][key] = len(pool)
    return report


def empty_rules():
    return {field: None for field in RULE_FIELDS}


def default_blocks():
    """界面初始状态：正文默认整篇套用（参数由用户填），其余档关闭。"""
    return [{'key': key, 'scope': 'all' if key == 'body' else 'off', 'rules': empty_rules()}
            for key in BLOCK_KEYS]


class Cancelled(RuntimeError):
    """用户中途取消排版：不保存、不留半成品，原文件一定没动。"""


def apply_to_file(source, out, blocks, page=None, heading_levels=None,
                  spec=None, options=None, should_stop=None):
    """从原始文档复制一份再套用规则：原文件永不改动，重复应用结果一致。

    spec/options 来自模板已解析出的要求（req）：会把说明删掉、把段落规则落到对应栏目，
    这样导出的文件和填写页生成的结果一致。格式套用不改动段落数量，所以先排版后删说明是安全的。

    should_stop 在阶段之间检查；一旦取消就删掉刚复制出来的中间副本并抛 Cancelled。
    """
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(source), str(out))

    def checkpoint(message):
        if should_stop and should_stop():
            try:
                out.unlink()
            except OSError:
                pass
            raise Cancelled(message)

    doc = Document(str(out))
    report = apply_layout(doc, blocks, page, heading_levels)
    checkpoint('已取消：本次排版没有写入，原文件未改动。')
    if spec:
        import formatapply
        pre = formatapply.apply(doc, spec, options or {}, phase='pre')
        report['removed'] = pre.get('removed', 0)
        report['notes_kept'] = pre.get('kept', 0)
        report['blank_lines'] = pre.get('blank_lines', 0)
        report['notes_skipped'] = pre.get('skipped') or []
        checkpoint('已取消：说明与格式的调整没有写入，原文件未改动。')
    cleaned = cleanup.clean(doc)
    report['textboxes_cleaned'] = cleaned['textboxes']
    report['marks_cleared'] = cleaned['highlights']
    ooxmlcheck.normalize(doc)
    doc.save(str(out))
    report['output'] = str(out)
    return report


# 样式库：点一张卡就把参数灌进左栏对应档，参数仍可改。
STYLE_PRESETS = [
    {'id': 'body', 'name': '正文', 'block': 'body',
     'summary': '宋体 / 小四 / 1.5 倍行距 / 首行缩进 2 字符',
     'rules': {'font_east': '宋体', 'font_ascii': 'Times New Roman', 'size_pt': 12.0,
               'line': {'type': 'multiple', 'value': 1.5}, 'first_line_chars': 2.0,
               'align': 'justify', 'bold': None, 'color': None, 'italic': None}},
    {'id': 'h1', 'name': '标题 1', 'block': 'heading1',
     'summary': '黑体 / 三号 / 居中 / 加粗',
     'rules': {'font_east': '黑体', 'font_ascii': 'Times New Roman', 'size_pt': 16.0,
               'line': {'type': 'multiple', 'value': 1.5}, 'first_line_chars': 0,
               'align': 'center', 'bold': True, 'color': None, 'italic': None}},
    {'id': 'h2', 'name': '标题 2', 'block': 'heading2',
     'summary': '黑体 / 四号 / 左对齐 / 加粗',
     'rules': {'font_east': '黑体', 'font_ascii': 'Times New Roman', 'size_pt': 14.0,
               'line': {'type': 'multiple', 'value': 1.5}, 'first_line_chars': 0,
               'align': 'left', 'bold': True, 'color': None, 'italic': None}},
    {'id': 'h3', 'name': '标题 3', 'block': 'heading3',
     'summary': '黑体 / 小四 / 左对齐 / 加粗',
     'rules': {'font_east': '黑体', 'font_ascii': 'Times New Roman', 'size_pt': 12.0,
               'line': {'type': 'multiple', 'value': 1.5}, 'first_line_chars': 0,
               'align': 'left', 'bold': True, 'color': None, 'italic': None}},
    {'id': 'quote', 'name': '引文', 'block': 'body',
     'summary': '楷体 / 小四 / 单倍行距 / 左右各缩进',
     'rules': {'font_east': '楷体', 'font_ascii': 'Times New Roman', 'size_pt': 12.0,
               'line': {'type': 'multiple', 'value': 1.0}, 'first_line_chars': 2.0,
               'align': 'justify', 'bold': None, 'color': None, 'italic': None}},
    {'id': 'toc', 'name': '目录', 'block': 'body',
     'summary': '宋体 / 小四 / 1.5 倍行距 / 两端对齐',
     'rules': {'font_east': '宋体', 'font_ascii': 'Times New Roman', 'size_pt': 12.0,
               'line': {'type': 'multiple', 'value': 1.5}, 'first_line_chars': 0,
               'align': 'justify', 'bold': None, 'color': None, 'italic': None}},
    {'id': 'table', 'name': '表格文字', 'block': 'table',
     'summary': '宋体 / 五号 / 单倍行距 / 不缩进',
     'rules': {'font_east': '宋体', 'font_ascii': 'Times New Roman', 'size_pt': 10.5,
               'line': {'type': 'multiple', 'value': 1.0}, 'first_line_chars': 0,
               'align': None, 'bold': None, 'color': None, 'italic': None}},
    {'id': 'header_footer', 'name': '页眉页脚', 'block': 'header_footer',
     'summary': '宋体 / 小五 / 单倍行距 / 居中',
     'rules': {'font_east': '宋体', 'font_ascii': 'Times New Roman', 'size_pt': 9.0,
               'line': {'type': 'multiple', 'value': 1.0}, 'first_line_chars': 0,
               'align': 'center', 'bold': None, 'color': None, 'italic': None}},
]


def preset_rules(preset_id):
    for preset in STYLE_PRESETS:
        if preset['id'] == preset_id:
            return dict(preset['rules']), preset['block']
    return None, None


LINE_CHOICES = [('1 倍', {'type': 'multiple', 'value': 1.0}),
                ('1.25 倍', {'type': 'multiple', 'value': 1.25}),
                ('1.5 倍', {'type': 'multiple', 'value': 1.5}),
                ('2 倍', {'type': 'multiple', 'value': 2.0}),
                ('固定值 20 磅', {'type': 'exact', 'value': 20.0}),
                ('固定值 22 磅', {'type': 'exact', 'value': 22.0}),
                ('最小值 15 磅', {'type': 'at_least', 'value': 15.0})]
ALIGN_CHOICES = [('左对齐', 'left'), ('居中', 'center'), ('两端对齐', 'justify'),
                 ('右对齐', 'right')]
INDENT_CHOICES = [('不缩进', 0), ('2 字符', 2.0), ('空 2 格', 2.0), ('4 字符', 4.0)]
