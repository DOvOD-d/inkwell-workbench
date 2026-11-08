"""Recognize filled-in example templates (colored examples + colored instructions).

The school's task-book template is an example document: example values are blue,
instructions are red, and fixed text is black.  This module turns such a document
into a fillable template without rebuilding the layout:

* label/value table rows become one marker per value cell;
* colored section headings become multiline markers placed after the heading;
* sample prose, references and instructional notes are removed;
* fixed text and row heights are kept.

AI is intentionally not required for the structural pass.  The existing
"AI 识别分类" step can then normalize labels/groups and UI names.
"""
from __future__ import annotations

import copy
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASIC = {'学院', '专业', '届次', '班次', '学生姓名', '学号', '指导教师姓名', '职称学位', '职称/学位'}
NOTE_RE = re.compile(
    r'中文字全部采用|英文、数字采用|行间距|每页都要|首行缩进|悬挂缩进|此处格式参考|'
    r'参考文献不少于|与论文系统上题目一致|请注意此处有变化|此处填写应与|此处空一行|整页设计')
FIXED_HEADINGS = ('要求', '说明', '进度', '安排')
VALUE_HINTS = re.compile(r'X{2,}|20\d{2}\s*年|[0-9]{4,}|年\s*[0-9]{1,2}\s*月|'
                         r'[0-9]{1,2}\s*月\s*[0-9]{1,2}\s*日', re.I)
HEADING_RE = re.compile(r'^（[一二三四五六七八九十]+）')


def _text(p):
    return ''.join(r.text for r in p.runs)


def _norm(text):
    return re.sub(r'\s+', '', text or '')


def _key(label, used=None):
    key = re.sub(r'[^\w\u4e00-\u9fff]+', '', label or '').strip('_')
    key = key[:24] or 'field'
    if used is None:
        return key
    base, n = key, 2
    while key in used:
        key = f'{base}_{n}'
        n += 1
    used.add(key)
    return key


def _run_color(run):
    rpr = run._element.find(qn('w:rPr'))
    if rpr is None:
        return ''
    for color in rpr.findall(qn('w:color')):
        val = (color.get(qn('w:val')) or '').strip()
        if re.fullmatch(r'[0-9A-Fa-f]{6}', val) and val.upper() not in ('000000', 'FFFFFF'):
            return val.upper()
    return ''


def _is_red(color):
    if not color:
        return False
    r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
    return r > 150 and r > g + 40 and r > b + 40 and g < 130


def _is_example_color(color):
    if not color or _is_red(color):
        return False
    r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
    return b > 120 and b > r + 25 and b > g + 10


def _cell_colors(cell):
    colors = []
    for p in cell.paragraphs:
        for run in p.runs:
            if run.text.strip():
                color = _run_color(run)
                if color:
                    colors.append(color)
    return colors


def _has_example(cell):
    return any(_is_example_color(c) for c in _cell_colors(cell))


def _has_red(cell):
    return any(_is_red(c) for c in _cell_colors(cell))


def _red_only(cell):
    saw = False
    for p in cell.paragraphs:
        for run in p.runs:
            if not run.text.strip():
                continue
            saw = True
            if not _is_red(_run_color(run)):
                return False
    return saw


def _looks_like_value(cell):
    text = _norm(cell.text)
    if not text:
        return False
    if _has_example(cell):
        return True
    if VALUE_HINTS.search(text):
        return True
    if len(text) > 20:
        return True
    return False


def _is_label(cell):
    text = _norm(cell.text)
    if not text or len(text) > 20:
        return False
    if _looks_like_value(cell) or _has_red(cell):
        return False
    if any(ch in text for ch in '。；;'):
        return False
    return True


def _unique_cells(row):
    out, seen = [], set()
    for ci, cell in enumerate(row.cells):
        if cell._tc in seen:
            continue
        seen.add(cell._tc)
        out.append((ci, cell))
    return out


def _grid_span(cell):
    tcpr = cell._tc.tcPr
    if tcpr is None:
        return 1
    el = tcpr.find(qn('w:gridSpan'))
    return int(el.get(qn('w:val'))) if el is not None else 1


def _clean_label(text):
    return re.sub(r'[\s：:]+', '', text or '').strip()


def _clean_heading(text):
    t = (text or '').strip()
    t = re.sub(r'^（[一二三四五六七八九十]+）', '', t)
    t = re.sub(r'^[一二三四五六七八九十]+、', '', t)
    t = t.strip(' ：:')
    return t


def _group_for(label):
    if label in BASIC or '姓名' in label or '学号' in label:
        return '基本信息'
    if '题目' in label:
        return '论文题目'
    if any(k in label for k in ('选题', '意义', '创新', '可行')):
        return '选题论证'
    if '研究内容' in label or '类型' in label:
        return '研究内容'
    if '参考' in label:
        return '参考文献'
    if '日期' in label:
        return '落款'
    return '其他'


def _run_index(p, want_color=None):
    for i, run in enumerate(p.runs):
        if not run.text.strip():
            continue
        color = _run_color(run)
        if want_color == 'example' and _is_example_color(color):
            return i
        if want_color == 'black' and not color:
            return i
        if want_color is None:
            return i
    return 0 if p.runs else -1


def _rpr(run):
    return run._element.find(qn('w:rPr')) if run is not None else None


def _clean_rpr(rpr, *, bold=None, size_pt=None, ea_font=None, ascii_font=None):
    if rpr is None:
        if bold is None and size_pt is None and not ea_font and not ascii_font:
            return None
        rpr = OxmlElement('w:rPr')
    else:
        rpr = copy.deepcopy(rpr)
    for tag in ('color', 'highlight', 'shd', 'u', 'em', 'vanish', 'webHidden'):
        for node in list(rpr.findall(qn('w:' + tag))):
            rpr.remove(node)
    if bold is not None:
        for tag in ('b', 'bCs'):
            for node in list(rpr.findall(qn('w:' + tag))):
                rpr.remove(node)
        if bold:
            rpr.append(rpr.makeelement(qn('w:b'), {}))
            rpr.append(rpr.makeelement(qn('w:bCs'), {}))
    if size_pt is not None:
        for tag in ('sz', 'szCs'):
            for node in list(rpr.findall(qn('w:' + tag))):
                rpr.remove(node)
        for tag in ('sz', 'szCs'):
            rpr.append(rpr.makeelement(qn('w:' + tag),
                                       {qn('w:val'): str(int(round(float(size_pt) * 2)))}))
    if ea_font or ascii_font:
        fonts = rpr.find(qn('w:rFonts'))
        if fonts is None:
            fonts = rpr.makeelement(qn('w:rFonts'), {})
            rpr.insert(0, fonts)
        if ea_font:
            fonts.set(qn('w:eastAsia'), ea_font)
        if ascii_font:
            fonts.set(qn('w:ascii'), ascii_font)
            fonts.set(qn('w:hAnsi'), ascii_font)
    return rpr


def _add_run(paragraph, text, rpr):
    run = paragraph.add_run(text)
    if rpr is not None:
        run._element.insert(0, copy.deepcopy(rpr))
    return run


def _clear_paragraph(paragraph):
    for child in list(paragraph._p):
        if child.tag != qn('w:pPr'):
            paragraph._p.remove(child)


def _copy_ppr(target, source):
    src = source._p.find(qn('w:pPr'))
    src = copy.deepcopy(src) if src is not None else None
    old = target._p.find(qn('w:pPr'))
    if old is not None:
        target._p.remove(old)
    if src is not None:
        target._p.insert(0, src)


def _clear_char_indents(paragraph):
    ppr = paragraph._p.find(qn('w:pPr'))
    if ppr is None:
        return
    for ind in ppr.findall(qn('w:ind')):
        for attr in ('firstLineChars', 'leftChars', 'rightChars', 'hangingChars'):
            if ind.get(qn('w:' + attr)) is not None:
                ind.set(qn('w:' + attr), '0')


def _reset_cell(cell):
    first = cell.paragraphs[0]
    _clear_paragraph(first)
    _strip_colors(first._p)
    for p in cell.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    return first


def _set_cell_marker(cell, key, source_p, source_run, *, bold=None, size_pt=None,
                     center=False, center_v=False, clear_indent=False):
    para = _reset_cell(cell)
    _copy_ppr(para, source_p)
    if center:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if clear_indent:
        para.paragraph_format.first_line_indent = Pt(0)
        _clear_char_indents(para)
    if center_v:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    rpr = _clean_rpr(_rpr(source_run), bold=bold, size_pt=size_pt,
                     ea_font='宋体', ascii_font='Times New Roman')
    _add_run(para, '{{ ' + key + ' }}', rpr)


def _insert_marker_before(paragraph, run, text, rpr):
    marker = paragraph.add_run(text)
    if rpr is not None:
        marker._element.insert(0, copy.deepcopy(rpr))
    run._element.addprevious(marker._element)
    return marker


def _replace_cell_prefix(cell, key, source_p, source_run, *, bold=None, size_pt=None,
                         center=False, center_v=False, clear_indent=False):
    """Replace only the leading sample in a cell paragraph; keep instructions and
    every following paragraph (and their formatting) untouched."""
    text = ''.join(r.text for r in source_p.runs)
    note = NOTE_RE.search(text)
    cut = note.start() if note else len(text)
    if not note:
        for hint in ('自选（', '自选(', '备选（', '备选('):
            idx = text.find(hint)
            if idx >= 0:
                cut = idx
                break
    rpr = _clean_rpr(_rpr(source_run), bold=bold, size_pt=size_pt,
                     ea_font='宋体', ascii_font='Times New Roman')
    offset = 0
    inserted = False
    for run in list(source_p.runs):
        value = run.text
        start, end = offset, offset + len(value)
        offset = end
        if end <= cut:
            run._element.getparent().remove(run._element)
        elif start < cut < end:
            run.text = value[cut - start:]
            _insert_marker_before(source_p, run, '{{ ' + key + ' }}', rpr)
            inserted = True
            break
        elif start >= cut:
            _insert_marker_before(source_p, run, '{{ ' + key + ' }}', rpr)
            inserted = True
            break
    if not inserted:
        marker = source_p.add_run('{{ ' + key + ' }}')
        if rpr is not None:
            marker._element.insert(0, copy.deepcopy(rpr))
    if center:
        source_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if clear_indent:
        source_p.paragraph_format.first_line_indent = Pt(0)
        _clear_char_indents(source_p)
    if center_v:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _strip_colors(element):
    for rpr in element.iter(qn('w:rPr')):
        for tag in ('color', 'highlight', 'shd'):
            for node in list(rpr.findall(qn('w:' + tag))):
                rpr.remove(node)


def _remove_textbox_shapes(doc):
    body = doc.element.body
    removed = 0
    for box in list(body.iter(qn('w:txbxContent'))):
        holder = box
        while holder is not None and holder.tag not in (qn('w:r'), qn('w:pict'), qn('w:drawing')):
            holder = holder.getparent()
        if holder is None:
            continue
        run = holder
        while run is not None and run.tag != qn('w:r'):
            run = run.getparent()
        target = run if run is not None else holder
        parent = target.getparent()
        if parent is not None:
            parent.remove(target)
            removed += 1
    return removed


def _row_height_cm(row):
    try:
        return row.height.cm if row.height is not None else 0.0
    except Exception:
        return 0.0


def _heading_prefix(p):
    text = _text(p).strip()
    if not text or HEADING_RE.match(text) is None:
        return None
    has_example = any(_is_example_color(_run_color(r)) for r in p.runs if r.text.strip())
    if not has_example:
        return None
    colon = min([x for x in (text.find('：'), text.find(':')) if x >= 0], default=-1)
    if 0 < colon <= 30 and text[colon + 1:].strip():
        return text[:colon + 1], True
    return text, False


def analyze(path):
    """Return a plan or raise ValueError when the document is not an example template."""
    path = Path(path)
    doc = Document(str(path))
    if not doc.tables:
        raise ValueError('文档里没有表格，无法按“示例模板”识别。')
    ti = max(range(len(doc.tables)), key=lambda i: len(doc.tables[i].rows))
    table = doc.tables[ti]
    if len(table.rows) < 3:
        raise ValueError('表格行数太少，无法识别示例模板。')

    fields, recipes, clear_cells = [], [], []
    used = set()
    field_seq = 0
    full_rows = set()

    # ---------- label/value rows
    for ri, row in enumerate(table.rows):
        cells = _unique_cells(row)
        span = sum(_grid_span(c) for _, c in cells)
        if len(cells) == 1 and span >= max(4, len(table.columns) - 1):
            full_rows.add(ri)
            continue
        last_label = None
        for ci, cell in cells:
            text = _norm(cell.text)
            if not text:
                continue
            if _red_only(cell):
                clear_cells.append([ti, ri, ci])
                continue
            if _is_label(cell):
                last_label = (ci, cell)
                continue
            if not _looks_like_value(cell) or last_label is None:
                continue
            label = _clean_label(last_label[1].text)
            if not label:
                continue
            display = label
            key = _key(label, used)
            field_seq += 1
            fid = f'ex{field_seq}'
            original = cell.text.strip()
            field = {
                'plan_id': fid,
                'key': key,
                'label': display,
                'type': 'text',
                'required': True,
                'group': _group_for(label),
                'mode': 'cell',
                'table': ti,
                'row': ri,
                'col': ci,
                'p': 0,
                'location': {'kind': 'cell', 'table': ti, 'row': ri, 'col': ci, 'p': 0},
                'where': f'表{ti + 1} 行{ri + 1} 列{ci + 1}',
                'original_content': original[:100],
                'context': f'{display}：{original[:80]}',
                'confidence': '示例',
                'source': '示例值',
                'style': {
                    'center': label in ('中文题目', '外文题目', '日期'),
                    'center_v': label in ('中文题目', '外文题目', '日期'),
                    'clear_indent': label in ('中文题目', '外文题目'),
                },
            }
            fields.append(field)

    # ---------- full-width section rows
    row_heights = {}
    section_rows = []
    for ri in sorted(full_rows):
        row = table.rows[ri]
        cell = _unique_cells(row)[0][1]
        col = _unique_cells(row)[0][0]
        paras = list(cell.paragraphs)
        if not paras:
            continue
        nonempty = [i for i, p in enumerate(paras) if _text(p).strip()]
        if not nonempty:
            continue
        title_i = nonempty[0]
        title_p = paras[title_i]
        title_text = _text(title_p).strip()
        title_text_clean = re.sub(r'（请注意此处有变化）', '', title_text).strip()
        title_run = _run_index(title_p)
        recipe = {
            'kind': 'cell', 'table': ti, 'row': ri, 'col': col,
            'title': {'text': title_text_clean, 'ppr': title_i, 'run': title_run},
            'order': ['title'], 'items': {}, 'fixed': [], 'blank': [], 'drop': [],
        }
        item_count = 0
        headings = {}
        for pi in nonempty:
            if pi == title_i:
                continue
            info = _heading_prefix(paras[pi])
            if info:
                headings[pi] = info
        heading_indices = sorted(headings)

        # fixed requirement sections such as （三）要求
        fixed_from = None
        for pi in heading_indices:
            prefix, inline = headings[pi]
            if any(k in _clean_heading(prefix) for k in FIXED_HEADINGS) and '主要研究内容' not in prefix:
                fixed_from = pi
                break

        # colored headings -> fields
        for idx, pi in enumerate(heading_indices):
            prefix, inline = headings[pi]
            label = _clean_heading(prefix)
            if fixed_from is not None and pi >= fixed_from:
                break
            if not label:
                continue
            next_i = heading_indices[idx + 1] if idx + 1 < len(heading_indices) else len(paras)
            sample_indices = []
            if inline:
                rest = _text(paras[pi]).split(prefix, 1)[-1].strip()
                if rest:
                    sample_indices.append(pi)
            sample_indices += [q for q in range(pi + 1, next_i) if _text(paras[q]).strip()]
            sample_text = '\n'.join(_text(paras[q]).strip() for q in sample_indices)
            if not sample_text:
                continue
            ftype = 'text' if '类型' in label else 'multiline'
            key = _key(label, used)
            field_seq += 1
            fid = f'ex{field_seq}'
            body_ppr = pi if inline else (sample_indices[0] if sample_indices else pi)
            body_run = _run_index(paras[body_ppr], want_color='black')
            if body_run < 0:
                body_run = _run_index(paras[body_ppr])
            field = {
                'plan_id': fid, 'key': key, 'label': label, 'type': ftype,
                'required': True, 'group': _group_for(label), 'mode': 'section',
                'table': ti, 'row': ri, 'col': col, 'p': pi,
                'location': {'kind': 'cell', 'table': ti, 'row': ri, 'col': col, 'p': pi},
                'where': f'表{ti + 1} 行{ri + 1} 整行',
                'original_content': sample_text[:100],
                'context': f'{title_text_clean} / {prefix} {sample_text[:80]}',
                'confidence': '示例', 'source': '示例正文',
                'style': {'center': False, 'center_v': False, 'clear_indent': False},
                'recipe_field': fid,
            }
            fields.append(field)
            recipe['items'][fid] = {
                'plan_id': fid, 'heading': prefix if inline else prefix,
                'inline': inline, 'ppr': pi, 'body_ppr': body_ppr,
                'heading_run': _run_index(paras[pi], want_color='example'),
                'body_run': body_run,
            }
            recipe['order'].append('item:' + fid)
            item_count += 1

        # a title-only section such as 主要参考文献
        if item_count == 0 and any(k in title_text_clean for k in ('参考文献', '文献')):
            sample_indices = [q for q in range(title_i + 1, len(paras)) if _text(paras[q]).strip()]
            if sample_indices:
                label = _clean_heading(title_text_clean)
                key = _key(label, used)
                field_seq += 1
                fid = f'ex{field_seq}'
                body_ppr = sample_indices[0]
                field = {
                    'plan_id': fid, 'key': key, 'label': label, 'type': 'multiline',
                    'required': True, 'group': '参考文献', 'mode': 'section',
                    'table': ti, 'row': ri, 'col': col, 'p': body_ppr,
                    'location': {'kind': 'cell', 'table': ti, 'row': ri, 'col': col, 'p': body_ppr},
                    'where': f'表{ti + 1} 行{ri + 1} 整行',
                    'original_content': '\n'.join(_text(paras[q]).strip() for q in sample_indices)[:100],
                    'context': title_text_clean,
                    'confidence': '示例', 'source': '示例正文',
                    'style': {'center': False, 'center_v': False, 'clear_indent': False},
                    'recipe_field': fid,
                }
                fields.append(field)
                recipe['items'][fid] = {
                    'plan_id': fid, 'heading': '', 'inline': False,
                    'ppr': body_ppr, 'body_ppr': body_ppr,
                    'heading_run': -1,
                    'body_run': _run_index(paras[body_ppr], want_color='black'),
                }
                recipe['order'].append('item:' + fid)
                item_count += 1

        # fixed / blank / drop order
        if item_count:
            # keep requirement paragraphs after a fixed heading
            if fixed_from is not None:
                for q in range(fixed_from, len(paras)):
                    text = _text(paras[q]).strip()
                    if not text or NOTE_RE.search(text):
                        continue
                    recipe['order'].append(f'fixed:{q}')
                    recipe['fixed'].append(q)
            # drop everything not already used
            used_paras = {title_i}
            used_paras |= set(heading_indices)
            for token in recipe['order']:
                if token.startswith('item:'):
                    item = recipe['items'][token.split(':', 1)[1]]
                    used_paras.add(item['ppr'])
                    used_paras.add(item['body_ppr'])
                elif token.startswith('fixed:'):
                    used_paras.add(int(token.split(':', 1)[1]))
            recipe['drop'] = [q for q in range(len(paras)) if q not in used_paras]
        else:
            # no field: preserve non-instruction text, blank a lone red spacer
            for q in range(title_i + 1, len(paras)):
                text = _text(paras[q]).strip()
                if not text:
                    continue
                if '此处空一行' in text:
                    recipe['order'].append(f'blank:{q}')
                    recipe['blank'].append(q)
                elif NOTE_RE.search(text):
                    recipe['drop'].append(q)
                else:
                    recipe['order'].append(f'fixed:{q}')
                    recipe['fixed'].append(q)
            recipe['drop'] += [q for q in range(len(paras))
                               if q not in {title_i, *recipe['fixed'], *recipe['blank']}
                               and q not in recipe['drop']]

        # only keep a recipe if it actually rebuilds something
        if item_count or recipe['fixed'] or recipe['blank']:
            recipes.append(recipe)
            if item_count:
                section_rows.append(ri)

    if not fields:
        raise ValueError('没有识别到可填写项。\n'
                         '这张模板可能没有彩色示例值，或结构不是“标签+示例内容”。\n'
                         '可以先用“侦察分析”检查空白位，或在 Word 中给填写处加 {{ 名称 }}。')

    # ---------- page-aware row heights for blank section rows
    try:
        section = doc.sections[0]
        content_cm = section.page_height.cm - section.top_margin.cm - section.bottom_margin.cm
        before = sum(_row_height_cm(table.rows[i]) for i in range(min(section_rows)))
        first_large = next((ri for ri in range(min(section_rows), len(table.rows))
                            if ri not in section_rows and _row_height_cm(table.rows[ri]) >= 8), None)
        page_rows = [ri for ri in section_rows if first_large is None or ri < first_large]
        available = content_cm - before - 1.7
        if len(page_rows) == 2 and 8 <= available <= 30:
            row_heights[page_rows[0]] = round(available * 0.44, 1)
            row_heights[page_rows[1]] = round(available * 0.56, 1)
        else:
            for ri in page_rows:
                row_heights[ri] = 4.0
        for ri in section_rows:
            if ri not in row_heights and _row_height_cm(table.rows[ri]) < 4:
                row_heights[ri] = 4.0
    except Exception:
        for ri in section_rows:
            row_heights[ri] = 4.0

    # ---------- title / description
    title = next((p.text.strip() for p in doc.paragraphs if p.text.strip()), path.stem)
    title = title[:60]
    all_text = '\n'.join(p.text for p in doc.paragraphs)
    all_text += '\n' + '\n'.join(c.text for t in doc.tables for row in t.rows for c in row.cells)
    bits = [f'由示例模板识别生成，共 {len(fields)} 个填写项；固定格式与进度安排保留，'
            '示例值、示例正文和红色说明已清理。']
    if '宋体' in all_text:
        bits.append('中文正文按原模板宋体五号、行距固定值20磅；短字段宋体小四加粗。')
    if '参考文献不少于10篇' in all_text:
        bits.append('参考文献不少于10篇。')
    description = ''.join(bits)
    category = '教学' if any(k in title for k in ('任务书', '论文', '学院')) else '自定义模板'
    for field in fields:
        if field['label'] == '论文（设计）类型':
            field['label'] = '论文（设计）类型（如Ⅰ（1）非公开发表论文）'
    group_order = {'基本信息': 0, '论文题目': 1, '选题论证': 2, '研究内容': 3,
                   '参考文献': 4, '落款': 5, '其他': 6}
    fields.sort(key=lambda f: group_order.get(f['group'], 9))

    return {
        'source': str(path),
        'title': title,
        'category': category,
        'description': description,
        'table': ti,
        'fields': fields,
        'recipes': recipes,
        'clear_cells': clear_cells,
        'row_heights': row_heights,
        'remove_textboxes': False,
    }


# ---------------------------------------------------------------- build

def _rebuild_recipe(cell, recipe, selected, source_paras):
    """Rebuild a full-width cell from recipe.order; unselected fields disappear."""
    def clone_ppr(p):
        return copy.deepcopy(p._p.find(qn('w:pPr')))

    def clone_run_rpr(p, index, clean=True, **kw):
        run = p.runs[index] if 0 <= index < len(p.runs) else (p.runs[0] if p.runs else None)
        if not clean:
            return copy.deepcopy(_rpr(run)) if run is not None else None
        return _clean_rpr(_rpr(run), ea_font='宋体', ascii_font='Times New Roman', **kw)

    title = recipe['title']
    title_src = source_paras[title['ppr']]
    title_ppr = clone_ppr(title_src)
    title_rpr = clone_run_rpr(title_src, title['run'], clean=False)
    fixed_clones = {q: copy.deepcopy(source_paras[q]._p) for q in recipe['fixed']}
    blank_pprs = {q: clone_ppr(source_paras[q]) for q in recipe['blank']}
    item_data = {}
    for pid, item in recipe['items'].items():
        hs = source_paras[item['ppr']]
        bs = source_paras[item['body_ppr']]
        item_data[pid] = {
            'heading_ppr': clone_ppr(hs),
            'body_ppr': clone_ppr(bs),
            'heading_rpr': clone_run_rpr(hs, item['heading_run'], clean=False) if item['heading'] else None,
            'body_rpr': clone_run_rpr(bs, item['body_run'], bold=False,
                                      size_pt=item.get('body_size_pt')),
        }

    first = _reset_cell(cell)
    made_first = False

    def new_para(ppr):
        nonlocal made_first
        if not made_first:
            para = first
            made_first = True
        else:
            para = cell.add_paragraph()
        old = para._p.find(qn('w:pPr'))
        if old is not None:
            para._p.remove(old)
        if ppr is not None:
            para._p.insert(0, copy.deepcopy(ppr))
        _strip_colors(para._p)
        return para

    for token in recipe['order']:
        if token == 'title':
            para = new_para(title_ppr)
            _add_run(para, title['text'], title_rpr)
        elif token.startswith('item:'):
            pid = token.split(':', 1)[1]
            if pid not in selected:
                continue
            item = recipe['items'][pid]
            data = item_data[pid]
            marker = '{{ ' + selected[pid] + ' }}'
            if item['inline'] and item['heading']:
                para = new_para(data['heading_ppr'])
                _add_run(para, item['heading'], data['heading_rpr'])
                _add_run(para, marker, data['body_rpr'])
            else:
                if item['heading']:
                    para = new_para(data['heading_ppr'])
                    _add_run(para, item['heading'], data['heading_rpr'])
                para = new_para(data['body_ppr'])
                if item.get('align') == 'left':
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    first = para.paragraph_format.first_line_indent
                    if first is not None and first > Pt(30):
                        para.paragraph_format.first_line_indent = Pt(0)
                        _clear_char_indents(para)
                    left = para.paragraph_format.left_indent
                    if left is not None and left > Pt(30):
                        para.paragraph_format.left_indent = Pt(0)
                        _clear_char_indents(para)
                _add_run(para, marker, data['body_rpr'])
        elif token.startswith('fixed:'):
            q = int(token.split(':', 1)[1])
            if not made_first:
                first._p.getparent().remove(first._p)
                made_first = True
            clone = fixed_clones[q]
            cell._tc.append(clone)
        elif token.startswith('blank:'):
            q = int(token.split(':', 1)[1])
            new_para(blank_pprs[q])


def build(source, out, plan, selected):
    """Build a marked template.

    selected: {plan_id: marker_key}.  Unselected label/value cells are cleared;
    unselected section fields are omitted.
    """
    source, out = Path(source), Path(out)
    doc = Document(str(source))
    tables = doc.tables

    # label/value cells
    for field in plan['fields']:
        if field['mode'] != 'cell':
            continue
        ti = int(field.get('table', plan.get('table', 0)))
        ri, ci = field['row'], field['col']
        table = tables[ti]
        cell = table.rows[ri].cells[ci]
        if field['plan_id'] not in selected:
            _reset_cell(cell)
            continue
        p = cell.paragraphs[0]
        run = next((r for r in p.runs if r.text.strip()), p.runs[0] if p.runs else None)
        style = field.get('style') or {}
        if field.get('preserve_suffix'):
            _replace_cell_prefix(cell, selected[field['plan_id']], p, run,
                                 bold=style.get('marker_bold'), size_pt=style.get('marker_size'),
                                 center=style.get('center'), center_v=style.get('center_v'),
                                 clear_indent=style.get('clear_indent'))
        else:
            _set_cell_marker(cell, selected[field['plan_id']], p, run,
                             bold=style.get('marker_bold'), size_pt=style.get('marker_size'),
                             center=style.get('center'), center_v=style.get('center_v'),
                             clear_indent=style.get('clear_indent'))

    # full-width cells
    for recipe in plan.get('recipes', []):
        ti = int(recipe.get('table', plan.get('table', 0)))
        ri, ci = recipe['row'], recipe['col']
        cell = tables[ti].rows[ri].cells[ci]
        _rebuild_recipe(cell, recipe, selected, list(cell.paragraphs))

    # clear red-only cells such as “手签”
    for item in plan.get('clear_cells', []):
        ti, ri, ci = item
        if ti < len(tables) and ri < len(tables[ti].rows):
            cells = tables[ti].rows[ri].cells
            if ci < len(cells):
                _reset_cell(cells[ci])

    # min row heights (python-docx converts EMU to the twips Word expects)
    for key, cm in (plan.get('row_heights') or {}).items():
        if isinstance(key, (tuple, list)):
            ti, ri = int(key[0]), int(key[1])
        else:
            ti, ri = int(plan.get('table', 0)), int(key)
        row = tables[ti].rows[ri]
        row.height = Cm(float(cm))
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    if plan.get('remove_textboxes'):
        _remove_textbox_shapes(doc)

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))

    import engine
    markers = engine.scan_fields(str(out))
    return {'out': str(out), 'markers': markers, 'fields': len(markers)}
