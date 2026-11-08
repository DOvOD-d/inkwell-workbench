"""Apply a template's own stated requirements to the document being generated.

Order matters: paragraph indices are only valid while nothing has been removed, so
this module formats first and deletes the guidance text afterwards.

Everything here is deterministic — the rules come from :mod:`reqparse`, which parses
the Chinese sentences itself; no model ever names a font directly.
"""
from __future__ import annotations

import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import bookmarkfill

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PARA_KEYS = ('line', 'first_line_chars', 'hanging_chars', 'align')


def _child(parent, tag, insert=True):
    found = parent.find(qn('w:' + tag))
    if found is not None or not insert:
        return found
    element = OxmlElement('w:' + tag)
    parent.append(element)
    return element


def _pPr(paragraph):
    ppr = paragraph._p.find(qn('w:pPr'))
    if ppr is None:
        ppr = OxmlElement('w:pPr')
        paragraph._p.insert(0, ppr)
    return ppr


def _order_ppr(ppr):
    """w:pPr children must follow the schema order or Word repairs the file."""
    order = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
             'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
             'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
             'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
             'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
             'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
             'sectPr', 'pPrChange']
    rank = {name: i for i, name in enumerate(order)}
    children = sorted(list(ppr), key=lambda el: rank.get(el.tag.split('}')[-1], 99))
    for child in children:
        ppr.append(child)


def set_spacing(paragraph, line, before=None, after=None):
    """``line`` is {'type': exact|at_least|multiple, 'value': float}; before/after are points."""
    spacing = _child(_pPr(paragraph), 'spacing')
    if line:
        kind = line.get('type')
        value = float(line.get('value') or 0)
        if kind in ('exact', 'at_least'):
            spacing.set(qn('w:line'), str(int(round(value * 20))))
            spacing.set(qn('w:lineRule'), 'exact' if kind == 'exact' else 'atLeast')
        else:
            spacing.set(qn('w:line'), str(int(round(max(value, 0.5) * 240))))
            spacing.set(qn('w:lineRule'), 'auto')
    for key, name in ((before, 'before'), (after, 'after')):
        if key is None:
            continue
        spacing.set(qn('w:' + name), str(int(round(float(key) * 20))))
    _order_ppr(spacing.getparent())
    return bool(line) or before is not None or after is not None


def set_indent(paragraph, first_chars=None, hanging_chars=None):
    if not first_chars and not hanging_chars:
        return False
    ind = _child(_pPr(paragraph), 'ind')
    fold_indent_aliases(ind)
    if hanging_chars:
        hundredths = str(int(round(float(hanging_chars) * 100)))
        ind.set(qn('w:hangingChars'), hundredths)
        ind.set(qn('w:hanging'), str(int(round(float(hanging_chars) * 210))))
        for attr in ('firstLineChars', 'firstLine'):
            if qn('w:' + attr) in ind.attrib:
                del ind.attrib[qn('w:' + attr)]
    else:
        hundredths = str(int(round(float(first_chars) * 100)))
        ind.set(qn('w:firstLineChars'), hundredths)
        ind.set(qn('w:firstLine'), str(int(round(float(first_chars) * 210))))
        for attr in ('hangingChars', 'hanging'):
            if qn('w:' + attr) in ind.attrib:
                del ind.attrib[qn('w:' + attr)]
    _order_ppr(ind.getparent())
    return True


ALIGN_VALUES = {'center': 'center', 'left': 'left', 'right': 'right', 'justify': 'both'}


def fold_indent_aliases(ind):
    """w:start / w:end 折进 w:left / w:right。

    OOXML 的横向缩进有两套写法，Word 两种都读；一个元素里同时写两套，
    下一次规范化保存会随机保留一边，表现就是“缩进莫名其妙没了”。
    """
    changed = False
    for strict, transitional in (('start', 'left'), ('end', 'right')):
        value = ind.get(qn('w:' + strict))
        if value is None:
            continue
        if ind.get(qn('w:' + transitional)) is None:
            ind.set(qn('w:' + transitional), value)
        del ind.attrib[qn('w:' + strict)]
        changed = True
    return changed


def set_pagination(paragraph, keep_next=None, keep_lines=None, widow_control=None):
    """分页三兄弟：None 表示不动。"""
    wanted = (('keepNext', keep_next), ('keepLines', keep_lines), ('widowControl', widow_control))
    changed = False
    ppr = _pPr(paragraph)
    for tag, value in wanted:
        if value is None:
            continue
        node = _child(ppr, tag)
        node.set(qn('w:val'), 'true' if value else 'false')
        changed = True
    if changed:
        _order_ppr(ppr)
    return changed


def set_shading(holder, fill, pattern='clear'):
    """给 pPr / rPr / tcPr 加底纹；fill 为空表示清掉。"""
    node = holder.find(qn('w:shd'))
    if not fill:
        if node is None:
            return False
        holder.remove(node)
        return True
    if node is None:
        node = OxmlElement('w:shd')
        holder.append(node)
    node.set(qn('w:val'), pattern or 'clear')
    node.set(qn('w:color'), 'auto')
    node.set(qn('w:fill'), str(fill).lstrip('#').upper())
    return True


def set_highlight(run, value):
    """run 的高亮色；'none' 清掉，None 不动。"""
    if value is None:
        return False
    rpr = run.find(qn('w:rPr'))
    if rpr is None:
        if value == 'none':
            return False
        rpr = OxmlElement('w:rPr')
        run.insert(0, rpr)
    node = rpr.find(qn('w:highlight'))
    if str(value).lower() == 'none':
        if node is None:
            return False
        rpr.remove(node)
        return True
    if node is None:
        node = OxmlElement('w:highlight')
        rpr.append(node)
    node.set(qn('w:val'), str(value).lower())
    return True


def apply_mark_rules(paragraph, rules):
    """段落标记（¶）自己的格式——决定回车后新打出来的字继承什么。"""
    if not rules:
        return False
    ppr = _pPr(paragraph)
    changed = apply_run_rules(ppr, rules)
    _order_ppr(ppr)
    return changed


def set_align(paragraph, value):
    if not value:
        return False
    jc = _child(_pPr(paragraph), 'jc')
    jc.set(qn('w:val'), ALIGN_VALUES.get(value, 'both'))
    _order_ppr(jc.getparent())
    return True


def set_cell_valign(cell, value):
    if not value:
        return False
    tc_pr = cell._tc.find(qn('w:tcPr'))
    if tc_pr is None:
        tc_pr = OxmlElement('w:tcPr')
        cell._tc.insert(0, tc_pr)
    node = _child(tc_pr, 'vAlign')
    node.set(qn('w:val'), value)
    return True


def apply_run_rules(run, rules, keep_color=False):
    changed = False
    rpr = run.find(qn('w:rPr'))
    if rpr is None:
        rpr = OxmlElement('w:rPr')
        run.insert(0, rpr)

    def tag(name, attributes):
        node = rpr.find(qn('w:' + name))
        if node is None:
            node = OxmlElement('w:' + name)
            rpr.append(node)
        for key, value in attributes.items():
            node.set(qn('w:' + key), value)

    fonts = {}
    if rules.get('font_east'):
        fonts['eastAsia'] = rules['font_east']
    if rules.get('font_ascii'):
        fonts['ascii'] = rules['font_ascii']
        fonts['hAnsi'] = rules['font_ascii']
    if fonts:
        existing = rpr.find(qn('w:rFonts'))
        node = existing if existing is not None else OxmlElement('w:rFonts')
        for key, value in fonts.items():
            node.set(qn('w:' + key), value)
        if existing is None:
            rpr.insert(0, node)
        changed = True
    if rules.get('size_pt'):
        half = str(int(round(float(rules['size_pt']) * 2)))
        tag('sz', {'val': half})
        tag('szCs', {'val': half})
        changed = True
    if 'bold' in rules:
        value = 'true' if rules['bold'] else '0'
        tag('b', {'val': value})
        tag('bCs', {'val': value})
        changed = True
    if 'italic' in rules:
        value = 'true' if rules['italic'] else '0'
        tag('i', {'val': value})
        tag('iCs', {'val': value})
        changed = True
    # Guidance highlight/colour must not survive onto the filled value.
    for name in ('highlight', 'shd'):
        for node in list(rpr.findall(qn('w:' + name))):
            rpr.remove(node)
            changed = True
    wanted = str(rules.get('color') or '').strip().lstrip('#').upper()
    color = rpr.find(qn('w:color'))
    if wanted and wanted not in ('AUTO',):
        if color is None:
            color = OxmlElement('w:color')
            rpr.append(color)
        if color.get(qn('w:val')) != wanted:
            color.set(qn('w:val'), wanted)
            changed = True
    elif color is not None and not keep_color:
        current = (color.get(qn('w:val')) or '').upper()
        if current not in ('', '000000', 'AUTO'):
            color.set(qn('w:val'), '000000')
            changed = True
    return changed


def _paragraph_range(doc, rng):
    paragraphs = bookmarkfill._range_paragraphs(doc, rng)
    start = max(0, int(rng.get('start_p', 0)))
    end = max(start, int(rng.get('end_p', start)))
    return paragraphs[start:end + 1]


def _find_bookmark(doc, name):
    for _kind, _part, root in bookmarkfill.story_parts(doc):
        for element in root.iter(qn('w:bookmarkStart')):
            if element.get(qn('w:name')) == name:
                return element
    return None


def _bookmark_ends(doc, bid):
    for _kind, _part, root in bookmarkfill.story_parts(doc):
        for element in root.iter(qn('w:bookmarkEnd')):
            if element.get(qn('w:id')) == bid:
                return element
    return None


def _bookmark_span(doc, name):
    """Paragraphs and runs that sit *inside* one invisible bookmark.

    Runs outside the range (an inline ``姓名：`` label sharing the paragraph) must not be
    restyled, and paragraph rules are skipped when the field starts mid-paragraph.
    """
    start = _find_bookmark(doc, name)
    if start is None:
        return None
    end = _bookmark_ends(doc, start.get(qn('w:id')))
    p_start = bookmarkfill._paragraph_of(start)
    if p_start is None:
        return None
    p_end = bookmarkfill._paragraph_of(end) if end is not None else p_start
    root = p_start.getparent()
    paragraphs = [node for node in root.iter(qn('w:p'))]
    if p_start not in paragraphs:
        return None
    first = paragraphs.index(p_start)
    last = paragraphs.index(p_end) if p_end in paragraphs else first
    if last < first:
        first, last = last, first
    from docx.text.paragraph import Paragraph
    runs = []
    inline = False
    started = False
    finished = False
    for node in paragraphs[first:last + 1]:
        if finished:
            break
        for element in node.iter():
            if element is start:
                started = True
                continue
            if end is not None and element is end:
                finished = True
                break
            if element.tag != qn('w:r'):
                continue
            if not started:
                if ''.join(t.text or '' for t in element.findall(qn('w:t'))).strip():
                    inline = True
                continue
            runs.append(element)
    return {'paragraphs': [Paragraph(p, None) for p in paragraphs[first:last + 1]],
            'runs': runs, 'inline': inline}


def _apply_paragraph_rules(paragraph, rules):
    changed = False
    if rules.get('line') or rules.get('space_before') is not None \
            or rules.get('space_after') is not None:
        changed = set_spacing(paragraph, rules.get('line'),
                              rules.get('space_before'), rules.get('space_after')) or changed
    if rules.get('first_line_chars') or rules.get('hanging_chars'):
        changed = set_indent(paragraph, rules.get('first_line_chars'),
                             rules.get('hanging_chars')) or changed
    if rules.get('align'):
        changed = set_align(paragraph, rules['align']) or changed
    if any(rules.get(key) is not None for key in ('keep_next', 'keep_lines', 'widow_control')):
        changed = set_pagination(paragraph, rules.get('keep_next'), rules.get('keep_lines'),
                                 rules.get('widow_control')) or changed
    if rules.get('shading_fill') is not None:
        changed = set_shading(_pPr(paragraph), rules['shading_fill']) or changed
    if rules.get('mark'):
        changed = apply_mark_rules(paragraph, rules['mark']) or changed
    return changed


def _cell(doc, table, row, col):
    try:
        return doc.tables[int(table)].rows[int(row)].cells[int(col)]
    except (IndexError, KeyError, ValueError):
        return None


KEEP_TAGS = (qn('w:pPr'), qn('w:bookmarkStart'), qn('w:bookmarkEnd'),
             qn('w:commentRangeStart'), qn('w:commentRangeEnd'))


def _has_bookmark(paragraph):
    p = paragraph._p
    return any(p.find(tag) is not None for tag in
               (qn('w:bookmarkStart'), qn('w:bookmarkEnd')))


def _clear_paragraph(paragraph):
    """Drop visible content but keep paragraph properties and invisible markers."""
    for child in list(paragraph._p):
        if child.tag not in KEEP_TAGS:
            paragraph._p.remove(child)


def _remove_paragraph(paragraph):
    node = paragraph._p
    parent = node.getparent()
    if parent is None:
        return False
    if _has_bookmark(paragraph):
        _clear_paragraph(paragraph)
        return True
    if parent.tag == qn('w:tc') and len([c for c in parent if c.tag == qn('w:p')]) <= 1:
        _clear_paragraph(paragraph)
        return True
    parent.remove(node)
    return True


def _trimmable(run):
    return ([c for c in run._element if c.tag not in (qn('w:rPr'), qn('w:t'))] == []
            and len(run._element.findall(qn('w:t'))) == 1)


def _trim_paragraph(paragraph, start, end):
    """Delete characters [start:end) from a paragraph without touching the rest."""
    text = ''.join(run.text or '' for run in paragraph.runs)
    if start >= end or start >= len(text):
        return False
    end = min(end, len(text))
    cursor = 0
    touched = False
    for run in list(paragraph.runs):
        value = run.text or ''
        run_start, run_end = cursor, cursor + len(value)
        cursor = run_end
        if run_end <= start or run_start >= end:
            continue
        if not _trimmable(run):
            return False
        head = value[:max(0, start - run_start)]
        tail = value[max(0, min(len(value), end - run_start)):]
        node = run._element.find(qn('w:t'))
        node.text = head
        node.set(qn('xml:space'), 'preserve')
        if tail:
            clone = OxmlElement('w:r')
            rpr = run._element.find(qn('w:rPr'))
            if rpr is not None:
                import copy
                clone.append(copy.deepcopy(rpr))
            extra = OxmlElement('w:t')
            extra.set(qn('xml:space'), 'preserve')
            extra.text = tail
            clone.append(extra)
            run._element.addnext(clone)
        touched = True
    return touched


def _note_paragraph(doc, note):
    container = note.get('container')
    if container == 'cell':
        cell = _cell(doc, note.get('table'), note.get('row'), note.get('col'))
        if cell is None:
            return None
        paragraphs = list(cell.paragraphs)
    elif container in ('header', 'footer'):
        root = None
        for kind, name, part in bookmarkfill.story_parts(doc):
            if name == note.get('part'):
                root = part
                break
        if root is None:
            return None
        paragraphs = bookmarkfill._paragraphs_of_root(root)
    elif container == 'body':
        paragraphs = list(doc.paragraphs)
    else:
        return None
    index = int(note.get('p', -1))
    if not 0 <= index < len(paragraphs):
        return None
    return paragraphs[index]


def _verify_note(paragraph, note):
    """Only touch a paragraph whose text still matches what recognition saw."""
    text = ''.join(run.text or '' for run in paragraph.runs)
    quote = (note.get('text') or '').strip()
    if not quote:
        return False
    if note.get('whole'):
        return quote in re.sub(r'\s+', '', text) or quote in text
    start = int(note.get('start', 0))
    return text[start:start + len(quote)] == quote if start < len(text) else quote in text


RUN_ONLY = ('size_pt', 'font_east', 'font_ascii', 'bold', 'italic')


def _field_paragraphs(doc, rng):
    paragraphs = bookmarkfill._range_paragraphs(doc, rng)
    start = max(0, int(rng.get('start_p', 0)))
    end = max(start, int(rng.get('end_p', start)))
    return paragraphs[start:end + 1]


def _box_text(box):
    lines = []
    for p in box.iter(qn('w:p')):
        lines.append(''.join(t.text or '' for t in p.iter(qn('w:t'))))
    return '\n'.join(lines)


def _compact(value):
    return re.sub(r'\s+', '', value or '')


SHAPE_TAGS = {qn('w:pict'), qn('w:drawing'),
              '{urn:schemas-microsoft-com:vml}shape',
              '{urn:schemas-microsoft-com:vml}rect',
              '{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}wsp'}


def _box_shape(box):
    node = box.getparent()
    picked = None
    while node is not None and node.tag != qn('w:p'):
        if node.tag in SHAPE_TAGS and picked is None:
            picked = node
        node = node.getparent()
    return picked


def _remove_textbox(doc, note):
    """Delete every callout box whose text was recorded as one requirement.

    Identical callouts are stored twice (``mc:Choice`` plus ``mc:Fallback``) and may
    repeat on the page, so all matches go, not just the first.
    """
    target = _compact(note.get('textbox'))
    if not target:
        return 0
    removed = 0
    for _kind, _part, root in bookmarkfill.story_parts(doc):
        for box in list(root.iter(qn('w:txbxContent'))):
            if _compact(_box_text(box)) != target:
                continue
            shape = _box_shape(box)
            if shape is None or shape.getparent() is None:
                continue
            shape.getparent().remove(shape)
            removed += 1
    return removed


def apply(doc, spec, options=None, phase='all'):
    """Apply the template's own requirements.

    ``phase='pre'`` runs before the bookmark replacement: paragraph rules and note
    removal, while the paragraph indices recorded at recognition time are still true.
    ``phase='post'`` runs afterwards and only restyles the runs inside each bookmark,
    which is index-free and therefore safe after paragraphs were deleted.
    """
    options = options or {}
    do_format = options.get('apply_format', True)
    do_remove = options.get('remove_notes', True)
    report = {'runs': 0, 'paragraphs': 0, 'removed': 0, 'kept': 0,
              'blank_lines': 0, 'skipped': []}
    if not spec:
        return report
    fields = {f.get('key'): f for f in (spec.get('fields') or [])}
    ranges = spec.get('field_ranges') or {}
    if do_format and phase in ('all', 'post'):
        for key, rules in (spec.get('field_rules') or {}).items():
            run_rules = {k: v for k, v in rules.items() if k in RUN_ONLY}
            field = fields.get(key)
            if not run_rules or not field or not field.get('bookmark'):
                continue
            span = _bookmark_span(doc, field['bookmark'])
            if span is None:
                report['skipped'].append(field.get('label') or key)
                continue
            for run in span['runs']:
                if apply_run_rules(run, run_rules):
                    report['runs'] += 1
    if do_format and phase in ('all', 'pre'):
        for key, rules in (spec.get('field_rules') or {}).items():
            rng = ranges.get(key)
            if not rng:
                continue
            try:
                paragraphs = _field_paragraphs(doc, rng)
            except (IndexError, KeyError, ValueError):
                report['skipped'].append(key)
                continue
            for paragraph in paragraphs:
                if _apply_paragraph_rules(paragraph, rules):
                    report['paragraphs'] += 1
        for entry in spec.get('container_rules') or []:
            rules = entry.get('rules') or {}
            if entry.get('container') == 'cell':
                cell = _cell(doc, entry.get('table'), entry.get('row'), entry.get('col'))
                if cell is None:
                    continue
                if rules.get('valign'):
                    set_cell_valign(cell, rules['valign'])
                exclude = {int(p) for p in entry.get('exclude_p') or []}
                paragraphs = [p for i, p in enumerate(cell.paragraphs) if i not in exclude]
            else:
                paragraph = _note_paragraph(doc, entry)
                paragraphs = [paragraph] if paragraph is not None else []
                if entry.get('scope') == 'following' and paragraph is not None:
                    body = list(doc.paragraphs)
                    if paragraph in body:
                        tail = body[body.index(paragraph) + 1:body.index(paragraph) + 7]
                        paragraphs = [p for p in tail if ''.join(
                            r.text or '' for r in p.runs).strip()][:6]
            for paragraph in paragraphs:
                if _apply_paragraph_rules(paragraph, rules):
                    report['paragraphs'] += 1
    if do_remove and phase in ('all', 'pre'):
        notes = [n for n in (spec.get('notes') or []) if n.get('container') != 'vision']
        # Callout boxes go first: deleting their host paragraph would hide them from us.
        for note in [n for n in notes if n.get('textbox')]:
            if _remove_textbox(doc, note):
                report['removed'] += 1
            else:
                report['kept'] += 1
        for note in sorted((n for n in notes if not n.get('textbox')),
                           key=lambda n: (n.get('container'), n.get('table') if
                                          n.get('table') is not None else -1,
                                          n.get('row') or 0, n.get('col') or 0,
                                          -(n.get('p') or 0))):
            paragraph = _note_paragraph(doc, note)
            if paragraph is None or not _verify_note(paragraph, note):
                report['kept'] += 1
                continue
            if note.get('plain'):
                # 黑字里的格式要求只套格式，原文是模板自己的正文，不能删
                report['kept'] += 1
                continue
            if note.get('whole'):
                if (note.get('rules') or {}).get('blank_lines'):
                    _clear_paragraph(paragraph)
                    report['blank_lines'] += 1
                else:
                    _remove_paragraph(paragraph)
                report['removed'] += 1
            elif _trim_paragraph(paragraph, int(note.get('start', 0)),
                                 int(note.get('end', 0))):
                report['removed'] += 1
            else:
                report['kept'] += 1
    return report


def check_content(spec, values, labels=None):
    """Warn-only checks for 不少于100字 / 参考文献不少于10篇 / 分数不能低于70分."""
    warnings = []
    labels = labels or {}
    for rule in spec.get('content_rules') or []:
        key = rule.get('field_key') or ''
        value = str((values or {}).get(key, '') or '').strip()
        label = labels.get(key) or rule.get('label') or '未指定栏目'
        content = rule.get('content') or {}
        if not value:
            continue
        count = len(re.sub(r'\s+', '', value))
        if content.get('min_chars') and count < content['min_chars']:
            warnings.append(f'「{label}」{count} 字，要求不少于 {content["min_chars"]} 字')
        if content.get('max_chars') and count > content['max_chars']:
            warnings.append(f'「{label}」{count} 字，要求不超过 {content["max_chars"]} 字')
        if content.get('min_refs'):
            refs = len([p for p in re.split(r'[\n;；]+', value) if p.strip()])
            if refs < content['min_refs']:
                warnings.append(f'「{label}」当前 {refs} 条，要求不少于 {content["min_refs"]} 条')
        if content.get('score_min'):
            numbers = [int(n) for n in re.findall(r'\d+', value)]
            if numbers and max(numbers) < content['score_min']:
                warnings.append(f'「{label}」分数 {max(numbers)}，要求不低于 {content["score_min"]} 分')
        if content.get('max_years'):
            from datetime import datetime
            years = [int(y) for y in re.findall(r'(?:19|20)\d{2}', value)]
            limit = datetime.now().year - int(content['max_years'])
            old = sorted(y for y in years if y < limit)
            if old:
                warnings.append(f'「{label}」有 {len(old)} 处早于 {limit} 年'
                                f'（最早 {old[0]} 年），要求 {content["max_years"]} 年以内')
    return warnings
