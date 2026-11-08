"""Fill Word content controls and legacy FORMText fields in place.

The template is never rewritten: a control field stores the part name and the
control index, so generation can write the user value into the existing
``w:sdtContent`` (or the result range of a legacy form field).
"""
from __future__ import annotations

import copy
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import bookmarkfill

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PLACEHOLDER = re.compile(
    r'单击此处|点击此处|请输入|请填写|在此输入|此处输入|填写|输入文字|选择日期|'
    r'placeholder|click here|type here|select a date', re.I)


def _text(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def _paragraph(el):
    node = el
    while node is not None:
        if node.tag == qn('w:p'):
            return node
        node = node.getparent()
    return None


def _run(el):
    node = el
    while node is not None:
        if node.tag == qn('w:r'):
            return node
        node = node.getparent()
    return None


def _is_inside(node, parent):
    cur = node.getparent()
    while cur is not None:
        if cur is parent:
            return True
        cur = cur.getparent()
    return False


def _text_before(el):
    p = _paragraph(el)
    if p is None:
        return ''
    anchor = _run(el)
    pieces = []
    for node in p.iter():
        if anchor is not None and node is anchor:
            break
        if node.tag == qn('w:t'):
            pieces.append(node.text or '')
    return ''.join(pieces).strip()


def _prev_paragraph_text(p):
    prev = p.getprevious() if p is not None else None
    while prev is not None:
        if prev.tag == qn('w:p'):
            text = _text(prev).strip()
            if text:
                return text
        prev = prev.getprevious()
    return ''


def _paragraph_index(root, el):
    p = _paragraph(el)
    if p is None:
        return 0
    for i, node in enumerate(root.iter(qn('w:p'))):
        if node is p:
            return i
    return 0


def _table_position(doc, el):
    tbl = el
    while tbl is not None and tbl.tag != qn('w:tbl'):
        tbl = tbl.getparent()
    if tbl is None:
        return None
    ti = None
    for i, table in enumerate(doc.tables):
        if table._tbl is tbl:
            ti = i
            break
    if ti is None:
        return None
    rows = [r for r in tbl if r.tag == qn('w:tr')]
    tr = el
    while tr is not None and tr.tag != qn('w:tr'):
        tr = tr.getparent()
    if tr is None or tr not in rows:
        return ti, 0, 0
    ri = rows.index(tr)
    cells = [c for c in tr if c.tag == qn('w:tc')]
    tc = el
    while tc is not None and tc.tag != qn('w:tc'):
        tc = tc.getparent()
    ci = cells.index(tc) if tc in cells else 0
    return ti, ri, ci


def _outer_sdts(root):
    for el in root.iter(qn('w:sdt')):
        parent = el.getparent()
        nested = False
        while parent is not None:
            if parent.tag == qn('w:sdt'):
                nested = True
                break
            parent = parent.getparent()
        if not nested:
            yield el


def _sdt_label(doc, sdt):
    pr = sdt.find(qn('w:sdtPr'))
    alias = tag = ''
    if pr is not None:
        node = pr.find(qn('w:alias'))
        if node is not None:
            alias = (node.get(qn('w:val')) or '').strip()
        node = pr.find(qn('w:tag'))
        if node is not None:
            tag = (node.get(qn('w:val')) or '').strip()
    for value in (alias, tag):
        if value and not PLACEHOLDER.search(value):
            return value
    before = re.sub(r'[\s：:]+$', '', _text_before(sdt)).strip()
    if before:
        return before[-24:]
    prev = re.sub(r'[\s：:]+$', '', _prev_paragraph_text(_paragraph(sdt))).strip()
    return prev[-24:] if prev else (alias or tag)


def _sdt_multiline(sdt):
    content = sdt.find(qn('w:sdtContent'))
    if content is None:
        return False
    return len(list(content.iter(qn('w:p')))) > 1 or len(_text(content)) > 40


def _sdt_choices(sdt):
    choices = []
    pr = sdt.find(qn('w:sdtPr'))
    if pr is None:
        return choices
    for name in ('w:dropDownList', 'w:comboBox'):
        holder = pr.find(qn(name))
        if holder is None:
            continue
        for item in holder.findall(qn('w:listItem')):
            label = (item.get(qn('w:displayText')) or item.get(qn('w:value')) or '').strip()
            if label:
                choices.append(label)
    return choices


def _is_checkbox(sdt):
    pr = sdt.find(qn('w:sdtPr'))
    if pr is None:
        return False
    return any(el.tag.endswith('}checkbox') for el in pr.iter())


def list_controls(doc):
    controls = []
    for kind, name, root in bookmarkfill.story_parts(doc):
        for index, sdt in enumerate(_outer_sdts(root)):
            if _is_checkbox(sdt):
                continue
            content = sdt.find(qn('w:sdtContent'))
            text = _text(content).strip() if content is not None else ''
            showing = False
            pr = sdt.find(qn('w:sdtPr'))
            if pr is not None:
                node = pr.find(qn('w:showingPlcHdr'))
                showing = node is not None and (node.get(qn('w:val')) or '1') not in ('0', 'false')
            label = _sdt_label(doc, sdt)
            pos = _table_position(doc, sdt)
            controls.append({
                'kind': 'control', 'part': kind, 'part_name': name, 'index': index,
                'alias': label, 'text': text, 'multiline': _sdt_multiline(sdt),
                'choices': _sdt_choices(sdt), 'placeholder': bool(showing) or bool(PLACEHOLDER.search(text)),
                'table': pos[0] if pos else None, 'row': pos[1] if pos else None,
                'col': pos[2] if pos else None, 'p': _paragraph_index(root, sdt),
            })
    return controls


def _find_control(doc, part_name, index):
    for _kind, name, root in bookmarkfill.story_parts(doc):
        if name != part_name:
            continue
        for i, sdt in enumerate(_outer_sdts(root)):
            if i == int(index):
                return sdt
    return None


def _first_rpr(el):
    for run in el.iter(qn('w:r')):
        rpr = run.find(qn('w:rPr'))
        if rpr is not None:
            return copy.deepcopy(rpr)
    return None


def _make_runs(value, rpr):
    nodes = []
    lines = str(value or '').replace('\r\n', '\n').replace('\r', '\n').split('\n')
    for line_i, line in enumerate(lines):
        if line_i:
            br = OxmlElement('w:r')
            if rpr is not None:
                br.append(copy.deepcopy(rpr))
            br.append(OxmlElement('w:br'))
            nodes.append(br)
        run = OxmlElement('w:r')
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = line
        run.append(t)
        nodes.append(run)
    return nodes


def _clear_paragraph(p):
    for child in list(p):
        if child.tag != qn('w:pPr'):
            p.remove(child)


def _set_sdt_text(sdt, value):
    content = sdt.find(qn('w:sdtContent'))
    if content is None:
        content = OxmlElement('w:sdtContent')
        sdt.append(content)
    rpr = _first_rpr(content)
    paras = [p for p in content if p.tag == qn('w:p')]
    if paras:
        first = paras[0]
        _clear_paragraph(first)
        for run in _make_runs(value, rpr):
            first.append(run)
        for p in paras[1:]:
            content.remove(p)
        for child in list(content):
            if child.tag != qn('w:p'):
                content.remove(child)
    else:
        for child in list(content):
            content.remove(child)
        for run in _make_runs(value, rpr):
            content.append(run)
    pr = sdt.find(qn('w:sdtPr'))
    if pr is not None:
        node = pr.find(qn('w:showingPlcHdr'))
        if node is not None:
            node.set(qn('w:val'), '0')


def replace_control(doc, part_name, index, value):
    sdt = _find_control(doc, part_name, index)
    if sdt is None:
        return False
    _set_sdt_text(sdt, value)
    return True


def _field_parts(begin):
    p = _paragraph(begin)
    run = _run(begin)
    if p is None or run is None:
        return None
    runs = [r for r in p.iter(qn('w:r'))]
    try:
        i0 = runs.index(run)
    except ValueError:
        return None
    separate = end = None
    for j in range(i0 + 1, len(runs)):
        chars = runs[j].findall(qn('w:fldChar'))
        for ch in chars:
            kind = ch.get(qn('w:fldCharType'))
            if kind == 'separate' and separate is None:
                separate = j
            elif kind == 'end':
                end = j
                break
        if end is not None:
            break
    if end is None:
        return None
    instruction = ''.join(
        (node.text or '') for r in runs[i0 + 1:separate if separate is not None else end]
        for node in r.findall(qn('w:instrText')))
    result = runs[(separate + 1 if separate is not None else i0 + 1):end]
    return p, runs, instruction.strip(), result, runs[end]


def list_formfields(doc):
    fields = []
    for kind, name, root in bookmarkfill.story_parts(doc):
        index = 0
        for begin in root.iter(qn('w:fldChar')):
            if begin.get(qn('w:fldCharType')) != 'begin':
                continue
            info = _field_parts(begin)
            if not info:
                continue
            p, runs, instruction, result, end_run = info
            if not re.match(r'FORMTEXT|FORMDROPDOWN', instruction, re.I):
                continue
            text = ''.join(_text(r) for r in result).strip()
            label = re.sub(r'[\s：:]+$', '', _text_before(begin)).strip()
            pos = _table_position(doc, begin)
            fields.append({
                'kind': 'formfield', 'part': kind, 'part_name': name, 'index': index,
                'alias': label[-24:], 'text': text, 'multiline': False,
                'choices': [], 'placeholder': not text,
                'table': pos[0] if pos else None, 'row': pos[1] if pos else None,
                'col': pos[2] if pos else None, 'p': _paragraph_index(root, begin),
            })
            index += 1
    return fields


def replace_formfield(doc, part_name, index, value):
    for _kind, name, root in bookmarkfill.story_parts(doc):
        if name != part_name:
            continue
        current = 0
        for begin in root.iter(qn('w:fldChar')):
            if begin.get(qn('w:fldCharType')) != 'begin':
                continue
            info = _field_parts(begin)
            if not info:
                continue
            p, runs, instruction, result, end_run = info
            if not re.match(r'FORMTEXT|FORMDROPDOWN', instruction, re.I):
                continue
            if current != int(index):
                current += 1
                continue
            rpr = _first_rpr(result[0]) if result else None
            for run in result:
                p.remove(run)
            for run in _make_runs(value, rpr):
                end_run.addprevious(run)
            return True
    return False


def fill_controls(doc, fields, values):
    done = set()
    for field in fields:
        key = field.get('key')
        if not key or key in done:
            continue
        value = values.get(key, '')
        ctrl = field.get('control') or {}
        if not ctrl:
            continue
        if ctrl.get('kind') == 'formfield':
            replace_formfield(doc, ctrl.get('part_name'), ctrl.get('index'), values.get(key, ''))
        else:
            replace_control(doc, ctrl.get('part_name'), ctrl.get('index'), values.get(key, ''))
        done.add(key)
    return done
