"""Invisible Word bookmarks for fillable ranges.

The template keeps every original paragraph, sample answer, color, highlight,
textbox and page break.  Recognition only adds ``w:bookmarkStart/End`` around
the sample range that a field should replace.  Generation removes the bookmark
range and writes the user's value, so the template looks like the original and
the filled document is clean.
"""
from __future__ import annotations

import copy
import re
from pathlib import Path

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _text(p):
    return ''.join(r.text for r in p.runs)


def story_parts(doc):
    """Yield (kind, name, root) for the body and every existing header/footer."""
    yield 'body', 'body', doc.element.body
    seen = set()
    for rel in doc.part.rels.values():
        if rel.reltype not in (RT.HEADER, RT.FOOTER):
            continue
        part = rel.target_part
        name = str(part.partname)
        if name in seen:
            continue
        seen.add(name)
        kind = 'header' if rel.reltype == RT.HEADER else 'footer'
        yield kind, name, part.element


def _paragraphs_of_root(root):
    return [Paragraph(p, None) for p in root.iter(qn('w:p'))]


def _part_root(doc, name):
    for _kind, part_name, root in story_parts(doc):
        if part_name == name:
            return root
    return None


def _range_paragraphs(doc, rng):
    container = (rng or {}).get('container', 'cell')
    if container == 'body':
        return list(doc.paragraphs)
    if container in ('header', 'footer'):
        root = _part_root(doc, rng.get('part'))
        if root is None:
            raise ValueError('找不到页眉/页脚部件：%s' % rng.get('part'))
        return _paragraphs_of_root(root)
    ti, ri, ci = int(rng['table']), int(rng['row']), int(rng['col'])
    return doc.tables[ti].rows[ri].cells[ci].paragraphs


def _paragraph_of(el):
    node = el
    while node is not None:
        if node.tag == qn('w:p'):
            return node
        node = node.getparent()
    return None


def _next_id(doc):
    ids = []
    for _kind, _name, root in story_parts(doc):
        for el in root.iter(qn('w:bookmarkStart')):
            try:
                ids.append(int(el.get(qn('w:id')) or 0))
            except ValueError:
                pass
    return max(ids, default=0) + 1


def _split_run(p, offset):
    """Split at a character offset; return the run that starts at the offset."""
    text = _text(p)
    if offset <= 0:
        return p.runs[0]._element if p.runs else None
    if offset >= len(text):
        return None
    pos = 0
    for run in list(p.runs):
        value = run.text
        end = pos + len(value)
        if pos < offset < end:
            local = offset - pos
            rpr = run._element.find(qn('w:rPr'))
            suffix = p.add_run(value[local:])
            if rpr is not None:
                suffix._element.insert(0, copy.deepcopy(rpr))
            run.text = value[:local]
            run._element.addnext(suffix._element)
            return suffix._element
        if offset == pos:
            return run._element
        pos = end
    return None


def _insert_at(p, offset, element):
    text = _text(p)
    if offset <= 0:
        index = 0
        children = list(p._p)
        while index < len(children) and children[index].tag == qn('w:pPr'):
            index += 1
        p._p.insert(index, element)
        return
    if offset >= len(text):
        p._p.append(element)
        return
    suffix = _split_run(p, offset)
    if suffix is not None:
        suffix.addprevious(element)
    else:
        p._p.append(element)


def add_bookmark(doc, paragraphs, start_p, start_off, end_p, end_off, name):
    start_p = max(0, min(int(start_p), len(paragraphs) - 1))
    end_p = max(0, min(int(end_p), len(paragraphs) - 1))
    sp = paragraphs[start_p]
    ep = paragraphs[end_p]
    start_off = max(0, min(int(start_off), len(_text(sp))))
    end_off = max(0, min(int(end_off), len(_text(ep))))
    if start_p > end_p or (start_p == end_p and start_off > end_off):
        start_p, end_p = end_p, start_p
        start_off, end_off = end_off, start_off
        sp, ep = paragraphs[start_p], paragraphs[end_p]
    bid = _next_id(doc)
    start_el = OxmlElement('w:bookmarkStart')
    start_el.set(qn('w:id'), str(bid))
    start_el.set(qn('w:name'), name)
    end_el = OxmlElement('w:bookmarkEnd')
    end_el.set(qn('w:id'), str(bid))
    if start_p == end_p and start_off == end_off:
        _insert_at(sp, start_off, start_el)
        start_el.addnext(end_el)
    else:
        _insert_at(ep, end_off, end_el)
        _insert_at(sp, start_off, start_el)
    return bid


def build(source, out, plan, selected):
    """Add invisible bookmarks for selected fields; keep all original content."""
    doc = Document(str(source))
    mapping = {}
    for field in plan.get('fields', []):
        pid = field.get('plan_id')
        if pid not in selected:
            continue
        rng = field.get('bookmark_range')
        if not rng:
            continue
        paras = _range_paragraphs(doc, rng)
        name = field.get('bookmark') or ('WTT_' + re.sub(r'[^0-9A-Za-z_]', '_', str(pid)))
        add_bookmark(doc, paras, rng.get('start_p', 0), rng.get('start', 0),
                     rng.get('end_p', rng.get('start_p', 0)), rng.get('end', rng.get('start', 0)),
                     name)
        mapping[pid] = name
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return {'bookmarks': sorted(mapping.values()), 'mapping': mapping}


def scan_bookmarks(path):
    doc = Document(str(path))
    names = []
    for _kind, _name, root in story_parts(doc):
        for el in root.iter(qn('w:bookmarkStart')):
            name = el.get(qn('w:name')) or ''
            if name:
                names.append(name)
    return names


def _clean_rpr(rpr):
    if rpr is None:
        return None
    rpr = copy.deepcopy(rpr)
    for tag in ('color', 'highlight', 'shd'):
        for el in list(rpr.findall(qn('w:' + tag))):
            rpr.remove(el)
    return rpr


def _value_rpr(doc, start, end, start_p, end_p):
    # Prefer the first sample run between the bookmark boundaries, but drop
    # the blue/red example color and highlighting.
    if start_p is end_p:
        node = start.getnext()
        while node is not None and node is not end:
            if node.tag == qn('w:r'):
                rpr = node.find(qn('w:rPr'))
                if rpr is not None and ''.join(
                        t.text or '' for t in node.iter(qn('w:t'))).strip():
                    return _clean_rpr(rpr)
            node = node.getnext()
    ppr = start_p.find(qn('w:pPr'))
    if ppr is not None:
        rpr = ppr.find(qn('w:rPr'))
        if rpr is not None:
            return _clean_rpr(rpr)
    ppr = end_p.find(qn('w:pPr'))
    if ppr is not None:
        rpr = ppr.find(qn('w:rPr'))
        if rpr is not None:
            return _clean_rpr(rpr)
    return None


def _remove_numbering(p):
    ppr = p.find(qn('w:pPr'))
    if ppr is None:
        return
    for num in list(ppr.findall(qn('w:numPr'))):
        ppr.remove(num)


def _has_text_before(p, bookmark):
    node = p.find(qn('w:pPr'))
    while node is not None and node is not bookmark:
        if node.tag == qn('w:r'):
            text = ''.join(t.text or '' for t in node.iter(qn('w:t')))
            if text.strip():
                return True
        node = node.getnext()
    return False


def _make_run(doc, rpr, value):
    run = OxmlElement('w:r')
    if rpr is not None:
        run.append(rpr)
    parts = re.split(r'([\n\t])', value)
    for part in parts:
        if part == '\n':
            run.append(OxmlElement('w:br'))
        elif part == '\t':
            run.append(OxmlElement('w:tab'))
        elif part:
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = part
            run.append(t)
    return run


def replace_bookmark(doc, name, value, keep_empty=False):
    start = None
    for _kind, _part, root in story_parts(doc):
        for el in root.iter(qn('w:bookmarkStart')):
            if el.get(qn('w:name')) == name:
                start = el
                break
        if start is not None:
            break
    if start is None:
        return False
    bid = start.get(qn('w:id'))
    end = None
    for _kind, _part, root in story_parts(doc):
        for el in root.iter(qn('w:bookmarkEnd')):
            if el.get(qn('w:id')) == bid:
                end = el
                break
        if end is not None:
            break
    if end is None:
        return False
    value = str(value or '').replace('\r\n', '\n').replace('\r', '\n')
    if not value and keep_empty:
        return True
    start_p = _paragraph_of(start)
    end_p = _paragraph_of(end)
    if start_p is None or end_p is None:
        return False
    rpr = _value_rpr(doc, start, end, start_p, end_p)
    if start_p is end_p:
        node = start.getnext()
        while node is not None and node is not end:
            nxt = node.getnext()
            start_p.remove(node)
            node = nxt
    else:
        node = start.getnext()
        while node is not None:
            nxt = node.getnext()
            start_p.remove(node)
            node = nxt
        node = end.getprevious()
        while node is not None:
            nxt = node.getprevious()
            if node.tag != qn('w:pPr'):
                end_p.remove(node)
            node = nxt
        parent = start_p.getparent()
        if parent is not None and parent is end_p.getparent():
            node = start_p.getnext()
            while node is not None and node is not end_p:
                nxt = node.getnext()
                if node.tag == qn('w:p'):
                    parent.remove(node)
                node = nxt
    if not _has_text_before(start_p, start):
        _remove_numbering(start_p)
    if end_p is not start_p:
        _remove_numbering(end_p)
    run = _make_run(doc, rpr, value)
    start.addnext(run)
    return True
