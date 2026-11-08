# -*- coding: utf-8 -*-
"""默认清洁规则：任何模板在导入、生成、排版时都删掉文本框与突出显示。

- 文本框：DrawingML（wps:txbx）与旧式 VML（v:textbox）都删；新版 Word 的文本框
  外层是 mc:AlternateContent（Choice 是 DrawingML、Fallback 是 VML 副本），
  删整个 AlternateContent 才能连副本一起清掉。
- 突出显示：删除所有 w:highlight，文字、字体、字号、底纹与版式不动。
"""
from __future__ import annotations

from docx.oxml.ns import qn

TXBX = qn('w:txbxContent')
DRAWING = qn('w:drawing')
PICT = qn('w:pict')
ALT = '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'


def _story_roots(doc):
    roots = [doc.element]
    try:
        roots.append(doc.styles.element)
    except Exception:
        pass
    for section in doc.sections:
        for name in ('header', 'footer', 'first_page_header', 'first_page_footer',
                     'even_page_header', 'even_page_footer'):
            try:
                element = getattr(getattr(section, name), '_element', None)
            except Exception:
                continue
            if element is not None:
                roots.append(element)
    return roots


def remove_textboxes(doc):
    """删除全文所有文本框（新旧两种），返回删除个数。"""
    removed = 0
    for root in _story_roots(doc):
        for node in list(root.iter(TXBX)):
            target = None
            ancestor = node.getparent()
            while ancestor is not None:
                if ancestor.tag == ALT:
                    target = ancestor
                    break
                if ancestor.tag in (DRAWING, PICT):
                    target = ancestor
                ancestor = ancestor.getparent()
            if target is None or target.getparent() is None:
                continue
            target.getparent().remove(target)
            removed += 1
    return removed


def clear_highlights(doc):
    """删除全文所有突出显示（w:highlight），返回删除个数。"""
    removed = 0
    for root in _story_roots(doc):
        for node in list(root.iter(qn('w:highlight'))):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
                removed += 1
    return removed


def clean(doc):
    """先删文本框（框里也可能有高亮），再删突出显示。"""
    boxes = remove_textboxes(doc)
    marks = clear_highlights(doc)
    return {'textboxes': boxes, 'highlights': marks}
