"""Fill literal Word placeholders without rebuilding the document package."""
from __future__ import annotations

import json
import os
import re
import tempfile
import unicodedata
from pathlib import Path
from xml.dom import Node, minidom
from zipfile import BadZipFile, ZipFile

from docx import Document

import bookmarkfill
import controlfill
import ooxmlcheck

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
V = 'urn:schemas-microsoft-com:vml'
TOKEN = re.compile(r'\{\{\s*([^{}\r\n]{1,100}?)\s*\}\}', re.UNICODE)
INVALID_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]')
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
IMAGE_SIGNS = {'.png': (b'\x89PNG',), '.jpg': (b'\xff\xd8\xff',), '.jpeg': (b'\xff\xd8\xff',),
               '.gif': (b'GIF8',), '.bmp': (b'BM',)}


def _normalize_part(directory, target):
    parts = (directory + '/' + target).split('/') if directory else target.split('/')
    stack = []
    for part in parts:
        if part == '..':
            if stack:
                stack.pop()
        elif part not in ('', '.'):
            stack.append(part)
    return '/'.join(stack)


def _rels_map(package, part):
    directory, _, filename = part.rpartition('/')
    rels_name = f'{directory}/_rels/{filename}.rels'
    if rels_name not in package.namelist():
        return {}
    doc = parse_xml(package.read(rels_name))
    mapping = {}
    for rel in doc.getElementsByTagName('Relationship'):
        if rel.getAttribute('TargetMode') == 'External':
            continue
        target = rel.getAttribute('Target')
        mapping[rel.getAttribute('Id')] = target[1:] if target.startswith('/') \
            else _normalize_part(directory, target)
    doc.unlink()
    return mapping


def scan_image_anchors(path):
    """扫描带 {{标记}} 替代文字(alt text)的图片，返回锚点清单：tag/part/relationship/media/format。"""
    anchors = []
    with ZipFile(path) as package:
        for name in word_parts(package):
            doc = parse_xml(package.read(name))
            rels = _rels_map(package, name)
            for node in doc.getElementsByTagNameNS(WP, 'docPr'):
                text = (node.getAttribute('descr') or '') + ' ' + (node.getAttribute('name') or '')
                for match in TOKEN.finditer(text):
                    holder = node.parentNode
                    while holder is not None and holder.localName not in ('inline', 'anchor'):
                        holder = holder.parentNode
                    if holder is None:
                        continue
                    blips = holder.getElementsByTagNameNS(A, 'blip')
                    if not blips:
                        continue
                    rid = blips[0].getAttributeNS(R, 'embed')
                    if not rid or not rels.get(rid):
                        continue
                    anchors.append({'tag': match.group(1).strip(), 'part': name,
                                    'relationship': rid, 'media': rels[rid],
                                    'fmt': Path(rels[rid]).suffix.lower()})
            doc.unlink()
    return anchors


def content_fingerprint(path):
    """文字内容指纹：拼接正文/页眉页脚全部可见文字做 SHA256，用于证明“内容未变”。"""
    import hashlib
    digest = hashlib.sha256()
    with ZipFile(path) as package:
        for name in sorted(word_parts(package)):
            doc = parse_xml(package.read(name))
            for t in doc.getElementsByTagNameNS(W, 't'):
                digest.update(node_text(t).encode('utf-8'))
            doc.unlink()
    return digest.hexdigest()


def word_parts(package):
    return [n for n in package.namelist() if re.fullmatch(
        r'word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml', n)]


def parse_xml(data):
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise ValueError('模板包含不支持的 XML 声明，请使用正常的 .docx 文件。')
    return minidom.parseString(data)


def _remove_textboxes(doc):
    """删除该 part 下所有文本框（DrawingML 与 VML），含 AlternateContent 里的兼容副本。"""
    removed = 0
    for node in list(doc.getElementsByTagNameNS(W, 'txbxContent')):
        target = None
        ancestor = node.parentNode
        while ancestor is not None:
            if ancestor.namespaceURI == MC and ancestor.localName == 'AlternateContent':
                target = ancestor
                break
            if (ancestor.namespaceURI == W and ancestor.localName in ('drawing', 'pict')):
                target = ancestor
            ancestor = ancestor.parentNode
        if target is None or target.parentNode is None:
            continue
        target.parentNode.removeChild(target)
        removed += 1
    return removed


def _clear_highlights(doc):
    """删除该 part 下所有突出显示（w:highlight）。"""
    removed = 0
    for node in list(doc.getElementsByTagNameNS(W, 'highlight')):
        parent = node.parentNode
        if parent is not None:
            parent.removeChild(node)
            removed += 1
    return removed


def paragraph_text_nodes(paragraph):
    nodes = []
    for node in paragraph.getElementsByTagNameNS(W, 't'):
        ancestor = node.parentNode
        while ancestor is not None and ancestor is not paragraph:
            if ancestor.namespaceURI == W and ancestor.localName == 'p':
                break
            ancestor = ancestor.parentNode
        if ancestor is paragraph:
            nodes.append(node)
    return nodes


def node_text(node):
    return ''.join(n.data for n in node.childNodes if n.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE))


def set_text(node, value):
    while node.firstChild:
        node.removeChild(node.firstChild)
    node.setAttribute('xml:space', 'preserve')
    node.appendChild(node.ownerDocument.createTextNode(value))


def scan_fields(path):
    found = []
    try:
        with ZipFile(path) as package:
            if 'word/document.xml' not in package.namelist():
                raise ValueError('这不是有效的 Word .docx 文档。')
            if any('vbaProject' in n for n in package.namelist()):
                raise ValueError('请使用不含宏的 .docx 模板。')
            for name in word_parts(package):
                if package.getinfo(name).file_size > 20_000_000:
                    raise ValueError('模板正文过大，无法导入。')
                doc = parse_xml(package.read(name))
                for p in doc.getElementsByTagNameNS(W, 'p'):
                    text = ''.join(map(node_text, paragraph_text_nodes(p)))
                    matches = list(TOKEN.finditer(text))
                    remainder = TOKEN.sub('', text)
                    if '{{' in remainder or '}}' in remainder:
                        raise ValueError('模板中存在不完整的标记。请使用 {{ 姓名 }}，并将一个标记放在同一段落内。')
                    for match in matches:
                        key = match.group(1).strip()
                        if not key:
                            raise ValueError('填写标记需要有名称，例如 {{ 姓名 }}。')
                        if key not in found:
                            found.append(key)
                doc.unlink()
    except BadZipFile as exc:
        raise ValueError('无法读取模板。旧版 .doc 请先用 Word 另存为 .docx。') from exc
    for anchor in scan_image_anchors(path):
        if anchor['tag'] not in found:
            found.append(anchor['tag'])
    return found


def validate_values(config, values, allow_missing=False):
    for field in config['fields']:
        value = str(values.get(field['key'], ''))
        if field.get('required') and not allow_missing and not value.strip():
            raise ValueError(f'请填写「{field["label"]}」。')
        if INVALID_XML.search(value):
            raise ValueError(f'「{field["label"]}」含有不支持的控制字符，请重新输入。')
        limit = field.get('max_length', 0)
        if limit and len(value) > limit:
            raise ValueError(f'「{field["label"]}」最多填写 {limit} 字，当前为 {len(value)} 字。')


def _attr(el, name):
    return el.getAttributeNS(W, name) or el.getAttribute('w:' + name)


def _run_props(run):
    for child in run.childNodes:
        if child.nodeType == Node.ELEMENT_NODE and child.namespaceURI == W and child.localName == 'rPr':
            return child
    return None


def _has_real_font_format(rpr):
    """有显式字号，或 rFonts 里指定了具体字体（只有 w:hint 不算设置过格式）。"""
    if rpr is None:
        return False
    if rpr.getElementsByTagNameNS(W, 'sz'):
        return True
    for f in rpr.getElementsByTagNameNS(W, 'rFonts'):
        if any(_attr(f, a) for a in ('ascii', 'eastAsia', 'hAnsi', 'cs')):
            return True
    return False


def _formatted_runs(container, skip_run):
    for run in container.getElementsByTagNameNS(W, 'r'):
        if run is skip_run:
            continue
        rpr = _run_props(run)
        if not _has_real_font_format(rpr):
            continue
        text = ''.join(node_text(t) for t in run.getElementsByTagNameNS(W, 't'))
        if text.strip():
            yield run, rpr, text


def _ancestor(node, local):
    while node is not None:
        if node.nodeType == Node.ELEMENT_NODE and node.namespaceURI == W and node.localName == local:
            return node
        node = node.parentNode
    return None


def find_donor_rpr(node):
    """标记所在 run 没有显式字体/字号时，就近借用可见文字的格式。

    顺序：同段落其他文字 → 同表格单元格 → 同表格行 → 同表格。找不到就不动。"""
    run = node.parentNode
    if run is None or run.nodeType != Node.ELEMENT_NODE or run.localName != 'r':
        return None
    own = _run_props(run)
    if _has_real_font_format(own):
        return None
    paragraph = _ancestor(run, 'p')
    if paragraph is None:
        return None
    for _run, rpr, text in _formatted_runs(paragraph, run):
        if not TOKEN.search(text):
            return rpr
    candidates = list(_formatted_runs(paragraph, run))
    if candidates:
        return candidates[0][1]
    scopes = []
    cell = _ancestor(paragraph, 'tc')
    if cell is not None:
        scopes.append(cell)
        row = _ancestor(cell, 'tr')
        table = _ancestor(cell, 'tbl')
        for scope in (row, table):
            if scope is not None:
                scopes.append(scope)
    for scope in scopes:
        for _run, rpr, text in _formatted_runs(scope, run):
            if not TOKEN.search(text):
                return rpr
        for _run, rpr, _text in _formatted_runs(scope, run):
            return rpr
    return None


def apply_donor(run, donor):
    """只把借来的字体/字号补进 run；下划线、加粗等已有格式全部保留。"""
    own = _run_props(run)
    if own is None:
        run.insertBefore(donor.cloneNode(True), run.firstChild)
        return
    donor_fonts = donor.getElementsByTagNameNS(W, 'rFonts')
    own_fonts = own.getElementsByTagNameNS(W, 'rFonts')
    if donor_fonts:
        if own_fonts:
            target = own_fonts[0]
            for name in ('ascii', 'eastAsia', 'hAnsi', 'cs'):
                value = _attr(donor_fonts[0], name)
                if value and not _attr(target, name):
                    target.setAttribute('w:' + name, value)
        else:
            own.insertBefore(donor_fonts[0].cloneNode(True), own.firstChild)
    for tag in ('sz', 'szCs'):
        elements = donor.getElementsByTagNameNS(W, tag)
        if elements and not own.getElementsByTagNameNS(W, tag):
            own.appendChild(elements[0].cloneNode(True))


def clone_paragraph(doc, source):
    clone = doc.createElementNS(W, 'w:p')
    for child in source.childNodes:
        if child.nodeType == Node.ELEMENT_NODE and child.namespaceURI == W and child.localName == 'pPr':
            clone.appendChild(child.cloneNode(True))
            break
    return clone


def write_value_inplace(doc, node, value, breaks_sink):
    """把值写入原 <w:t>；换行用软回车占位（后续再拆成真实段落），制表符转 w:tab。"""
    parts = re.split(r'([\n\t])', value)
    set_text(node, parts[0])
    anchor = node
    for part in parts[1:]:
        if part == '\n':
            element = doc.createElementNS(W, 'w:br')
            breaks_sink.append(element)
        elif part == '\t':
            element = doc.createElementNS(W, 'w:tab')
        else:
            element = doc.createElementNS(W, 'w:t')
            set_text(element, part)
        node.parentNode.insertBefore(element, anchor.nextSibling)
        anchor = element


def split_paragraph_at_break(doc, br):
    """把一个软回车升级为真实分段：回车后面的内容整体搬到复制了段落属性的新段落。"""
    run = br.parentNode
    paragraph = run.parentNode
    tail = br.nextSibling
    run.removeChild(br)
    new_paragraph = clone_paragraph(doc, paragraph)
    remainder = doc.createElementNS(W, 'w:r')
    rpr = _run_props(run)
    if rpr is not None:
        remainder.appendChild(rpr.cloneNode(True))
    while tail is not None:
        nxt = tail.nextSibling
        run.removeChild(tail)
        remainder.appendChild(tail)
        tail = nxt
    has_content = any(child.localName in ('t', 'br', 'tab', 'drawing', 'pict', 'fldChar')
                      for child in remainder.childNodes)
    if has_content:
        new_paragraph.appendChild(remainder)
    sibling = run.nextSibling
    while sibling is not None:
        nxt = sibling.nextSibling
        paragraph.removeChild(sibling)
        new_paragraph.appendChild(sibling)
        sibling = nxt
    paragraph.parentNode.insertBefore(new_paragraph, paragraph.nextSibling)
    return new_paragraph


def shrink_trailing_empty_paragraph(doc):
    """表格结尾后的空段落压缩到最小行高，避免长内容把表格顶到页尾时多出一整页空白。
    OOXML 要求表格后必须保留段落，因此不删除只压缩；其他结构不动。"""
    for body in doc.getElementsByTagNameNS(W, 'body'):
        kids = [c for c in body.childNodes if c.nodeType == Node.ELEMENT_NODE]
        if len(kids) < 3:
            continue
        if kids[-1].localName != 'sectPr' or kids[-2].localName != 'p' or kids[-3].localName != 'tbl':
            continue
        paragraph = kids[-2]
        xml = paragraph.toxml()
        for marker in ('<w:t', 'w:br', 'w:drawing', 'w:pict', 'w:object', 'fldChar', 'w:tab', 'w:bookmarkStart'):
            if marker in xml:
                return False
        props = next((c for c in paragraph.childNodes
                      if c.nodeType == Node.ELEMENT_NODE and c.localName == 'pPr'), None)
        if props is not None:
            for spacing in props.getElementsByTagNameNS(W, 'spacing'):
                if spacing.getAttributeNS(W, 'lineRule') == 'exact' and \
                        int(spacing.getAttributeNS(W, 'line') or '0') <= 20:
                    return False
        if props is None:
            props = doc.createElementNS(W, 'w:pPr')
            paragraph.insertBefore(props, paragraph.firstChild)
        for spacing in props.getElementsByTagNameNS(W, 'spacing'):
            props.removeChild(spacing)
        spacing = doc.createElementNS(W, 'w:spacing')
        spacing.setAttribute('w:line', '20')
        spacing.setAttribute('w:lineRule', 'exact')
        props.appendChild(spacing)
        for old in props.getElementsByTagNameNS(W, 'rPr'):
            props.removeChild(old)
        mark_props = doc.createElementNS(W, 'w:rPr')
        size = doc.createElementNS(W, 'w:sz')
        size.setAttribute('w:val', '2')
        size_cs = doc.createElementNS(W, 'w:szCs')
        size_cs.setAttribute('w:val', '2')
        mark_props.appendChild(size)
        mark_props.appendChild(size_cs)
        props.appendChild(mark_props)
        return True
    return False


_RPR_ORDER = ('rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
              'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
              'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position',
              'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
              'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish',
              'oMath')


def _reorder_rpr(rpr):
    kids = [c for c in rpr.childNodes if c.nodeType == Node.ELEMENT_NODE]
    if len(kids) < 2:
        return
    def rank(el):
        try:
            return _RPR_ORDER.index(el.localName)
        except ValueError:
            return len(_RPR_ORDER)
    for el in sorted(kids, key=rank):
        rpr.appendChild(el)


def _direct_children(parent, local_name):
    return [child for child in parent.childNodes
            if child.nodeType == Node.ELEMENT_NODE and child.localName == local_name]


def _cell_valign(tc):
    for pr in _direct_children(tc, 'tcPr'):
        for valign in _direct_children(pr, 'vAlign'):
            return (valign.getAttributeNS(W, 'val') or valign.getAttribute('w:val') or '').lower()
    return ''


def _set_cell_valign(doc, tc, value):
    props = _direct_children(tc, 'tcPr')
    if props:
        tcpr = props[0]
    else:
        tcpr = doc.createElementNS(W, 'w:tcPr')
        tc.insertBefore(tcpr, tc.firstChild)
    valigns = _direct_children(tcpr, 'vAlign')
    if valigns:
        valigns[0].setAttribute('w:val', value)
    else:
        node = doc.createElementNS(W, 'w:vAlign')
        node.setAttribute('w:val', value)
        tcpr.appendChild(node)


def _align_label_value_tables(doc):
    """标签单元格底端对齐、值单元格却顶端对齐时，把值也改到底端，贴住下划线。"""
    changed = False
    for table in doc.getElementsByTagNameNS(W, 'tbl'):
        for row in _direct_children(table, 'tr'):
            cells = _direct_children(row, 'tc')
            if len(cells) < 2:
                continue
            label = ''.join(node_text(t) for t in cells[0].getElementsByTagNameNS(W, 't'))
            if not label.strip().endswith(('：', ':')):
                continue
            if _cell_valign(cells[0]) != 'bottom':
                continue
            for cell in cells[1:]:
                if _cell_valign(cell) == 'top':
                    _set_cell_valign(doc, cell, 'bottom')
                    changed = True
    return changed


def _field_format(config, key):
    fmt = config.get('format') or {}
    if not fmt:
        return None
    size = (fmt.get('field_size_pt') or {}).get(key)
    if size in (None, ''):
        size = fmt.get('default_size_pt') or None
    bold = key in (fmt.get('field_bold') or [])
    font = (fmt.get('field_font') or {}).get(key)
    if not size and not bold and not font:
        return None
    return {'size': size, 'bold': bold, 'font': font}


def _ensure_rpr(doc, run):
    rpr = _run_props(run)
    if rpr is None:
        rpr = doc.createElementNS(W, 'w:rPr')
        run.insertBefore(rpr, run.firstChild)
    return rpr


def _set_rpr_tag(doc, rpr, tag, val=None):
    for old in list(rpr.getElementsByTagNameNS(W, tag)):
        rpr.removeChild(old)
    el = doc.createElementNS(W, 'w:' + tag)
    if val is not None:
        el.setAttribute('w:val', val)
    rpr.appendChild(el)


def _apply_run_format(doc, run, spec):
    if not spec:
        return
    rpr = _ensure_rpr(doc, run)
    if spec.get('font'):
        fonts = rpr.getElementsByTagNameNS(W, 'rFonts')
        f = fonts[0] if fonts else doc.createElementNS(W, 'w:rFonts')
        for attr in ('ascii', 'eastAsia', 'hAnsi'):
            f.setAttribute('w:' + attr, spec['font'])
        if not fonts:
            rpr.insertBefore(f, rpr.firstChild)
    if spec.get('bold'):
        _set_rpr_tag(doc, rpr, 'b', 'true')
        _set_rpr_tag(doc, rpr, 'bCs', 'true')
    if spec.get('size'):
        half = str(int(round(float(spec['size']) * 2)))
        _set_rpr_tag(doc, rpr, 'sz', half)
        _set_rpr_tag(doc, rpr, 'szCs', half)
    _reorder_rpr(rpr)


def _clone_rpr(run):
    rpr = _run_props(run)
    return rpr.cloneNode(True) if rpr is not None else None


def _make_text_run(doc, rpr, text):
    run = doc.createElementNS(W, 'w:r')
    if rpr is not None:
        run.appendChild(rpr)
    node = doc.createElementNS(W, 'w:t')
    node.setAttribute('xml:space', 'preserve')
    node.appendChild(doc.createTextNode(text))
    run.appendChild(node)
    return run


def _split_paragraph_tokens(doc, paragraph, keys):
    """把与文字混排在同一 run 里的 {{标记}} 拆成独立 run（仅当该 run 只含 rPr/t 子节点）。
    拆分后标记独占 run，格式控制即可安全套用，而不必跳过。"""
    nodes = paragraph_text_nodes(paragraph)
    texts = [node_text(t) for t in nodes]
    text = ''.join(texts)
    found = [m for m in TOKEN.finditer(text) if m.group(1).strip() in keys]
    if not found:
        return False
    offsets, cursor = [], 0
    for t in texts:
        offsets.append(cursor)
        cursor += len(t)
    changed = False
    for m in reversed(found):
        start = next(i for i, off in enumerate(offsets) if off <= m.start() < off + len(texts[i]))
        end = next(i for i, off in enumerate(offsets) if off < m.end() <= off + len(texts[i]))
        if start != end:
            continue
        left = m.start() - offsets[start]
        right = m.end() - offsets[start]
        if left == 0 and right == len(texts[start]):
            continue
        run = nodes[start].parentNode
        if run.nodeType != Node.ELEMENT_NODE or run.localName != 'r':
            continue
        if any(c.nodeType == Node.ELEMENT_NODE and c.localName not in ('rPr', 't')
               for c in run.childNodes):
            continue
        if len(run.getElementsByTagNameNS(W, 't')) != 1:
            continue
        ts = texts[start]
        pre, tok, suf = ts[:left], ts[left:right], ts[right:]
        set_text(nodes[start], pre)
        ref = run
        tok_run = _make_text_run(doc, _clone_rpr(run), tok)
        run.parentNode.insertBefore(tok_run, run.nextSibling)
        ref = tok_run
        if suf:
            suf_run = _make_text_run(doc, _clone_rpr(run), suf)
            ref.parentNode.insertBefore(suf_run, ref.nextSibling)
        changed = True
    return changed


def manual_format_rules(config):
    """Turn the 格式/字号 dialog settings into per-field rules.

    The dialog used to reach only the legacy ``{{ }}`` path, so bookmark templates
    silently ignored it; folding it into the same rule structure fixes that.
    """
    fmt = config.get('format') or {}
    if not fmt:
        return {}
    rules = {}
    default_size = fmt.get('default_size_pt')
    for field in config.get('fields', []):
        key = field.get('key')
        if not key:
            continue
        entry = {}
        if default_size:
            entry['size_pt'] = default_size
        size = (fmt.get('field_size_pt') or {}).get(key)
        if size:
            entry['size_pt'] = size
        if key in (fmt.get('field_bold') or []):
            entry['bold'] = True
        font = (fmt.get('field_font') or {}).get(key)
        if font:
            if re.match(r'\s*[A-Za-z]', str(font)):
                entry['font_ascii'] = str(font).strip()
            else:
                entry['font_east'] = str(font).strip()
        if entry:
            rules[key] = entry
    return rules


def render_word(template, output, config, values, validate=True, keep_empty_bookmarks=False,
                allow_missing=False, report=None):
    template, output = Path(template), Path(output)
    if template.resolve() == output.resolve():
        raise ValueError('生成文件不能覆盖模板，请选择其他保存位置。')
    if validate:
        validate_values(config, values, allow_missing=allow_missing)
    fmt_skipped = set()
    req_report = None
    bookmark_fields = [f for f in config['fields'] if f.get('bookmark')]
    control_fields = [f for f in config['fields'] if f.get('control')]
    pre_path = None
    # The pre-pass is needed for bookmarks/controls anyway; 'req' is the only extra
    # trigger, so legacy {{ }} templates keep their original byte-for-byte path.
    pre_needed = bool(bookmark_fields or control_fields or config.get('req'))
    if pre_needed:
        existing = set(bookmarkfill.scan_bookmarks(template))
        for field in bookmark_fields:
            if field['bookmark'] not in existing:
                raise ValueError(f'模板缺少填写位「{field.get("label", field["key"])}」，请重新生成模板。')
        doc = Document(str(template))
        spec = None
        manual = manual_format_rules(config)
        if config.get('req') or manual:
            import formatapply
            spec = dict(config.get('req') or {})
            spec.setdefault('fields', config['fields'])
            merged = {k: dict(v) for k, v in (spec.get('field_rules') or {}).items()}
            for key, rules in manual.items():
                merged.setdefault(key, {}).update(rules)   # 手动设置优先于模板说明
            spec['field_rules'] = merged
        if spec:
            try:
                req_report = formatapply.apply(doc, spec, config.get('req_options'),
                                               phase='pre')
            except Exception as exc:                      # never lose the filled document
                req_report = {'error': str(exc)}
        for field in bookmark_fields:
            value = str(values.get(field['key'], ''))
            bookmarkfill.replace_bookmark(doc, field['bookmark'], value,
                                          keep_empty=keep_empty_bookmarks)
        if control_fields:
            controlfill.fill_controls(doc, control_fields, values)
        if spec and not (req_report or {}).get('error'):
            try:
                post = formatapply.apply(doc, spec, config.get('req_options'),
                                         phase='post')
                for name in ('runs', 'paragraphs', 'removed', 'kept', 'blank_lines'):
                    req_report[name] = req_report.get(name, 0) + post.get(name, 0)
                req_report['skipped'].extend(post.get('skipped') or [])
            except Exception as exc:
                req_report['post_error'] = str(exc)
        output.parent.mkdir(parents=True, exist_ok=True)
        handle, pre_path = tempfile.mkstemp(suffix='.docx', dir=output.parent)
        os.close(handle)
        ooxmlcheck.normalize(doc)   # 来源模板元素顺序不合规时，先排好再写盘
        doc.save(pre_path)
        template = Path(pre_path)
    image_data = {}
    for image in config.get('images', []):
        value = values.get(image['key'], '__original__')
        if value in ('', '__original__'):
            continue
        try:
            data = Path(value).read_bytes()
        except OSError:
            raise ValueError('所选图片已失效或无法读取（可能已被移动、删除或换过电脑），请重新选择图片。')
        if not data:
            raise ValueError('所选图片是空文件，请重新选择图片。')
        image_data[image['key']] = data
    keys = scan_fields(template)
    fields = {f['key'] for f in config['fields']}
    fields |= {i['key'] for i in config.get('images', [])}
    if set(keys) - fields:
        raise ValueError('模板字段与配置不一致，请重新导入模板。')
    anchors = scan_image_anchors(template)
    anchor_data = {}
    for anchor in anchors:
        value = str(values.get(anchor['tag'], '')).strip()
        if not value or value == '__original__':
            continue
        try:
            data = Path(value).read_bytes()
        except OSError:
            raise ValueError(f'图片「{anchor["tag"]}」所选文件已失效或无法读取，请重新选择。')
        signs = IMAGE_SIGNS.get(anchor['fmt'])
        if not data or (signs and not any(data.startswith(s) for s in signs)):
            raise ValueError(f'图片「{anchor["tag"]}」与原图格式不一致（原图为 {anchor["fmt"] or "未知"}），'
                             '请选择同格式的图片。')
        anchor_data[(anchor['part'], anchor['relationship'])] = (anchor, data)
    output.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(suffix='.docx', dir=output.parent)
    os.close(handle)
    try:
        with ZipFile(template) as source, ZipFile(temporary, 'w') as target:
            changes = {}
            for name in word_parts(source):
                data = source.read(name)
                doc = parse_xml(data)
                changed = False
                for p in doc.getElementsByTagNameNS(W, 'p'):
                    nodes = paragraph_text_nodes(p)
                    original = [node_text(n) for n in nodes]
                    text = ''.join(original)
                    matches = list(TOKEN.finditer(text))
                    if not matches:
                        continue
                    if config.get('format'):
                        want = {m.group(1).strip() for m in matches
                                if _field_format(config, m.group(1).strip())}
                        if want and _split_paragraph_tokens(doc, p, want):
                            nodes = paragraph_text_nodes(p)
                            original = [node_text(n) for n in nodes]
                            text = ''.join(original)
                            matches = list(TOKEN.finditer(text))
                            if not matches:
                                continue
                    offsets, cursor = [], 0
                    for value in original:
                        offsets.append(cursor)
                        cursor += len(value)
                    updated = original[:]
                    positions = {}
                    for match in reversed(matches):
                        start = next(i for i, off in enumerate(offsets) if off <= match.start() < off + len(original[i]))
                        end = next(i for i, off in enumerate(offsets) if off < match.end() <= off + len(original[i]))
                        value = str(values.get(match.group(1).strip(), '')).replace('\r\n', '\n').replace('\r', '\n')
                        width = config.get('underline_widths', {}).get(match.group(1).strip(), 0)
                        if width and nodes[start].parentNode.getElementsByTagNameNS(W, 'u'):
                            used = sum(2 if unicodedata.east_asian_width(c) in ('W', 'F') else 1 for c in value)
                            value += ' ' * max(0, width - used)
                        left = match.start() - offsets[start]
                        right = match.end() - offsets[end]
                        positions[match.group(1).strip()] = (start, end, left, right)
                        if start == end:
                            updated[start] = updated[start][:left] + value + updated[start][right:]
                        else:
                            updated[start] = updated[start][:left] + value
                            for i in range(start + 1, end):
                                updated[i] = ''
                            updated[end] = updated[end][right:]
                    fmt_specs = {}
                    for match in matches:
                        key = match.group(1).strip()
                        spec = _field_format(config, key)
                        if spec:
                            fmt_specs[key] = spec
                    whole_keys = {}
                    for key, (s, e, l, r) in positions.items():
                        if key in fmt_specs and s == e and l == 0 and r == len(original[s]):
                            whole_keys[s] = key
                    breaks = []
                    for idx, (node, value) in enumerate(zip(nodes, updated)):
                        if node_text(node) == value:
                            continue
                        run = node.parentNode
                        if run.nodeType == Node.ELEMENT_NODE and run.localName == 'r':
                            donor = find_donor_rpr(node)
                            if donor is not None:
                                apply_donor(run, donor)
                            key = whole_keys.get(idx)
                            if key:
                                if len(run.getElementsByTagNameNS(W, 't')) == 1:
                                    _apply_run_format(doc, run, fmt_specs[key])
                                else:
                                    fmt_skipped.add(key)
                        write_value_inplace(doc, node, value, breaks)
                        changed = True
                    for key in fmt_specs:
                        if key not in whole_keys.values():
                            fmt_skipped.add(key)
                    for br in breaks:
                        split_paragraph_at_break(doc, br)
                        changed = True
                for image in config.get('images', []):
                    value = values.get(image['key'], '__original__')
                    if value == '__original__':
                        continue
                    if value:
                        changes[image['part']] = image_data[image['key']]
                        for node in doc.getElementsByTagNameNS(V, 'imagedata'):
                            if node.getAttributeNS(R, 'id') == image['relationship']:
                                for attribute in ('croptop', 'cropbottom', 'cropleft', 'cropright'):
                                    if node.hasAttribute(attribute):
                                        node.removeAttribute(attribute)
                                        changed = True
                    else:
                        for node in list(doc.getElementsByTagNameNS(V, 'imagedata')):
                            if node.getAttributeNS(R, 'id') == image['relationship']:
                                shape = node.parentNode
                                shape.parentNode.removeChild(shape)
                                changed = True
                if _align_label_value_tables(doc):
                    changed = True
                if name == 'word/document.xml' and shrink_trailing_empty_paragraph(doc):
                    changed = True
                if _remove_textboxes(doc):
                    changed = True
                if _clear_highlights(doc):
                    changed = True
                changes[name] = doc.toxml(encoding='utf-8') if changed else data
                doc.unlink()
            styles_name = 'word/styles.xml'
            if styles_name in source.namelist() and styles_name not in changes:
                styles_doc = parse_xml(source.read(styles_name))
                if _clear_highlights(styles_doc):
                    changes[styles_name] = styles_doc.toxml(encoding='utf-8')
                styles_doc.unlink()
            touched = {}
            for (part, rid), (anchor, data) in anchor_data.items():
                changes[anchor['media']] = data
                if part not in touched:
                    touched[part] = parse_xml(changes.get(part) or source.read(part))
                removed = 0
                for blip in touched[part].getElementsByTagNameNS(A, 'blip'):
                    if blip.getAttributeNS(R, 'embed') != rid:
                        continue
                    for ext in list(blip.getElementsByTagNameNS(A, 'extLst')):
                        blip.removeChild(ext)
                        removed += 1
            for part, doc in touched.items():
                changes[part] = doc.toxml(encoding='utf-8')
                doc.unlink()
            for info in source.infolist():
                target.writestr(info, changes.get(info.filename, source.read(info.filename)))
        os.replace(temporary, output)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
        if pre_path and os.path.exists(pre_path):
            os.unlink(pre_path)
    if report is not None:
        report['req'] = req_report
    return sorted(fmt_skipped)


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temp = tempfile.mkstemp(suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(handle, 'w', encoding='utf-8') as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2)
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))
