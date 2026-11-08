# -*- coding: utf-8 -*-
"""生成后的 OOXML 自检：把“Word 打开提示文件已损坏”挡在交付之前。

思路来自 OfficeCLI 的 validate + view issues：先查结构合法性（元素顺序、书签配对、
表格骨架、样式与关系引用），再查取值是否落在 Word 允许的区间。
返回问题清单，界面如实列出，不静默修、也不假装没问题。
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_EMBED = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
R_ID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

LEVEL_ERROR = '错误'
LEVEL_WARN = '提醒'


def _issue(issues, level, where, message, fix=''):
    issues.append({'level': level, 'where': where, 'message': message, 'fix': fix})
    return issues[-1]


def check(path):
    """→ 问题列表；空列表表示结构合法。"""
    path = Path(path)
    issues = []
    doc = Document(str(path))
    body = doc.element.body
    _check_order(body, issues)
    _check_bookmarks(body, issues)
    _check_tables(body, issues)
    _check_values(body, issues)
    _check_styles(doc, body, issues)
    _check_relationships(path, body, issues)
    _check_section(body, issues)
    return issues


def _normalize_holder(holder, order):
    """把 pPr/rPr 的子元素按 schema 顺序重排；返回是否改动过。"""
    rank = {name: index for index, name in enumerate(order)}
    children = list(holder)
    ordered = sorted(children, key=lambda el: rank.get(el.tag.split('}')[-1], len(order)))
    if [id(node) for node in ordered] == [id(node) for node in children]:
        return 0
    for node in ordered:
        holder.append(node)      # append 会把元素移到末尾，最终顺序即 ordered
    return 1


def normalize(doc):
    """保存前统一修正 pPr/rPr 顺序，来源模板本身不合规也不会带出“需要修复”的文件。"""
    roots = [doc.element.body]
    try:
        roots.append(doc.styles.element)
    except Exception:
        pass
    fixed = 0
    for root in roots:
        for tag, order in (('pPr', _PPR_ORDER), ('rPr', _RPR_ORDER)):
            for holder in list(root.iter(qn('w:' + tag))):
                fixed += _normalize_holder(holder, order)
    return fixed


def _where(element):
    tag = element.tag.split('}')[-1]
    parent = element.getparent()
    chain = []
    while parent is not None and len(chain) < 3:
        chain.append(parent.tag.split('}')[-1])
        parent = parent.getparent()
    return f'{"/".join(reversed(chain))}/{tag}'


def _check_order(body, issues):
    """w:pPr / w:rPr 的子元素必须按 schema 顺序，乱序 Word 会提示需要修复。"""
    for tag, order in (('pPr', _PPR_ORDER), ('rPr', _RPR_ORDER)):
        for holder in body.iter(qn('w:' + tag)):
            names = [child.tag.split('}')[-1] for child in holder]
            ranks = [order.index(name) for name in names if name in order]
            if any(ranks[i] > ranks[i + 1] for i in range(len(ranks) - 1)):
                bad = next(names[i + 1] for i in range(len(names) - 1)
                           if names[i] in order and names[i + 1] in order
                           and order.index(names[i]) > order.index(names[i + 1]))
                _issue(issues, LEVEL_ERROR, _where(holder),
                       f'{tag} 里 {bad} 的位置不对，Word 打开时会提示“文件需要修复”',
                       '已按 schema 顺序重排')


_PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
              'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
              'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
              'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
              'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
              'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
              'sectPr', 'pPrChange']
_RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
              'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
              'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs',
              'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs',
              'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange']


def _check_bookmarks(body, issues):
    starts, ends, ids = {}, {}, {}
    for node in body.iter(qn('w:bookmarkStart')):
        name = node.get(qn('w:name')) or ''
        bid = node.get(qn('w:id')) or ''
        starts.setdefault(name, []).append(bid)
        ids.setdefault(bid, []).append(name)
    for node in body.iter(qn('w:bookmarkEnd')):
        ends.setdefault(node.get(qn('w:id')) or '', []).append(node)
    for name, bids in starts.items():
        if name.startswith('_') or not name:
            continue
        if len(bids) > 1:
            _issue(issues, LEVEL_ERROR, '书签', f'书签名「{name}」重复 {len(bids)} 次，'
                   'Word 要求书签名唯一', '填写位书签名按位置编号，不会重名')
        for bid in bids:
            if bid not in ends:
                _issue(issues, LEVEL_ERROR, '书签', f'书签「{name}」只有起点没有终点，文档会打不开')
    for bid, nodes in ends.items():
        if bid not in ids:
            _issue(issues, LEVEL_WARN, '书签', f'有书签终点找不到对应起点（id={bid}）')


def _check_tables(body, issues):
    for table in body.iter(qn('w:tbl')):
        grid = table.find(qn('w:tblGrid'))
        columns = len(grid.findall(qn('w:gridCol'))) if grid is not None else 0
        for row in table.findall(qn('w:tr')):
            cells = row.findall(qn('w:tc'))
            if not cells:
                _issue(issues, LEVEL_ERROR, _where(row), '表格里有空行（没有任何单元格）')
                continue
            span = 0
            for cell in cells:
                properties = cell.find(qn('w:tcPr'))
                if properties is not None:
                    wide = properties.find(qn('w:gridSpan'))
                    span += int(wide.get(qn('w:val'))) if wide is not None else 1
                else:
                    span += 1
                if not [child for child in cell if child.tag in (qn('w:p'), qn('w:tbl'))]:
                    _issue(issues, LEVEL_ERROR, _where(cell), '单元格末尾没有段落，Word 会拒绝打开')
            if columns and span != columns:
                _issue(issues, LEVEL_WARN, _where(row),
                       f'这一行占了 {span} 列，但表格网格是 {columns} 列')


def _check_values(body, issues):
    for node in body.iter(qn('w:sz')):
        try:
            value = int(node.get(qn('w:val')))
        except (TypeError, ValueError):
            continue
        if value <= 0 or value > 4000:
            _issue(issues, LEVEL_WARN, _where(node), f'字号值 {value}（半磅）超出 Word 常用范围')
    for node in body.iter(qn('w:spacing')):
        for attr, limit in (('line', 86400), ('before', 31680), ('after', 31680)):
            raw = node.get(qn('w:' + attr))
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                _issue(issues, LEVEL_ERROR, _where(node), f'行距 {attr}={raw} 不是整数')
                continue
            if value < 0 or value > limit:
                _issue(issues, LEVEL_WARN, _where(node),
                       f'{attr}={value} 超出 Word 允许范围（0–{limit}）')
    for node in body.iter(qn('w:pgSz')):
        for attr in ('w', 'h'):
            raw = node.get(qn('w:' + attr))
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                _issue(issues, LEVEL_ERROR, _where(node), f'纸张 {attr}={raw} 不是整数')
                continue
            if not 240 <= value <= 31680:
                _issue(issues, LEVEL_ERROR, _where(node),
                       f'纸张 {attr}={value} 超出 Word 允许范围（240–31680 缇，约 0.4–55.9cm）')
    for node in body.iter(qn('w:pgMar')):
        for attr in ('top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'):
            raw = node.get(qn('w:' + attr))
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                _issue(issues, LEVEL_ERROR, _where(node), f'页边距 {attr}={raw} 不是整数')
                continue
            if value < 0 or value > 31680:
                _issue(issues, LEVEL_ERROR, _where(node),
                       f'页边距 {attr}={value} 超出 Word 允许范围（0–31680 缇）')
    for node in body.iter(qn('w:color')):
        value = (node.get(qn('w:val')) or '').upper()
        if value and value not in ('AUTO',) and not re.fullmatch(r'[0-9A-F]{6}', value):
            _issue(issues, LEVEL_WARN, _where(node), f'颜色值“{value}”不是 6 位十六进制')


def _check_styles(doc, body, issues):
    known = set()
    for style in doc.styles:
        known.add(getattr(style, 'style_id', None))
    known.discard(None)
    for tag in ('pStyle', 'rStyle', 'tblStyle'):
        for node in body.iter(qn('w:' + tag)):
            value = node.get(qn('w:val'))
            if value and value not in known:
                _issue(issues, LEVEL_WARN, _where(node),
                       f'引用了文档里没有的样式“{value}”，Word 会退回默认样式')


def _check_relationships(path, body, issues):
    """图片/超链接引用的关系必须真的在包里；缺文件是 Word 打不开的常见原因。"""
    with zipfile.ZipFile(str(path)) as archive:
        members = set(archive.namelist())
        targets = {}
        for part in [name for name in members if name.endswith('.rels')]:
            base = posixpath.dirname(posixpath.dirname(part))
            try:
                xml = archive.read(part).decode('utf-8', 'replace')
            except OSError:
                continue
            for match in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*/>', xml):
                identifier, target = match.group(1), match.group(2)
                if target.startswith(('http', 'file:', 'mailto:')) or 'External' in match.group(0):
                    continue
                resolved = posixpath.normpath(posixpath.join(base, target)) if base else target
                targets[identifier] = resolved.lstrip('/')
    for tag in (qn('a:blip'), qn('w:imagedata')):
        for node in body.iter(tag):
            for attribute in (R_EMBED, R_ID):
                identifier = node.get(attribute)
                if not identifier:
                    continue
                where = targets.get(identifier)
                if where and where not in members:
                    _issue(issues, LEVEL_ERROR, _where(node),
                           f'图片引用 {identifier} 指向 {where}，但包里找不到这个文件')


def _check_section(body, issues):
    children = list(body)
    if not children:
        _issue(issues, LEVEL_ERROR, 'body', '文档主体是空的')
        return
    last = children[-1]
    if last.tag != qn('w:sectPr'):
        _issue(issues, LEVEL_WARN, 'body', '末尾没有分节属性，Word 会用默认页面设置')
    for child in children[:-1]:
        if child.tag == qn('w:sectPr'):
            _issue(issues, LEVEL_WARN, _where(child), '分节属性出现在文档中间（正常应在段落里）')


def summarize(issues):
    """给界面用的一行行文字。"""
    if not issues:
        return ['自检通过：书签配对、表格骨架、样式引用、取值范围都符合 Word 的要求。']
    lines = []
    for issue in issues[:20]:
        fix = f'（{issue["fix"]}）' if issue.get('fix') else ''
        lines.append(f'{issue["level"]}｜{issue["where"]}：{issue["message"]}{fix}')
    if len(issues) > 20:
        lines.append(f'…还有 {len(issues) - 20} 条同类问题')
    return lines
