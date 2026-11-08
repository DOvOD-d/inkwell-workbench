"""Placeholder-free recognition.

The document is the source of truth: label/value cells, section headings,
underlines, bracket prompts, content controls and legacy form fields are all
read directly from OOXML.  No ``{{ }}`` token is required or produced.
"""
from __future__ import annotations

import re

from docx import Document
from docx.oxml.ns import qn

import bookmarkfill
import controlfill
import formparse

NOTE_RE = formparse.NOTE_RE
HEADING_RE = formparse.HEADING_RE
CN_HEADING_RE = formparse.CN_HEADING_RE
NUMBERED_RE = formparse.NUMBERED_RE

BLANK_RE = re.compile(r'[_＿﹍﹏]{2,}|…{2,}|\.{4,}|·{3,}|[ \u3000\t]{4,}')
LINE_RE = re.compile(r'^[\s\u3000_＿﹍﹏—\-·.…]+$')
DATE_RE = re.compile(
    r'(?:(?:(?:19|20)\d{2})\s*年|[_＿﹏]{2,}\s*年)\s*'
    r'(?:[_＿﹏]{2,}|\d{1,2})?\s*月\s*(?:[_＿﹏]{2,}|\d{1,2})?\s*日')
BRACKET_RE = re.compile(
    r'【([^】\n]{0,30})】|［([^］\n]{0,30})］|\[([^\]\n]{0,30})\]|'
    r'（([^）\n]{0,20})）|\(([^)\n]{0,20})\)')
PLACEHOLDER_WORDS = re.compile(
    r'填写|请输入|在此输入|此处|姓名|名称|标题|日期|时间|单位|学院|专业|学号|'
    r'电话|手机|邮箱|地址|题目|签字|意见|内容|摘要|备注|编号|金额|数量')
PLACEHOLDER_ONLY = re.compile(r'^(?:[\s\u3000_＿﹍…]*|[Xx×＊*]{2,})$')
MULTI_WORDS = ('内容', '说明', '意见', '原因', '简介', '安排', '计划', '总结',
               '描述', '情况', '方案', '措施', '要求', '目标', '任务', '过程',
               '经历', '论文', '摘要', '参考文献', '备注', '建议', '评语', '论证')
NOISE_LABELS = ('填表说明', '填写说明', '说明', '注意事项', '注意', '注',
                '请阅读', '温馨提示', '以下', '上述', '如下', '格式', '要求')
SKIP_LABELS = ('盖章', '公章', '签字', '签名', '手签')
CONTROL_NOISE = ('页码', '目录', '页眉', '页脚')


def _text(p):
    return ''.join(r.text for r in p.runs)


def _clean(text):
    text = re.sub(r'[\s\u3000]+', '', text or '')
    text = re.sub(r'^[（(【\[［]+|[）)】\]］]+$', '', text)
    text = re.sub(r'[：:；;,，。、\.]+$', '', text)
    return text.strip()


def _short_label(text):
    text = _clean(text)
    if len(text) > 24:
        text = text[-24:]
    return text


def _noise(label):
    plain = _clean(label)
    if not plain:
        return True
    if plain in NOISE_LABELS:
        return True
    if any(plain.startswith(x) for x in ('填写说明', '填表说明', '注意事项', '备注说明')):
        return True
    return len(plain) > 30


def _is_heading(text):
    text = text.strip()
    if not text or len(text) > 40:
        return False
    return bool(HEADING_RE.match(text) or CN_HEADING_RE.match(text)
                or NUMBERED_RE.match(text))


def _heading_text(text):
    text = re.sub(r'^[（(]?[一二三四五六七八九十百\d]+[）)、\.．]\s*', '', text.strip())
    return _short_label(text)


def _label_before(text, start, floor=0):
    seg = text[floor:start].replace('\u3000', ' ')
    seg = re.split(r'[|｜]', seg)[-1]
    seg = re.split(r'\s{2,}', seg)[-1].strip()
    seg = re.sub(r'[：:]\s*$', '', seg)
    if '：' in seg or ':' in seg:
        seg = re.split(r'[：:]', seg)[-1]
    return _short_label(seg)


def _infer_type(label, sample='', multiline=False):
    if multiline or any(word in (label or '') for word in MULTI_WORDS):
        return 'multiline'
    if len((sample or '').replace('_', '').strip()) > 25:
        return 'multiline'
    return 'text'


def _group(label, heading):
    group = formparse._group_for_single(label, '')
    if group == '填写内容' and heading:
        return heading
    return group


def _date_match(text):
    for match in DATE_RE.finditer(text or ''):
        if re.search(r'[_\u3000]', match.group(0)):
            return match
    return None


def _is_placeholder(content, paren=False):
    content = content or ''
    if PLACEHOLDER_ONLY.match(content):
        return True
    if paren:
        if len(content) > 12:
            return False
        if re.search(r'(由.+)?(填写|输入)$|说明|注意|以下|上述|本人', content):
            return False
        return bool(re.search(r'请|输入|填写|此处|在此|点击|选择', content))
    return bool(PLACEHOLDER_WORDS.search(content))


def _bracket_label(content, text, start):
    label = _clean(content or '')
    label = re.sub(r'^(?:请|请在此|在此|此处)?(?:填写|输入)', '', label)
    label = re.sub(r'(?:填写|输入)(?:处|栏|框)?$', '', label)
    label = _short_label(label)
    if label and re.search(r'[\u4e00-\u9fff]{2,}', label):
        return label
    return _label_before(text, start) or '填写项'


def _make_field(fields, label, type_, group, where, context, original, confidence,
                source, location, bookmark_range=None, control=None):
    fields.append({
        'plan_id': '', 'key': '', 'label': label, 'type': type_,
        'required': False, 'group': group, 'where': where, 'context': context,
        'original_content': (original or '')[:100], 'confidence': confidence,
        'source': source, 'location': location,
        'bookmark_range': bookmark_range, 'bookmark': '',
        'control': control,
    })


def _covered(existing, p, start, end):
    for field in existing:
        rng = field.get('bookmark_range') or {}
        if not rng:
            continue
        if rng.get('start_p', 0) <= p <= rng.get('end_p', 0):
            if rng.get('start_p') == p and rng.get('end_p') == p:
                if not (start < rng.get('end', 0) and end > rng.get('start', 0)):
                    continue
            return True
    return False


def _prev_text(paragraphs, index):
    for q in range(index - 1, -1, -1):
        value = _text(paragraphs[q]).strip()
        if value:
            return value
    return ''


def _prev_label(prev):
    if not prev or NOTE_RE.search(prev):
        return ''
    label = _short_label(prev)
    if not label or _noise(label) or _is_heading(prev):
        return ''
    if len(label) > 14:
        return ''
    return label


def _scan_list(paragraphs, fields, container='body', part_name='body',
               table=None, row=None, col=None, existing=()):
    heading = ''
    used = {}
    limit = len(paragraphs)

    def where_for(index):
        if container == 'cell':
            return f'表{table + 1} 行{row + 1} 列{col + 1} · 第{index + 1}段'
        name = {'header': '页眉', 'footer': '页脚'}.get(container, '正文')
        return f'{name} 第{index + 1}段'

    def add(index, start, end, label, type_, sample, confidence='高', source='空白位',
            start_p=None, end_p=None):
        label = _short_label(label)
        if not label or _noise(label):
            return False
        start_p = index if start_p is None else start_p
        end_p = index if end_p is None else end_p
        new_size = (end_p - start_p) * 10000 + (end - start)
        for old in list(fields):
            if old.get('label') != label:
                continue
            loc = old.get('location') or {}
            old_rng = old.get('bookmark_range') or {}
            if loc.get('kind') != container:
                continue
            if container == 'cell':
                if (loc.get('table'), loc.get('row'), loc.get('col')) != (table, row, col):
                    continue
            elif old_rng.get('start_p', -1) != start_p:
                continue
            old_size = ((old_rng.get('end_p', 0) - old_rng.get('start_p', 0)) * 10000
                        + (old_rng.get('end', 0) - old_rng.get('start', 0)))
            if old.get('source') == '表单结构' and old_size <= 0 < new_size:
                fields.remove(old)
                break
            return False
        spans = used.setdefault(start_p, [])
        if start_p == end_p and any(start < e and end > s for s, e in spans):
            return False
        if existing and _covered(existing, start_p, start, end):
            return False
        loc = {'kind': container, 'p': index}
        if container == 'cell':
            loc.update(table=table, row=row, col=col)
        elif container in ('header', 'footer'):
            loc['part'] = part_name
        rng = {'container': 'cell' if container == 'cell' else container,
               'part': part_name, 'start_p': start_p, 'start': start,
               'end_p': end_p, 'end': end}
        if container == 'cell':
            rng.update(table=table, row=row, col=col)
        _make_field(fields, label, type_, _group(label, heading), where_for(index),
                    _text(paragraphs[index])[:120], sample, confidence, source,
                    loc, rng)
        spans.append((start, end))
        return True

    for i, p in enumerate(paragraphs):
        text = _text(p)
        if not text.strip():
            continue
        stripped = text.strip()
        if NOTE_RE.search(stripped) or formparse._is_note(p):
            continue
        if _is_heading(stripped):
            heading = _heading_text(stripped)

        # date placeholders are one field, not three
        date = _date_match(text)
        if date and not (existing and _covered(existing, i, date.start(), date.end())):
            label = _label_before(text, date.start()) or '日期'
            add(i, date.start(), date.end(), label, 'text', date.group(0), '高', '日期')

        # a line made only of underscores is a cover/line field
        if LINE_RE.match(stripped) and len(stripped) >= 4 and re.search(r'[_＿﹍…]', stripped):
            label = _prev_label(_prev_text(paragraphs, i))
            if label:
                start = len(text) - len(text.lstrip())
                end = len(text.rstrip())
                add(i, start, end, label, _infer_type(label, '', len(stripped) > 30),
                    stripped, '中', '下划线')
            continue

        # bracketed prompts
        for match in BRACKET_RE.finditer(text):
            group = next((gi for gi in range(1, 6) if match.group(gi) is not None), None)
            if group is None:
                continue
            content = match.group(group)
            if not _is_placeholder(content, paren=group >= 4):
                continue
            if any(word in content for word in SKIP_LABELS):
                continue
            start, end = match.start(), match.end()
            label = _bracket_label(content, text, start)
            add(i, start, end, label, _infer_type(label, content), content, '高', '占位提示')

        # underline / space / dot blanks
        for match in BLANK_RE.finditer(text):
            start, end = match.start(), match.end()
            if date and date.start() <= start < date.end():
                continue
            if any(start < e and end > s for s, e in used.get(i, [])):
                continue
            if not match.group(0).strip():
                before = text[:start]
                after = text[end:]
                if not (re.search(r'[：:]\s*$', before)
                        or re.match(r'\s*[\u4e00-\u9fff]{1,10}[：:]', after)):
                    continue
            label = _label_before(text, start)
            if not label and len(stripped) <= 24 and not stripped.endswith(('：', ':')):
                label = _short_label(stripped)
            add(i, start, end, label, _infer_type(label, match.group(0)),
                match.group(0), '高', '空白位')

        # "标签：" followed by empty line(s) or an underline line
        if stripped.endswith(('：', ':')):
            label = _short_label(stripped[:-1])
            if not label or _noise(label):
                continue
            j = i + 1
            while j < limit and not _text(paragraphs[j]).strip():
                j += 1
            if j > i + 1:
                add(i + 1, 0, 0, label, _infer_type(label, '', True), '',
                    '中', '标签空行', start_p=i + 1, end_p=j - 1)
            elif j < limit:
                nxt = _text(paragraphs[j]).strip()
                if LINE_RE.match(nxt) and re.search(r'[_＿﹍…]', nxt):
                    start = len(_text(paragraphs[j])) - len(_text(paragraphs[j]).lstrip())
                    end = len(_text(paragraphs[j]).rstrip())
                    add(j, start, end, label, _infer_type(label, nxt), nxt, '中', '标签下划线')
    return heading


def _scan_tables(doc, fields):
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            seen = set()
            cells = []
            for ci, cell in enumerate(row.cells):
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                cells.append((ci, cell))
            field_cols = set()
            for field in fields:
                loc = field.get('location') or {}
                if loc.get('kind') == 'cell' and \
                        (loc.get('table'), loc.get('row')) == (ti, ri):
                    field_cols.add(loc.get('col'))
            for ci, cell in cells:
                if ci in field_cols or ci + 1 in field_cols:
                    continue
                _scan_list(list(cell.paragraphs), fields, 'cell', 'body', ti, ri, ci)


def _drop_covered_table_fields(fields, controls):
    drop = set()
    for ctrl in controls:
        if ctrl.get('table') is None:
            continue
        key = (ctrl['table'], ctrl['row'], ctrl['col'])
        for field in fields:
            loc = field.get('location') or {}
            if loc.get('kind') == 'cell' and \
                    (loc.get('table'), loc.get('row'), loc.get('col')) == key:
                drop.add(id(field))
    return [f for f in fields if id(f) not in drop]


def _control_fields(doc):
    raw = controlfill.list_controls(doc) + controlfill.list_formfields(doc)
    extra = []
    for ctrl in raw:
        alias = _clean(ctrl.get('alias') or '')
        if any(word in alias for word in CONTROL_NOISE):
            continue
        label = alias or f'控件{len(extra) + 1}'
        original = ctrl.get('text') or ''
        type_ = 'multiline' if ctrl.get('multiline') else 'text'
        base = ('页眉' if ctrl.get('part') == 'header' else
                '页脚' if ctrl.get('part') == 'footer' else '正文')
        where = base + ('表单域' if ctrl.get('kind') == 'formfield' else '内容控件')
        loc = {'kind': ctrl.get('kind'), 'part': ctrl.get('part'),
               'part_name': ctrl.get('part_name'), 'index': ctrl.get('index'),
               'p': ctrl.get('p', 0), 'table': ctrl.get('table'),
               'row': ctrl.get('row'), 'col': ctrl.get('col')}
        _make_field(extra, label, type_, _group(label, ''), where,
                    original[:120], original, '高',
                    '内容控件' if ctrl.get('kind') == 'control' else '表单域',
                    loc, None,
                    {'kind': ctrl.get('kind'), 'part_name': ctrl.get('part_name'),
                     'index': ctrl.get('index')})
    return extra, raw


def _renumber(fields):
    seen = set()
    for i, field in enumerate(fields, 1):
        field['plan_id'] = f'sp{i}'
        if not field.get('control'):
            field['bookmark'] = f'WTT5_{i:03d}'
        base = re.sub(r'[^\w\u4e00-\u9fff]+', '', field.get('label') or '')[:32] or f'field{i}'
        base = base.strip('_') or f'field{i}'
        key, n = base, 1
        while key in seen:
            n += 1
            key = f'{base}{n}'
        seen.add(key)
        field['key'] = key


def _sort_key(field):
    loc = field.get('location') or {}
    kind = loc.get('kind')
    if kind == 'cell':
        rng = field.get('bookmark_range') or {}
        return (0, loc.get('table', 0), loc.get('row', 0), loc.get('col', 0),
                rng.get('start_p', 0), 0)
    if kind == 'body':
        return (0, -1, -1, -1, loc.get('p', 0), 0)
    return (1, 0, 0, 0, loc.get('p', 0), loc.get('index', 0))


def analyze(path):
    path = str(path)
    doc = Document(path)
    base = None
    try:
        base = formparse.analyze(path)
    except Exception:
        base = None
    fields = list(base['fields']) if base else []
    for field in fields:
        field.setdefault('source', '表单结构')
        field.setdefault('confidence', '高')

    extra, controls = _control_fields(doc)
    if controls:
        fields = _drop_covered_table_fields(fields, controls)
    fields += extra

    _scan_list(list(doc.paragraphs), fields)
    for kind, part_name, root in bookmarkfill.story_parts(doc):
        if kind == 'body':
            continue
        _scan_list(bookmarkfill._paragraphs_of_root(root), fields, kind, part_name)
    _scan_tables(doc, fields)

    fields = [f for f in fields if f.get('control') or f.get('bookmark_range')]
    if not fields:
        raise ValueError(
            '没有识别到可填写栏目。\n'
            '这份文档可能只是规范/说明材料，或填写位置没有可识别的标签、下划线、'
            '括号提示、内容控件或空白单元格。')

    for field in fields:
        field['required'] = False
    fields.sort(key=_sort_key)
    _renumber(fields)
    title = (base or {}).get('title') or next(
        (p.text.strip() for p in doc.paragraphs if p.text.strip()), '未命名模板')
    category = (base or {}).get('category') or (
        '教学' if any(k in title for k in ('任务书', '开题', '论文', '学院', '学校')) else '自定义模板')
    return {
        'source': path, 'title': title[:60], 'category': category,
        'description': (f'自动识别 {len(fields)} 个填写项；未使用任何占位符，'
                        '原文档文字、示例、颜色和版式保持不变。'),
        'table': 0, 'fields': fields, 'recipes': [], 'clear_cells': [],
        'row_heights': {}, 'remove_textboxes': False, 'smart': True,
    }
