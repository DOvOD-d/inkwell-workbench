"""Placeholder-free recognition.

The document is the source of truth: label/value cells, section headings,
underlines, bracket prompts, content controls and legacy form fields are all
read directly from OOXML.  No ``{{ }}`` token is required or produced.
"""
from __future__ import annotations

import re

from docx import Document

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

# 规范/示例材料里的参考文献条目长得像“标签：内容”，其实一个字都不用填。
CITATION_RE = re.compile(r'\[(?:J|M|C|D|S|R|N|P|A|G|DB|EB|OL|MT|CP)[,:\]]', re.IGNORECASE)
# 真正的栏目名是短名词；带 * [ ] 、句尾标点或动词的都是示例文字。
LABEL_BAD_RE = re.compile(r'[\[\]【】＊*]|\.{3,}|·{2,}|[。.;；]$')
LABEL_VERB_RE = re.compile(r'选择|填写|请输入|请|应当|应|必须|须|不得|不要|可以|需要|注意|例如|以下|上述')

# 同一份表里反复出现的基础信息（封面、任务书部分、开题部分）本来就是同一个值，
# 归并成一个填写位，填一次写到所有位置；显示名统一用最完整的那种写法。
BASIC_ALIASES = {
    '姓名': '学生姓名', '学生': '学生姓名', '学生姓名': '学生姓名', '学生名字': '学生姓名',
    '名字': '学生姓名', '本人姓名': '学生姓名',
    '学号': '学号', '学生学号': '学号',
    '学院': '学院', '院系': '学院', '所在学院': '学院', '所属学院': '学院',
    '院（系）': '学院', '系（院）': '学院', '学院名称': '学院',
    '专业': '专业', '所学专业': '专业', '学生专业': '专业',
    '班级': '班级', '班次': '班级',
    '届次': '届次', '届别': '届次', '毕业届次': '届次',
    '指导教师': '指导教师', '指导老师': '指导教师', '指导教师姓名': '指导教师',
    '教师姓名': '指导教师',    '中文题目': '中文题目', '论文题目（中文）': '中文题目', '题目（中文）': '中文题目',
    '设计题目（中文）': '中文题目', '论文（设计）题目（中文）': '中文题目',
    '外文题目': '外文题目', '英文题目': '外文题目', '题目（外文）': '外文题目',
    '题目（英文）': '外文题目', '论文题目（外文）': '外文题目',
    # “职称”和“学位”单独出现时是两个栏目，不能并成一个；只归并本来就写在一起的写法。
    '职称/学位': '职称/学位', '职称（学位）': '职称/学位', '职务/职称': '职称/学位',
    '电话': '联系电话', '联系电话': '联系电话', '手机': '联系电话', '手机号码': '联系电话',
    '邮箱': '电子邮箱', '电子邮箱': '电子邮箱', 'email': '电子邮箱',
}


def _plain(label):
    return re.sub(r'[\s\u3000]', '', label or '')


def canonical_label(label):
    """把“学生 / 姓名 / 学生姓名”这类同义写法归并成同一个栏目名。"""
    plain = _plain(label)
    if not plain:
        return (label or '').strip()
    mapped = BASIC_ALIASES.get(plain) or BASIC_ALIASES.get(plain.lower())
    if mapped:
        return mapped
    stripped = re.sub(r'[（）()】［\[\]]', '', plain)
    mapped = BASIC_ALIASES.get(stripped) or BASIC_ALIASES.get(stripped.lower())
    return mapped or (label or '').strip()


def _label_ok(label, container):
    plain = _clean(label)
    if not plain or LABEL_BAD_RE.search(plain) or LABEL_VERB_RE.search(plain):
        return False
    if len(re.findall(r'[\u4e00-\u9fff]', plain)) < 2:
        return False
    if container != 'cell' and len(plain) > 14:
        return False
    return True


def _text(p):
    return ''.join(r.text for r in p.runs)


def _clean(text):
    text = re.sub(r'[\s\u3000]+', '', text or '')
    text = re.sub(r'[：:；;,，。、\.]+$', '', text)
    # 只有整段被括号包住时才去掉括号；“经费总额（元）”里的括号是名字的一部分
    if text[:1] in '（(【[［' and text[-1:] in '）)】]］':
        inner = text[1:-1]
        if inner and not any(ch in inner for ch in '（(【[［'):
            text = inner
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
        if not label or _noise(label) or not _label_ok(label, container):
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
        if container != 'cell' and CITATION_RE.search(stripped):
            continue          # 规范材料里的参考文献示例行不是栏目

        # 模板自己写的 {{ 记号 }} 最权威：记号名就是栏目名，一段里可以有多个
        tokens = list(TOKEN_RE.finditer(text))
        if tokens:
            for match in tokens:
                name = _clean(match.group(1))
                if not name:
                    continue
                add(i, match.start(), match.end(), name, _infer_type(name, ''),
                    match.group(0), '高', '占位符')
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
                # 下一段是参考文献示例（规范材料里的“期刊论文：作者.题名[J]…”），不是要填的空
                if j < limit and CITATION_RE.search(_text(paragraphs[j])):
                    continue
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
                # 合并单元格会重复出现：比较元素本身，id() 在包装对象回收后可能复用
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
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


TICK_RE = re.compile(r'([^\s（）()]{1,14})[（(][\s\u3000]*[）)]')


def _tick_choices(text):
    """“自选（ ）备选（ ）”这种勾选栏：把每个括号打上 √，做成下拉选项。"""
    matches = list(TICK_RE.finditer(text or ''))
    if len(matches) < 2:
        return []
    if any(len(match.group(1)) < 2 for match in matches):
        return []
    options = []
    for index in range(len(matches)):
        parts, cursor = [], 0
        for position, match in enumerate(matches):
            parts.append(text[cursor:match.start(1)])
            parts.append(match.group(1))
            parts.append('（√）' if position == index else '（ ）')
            cursor = match.end()
        parts.append(text[cursor:])
        options.append(''.join(parts).strip())
    return options


def _as_select(field):
    choices = _tick_choices(field.get('original_content') or '')
    if not choices or field.get('type') != 'text':
        return
    original = field['original_content']
    field['type'] = 'select'
    field['options'] = [{'label': option, 'value': option} for option in choices]
    if original not in choices:
        field['options'].insert(0, {'label': original + '（不勾）', 'value': original})
    field['where'] = (field.get('where') or '') + ' · 勾选项'


def _raw_key(label, index):
    base = (re.sub(r'[^\w\u4e00-\u9fff]+', '', label or '')[:32]).strip('_')
    return base or f'field{index}'


TOKEN_RE = re.compile(r'\{\{\s*([^{}]+?)\s*\}\}')
NOISE_CONTAINS = ('公章', '盖章', '签字', '签名', '手签', '填表说明', '注意事项')


def _strip_token(label):
    """{{专家1姓名}} 这类记号：栏目名就是花括号里的内容。"""
    text = (label or '').strip()
    inner = TOKEN_RE.findall(text)
    if not inner:
        return text
    return inner[0].strip() or text


def _digitless(label):
    return re.sub(r'\d+', '', canonical_label(label))


def _noise_label(label):
    plain = _clean(label)
    return bool(plain) and any(word in plain for word in NOISE_CONTAINS)


def _drop_token_guesses(doc, fields):
    """带 {{ 记号 }} 的位置由记号名命名，丢掉结构猜测出来的同位置字段。"""
    keep = []
    for field in fields:
        loc = field.get('location') or {}
        if loc.get('kind') == 'cell' and loc.get('table') is not None:
            try:
                cell = doc.tables[loc['table']].rows[loc['row']].cells[loc['col']]
            except (IndexError, KeyError, AttributeError):
                keep.append(field)
                continue
            if TOKEN_RE.search(cell.text or ''):
                continue
        keep.append(field)
    return keep


def tidy_labels(fields):
    """清掉花括号记号、去掉公章/签字类噪声，并把同一格里重名的栏目按段编号。"""
    kept = []
    for field in fields:
        label = _strip_token(field.get('label') or '')
        if label and not field.get('control') and (_noise(label) or _noise_label(label)):
            continue
        field['label'] = label or (field.get('label') or '')
        kept.append(field)
    fields[:] = kept
    by_cell = {}
    for field in fields:
        loc = field.get('location') or {}
        if loc.get('kind') != 'cell':
            continue
        by_cell.setdefault((loc.get('table'), loc.get('row'), loc.get('col')), []).append(field)
    for group in by_cell.values():
        same = {}
        for field in group:
            same.setdefault(field['label'], []).append(field)
        for label, items in same.items():
            if len(items) < 2:
                continue
            for index, field in enumerate(items, 1):
                field['label'] = f'{label}·第{index}段'


def _common_suffix(labels):
    """一列各行栏目名的公共尾巴就是列标题：专家1姓名/专家2姓名 → 姓名。"""
    if not labels:
        return ''
    shortest = min(labels, key=len)
    for size in range(len(shortest), 1, -1):
        tail = shortest[-size:]
        if all(label.endswith(tail) for label in labels):
            return tail
    return shortest


def _runs(rows):
    groups, current = [], []
    for row in rows:
        if current and row != current[-1] + 1:
            groups.append(current)
            current = []
        current.append(row)
    if current:
        groups.append(current)
    return groups


def detect_grids(fields):
    """把“结构相同的连续多行”（专家×5、成员×7、经费×6）改成表格格填写。

    判据：同一张表里连续 ≥3 行，每行的栏目“族”（去掉行号后的名字）集合完全一样，
    且每行至少两个栏目。命中的格子各自独立成键，不再被当成同名栏目合并。
    返回的 rows 先放字段对象，等键名分配完再换成 key。
    """
    by_row = {}
    for field in fields:
        loc = field.get('location') or {}
        if loc.get('kind') != 'cell' or field.get('control') or field.get('grid'):
            continue
        by_row.setdefault((loc.get('table'), loc.get('row')), []).append(field)
    shapes = {}
    for (table, row), items in by_row.items():
        families = tuple(sorted({_digitless(f['label']) for f in items}))
        if len(items) < 2 or not all(families):
            continue
        shapes.setdefault(table, {}).setdefault(families, []).append(row)
    grids = []
    for table, groups in shapes.items():
        for rows in groups.values():
            for run in _runs(sorted(rows)):
                if len(run) < 3:
                    continue
                block = [f for row in run for f in by_row[(table, row)]]
                columns = {}
                for field in block:
                    col = (field.get('location') or {}).get('col', 0)
                    columns.setdefault(_digitless(field['label']), []).append((col, field))
                if len(columns) < 2:
                    continue
                ordered = sorted(columns.items(), key=lambda item: min(c for c, _f in item[1]))
                titles = []
                for family, group in ordered:
                    labels = [f['label'] for _c, f in group]
                    titles.append(family if len(set(labels)) == 1 else _common_suffix(labels))
                # 紧挨在块上面的那一行往往是表头（姓 名 / 所属单位 / 职务职称），
                # 表头文字就是列标题，它本身不是要填的栏目。
                header = [f for f in by_row.get((table, run[0] - 1), [])
                          if (f.get('location') or {}).get('col') is not None]
                header.sort(key=lambda f: (f.get('location') or {}).get('col', 0))
                if (len(header) == len(ordered)
                        and all(len(f['label']) <= 12 for f in header)):
                    titles = [f['label'] for f in header]
                # 块上面那行里，凡是列标题或被整块重复引用的合并标签，都不是要填的栏目
                repeated = {f['label'] for row in run for f in by_row[(table, row)]}
                for field in by_row.get((table, run[0] - 1), []):
                    if field['label'] in set(titles) or field['label'] in repeated:
                        field['header'] = True
                title = f'{titles[0]}（按原表保留 {len(run)} 行）'
                matrix = []
                for number, row in enumerate(run, 1):
                    line = []
                    for (family, group), column_title in zip(ordered, titles):
                        field = next((f for _c, f in group
                                      if (f.get('location') or {}).get('row') == row), None)
                        if field is None:
                            line.append(None)
                            continue
                        if len({f['label'] for _c, f in group}) == 1:
                            field['label'] = f'{family}{number}'
                        field['grid'] = title
                        field['group'] = titles[0]
                        line.append(field)
                    matrix.append(line)
                grids.append({'group': titles[0], 'title': title, 'columns': titles,
                              'rows': matrix, 'table': table})
    return grids


def group_by_label(fields):
    """同一份表里出现多处的基础信息并成一个栏目；只出现一处的保持原样。

    归并判据用同义写法表（学生 / 姓名 / 学生姓名 是同一个值），
    但只有真的多于一个位置时才改 key 和显示名，避免把 “姓名” 莫名改成 “学生姓名”。
    """
    canonical = {}
    for i, field in enumerate(fields, 1):
        if field.get('grid'):
            field['key'] = _raw_key(field.get('label') or '', i)
            field.setdefault('slots', 1)
            continue          # 表格格每格独立成键，绝不参与同名合并
        label = field.get('label') or ''
        canon = _raw_key(canonical_label(label), i)
        canonical.setdefault(canon, []).append((i, field))
    changed = False
    for canon, members in canonical.items():
        if len(members) > 1:
            # 统一叫法用这一组里最完整的写法：学生 + 学生姓名 → 学生姓名；
            # 只有 “姓名” 时不会凭空叫成 “学生姓名”。
            unified = max((f.get('label') or '' for _i, f in members), key=len) or canon
            existing = {f.get('key') for _i, f in members if f.get('key')}
            if len(existing) == 1:
                key = existing.pop()       # 本来就共用一个键，保持不动（{{记号}} 模板尤其不能改）
            else:
                key = _raw_key(unified, members[0][0])
            multiline = any(f.get('type') == 'multiline' for _i, f in members)
            for index, field in members:
                if field.get('key') != key:
                    field['key'] = key
                    changed = True
                if field.get('label') != unified:
                    field['label'] = unified
                    changed = True
                if multiline and field.get('type') != 'multiline':
                    field['type'] = 'multiline'
                    changed = True
                field['slots'] = len(members)
        else:
            index, field = members[0]
            if not field.get('key'):
                field['key'] = _raw_key(field.get('label') or '', index)
            field.setdefault('slots', 1)
    return changed


def _renumber(fields):
    """槽位编号 + 书签名唯一 + 同名栏目共用 key。"""
    for i, field in enumerate(fields, 1):
        field['plan_id'] = f'sp{i}'
        if not field.get('control'):
            field['bookmark'] = f'WTT5_{i:03d}'
    group_by_label(fields)


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


def _spec_document(doc, fields):
    """规范/说明材料：没有任何表格空白格或控件，正文全是“例如：/式样/著录”式示例。"""
    if any((field.get('location') or {}).get('kind') == 'cell' for field in fields):
        return False
    if any(field.get('control') for field in fields):
        return False
    text = '\n'.join(_text(p) for p in doc.paragraphs)
    return len(re.findall(r'例如[:：]|示例|式样|著录|注意事项', text)) >= 3


HEADING_LABEL_RE = re.compile(
    r'^(?:[一二三四五六七八九十]+[、.．]|[（(][一二三四五六七八九十]+[)）])')
NUMBERED_LABEL_RE = re.compile(r'^(?:\d+[、.．]|\d+[)）])')
IMAGE_SLOT_RE = re.compile(r'截图|图片|照片')


def _placeholder_only(text):
    return bool(PLACEHOLDER_ONLY.match(text or '')) or bool(
        BLANK_RE.search(text or '') and not re.search(r'[\u4e00-\u9fffA-Za-z0-9]{2,}', text))


def _drop_unsafe(fields):
    """两条安全线，宁可少识别也不出错：

    1. “一、实验目的”“二、实验设备及环境”这类小节标题被当成栏目时，它的写入范围会
       覆盖整段正文，一旦生成就把原文示例顶掉。标题位只要已有实质内容就一律不要。
    2. “第1题运行结果的截图如下：”这种只留插入点、没有任何文字空位的图片占位，
       不是文本栏目；本工具也不往这些位置插图。
    """
    keep = []
    for field in fields:
        label = field.get('label') or ''
        rng = field.get('bookmark_range') or {}
        content = (field.get('original_content') or '').strip()
        token = '{{' in content
        if not field.get('control') and not token:
            section = bool(HEADING_LABEL_RE.match(label))
            numbered = bool(NUMBERED_LABEL_RE.match(label))
            if section and content and not _placeholder_only(content):
                continue
            if (numbered and content and len(content) > 12
                    and rng.get('start_p', 0) != rng.get('end_p', 0)):
                continue
        if (IMAGE_SLOT_RE.search(label) and not content and not field.get('control')
                and rng.get('start_p') == rng.get('end_p')
                and rng.get('start', 0) == rng.get('end', 0)):
            continue
        keep.append(field)
    return keep


def analyze(path):
    path = str(path)
    doc = Document(path)
    base = None
    try:
        base = formparse.analyze(path)
    except Exception:
        base = None
    fields = list(base['fields']) if base else []
    if fields:
        fields = _drop_token_guesses(doc, fields)
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
    fields = _drop_unsafe(fields)
    if fields and _spec_document(doc, fields):
        raise ValueError(
            '这份文档是规范/说明材料：没有表格空白格、内容控件或下划线空位，'
            '正文里的“例如：”“式样”都是示例条目，不是要填的栏目。\n'
            '原文件不会被改动，也不会录入模板库。')
    if not fields:
        raise ValueError(
            '没有识别到可填写栏目。\n'
            '这份文档可能只是规范/说明材料，或填写位置没有可识别的标签、下划线、'
            '括号提示、内容控件或空白单元格。')

    tidy_labels(fields)
    grids = detect_grids(fields)
    fields[:] = [field for field in fields if not field.get('header')]
    for field in fields:
        field['required'] = False
        _as_select(field)
    fields.sort(key=_sort_key)
    _renumber(fields)
    for grid in grids:
        grid['rows'] = [[field['key'] if field else '' for field in row] for row in grid['rows']]
    title = (base or {}).get('title') or next(
        (p.text.strip() for p in doc.paragraphs if p.text.strip()), '未命名模板')
    category = (base or {}).get('category') or (
        '教学' if any(k in title for k in ('任务书', '开题', '论文', '学院', '学校')) else '自定义模板')
    return {
        'source': path, 'title': title[:60], 'category': category,
        'description': (f'自动识别 {len(fields)} 个填写项；未使用任何占位符，'
                        '原文档文字、示例、颜色和版式保持不变。'),
        'table': 0, 'fields': fields, 'recipes': [], 'clear_cells': [],
        'grids': grids,
        'row_heights': {}, 'remove_textboxes': False, 'smart': True,
    }
