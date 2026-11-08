# -*- coding: utf-8 -*-
"""文本模型识别：把整份文档的文字与结构交给文本模型，判断哪些位置要填写。

为什么不是截图识别：截图把版面压成像素，表格线、合并单元格、下划线都得靠模型“看”，
错一个就定位不到；而 .docx 本身就带着段落、单元格、下划线和占位符，把这份结构直接
交给文本模型，模型只需要判断“哪一格是空着的、该叫什么”，定位由程序按原文精确回填。

分工固定：
  smartparse  决定“位置”：所有锚点都从 OOXML 里来，模型给不出位置就不会被采纳；
  textparse   决定“叫什么/什么类型/怎么分组”：只改标签与属性，不改位置；
  任何模型输出先过 _validate，锚点不存在、标签不在原文、单元格里已有成段正文，一律丢弃。
"""
from __future__ import annotations

import copy
import re
from docx import Document

import ai
import feedback
import smartparse
import visionparse

OUTLINE_MAX_CHARS = 3800    # 单次请求的结构清单上限：留足输出空间，避免 JSON 被截断
MAX_FIELDS = 120            # 单次请求最多接受的栏目数
MAX_CHUNKS = 10             # 最多问几次；再大的文档只增强前 10 段，其余靠结构识别
SLOT_BUDGET = 24            # 每段最多包含多少个空位：按输出量切，避免 JSON 被截断
MIN_SPLIT_LINES = 12        # 失败二分到这么少就不再拆
MAX_OUTPUT_TOKENS = 2600
CLEAN_TYPE = {'text', 'multiline', 'date', 'number', 'select'}
TYPE_ALIASES = {'单行': 'text', '多行': 'multiline', '日期': 'date',
                '数字': 'number', '选择': 'select'}
CONFIDENCE_ALIASES = {'高': '高', '中': '中', '低': '低',
                      'high': '高', 'medium': '中', 'low': '低'}
GENERIC_GROUPS = {'', '填写内容', '自定义模板'}
MAX_LABEL = 24

TOKEN_RE = re.compile(r'\{\{\s*([^{}]{1,40}?)\s*\}\}')
BLANK_RE = re.compile(r'[_＿﹍﹏]{2,}|…{2,}|\.{4,}|·{3,}')
NOISE_LABEL = re.compile(r'(签字|签名|盖章|公章|手签|示例|页码|目录|页眉|页脚)')
NOISE_EXACT = {'说明', '填写说明', '填表说明', '注意事项', '格式要求'}
PROMPT_RE = re.compile(r'【[^】]*?(?:请填写|请输入|填写|输入)[^】]*?】'
                       r'|[（(][^）)]*?(?:请填写|请输入|在此输入|此处填写)[^）)]*?[）)]')


def _wire(text):
    """比较用：去掉所有空白，中文标签常被排版成“姓    名”。"""
    return re.sub(r'[\s\u3000]+', '', text or '')


MARKER_PREFIX = re.compile(r'^(?:占位符|占位及|占位|空单元格|下划线|空位)[:：]?')


def _clean_label(text):
    label = re.sub(r'[\s\u3000]+', '', str(text or ''))
    label = MARKER_PREFIX.sub('', label)
    label = re.sub(r'[：:；;,，。、\.]+$', '', label)
    return label[:MAX_LABEL]


def _has_slot(text):
    return bool(TOKEN_RE.search(text or '') or BLANK_RE.search(text or ''))


def _underline(p):
    return any(r.underline for r in p.runs if r.text.strip())


# ---------------------------------------------------------------- 结构清单
def _paragraph_item(p, index):
    text = p.text.strip()
    if not text:
        return None
    item = {'p': index, 'text': text}
    if _underline(p):
        item['underline'] = True
    if BLANK_RE.search(text):
        item['blank'] = True
    tokens = TOKEN_RE.findall(text)
    if tokens:
        item['tokens'] = [t.strip() for t in tokens]
    style = p.style.name if p.style is not None else ''
    if style and style != 'Normal':
        item['style'] = style
    return item


def _cell_item(cell, row, col):
    texts, tokens, blank, underlined = [], [], False, False
    for p in cell.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        texts.append(text)
        tokens.extend(t.strip() for t in TOKEN_RE.findall(text))
        blank = blank or bool(BLANK_RE.search(text))
        underlined = underlined or _underline(p)
    item = {'r': row, 'c': col, 'text': ' '.join(texts)}
    if tokens:
        item['tokens'] = tokens
    if blank:
        item['blank'] = True
    if underlined:
        item['underline'] = True
    return item


def build_outline(doc):
    paragraphs = [item for item in (_paragraph_item(p, i) for i, p in enumerate(doc.paragraphs))
                  if item]
    tables = []
    for ti, table in enumerate(doc.tables):
        rows = []
        for ri, row in enumerate(table.rows):
            seen, cells = set(), []
            for ci, cell in enumerate(row.cells):
                # 合并单元格在 row.cells 里会重复出现；比较元素本身而不是 id()，
                # 否则临时包装对象被回收后 id 可能复用，漏判成两个格子。
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                cells.append(_cell_item(cell, ri, ci))
            rows.append(cells)
        tables.append({'t': ti, 'rows': rows})
    return {'paragraphs': paragraphs, 'tables': tables}


def _clip(text, limit=240):
    flat = re.sub(r'\s+', ' ', text or '').strip()
    return flat if len(flat) <= limit else flat[:limit] + '…'


def _render_block(paragraphs, table=None):
    lines = []
    if table is None:
        for item in paragraphs:
            marks = []
            if item.get('underline'):
                marks.append('下划线')
            if item.get('blank'):
                marks.append('空位')
            if item.get('tokens'):
                marks.append('占位符:' + ','.join(item['tokens']))
            if item.get('style'):
                marks.append(item['style'])
            head = 'p%d' % item['p'] + (' [%s]' % ' '.join(marks) if marks else '')
            lines.append(f'{head} :: {_clip(item["text"])}')
        return lines
    lines.append(f'=== 表{table["t"]} ===')
    for cells in table['rows']:
        for item in cells:
            marks = []
            if not item['text']:
                marks.append('空单元格')
            if item.get('underline'):
                marks.append('下划线')
            if item.get('blank'):
                marks.append('空位')
            if item.get('tokens'):
                marks.append('占位符:' + ','.join(item['tokens']))
            head = f't{table["t"]} r{item["r"]} c{item["c"]}'
            if marks:
                head += ' [' + ' '.join(marks) + ']'
            lines.append(f'{head} :: {_clip(item["text"])}' if item['text'] else head)
    return lines


def _chunks(outline):
    """把清单切成若干段，保证每段不超过预算；表格整张优先，太大的表按行切。"""
    blocks = []
    pending = []
    pending_size = 0

    def flush():
        nonlocal pending, pending_size
        if pending:
            blocks.append(_render_block(pending))
            pending, pending_size = [], 0

    for item in outline['paragraphs']:
        size = len(item['text']) + 20
        if pending and pending_size + size > OUTLINE_MAX_CHARS:
            flush()
        pending.append(item)
        pending_size += size
    flush()

    for table in outline['tables']:
        rows = table['rows']
        batch, size = [], 0
        for cells in rows:
            row_size = sum(len(c['text']) + 24 for c in cells)
            if batch and size + row_size > OUTLINE_MAX_CHARS:
                blocks.append(_render_block([], {'t': table['t'], 'rows': batch}))
                batch, size = [], 0
            batch.append(cells)
            size += row_size
        if batch:
            blocks.append(_render_block([], {'t': table['t'], 'rows': batch}))

    chunks, current, size, slots = [], [], 0, 0
    for block in blocks:
        for line in block:
            line_size = len(line) + 1
            line_slots = 1 if any(mark in line for mark in
                                  ('空单元格', '下划线', '空位', '占位符')) else 0
            if current and (size + line_size > OUTLINE_MAX_CHARS
                            or slots + line_slots > SLOT_BUDGET):
                chunks.append(current)
                current, size, slots = [], 0, 0
            current.append(line)
            size += line_size
            slots += line_slots
    if current:
        chunks.append(current)
    return chunks


SYSTEM = '''你是 Word 表单结构分析员。用户给你一份 .docx 的完整结构清单：正文段落（p0、p1…）
和表格单元格（t0 r0 c0…），每行是坐标 :: 原文。方括号里是标记：
下划线=该段有下划线；空位=有连续下划线或空白；占位符:xxx=文档自带的 {{xxx}} 记号；
空单元格=这一格没有内容。

任务：找出所有“需要人来填写”的位置，并给出栏目名。

是填写位：空单元格、下划线、连续空白、{{占位符}}、括号提示（请填写/请输入/在此输入）、日期空位。
不是填写位：已有成段文字、说明/注意事项/示例/参考文献、签名盖章、小节标题本身。

输出规则：
1. label 必须是清单里真实出现过的标签（冒号前的词、表头、占位符名），原样照抄，不翻译、不扩写。
2. 表格里值是右边或下边那一格：anchor 指向值所在的格子，label 用同一行左侧的标签。
3. 同一个值在封面和表格里都出现时，分别列出（各用各的坐标），label 用最完整的写法。
4. type 取 text（单行）、multiline（整段/多行）、date（日期）、number（数字）、select（勾选/择一）。
5. group 用所在章节名或表名，实在没有就写“填写内容”。
6. anchor 原样引用坐标：正文 {"kind":"body","p":3}；单元格 {"kind":"cell","t":0,"r":0,"c":1}。
7. 不确定就不要输出；文档里没有的东西不要编。只处理本次清单里出现的位置。

只输出 JSON：
{"fields":[{"label":"项目名称","anchor":{"kind":"body","p":1},"type":"text","group":"封面",
"confidence":"高","why":"下划线空位"}],"note":""}'''


# ---------------------------------------------------------------- 校验
def _cell_of(doc, table, row, col):
    try:
        return doc.tables[table].rows[row].cells[col]
    except (IndexError, KeyError, AttributeError):
        return None


def _check_label(label, doc, anchor):
    if not label or len(label) > MAX_LABEL:
        return '标签为空或过长'
    if NOISE_LABEL.search(label) or label in NOISE_EXACT:
        return '签名/说明类不是填写位'
    if anchor.get('kind') == 'body':
        index = anchor['p']
        if not 0 <= index < len(doc.paragraphs):
            return '段落坐标越界'
        text = doc.paragraphs[index].text
        if _wire(label) not in _wire(text):
            return '标签不在该段落原文里'
        return ''
    cell = _cell_of(doc, anchor['t'], anchor['r'], anchor['c'])
    if cell is None:
        return '单元格坐标越界'
    row_text = ' '.join(c.text for c in doc.tables[anchor['t']].rows[anchor['r']].cells)
    if _wire(label) not in _wire(row_text) and _wire(label) not in _wire(cell.text):
        return '标签不在该行原文里'
    return ''


def _cell_ok(cell):
    """只接受空格子或带空位标记的格子。

    表头行（“姓名 / 出生年月 / 职称…”）也长得像短格子，模型很容易把它们当成值格；
    已经填了字的格子（示例值、往年填空）更不能让模型决定往里写，位置一律交给结构结果。
    """
    text = (cell.text or '').strip()
    return not text or _has_slot(text)


def _coerce_type(value, label):
    kind = str(value or 'text').strip().lower()
    kind = TYPE_ALIASES.get(kind, kind)
    if kind not in CLEAN_TYPE:
        kind = 'text'
    if kind == 'date' and not re.search(r'日|时间|日期', label):
        kind = 'text'
    return kind


def _coerce_confidence(value):
    return CONFIDENCE_ALIASES.get(str(value or '').strip().lower(), '中')


def _coerce_anchor(value):
    if not isinstance(value, dict):
        return None
    kind = str(value.get('kind') or '').strip().lower()
    try:
        if kind == 'body':
            index = int(value.get('p'))
            return {'kind': 'body', 'p': index} if index >= 0 else None
        if kind == 'cell':
            table, row, col = (int(value.get(k)) for k in ('t', 'r', 'c'))
            if min(table, row, col) < 0:
                return None
            return {'kind': 'cell', 't': table, 'r': row, 'c': col}
    except (TypeError, ValueError):
        return None
    return None


def _validate(doc, item):
    """把模型给的一条结果收紧成可执行结果；不合法返回 (None, 原因)。"""
    label = _clean_label(item.get('label'))
    anchor = _coerce_anchor(item.get('anchor'))
    if anchor is None:
        return None, '缺少可用坐标'
    reason = _check_label(label, doc, anchor)
    if reason:
        return None, reason
    if anchor['kind'] == 'cell':
        cell = _cell_of(doc, anchor['t'], anchor['r'], anchor['c'])
        if cell is None:
            return None, '单元格坐标越界'
        if not _cell_ok(cell):
            # 模型常把坐标指到左侧的标签格：值其实在右边一格
            right = _cell_of(doc, anchor['t'], anchor['r'], anchor['c'] + 1)
            if right is None or right._tc is cell._tc or not _cell_ok(right):
                return None, '单元格里已有成段正文'
            anchor = {'kind': 'cell', 't': anchor['t'], 'r': anchor['r'],
                      'c': anchor['c'] + 1}
    else:
        paragraph = doc.paragraphs[anchor['p']]
        text = paragraph.text
        # 尾随冒号、括号提示（【请填写…】）都算空位；这里只负责“是不是填写位”，
        # 具体写到哪里仍然由 visionparse._locate 在原文里找，找不到就不会新增。
        if not (_has_slot(text) or _underline(paragraph) or PROMPT_RE.search(text)
                or text.rstrip().endswith(('：', ':'))):
            return None, '该段落没有空位标记'
    return {
        'label': label, 'anchor': anchor,
        'type': _coerce_type(item.get('type'), label),
        'group': _clean_label(item.get('group')) or '填写内容',
        'confidence': _coerce_confidence(item.get('confidence')),
        'why': str(item.get('why') or '')[:40],
    }, ''


# ---------------------------------------------------------------- 调用
def _budget(lines):
    """按这段里空位的数量给输出留 token：空位越多，栏目越多。"""
    slots = sum(1 for line in lines
                if any(mark in line for mark in ('空单元格', '下划线', '空位', '占位符')))
    return max(1600, min(MAX_OUTPUT_TOKENS, 900 + 55 * slots))


def _ask_lines(client, lines, head, out, errors, progress):
    """一段一段问；某段返回的 JSON 不完整（输出被截断）就二分再问。"""
    if not lines:
        return
    try:
        data = client.ask_json(SYSTEM, head + '\n'.join(lines), expect_key='fields',
                               temperature=0.0, max_tokens=_budget(lines))
    except feedback.Cancelled:
        raise                       # 用户点了停止：不能当成“这次请求失败”继续重试
    except Exception as exc:
        if len(lines) > MIN_SPLIT_LINES:
            middle = len(lines) // 2
            if progress:
                progress('文本模型：这一段返回不完整，拆成两半重试…')
            _ask_lines(client, lines[:middle], head, out, errors, progress)
            _ask_lines(client, lines[middle:], head, out, errors, progress)
        else:
            errors.append(str(exc)[:120])
        return
    fields = data.get('fields')
    if not isinstance(fields, list):
        errors.append('模型返回的 fields 不是数组')
        return
    out.extend(item for item in fields[:MAX_FIELDS] if isinstance(item, dict))


def _ask(client, chunks, progress):
    results, errors = [], []
    total = len(chunks)
    for index, chunk in enumerate(chunks[:MAX_CHUNKS]):
        if progress:
            progress(f'文本模型：正在通读第 {index + 1}/{total} 段结构…')
        head = (f'这是第 {index + 1}/{total} 段结构清单，只处理这一段里出现的位置。\n'
                if total > 1 else '')
        _ask_lines(client, chunk, head, results, errors, progress)
    if not results and errors:
        raise ai.AIError('文本模型没有返回可用结果：' + '；'.join(errors[:2]))
    if total > MAX_CHUNKS:
        errors.append(f'文档结构有 {total} 段，只增强了前 {MAX_CHUNKS} 段')
    return results, errors


# ---------------------------------------------------------------- 合并
def _same_place(field, anchor):
    loc = field.get('location') or {}
    if anchor['kind'] == 'body':
        return loc.get('kind') == 'body' and loc.get('p') == anchor['p']
    return (loc.get('kind') == 'cell' and loc.get('table') == anchor['t']
            and loc.get('row') == anchor['r'] and loc.get('col') == anchor['c'])


def merge(source, plan, raw, model='', errors=()):
    """把模型结果并进结构结果：只改标签与属性，位置永远以结构结果为准。"""
    plan = copy.deepcopy(plan)
    fields = list(plan.get('fields') or [])
    doc = Document(str(source))
    used = [visionparse._norm_range(f.get('bookmark_range')) for f in fields
            if f.get('bookmark_range')]
    known = {visionparse._norm_label(f.get('label')) for f in fields if f.get('label')}
    accepted, renamed, retyped, added, rejected, disputes = [], 0, 0, 0, [], []
    seen = set()

    for item in raw:
        field, reason = _validate(doc, item)
        if field is None:
            label = _clean_label(item.get('label'))
            rejected.append(f'{label or "?"}：{reason}')
            continue
        key = (field['label'], tuple(sorted(field['anchor'].items())))
        if key in seen:
            continue
        seen.add(key)
        norm = visionparse._norm_label(field['label'])
        same_name = [f for f in fields if visionparse._norm_label(f.get('label')) == norm]
        target = next((f for f in same_name if _same_place(f, field['anchor'])), None)
        if target is None and same_name:
            target = same_name[0]
        if target is None:
            target = next((f for f in fields if _same_place(f, field['anchor'])), None)
        if target is not None:
            accepted.append(field['label'])
            if visionparse._generic_label(target.get('label')):
                target['label'] = field['label']
                renamed += 1
            elif (_wire(target['label']) != _wire(field['label'])
                  and visionparse._norm_label(target['label'])
                  != visionparse._norm_label(field['label'])):
                disputes.append({'label': target['label'], 'alt': field['label'],
                                 'reason': f'模型读作 {field["label"]}'})
            if field['confidence'] == '高':
                if target.get('type') in ('text', None) and field['type'] != 'text':
                    target['type'] = field['type']
                    retyped += 1
                if _wire(field['group']) not in GENERIC_GROUPS \
                        and target.get('group') in GENERIC_GROUPS | {f'表{field["anchor"].get("t", 0) + 1}'}:
                    target['group'] = field['group']
            target['text_checked'] = True
            continue

        if field['confidence'] == '低':
            rejected.append(f'{field["label"]}：置信度低')
            continue
        located = visionparse._locate(doc, field, known, used)
        if not located:
            rejected.append(f'{field["label"]}：原文里找不到可写入的位置')
            disputes.append({'label': field['label'], 'reason': '结构里没有对应空位'})
            continue
        rng, p, context = located
        rng = visionparse._norm_range(rng)
        if any(visionparse._overlaps(rng, old) for old in used):
            rejected.append(f'{field["label"]}：与已有栏目重叠')
            continue
        new_field = visionparse._new_field(
            {'label': field['label'], 'type': field['type'], 'group': field['group'],
             'confidence': field['confidence'], 'sample': '', 'section': ''},
            rng, p, context, source='文本模型')
        new_field['text_added'] = True
        fields.append(new_field)
        used.append(rng)
        known.add(visionparse._norm_label(field['label']))
        added += 1

    # 格阵（成员×7、预算×6）的每格键名依赖栏目名与编号。合并后必须重算，
    # 否则 plan['grids'] 还引用旧键名，生成时会写进错误的格子。
    for field in fields:
        field.pop('grid', None)
        field.pop('header', None)
    grids = smartparse.detect_grids(fields)
    fields[:] = [field for field in fields if not field.get('header')]
    fields.sort(key=smartparse._sort_key)
    smartparse._renumber(fields)
    for grid in grids:
        grid['rows'] = [[field['key'] if field else '' for field in row]
                        for row in grid['rows']]
    plan['fields'] = fields
    plan['grids'] = grids
    plan['text'] = {
        'used': True, 'model': model, 'matched': len(accepted), 'renamed': renamed,
        'retyped': retyped, 'added': added, 'rejected': rejected[:20],
        'disputes': disputes[:12], 'errors': list(errors)[:5],
    }
    note = (f'文本模型（{model or "未命名"}）通读结构：校正 {len(accepted)} 项、'
            f'改名 {renamed} 项、补识别 {added} 项')
    if rejected:
        note += f'、{len(rejected)} 项未采纳'
    plan['description'] = (plan.get('description') or '') + ' ' + note + '。'
    return plan


def enhance(source, plan, cfg, progress=None):
    """结构结果 + 文本模型增强；模型不可用或失败时原样返回结构结果，不打断流程。"""
    plan = copy.deepcopy(plan)
    client = ai.Client(cfg)
    ok, message = client.usable()
    if not ok:
        plan['text'] = {'used': False, 'error': message}
        return plan
    if client.provider == ai.EMBEDDED_BACKEND:
        plan['text'] = {'used': False,
                        'error': '内置 0.6B 小模型不做整篇识别，请配置文本模型或本机 Ollama'}
        return plan
    try:
        doc = Document(str(source))
        outline = build_outline(doc)
        chunks = _chunks(outline)
        if progress:
            progress(f'文本模型：文档结构 {len(chunks)} 段，开始识别…')
        raw, errors = _ask(client, chunks, progress)
        if progress:
            progress(f'文本模型：整理 {len(raw)} 条候选栏目…')
        if not raw:
            # 模型一条都没给出：保留原 plan，不做无意义的重排与键位重算
            plan['text'] = {'used': True, 'model': client.model, 'matched': 0,
                            'renamed': 0, 'retyped': 0, 'added': 0, 'rejected': [],
                            'disputes': [], 'errors': list(errors)[:5]}
            return plan
        return merge(source, plan, raw, client.model, errors)
    except feedback.Cancelled:
        raise
    except Exception as exc:
        plan['text'] = {'used': False, 'error': str(exc)[:200]}
        plan['description'] = (plan.get('description') or '') + \
            f' 文本模型未完成：{str(exc)[:140]}。'
        return plan
