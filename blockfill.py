"""发现并规范化跨段落的空 {{ ... }} 区域。

空括号区域没有字段名，无法直接注册。这里先给每个区域一个上下文名，
再把原括号对替换成单个 {{ 名称 }} 标记，输出到新文件；原文件不动。
"""
from __future__ import annotations

import copy
import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HEADING = re.compile(r'^[一二三四五六七八九十]+、')
TASK = re.compile(r'任务\s*([一二三四五六七八九十百\d]+)')
QUESTION = re.compile(r'第\s*([一二三四五六七八九十\d]+)\s*题')
THINK = re.compile(r'思考题\s*([一二三四五六七八九十\d]+)')
TOKEN = re.compile(r'\{\{\s*([^{}\r\n]{1,100}?)\s*\}\}')


def _text(p):
    return ''.join(t.text or '' for t in p.iter(qn('w:t')))


def _paragraphs(doc):
    return list(doc.element.body.iter(qn('w:p')))


def _is_blank(p):
    if _text(p).strip():
        return False
    for tag in ('w:drawing', 'w:pict', 'w:object', 'w:br', 'w:tab'):
        if p.findall('.//' + qn(tag)):
            return False
    return True


def _clean_label(text):
    value = str(text or '').replace('{{', '').replace('}}', '')
    value = re.sub(r'[：:，,。；;、!！?？]', '', value)
    return re.sub(r'\s+', '', value)[:30]


def _task(text):
    match = TASK.search(text or '')
    return match.group(1) if match else ''


def _question(text):
    match = QUESTION.search(text or '')
    return match.group(1) if match else _task(text)


def _think(text):
    match = THINK.search(text or '')
    return match.group(1) if match else _task(text)


def _is_screenshot(*texts):
    joined = ' '.join(texts)
    return '截图' in joined or '图片' in joined or \
        ('运行结果' in joined and '图' in joined)


def _fallback_name(label, context, section):
    joined = ' '.join((label, context, section))
    if '程序分析' in joined:
        q = _question(joined)
        return (f'第{q}题程序分析' if q else '程序分析')
    if '思考题' in joined:
        q = _think(joined)
        return (f'思考题{q}回答内容' if q else '思考题回答内容')
    task = _task(joined)
    if task:
        if '实验步骤' in section or '实验步骤' in context:
            return f'任务{task}实验步骤代码内容'
        if '代码' in joined or '实验内容' in section:
            return f'任务{task}代码内容'
        return f'任务{task}内容'
    q = QUESTION.search(joined)
    if q:
        return f'第{q.group(1)}题内容'
    base = _clean_label(label) or _clean_label(section).replace('实验内容', '').replace('实验步骤', '')
    if not base:
        base = '填写'
    if not base.endswith('内容'):
        base += '内容'
    return base[:16]


def _sanitize_name(value):
    text = re.sub(r'[{}\[\]"\'“”‘’\r\n\t]', '', str(value or '')).strip()
    text = re.sub(r'[：:，,。；;、!！?？]+$', '', text)
    text = re.sub(r'\s+', '', text)
    return text[:24]


def _unique(name, used):
    name = _sanitize_name(name) or '填写内容'
    base, n = name, 2
    while name in used:
        name = f'{base}{n}'
        n += 1
    used.add(name)
    return name


def find_blocks(docx_path):
    """按文档顺序找跨段落的空括号对，返回带上下文的块列表。"""
    doc = Document(docx_path)
    paras = _paragraphs(doc)
    blocks, pending = [], None
    last_texts, last_heading = [], ''
    for i, p in enumerate(paras):
        text = _text(p)
        stripped = text.strip()
        if stripped and HEADING.match(stripped):
            last_heading = stripped
        if pending is None:
            if stripped == '{{' or stripped.endswith('{{'):
                own_label = _clean_label(text)
                label = own_label
                if not label:
                    label = _clean_label(last_texts[-1] if last_texts else '')
                context = ' / '.join(x for x in last_texts[-2:] if x)
                shot = _is_screenshot(own_label or label)
                blocks.append({
                    'index': len(blocks), 'open_index': i, 'close_index': None,
                    'parent': p.getparent(), 'label': label,
                    'section': last_heading, 'context': context,
                    'is_screenshot': shot,
                    'fallback': _fallback_name(label, context, last_heading),
                    'name': '', 'name_source': '',
                })
                pending = blocks[-1]
        else:
            if stripped == '}}' or stripped.startswith('}}'):
                if p.getparent() is pending['parent']:
                    pending['close_index'] = i
                    pending = None
        if stripped:
            last_texts.append(stripped[:80])
            last_texts = last_texts[-3:]
    return [b for b in blocks if b['close_index'] is not None]


def model_items(blocks, used_names=()):
    used = list(used_names)
    items = []
    for block in blocks:
        if block['is_screenshot']:
            continue
        items.append({
            'i': block['index'],
            'section': block['section'],
            'label': block['label'],
            'context': block['context'],
            'hint': block['fallback'],
        })
        used.append(block['fallback'])
    return items, used


def resolve_names(blocks, suggested, existing=()):
    """合并模型建议与回退名，返回每块的最终名称（截图保持空）。"""
    used = {x for x in existing if x}
    for block in blocks:
        if block['is_screenshot']:
            block['name'] = ''
            block['name_source'] = 'screenshot'
            continue
        name = _sanitize_name(suggested.get(block['index'], ''))
        source = 'model' if name else 'rule'
        if not name:
            name = block['fallback']
        block['name'] = _unique(name, used)
        block['name_source'] = source
    return blocks


# ---------------- XML rewrite helpers

def _strip_token(p, token):
    nodes = list(p.iter(qn('w:t')))
    full = ''.join(n.text or '' for n in nodes)
    idx = full.rfind(token)
    if idx < 0:
        return False
    pos, end = 0, idx + len(token)
    for node in nodes:
        value = node.text or ''
        start = max(idx, pos)
        stop = min(end, pos + len(value))
        if start < stop:
            node.text = value[:start - pos] + value[stop - pos:]
        pos += len(value)
    return True


def _marker_run(p, name):
    first = next((el for el in p.iter(qn('w:r'))), None)
    rpr = first.find(qn('w:rPr')) if first is not None else None
    run = OxmlElement('w:r')
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    t = OxmlElement('w:t')
    t.text = '{{ ' + name + ' }}'
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    run.append(t)
    return run


def _set_marker(p, name):
    for child in list(p):
        if child.tag != qn('w:pPr'):
            p.remove(child)
    p.append(_marker_run(p, name))


def _insert_marker_after(p, name):
    new_p = OxmlElement('w:p')
    ppr = p.find(qn('w:pPr'))
    if ppr is not None:
        new_p.append(copy.deepcopy(ppr))
    new_p.append(_marker_run(p, name))
    p.addnext(new_p)
    return new_p


def normalize(src, dst, blocks):
    """把识别到的空括号块写成命名字段，输出到 dst。原文件不修改。"""
    doc = Document(src)
    paras = _paragraphs(doc)
    made, skipped = [], []
    for block in blocks:
        oi, ci = block['open_index'], block['close_index']
        if not (0 <= oi < ci < len(paras)):
            continue
        p_open, p_close = paras[oi], paras[ci]
        if p_open.getparent() is not p_close.getparent():
            continue
        if block['is_screenshot']:
            if _text(p_open).strip() == '{{':
                p_open.getparent().remove(p_open)
            else:
                _strip_token(p_open, '{{')
            if _text(p_close).strip() == '}}':
                p_close.getparent().remove(p_close)
            else:
                _strip_token(p_close, '}}')
            skipped.append(block)
            continue
        name = block.get('name') or block.get('fallback')
        if not name:
            continue
        if _text(p_open).strip() == '{{':
            _set_marker(p_open, name)
            marker = p_open
        else:
            _strip_token(p_open, '{{')
            marker = _insert_marker_after(p_open, name)
        parent = marker.getparent()
        node = marker.getnext()
        while node is not None and node is not p_close:
            nxt = node.getnext()
            if node.tag == qn('w:p') and _is_blank(node):
                parent.remove(node)
            node = nxt
        if p_close.getparent() is parent:
            if _text(p_close).strip() == '}}':
                parent.remove(p_close)
            else:
                _strip_token(p_close, '}}')
        made.append(name)
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst)
    return {'fields': made, 'skipped': len(skipped), 'output': str(dst)}


def existing_names(path):
    """宽松提取已有具名标记；不因别的空括号对而失败。"""
    try:
        doc = Document(path)
    except Exception:
        return []
    names = []
    for p in _paragraphs(doc):
        for match in TOKEN.finditer(_text(p)):
            name = match.group(1).strip()
            if name and name not in names:
                names.append(name)
    return names
