"""Word 页面的读取与“少量截图复核”。

识别主力是 textparse（结构 + 文本模型），这里只负责两件事：
  1. render_selected：把 docx 排成 PDF，只取含争议栏目文字的页面；
  2. verify：只对文本模型与结构结果有分歧的少数栏目做截图复核，
     复核只用来改栏目名与类型，不新增位置。
"""
from __future__ import annotations

import copy
import io
import re
import shutil
import tempfile
from pathlib import Path

import pypdfium2 as pdfium

import ai
import bookmarkfill
import feedback
import formparse
import preview
import smartparse

VERIFY_SYSTEM = (
    '你是中文 Word 表单的校对员。用户会给你几页截图，以及一批“有疑问的栏目名”。'
    '你只需要回答这些栏目：页面上是否存在、准确的栏目名是什么、是单行还是整段。'
    '不要新增栏目，不要输出填写内容，不要评价版式。只输出 JSON。'
)

VERIFY_PROMPT = (
    '【要核对的栏目】{labels}\n'
    '【输出格式】{{"checks":[{{"label":"带疑问的栏目名","exists":true,'
    '"real_label":"页面上准确的栏目名","type":"text|multiline|date|number|select",'
    '"confidence":"高|中|低"}}]}}\n'
    '规则：real_label 必须照抄页面上出现的字；页面上找不到就 exists=false。'
)

NOISE_WORDS = (
    '填写说明', '填表说明', '注意事项', '格式要求', '例如', '页码', '目录',
    '签字', '签名', '盖章', '公章', '手签', '空一行',
    '宋体', '黑体', '字号', '行距', '缩进', '排版', '页边距', '字体',
)


def _text(p):
    return ''.join(r.text for r in p.runs)


def _norm_label(text):
    text = str(text or '').replace('\u3000', '').replace('\xa0', '')
    text = re.sub(r'[\s：:；;，,。、\.\-—_（）()【】\[\]“”"\'`/\\]+', '', text)
    return text.lower()


def _clean_label(text):
    text = str(text or '').strip()
    text = re.sub(r'^[\s\d一二三四五六七八九十]+[、.．)）]\s*', '', text)
    text = text.strip('“”"\'` ')
    text = re.sub(r'[：:；;，,。、\s]+$', '', text)
    text = re.sub(r'（(?:请填写|请填|填写|示例|选填|必填|单位|格式)[^）]*）', '', text)
    text = re.sub(r'\((?:请填写|请填|填写|示例|选填|必填|单位|格式)[^)]*\)', '', text)
    return text.strip()


def _noise_label(label):
    plain = _norm_label(label)
    if len(plain) < 2 or len(label) > 30:
        return True
    if re.fullmatch(r'[\dA-Za-z_\-]+', plain):
        return True
    if plain in ('日期', '时间', '年月日', '序号'):
        return True
    return any(word in plain for word in NOISE_WORDS)





def _meaningful(text):
    plain = re.sub(r'[\s\u3000_＿﹍﹏…\.·\-—]+', '', text or '')
    return bool(plain)


def _is_hint_only(text):
    text = (text or '').strip()
    if not text:
        return False
    if re.fullmatch(r'[（(【\[][^）)】\]]{0,24}[）)】\]]', text):
        return True
    return bool(re.match(r'^(?:请|在此|此处|填写|输入|签名|签字|盖章|日期|格式|说明)', text)
                and len(text) <= 24)


def _signature_line(text):
    return any(word in (text or '') for word in ('签名', '签字', '盖章', '公章'))


def _find_label_span(text, label):
    """Find a standalone occurrence, rejecting ``姓名`` inside ``学生姓名``."""
    if not text or not label:
        return None
    start = text.find(label)
    while start >= 0:
        end = start + len(label)
        if _boundary_ok(text, start, end):
            return start, end
        start = text.find(label, start + 1)
    pattern = r'[\s\u3000\xa0]*'.join(re.escape(ch) for ch in label)
    for match in re.finditer(pattern, text):
        if _boundary_ok(text, match.start(), match.end()):
            return match.start(), match.end()
    return None


def _boundary_ok(text, start, end):
    before = text[start - 1] if start > 0 else ''
    after = text[end] if end < len(text) else ''
    if before and re.match(r'[\u4e00-\u9fff]', before):
        return False
    if after and re.match(r'[\u4e00-\u9fff]', after):
        return False
    return True


def _after_label(text, end):
    pos = end
    while pos < len(text) and text[pos] in ' \u3000\t':
        pos += 1
    rest = text[pos:]
    hint = re.match(r'[（(][^）)]{0,24}[）)]', rest)
    if hint and re.search(r'填写|请|说明|示例|选填|必填', hint.group(0)):
        pos += hint.end()
    while pos < len(text) and text[pos] in ' \u3000\t':
        pos += 1
    if pos < len(text) and text[pos] in '：:':
        pos += 1
    return pos


def _tail_end(text, start):
    end = len(text)
    note = smartparse.NOTE_RE.search(text, start)
    if note:
        end = min(end, note.start())
    next_label = re.search(r'[\u4e00-\u9fff（）()/\s]{2,24}[：:]', text[start:])
    if next_label and next_label.start() > 0:
        end = min(end, start + next_label.start())
    while end > start and text[end - 1] in ' \u3000\t':
        end -= 1
    return end


def _range(kind, part, p, start, end_p, end, ti=None, ri=None, ci=None):
    value = {'container': kind, 'part': part,
             'start_p': p, 'start': start, 'end_p': end_p, 'end': end}
    if kind == 'cell':
        value.update(table=ti, row=ri, col=ci)
    return value


def _norm_range(rng):
    value = dict(rng or {})
    value.setdefault('container', 'cell')
    value.setdefault('part', 'body')
    return value


def _overlaps(a, b):
    a, b = _norm_range(a), _norm_range(b)
    if a.get('container') != b.get('container'):
        return False
    if a.get('container') == 'cell':
        if (a.get('table'), a.get('row'), a.get('col')) != \
                (b.get('table'), b.get('row'), b.get('col')):
            return False
    elif a.get('container') in ('header', 'footer'):
        if a.get('part') != b.get('part'):
            return False
    if a.get('start_p') > b.get('end_p') or b.get('start_p') > a.get('end_p'):
        return False
    if a.get('start_p') == a.get('end_p') == b.get('start_p') == b.get('end_p'):
        return a.get('start') == b.get('start')
    return True


def _covered(rng, used):
    return any(_overlaps(rng, old) for old in used)


def empty_plan(source, reason=''):
    return {
        'source': str(source), 'title': '',
        'category': '',
        'description': '结构识别没有发现明确填写位，已改用文本模型复核文档结构。',
        'table': 0, 'fields': [], 'recipes': [], 'clear_cells': [],
        'row_heights': {}, 'remove_textboxes': False, 'smart': True,
    }


def _generic_label(label):
    return not label or bool(re.search(r'填写项|字段\d*|控件\d*|填空\d*|^第.+项$', label))


def _containers(doc):
    result = [('body', 'body', list(doc.paragraphs), None, None, None, None)]
    seen = set()
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                result.append(('cell', 'body', list(cell.paragraphs),
                               ti, ri, ci, cell))
    for kind, name, root in bookmarkfill.story_parts(doc):
        if kind == 'body':
            continue
        result.append((kind, name, bookmarkfill._paragraphs_of_root(root),
                       None, None, None, None))
    return result


def _looks_like_label(text, known):
    plain = _norm_label(text)
    if not plain:
        return False
    if plain in known:
        return True
    return (text or '').strip().endswith(('：', ':')) and len(text.strip()) <= 28


def _lookahead(paras, index, known, used, kind, part, ti, ri, ci):
    blanks = []
    for q in range(index + 1, min(len(paras), index + 6)):
        text = _text(paras[q])
        if not text.strip():
            blanks.append(q)
            continue
        if blanks:
            rng = _range(kind, part, blanks[0], 0, blanks[-1], 0, ti, ri, ci)
            return None if _covered(rng, used) else rng
        if formparse._is_note(paras[q]):
            continue
        if _signature_line(text):
            return None
        if _is_hint_only(text):
            end = len(text.rstrip())
            rng = _range(kind, part, q, end, q, end, ti, ri, ci)
            return None if _covered(rng, used) else rng
        if smartparse._is_heading(text) and not re.search(r'[：:]', text):
            return None
        if _looks_like_label(text, known):
            return None
        end_p, end = q, len(text)
        for r in range(q + 1, min(len(paras), q + 8)):
            value = _text(paras[r])
            if not value.strip():
                continue
            if formparse._is_note(paras[r]) or _signature_line(value):
                break
            if smartparse._is_heading(value) and not re.search(r'[：:]', value):
                break
            if _looks_like_label(value, known):
                break
            end_p, end = r, len(value)
        rng = _range(kind, part, q, 0, end_p, end, ti, ri, ci)
        return None if _covered(rng, used) else rng
    if blanks:
        rng = _range(kind, part, blanks[0], 0, blanks[-1], 0, ti, ri, ci)
        return None if _covered(rng, used) else rng
    return None


def _adjacent_value(doc, ti, ri, ci, cell, known, used):
    try:
        cells = doc.tables[ti].rows[ri].cells
    except (IndexError, KeyError):
        return None
    for column in range(ci + 1, len(cells)):
        target = cells[column]
        if target._tc is cell._tc:
            continue
        text = ' '.join(_text(p) for p in target.paragraphs).strip()
        if text and (_looks_like_label(text, known) or _signature_line(text)):
            break
        rng = formparse._cell_sample_range(target, ti, ri, column)
        rng = _norm_range(rng)
        return None if _covered(rng, used) else rng
    return None


def _candidate(kind, part, paras, p, ti, ri, ci, cell, text, span, known, used, doc):
    label_end = span[1]
    after = _after_label(text, label_end)
    tail = text[after:]
    if kind == 'cell' and not _meaningful(tail):
        adjacent = _adjacent_value(doc, ti, ri, ci, cell, known, used)
        if adjacent:
            return adjacent
    if _meaningful(tail) and not _is_hint_only(tail):
        end = _tail_end(text, after)
        start = after + (len(tail) - len(tail.lstrip(' \u3000\t')))
        if end > start:
            rng = _range(kind, part, p, start, p, end, ti, ri, ci)
            return None if _covered(rng, used) else rng
    blank = re.match(r'[_\u3000\s＿﹍﹏…\.·\-—]+', tail)
    if blank:
        rng = _range(kind, part, p, after, p, after + blank.end(), ti, ri, ci)
        if not _covered(rng, used):
            return rng
    ahead = _lookahead(paras, p, known, used, kind, part, ti, ri, ci)
    if ahead:
        return ahead
    end = len(text.rstrip())
    rng = _range(kind, part, p, end, p, end, ti, ri, ci)
    return None if _covered(rng, used) else rng


def _locate(doc, vision, known, used):
    label = _clean_label(vision.get('label'))
    if not label:
        return None
    candidates = []
    for kind, part, paras, ti, ri, ci, cell in _containers(doc):
        for p, para in enumerate(paras):
            text = _text(para)
            if not text.strip() or formparse._is_note(para):
                continue
            if _signature_line(text) and not any(w in label for w in ('意见', '评语')):
                continue
            span = _find_label_span(text, label)
            if not span:
                continue
            rng = _candidate(kind, part, paras, p, ti, ri, ci, cell,
                             text, span, known, used, doc)
            if not rng:
                continue
            score = 10
            if kind == 'cell':
                score += 4
            if _norm_label(text) == _norm_label(label):
                score += 4
            if text.strip().startswith(label):
                score += 2
            if len(text) <= 40:
                score += 1
            candidates.append((score, rng, p, text))
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1:]


def _new_field(vision, rng, p, context, source='视觉模型'):
    rng = _norm_range(rng)
    kind = rng.get('container')
    page = vision.get('page')
    prefix = f'第{page}页 · ' if page else ''
    origin = '文本模型定位' if source == '文本模型' else '视觉定位'
    if kind == 'cell':
        where = (f'{prefix}表{rng.get("table", 0) + 1} 行{rng.get("row", 0) + 1} '
                 f'列{rng.get("col", 0) + 1} · {origin}')
        location = {'kind': 'cell', 'table': rng.get('table'),
                    'row': rng.get('row'), 'col': rng.get('col'), 'p': p}
    elif kind in ('header', 'footer'):
        name = '页眉' if kind == 'header' else '页脚'
        where = f'{prefix}{name} · {origin}'
        location = {'kind': kind, 'part': rng.get('part'), 'p': p}
    else:
        where = f'{prefix}正文 第{p + 1}段 · 视觉定位'
        location = {'kind': 'body', 'p': p}
    return {
        'plan_id': '', 'key': '', 'label': _clean_label(vision.get('label')),
        'type': vision.get('type') or 'text',
        'required': bool(vision.get('required', False)),
        'group': vision.get('group') or '填写内容',
        'where': where, 'context': (context or '')[:120],
        'original_content': (vision.get('sample') or vision.get('hint') or '')[:100],
        'confidence': vision.get('confidence') or '中',
        'source': source, 'location': location, 'bookmark_range': rng,
        'bookmark': '', 'control': None, 'vision': True,
        'vision_page': page, 'vision_section': vision.get('section') or '',
        'vision_confidence': vision.get('confidence') or '中',
    }


# ---------------------------------------------------------------- 争议项复核
MAX_CHECK_PAGES = 3
MAX_CHECK_LABELS = 6


def render_selected(source, labels, max_pages=MAX_CHECK_PAGES, progress=None):
    """只渲染含争议栏目文字的那几页：先查 PDF 文字层，找到了才截图。"""
    wanted = [(_norm_label(label), label) for label in labels if _norm_label(label)]
    if not wanted:
        return [], 0
    work = Path(tempfile.mkdtemp(prefix='wtt-check-'))
    pdf_path = work / 'source.pdf'
    try:
        if progress:
            progress('视觉复核：正在把页面排版成 PDF…')
        preview.export_pdf_sync(str(source), pdf_path, progress)
        doc = pdfium.PdfDocument(pdf_path.read_bytes())
        try:
            picked = []
            for index in range(len(doc)):
                if len(picked) >= max_pages:
                    break
                page = doc[index]
                try:
                    page_text = _norm_label(page.get_textpage().get_text_range())
                except Exception:
                    continue
                hits = [raw for norm, raw in wanted if norm in page_text]
                if not hits:
                    continue
                if progress:
                    progress(f'视觉复核：正在截图第 {index + 1} 页…')
                image = page.render(scale=1.6).to_pil().convert('RGB')
                buffer = io.BytesIO()
                image.save(buffer, 'JPEG', quality=80, optimize=True)
                picked.append({'page': index + 1, 'labels': hits,
                               'text': page_text, 'image': buffer.getvalue()})
            return picked, len(doc)
        finally:
            doc.close()
    finally:
        shutil.rmtree(work, ignore_errors=True)


def verify(source, plan, cfg, progress=None):
    """对少数有分歧的栏目做截图复核；只改名字与类型，不新增、不删除位置。"""
    plan = copy.deepcopy(plan)
    disputes = [item for item in ((plan.get('text') or {}).get('disputes') or [])
                if item.get('label')]
    stats = {'checked': False, 'pages': 0, 'labels': len(disputes),
             'confirmed': 0, 'fixed': 0}
    if not disputes:
        plan['vision'] = stats
        return plan
    ok, message = ai.vision_usable(cfg)
    if not ok:
        stats['skipped'] = message
        plan['vision'] = stats
        return plan
    labels = []
    for item in disputes[:MAX_CHECK_LABELS]:
        for name in (item.get('label'), item.get('alt')):
            if name and name not in labels:
                labels.append(name)
    try:
        picked, total = render_selected(source, labels, progress=progress)
    except feedback.Cancelled:
        raise
    except Exception as exc:
        stats['error'] = str(exc)[:160]
        plan['vision'] = stats
        return plan
    if not picked:
        stats['note'] = '页面上没有找到这些栏目文字，跳过复核'
        plan['vision'] = stats
        return plan
    client = ai.vision_client(cfg)
    if progress:
        progress(f'视觉复核：请视觉模型核对 {len(picked)} 页里的 {len(labels)} 个栏目…')
    try:
        answer = client.ask_vision_json(
            VERIFY_SYSTEM, VERIFY_PROMPT.format(labels='、'.join(labels)),
            [item['image'] for item in picked], expect_key='checks',
            temperature=0.0, max_tokens=1200)
    except feedback.Cancelled:
        raise
    except Exception as exc:
        stats['error'] = str(exc)[:160]
        plan['vision'] = stats
        return plan

    by_label = {_norm_label(f.get('label')): f for f in plan.get('fields') or []}
    page_text = ' '.join(item['text'] for item in picked)
    taken = set(by_label)
    confirmed = fixed = 0
    for item in (answer.get('checks') or [])[:MAX_CHECK_LABELS]:
        if not isinstance(item, dict):
            continue
        target = by_label.get(_norm_label(item.get('label')))
        if target is None or item.get('exists') is False:
            continue
        confirmed += 1
        name = _clean_label(item.get('real_label'))
        if (name and not _noise_label(name) and _norm_label(name) != _norm_label(target['label'])
                and _norm_label(name) in page_text and _norm_label(name) not in taken):
            taken.discard(_norm_label(target['label']))
            target['label'] = name
            taken.add(_norm_label(name))
            fixed += 1
        kind = str(item.get('type') or '').strip().lower()
        if kind in ('multiline', 'date', 'number', 'select') and target.get('type') == 'text':
            target['type'] = kind
            fixed += 1
    stats.update({'checked': True, 'pages': len(picked), 'confirmed': confirmed,
                  'fixed': fixed, 'total_pages': total})
    plan['vision'] = stats
    plan['description'] = (plan.get('description') or '') + \
        f' 视觉复核 {len(picked)} 页：确认 {confirmed} 项' + (
            f'、改正 {fixed} 项' if fixed else '') + '。'
    return plan



