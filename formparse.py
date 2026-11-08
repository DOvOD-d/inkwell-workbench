"""Structure-first recognition for blank forms and multi-table documents.

The colored-example recognizer in exampleparse works well when a template has
blue examples.  Many school forms are blank instead: a title paragraph followed
by empty paragraphs, or several tables containing blank columns.  This module
infers those fields from the OOXML structure:

* label/value rows become short fields, including empty value cells;
* full-width cells become one or more multiline fields named after their
  section headings, with sample prose and red instructions removed;
* every table is handled, so composite documents (cover + task book + opening
  report + guidance records) are covered;
* a pure text/guidance document has no table fields and is rejected instead of
  being registered with made-up fields.

The plan format is compatible with exampleparse.build(), so the same review,
rename and registration UI works for both recognizers.
"""
from __future__ import annotations

import re

from docx import Document

import exampleparse as ep

NOTE_RE = re.compile(
    r'中文字全部采用|英文、数字采用|行间距|每页都要|首行缩进|悬挂缩进|此处格式参考|'
    r'参考文献不少于|与论文系统上题目一致|请注意此处有变化|此处填写应与|此处空一行|整页设计|'
    r'此处成绩|一定不能涂改|分数不能低于|不少于\s*\d+\s*字|仅供参考|不要跳页|'
    r'此页只填写|至少要有|最后一次记录|空一行|手签|此处填写')
HEADING_RE = re.compile(r'^（[一二三四五六七八九十百]+）')
CN_HEADING_RE = re.compile(r'^[一二三四五六七八九十]+、')
NUMBERED_RE = re.compile(r'^\d+[\.、]')
SIGN_WORDS = ('签字', '签名', '盖章', '公章')


def _text(p):
    return ''.join(r.text for r in p.runs)


def _norm(text):
    return re.sub(r'\s+', '', text or '')


def _clean_label(text):
    return re.sub(r'[\s：:]+', '', text or '').strip()


def _clean_title(text):
    t = re.sub(r'（请注意此处有变化）', '', text or '').strip()
    t = re.sub(r'[（(]此处[^）)]*[）)]', '', t)
    t = re.sub(r'\s+', '', t)
    return t.strip('：:').strip()


def _title_text(p):
    runs = [r.text for r in p.runs if r.text.strip() and not ep._is_red(ep._run_color(r))]
    text = ''.join(runs) if runs else _text(p)
    return text.strip()


def _vmerge(cell):
    tcpr = cell._tc.tcPr
    if tcpr is None:
        return ''
    el = tcpr.find(ep.qn('w:vMerge'))
    if el is None:
        return ''
    return el.get(ep.qn('w:val')) or 'continue'


def _bold_or_colored(p):
    for r in p.runs:
        if not r.text.strip():
            continue
        color = ep._run_color(r)
        if r.bold or (color and not ep._is_red(color) and color.upper() not in ('000000', 'AUTO')):
            return True
    return False


def _heading_info(p):
    """Return (prefix, inline, run_index) for a section sub-heading."""
    text = _text(p).strip()
    if not text or HEADING_RE.match(text) is None:
        return None
    if not _bold_or_colored(p):
        return None
    colon = min([x for x in (text.find('：'), text.find(':')) if x >= 0], default=-1)
    if 0 < colon <= 32 and text[colon + 1:].strip():
        prefix, inline = text[:colon + 1], True
    else:
        prefix, inline = text, False
    run = ep._run_index(p, want_color='example')
    if run < 0:
        run = ep._run_index(p)
    return prefix, inline, run


def _run_for(p, prefer_black=True):
    if prefer_black:
        idx = ep._run_index(p, want_color='black')
        if idx >= 0:
            return p.runs[idx]
    return p.runs[0] if p.runs else None


def _is_note(p):
    text = _text(p).strip()
    if not text:
        return False
    if NOTE_RE.search(text):
        return True
    runs = [r for r in p.runs if r.text.strip()]
    return bool(runs) and all(ep._is_red(ep._run_color(r)) for r in runs)


def _is_signature_or_date(p):
    text = _norm(_text(p))
    if not text:
        return False
    if any(w in text for w in ('签名', '签字', '手签')):
        return True
    return bool(re.search(r'\d{4}年|\d+月|\d+日|年月日', text))


def _is_dots(p):
    text = _text(p).strip()
    return bool(text) and len(text) >= 3 and re.fullmatch(r'[.。…_—\-·\s]+', text) is not None


def _bookmark_range(ti, ri, ci, start_p, start, end_p, end):
    return {'table': ti, 'row': ri, 'col': ci,
            'start_p': start_p, 'start': start, 'end_p': end_p, 'end': end}


def _sample_end(p):
    text = _text(p)
    note = NOTE_RE.search(text)
    if note and note.start() > 0:
        return note.start()
    return len(text)


def _cell_sample_range(cell, ti, ri, ci):
    paras = list(cell.paragraphs)
    start_p = end_p = 0
    start = end = 0
    found = False
    for qi, p in enumerate(paras):
        text = _text(p)
        note = NOTE_RE.search(text)
        if note:
            cut = note.start()
            if cut > 0 and text[:cut].strip():
                if not found:
                    start_p, start = qi, 0
                end_p, end = qi, cut
                found = True
            break
        if not text.strip():
            continue
        if not found:
            start_p, start = qi, 0
        end_p, end = qi, len(text)
        found = True
    return _bookmark_range(ti, ri, ci, start_p, start, end_p, end)


def _section_bookmark_range(paras, title_i, q, prefix_text, sample_indices,
                            ti, ri, ci, inline):
    samples = []
    for x in sample_indices:
        if x == q:
            samples.append(x)
            continue
        text = _text(paras[x])
        note = NOTE_RE.search(text)
        if note and note.start() == 0:
            continue
        if _is_signature_or_date(paras[x]) or _is_dots(paras[x]):
            continue
        samples.append(x)
    if samples:
        if inline:
            text = _text(paras[q])
            pos = text.find(prefix_text)
            start = (pos + len(prefix_text)) if pos >= 0 else 0
            start_p = q
        else:
            start_p, start = samples[0], 0
        end_p = samples[-1]
        end = _sample_end(paras[end_p])
        if end_p == start_p and end <= start:
            end = max(start, _sample_end(paras[start_p]))
        return _bookmark_range(ti, ri, ci, start_p, start, end_p, end)
    empties = [x for x in range(q + 1, len(paras)) if not _text(paras[x]).strip()]
    if empties:
        return _bookmark_range(ti, ri, ci, empties[0], 0, empties[0], 0)
    return _bookmark_range(ti, ri, ci, q, _sample_end(paras[q]), q, _sample_end(paras[q]))


def _whole_section_bookmark_range(paras, title_i, sample_indices, ti, ri, ci):
    if sample_indices:
        end_p = sample_indices[-1]
        return _bookmark_range(ti, ri, ci, sample_indices[0], 0, end_p, _sample_end(paras[end_p]))
    dots = [q for q in range(title_i + 1, len(paras)) if _is_dots(paras[q])]
    if dots:
        return _bookmark_range(ti, ri, ci, dots[0], 0, dots[-1], len(_text(paras[dots[-1]])))
    for q in range(title_i + 1, len(paras)):
        text = _text(paras[q])
        if not text.strip():
            continue
        if _is_signature_or_date(paras[q]) or _is_dots(paras[q]):
            continue
        runs = [r for r in paras[q].runs if r.text.strip()]
        if not runs or ep._is_red(ep._run_color(runs[0])):
            continue
        note = NOTE_RE.search(text)
        if note and note.start() >= 30 and text[:note.start()].strip():
            return _bookmark_range(ti, ri, ci, q, 0, q, note.start())
    empties = [q for q in range(title_i + 1, len(paras)) if not _text(paras[q]).strip()]
    if empties:
        return _bookmark_range(ti, ri, ci, empties[0], 0, empties[0], 0)
    return _bookmark_range(ti, ri, ci, title_i, len(_text(paras[title_i])),
                           title_i, len(_text(paras[title_i])))


def _field_style(label, context, kind='basic'):
    taskbook = (context == '任务书')
    size = 12 if taskbook and kind in ('basic', 'title', 'date') else 10.5
    return {'center': True, 'center_v': True,
            'clear_indent': kind == 'title',
            'marker_bold': True, 'marker_size': size}


def _group_for_single(label, context):
    if any(w in label for w in ('选题', '意义', '创新', '可行')):
        return '选题论证'
    if '研究内容' in label or '类型' in label:
        return '研究内容'
    if '参考' in label:
        return '参考文献'
    if any(w in label for w in ('进度', '安排')):
        return '进度安排'
    if '日期' in label or '签名' in label:
        return '落款'
    if '意见' in label or '评语' in label:
        return '指导教师意见'
    if any(w in label for w in ('来源', '开题时间')):
        return '开题信息'
    if '题目' in label:
        return '论文题目'
    if any(w in label for w in ('学院', '专业', '届次', '班次', '姓名', '学号', '指导教师', '职称')):
        return '基本信息'
    if context == '开题报告':
        return '正文'
    return '填写内容'


def _group_for_multi(context, label):
    if context == '封面':
        return '封面'
    if context == '任务书':
        if any(w in label for w in ('学院', '专业', '届次', '班次', '姓名', '学号', '指导教师', '职称')):
            return '任务书·基本信息'
        if '题目' in label:
            return '任务书·论文题目'
        if '日期' in label or '签名' in label:
            return '任务书·落款'
        return '任务书·正文'
    if context == '开题报告':
        if any(w in label for w in ('意见', '评语')):
            return '开题报告·指导教师意见'
        if any(w in label for w in ('来源', '开题时间')):
            return '开题报告·开题信息'
        if '题目' in label:
            return '开题报告·论文题目'
        if any(w in label for w in ('学院', '专业', '姓名', '指导教师')):
            return '开题报告·基本信息'
        return '开题报告·正文'
    if context == '指导记录':
        if any(w in label for w in ('评语', '成绩')):
            return '指导记录·评语成绩'
        return '指导记录'
    return context or '填写内容'


def _table_contexts(doc, title):
    n = len(doc.tables)
    if n == 1:
        if '任务书' in title:
            return [{'prefix': '', 'context': '任务书'}]
        if '开题报告' in title:
            return [{'prefix': '', 'context': '开题报告'}]
        return [{'prefix': '', 'context': '表单'}]
    if n >= 4 and ('工作情况记载表' in title or '记载表' in title):
        names = [('封面', '封面'), ('任务书', '任务书'), ('开题报告', '开题'), ('指导记录', '指导')]
        return [{'prefix': short + '_', 'context': ctx} for ctx, short in names] + \
               [{'prefix': f'表{i + 1}_', 'context': f'表{i + 1}'} for i in range(4, n)]
    return [{'prefix': f'表{i + 1}_', 'context': f'表{i + 1}'} for i in range(n)]


def _add_simple(fields, used, ti, ri, label_cell, value_cell, label_ci, value_ci,
                prefix, context):
    label = _clean_label(label_cell.text)
    if not label:
        return None
    if '题目（中文）' in label or '中文题目' in label:
        label = '中文题目'
    elif '题目（外文）' in label or '外文题目' in label:
        label = '外文题目'
    elif label.startswith('指导教师') and ('职称' in label or '姓名' in label):
        label = '指导教师'
    if any(w in label for w in SIGN_WORDS):
        return None
    if value_cell is None:
        return None
    if ep._red_only(value_cell):
        return None
    value_text = value_cell.text
    preserve_suffix = bool(
        NOTE_RE.search(value_text)
        or any(x in value_text for x in ('自选（', '自选(', '备选（', '备选(')))
    key = re.sub(r'[^\w\u4e00-\u9fff]+', '', label)[:32].strip('_') or 'field'
    base, n = key, 2
    while key in used:
        key = f'{base}{n}'
        n += 1
    used.add(key)
    fid = f'fm{len(fields) + 1}'
    kind = 'title' if '题目' in label else ('date' if '日期' in label or '时间' in label else 'basic')
    field = {
        'plan_id': fid, 'key': key, 'label': label,
        'type': 'text', 'required': True,
        'group': (_group_for_multi(context, label) if prefix else _group_for_single(label, context)),
        'mode': 'cell', 'table': ti, 'row': ri, 'col': value_ci, 'p': 0,
        'location': {'kind': 'cell', 'table': ti, 'row': ri, 'col': value_ci, 'p': 0},
        'where': f'表{ti + 1} 行{ri + 1} 列{value_ci + 1}',
        'original_content': value_cell.text.strip()[:100],
        'context': f'{context} / {label}',
        'confidence': '表单', 'source': '表单结构',
        'preserve_suffix': preserve_suffix,
        'bookmark': 'WTT_' + fid,
        'bookmark_range': _cell_sample_range(value_cell, ti, ri, value_ci),
        'style': _field_style(label, context, kind),
    }
    fields.append(field)
    return ('field', fid)


def _section_fixed_recipe(recipe, paras, title_i, body_start):
    recipe['order'] = ['title']
    for q in range(body_start, len(paras)):
        if not _text(paras[q]).strip():
            recipe['order'].append(f'blank:{q}')
            recipe['blank'].append(q)
        else:
            recipe['order'].append(f'fixed:{q}')
            recipe['fixed'].append(q)
    recipe['drop'] = [q for q in range(len(paras))
                      if q != title_i and q not in recipe['fixed'] and q not in recipe['blank']]


def _add_section(fields, used, recipes, ti, ri, ci, cell, context, prefix,
                 record_label=None):
    paras = list(cell.paragraphs)
    nonempty = [i for i, p in enumerate(paras) if _text(p).strip()]
    if not nonempty:
        return
    title_i = nonempty[0]
    title_raw = _title_text(paras[title_i])
    title = _clean_title(title_raw)
    if not title:
        return
    title_run = ep._run_index(paras[title_i])
    recipe = {
        'table': ti, 'row': ri, 'col': ci,
        'title': {'text': title, 'ppr': title_i, 'run': title_run},
        'order': ['title'], 'items': {}, 'fixed': [], 'blank': [], 'drop': [],
    }
    # fixed whole-section forms: progress schedule / main guidance work
    body_texts = [_text(p).strip() for p in paras[title_i + 1:] if _text(p).strip()]
    numbered = sum(1 for x in body_texts if NUMBERED_RE.match(x))
    fixed_section = False
    if any(k in title for k in ('写作进度安排', '指导的主要工作情况')) and numbered >= 3:
        fixed_section = True
    if any(k in title for k in ('填表说明',)) or title in ('说明',):
        fixed_section = True
    if fixed_section:
        _section_fixed_recipe(recipe, paras, title_i, title_i + 1)
        recipes.append(recipe)
        return

    # sub-headings (colored or bold): taskbook sections, opening topic reason/basic content
    headings = []
    for q in range(title_i + 1, len(paras)):
        info = _heading_info(paras[q])
        if info:
            headings.append((q, *info))
    fixed_from = None
    for q, prefix_text, inline, run_idx in headings:
        label = ep._clean_heading(prefix_text)
        if any(k in label for k in ('要求', '说明', '注意事项')) and label not in ('主要研究内容与要求',):
            fixed_from = q
            break

    made = 0
    for idx, (q, prefix_text, inline, run_idx) in enumerate(headings):
        if fixed_from is not None and q >= fixed_from:
            break
        label = ep._clean_heading(prefix_text)
        if not label:
            continue
        nxt = headings[idx + 1][0] if idx + 1 < len(headings) else len(paras)
        sample_indices = []
        if inline and _text(paras[q]).split(prefix_text, 1)[-1].strip():
            sample_indices.append(q)
        sample_indices += [x for x in range(q + 1, nxt) if _text(paras[x]).strip()]
        body_ppr = q if inline else (sample_indices[0] if sample_indices else q)
        body_run = _run_for(paras[body_ppr])
        if body_run is None:
            body_run = paras[q].runs[run_idx] if run_idx >= 0 and run_idx < len(paras[q].runs) else None
        key = re.sub(r'[^\w\u4e00-\u9fff]+', '', label)[:32].strip('_') or 'field'
        base, n = key, 2
        while key in used:
            key = f'{base}{n}'
            n += 1
        used.add(key)
        fid = f'fm{len(fields) + 1}'
        field = {
            'plan_id': fid, 'key': key, 'label': label,
            'type': 'text' if any(k in label for k in ('类型', '成绩', '日期')) else 'multiline',
            'required': True,
            'group': (_group_for_multi(context, label) if prefix else _group_for_single(label, context)),
            'mode': 'section', 'table': ti, 'row': ri, 'col': ci, 'p': q,
            'location': {'kind': 'cell', 'table': ti, 'row': ri, 'col': ci, 'p': q},
            'where': f'表{ti + 1} 行{ri + 1} 整行',
            'original_content': '\n'.join(_text(paras[x]).strip() for x in sample_indices)[:100],
            'context': f'{title} / {prefix_text}',
            'confidence': '表单', 'source': '表单结构',
            'style': _field_style(label, context, 'title' if '题目' in label else 'basic'),
            'recipe_field': fid,
            'bookmark': 'WTT_' + fid,
            'bookmark_range': _section_bookmark_range(
                paras, title_i, q, prefix_text, sample_indices, ti, ri, ci, inline),
        }
        fields.append(field)
        recipe['items'][fid] = {
            'plan_id': fid, 'heading': prefix_text, 'inline': inline,
            'ppr': q, 'body_ppr': body_ppr, 'heading_run': run_idx,
            'body_run': ep._run_index(paras[body_ppr], want_color='black') if body_ppr != q else 0,
            'body_size_pt': 10.5, 'heading_size_pt': 10.5, 'align': 'left',
        }
        recipe['order'].append('item:' + fid)
        made += 1

    if made:
        used_paras = {title_i}
        used_paras |= {q for q, _p, _i, _r in headings}
        for item in recipe['items'].values():
            used_paras.add(item['body_ppr'])
        if fixed_from is not None:
            for q in range(fixed_from, len(paras)):
                if not _text(paras[q]).strip():
                    recipe['order'].append(f'blank:{q}')
                    recipe['blank'].append(q)
                else:
                    recipe['order'].append(f'fixed:{q}')
                    recipe['fixed'].append(q)
                used_paras.add(q)
        # Keep notes, signatures, dates and empty paragraphs; drop only sample prose/dots.
        for q in range(title_i + 1, len(paras)):
            if q in used_paras:
                continue
            if not _text(paras[q]).strip():
                recipe['order'].append(f'blank:{q}')
                recipe['blank'].append(q)
            elif _is_dots(paras[q]):
                continue
            elif _is_note(paras[q]) or _is_signature_or_date(paras[q]):
                recipe['order'].append(f'fixed:{q}')
                recipe['fixed'].append(q)
        recipes.append(recipe)
        recipe['drop'] = [q for q in range(len(paras))
                          if q != title_i and all(f':{q}' not in x for x in recipe['order'])]
        return

    # no sub-headings: one field for the whole section
    label = record_label or title
    if not record_label and any(k in title for k in ('成绩',)):
        label = '论文成绩'
    if not record_label and '指导教师对毕业论文' in title:
        label = '指导教师评语'
    sample_indices = [
        q for q in range(title_i + 1, len(paras))
        if _text(paras[q]).strip() and not _is_note(paras[q])
        and not _is_signature_or_date(paras[q]) and not _is_dots(paras[q])
    ]
    if sample_indices:
        body_ppr = sample_indices[0]
    else:
        empties = [q for q in range(title_i + 1, len(paras)) if not _text(paras[q]).strip()]
        dots = [q for q in range(title_i + 1, len(paras)) if _is_dots(paras[q])]
        body_ppr = (empties[0] if empties else (dots[0] if dots else title_i))
    body_run = _run_for(paras[body_ppr])
    if body_run is None:
        body_run = paras[title_i].runs[title_run] if title_run >= 0 and title_run < len(paras[title_i].runs) else None
    key = re.sub(r'[^\w\u4e00-\u9fff]+', '', label)[:32].strip('_') or 'field'
    base, n = key, 2
    while key in used:
        key = f'{base}{n}'
        n += 1
    used.add(key)
    fid = f'fm{len(fields) + 1}'
    ftype = 'text' if any(k in label for k in ('成绩', '日期', '签名', '时间', '来源')) else 'multiline'
    field = {
        'plan_id': fid, 'key': key, 'label': label, 'type': ftype,
        'required': True,
        'group': (_group_for_multi(context, label) if prefix else _group_for_single(label, context)),
        'mode': 'section', 'table': ti, 'row': ri, 'col': ci, 'p': title_i,
        'location': {'kind': 'cell', 'table': ti, 'row': ri, 'col': ci, 'p': title_i},
        'where': f'表{ti + 1} 行{ri + 1} 整行',
        'original_content': '\n'.join(_text(paras[q]).strip() for q in sample_indices)[:100],
        'context': title,
        'confidence': '表单', 'source': '表单结构',
        'style': _field_style(label, context, 'basic'),
        'recipe_field': fid,
        'bookmark': 'WTT_' + fid,
        'bookmark_range': _whole_section_bookmark_range(
            paras, title_i, sample_indices, ti, ri, ci),
    }
    fields.append(field)
    recipe['items'][fid] = {
        'plan_id': fid, 'heading': '', 'inline': False,
        'ppr': title_i, 'body_ppr': body_ppr,
        'heading_run': -1, 'body_run': 0,
        'body_size_pt': 10.5, 'align': 'left',
    }
    recipe['order'].append('item:' + fid)
    for q in range(title_i + 1, len(paras)):
        if q == body_ppr:
            continue
        if not _text(paras[q]).strip():
            recipe['order'].append(f'blank:{q}')
            recipe['blank'].append(q)
        elif _is_dots(paras[q]):
            continue
        elif _is_note(paras[q]) or _is_signature_or_date(paras[q]):
            recipe['order'].append(f'fixed:{q}')
            recipe['fixed'].append(q)
    recipes.append(recipe)
    recipe['drop'] = [q for q in range(len(paras))
                      if q != title_i and all(f':{q}' not in x for x in recipe['order'])]


def analyze(path):
    path = str(path)
    doc = Document(path)
    if not doc.tables:
        raise ValueError('文档没有表格，未发现可填写栏目。\n'
                         '这更像说明/规范文档，不应作为模板录入。')
    title = next((p.text.strip() for p in doc.paragraphs if p.text.strip()), '')
    contexts = _table_contexts(doc, title)
    fields, recipes, clear_cells = [], [], []
    used = set()

    for ti, table in enumerate(doc.tables):
        ctx = contexts[ti] if ti < len(contexts) else {'prefix': f'表{ti + 1}_', 'context': f'表{ti + 1}'}
        prefix, context = ctx['prefix'], ctx['context']
        handled_rows = set()
        for ri, row in enumerate(table.rows):
            cells = ep._unique_cells(row)
            if not cells:
                continue
            span = sum(ep._grid_span(c) for _, c in cells)
            # guidance record rows: 第一次/第二次 ... + 指导意见 body
            if len(cells) >= 2:
                first_text = _norm(cells[0][1].text)
                if re.fullmatch(r'第[一二三四五六七八九十]+次', first_text) and \
                        '指导意见' in (cells[1][1].text or ''):
                    label = first_text + '指导意见'
                    _add_section(fields, used, recipes, ti, ri, cells[1][0], cells[1][1],
                                 '指导记录', prefix, record_label=label)
                    handled_rows.add(ri)
                    continue
            full = len(cells) == 1 and span >= max(3, len(table.columns) - 1)
            if full:
                ci, cell = cells[0]
                _add_section(fields, used, recipes, ti, ri, ci, cell, context, prefix)
                handled_rows.add(ri)
                continue
            # label/value pairs
            start = 0
            if len(cells) >= 3 and _vmerge(cells[0][1]) == 'restart':
                start = 1
            i = start
            while i < len(cells):
                _, label_cell = cells[i]
                if not _norm(label_cell.text):
                    i += 1
                    continue
                if i + 1 >= len(cells):
                    break
                value_ci, value_cell = cells[i + 1]
                result = _add_simple(fields, used, ti, ri, label_cell, value_cell,
                                     cells[i][0], value_ci, prefix, context)
                if result and result[0] == 'clear':
                    clear_cells.append([result[1], result[2], result[3]])
                i += 2

    if not fields:
        raise ValueError('没有识别到可填写栏目。\n'
                         '这份文档可能只是规范/说明材料，没有可填位。')
    # remove duplicate recipe fields keys? keep as planned
    group_order = {'封面': 0, '任务书·基本信息': 1, '任务书·论文题目': 2, '任务书·正文': 3,
                   '任务书·落款': 4, '开题报告·基本信息': 5, '开题报告·论文题目': 6,
                   '开题报告·开题信息': 7, '开题报告·正文': 8, '开题报告·指导教师意见': 9,
                   '指导记录': 10, '指导记录·评语成绩': 11,
                   '基本信息': 0, '论文题目': 1, '开题信息': 2, '选题论证': 3,
                   '研究内容': 4, '进度安排': 5, '参考文献': 6, '正文': 7,
                   '指导教师意见': 8, '落款': 9, '填写内容': 10, '其他': 11}
    fields.sort(key=lambda f: (group_order.get(f['group'], 50),
                               int(f['plan_id'][2:]) if f['plan_id'][2:].isdigit() else 0))
    description = (f'由表单结构识别生成，共 {len(fields)} 个填写项；'
                   '空白大段按标题生成正文项，示例正文和红色说明已清理，固定栏目保留。')
    return {
        'source': path, 'title': title or doc.core_properties.title or '未命名模板',
        'category': '教学' if any(k in title for k in ('任务书', '开题报告', '论文', '学院')) else '自定义模板',
        'description': description, 'table': 0,
        'fields': fields, 'recipes': recipes, 'clear_cells': clear_cells,
        'row_heights': {}, 'remove_textboxes': False,
    }
