"""Turn any .docx into a fillable template without asking for placeholders."""
from __future__ import annotations

from pathlib import Path

import bookmarkfill
import reqparse
import smartparse
import textparse
import visionparse


def config_fields(plan):
    hints = ((plan.get('req') or {}).get('hints')) or {}
    fields = []
    for field in plan.get('fields', []):
        item = {
            'key': field['key'], 'label': field.get('label') or field['key'],
            'type': field.get('type', 'text'), 'required': bool(field.get('required', False)),
            'group': field.get('group') or '填写内容', 'max_length': 0, 'default': '',
        }
        if field.get('bookmark'):
            item['bookmark'] = field['bookmark']
        if field.get('control'):
            item['control'] = field['control']
        if field.get('choices'):
            item['choices'] = field['choices']
        if field.get('options'):
            item['options'] = field['options']
        if field.get('grid'):
            item['grid'] = field['grid']
        moved = hints.get(field['key'])
        if moved:
            item['hint'] = '模板要求：' + '；'.join(str(t).strip() for t in moved[:3])
        fields.append(item)
    return fields


def build_from_plan(plan, source, out_dir, title=None, category=None, description=None):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    selected = {f['plan_id']: f['key'] for f in plan.get('fields', [])}
    bookmarkfill.build(str(source), str(out_dir / 'template.docx'), plan, selected)
    config = {
        'id': '', 'title': (title or plan.get('title') or Path(source).stem)[:60],
        'category': category or plan.get('category') or '自定义模板',
        'description': description or plan.get('description') or '自动识别填写位；原文与版式保持不变。',
        'file': 'template.docx', 'fields': config_fields(plan),
        'grids': plan.get('grids') or [], 'images': [],
    }
    if plan.get('req'):
        config['req'] = plan['req']
        config['req_options'] = {'apply_format': True, 'remove_notes': True}
    if plan.get('req_summary'):
        config['req_summary'] = plan['req_summary']
    return config


def prepare(source, out_dir, title=None, category=None, text_cfg=None, vision_cfg=None,
            auto_format=True, progress=None):
    structural_error = ''
    try:
        plan = smartparse.analyze(str(source))
    except Exception as exc:
        # 规范/说明材料是明确拒绝的，不能因为配了模型就绕过去再猜一遍
        if not text_cfg or '规范/说明材料' in str(exc):
            raise
        structural_error = str(exc)
        plan = visionparse.empty_plan(source, structural_error)
    if text_cfg:
        if progress:
            progress('文本模型：正在通读文档结构，判断哪些位置要填写…')
        plan = textparse.enhance(str(source), plan, text_cfg, progress=progress)
    if vision_cfg:
        if progress:
            progress('视觉复核：正在核对有疑问的栏目…')
        plan = visionparse.verify(str(source), plan, vision_cfg, progress=progress)
    if auto_format:
        if progress:
            progress('正在读取模板自己写明的格式要求…')
        plan = reqparse.enhance(str(source), plan, progress=progress)
    if not plan.get('fields'):
        raise ValueError(
            '结构识别和文本模型都没有找到可填写栏目。\n'
            '如果这份文档确实有填写位，请确认 AI 设置里的文本模型可用，'
            '或在 Word 中补一个栏目名/下划线后再识别。')
    if structural_error:
        plan['structural_error'] = structural_error
    config = build_from_plan(plan, source, out_dir, title, category)
    return plan, config
