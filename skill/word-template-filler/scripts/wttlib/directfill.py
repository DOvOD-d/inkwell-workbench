"""Turn any .docx into a fillable template without asking for placeholders."""
from __future__ import annotations

from pathlib import Path

import bookmarkfill
import smartparse


def config_fields(plan):
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
        fields.append(item)
    return fields


def build_from_plan(plan, source, out_dir, title=None, category=None, description=None):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    selected = {f['plan_id']: f['key'] for f in plan.get('fields', [])}
    bookmarkfill.build(str(source), str(out_dir / 'template.docx'), plan, selected)
    return {
        'id': '', 'title': (title or plan.get('title') or Path(source).stem)[:60],
        'category': category or plan.get('category') or '自定义模板',
        'description': description or plan.get('description') or '自动识别填写位；原文与版式保持不变。',
        'file': 'template.docx', 'fields': config_fields(plan),
        'grids': [], 'images': [],
    }


def prepare(source, out_dir, title=None, category=None):
    plan = smartparse.analyze(str(source))
    config = build_from_plan(plan, source, out_dir, title, category)
    return plan, config
