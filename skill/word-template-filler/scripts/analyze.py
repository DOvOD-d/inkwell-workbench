"""Analyze any .docx and emit a fields.json + prepared template pair.

Usage:
    python analyze.py input.docx [--out workdir]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / 'vendor'))
sys.path.insert(0, str(HERE / 'wttlib'))

from docx import Document  # noqa: E402

import directfill  # noqa: E402


def outline(path, limit=300):
    doc = Document(str(path))
    lines = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            lines.append(text)
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            cells = []
            for cell in row.cells:
                text = cell.text.strip().replace('\n', ' / ')
                if text and text not in cells:
                    cells.append(text)
            if cells:
                lines.append(f'[表{ti + 1} 行{ri + 1}] ' + ' | '.join(cells))
            if len(lines) >= limit:
                return lines[:limit]
    return lines[:limit]


def main():
    ap = argparse.ArgumentParser(description='识别 Word 模板的填写位，不需要占位符。')
    ap.add_argument('input', help='输入 .docx')
    ap.add_argument('--out', help='工作目录（默认：输入文件同名的 .wtt 文件夹）')
    ap.add_argument('--quiet', action='store_true')
    args = ap.parse_args()

    source = Path(args.input).expanduser().resolve()
    if not source.is_file():
        print(f'文件不存在：{source}', file=sys.stderr)
        return 2
    if source.suffix.lower() != '.docx':
        print('只支持 .docx；旧版 .doc 请先用 Word 另存为 .docx。', file=sys.stderr)
        return 2
    out_dir = Path(args.out).expanduser().resolve() if args.out else \
        source.parent / (source.stem + '.wtt')

    try:
        plan, config = directfill.prepare(source, out_dir)
    except Exception as exc:
        print(f'识别失败：{exc}', file=sys.stderr)
        return 2

    config['id'] = 'skill-' + source.stem
    fields = []
    for item in config['fields']:
        field = {
            'key': item['key'], 'label': item.get('label', item['key']),
            'type': item.get('type', 'text'), 'required': bool(item.get('required', True)),
            'group': item.get('group', '填写内容'),
        }
        fields.append(field)

    detail = {field['key']: field for field in fields}
    for orig in plan.get('fields', []):
        if orig.get('key') in detail:
            detail[orig['key']]['where'] = orig.get('where', '')
            detail[orig['key']]['context'] = (orig.get('context') or '')[:160]
            detail[orig['key']]['sample'] = (orig.get('original_content') or '')[:160]

    data = {
        'version': 1,
        'source': str(source),
        'title': config['title'],
        'category': config['category'],
        'description': config['description'],
        'template': str(out_dir / 'template.docx'),
        'values': {field['key']: '' for field in fields},
        'fields': fields,
        'outline': outline(source),
        'config': config,
    }
    target = out_dir / 'fields.json'
    target.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    if not args.quiet:
        print(f'识别完成：{len(fields)} 个填写项')
        print(f'工作目录：{out_dir}')
        print(f'字段清单：{target}')
        for field in fields:
            mark = '*' if field['required'] else ' '
            print(f'  {mark} {field["key"]}  [{field["type"]}/{field["group"]}] {field["label"]}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
