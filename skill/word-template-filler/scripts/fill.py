"""Fill a prepared template from fields.json + values.json.

Usage:
    python fill.py fields.json values.json --out output.docx
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / 'vendor'))
sys.path.insert(0, str(HERE / 'wttlib'))

import engine  # noqa: E402


def load_values(path):
    raw = json.loads(Path(path).read_text(encoding='utf-8'))
    if isinstance(raw, dict) and isinstance(raw.get('values'), dict):
        raw = raw['values']
    if not isinstance(raw, dict):
        raise ValueError('values.json 必须是 JSON 对象')
    return raw


def main():
    ap = argparse.ArgumentParser(description='把填写内容写入已识别的 Word 模板。')
    ap.add_argument('fields', help='analyze.py 生成的 fields.json')
    ap.add_argument('values', help='填写内容 JSON')
    ap.add_argument('--out', required=True, help='输出 .docx')
    ap.add_argument('--strict', action='store_true', help='必填项为空时拒绝生成（默认只提醒）')
    args = ap.parse_args()

    data = json.loads(Path(args.fields).read_text(encoding='utf-8'))
    config = data.get('config') or {}
    fields = config.get('fields') or []
    if not fields:
        print('fields.json 中没有 config.fields，请重新运行 analyze.py。', file=sys.stderr)
        return 2
    values = load_values(args.values)

    alias = {}
    for field in fields:
        alias[field['key']] = field['key']
        label = str(field.get('label') or '').strip()
        if label:
            alias.setdefault(label, field['key'])
    normalized = {}
    unknown = []
    for key, value in values.items():
        if key in alias:
            normalized[alias[key]] = value
        else:
            unknown.append(key)

    missing = [f for f in fields if f.get('required') and not str(normalized.get(f['key'], '')).strip()]
    if missing:
        names = '、'.join(f'{f["key"]}（{f.get("label")}）' for f in missing)
        if args.strict:
            print('缺少必填项：' + names, file=sys.stderr)
            return 3
        print('提示：以下必填项为空，仍会生成：' + names)

    template = Path(data.get('template') or '')
    if not template.is_absolute():
        template = Path(args.fields).resolve().parent / template
    output = Path(args.out).expanduser().resolve()
    try:
        engine.render_word(template, output, config, normalized,
                           validate=not args.strict, allow_missing=True,
                           keep_empty_bookmarks=True)
    except Exception as exc:
        print(f'生成失败：{exc}', file=sys.stderr)
        return 2
    if unknown:
        print('提示：这些键没有对应字段，已忽略：' + '、'.join(unknown))
    print(f'已生成：{output}')
    print(f'填写项：{len(normalized)} / {len(fields)}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
