"""Basic post-fill verification for a generated .docx.

Usage:
    python verify.py output.docx [--values values.json] [--fields fields.json]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / 'vendor'))
sys.path.insert(0, str(HERE / 'wttlib'))

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

PLACEHOLDER = re.compile(r'单击此处|点击此处|请输入|请填写|在此输入|此处输入|placeholder|click here', re.I)


def all_text(doc):
    chunks = []
    for node in doc.element.body.iter(qn('w:t')):
        chunks.append(node.text or '')
    for section in doc.sections:
        for part in (section.header, section.footer):
            try:
                for node in part._element.iter(qn('w:t')):
                    chunks.append(node.text or '')
            except Exception:
                pass
    return '\n'.join(chunks)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--values')
    ap.add_argument('--fields')
    ap.add_argument('--strict', action='store_true',
                    help='把残留占位提示也视为失败（默认只提示）')
    args = ap.parse_args()
    path = Path(args.docx)
    if not path.is_file():
        print('文件不存在：' + str(path), file=sys.stderr)
        return 2
    doc = Document(str(path))
    text = all_text(doc)
    placeholders = sorted(set(PLACEHOLDER.findall(text)))
    missing = []
    if args.values:
        values = json.loads(Path(args.values).read_text(encoding='utf-8'))
        if isinstance(values, dict) and isinstance(values.get('values'), dict):
            values = values['values']
        for key, value in values.items():
            probe = re.sub(r'\s+', '', str(value))[:30]
            if probe and probe not in re.sub(r'\s+', '', text):
                missing.append(key)
    result = {
        'ok': not missing and (not args.strict or not placeholders),
        'paragraphs': len(doc.paragraphs),
        'tables': len(doc.tables),
        'placeholders_left': placeholders,
        'values_not_found': missing,
    }
    if placeholders and not args.strict:
        result['note'] = '未填字段保留了模板原占位提示；如要求全部填写，请使用 --strict。'
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result['ok'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
