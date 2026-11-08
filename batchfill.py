# -*- coding: utf-8 -*-
"""一个名单 → N 份填好的 Word。

做法参照 OfficeCLI 的 template merge：模板只做一次，数据表负责批量。
数据文件支持 .csv / .txt(制表符) / .json（对象数组）/ .xlsx（标准库直接解 zip，不依赖 openpyxl）。
"""
from __future__ import annotations

import csv
import io
import json
import re
import zipfile
from pathlib import Path

MAX_COLUMNS = 64


def _norm(text):
    return re.sub(r'[\s\u3000（）()【】\[\]:：，,、.。]', '', str(text or ''))


def read_table(path):
    """→ (表头, 行列表)。行是 {表头: 值}，全空行丢掉。"""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == '.json':
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        if isinstance(data, dict):
            data = data.get('rows') or data.get('data') or []
        rows = [dict(item) for item in data if isinstance(item, dict)]
        headers = []
        for row in rows:
            for key in row:
                if key not in headers:
                    headers.append(key)
        return headers, rows
    if suffix == '.xlsx':
        return _read_xlsx(path)
    text = path.read_text(encoding='utf-8-sig', errors='replace')
    dialect = csv.excel_tab if '\t' in text.splitlines()[0] else csv.excel
    reader = csv.reader(io.StringIO(text), dialect)
    rows = [row for row in reader if any(str(cell).strip() for cell in row)]
    if not rows:
        return [], []
    headers = [str(cell).strip() for cell in rows[0]]
    records = []
    for row in rows[1:]:
        record = {}
        for index, header in enumerate(headers):
            if header:
                record[header] = str(row[index]).strip() if index < len(row) else ''
        records.append(record)
    return headers, records


SS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def _read_xlsx(path):
    """只读第一张工作表：sharedStrings + sheet，够用且不引新依赖。"""
    from lxml import etree
    with zipfile.ZipFile(str(path)) as archive:
        members = archive.namelist()
        shared = []
        if 'xl/sharedStrings.xml' in members:
            root = etree.fromstring(archive.read('xl/sharedStrings.xml'))
            for item in root.findall(SS + 'si'):
                shared.append(''.join(node.text or '' for node in item.iter(SS + 't')))
        name = next((entry for entry in members
                     if entry.startswith('xl/worksheets/sheet')), None)
        if name is None:
            return [], []
        sheet = etree.fromstring(archive.read(name))
        rows = []
        for row in sheet.iter(SS + 'row'):
            cells = {}
            for cell in row.iter(SS + 'c'):
                reference = cell.get('r') or ''
                match = re.match(r'([A-Z]+)', reference)
                if match is None:
                    continue
                value = ''
                holder = cell.find(SS + 'v')
                if holder is not None:
                    value = holder.text or ''
                    if (cell.get('t') or '') == 's':
                        try:
                            value = shared[int(value)]
                        except (IndexError, ValueError):
                            value = ''
                else:
                    inline = cell.find(SS + 'is')
                    if inline is not None:
                        value = ''.join(node.text or '' for node in inline.iter(SS + 't'))
                cells[_column_index(match.group(1))] = (value or '').strip()
            rows.append([cells.get(i, '') for i in range(max(cells or {0: 0}) + 1)])
    rows = [row for row in rows if any(row)]
    if not rows:
        return [], []
    headers = [str(cell).strip() for cell in rows[0]]
    records = []
    for row in rows[1:]:
        record = {}
        for index, header in enumerate(headers):
            if header:
                record[header] = row[index] if index < len(row) else ''
        records.append(record)
    return headers, records


def _column_index(letters):
    value = 0
    for char in letters:
        value = value * 26 + (ord(char) - ord('A') + 1)
    return value - 1


def match_columns(headers, fields):
    """表头 → 栏目 key 的映射：先去标点比全等，再比包含关系。"""
    by_header = {}
    keys = {_norm(field['key']): field['key'] for field in fields}
    labels = {}
    for field in fields:
        labels.setdefault(_norm(field.get('label') or field['key']), field['key'])
    for header in headers:
        norm = _norm(header)
        if not norm:
            continue
        if norm in keys:
            by_header[header] = keys[norm]
        elif norm in labels:
            by_header[header] = labels[norm]
        else:
            hit = [value for norm_key, value in list(keys.items()) + list(labels.items())
                   if len(norm_key) >= 2 and (norm_key in norm or norm in norm_key)]
            if len(set(hit)) == 1:
                by_header[header] = hit[0]
    return by_header


def run(template_docx, output_dir, config, headers, records, mapping, name_column=None,
        progress=None, should_stop=None):
    """逐行生成，返回 {done, failed, skipped, files, problems, cancelled}。

    should_stop 每做完一份问一次：取消只停在未来的一份上，已经写出去的文件不回滚。
    """
    from engine import render_word
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    keys = {field['key'] for field in config['fields']}
    report = {'done': 0, 'failed': [], 'skipped': [], 'files': [], 'problems': [],
              'cancelled': 0}
    unmatched = [header for header in headers if header not in mapping]
    if unmatched:
        report['problems'].append('这些列没有对应栏目，已忽略：' + '、'.join(unmatched))
    missing = sorted(keys - {mapping[header] for header in mapping if header in mapping})
    if missing:
        report['problems'].append('这些栏目在名单里没有对应列，会留空：'
                                  + '、'.join(_display(config, key) for key in missing[:8]))
    for number, record in enumerate(records, 1):
        if should_stop and should_stop():
            report['cancelled'] = len(records) - number + 1
            break
        if not any(str(value).strip() for value in record.values()):
            report['skipped'].append(number)
            continue
        values = {field['key']: '' for field in config['fields']}
        for header, field_key in mapping.items():
            values[field_key] = str(record.get(header, '') or '')
        name = _file_name(record, headers, name_column, number)
        target = output_dir / f'{name}.docx'
        if progress:
            progress(f'第 {number}/{len(records)} 份：{target.name}')
        try:
            render_word(Path(template_docx), target, config, values,
                        validate=False, allow_missing=True)
        except Exception as exc:
            report['failed'].append((name, str(exc)[:120]))
            continue
        report['files'].append(str(target))
        report['done'] += 1
    return report


def _display(config, key):
    for field in config['fields']:
        if field['key'] == key:
            return field.get('label') or key
    return key


def _file_name(record, headers, name_column, number):
    parts = []
    for header in (name_column, headers[0] if headers else None):
        if header and str(record.get(header, '')).strip():
            parts.append(str(record[header]).strip())
            break
    base = parts[0] if parts else f'第{number}份'
    cleaned = re.sub(r'[\\/:*?"<>|]', '_', base)[:60]
    return f'{number:02d}_{cleaned}'
