"""五阶段 Word 模板流水线：侦察 → 候选 → (人工确认) → 白名单写入 → 分层验收。

只依赖 python-docx，对任意 .docx 通用。各阶段输入输出都是 JSON 文件，
Phase 3（确认）由人/助手对白名单 JSON 做删改后交给 Phase 4。
用法见同目录 README-五阶段流水线.md。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as ET

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn

import cleanup
import ooxmlcheck

TOKEN = re.compile(r"\{\{\s*([^{}\r\n]{1,100}?)\s*\}\}", re.UNICODE)
LIMIT_RE = re.compile(r"限\s*(\d+)\s*字")
COLONS = ("：", ":")


# ---------------------------------------------------------------- helpers

def norm(s):
    return re.sub(r"\s+", "", s or "")


def para_text(p):
    return "".join(r.text for r in p.runs)


def token_positions(doc):
    """结构事实：每个 {{ key }} 标记出现在哪个可定位容器（顶层段落/单元格/页眉页脚）。"""
    out = []
    for i, p in enumerate(doc.paragraphs):
        for m in TOKEN.finditer(para_text(p)):
            out.append({"key": m.group(1).strip(), "target": {"kind": "paragraph", "p": i},
                        "runs": len(p.runs), "cross_run": True})
    for ti, t in enumerate(doc.tables):
        for ri, row in enumerate(t.rows):
            seen = set()
            for ci, cell in enumerate(row.cells):
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                for pi, p in enumerate(cell.paragraphs):
                    text = para_text(p)
                    for m in TOKEN.finditer(text):
                        out.append({"key": m.group(1).strip(),
                                    "target": {"kind": "cell", "table": ti, "row": ri,
                                               "col": ci, "p": pi},
                                    "runs": len(p.runs),
                                    "cross_run": _spans_runs(p, m)})
    return out


def _spans_runs(p, m):
    off = 0
    for r in p.runs:
        if off < m.end() and off + len(r.text) > m.start():
            if off <= m.start() and off + len(r.text) >= m.end():
                return len([x for x in p.runs
                            if m.start() < off + len(r.text) and off < m.end() - len(r.text)]) > 0
        off += len(r.text)
    return True


def v_merge_of(tc):
    el = tc.tcPr.find(qn("w:vMerge")) if tc.tcPr is not None else None
    if el is None:
        return None
    val = el.get(qn("w:val"))
    return val or "continue"


def grid_span_of(tc):
    el = tc.tcPr.find(qn("w:gridSpan")) if tc.tcPr is not None else None
    return int(el.get(qn("w:val"))) if el is not None else 1


# ---------------------------------------------------------------- phase 1

def phase_recon(docx_path):
    doc = Document(docx_path)
    body = doc.element.body
    n_sdt = sum(1 for e in body.iter() if e.tag == qn("w:sdt"))
    n_fld = sum(1 for e in body.iter() if e.tag in (qn("w:fldChar"), qn("w:instrText")))
    hf = {}
    for si, sec in enumerate(doc.sections):
        for role, part in (("header", sec.header), ("footer", sec.footer)):
            if part.is_linked_to_previous:
                continue
            txt = "\n".join(para_text(p) for p in part.paragraphs)
            hf[f"s{si}.{role}"] = {"chars": len(txt.strip()), "tokens": [m.group(1).strip()
                                                                           for m in TOKEN.finditer(txt)]}
    tables = []
    for ti, t in enumerate(doc.tables):
        cells, seen, occ = [], {}, {}
        for ri, row in enumerate(t.rows):
            for ci, cell in enumerate(row.cells):
                tc = cell._tc
                key = seen.get(tc)
                if key is None:
                    key = len(seen)
                    seen[tc] = key
                    runs = sum(len(p.runs) for p in cell.paragraphs)
                    cells.append({"row": ri, "col": ci, "tc": key,
                                  "grid_span": grid_span_of(tc),
                                  "v_merge": v_merge_of(tc),
                                  "text": (cell.text or "")[:60],
                                  "blank": not cell.text.strip(),
                                  "ends_with_colon": cell.text.rstrip().endswith(COLONS),
                                  "runs": runs,
                                  "paras": len(cell.paragraphs)})
                occ.setdefault(key, []).append([ri, ci])
        merged = [{"anchor": occ[k][0], "occurrences": occ[k],
                   "grid_span": next((c["grid_span"] for c in cells if c["tc"] == k), 1),
                   "v_merge": next((c["v_merge"] for c in cells if c["tc"] == k), None)}
                  for k in occ if len(occ[k]) > 1]
        tables.append({"index": ti, "rows": len(t.rows), "cols": len(t.columns),
                       "unique_tc": len(seen), "cells": cells, "merged_regions": merged})
    paras = []
    for i, p in enumerate(doc.paragraphs):
        text = para_text(p)
        fmt_only = [j for j, r in enumerate(p.runs)
                    if not r.text and r._element.find(qn("w:rPr")) is not None]
        has_image = any(e.tag in (qn("w:drawing"), qn("w:pict"), qn("w:object"))
                        for e in p._element.iter())
        paras.append({"index": i, "empty": not text.strip(), "text": text[:80],
                      "ends_with_colon": text.rstrip().endswith(COLONS),
                      "runs": len(p.runs), "format_only_runs": fmt_only,
                      "has_image": has_image})
    import engine as _engine
    anchors = []
    for a in _engine.scan_image_anchors(docx_path):
        anchors.append({"tag": a["tag"], "media": a["media"], "fmt": a["fmt"],
                        "part": a["part"], "relationship": a["relationship"]})
    with zipfile.ZipFile(docx_path) as z:
        names = z.namelist()
        parts = {"headers": [n for n in names if re.fullmatch(r"word/header\d*\.xml", n)],
                 "footers": [n for n in names if re.fullmatch(r"word/footer\d*\.xml", n)],
                 "media": [n for n in names if n.startswith("word/media/")],
                 "settings": "word/settings.xml" in names}
    return {"file": str(docx_path), "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables), "section_count": len(doc.sections),
            "has_header_footer": bool(hf), "header_footer": hf,
            "content_controls": n_sdt, "form_fields": n_fld,
            "package": parts, "tables": tables, "paragraphs": paras,
            "markers": token_positions(doc), "image_anchors": anchors}


# ---------------------------------------------------------------- phase 2

def _label_from(text):
    t = text.strip().rstrip("＊*")
    for c in COLONS:
        t = t.replace(c, "")
    return t.strip()[:20]


def phase_candidates(recon, requirement=""):
    tokens = {}
    for m in recon["markers"]:
        tokens.setdefault(m["key"], m)
    fields = []
    used_cells = set()

    def add(loc, label, orig, auth, conf, cap="", merge="", cross=False, source="结构"):
        base = re.sub(r"[^\w\u4e00-\u9fff]+", "_", label or loc["kind"])[:24].strip("_") or "field"
        fid, n = base, 2
        while any(f["field_id"] == fid for f in fields):
            fid = f"{base}_{n}"
            n += 1
        fields.append({"field_id": fid, "label": label or fid, "location": loc,
                       "original_content": (orig or "")[:60], "authorization_type": auth,
                       "capacity_note": cap, "merge_flag": merge, "cross_run_flag": cross,
                       "confidence": conf, "source": source})
        return fid

    anchor_tags = {a["tag"]: a for a in recon.get("image_anchors", [])}

    for key, m in tokens.items():
        loc = m["target"]
        cap = ""
        merge = ""
        if loc["kind"] == "cell":
            cell = find_cell(recon, loc["table"], loc["row"], loc["col"])
            if cell is not None:
                merge = _merge_note(recon, loc["table"], cell)
                cap = _cap_note(recon, loc["table"], cell)
        add(loc, key, "{{" + key + "}}", "原样", "高", cap, merge, m["cross_run"], "标记")
        used_cells.add((loc.get("table"), loc.get("row"), loc.get("col")))

    def marker_anchor(m):
        loc = m["target"]
        if loc.get("kind") != "cell":
            return None
        c = find_cell(recon, loc["table"], loc["row"], loc["col"])
        if c is None:
            return None
        return (loc["table"], c["row"], c["col"])

    marked_anchors = {marker_anchor(m) for m in recon["markers"]} - {None}

    for t in recon["tables"]:
        ti = t["index"]
        for cell in t["cells"]:
            if cell["blank"] and (ti, cell["row"], cell["col"]) not in marked_anchors:
                lab = _left_label(t, cell, ti) or _up_label(t, cell, ti) or _row_label(t, cell)
                if lab:
                    add({"kind": "cell", "table": ti, "row": cell["row"], "col": cell["col"], "p": 0},
                        _label_from(lab), "", "不确定",
                        "中" if lab else "低", _cap_note(recon, ti, cell),
                        _merge_note(recon, ti, cell))

    # 段落标签：冒号结尾段落（尾部空格即填写位）；或标签段后紧跟空段
    pend = None
    for p in recon["paragraphs"]:
        if p["ends_with_colon"]:
            tail_ws = p["text"] != p["text"].rstrip()
            add({"kind": "paragraph", "p": p["index"]}, _label_from(p["text"]), "",
                "不确定", "中" if tail_ws else "低",
                "填写位为段落尾部空格" if tail_ws else "标签段，填写位待确认")
            pend = None if tail_ws else p
            continue
        if pend is not None:
            if p["empty"] and not p.get("has_image"):
                add({"kind": "paragraph", "p": p["index"]}, _label_from(pend["text"]), "",
                    "不确定", "中", "标签段后的空段落")
            pend = None

    # 用户点名栏目（关键词匹配）
    hit = []
    for kw in re.split(r"[，,、;；\n]", requirement or ""):
        kw = kw.strip()
        if len(kw) >= 2 and any(kw in norm(f["label"]) for f in fields):
            hit.append(kw)

    for tag, a in anchor_tags.items():
        add({"kind": "image", "part": a["part"], "relationship": a["relationship"],
             "media": a["media"]}, tag, "{{" + tag + "}}", "原样", "高",
            f"图片锚点 {a['fmt']}，原位替换保留尺寸", "")

    uncertain = []
    for t in recon["tables"]:
        for cell in t["cells"]:
            if cell["blank"] and (t["index"], cell["row"], cell["col"]) not in marked_anchors \
                    and not _left_label(t, cell, t["index"]) and not _up_label(t, cell, t["index"]):
                uncertain.append({"target": {"kind": "cell", "table": t["index"],
                                             "row": cell["row"], "col": cell["col"]},
                                  "reason": "空白但无可识别标签，未授权"})
    return {"file": recon["file"], "fields": fields, "uncertain_blanks": uncertain,
            "user_keywords_matched": hit,
            "note": "空白本身不是授权；uncertain_blanks 需用户明确指向才可进白名单"}


def _in_tokens(ti, cell, recon):
    return any(m["target"].get("kind") == "cell" and m["target"].get("table") == ti
               and m["target"].get("row") == cell["row"] and m["target"].get("col") == cell["col"]
               for m in recon["markers"])


def find_cell(recon, ti, row, col):
    """把 row.cells 的网格坐标归一到去重后的锚点单元格条目。"""
    t = recon["tables"][ti]
    for c in t["cells"]:
        if c["row"] == row and c["col"] == col:
            return c
    for mr in t["merged_regions"]:
        if [row, col] in mr["occurrences"]:
            ar, ac = mr["anchor"]
            for c in t["cells"]:
                if c["row"] == ar and c["col"] == ac:
                    return c
    return None


def _cap_note(recon, ti, cell):
    t = recon["tables"][ti]
    anchor = next((c for c in t["cells"] if c["row"] == cell["row"] and c["col"] == cell["col"]), None)
    note = ""
    if anchor:
        if anchor.get("text") and (mm := LIMIT_RE.search(anchor["text"])):
            note += f"限{mm.group(1)}字 "
        note += f"可用段落数≈{anchor['paras']}"
    return note


def _merge_note(recon, ti, cell):
    t = recon["tables"][ti]
    for mr in t["merged_regions"]:
        if [cell["row"], cell["col"]] in mr["occurrences"]:
            other = [o for o in mr["occurrences"] if o != [cell["row"], cell["col"]]]
            return (f"合并区域 anchor={mr['anchor']} 共享格={other} "
                    "写入前需 id(cell._tc) 去重")
    return ""


def _row_label(t, cell):
    left = [c for c in t["cells"] if c["row"] == cell["row"] and c["col"] < cell["col"]
            and not c["blank"] and len(norm(c["text"])) <= 12]
    if left:
        best = max(left, key=lambda x: x["col"])
        return best["text"]
    return ""


def _left_label(t, cell, ti):
    for c in sorted(t["cells"], key=lambda x: (x["row"], x["col"])):
        if c["row"] == cell["row"] and c["col"] < cell["col"] and c["ends_with_colon"]:
            return c["text"]
    return ""


def _up_label(t, cell, ti):
    for c in t["cells"]:
        if c["col"] == cell["col"] and c["row"] < cell["row"] and c["ends_with_colon"]:
            last = max(x["row"] for x in t["cells"] if x["col"] == cell["col"] and x["row"] < cell["row"])
            if cell["row"] - last == 1:
                return c["text"]
    return ""


# ---------------------------------------------------------------- phase 2b

def phase_diff(docx_a, docx_b):
    """双样张差异比对：两份同结构、不同人填写的成品对比，变了的段落=变量候选。"""
    import difflib
    from docx import Document as _Doc
    da, db = _Doc(docx_a), _Doc(docx_b)
    ta = [para_text(p) for p in da.paragraphs]
    tb = [para_text(p) for p in db.paragraphs]
    fields = []
    sm = difflib.SequenceMatcher(None, [norm(x) for x in ta], [norm(x) for x in tb])
    changed = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        for i in range(i1, min(i2, len(ta))):
            changed.append(i)
    used_rows = set()
    used_ids = set()
    for i in changed:
        if i in used_rows or i >= len(ta) or i >= len(tb):
            continue
        used_rows.add(i)
        base, alt = ta[i].strip(), tb[i].strip()
        if not base and not alt:
            continue
        prefix = ""
        if norm(base) and norm(alt) and norm(base) != norm(alt):
            cp = 0
            for x, y in zip(base, alt):
                if x != y:
                    break
                cp += 1
            prefix = base[:cp].rstrip()
        label = re.sub(r"[：:].*$", "", prefix).strip()
        if len(label) < 2:
            label = (base or alt)[:12]
        label = re.sub(r"[^\w\u4e00-\u9fff]+", "_", label)[:20].strip("_") or f"diff_p{i}"
        fid, n = label, 2
        while fid in used_ids:
            fid = f"{label}_{n}"
            n += 1
        used_ids.add(fid)
        fields.append({"field_id": fid, "label": label or fid,
                       "location": {"kind": "paragraph", "p": i},
                       "original_content": base[:60], "alt_content": alt[:60],
                       "authorization_type": "不确定",
                       "capacity_note": "两份样例此处文字不同，推断为填写位",
                       "merge_flag": "", "cross_run_flag": False,
                       "confidence": "中", "source": "样例差异"})
    return {"file_a": str(docx_a), "file_b": str(docx_b),
            "paragraphs_a": len(ta), "paragraphs_b": len(tb),
            "structure_same": len(ta) == len(tb), "fields": fields}


# ---------------------------------------------------------------- phase 4

def _set_run_text(run, text):
    run.text = text  # run 级赋值，保留 rPr


def _borrow_rpr(para, anchor_run):
    """空白位插入文本时借格式：优先相邻有文字 run 的 rPr，深拷贝。"""
    src = None
    if anchor_run is not None:
        src = anchor_run._element.find(qn("w:rPr"))
    if src is None:
        for r in para.runs:
            if r.text.strip():
                src = r._element.find(qn("w:rPr"))
                break
    return deepcopy(src) if src is not None else None


def _write_value_para(para, value, donor_run=None):
    """把一个（可能多行的）值写入段落：优先复用空 run，否则新建 run 并借 rPr。

    复用判定很严格：只有除 <w:rPr> 外没有任何子节点的 run 才算“空”。
    含图片(drawing/pict)或其他内容的 run 即使没有文字也不能清空，否则会把内联图片抹掉。"""
    lines = str(value).split("\n")

    def truly_empty(run):
        if run.text:
            return False
        return all(c.tag == qn("w:rPr") for c in run._element)

    target_run = next((r for r in para.runs if truly_empty(r)), None)
    if target_run is None:
        target_run = para.add_run("")
        rpr = _borrow_rpr(para, donor_run or para.runs[-1] if para.runs else donor_run)
        if rpr is not None:
            target_run._element.insert(0, rpr)
    _set_run_text(target_run, lines[0])
    for ln in lines[1:]:
        target_run.add_break()
        new_r = para.add_run(ln)
        rpr = target_run._element.find(qn("w:rPr"))
        if rpr is not None:
            new_r._element.insert(0, deepcopy(rpr))
        target_run = new_r


def _replace_tokens_in_para(p, mapping):
    """token 替换：单 run 直接改文本；跨 run 重建 runs 并深拷贝 rPr。同段多字段从后往前。"""
    writes = []
    ms = list(TOKEN.finditer(para_text(p)))
    for m in reversed(ms):
        key = m.group(1).strip()
        if key not in mapping:
            continue
        runs = p.runs
        spans, off = [], 0
        for i, r in enumerate(runs):
            s, e = off, off + len(r.text)
            if e > m.start() and s < m.end():
                spans.append((i, s, e))
            off = e
        if not spans:
            continue
        value = str(mapping[key])
        if len(spans) == 1:
            i, s, e = spans[0]
            t = runs[i].text
            _set_run_text(runs[i], t[: m.start() - s] + value + t[m.end() - s:])
        else:
            i0, s0, e0 = spans[0]
            i1, s1, e1 = spans[-1]
            donor = deepcopy(runs[i0]._element.find(qn("w:rPr"))) if \
                runs[i0]._element.find(qn("w:rPr")) is not None else None
            new_run = p.add_run("")
            if donor is not None:
                new_run._element.insert(0, deepcopy(donor))
            _set_run_text(new_run, value)
            runs[i0]._element.addprevious(new_run._element)
            runs[i0].text = runs[i0].text[: m.start() - s0]
            for i in range(i0 + 1, i1):
                _set_run_text(runs[i], "")
            runs[i1].text = runs[i1].text[m.end() - s1:]
        writes.append({"key": key, "para": p})
    return writes


def _find_para(doc, target):
    if target["kind"] == "paragraph":
        return doc.paragraphs[target["p"]]
    if target["kind"] == "cell":
        cell = doc.tables[target["table"]].rows[target["row"]].cells[target["col"]]
        return cell.paragraphs[target.get("p", 0)], cell
    raise ValueError("未知 target 类型: %r" % target)


def phase_fill(docx_path, whitelist_path, values_path, mode, out_path, log_path):
    doc = Document(docx_path)
    white = json.loads(Path(whitelist_path).read_text(encoding="utf-8"))
    values = json.loads(Path(values_path).read_text(encoding="utf-8")) if values_path else {}
    items = white["fields"] if isinstance(white, dict) else white
    tc_seen = set()
    log = []
    for item in items:
        fid = item["field_id"]
        target = item.get("target", item.get("location"))
        value = values.get(fid, item.get("value", ""))
        if mode == "mark":
            value = "{{ " + fid + " }}"
        para = None
        container_tc = None
        if target["kind"] == "image":
            log.append({"field_id": fid, "target": target, "mode": mode,
                        "method": "图片锚点，无需文字写入", "before": "", "after": "",
                        "written": value})
            continue
        if target["kind"] == "cell":
            cell = doc.tables[target["table"]].rows[target["row"]].cells[target["col"]]
            container_tc = cell._tc
            para = cell.paragraphs[target.get("p", 0)]
        else:
            para = doc.paragraphs[target["p"]]
        before = para_text(para)
        if mode == "mark" and item.get("replace_para"):
            for r in list(para.runs):
                if any(c.tag in (qn("w:drawing"), qn("w:pict"), qn("w:object"))
                       for c in r._element):
                    continue
                r.text = ""
            for hl in list(para._element.iterchildren(qn("w:hyperlink"))):
                para._element.remove(hl)
            method_extra = "整段替换(样例差异授权)"
        else:
            method_extra = ""
        ms = list(TOKEN.finditer(before))
        has_own_token = mode == "fill" and any(m.group(1).strip() == fid for m in ms)
        slot = (container_tc, target.get("p", 0)) if container_tc is not None else None
        if not has_own_token and slot is not None and slot in tc_seen:
            raise ValueError(f"{fid}: 合并单元格追加写入重复落点 tc={id(slot[0])} p={slot[1]}，已阻断")
        if slot is not None:
            tc_seen.add(slot)
        if has_own_token:
            _replace_tokens_in_para(para, {fid: value})
            method = "run级替换(含跨run重建)"
        elif mode == "fill":
            _write_value_para(para, value)
            method = "空白run写入/新建run借rPr"
        elif any(m.group(1).strip() == fid for m in ms):
            method = "已存在同名标记，跳过"
        else:
            _write_value_para(para, value)
            method = "插入标记token"
        log.append({"field_id": fid, "target": target, "mode": mode,
                    "method": method + method_extra,
                    "before": before, "after": para_text(para), "written": value})
    _align_label_value_cells(doc)
    cleanup.clean(doc)
    ooxmlcheck.normalize(doc)
    doc.save(out_path)
    Path(log_path).write_text(json.dumps({"source": str(docx_path), "output": str(out_path),
                                          "writes": log}, ensure_ascii=False, indent=2), encoding="utf-8")
    return log


def _align_label_value_cells(doc):
    """标签为底端对齐、值为顶端对齐时，把值也改到底端，避免填入值离下划线太远。"""
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            label = cells[0]
            if not label.text.strip().endswith(("：", ":")):
                continue
            if label.vertical_alignment != WD_CELL_VERTICAL_ALIGNMENT.BOTTOM:
                continue
            seen = set()
            for cell in cells[1:]:
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                if cell.vertical_alignment == WD_CELL_VERTICAL_ALIGNMENT.TOP:
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM


# ---------------------------------------------------------------- phase 5

def _c14n_equal(data_a, data_b):
    try:
        from lxml import etree
        ca = etree.tostring(etree.fromstring(data_a), method="c14n2")
        cb = etree.tostring(etree.fromstring(data_b), method="c14n2")
        return ca == cb
    except Exception:
        return False


def _rels_equal(data_a, data_b):
    """比较 .rels 的关系语义：忽略空白、声明和元素顺序，只比 Id/Type/Target 等属性。"""
    try:
        ra, rb = ET.fromstring(data_a), ET.fromstring(data_b)
    except ET.ParseError:
        return False

    def signature(root):
        items = []
        for el in root.iter():
            if el.tag.rsplit('}', 1)[-1] == 'Relationship':
                items.append((el.tag, tuple(sorted(el.attrib.items()))))
        return sorted(items)

    return signature(ra) == signature(rb)


def phase_verify(before, after, log_path, out_path, allowed=("word/document.xml",)):
    report = {"passed": [], "pending": [], "blocked": []}
    log = json.loads(Path(log_path).read_text(encoding="utf-8"))["writes"]
    doc = Document(after)

    def read_target(t):
        if t["kind"] == "image":
            import engine as _e
            return ' '.join(a["tag"] for a in _e.scan_image_anchors(after))
        if t["kind"] == "cell":
            cell = doc.tables[t["table"]].rows[t["row"]].cells[t["col"]]
            return cell.paragraphs[t.get("p", 0)].text
        return doc.paragraphs[t["p"]].text

    # L1 内容核对
    for w in log:
        if w["target"].get("kind") == "image":
            if w["field_id"] in read_target(w["target"]).split():
                report["passed"].append(f"图片锚点 {w['field_id']} 保持原位")
            else:
                report["blocked"].append(f"图片锚点 {w['field_id']} 丢失或变更")
            continue
        actual = read_target(w["target"])
        src = norm(w["written"])
        if src and src in norm(actual):
            report["passed"].append(f"内容 {w['field_id']} 落位一致")
        elif not src:
            report["passed"].append(f"内容 {w['field_id']} 空值跳过")
        else:
            report["blocked"].append(f"内容 {w['field_id']} 未落位: 期望≈{w['written'][:20]} 实际={actual[:40]}")
    vals = [w["written"] for w in log if w["written"].strip() and w["mode"] == "fill"]
    for w in log:
        if w["mode"] != "fill":
            continue
        actual = norm(read_target(w["target"]))
        for v in vals:
            if v != w["written"] and norm(v) and norm(v) in actual and len(norm(v)) > 6:
                report["blocked"].append(f"疑似串字段: {w['field_id']} 位置出现 {v[:16]}")
    # L1.5 内容指纹：除授权写入外，文字总量应仅按预期变化
    import engine as _engine
    fb, fa = _engine.content_fingerprint(before), _engine.content_fingerprint(after)
    if fb == fa:
        report["passed"].append("内容指纹一致：除位置外文字未变")
    else:
        extra = set()
        for w in log:
            if w["mode"] == "fill" and w["written"].strip():
                extra.add(norm(w["written"]))
        report["pending"].append(
            f"内容指纹不同（填充模式下属预期；mark 模式应一致）。before={fb[:12]} after={fa[:12]}")
    # L2 结构白名单
    zb, za = zipfile.ZipFile(before), zipfile.ZipFile(after)
    DIRS_ONLY = lambda s: {n for n in s if not n.endswith('/')}
    nb, na = DIRS_ONLY(set(zb.namelist())), DIRS_ONLY(set(za.namelist()))
    if nb != na:
        report["blocked"].append(f"包成员变化: 新增{na - nb} 删除{nb - na}")
    else:
        report["passed"].append("包成员列表一致")
    real_changed, reordered = [], []
    for n in sorted(nb & na):
        ba, bb = zb.read(n), za.read(n)
        if ba == bb:
            continue
        if n.endswith(".rels") and _rels_equal(ba, bb):
            reordered.append(n)
        elif n.endswith(".xml") and _c14n_equal(ba, bb):
            reordered.append(n)
        else:
            real_changed.append(n)
    if reordered:
        report["pending"].append(f"序列化重排(语义一致，非结构变更): {reordered}")
    out_of_scope = [n for n in real_changed if n not in allowed
                    and not n.startswith("_rels") and n != "[Content_Types].xml"]
    for n in out_of_scope:
        if any(k in n for k in ("header", "footer", "media/", "settings", "styles")):
            report["blocked"].append(f"范围外变更(阻断): {n}")
        else:
            report["pending"].append(f"未预期变更部件: {n}")
    if not out_of_scope:
        report["passed"].append("页眉/页脚/媒体/样式/设置 无真实变更")
    swapped = {w["target"].get("media") for w in log
               if w["target"].get("kind") == "image"} - {None}
    for name in sorted(nb & na):
        if name.startswith("word/media/") and zb.read(name) != za.read(name):
            if name in swapped:
                report["passed"].append(f"授权换图: {name}")
            else:
                report["blocked"].append(f"二进制媒体部件被改动: {name}")
    tb = Document(before).tables
    ta = doc.tables
    if len(tb) != len(ta):
        report["blocked"].append("表格数量变化")
    else:
        for ti, (b, a) in enumerate(zip(tb, ta)):
            sb = [(c.row, c.col, c.tc) for c in recon_cells(b)]
            sa = [(c.row, c.col, c.tc) for c in recon_cells(a)]
            sigb = [(x[0], x[1], grid_span_of(x[2]), v_merge_of(x[2])) for x in sb]
            siga = [(x[0], x[1], grid_span_of(x[2]), v_merge_of(x[2])) for x in sa]
            if sigb != siga:
                report["blocked"].append(f"表{ti} gridSpan/vMerge 与基线不一致")
            if len(b.rows) != len(a.rows):
                report["blocked"].append(f"表{ti} 行数变化")
        report["passed"].append("表格结构与基线一致")
    # L3 可视
    report["pending"].append("可视排版尚未验证（DOCX→PDF 需 Word 渲染，本机未自动执行；页数/截字/跨页请人工抽查）")
    result = {"passed_count": len(report["passed"]), "pending_count": len(report["pending"]),
              "blocked_count": len(report["blocked"]), **report}
    Path(out_path).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def recon_cells(t):
    out, seen = [], set()
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            if cell._tc in seen:
                continue
            seen.add(cell._tc)
            out.append(type("X", (), {"row": ri, "col": ci, "tc": cell._tc})())
    return out


# ---------------------------------------------------------------- cli

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("recon")
    r.add_argument("--docx", required=True)
    r.add_argument("--out", default="recon.json")
    c = sub.add_parser("candidates")
    c.add_argument("--recon", required=True)
    c.add_argument("--require", default="")
    c.add_argument("--out", default="candidates.json")
    d = sub.add_parser("diff")
    d.add_argument("--a", required=True)
    d.add_argument("--b", required=True)
    d.add_argument("--out", default="diff.json")
    f = sub.add_parser("fill")
    f.add_argument("--docx", required=True)
    f.add_argument("--whitelist", required=True)
    f.add_argument("--values")
    f.add_argument("--mode", choices=["fill", "mark"], default="fill")
    f.add_argument("--out", required=True)
    f.add_argument("--log", default="fill_log.json")
    v = sub.add_parser("verify")
    v.add_argument("--before", required=True)
    v.add_argument("--after", required=True)
    v.add_argument("--log", required=True)
    v.add_argument("--out", default="verify.json")
    v.add_argument("--allow", default="")
    a = ap.parse_args()
    if a.cmd == "recon":
        data = phase_recon(a.docx)
        Path(a.out).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"recon -> {a.out} | 段落{data['paragraph_count']} 表{data['table_count']} "
              f"标记{len(data['markers'])} 合并区{sum(len(t['merged_regions']) for t in data['tables'])}")
    elif a.cmd == "candidates":
        recon = json.loads(Path(a.recon).read_text(encoding="utf-8"))
        data = phase_candidates(recon, a.require)
        Path(a.out).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        conf = {}
        for f_ in data["fields"]:
            conf[f_["confidence"]] = conf.get(f_["confidence"], 0) + 1
        print(f"candidates -> {a.out} | 字段{len(data['fields'])} {conf} "
              f"不确定空白{len(data['uncertain_blanks'])}")
    elif a.cmd == "diff":
        data = phase_diff(a.a, a.b)
        Path(a.out).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"diff -> {a.out} | 段落A{data['paragraphs_a']} 段落B{data['paragraphs_b']} "
              f"结构一致={data['structure_same']} 差异字段{len(data['fields'])}")
    elif a.cmd == "fill":
        log = phase_fill(a.docx, a.whitelist, a.values, a.mode, a.out, a.log)
        print(f"fill {a.mode} -> {a.out} | 写入{len(log)} 字段，日志 {a.log}")
    elif a.cmd == "verify":
        allowed = tuple(x.strip() for x in a.allow.split(",") if x.strip()) or ("word/document.xml",)
        res = phase_verify(a.before, a.after, a.log, a.out, allowed)
        print(f"verify -> {a.out} | 通过{res['passed_count']} 待办{res['pending_count']} 阻断{res['blocked_count']}")
        for line in res["blocked"]:
            print("  BLOCK:", line)
        sys.exit(1 if res["blocked_count"] else 0)


if __name__ == "__main__":
    main()
