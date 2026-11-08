# -*- coding: utf-8 -*-
"""联网文献检索：e2sci 主入口、平台跳转、OpenAlex/Crossref 题录与 GB/T 7714 格式化。

只使用标准库，不依赖第三方网络包；不下载全文、不绕过登录。
"""
from __future__ import annotations

import json
import re
import urllib.parse
import urllib.request

UA = "WordTemplateWorkbench/5.13 (local desktop; metadata only)"

# 用户指定的 e2sci 推广/登录参数，原样保留。
E2SCI_PARTNER_UUID = "ebdc5ff8ffaa41feb414d9a283d32da9"
E2SCI_URL = ("https://www.e2sci.com/?partner_uuid=" + E2SCI_PARTNER_UUID +
             "&local_login=1")

PLATFORMS = [
    {"id": "e2sci", "name": "e2sci 元空", "url": E2SCI_URL, "query": False,
     "note": "打开后把检索词粘贴到输入框，点 Search"},
    {"id": "cnki", "name": "知网 CNKI",
     "url": "https://kns.cnki.net/kns8s/defaultresult/index?dbcode=CFLS&kw={q}", "query": True},
    {"id": "wanfang", "name": "万方",
     "url": "https://s.wanfangdata.com.cn/paper?q={q}", "query": True},
    {"id": "cqvip", "name": "维普",
     "url": "http://qikan.cqvip.com/Qikan/Search/Index?key={q}", "query": True},
    {"id": "google", "name": "Google Scholar",
     "url": "https://scholar.google.com/scholar?q={q}", "query": True},
    {"id": "baidu", "name": "百度学术",
     "url": "https://xueshu.baidu.com/s?wd={q}", "query": True},
    {"id": "pubmed", "name": "PubMed",
     "url": "https://pubmed.ncbi.nlm.nih.gov/?term={q}", "query": True},
    {"id": "arxiv", "name": "arXiv",
     "url": "https://arxiv.org/search/?searchtype=all&query={q}", "query": True},
    {"id": "semantic", "name": "Semantic Scholar",
     "url": "https://www.semanticscholar.org/search?q={q}", "query": True},
    {"id": "openalex", "name": "OpenAlex",
     "url": "https://openalex.org/works?search={q}", "query": True},
]

STOP = {"基于", "面向", "研究", "设计", "实现", "系统", "方法", "技术", "分析", "应用",
        "论文", "毕业", "高校", "一种", "探讨", "浅析", "初探", "及其", "中的", "下的",
        "问题", "对策", "现状", "发展", "影响", "策略", "以", "为", "的", "与", "和", "及",
        "工具", "平台", "辅助", "管理", "构建", "探究", "实践", "思考", "路径", "模式"}

LEX = [
    "生成式人工智能", "自然语言处理", "大语言模型", "语言模型", "人工智能", "机器学习",
    "深度学习", "毕业论文", "毕业设计", "论文写作", "学术写作", "写作辅助", "写作教学",
    "写作平台", "辅助工具", "写作工具", "文档排版", "自动排版", "模板识别", "格式规范",
    "参考文献", "文献管理", "文献检索", "文本生成", "提示词", "智能体", "知识图谱",
    "推荐系统", "教育技术", "需求分析", "系统设计", "系统实现", "数据库设计", "数据可视化",
    "管理系统", "信息管理", "前端开发", "后端开发", "英语写作", "英语教学", "课程思政",
    "跨境电商", "数字经济", "乡村振兴", "养老服务", "供应链", "市场营销", "人力资源",
    "财务管理", "心理健康", "文化传播", "小程序", "大模型", "AIGC", "ChatGPT", "LLM",
    "Web", "Python", "Java", "Vue", "SpringBoot", "MySQL",
]


def clean_query(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())[:240]


def simplify(text: str, limit: int = 6) -> str:
    """把整句题目压成可检索的实词；不引入分词依赖。"""
    s = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]+", " ", text or "")
    out = []
    i = 0
    while i < len(s):
        hit = None
        for n in range(min(10, len(s) - i), 1, -1):
            w = s[i:i + n]
            if w in LEX:
                hit = w
                break
        if hit:
            if hit not in out:
                out.append(hit)
            i += len(hit)
            continue
        j = i + 1
        while j < len(s):
            if any(s[j:j + n] in LEX for n in range(2, min(10, len(s) - j) + 1)):
                break
            if s[j] == " ":
                break
            j += 1
        chunk = s[i:j].strip()
        for part in re.split(r"基于|面向|的|与|和|及|研究|设计|实现|系统|工具|方法|技术|分析|应用|一种|及其|中的|以|为|对|在", chunk):
            part = part.strip()
            if 2 <= len(part) <= 12 and part not in STOP and part not in out:
                out.append(part)
        i = max(j, i + 1)
    out.sort(key=len, reverse=True)
    return " ".join(out[:limit]) if out else clean_query(text)


def platform_url(pid: str, query: str = "") -> str:
    for item in PLATFORMS:
        if item["id"] == pid:
            if item.get("query"):
                return item["url"].format(q=urllib.parse.quote(clean_query(query) or "毕业论文"))
            return item["url"]
    raise KeyError(pid)


def check_online(timeout: int = 8) -> tuple:
    """检测是否能访问 e2sci；返回 (ok, message)。"""
    try:
        req = urllib.request.Request(E2SCI_URL, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status < 500:
                return True, "联网正常 · e2sci 可访问"
            return False, f"e2sci 返回 HTTP {resp.status}"
    except Exception as exc:  # noqa: BLE001
        return False, "无法联网：" + str(exc)


def _get_json(url: str, timeout: int = 18) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def search_openalex(query: str, rows: int = 20, year: int | None = None) -> list:
    url = ("https://api.openalex.org/works?search=" +
           urllib.parse.quote(clean_query(query)) +
           f"&per-page={max(1, min(50, rows))}&sort=relevance_score:desc")
    if year:
        url += f"&filter=from_publication_date:{int(year)}-01-01"
    data = _get_json(url)
    out = []
    for w in data.get("results", []):
        authors = [a.get("author", {}).get("display_name", "")
                   for a in w.get("authorships", []) if a.get("author")]
        loc = w.get("primary_location") or {}
        src = (loc.get("source") or {}) if isinstance(loc, dict) else {}
        b = w.get("biblio") or {}
        doi = (w.get("doi") or "").replace("https://doi.org/", "")
        out.append({
            "title": w.get("display_name") or "",
            "authors": "; ".join([a for a in authors if a][:8]),
            "year": str(w.get("publication_year") or ""),
            "journal": src.get("display_name") or "",
            "volume": b.get("volume") or "", "issue": b.get("issue") or "",
            "pages": "-".join(x for x in (b.get("first_page"), b.get("last_page")) if x),
            "doi": doi,
            "url": loc.get("landing_page_url") or (f"https://doi.org/{doi}" if doi else ""),
            "type": "article", "source": "OpenAlex",
        })
    return out


def search_crossref(query: str, rows: int = 20, year: int | None = None) -> list:
    url = ("https://api.crossref.org/works?query=" +
           urllib.parse.quote(clean_query(query)) +
           f"&rows={max(1, min(50, rows))}&select=DOI,title,author,container-title,"
           "issued,volume,issue,page,type,URL")
    if year:
        url += f"&filter=from-pub-date:{int(year)}"
    data = _get_json(url)
    out = []
    for w in (data.get("message") or {}).get("items", []):
        authors = []
        for a in w.get("author", []) or []:
            fam, giv = a.get("family", ""), a.get("given", "")
            if fam and giv and re.search(r"[\u4e00-\u9fff]", fam + giv):
                authors.append(fam + giv)
            elif fam:
                authors.append(f"{fam}, {giv}".strip(", "))
            elif a.get("name"):
                authors.append(a["name"])
        issued = (w.get("issued") or {}).get("date-parts") or [[None]]
        out.append({
            "title": (w.get("title") or [""])[0],
            "authors": "; ".join(authors[:8]),
            "year": str((issued[0][0] if issued and issued[0] else "") or ""),
            "journal": (w.get("container-title") or [""])[0],
            "volume": w.get("volume") or "", "issue": w.get("issue") or "",
            "pages": (w.get("page") or "").replace("–", "-"),
            "doi": w.get("DOI") or "", "url": w.get("URL") or "",
            "type": "conference" if "proceedings" in str(w.get("type")) else "article",
            "source": "Crossref",
        })
    return out


def search(source: str, query: str, rows: int = 20, year: int | None = None) -> list:
    src = (source or "openalex").lower()
    if src == "openalex":
        try:
            items = search_openalex(query, rows, year)
        except Exception as exc:  # noqa: BLE001
            if "429" not in str(exc) and "Too Many" not in str(exc):
                raise
            items = search_crossref(query, rows, year)
            for item in items:
                item["source"] = "Crossref（OpenAlex 限流回退）"
    else:
        items = search_crossref(query, rows, year)
    seen, out = set(), []
    for it in items:
        key = (re.sub(r"[^\w\u4e00-\u9fff]", "", (it.get("title") or "").lower()), it.get("year"))
        if key in seen or not it.get("title"):
            continue
        seen.add(key)
        out.append(it)
    return out


def _split_authors(authors: str) -> list:
    text = str(authors or "")
    if ";" in text or "；" in text:
        raw = re.split(r"\s*[;；]\s*", text)
    elif re.search(r"\s+and\s+", text):
        raw = re.split(r"\s+and\s+", text)
    else:
        raw = re.split(r"\s*[,，、]\s*", text)
    return [a.strip() for a in raw if a.strip()]


def _en_author(name: str) -> str:
    name = name.replace(".", " ").strip()
    if "," in name:
        sur, given = name.split(",", 1)
        parts = [sur.strip()] + given.split()
    else:
        parts = name.split()
    if not parts:
        return ""
    return (parts[0].upper() + (" " + " ".join(p[0].upper() for p in parts[1:]) if len(parts) > 1 else "")).strip()


def format_gbt(ref: dict, number: int | None = None) -> str:
    title = (ref.get("title") or "").strip().rstrip(".")
    journal = (ref.get("journal") or "").strip()
    year = (ref.get("year") or "").strip()
    volume = (ref.get("volume") or "").strip()
    issue = (ref.get("issue") or "").strip()
    pages = (ref.get("pages") or "").strip().replace("–", "-")
    doi = (ref.get("doi") or "").strip()
    names = _split_authors(ref.get("authors"))
    if names:
        chinese = bool(re.search(r"[\u4e00-\u9fff]", "".join(names)))
        shown = names[:3]
        if chinese:
            authors = ", ".join(shown) + (", 等" if len(names) > 3 else "")
        else:
            authors = ", ".join(_en_author(n) for n in shown) + (", et al" if len(names) > 3 else "")
    else:
        authors = ""
    kind = (ref.get("type") or "article").lower()
    mark = "C" if kind in ("conference", "proceedings") else "J"
    tail = journal
    if year:
        tail = f"{tail}, {year}" if tail else year
    if volume:
        tail = (f"{tail}, {volume}" if tail else volume) + (f"({issue})" if issue else "")
    elif issue:
        tail = f"{tail}, ({issue})" if tail else f"({issue})"
    if pages:
        tail = f"{tail}: {pages}" if tail else pages
    if tail:
        text = (f"{authors}. " if authors else "") + f"{title}[{mark}]. {tail}."
    else:
        text = (f"{authors}. " if authors else "") + f"{title}[{mark}]."
    if doi:
        text += f" DOI: {doi}."
    return f"[{number}] {text}" if number else text


def format_gbt_list(refs: list) -> str:
    return "\n".join(format_gbt(r, i) for i, r in enumerate(refs, 1))
