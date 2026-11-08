"""AI 客户端：本机 Ollama 与任意 OpenAI 兼容 API。仅用标准库，无额外依赖。

配置持久化在数据目录 ai_config.json：
{"backend": "ollama|openai", "base_url": "...", "model": "...", "api_key": "",
 "temperature": 0.2, "timeout": 120, "max_tokens": 2048,
 "vision_backend": "ollama|openai", "vision_base_url": "...",
 "vision_model": "...", "vision_api_key": "", "vision_max_pages": 12}

文本模型用于字段分类和正文生成；视觉模型用于把 Word 页面截图识别成栏目。
默认视觉服务是托管的 Qwen3-VL（开源权重），也可指向本机 Ollama。
隐私：backend=ollama 时请求只发往本机地址；云端服务时文档摘录/页面截图会发往
所配置的远程服务，界面已提示。所有网络调用都由用户点击按钮显式触发。
"""
from __future__ import annotations

import ipaddress
import json
import re
import shutil
import subprocess
import urllib.error
import urllib.request
from urllib.parse import urlsplit
import base64

OLLAMA_DEFAULT = 'http://127.0.0.1:11434'
OPENAI_DEFAULT = 'https://api.deepseek.com'
DOUBAO_DEFAULT = 'https://ark.cn-beijing.volces.com/api/v3'
DOUBAO_MODEL = 'doubao-seed-2-1-pro-260915'
RECOMMENDED_OLLAMA_MODEL = 'qwen3:4b-instruct'
# 思考型模型专门负责“需要推理”的活（标题层级判定），指示版负责产出型任务
THINK_MODEL_DEFAULT = 'qwen3:4b'
# 思考版（qwen3:4b / deepseek-r1）在 Ollama 上会把正文写成独白，正文生成请用非思考版

SILICONFLOW_DEFAULT = 'https://api.siliconflow.cn/v1'
DASHSCOPE_DEFAULT = 'https://dashscope.aliyuncs.com/compatible-mode/v1'
OPENROUTER_DEFAULT = 'https://openrouter.ai/api/v1'
ANTHROPIC_DEFAULT = 'https://api.anthropic.com/v1'
ANTHROPIC_VERSION = '2023-06-01'
QWEN_VL_MODEL = 'Qwen/Qwen3-VL-32B-Instruct'
QWEN_VL_DASHSCOPE_MODEL = 'qwen3-vl-plus'
QWEN_VL_OLLAMA_MODEL = 'qwen3-vl:8b'

VISION_PRESETS = (
    ('硅基流动 Qwen3-VL-32B（开源权重，推荐）', 'siliconflow',
     'openai', SILICONFLOW_DEFAULT, QWEN_VL_MODEL),
    ('阿里云百炼 Qwen3-VL-Plus（托管）', 'dashscope',
     'openai', DASHSCOPE_DEFAULT, QWEN_VL_DASHSCOPE_MODEL),
    ('OpenRouter Qwen3-VL', 'openrouter',
     'openai', OPENROUTER_DEFAULT, 'qwen/qwen3-vl-235b-a22b-instruct'),
    ('本机 Ollama Qwen3-VL（无需联网）', 'ollama',
     'ollama', OLLAMA_DEFAULT, QWEN_VL_OLLAMA_MODEL),
    ('自定义视觉服务', 'custom', 'openai', '', ''),
)


class AIError(RuntimeError):
    pass


# 不少中转站挂在 Cloudflare 后面，按客户端指纹拦 Python-urllib（error code: 1010）。
USER_AGENT = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
              '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36')
BLOCK_MARKS = ('error code: 1010', 'error code: 1020', 'cloudflare', 'just a moment',
               'cf-browser-verification', 'request blocked')


def is_blocked(code, body):
    text = (body or '').lower()
    return int(code) in (403, 451, 503) and any(mark in text for mark in BLOCK_MARKS)


EMBEDDED_BACKEND = 'embedded'
EMBEDDED_LABEL = '内置小模型 Qwen3-0.6B（离线，质量有限）'
ANTHROPIC_BACKEND = 'anthropic'


def _user_env(name):
    """读用户级环境变量。刚改过的值当前进程读不到，注册表里一定有。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment') as key:
            value, _ = winreg.QueryValueEx(key, name)
            return str(value or '')
    except Exception:
        return ''


def _root_of(base_url):
    from urllib.parse import urlsplit
    url = (base_url or OLLAMA_DEFAULT).strip().rstrip('/')
    if '://' not in url:
        url = 'http://' + url
    parts = urlsplit(url)
    return f'{parts.scheme}://{parts.netloc}'


def ollama_exe():
    """按 PATH、用户 PATH、OLLAMA_HOME、默认安装位置找 ollama.exe。"""
    import os
    import shutil
    from pathlib import Path
    found = shutil.which('ollama')
    if found:
        return found
    dirs = []
    for source in (_user_env('Path'), os.environ.get('PATH', '')):
        dirs.extend(part for part in source.split(';') if part.strip())
    home = _user_env('OLLAMA_HOME') or os.environ.get('OLLAMA_HOME', '')
    if home:
        dirs.extend([str(Path(home) / 'app'), home])
    local = os.environ.get('LOCALAPPDATA') or ''
    if local:
        dirs.append(str(Path(local) / 'Programs' / 'Ollama'))
    for item in dirs:
        try:
            exe = Path(item.strip('"').strip()) / 'ollama.exe'
            if exe.is_file():
                return str(exe)
        except OSError:
            continue
    return ''


def ollama_running(base_url=''):
    root = _root_of(base_url)
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    for path in ('/api/tags', '/v1/models'):
        try:
            with opener.open(root + path, timeout=2.0) as resp:
                if resp.status == 200:
                    return True
        except Exception:
            continue
    return False


def points_to_local_ollama(cfg):
    """配置里是否把（文本或视觉）指向本机的 Ollama 服务。"""
    cfg = cfg or {}
    for key, base_key in (('backend', 'base_url'), ('vision_backend', 'vision_base_url')):
        if cfg.get(key) != 'ollama':
            continue
        root = _root_of(cfg.get(base_key) or OLLAMA_DEFAULT)
        host = urlsplit(root).hostname or ''
        if host in ('localhost', '127.0.0.1', '::1') or (host and not _public(host)):
            return True
    return False


def _public(host):
    try:
        return not ipaddress.ip_address(host).is_private
    except ValueError:
        return True


def start_ollama(wait=45.0):
    """本机 Ollama 没在跑就把它拉起来；模型目录用注册表里的用户级设置。"""
    import os
    import subprocess
    import time
    if ollama_running():
        return True, '本机 Ollama 已在运行'
    exe = ollama_exe()
    if not exe:
        return False, ('没找到 ollama.exe：请安装 Ollama，或在 AI 设置里改用'
                       '「内置小模型」/ 云端服务')
    env = dict(os.environ)
    for name in ('OLLAMA_MODELS', 'OLLAMA_HOME'):
        value = _user_env(name)
        if value:
            env[name] = value
    flags = (getattr(subprocess, 'DETACHED_PROCESS', 0)
             | getattr(subprocess, 'CREATE_NEW_PROCESS_GROUP', 0))
    try:
        subprocess.Popen([exe, 'serve'], env=env, cwd=os.path.dirname(exe),
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                         creationflags=flags)
    except OSError as exc:
        return False, f'启动 Ollama 失败：{exc}'
    deadline = time.time() + wait
    while time.time() < deadline:
        if ollama_running():
            return True, f'已自动启动本机 Ollama（{exe}）'
        time.sleep(1.0)
    return False, f'Ollama 进程已启动但 {wait:.0f} 秒内没就绪，可手动运行：ollama serve'


STRUCTURE_SYSTEM = (
    '你是 Word 长文档的结构分析助手。用户会给你一份文档的段落清单，每行带编号、'
    '样式名、字号、是否加粗和对齐方式。你的任务是判断哪些段落是标题、是几级标题。'
    '判断依据（按优先级）：① Word 样式名（Heading/标题 1-3、styleId 为 1-3）；'
    '② 编号规律（“第一章”“一、”“1.1”“1.1.1”“（一）”等）；③ 版式线索（字号比正文大、加粗、'
    '居中、段前段后留白）。正文段落、表格内容、图注、页眉页脚一律不算标题。'
    '级别映射：章/一级序号→1，节/二级序号→2，小节/三级序号→3；只标注明确可判断的，'
    '拿不准的宁可不标。只输出 JSON：{"headings":[{"i":段落编号,"level":1}]}，不要输出解释。'
)


def structure_prompt(items, title=''):
    """items: [{'i','text','style','size','bold','align'}]，按文档顺序。"""
    lines = [f'文档标题：{title or "（未提供）"}', '段落清单：']
    for item in items:
        bits = [f'#{item["i"]}']
        if item.get('style'):
            bits.append('样式=' + str(item['style'])[:20])
        if item.get('size'):
            bits.append(f'{float(item["size"]):g}pt')
        if item.get('bold'):
            bits.append('加粗')
        if item.get('align') in ('center', 'right'):
            bits.append('居中' if item['align'] == 'center' else '右对齐')
        text = ' '.join(str(item.get('text') or '').split())[:60]
        lines.append(' '.join(bits) + '：' + text)
    return chr(10).join(lines)


def think_client(cfg):
    """思考模型客户端：本机 Ollama 时用 think_model（默认 qwen3:4b），否则沿用当前后端。"""
    cfg = dict(cfg or {})
    model = (cfg.get('think_model') or THINK_MODEL_DEFAULT).strip()
    if (cfg.get('backend') or '') == 'ollama' and model:
        think_cfg = dict(cfg)
        think_cfg.update({'model': model, 'timeout': max(600, int(cfg.get('timeout', 120) or 120)),
                          'max_tokens': max(2048, int(cfg.get('max_tokens', 2048) or 2048))})
        return Client(think_cfg)
    return Client(cfg)


HEADING_NUMBER_RULES = (
    (re.compile(r'^第[一二三四五六七八九十百零〇\d]+\s*[章篇]'), 1),
    (re.compile(r'^\d+\.\d+\.\d+'), 3),
    (re.compile(r'^\d+\.\d+'), 2),
    (re.compile(r'^[一二三四五六七八九十]+\s*、'), 2),
    (re.compile(r'^（[一二三四五六七八九十]+）'), 3),
    (re.compile(r'^[(（]\d+[)）]'), 3),
)


def numbered_headings(items):
    """靠序号规律确定性判定标题（标题都很短，正文很少这样开头）。"""
    found = {}
    for item in items:
        text = str(item.get('text') or '').strip()
        if not text or len(text) > 40:
            continue
        for pattern, level in HEADING_NUMBER_RULES:
            if pattern.match(text):
                found[int(item['i'])] = level
                break
    return found


def infer_headings(client, items, title='', max_tokens=6000):
    """让思考模型判定标题层级，返回 {段落编号: 级别}（已做范围与文本校验）。"""
    if not items:
        return {}, ''
    obj = client.ask_json(STRUCTURE_SYSTEM, structure_prompt(items, title), 'headings',
                          temperature=0.0, max_tokens=max_tokens, think=True)
    known = {int(item['i']): item for item in items}
    levels = {}
    for entry in obj.get('headings') or []:
        try:
            index = int(entry.get('i'))
            level = int(entry.get('level'))
        except (TypeError, ValueError):
            continue
        item = known.get(index)
        if item is None or level not in (1, 2, 3):
            continue
        text = str(item.get('text') or '').strip()
        if not text or len(text) > 60:
            continue
        levels[index] = level
    numbered = numbered_headings(items)
    for index, level in numbered.items():
        levels[index] = level
    note = obj.get('note') or ''
    if numbered:
        note = (note + '；' if note else '') + f'序号规则补 {len(numbered)} 个'
    return levels, note


def embedded_assets():
    """内置模型文件是否就位（不启动服务，可在界面线程调用）。"""
    try:
        import embedded_ai
    except ImportError:
        return None
    return embedded_ai.find_assets()


def embedded_endpoint():
    """启动或复用内置 llama-server，返回 (OpenAI 兼容 base_url, 模型名)。

    会阻塞几十秒到几分钟（首次加载权重），只能在后台线程调用。
    """
    import embedded_ai
    client = embedded_ai.shared()
    base = client.ensure_server()
    assets = client.assets() or {}
    model = assets['model'].stem if assets.get('model') else 'embedded'
    return base, model


def default_config():
    return {'backend': 'ollama', 'base_url': OLLAMA_DEFAULT, 'model': '',
            'api_key': '', 'temperature': 0.2, 'timeout': 120, 'max_tokens': 2048,
            'num_ctx': 8192,
            'vision_backend': 'openai', 'vision_base_url': SILICONFLOW_DEFAULT,
            'vision_model': QWEN_VL_MODEL, 'vision_api_key': '',
            'vision_max_pages': 12}


def load_config(path):
    cfg = default_config()
    try:
        saved = json.loads(path.read_text(encoding='utf-8'))
        if isinstance(saved, dict):
            cfg.update(saved)
    except (OSError, ValueError):
        pass
    return cfg


def save_config(path, cfg):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix('.tmp')
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
    tmp.replace(path)


def _same_base(a, b):
    return (a or '').strip().rstrip('/').lower() == (b or '').strip().rstrip('/').lower()


def vision_config(cfg):
    """Return a Client-compatible config for the vision model."""
    cfg = dict(cfg or {})
    backend = cfg.get('vision_backend') or 'openai'
    base = (cfg.get('vision_base_url') or '').strip().rstrip('/')
    model = (cfg.get('vision_model') or '').strip()
    key = (cfg.get('vision_api_key') or '').strip()
    same_backend = backend == (cfg.get('backend') or '')
    if same_backend and not base:
        # 视觉与文本用同一个服务（比如同一个中转站）时不必重复填地址和 Key
        base = (cfg.get('base_url') or '').strip().rstrip('/')
    if not key and same_backend and _same_base(base, cfg.get('base_url')):
        key = (cfg.get('api_key') or '').strip()
    return {
        'backend': backend,
        'base_url': base or (OLLAMA_DEFAULT if backend == 'ollama'
                             else ANTHROPIC_DEFAULT if backend == ANTHROPIC_BACKEND else ''),
        'model': model, 'api_key': key,
        'temperature': 0.0,
        'timeout': max(120, int(cfg.get('timeout', 120) or 120)),
        'max_tokens': max(4096, int(cfg.get('max_tokens', 2048) or 2048)),
        'num_ctx': int(cfg.get('num_ctx', 8192) or 8192),
        'vision': True,
    }


def vision_usable(cfg):
    cfg = dict(cfg or {})
    model = (cfg.get('vision_model') or '').strip()
    if not model:
        return False, '请先在 AI 设置 → 视觉模型选择开源视觉模型'
    vc = vision_config(cfg)
    if not vc.get('base_url'):
        return False, '请先填写视觉模型服务地址'
    ok, message = Client(vc).usable()
    if not ok:
        return False, ('视觉模型与文本模型不是同一个服务时，需要单独填写视觉模型 API Key。'
                       if not vc.get('api_key') else message)
    return True, ''


def vision_client(cfg):
    return Client(vision_config(cfg))


class Client:
    def __init__(self, cfg):
        self.cfg = dict(cfg or {})
        self.provider = self.cfg.get('backend', 'ollama')
        # 内置小模型走 llama-server 的 OpenAI 兼容接口，请求形状与云端一致
        self.backend = 'openai' if self.provider == EMBEDDED_BACKEND else self.provider
        self.base_url = (self.cfg.get('base_url') or '').strip().rstrip('/') \
            or (OLLAMA_DEFAULT if self.backend == 'ollama'
                else ANTHROPIC_DEFAULT if self.backend == ANTHROPIC_BACKEND
                else '' if self.provider == EMBEDDED_BACKEND else OPENAI_DEFAULT)
        if self.base_url and '://' not in self.base_url:
            self.base_url = ('http://' if self.backend == 'ollama' else 'https://') + self.base_url
        self.model = (self.cfg.get('model') or '').strip()
        self.api_key = (self.cfg.get('api_key') or '').strip()
        self.temperature = float(self.cfg.get('temperature', 0.2) or 0.2)
        self.timeout = float(self.cfg.get('timeout', 120) or 120)
        self.max_tokens = int(self.cfg.get('max_tokens', 2048) or 2048)
        self.num_ctx = int(self.cfg.get('num_ctx', 8192) or 8192)
        self.autostart_note = ''

    def _ensure_embedded(self):
        if self.provider != EMBEDDED_BACKEND:
            return
        if not self.base_url:
            self.base_url, resolved = embedded_endpoint()
            if not self.model:
                self.model = resolved

    def _ensure_local_server(self):
        """选了本机 Ollama 但服务没跑：自动拉起来，别让用户去开命令行。"""
        if self.provider == EMBEDDED_BACKEND or self.backend != 'ollama':
            return
        if self._is_remote():
            return
        if ollama_running(self.base_url):
            return
        ok, message = start_ollama()
        if not ok:
            raise AIError(message)
        self.autostart_note = message

    @staticmethod
    def _quiet_thinking(messages):
        """Qwen3 思考模式会让 0.6B 长篇自言自语；补 /no_think 让它直接给结果。"""
        out = []
        for message in messages:
            content = message.get('content', '')
            if isinstance(content, str) and '/no_think' not in content:
                content = content.rstrip() + ' /no_think'
            out.append({'role': message.get('role', 'user'), 'content': content})
        return out

    def usable(self):
        if self.provider == EMBEDDED_BACKEND:
            if not embedded_assets():
                return False, '没找到内置模型文件（ai_models / ai_runtime 需与本程序同目录）'
            return True, ''
        if not self.model:
            return False, '请先选择或填写模型名'
        if self.backend in ('openai', ANTHROPIC_BACKEND) and not self.api_key and self._is_remote():
            return False, '云端服务需要填写 API Key'
        return True, ''

    def _is_remote(self):
        try:
            host = (urlsplit(self.base_url).hostname or '').strip().lower()
        except ValueError:
            return True
        if host in ('', 'localhost', '127.0.0.1', '::1'):
            return False
        try:
            return not ipaddress.ip_address(host).is_private
        except ValueError:
            return True

    def describe(self):
        if self.provider == EMBEDDED_BACKEND:
            return EMBEDDED_LABEL
        backend = ('本机 Ollama' if self.backend == 'ollama'
                   else 'Anthropic Messages' if self.backend == ANTHROPIC_BACKEND
                   else '兼容 API')
        return f'{backend} · {self.model or "未选择模型"}'

    # ---------------- low level
    def _endpoint(self, suffix):
        base = self.base_url
        for tail in ('/chat/completions', '/api/chat', '/api/tags', '/models'):
            if base.endswith(tail):
                base = base[:-len(tail)]
                break
        return base.rstrip('/') + '/' + suffix

    def _messages_url(self):
        """Anthropic 风格端点：地址里已经带 /v1 就接 /messages，否则补 /v1/messages。"""
        base = self.base_url
        for tail in ('/messages', '/chat/completions', '/api/chat'):
            if base.endswith(tail):
                base = base[:-len(tail)]
                break
        if base.endswith('/v1'):
            return base + '/messages'
        return base.rstrip('/') + '/v1/messages'

    def _anthropic_headers(self):
        headers = {'anthropic-version': ANTHROPIC_VERSION}
        if self.api_key:
            headers['x-api-key'] = self.api_key
        return headers

    @staticmethod
    def _split_system(messages):
        """Anthropic 把 system 放在请求体顶层，不占对话轮次。"""
        system, turns = [], []
        for message in messages:
            role = message.get('role', 'user')
            content = message.get('content', '')
            if role == 'system':
                system.append(_content_text(content))
            else:
                turns.append({'role': 'assistant' if role == 'assistant' else 'user',
                              'content': content})
        if turns and turns[0]['role'] == 'assistant':
            turns.insert(0, {'role': 'user', 'content': '（接续上文）'})
        return '\n\n'.join(part for part in system if part), turns

    def _curl(self, url, payload=None, headers=None):
        """被按客户端指纹拦下时的兜底：改用系统 curl（Win10+ 自带）。"""
        exe = shutil.which('curl') or shutil.which('curl.exe')
        if not exe:
            return None, '本机没有 curl，没法换第二种客户端重试'
        limit = int(max(20, min(float(self.timeout), 600)))
        args = [exe, '-sS', '--max-time', str(limit), '-w', '\n%{http_code}',
                '-H', 'User-Agent: ' + USER_AGENT]
        for k, v in (headers or {}).items():
            args += ['-H', f'{k}: {v}']
        if payload is not None:
            args += ['-X', 'POST', '-H', 'Content-Type: application/json',
                     '--data-binary', json.dumps(payload, ensure_ascii=False)]
        args.append(url)
        try:
            done = subprocess.run(args, capture_output=True, timeout=limit + 20)
        except (OSError, subprocess.TimeoutExpired) as exc:
            return None, f'curl 也没跑通：{str(exc)[:120]}'
        raw = (done.stdout or b'').decode('utf-8', 'replace')
        head, _, code = raw.rpartition('\n')
        status = (code or '').strip() or str(done.returncode)
        if not head.strip() and done.returncode:
            return None, 'curl 失败：' + (done.stderr or b'').decode('utf-8', 'replace')[:160]
        if not (status.isdigit() and 200 <= int(status) < 300):
            return None, f'curl 返回 {status}：{head[:200]}'
        return head, ''

    def _fail(self, url, code, detail, headers=None, payload=None):
        """403 里带 Cloudflare 特征的，先换 curl 重试一次；换不动就把原因说清楚。"""
        if not is_blocked(code, detail):
            raise AIError(f'服务返回 {code}：{detail[:300]}')
        body, why = self._curl(url, payload, headers)
        if body is None:
            raise AIError(f'服务返回 {code}：{detail[:200]}\n\n'
                          '这类服务在 Cloudflare 后面按浏览器指纹拦非浏览器客户端。'
                          + (why or '换用官方地址、本机代理，或改用 OpenAI 兼容协议再试。'))
        return body

    def _post(self, url, payload, headers=None):
        headers = dict(headers or {})
        headers.setdefault('User-Agent', USER_AGENT)
        data = json.dumps(payload).encode('utf-8')
        req = urllib.request.Request(url, data=data, method='POST')
        req.add_header('Content-Type', 'application/json')
        for k, v in headers.items():
            req.add_header(k, v)
        block = None
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                body = resp.read().decode('utf-8', 'replace')
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode('utf-8', 'replace')[:400] if hasattr(exc, 'read') else ''
            block = self._fail(url, exc.code, detail, headers, payload)
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            raise AIError(f'无法连接 {url}：{exc}') from exc
        if block is not None:
            body = block
        try:
            return json.loads(body)
        except ValueError as e:
            raise AIError(f'返回不是 JSON：{body[:120]}') from e

    def _get(self, url, headers=None):
        headers = dict(headers or {})
        headers.setdefault('User-Agent', USER_AGENT)
        req = urllib.request.Request(url)
        for k, v in headers.items():
            req.add_header(k, v)
        block = None
        try:
            with urllib.request.urlopen(req, timeout=min(30.0, float(self.timeout))) as resp:
                body = resp.read().decode('utf-8', 'replace')
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode('utf-8', 'replace')[:400] if hasattr(exc, 'read') else ''
            block = self._fail(url, exc.code, detail, headers)
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            raise AIError(f'无法连接 {url}：{exc}') from exc
        if block is not None:
            body = block
        try:
            return json.loads(body)
        except ValueError as e:
            raise AIError(f'返回不是 JSON：{body[:120]}') from e

    def chat(self, messages, temperature=None, max_tokens=None, json_mode=False,
             think=False):
        ok, msg = self.usable()
        if not ok:
            raise AIError(msg)
        self._ensure_embedded()
        self._ensure_local_server()
        temp = self.temperature if temperature is None else float(temperature)
        limit = self.max_tokens if max_tokens is None else int(max_tokens)
        if self.backend == 'ollama':
            options = {'temperature': temp, 'num_ctx': self.num_ctx}
            if limit:
                options['num_predict'] = limit
            payload = {'model': self.model, 'messages': messages,
                       'stream': False, 'options': options,
                       'think': bool(think), 'keep_alive': '30m'}
            if json_mode:
                payload['format'] = 'json'
            out = self._post(self._endpoint('api/chat'), payload)
            if out.get('error'):
                raise AIError(str(out['error'])[:300])
            content = (out.get('message') or {}).get('content', '')
            if not content and (out.get('message') or {}).get('thinking'):
                # 旧版 Ollama 不认 think 参数：退一步用 /no_think 提示词再试一次
                retry = dict(payload)
                retry['messages'] = self._quiet_thinking(messages)
                out = self._post(self._endpoint('api/chat'), retry)
                content = (out.get('message') or {}).get('content', '')
            if not content and (out.get('message') or {}).get('thinking'):
                reasoning = (out.get('message') or {}).get('thinking') or ''
                if json_mode:
                    # 思考模型常把最终 JSON 写在推理末尾；能回收就直接用
                    salvaged = _extract_json(reasoning)
                    if salvaged is not None:
                        return json.dumps(salvaged, ensure_ascii=False)
                raise AIError('模型只输出了思考过程，请换用非思考模型，或把输出上限调大')
            return content
        if self.backend == ANTHROPIC_BACKEND:
            system, turns = self._split_system(messages)
            payload = {'model': self.model, 'max_tokens': limit or 2048,
                       'messages': turns, 'temperature': temp}
            if system:
                payload['system'] = system
            url = self._messages_url()
            try:
                out = self._post(url, payload, self._anthropic_headers())
            except AIError as exc:
                # 较新的 claude 通过中转站会拒绝 temperature，去掉再试一次
                if 'temperature' not in str(exc).lower():
                    raise
                payload.pop('temperature', None)
                out = self._post(url, payload, self._anthropic_headers())
            content = _content_text(out.get('content'))
            if not content:
                raise AIError('模型没有返回正文：' + json.dumps(out, ensure_ascii=False)[:260])
            return content
        headers = {'Authorization': f'Bearer {self.api_key}'} if self.api_key else {}
        payload = {'model': self.model,
                   'messages': (self._quiet_thinking(messages)
                                if self.provider == EMBEDDED_BACKEND else messages),
                   'temperature': temp}
        if limit:
            payload['max_tokens'] = limit
        out = self._post(self._endpoint('chat/completions'), payload, headers)
        try:
            choice = out['choices'][0]
            content = choice['message'].get('content') or ''
            if not content:
                content = choice['message'].get('reasoning_content') or ''
        except (KeyError, IndexError):
            raise AIError(json.dumps(out, ensure_ascii=False)[:300]) from None
        return content

    def chat_vision(self, messages, images, temperature=None, max_tokens=None,
                    json_mode=False, think=False):
        """Send page screenshots to an OpenAI-compatible VLM or Ollama vision model."""
        ok, msg = self.usable()
        if not ok:
            raise AIError(msg)
        self._ensure_embedded()
        self._ensure_local_server()
        temp = self.temperature if temperature is None else float(temperature)
        limit = self.max_tokens if max_tokens is None else int(max_tokens)
        imgs = [_as_data_url(x) for x in (images or []) if x]
        msgs = []
        for message in messages:
            content = message.get('content', '')
            msgs.append({'role': message.get('role', 'user'),
                         'content': _content_text(content)})
        for message in reversed(msgs):
            if message.get('role') != 'user':
                continue
            if self.backend == 'ollama':
                message['images'] = [_image_b64(x) for x in imgs]
            elif self.backend == ANTHROPIC_BACKEND:
                blocks = [{'type': 'image',
                           'source': {'type': 'base64', 'media_type': media, 'data': data}}
                          for media, data in (_split_data_url(x) for x in imgs)]
                blocks.append({'type': 'text', 'text': message.get('content', '')})
                message['content'] = blocks
            else:
                parts = [{'type': 'text', 'text': message.get('content', '')}]
                parts += [{'type': 'image_url',
                           'image_url': {'url': x}} for x in imgs]
                message['content'] = parts
            break
        if self.backend == ANTHROPIC_BACKEND:
            system, turns = self._split_system(msgs)
            payload = {'model': self.model, 'max_tokens': limit or 4096,
                       'messages': turns, 'temperature': temp}
            if system:
                payload['system'] = system
            url = self._messages_url()
            try:
                out = self._post(url, payload, self._anthropic_headers())
            except AIError as exc:
                if 'temperature' not in str(exc).lower():
                    raise
                payload.pop('temperature', None)
                out = self._post(url, payload, self._anthropic_headers())
            content = _content_text(out.get('content'))
            if not content:
                raise AIError('视觉模型没有返回内容：' + json.dumps(
                    out, ensure_ascii=False)[:260])
            return content
        if self.backend == 'ollama':
            options = {'temperature': temp, 'num_ctx': self.num_ctx}
            if limit:
                options['num_predict'] = limit
            payload = {'model': self.model, 'messages': msgs,
                       'stream': False, 'options': options,
                       'think': bool(think), 'keep_alive': '30m'}
            if json_mode:
                payload['format'] = 'json'
            out = self._post(self._endpoint('api/chat'), payload)
            if out.get('error'):
                raise AIError(str(out['error'])[:300])
            message = out.get('message') or {}
            content = _content_text(message.get('content')) or _content_text(message.get('thinking'))
            if not content:
                raise AIError('视觉模型没有返回内容，请确认所用模型支持图片输入')
            return content
        headers = {'Authorization': f'Bearer {self.api_key}'} if self.api_key else {}
        payload = {'model': self.model, 'messages': msgs, 'temperature': temp}
        if limit:
            payload['max_tokens'] = limit
        try:
            host = (urlsplit(self.base_url).hostname or '').lower()
        except ValueError:
            host = ''
        if 'dashscope.aliyuncs.com' in host:
            payload['enable_thinking'] = False
        out = self._post(self._endpoint('chat/completions'), payload, headers)
        try:
            message = out['choices'][0]['message']
        except (KeyError, IndexError):
            raise AIError(json.dumps(out, ensure_ascii=False)[:300]) from None
        content = _content_text(message.get('content')) or \
            _content_text(message.get('reasoning_content'))
        if not content:
            raise AIError('视觉模型没有返回内容：' + json.dumps(
                message, ensure_ascii=False)[:200])
        return content

    def list_models(self):
        self._ensure_local_server()
        if self.provider == EMBEDDED_BACKEND:
            self._ensure_embedded()
            return [self.model] if self.model else []
        if self.backend == 'ollama':
            try:
                data = self._get(self._endpoint('api/tags'))
            except Exception as exc:
                raise AIError(f'Ollama 不可达：{exc}') from exc
            return [m.get('name', '') for m in data.get('models', []) if m.get('name')]
        headers = ({'Authorization': f'Bearer {self.api_key}'} if self.api_key else {})
        if self.backend == ANTHROPIC_BACKEND:
            headers = self._anthropic_headers()
        try:
            data = self._get(self._endpoint('models'), headers)
            ids = [m.get('id', '') for m in data.get('data', []) if m.get('id')]
            return ids[:100]
        except Exception as exc:
            raise AIError(f'模型列表获取失败：{exc}') from exc

    def ping(self):
        models = self.list_models()
        return f'连接成功，发现 {len(models)} 个模型' if models else '连接成功，但服务没有返回模型'

    # ---------------- structured helpers
    def ask_json(self, system, user, expect_key='', temperature=0.1, max_tokens=None,
                 think=False):
        messages = [
            {'role': 'system', 'content': system + '\n只输出一个 JSON 对象，不要输出解释、'
             '不要使用 markdown 代码围栏。'},
            {'role': 'user', 'content': user},
        ]
        limit = max_tokens or max(self.max_tokens, 4096)
        last_err = None
        for _ in range(2):
            raw = self.chat(messages, temperature=temperature, max_tokens=limit,
                            json_mode=True, think=think)
            obj = _extract_json(raw)
            if obj is not None and (not expect_key or expect_key in obj):
                return obj
            last_err = raw
            messages.append({'role': 'assistant', 'content': raw})
            messages.append({'role': 'user', 'content': '输出无效，请重新只输出要求的 JSON 对象。'})
        raise AIError(f'模型未按格式返回 JSON：{str(last_err)[:200]}')

    def ask_vision_json(self, system, user, images, expect_key='fields',
                        temperature=0.0, max_tokens=None):
        messages = [
            {'role': 'system', 'content': system + '\n只输出一个 JSON 对象，不要输出解释、'
             '不要使用 markdown 代码围栏。'},
            {'role': 'user', 'content': user},
        ]
        limit = max_tokens or (1600 if self.backend == 'ollama'
                               else max(self.max_tokens, 4096))
        last_err = None
        for _ in range(2):
            raw = self.chat_vision(messages, images, temperature=temperature,
                                   max_tokens=limit, json_mode=True)
            obj = _extract_json(raw)
            if obj is not None and (not expect_key or expect_key in obj):
                return obj
            last_err = raw
            messages.append({'role': 'assistant', 'content': raw})
            messages.append({'role': 'user',
                             'content': '输出无效，请重新只输出要求的 JSON 对象。'})
        raise AIError(f'视觉模型未按格式返回 JSON：{str(last_err)[:200]}')


def _content_text(value):
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        pieces = []
        for item in value:
            if isinstance(item, dict):
                if item.get('type') in ('text', 'output_text'):
                    pieces.append(str(item.get('text') or ''))
                elif 'content' in item:
                    pieces.append(_content_text(item.get('content')))
            elif isinstance(item, str):
                pieces.append(item)
        return ''.join(pieces)
    return '' if value is None else str(value)


def _as_data_url(value):
    if isinstance(value, bytes):
        mime = 'image/png' if value[:8] == b'\x89PNG\r\n\x1a\n' else 'image/jpeg'
        return f'data:{mime};base64,' + base64.b64encode(value).decode('ascii')
    text = str(value or '')
    if text.startswith(('data:', 'http://', 'https://')):
        return text
    return 'data:image/jpeg;base64,' + text


def _image_b64(value):
    text = _as_data_url(value)
    return text.split(',', 1)[1] if text.startswith('data:') and ',' in text else text


def _split_data_url(value):
    """拆成 (media_type, base64)，Anthropic 的图片块要这个形状。"""
    text = _as_data_url(value)
    if text.startswith('data:') and ',' in text:
        head, data = text.split(',', 1)
        return (head[5:].split(';')[0] or 'image/png'), data
    return 'image/png', text


def _extract_json(text):
    text = re.sub(r'^```[a-zA-Z]*\s*|\s*```$', '', (text or '').strip())
    start = text.find('{')
    if start < 0:
        return None
    depth, in_str, esc = 0, False, False
    for i, ch in enumerate(text[start:], start):
        if in_str:
            if esc:
                esc = False
            elif ch == '\\':
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(text[start:i + 1])
                except ValueError:
                    return None
    return None


# ---------------- prompt builders (used by GUI)

FIELD_SYSTEM = (
    '你在辅助一个 Word 模板工具识别“可填写字段”并归类。给定文档结构摘录、候选位置和当前分组，'
    '请为每个候选给出：规范的中文字段名 label（2-10字，来自文档上下文，不要杜撰）、'
    '输入方式 type（短答案 text，长段落 multiline）、'
    '分组 group（如 基本信息/项目信息/人员/正文/经费/落款，2-6个字，同类字段同组）。'
    '同时根据文档标题和内容给出 template_title（简短模板名）与 category'
    '（科研项目/行政公文/人事/财务/教学/其他 之一）。'
    '已有明确结构或占位提示的候选，label 必须沿用现名，不得改名。只能基于给出的候选编号作答，不得新增字段。')


def field_prompt(outline, candidates, template_title=''):
    lines = ['【文档标题】' + (template_title or '（未提供）')]
    if outline:
        lines.append('【文档正文摘录（段落序号: 内容）】')
        lines += [f'{i}: {str(t)[:70]}' for i, t in outline[:160]]
    lines.append('【候选字段（k 从 0 开始）】')
    for k, c in enumerate(candidates):
        lines.append(f'{k}: 位置 {c.get("where", "")}；现名 {c.get("label", "")}；'
                     f'输入方式 {c.get("type", "")}；结构明确 {"是" if c.get("marked") else "否"}；'
                     f'上下文“{str(c.get("context", ""))[:70]}”')
    lines.append('【输出格式】')
    lines.append('{"template_title":"项目任务书","category":"科研项目","fields":'
                 '[{"i":0,"label":"项目名称","type":"text","group":"项目信息"}]}')
    return '\n'.join(lines)


BODY_SYSTEM = (
    '你是熟悉中文公文、科研项目申报材料的写作助手。根据模板信息、字段清单、已填写内容和写作要求，'
    '为指定字段撰写可直接粘贴进 Word 的正文。'
    '要求：只输出正文本身，不要标题、不要解释、不要 markdown 代码围栏，不要“以下是”之类的开场；'
    '语言正式、客观、具体，层次清楚，可用“一、”“1.”“（1）”或自然段，但不要使用 markdown 符号；'
    '不得编造未提供的人名、单位、数据、时间等具体事实，可用“拟”“计划”“将”等表述；'
    '不重复字段名作为小标题；长度接近目标字数，且不得超过给定上限。')


def body_prompt(template_title, description, fields, values, target, requirement='',
                target_chars=600, max_length=0, existing='', append=False):
    filled = []
    for field in fields:
        key = field.get('key')
        value = str(values.get(key, '') or '').strip()
        if value and key != target and len(value) <= 400:
            filled.append(f'- {field.get("label", key)}：{value}')
    lines = [f'【模板】{template_title}']
    if description:
        lines.append(f'【模板说明】{description}')
    if filled:
        lines.append('【已填写内容（可用于保持上下文一致）】')
        lines += filled[:30]
    lines.append(f'【本次要写的字段】{target}')
    lines.append(f'【目标字数】约 {target_chars} 字' + (f'，上限 {max_length} 字' if max_length else ''))
    if existing:
        lines.append('【该字段现有内容】' + existing[:1200])
        lines.append('【处理方式】' + ('在现有内容后续写，保持衔接，不要重复已有段落。' if append else '重写该字段内容。'))
    if requirement:
        lines.append('【写作要求】' + requirement)
    else:
        lines.append('【写作要求】内容具体、可执行，符合该模板的正式语体。')
    return '\n'.join(lines)


FORMAT_SYSTEM = (
    '你在辅助 Word 模板工具设置“填写内容的字体格式”。中文字号对照：初号=42,小初=36,一号=26,'
    '小一=24,二号=22,小二=18,三号=16,小三=15,四号=14,小四=12,五号=10.5,小五=9。'
    '根据用户要求输出：body_size_pt（正文字号，磅，可为null）、apply_to_all（该字号是否也应用到所有短字段，'
    '布尔）、default_size_pt（所有填写内容的默认字号，可为null）、field_size_pt（按字段名覆盖字号）、'
    'field_bold（需加粗的字段名数组）、field_font（按字段名指定字体名）、note（一句话说明）。'
    '字段名只能取自给定清单；用户没提到的就不要输出；字号必须在 8 到 42 之间。')


def format_prompt(fields, current, request, body_fields=None):
    return ('【字段清单】' + '、'.join(fields) + '\n'
            '【正文字段（多行）】' + '、'.join(body_fields or []) + '\n'
            '【当前设置】' + json.dumps(current, ensure_ascii=False) + '\n'
            '【用户要求】' + request + '\n'
            '【输出格式】\n'
            '{"body_size_pt":12,"apply_to_all":false,"default_size_pt":null,'
            '"field_size_pt":{"学号":10.5},"field_bold":["项目名称"],'
            '"field_font":{"项目名称":"黑体"},"note":"正文小四"}')
