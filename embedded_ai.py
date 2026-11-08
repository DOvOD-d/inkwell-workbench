"""内置离线小模型：用 llama.cpp 运行 GGUF，仅负责字段命名等短输出任务。

运行时和模型作为 exe 旁的两个目录分发：
  ai_runtime/llama-server.exe + dll
  ai_models/*.gguf
源目录调试时也支持 WTT_AI_HOME 指向包含这两个目录的目录。
"""
from __future__ import annotations

import atexit
import json
import os
import re
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path

import ai

SCHEMA = {
    'type': 'object',
    'properties': {
        'names': {
            'type': 'array',
            'items': {
                'type': 'object',
                'properties': {
                    'i': {'type': 'integer'},
                    'name': {'type': 'string'},
                },
                'required': ['i', 'name'],
                'additionalProperties': False,
            },
        }
    },
    'required': ['names'],
    'additionalProperties': False,
}

SYSTEM = (
    '你是 Word 表单字段命名助手。为每个空白填写区域起一个 4-12 个中文字符的字段名，用于表单显示。'
    '规则：结合章节、任务编号和上下文；例如「任务一代码内容」「第2题程序分析」「思考题1回答内容」。'
    '不要包含 {{ }}、引号、标点或换行；不要用“字段1”；同一表单内名称不能重复。'
    '如果给了参考名且准确，可以沿用；不准确则给出更准确的名称。'
    '只输出 JSON：{"names":[{"i":0,"name":"..."}]}。 /no_think'
)


def _bases():
    env = os.environ.get('WTT_AI_HOME')
    if env:
        yield Path(env)
    if getattr(sys, 'frozen', False):
        yield Path(sys.executable).resolve().parent
    yield Path(__file__).resolve().parent


def _pick_model(models):
    if not models:
        return None
    preferred = [m for m in models if 'qwen3' in m.name.lower()]
    preferred = preferred or models
    return sorted(preferred, key=lambda p: (len(p.name), p.name))[0]


def find_assets():
    """返回 {'server','model','label'} 或 None。"""
    env_model = os.environ.get('WTT_AI_MODEL')
    for base in _bases():
        for root in (base, base / 'embedded_ai', base / 'embedded-test'):
            runtime = root / 'ai_runtime'
            model_dir = root / 'ai_models'
            server = runtime / 'llama-server.exe'
            if not server.is_file():
                server = root / 'llama-server.exe'
            if not server.is_file():
                continue
            if env_model:
                model = Path(env_model)
            else:
                models = list(model_dir.glob('*.gguf')) if model_dir.is_dir() else []
                model = _pick_model(models) if models else None
            if not model or not model.is_file():
                continue
            match = re.match(r'(.+?)[-_.]Q\d', model.stem, re.IGNORECASE)
            label = (match.group(1) if match else model.stem).replace('_', ' ').replace('-', ' ')
            return {'server': server, 'model': model, 'label': label}
    return None


def missing_hint():
    return ('未找到内置模型。请把 ai_runtime/（含 llama-server.exe）和 '
            'ai_models/（含 .gguf）放在程序所在目录。')


class EmbeddedAI:
    def __init__(self, log_path=None):
        self._lock = threading.RLock()
        self._proc = None
        self._base = ''
        self._model_name = ''
        self._log = None
        self._log_path = Path(log_path) if log_path else Path(tempfile.gettempdir()) / 'wtt_embedded_ai.log'
        self._registered = False

    def assets(self):
        return find_assets()

    def label(self):
        assets = self.assets()
        return assets['label'] if assets else ''

    def _free_port(self):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('127.0.0.1', 0))
            return s.getsockname()[1]

    def _health(self):
        if not self._base:
            return False
        for path in ('/health', '/v1/models'):
            try:
                opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
                with opener.open(self._base + path, timeout=1.5) as resp:
                    return resp.status == 200
            except Exception:
                continue
        return False

    def _read_log_tail(self):
        try:
            data = self._log_path.read_bytes()
            return data[-1800:].decode('utf-8', 'replace')
        except OSError:
            return ''

    def ensure_server(self):
        """启动并等待 llama-server，返回 OpenAI 兼容 base_url（含 /v1）。"""
        with self._lock:
            if self._proc is not None and self._proc.poll() is None and self._health():
                return self._base + '/v1'
            if self._proc is not None and self._proc.poll() is not None:
                self._close_log()
                self._proc = None
            assets = self.assets()
            if not assets:
                raise RuntimeError(missing_hint())
            port = self._free_port()
            threads = str(max(2, min(8, os.cpu_count() or 4)))
            cmd = [str(assets['server']), '-m', str(assets['model']),
                   '--host', '127.0.0.1', '--port', str(port),
                   '-c', '4096', '-t', threads, '-ngl', '0', '--jinja']
            self._log_path.parent.mkdir(parents=True, exist_ok=True)
            self._log = open(self._log_path, 'ab')
            self._log.write(('\n=== start ' + time.strftime('%Y-%m-%d %H:%M:%S') + ' ===\n').encode())
            self._log.flush()
            flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0) if os.name == 'nt' else 0
            try:
                self._proc = subprocess.Popen(
                    cmd, cwd=str(assets['server'].parent),
                    stdout=self._log, stderr=subprocess.STDOUT,
                    creationflags=flags)
            except OSError as exc:
                self._close_log()
                raise RuntimeError(f'无法启动内置模型运行时：{exc}') from exc
            if not self._registered:
                atexit.register(self.stop)
                self._registered = True
            self._base = f'http://127.0.0.1:{port}'
            self._model_name = assets['model'].stem
            deadline = time.time() + 240
            while time.time() < deadline:
                if self._proc.poll() is not None:
                    detail = self._read_log_tail()
                    self._proc = None
                    self._close_log()
                    raise RuntimeError('内置模型启动失败。\n' + detail)
                if self._health():
                    return self._base + '/v1'
                time.sleep(0.4)
            self.stop()
            raise RuntimeError('内置模型启动超时，请检查电脑内存是否足够。')

    def _close_log(self):
        if self._log is not None:
            try:
                self._log.close()
            except OSError:
                pass
            self._log = None

    def stop(self):
        with self._lock:
            proc, self._proc = self._proc, None
            if proc is not None and proc.poll() is None:
                try:
                    proc.terminate()
                    proc.wait(timeout=5)
                except Exception:
                    try:
                        proc.kill()
                    except Exception:
                        pass
            self._close_log()

    def _post(self, url, payload, timeout=180):
        data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        req = urllib.request.Request(url, data=data, method='POST')
        req.add_header('Content-Type', 'application/json')
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        try:
            with opener.open(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode('utf-8', 'replace'))
        except urllib.error.HTTPError as exc:
            body = exc.read().decode('utf-8', 'replace')[:300]
            raise RuntimeError(f'内置模型返回 {exc.code}：{body}') from exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            raise RuntimeError(f'内置模型连接失败：{exc}') from exc

    def _chat_json(self, system, user, schema, max_tokens=512):
        base = self._base + '/v1/chat/completions'
        payload = {
            'model': self._model_name,
            'messages': [{'role': 'system', 'content': system},
                         {'role': 'user', 'content': user}],
            'temperature': 0.1, 'max_tokens': max_tokens, 'stream': False,
            'response_format': {'type': 'json_schema',
                                'json_schema': {'name': 'field_names', 'schema': schema}},
        }
        attempts = [payload,
                    {k: v for k, v in payload.items() if k != 'response_format'},
                    {**{k: v for k, v in payload.items() if k != 'response_format'},
                     'response_format': {'type': 'json_object'}}]
        last = None
        for body in attempts:
            try:
                out = self._post(base, body)
                content = (out.get('choices') or [{}])[0].get('message', {}).get('content', '')
                obj = ai._extract_json(content)
                if isinstance(obj, dict):
                    return obj
                last = content[:200]
            except RuntimeError as exc:
                last = str(exc)
        raise RuntimeError(f'内置模型没有返回有效 JSON：{last}')

    def suggest_names(self, items, template_title='', existing=(), progress=None,
                      should_stop=None):
        """items: [{'i':int,'section':str,'label':str,'context':str,'hint':str}]

        progress(done, total) 与 should_stop() 供界面使用：取消只停在未来的一批上，
        已经补出来的名字照常保留。
        """
        if not items:
            return {}, []
        self.ensure_server()
        result, errors = {}, []
        used = {x for x in existing if x}
        total = len(items)
        for start in range(0, total, 6):
            if should_stop and should_stop():
                errors.append(f'已取消：剩余 {total - start} 处没有补全，字段名保持原样。')
                break
            batch = items[start:start + 6]
            lines = []
            for pos, item in enumerate(batch):
                lines.append(
                    f"{pos}: 章节={item.get('section','')[:40]}；标签={item.get('label','')[:80]}；"
                    f"上下文={item.get('context','')[:220]}；参考名={item.get('hint','')[:24]}")
            user = (f'模板标题：{template_title}\n已有字段名：' + '、'.join(sorted(used)[:40]) +
                    '\n空白区域：\n' + '\n'.join(lines) +
                    '\n只输出 JSON，i 使用上面的编号。')
            try:
                obj = self._chat_json(SYSTEM, user, SCHEMA, max_tokens=512)
            except Exception as exc:
                errors.append(str(exc))
                continue
            for entry in obj.get('names', []):
                try:
                    pos = int(entry.get('i'))
                    name = _clean_name(entry.get('name'))
                except (TypeError, ValueError):
                    continue
                if 0 <= pos < len(batch) and name:
                    result[batch[pos]['i']] = name
            if progress:
                progress(min(start + 6, total), total)
        return result, errors


def _clean_name(value):
    text = re.sub(r'[{}\[\]"\'“”‘’\r\n\t]', '', str(value or '')).strip()
    text = re.sub(r'[：:，,。；;、!！?？]+$', '', text)
    text = re.sub(r'\s+', '', text)
    if not (2 <= len(text) <= 24):
        return ''
    return text


_instance = None
_instance_lock = threading.Lock()


def shared(log_path=None):
    global _instance
    with _instance_lock:
        if _instance is None:
            _instance = EmbeddedAI(log_path)
        elif log_path:
            _instance._log_path = Path(log_path)
        return _instance
