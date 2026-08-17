# -*- coding: utf-8 -*-
"""
租车经营分析工作台 · 本地预览服务
================================
用法：
    python server.py
默认 http://localhost:8080  （端口可在 config.json 的 server_port 修改）

能力：
    GET  /              预览根目录 index.html（TrailScope 风格工作台）
    GET  /data.json     当前统一订单数据（index.html 加载用）
    GET  /schema_mapping.json  表头统一映射配置
    GET  /api/metrics   返回最新 metrics.json（Python 看板用）
    POST /api/refresh   重跑整条管线（ingest→transform→analyze→render），
                        并同步刷新前端 data.json，返回「哪些文件变了」+ 最新指标摘要
"""
import os
import sys
import json
import hashlib
import shutil
import subprocess
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUTPUT = os.path.join(ROOT, 'tools', 'sample_outputs')
DASHBOARD = os.path.join(OUTPUT, '分析看板.html')
GEN_SCRIPT = os.path.join(HERE, 'generate_dashboard.py')

CONFIG = {}
CFG_PATH = os.path.join(HERE, 'config.json')
if os.path.exists(CFG_PATH):
    try:
        with open(CFG_PATH, 'r', encoding='utf-8') as f:
            CONFIG = json.load(f)
    except Exception:
        pass
PORT = int(CONFIG.get('server_port', 8080))


def sha256_of(path):
    if not os.path.exists(path):
        return None
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def run_pipeline():
    """重跑管线，并同步前端 data.json。返回 (ok, msg, changed_files)。"""
    baseline = {n: sha256_of(os.path.join(OUTPUT, n))
                for n in ['分析看板.html', 'metrics.json', 'unified_orders.json']}
    try:
        subprocess.run([sys.executable, GEN_SCRIPT], cwd=HERE,
                       check=True, capture_output=True, text=True,
                       encoding='utf-8', errors='replace')
    except subprocess.CalledProcessError as e:
        return False, f"管线运行失败：{e.stderr[:500]}", []
    # 同步前端 data.json（根目录）
    src = os.path.join(OUTPUT, 'unified_orders.json')
    dst = os.path.join(ROOT, 'data.json')
    if os.path.exists(src):
        shutil.copyfile(src, dst)
    after = {n: sha256_of(os.path.join(OUTPUT, n))
             for n in ['分析看板.html', 'metrics.json', 'unified_orders.json']}
    changed = [n for n in after if baseline.get(n) != after.get(n)]
    return True, f"重跑完成，{'无变化' if not changed else '更新了 ' + '、'.join(changed) + '（含前端 data.json）'}", changed


def load_metrics():
    mp = os.path.join(OUTPUT, 'metrics.json')
    if not os.path.exists(mp):
        return None
    try:
        with open(mp, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def ctype_for(ext):
    return {
        '.html': 'text/html; charset=utf-8',
        '.js': 'application/javascript; charset=utf-8',
        '.css': 'text/css; charset=utf-8',
        '.json': 'application/json; charset=utf-8',
        '.png': 'image/png',
        '.svg': 'image/svg+xml',
    }.get(ext, 'application/octet-stream')


class Handler(BaseHTTPRequestHandler):
    def _serve_file(self, path, ctype=None):
        try:
            with open(path, 'rb') as f:
                data = f.read()
        except Exception:
            self._send(404, {'error': 'file not found'})
            return
        ext = os.path.splitext(path)[1].lower()
        self.send_response(200)
        self.send_header('Content-Type', ctype or ctype_for(ext))
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send(self, code, body, ctype='application/json; charset=utf-8'):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        route = parsed.path
        # 根目录静态文件（index.html / data.json / schema_mapping.json / 3_报表产出 下文件）
        if route in ('/', '/index.html'):
            self._serve_file(os.path.join(ROOT, 'index.html'))
            return
        if route == '/api/metrics':
            self._send(200, load_metrics() or {'error': 'no metrics'})
            return
        rel = os.path.normpath(route.lstrip('/'))
        # 防目录穿越
        if rel.startswith('..') or os.path.isabs(rel):
            self._send(404, {'error': 'not found'})
            return
        fp = os.path.join(ROOT, rel)
        if os.path.isfile(fp):
            self._serve_file(fp)
            return
        self._send(404, {'error': 'not found'})

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/refresh':
            ok, msg, changed = run_pipeline()
            self._send(200, {
                'ok': ok,
                'message': msg,
                'changed': changed,
                'metrics': load_metrics(),
            })
            return
        self._send(404, {'error': 'not found'})

    def log_message(self, fmt, *args):
        sys.stdout.write("[server] " + (fmt % args) + "\n")


def main():
    os.makedirs(OUTPUT, exist_ok=True)
    srv = ThreadingHTTPServer(('0.0.0.0', PORT), Handler)
    print(f"租车经营分析工作台已启动： http://localhost:{PORT}")
    print("  预览工作台：浏览器打开上面的地址（index.html）")
    print("  刷新数据：点击看板右上角按钮，或 POST /api/refresh")
    print("  退出服务：Ctrl+C")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")


if __name__ == '__main__':
    main()
