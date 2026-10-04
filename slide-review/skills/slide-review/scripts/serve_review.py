# -*- coding: utf-8 -*-
"""serve_review.py —— 审阅页本地服务：静态（显式禁用缓存）+ 应用/预览/撤销接口。

为什么要用它而不是 `python -m http.server`：
1. no-store：http.server 只发 Last-Modified，浏览器会按启发式规则把渲染图缓存住
   —— 改了图、重新生成页面后，用户刷新仍看到旧图（实测踩过）。
2. POST /apply：接收网页攒好的改动清单（JSON），落盘到 <dir>/apply/pending_<ts>.json，
   并以当前解释器后台启动 apply_changes.py 写回 pptx + 重渲染；前端凭返回的
   job 轮询 GET /status?job=<ts>（读 apply/status_<ts>.json）直到完成。
3. POST /preview：预览——同样的改动写回临时副本、只渲染涉及页（preview_changes.py），
   不触碰原件；预览图 apply/preview_<ts>_slideNN.png 由前端直接静态取用，
   状态走 GET /status?job=<ts>&type=preview。
4. POST /revert：撤销最近一次应用——从该次备份恢复 pptx 并重渲染涉及页
   （revert_changes.py），状态走 GET /status?job=<ts>&type=revert。

用法: python serve_review.py <port> <dir>
"""
import sys, os, json, time, re, subprocess, functools, http.server, socketserver

APPLY_PY = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'apply_changes.py')
PREVIEW_PY = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'preview_changes.py')
REVERT_PY = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'revert_changes.py')

# docx 审阅需要另外的 docx 工具包（apply_changes_docx.py 三件套）；把其目录设进
# SLIDE_REVIEW_DOCX_TOOLKIT 环境变量即启用，缺省仅支持 PPT
_DOCX = os.environ.get('SLIDE_REVIEW_DOCX_TOOLKIT', '')
SCRIPTS = {
    'ppt': {'/apply': APPLY_PY, '/preview': PREVIEW_PY, '/revert': REVERT_PY},
    'docx': ({'/apply': os.path.join(_DOCX, 'apply_changes_docx.py'),
              '/preview': os.path.join(_DOCX, 'preview_changes_docx.py'),
              '/revert': os.path.join(_DOCX, 'revert_changes_docx.py')}
             if _DOCX and os.path.isdir(_DOCX) else None),
}


class ReviewHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def log_message(self, *a):
        pass

    # ---- helpers ----
    def json_resp(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    # ---- POST /apply、/preview、/revert ----
    def do_POST(self):
        ep = self.path.rstrip('/')
        if ep not in ('/apply', '/preview', '/revert'):
            return self.json_resp(404, {'ok': False, 'msg': 'unknown endpoint'})
        n = int(self.headers.get('Content-Length') or 0)
        if n <= 0 or n > 32_000_000:   # 图片替换走 base64 dataURL，放宽到 32MB
            return self.json_resp(400, {'ok': False, 'msg': 'body 为空或超过 32MB'})
        try:
            data = json.loads(self.rfile.read(n).decode('utf-8'))
        except Exception:
            return self.json_resp(400, {'ok': False, 'msg': '不是合法 JSON'})

        try:
            meta = json.load(open(os.path.join(self.directory, 'meta.json'),
                                  encoding='utf-8'))
        except Exception:
            return self.json_resp(500, {'ok': False, 'msg': '本目录没有 meta.json，'
                    '请用新版 make_review.py 重新生成审阅页'})
        kind = meta.get('kind', 'ppt')
        if data.get('key') != meta.get('key') or (
                ep != '/revert' and data.get('pptx') != meta.get('pptx')):
            return self.json_resp(403, {'ok': False, 'msg': '请求与审阅页不匹配'
                    '（key/pptx 校验失败），请刷新页面后重试'})

        apply_dir = os.path.join(self.directory, 'apply')
        os.makedirs(apply_dir, exist_ok=True)
        ts = time.strftime('%Y%m%d_%H%M%S')

        if ep == '/revert':
            job = data.get('ts') or ''
            bak_ext = 'docx' if kind == 'docx' else 'pptx'
            if not re.fullmatch(r'[\d_]+', job) or \
                    not os.path.isfile(os.path.join(apply_dir,
                                                    'backup_%s.%s' % (job, bak_ext))):
                return self.json_resp(400, {'ok': False, 'msg': '撤销任务无效或备份不存在'})
            log = open(os.path.join(apply_dir, 'log_revert_%s.txt' % job), 'w',
                       encoding='utf-8')
            subprocess.Popen([sys.executable, SCRIPTS[kind]['/revert'], self.directory, job],
                             stdout=log, stderr=subprocess.STDOUT)
            return self.json_resp(200, {'ok': True, 'job': job, 'kind': 'revert'})

        changes = data.get('changes')
        if not isinstance(changes, list) or not changes:
            return self.json_resp(400, {'ok': False, 'msg': 'changes 为空'})
        if len(changes) > 200:
            return self.json_resp(400, {'ok': False, 'msg': '单次最多 200 条改动'})
        for c in changes:
            if not isinstance(c, dict) or not isinstance(c.get('fields'), dict) \
                    or not c.get('fields'):
                return self.json_resp(400, {'ok': False, 'msg': '存在空改动条目'})
            if not isinstance(c.get('page'), int) or not isinstance(c.get('id'), int):
                return self.json_resp(400, {'ok': False, 'msg': '改动条目缺 page/id'})

        fname = 'pending_%s.json' % ts if ep == '/apply' \
            else 'pending_preview_%s.json' % ts
        data['received'] = time.strftime('%Y-%m-%d %H:%M:%S')
        with open(os.path.join(apply_dir, fname), 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=1)

        log = open(os.path.join(apply_dir, 'log_%s.txt' % ts), 'w', encoding='utf-8')
        if not SCRIPTS.get(kind):
            return self.json_resp(400, {'ok': False, 'msg':
                    'kind=' + kind + ' 的写回工具未安装（本插件仅含 PPT 审阅；'
                    'docx 需设 SLIDE_REVIEW_DOCX_TOOLKIT）'})
        subprocess.Popen([sys.executable, SCRIPTS[kind][ep], self.directory, ts],
                         stdout=log, stderr=subprocess.STDOUT)
        self.json_resp(200, {'ok': True, 'job': ts, 'kind': ep.lstrip('/')})

    # ---- GET /status?job=xxx[&type=preview|revert] ----
    def do_GET(self):
        if self.path.startswith('/status'):
            m = re.search(r'[?&]job=([\w-]+)', self.path)
            if not m:
                return self.json_resp(400, {'ok': False, 'msg': '缺 job'})
            kind = 'preview' if 'type=preview' in self.path else \
                   'revert' if 'type=revert' in self.path else ''
            name = ('status_%s_%s.json' % (kind, m.group(1))) if kind \
                else 'status_%s.json' % m.group(1)
            st = os.path.join(self.directory, 'apply', name)
            if not os.path.isfile(st):
                return self.json_resp(200, {'state': 'running', 'msg': '排队中'})
            try:
                return self.json_resp(200, json.load(open(st, encoding='utf-8')))
            except Exception:
                return self.json_resp(200, {'state': 'running', 'msg': '状态文件写入中'})
        super().do_GET()


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    directory = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.getcwd()
    handler = functools.partial(ReviewHandler, directory=directory)
    socketserver.ThreadingTCPServer.allow_reuse_address = True
    with socketserver.ThreadingTCPServer(('127.0.0.1', port), handler) as httpd:
        print('serving %s at http://127.0.0.1:%d/ (no-store + apply API)' % (
            directory, port), flush=True)
        httpd.serve_forever()


if __name__ == '__main__':
    main()
