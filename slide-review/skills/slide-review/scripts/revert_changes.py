# -*- coding: utf-8 -*-
"""revert_changes.py —— 撤销最近一次应用：从该次备份恢复 pptx 并重渲染涉及页。

由 serve_review.py 在 POST /revert 后以后台进程调用，不手动跑：
    python revert_changes.py <review_dir> <ts>     # ts = 被撤销的那次应用

流程：读 status_<ts>.json（取备份与改动页）→ 锁检测 → 备份覆盖回 pptx →
render_pages 只渲染该次涉及页并替换 img/ → make_review --no-render 重建审阅页 →
状态写 status_revert_<ts>.json。仅允许撤销最近一次（前端只展示最近一次入口）。
"""
import os, sys, json, shutil, traceback

KIT = os.path.dirname(os.path.abspath(__file__))
RENDER_WS = KIT                      # 插件内 render_all.py 与本脚本同目录

review_dir = os.path.abspath(sys.argv[1])
ts = sys.argv[2]
apply_dir = os.path.join(review_dir, 'apply')
STATUS = os.path.join(apply_dir, 'status_revert_%s.json' % ts)


def write_status(state, **kw):
    st = {'state': state, 'kind': 'revert', 'ts': ts}
    st.update(kw)
    tmp = STATUS + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(st, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STATUS)


def fail(msg):
    write_status('error', msg=msg)
    print('FAIL: %s' % msg)
    sys.exit(1)


try:
    meta = json.load(open(os.path.join(review_dir, 'meta.json'), encoding='utf-8'))
    applied = json.load(open(os.path.join(apply_dir, 'status_%s.json' % ts),
                             encoding='utf-8'))
    backup = applied['backup']
    pages = sorted({int(r['page']) for r in applied.get('results', []) if r.get('ok')})
    pptx_path = meta['pptx']
    img_dir = os.path.join(review_dir, 'img')
    write_status('running', msg='从备份恢复')
except Exception as e:
    fail('读取撤销任务失败: %s' % e)

if not os.path.isfile(backup):
    fail('找不到该次应用的备份: %s' % backup)

try:
    try:
        with open(pptx_path, 'r+b'):
            pass
    except PermissionError:
        fail('pptx 被占用（可能正被 WPS/PowerPoint 打开），请关闭后重试')

    shutil.copy2(backup, pptx_path)

    write_status('running', msg='重渲染涉及页 %s' % pages)
    if pages:
        sys.path.insert(0, RENDER_WS)
        from render_all import render_pages
        tmp = os.path.join(apply_dir, 'revert_%s' % ts)
        if not render_pages(pptx_path, tmp, pages, meta['img_w'], meta['img_h']):
            raise RuntimeError('COM 渲染失败（pptx 已恢复，但页面图未更新）')
        for png in os.listdir(tmp):
            if png.endswith('.png'):
                shutil.copy2(os.path.join(tmp, png), os.path.join(img_dir, png))
        shutil.rmtree(tmp, ignore_errors=True)

    write_status('running', msg='重建审阅页')
    import subprocess
    r = subprocess.run([sys.executable, os.path.join(KIT, 'make_review.py'),
                        pptx_path, '--no-render', '--frames', '--out', review_dir],
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError('make_review 失败: %s' % (r.stderr or r.stdout)[-400:])

    write_status('ok', msg='已撤销 %s 的应用并恢复页面' % ts, pages=pages)
    print('OK revert %s' % ts)
except Exception as e:
    traceback.print_exc()
    write_status('error', msg='%s' % e)
    sys.exit(1)
