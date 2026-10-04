# -*- coding: utf-8 -*-
"""preview_changes.py —— 预览：把改动写到 pptx 副本并渲染涉及页，不触碰原件。

由 serve_review.py 在 POST /preview 后以后台进程调用，不手动跑：
    python preview_changes.py <review_dir> <ts>

流程：读 apply/pending_preview_<ts>.json（changes 与正式应用同构）→ 复制原 pptx
为 apply/preview_src_<ts>.pptx → 复用 apply_changes 的写回原语改副本 → 只渲染涉及
页（render_all.render_pages）→ 预览图挪为 apply/preview_<ts>_slideNN.png → 前端
用它替换该页显示（角标「预览·未应用」），正式应用流程不受影响。

预览所见 = 应用所得：写回原语、字段语义与 apply_changes 完全一致。
"""
import os, sys, json, shutil, traceback

KIT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, KIT)
import apply_changes as AC                      # 复用写回原语（main 已保护，导入无副作用）

review_dir = os.path.abspath(sys.argv[1])
ts = sys.argv[2]
apply_dir = os.path.join(review_dir, 'apply')
os.makedirs(apply_dir, exist_ok=True)
STATUS = os.path.join(apply_dir, 'status_preview_%s.json' % ts)


def write_status(state, **kw):
    st = {'state': state, 'kind': 'preview', 'ts': ts, 'now': int(AC.time.time())}
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
    pending = json.load(open(os.path.join(apply_dir, 'pending_preview_%s.json' % ts),
                             encoding='utf-8'))
    meta = json.load(open(os.path.join(review_dir, 'meta.json'), encoding='utf-8'))
    changes = pending['changes']
    pptx_path = meta['pptx']
    write_status('running', msg='生成预览副本')
except Exception as e:
    fail('读取预览任务失败: %s' % e)

if not os.path.isfile(pptx_path):
    fail('找不到 pptx: %s' % pptx_path)

src_tmp = os.path.join(apply_dir, 'preview_src_%s.pptx' % ts)
try:
    shutil.copy2(pptx_path, src_tmp)
    from pptx import Presentation
    prs = Presentation(src_tmp)
    results = AC.apply_to_prs(prs, changes)
    n_ok = sum(1 for r in results if r['ok'])
    if not n_ok:
        fail('所有改动都无法写回（预览中止）：%s' %
             '；'.join('%s' % r.get('msg') for r in results if not r['ok']))
    prs.save(src_tmp)

    write_status('running', msg='渲染预览页', pages=sorted({int(c['page']) for c in changes}))
    sys.path.insert(0, AC.RENDER_WS)
    from render_all import render_pages
    tmp_dir = os.path.join(apply_dir, 'preview_%s' % ts)
    pages = sorted({int(c['page']) for c in changes})
    if not render_pages(src_tmp, tmp_dir, pages, meta['img_w'], meta['img_h']):
        raise RuntimeError('COM 渲染失败')
    pngs = []
    for png in os.listdir(tmp_dir):
        if png.endswith('.png'):
            shutil.copy2(os.path.join(tmp_dir, png),
                         os.path.join(apply_dir, 'preview_%s_%s' % (ts, png)))
            pngs.append('preview_%s_%s' % (ts, png))
    shutil.rmtree(tmp_dir, ignore_errors=True)
    write_status('ok', msg='预览已生成（未写回原件）', pages=pages, pngs=pngs,
                 results=results)
    print('OK preview %s' % pngs)
except Exception as e:
    traceback.print_exc()
    write_status('error', msg='%s' % e)
    sys.exit(1)
finally:
    if os.path.exists(src_tmp):
        try:
            os.remove(src_tmp)
        except OSError:
            pass
