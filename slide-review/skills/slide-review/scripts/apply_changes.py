# -*- coding: utf-8 -*-
"""apply_changes.py —— 执行审阅页「应用改动」清单：按形状 id 写回 pptx 并重渲染。

由 serve_review.py 在 POST /apply 后以后台进程调用，不手动跑：
    python apply_changes.py <review_dir> <ts>

流程：读 apply/pending_<ts>.json + meta.json → 备份 pptx → python-pptx 逐条写回
（text / size / bold / color / x / y / w / h）→ COM 全册重渲染到临时目录 →
非改动页与 img/ 旧图逐像素 diff（应零差异，违反即回滚 pptx 并报错）→ 改动页
新图替换 img/ → 重跑 make_review --no-render 重建 index.html → 全程把进度写进
apply/status_<ts>.json（serve_review.py 的 /status 接口读它返回给前端轮询）。

写回原语（find_shape / set_text / set_style / set_geom / apply_to_prs）设计为可
导入：preview_changes.py 复用同一套写回逻辑生成预览，保证「预览所见 = 应用所得」。
安全纪律：写回前检测 pptx 未被占用；文字替换统一为首段格式时在 result 里注明；
渲染 diff 异常一律从备份回滚，不污染交付文件。
"""
import os, sys, re, json, time, shutil, traceback
import lxml.etree as etree
from pptx.oxml.ns import qn

KIT = os.path.dirname(os.path.abspath(__file__))
RENDER_WS = KIT                      # 插件内 render_all.py 与本脚本同目录
EMU = 914400.0
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
AP = '{%s}' % A

# 运行期状态（main() 赋值；write_status/fail 只在 main 流程里调用）
T0 = 0
_ts = ''
STATUS = ''


def write_status(state, **kw):
    st = {'state': state, 'ts': _ts, 't0': T0, 'now': int(time.time())}
    st.update(kw)
    tmp = STATUS + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(st, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STATUS)


def fail(msg, results=None, backup=None):
    write_status('error', msg=msg, results=results or [], backup=backup)
    print('FAIL: %s' % msg)
    sys.exit(1)


# ---------------------------------------------------------------------------
# 写回原语（可导入复用）
# ---------------------------------------------------------------------------
def find_shape(slide, sid):
    """递归查找（含组内子形状），返回 (shape, 组链)。

    组链从最外层组到直接父组，用于几何改动的坐标系换算；顶层形状组链为空。
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    chain = []

    def walk(shapes):
        for sp in shapes:
            if sp.shape_id == sid:
                return sp
            if sp.shape_type == MSO_SHAPE_TYPE.GROUP:
                chain.append(sp)
                r = walk(sp.shapes)
                if r is not None:
                    return r
                chain.pop()
        return None

    sp = walk(slide.shapes)
    return (sp, list(chain)) if sp is not None else (None, [])


def _grp_params(g):
    """组的 (off.x, off.y, ext.cx, ext.cy, chOff.x, chOff.y, chExt.cx, chExt.cy)，EMU。"""
    pr = g._element.find(qn('p:grpSpPr'))
    xf = None if pr is None else pr.find(qn('a:xfrm'))
    if xf is None:
        return None
    off, ext = xf.find(qn('a:off')), xf.find(qn('a:ext'))
    cho, che = xf.find(qn('a:chOff')), xf.find(qn('a:chExt'))
    if None in (off, ext, cho, che):
        return None
    return (int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy')),
            int(cho.get('x')), int(cho.get('y')), int(che.get('cx')), int(che.get('cy')))


def _abs_geom(sp, chain):
    """sp 当前 off/ext（直接父组坐标系，EMU）→ 幻灯片绝对 EMU。"""
    x, y, w, h = sp.left, sp.top, sp.width, sp.height
    for g in reversed(chain):                    # 直接父组 → … → 最外层组 → 绝对
        p = _grp_params(g)
        if p is None:
            continue
        kx = p[2] / p[6] if p[6] else 1.0        # ext / chExt
        ky = p[3] / p[7] if p[7] else 1.0
        x = p[0] + (x - p[4]) * kx
        y = p[1] + (y - p[5]) * ky
        w, h = w * kx, h * ky
    return x, y, w, h


def _set_geom_abs(sp, chain, tx, ty, tw, th):
    """目标绝对 EMU → 逐层逆变换到直接父组坐标系后写入。"""
    from pptx.util import Emu
    x, y, w, h = tx, ty, tw, th
    for g in chain:                              # 绝对 → 最外层组子坐标 → … → 直接父组坐标
        p = _grp_params(g)
        if p is None:
            continue
        kx = p[6] / p[2] if p[2] else 1.0        # chExt / ext
        ky = p[7] / p[3] if p[3] else 1.0
        x = p[4] + (x - p[0]) * kx
        y = p[5] + (y - p[1]) * ky
        w, h = w * kx, h * ky
    sp.left = Emu(int(round(x)))
    sp.top = Emu(int(round(y)))
    sp.width = Emu(int(round(w)))
    sp.height = Emu(int(round(h)))


def _toxml(el):
    return etree.tostring(el) if el is not None else b''


def _deepcopy(el):
    from copy import deepcopy
    return deepcopy(el)


def set_text(sp, text, res):
    """整框替换文字：\\n 拆段；保留每段首个 run 的格式；原多格式时统一为首格式。"""
    if not getattr(sp, 'has_text_frame', False):
        res['ok'] = False
        res['msg'] = '该形状不含文本（图片/表格等），不能改文字'
        return
    txB = sp.text_frame._txBody
    paras = txB.findall(AP + 'p')
    # 混排检测：所有 run 的 rPr 是否一致
    rprs = set()
    for p in paras:
        for r in p.findall(AP + 'r'):
            rPr = r.find(AP + 'rPr')
            rprs.add(re.sub(rb'\s+', b'', _toxml(rPr)) if rPr is not None else b'')
    if len(rprs) > 1:
        res['msg'] = '原文本含 %d 种格式，已统一为第一种' % len(rprs)
    tmpl = None                                   # 首个带格式的 run 作样式模板
    for p in paras:
        rs = p.findall(AP + 'r')
        if rs:
            tmpl = rs[0]
            break
    lines = text.split('\n')
    while len(paras) > len(lines):                # 删多余段
        txB.remove(paras.pop())
    while len(paras) < len(lines):                # 补段（复制末段骨架，清空 run）
        newp = _deepcopy(paras[-1])
        for e in list(newp):
            if etree.QName(e).localname in ('r', 'br', 'fld'):
                newp.remove(e)
        txB.append(newp)
        paras.append(newp)
    for p, line in zip(paras, lines):
        rs = p.findall(AP + 'r')
        if rs:
            keep = rs[0]
            for r in rs[1:]:
                p.remove(r)
        elif tmpl is not None:
            keep = _deepcopy(tmpl)
            p.append(keep)
        else:
            keep = etree.SubElement(p, AP + 'r')
        for e in list(keep):
            if etree.QName(e).localname == 't':
                keep.remove(e)
        etree.SubElement(keep, AP + 't').text = line


def set_style(sp, fields, res):
    if not getattr(sp, 'has_text_frame', False):
        res['ok'] = False
        res['msg'] = '该形状不含文本，不能改样式'
        return
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    n = 0
    for p in sp.text_frame.paragraphs:
        for r in p.runs:
            if 'size' in fields:
                r.font.size = Pt(float(fields['size']))
            if 'bold' in fields:
                r.font.bold = bool(fields['bold'])
            if 'color' in fields:
                r.font.color.rgb = RGBColor.from_string(str(fields['color']))
            n += 1
    if not n:
        res['msg'] = '该文本框没有可设置格式的文字 run，样式未生效'


def set_geom(sp, fields, res, chain):
    """按绝对英寸坐标改几何；组内子形状经 _abs_geom/_set_geom_abs 换算。"""
    from pptx.util import Emu
    cur = dict(zip(('x', 'y', 'w', 'h'), (v / EMU for v in _abs_geom(sp, chain))))
    tgt = dict(cur)
    for k in ('x', 'y', 'w', 'h'):
        if k in fields:
            v = float(fields[k])
            if abs(v - cur[k]) > 0.004:
                tgt[k] = v
    if tgt == cur:
        return
    _set_geom_abs(sp, chain,
                  *(int(round(tgt[k] * EMU)) for k in ('x', 'y', 'w', 'h')))


def _remove_anim_refs(sld, sid):
    """删形状前清掉 timing 里对它的引用：spTgt(spid) 所在效果 par 整体移除，
    避免悬空 spid 让放映异常。同一 par 多目标时整组删除（粗粒度可接受）。"""
    root = sld._element
    pars = set()
    for tgt in root.iter(qn('p:spTgt')):
        if tgt.get('spid') == str(sid):
            par = tgt.getparent()                     # 效果 par（内含带 nodeType 的 cTn）
            while par is not None and par.getparent() is not None \
                    and par.getparent().tag != qn('p:childTnLst'):
                par = par.getparent()                 # 向上到击组 par（直接挂在 childTnLst 下）
            if par is not None and par.getparent() is not None \
                    and par.getparent().tag == qn('p:childTnLst'):
                pars.add(par)
    for par in pars:
        par.getparent().remove(par)


def _replace_image(slide, sp, data_url, res):
    """图片替换：新图加为独立 image part，blip 的 r:embed 指向它（旧 part 留存不破坏共享）。
    data_url = data:image/png;base64,... （前端统一 canvas 重采样输出）。"""
    import base64
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    if sp.shape_type != MSO_SHAPE_TYPE.PICTURE:
        res['ok'] = False
        res['msg'] = '该形状不是图片（%s），不能替换' % sp.shape_type
        return
    head, _, b64 = data_url.partition(',')
    ext = 'png' if 'image/png' in head else 'jpg' if 'image/jpeg' in head else None
    if ext is None:
        res['ok'] = False
        res['msg'] = '只支持 PNG/JPEG 图片'
        return
    data = base64.b64decode(b64)
    import io, tempfile
    blip = sp._element.blipFill.find(qn('a:blip'))
    if blip is None or not blip.get(qn('r:embed')):
        res['ok'] = False
        res['msg'] = '该图片没有可替换的媒体引用'
        return
    tfd, tpath = tempfile.mkstemp(suffix='.' + ext)
    try:
        with os.fdopen(tfd, 'wb') as f:
            f.write(data)
        image_part, new_rid = slide.part.get_or_add_image_part(tpath)
    finally:
        os.remove(tpath)
    blip.set(qn('r:embed'), new_rid)
    res['msg'] = res.get('msg', '') + '已替换为新图（%.0f KB）' % (len(data) / 1024)


def apply_to_prs(prs, changes):
    """把改动逐条写回 Presentation 对象，返回 results（顺序与 changes 一致）。"""
    results = []
    slides = list(prs.slides)
    for c in changes:
        page, sid = int(c['page']), int(c['id'])
        fields = c.get('fields') or {}
        res = {'page': page, 'id': sid, 'op': c.get('op', ''),
               'cid': c.get('cid'), 'ok': True}
        try:
            if page < 1 or page > len(slides):
                raise ValueError('页码超出范围（共 %d 页）' % len(slides))
            sp, chain = find_shape(slides[page - 1], sid)
            if sp is None:
                raise ValueError('该页形状中找不到 id=%d（可能被删或重组过）' % sid)
            if fields.get('del'):
                _remove_anim_refs(slides[page - 1], sid)
                el = sp._element
                el.getparent().remove(el)
                res['msg'] = '已删除形状（含其动画引用）'
                results.append(res)
                continue
            if 'img' in fields:
                _replace_image(slides[page - 1], sp, str(fields['img']), res)
            if 'text' in fields:
                set_text(sp, str(fields['text']), res)
            if res['ok'] and ({'size', 'bold', 'color'} & fields.keys()):
                set_style(sp, fields, res)
            if res['ok'] and ({'x', 'y', 'w', 'h'} & fields.keys()):
                set_geom(sp, fields, res, chain)
        except Exception as e:
            res['ok'] = False
            res['msg'] = str(e)
        results.append(res)
    return results


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main():
    global T0, _ts, STATUS
    review_dir = os.path.abspath(sys.argv[1])
    _ts = sys.argv[2]
    ts = _ts
    apply_dir = os.path.join(review_dir, 'apply')
    os.makedirs(apply_dir, exist_ok=True)
    STATUS = os.path.join(apply_dir, 'status_%s.json' % ts)
    T0 = int(time.time())
    try:
        pending = json.load(open(os.path.join(apply_dir, 'pending_%s.json' % ts),
                                 encoding='utf-8'))
        meta = json.load(open(os.path.join(review_dir, 'meta.json'), encoding='utf-8'))
        changes = pending['changes']
        pptx_path = meta['pptx']
        img_dir = os.path.join(review_dir, 'img')
        write_status('running', total=len(changes), msg='开始写回')
    except Exception as e:
        fail('读取任务失败: %s' % e)

    if not os.path.isfile(pptx_path):
        fail('找不到 pptx: %s' % pptx_path)

    # ---- 备份 ----
    backup = os.path.join(apply_dir, 'backup_%s.pptx' % ts)
    shutil.copy2(pptx_path, backup)

    # ---- 文件占用检测（WPS/PowerPoint 开着会锁文件，写回必失败）----
    try:
        with open(pptx_path, 'r+b'):
            pass
    except PermissionError:
        fail('pptx 被占用（可能正被 WPS/PowerPoint 打开），请关闭后重新应用。备份: %s' % backup,
             backup=backup)

    results = []
    try:
        write_status('running', msg='写回形状', total=len(changes))
        from pptx import Presentation
        prs = Presentation(pptx_path)
        results = apply_to_prs(prs, changes)
        for i, res in enumerate(results, 1):
            write_status('running', msg='写回形状 %d/%d' % (i, len(changes)),
                         total=len(changes), results=results)

        bad = [r for r in results if not r['ok']]
        if len(bad) == len(results) and results:
            fail('所有改动都失败，pptx 未保存', results=results, backup=backup)
        write_status('running', msg='保存 pptx', total=len(changes), results=results)
        prs.save(pptx_path)

        # ---- 渲染 diff：非改动页必须与旧图零差异，违反即回滚 ----
        write_status('running', msg='COM 渲染中（约几十秒）', total=len(changes),
                     results=results)
        sys.path.insert(0, RENDER_WS)
        from render_all import render
        from PIL import Image, ImageChops
        tmp = os.path.join(apply_dir, 'render_%s' % ts)
        if not render(pptx_path, tmp, meta['img_w'], meta['img_h']):
            raise RuntimeError('COM 渲染失败，pptx 已写回但未通过校验')

        changed_pages = {int(c['page']) for c in changes}
        cmp_tmp = os.path.join(tmp, '_cmp')      # 前后对比留存（改动页 before/after 渲染图）
        os.makedirs(cmp_tmp, exist_ok=True)
        unexpected = []
        for png in sorted(f for f in os.listdir(tmp) if f.endswith('.png')):
            page = int(re.search(r'(\d+)', png).group(1))
            src, dst = os.path.join(tmp, png), os.path.join(img_dir, png)
            if page in changed_pages or not os.path.exists(dst):
                if page in changed_pages:
                    if os.path.exists(dst):      # 应用前旧图（img/ 即将被覆盖，先留底）
                        shutil.copy2(dst, os.path.join(cmp_tmp, 'before_' + png))
                    shutil.copy2(src, dst)
                    shutil.copy2(src, os.path.join(cmp_tmp, 'after_' + png))
                    continue
                shutil.copy2(src, dst)
                continue
            with Image.open(dst) as a, Image.open(src) as b:
                if a.size != b.size or ImageChops.difference(
                        a.convert('RGB'), b.convert('RGB')).getbbox() is not None:
                    unexpected.append(page)
        if unexpected:
            shutil.copy2(backup, pptx_path)          # 回滚
            shutil.rmtree(tmp, ignore_errors=True)   # _cmp 在 tmp 内，回滚后对比留存一并作废
            raise RuntimeError('第 %s 页出现意外变化（只允许改动页变化），已从备份回滚 pptx'
                               % '、'.join(map(str, sorted(set(unexpected)))))
        cmp_dir = None
        if os.listdir(cmp_tmp):
            cmp_dir = os.path.join(apply_dir, 'compare_%s' % ts)
            shutil.move(cmp_tmp, cmp_dir)
        shutil.rmtree(tmp, ignore_errors=True)

        # ---- 重建 index.html（新 v；--no-render 复用刚替换的 img；--frames 帧图已在则自动复用，
        #      漏带会把 meta 的帧数据冲掉、播放器整块消失——实测踩过）----
        write_status('running', msg='重建审阅页', total=len(changes), results=results)
        import subprocess
        r = subprocess.run([sys.executable, os.path.join(KIT, 'make_review.py'),
                            pptx_path, '--no-render', '--frames', '--out', review_dir],
                           capture_output=True, text=True)
        if r.returncode != 0:
            raise RuntimeError('make_review 失败: %s' % (r.stderr or r.stdout)[-400:])

        n_ok = sum(1 for r in results if r['ok'])
        write_status('ok', msg='已应用 %d/%d 条改动' % (n_ok, len(results)),
                     total=len(changes), results=results, backup=backup,
                     compare=('compare_%s' % ts) if cmp_dir else None,
                     cmp_pages=sorted(changed_pages) if cmp_dir else [])
        print('OK %d/%d' % (n_ok, len(results)))
    except Exception as e:
        traceback.print_exc()
        fail('%s' % e, results=results, backup=backup)


if __name__ == '__main__':
    main()
