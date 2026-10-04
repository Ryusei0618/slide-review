# -*- coding: utf-8 -*-
"""make_review.py —— 生成"逐页审阅页"（用户在渲染图上拖拽圈选、写意见、导出带坐标的汇总）。

用途：课件/讲义交付后，把每页渲染图铺成一个本地网页，用户可以在图上直接框住
要改的地方，框下方自动列出命中的形状（id + 名称 + 文字摘要），导出的意见带上
英寸坐标与元素 id —— 修改方据此能精确定位形状，不必靠"页面上方靠右那个框"描述。

用法:
    python make_review.py <pptx> [--out DIR] [--port 8765] [--no-render]

    --out       输出目录，默认 <pptx 同目录>/<文件名>_review
    --port      推荐端口（脚本会探测占用情况，打印最终可用的端口）
    --no-render 不重新渲染，复用 <out>/img 下已有的 PNG

生成后用本地静态服务打开（脚本不自己起服务，便于调用方管理生命周期）：
    python -m http.server <port> --directory <out>
"""
import os, re, sys, json, time, socket, zipfile, argparse
import lxml.etree as ET

KIT = os.path.dirname(os.path.abspath(__file__))          # _build_kit\脚本
RENDER_WS = KIT                      # 插件内 render_all.py 与本脚本同目录

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 914400.0
IMG_W = 1600                                              # 渲染图固定宽 1600px


# --------------------------------------------------------------------------
# 解析
# --------------------------------------------------------------------------
def slide_files(z):
    names = [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)]
    return sorted(names, key=lambda n: int(re.search(r'(\d+)', n.split('/')[-1]).group(1)))


def text_of(el, limit=44):
    t = re.sub(r'\s+', '', ''.join(el.itertext()))
    return t[:limit]


def max_sz(el):
    szs = [int(r.get('sz')) for r in el.iter('{%s}rPr' % A) if r.get('sz')]
    return max(szs) if szs else 0


def full_text_of(el):
    """sp 的完整文字（各段落以 \\n 连接）；只取自身 txBody，不含子形状/表格。"""
    tx = el.find('{%s}txBody' % P)          # p:sp 的 txBody 在 presentationml 命名空间
    if tx is None:
        return ''
    lines = [''.join(p.itertext()) for p in tx.findall('{%s}p' % A)]
    while lines and not lines[-1]:
        lines.pop()
    return '\n'.join(lines)


def sz_info(el):
    """(第一个显式字号 pt, 是否混排)。无显式字号返回 (None, False)。"""
    szs = []
    for p in el.iter('{%s}p' % A):
        for r in p.findall('{%s}r' % A):
            rPr = r.find('{%s}rPr' % A)
            if rPr is not None and rPr.get('sz'):
                szs.append(int(rPr.get('sz')) / 100.0)
    if not szs:
        return None, False
    return szs[0], len(set(szs)) > 1


NV_TAG = {'sp': 'nvSpPr', 'pic': 'nvPicPr', 'cxnSp': 'nvCxnSpPr',
          'graphicFrame': 'nvGraphicFramePr', 'grpSp': 'nvGrpSpPr'}


def walk_tree(el, tr, rows, depth, parent_id, title_cand, badge, sw_in):
    """递归登记形状（含组内子形状）。

    tr=(Sx,Sy,Cx,Cy) 是该层局部坐标→幻灯片绝对英寸的仿射映射：abs = C + S*local。
    组变换：子局部 l → 组局部 g.off + (l-chOff)*(g.ext/chExt) → 绝对 C + S*上式。
    """
    for ch in el:
        tag = ET.QName(ch).localname
        if tag not in NV_TAG:
            continue
        nv = ch.find('{%s}%s/{%s}cNvPr' % (P, NV_TAG[tag], P))
        if nv is None:
            continue
        xf = (ch.find('{%s}xfrm' % P) if tag == 'graphicFrame'
              else ch.find('.//{%s}xfrm' % A))
        if xf is None:
            continue
        off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
        if off is None or ext is None:
            continue
        Sx, Sy, Cx, Cy = tr
        gx, gy = int(off.get('x')) / EMU, int(off.get('y')) / EMU
        gw, gh = int(ext.get('cx')) / EMU, int(ext.get('cy')) / EMU
        x, y = Cx + Sx * gx, Cy + Sy * gy
        w, h = Sx * gw, Sy * gh
        txt = text_of(ch)
        row = dict(id=int(nv.get('id')), name=nv.get('name') or '',
                   x=x, y=y, w=w, h=h, txt=txt, tag=tag, parent=parent_id,
                   full='', sz=None, mixed=False)
        if tag == 'sp':
            row['full'] = full_text_of(ch)
            row['sz'], row['mixed'] = sz_info(ch)
        rows.append(row)
        # 标题 / 类型徽章候选只在顶层形状里选，避免组内小字干扰
        if depth == 0 and txt and y + h / 2 < 1.7:
            title_cand.append((max_sz(ch), y, txt))
        if depth == 0 and txt and len(txt) <= 5 and 0.5 < y < 1.4 and x > sw_in * 0.7:
            badge[0] = txt
        if tag == 'grpSp':
            chOff, chExt = xf.find('{%s}chOff' % A), xf.find('{%s}chExt' % A)
            if chOff is None or chExt is None:
                continue
            cox, coy = int(chOff.get('x')) / EMU, int(chOff.get('y')) / EMU
            cex, cey = int(chExt.get('cx')) / EMU, int(chExt.get('cy')) / EMU
            if not (cex and cey and gw and gh):
                continue
            s2x, s2y = Sx * gw / cex, Sy * gh / cey
            c2x, c2y = Cx + Sx * gx - s2x * cox, Cy + Sy * gy - s2y * coy
            walk_tree(ch, (s2x, s2y, c2x, c2y), rows, depth + 1,
                      row['id'], title_cand, badge, sw_in)


def parse_shapes_and_groups(pptx, sw_in):
    """返回 (pages, groups, titles)。坐标单位为英寸，供前端换算成像素。

    递归展开组：组内子形状也登记，parent 记父组 id（顶层为 0）。
    页标题取靠上区域字号最大的文本，并附上同带靠右的短标签（类型徽章）以便区分同节多页。
    """
    z = zipfile.ZipFile(pptx)
    pages, groups, titles = {}, {}, {}
    for fn in slide_files(z):
        idx = int(re.search(r'slide(\d+)', fn).group(1))
        root = ET.fromstring(z.read(fn))
        spTree = root.find('.//{%s}cSld/{%s}spTree' % (P, P))
        rows, title_cand, badge = [], [], ['']
        if spTree is not None:
            walk_tree(spTree, (1.0, 1.0, 0.0, 0.0), rows, 0, 0,
                      title_cand, badge, sw_in)
        pages[idx] = rows
        t = max(title_cand)[2] if title_cand else ''
        if t and badge[0] and badge[0] not in t:
            t += '（%s）' % badge[0]
        titles[idx] = t

        # ---- 击序：从 timing 解析每组点击的成员，用形状文字/尺寸描述 ----
        byid = {r['id']: r for r in rows}
        desc = []
        tim = root.find('{%s}timing' % P)
        if tim is not None:
            ms = tim.find('.//{%s}cTn[@nodeType="mainSeq"]' % P)
            if ms is not None:
                seen = set()
                for gi, gpar in enumerate(ms.find('{%s}childTnLst' % P).findall('{%s}par' % P), 1):
                    ids = []
                    for epar in gpar.iter('{%s}par' % P):
                        ctn = epar.find('{%s}cTn' % P)
                        if ctn is None or ctn.get('nodeType') not in (
                                'clickEffect', 'withEffect', 'afterEffect'):
                            continue
                        tgt = ctn.find('.//{%s}spTgt' % P)
                        sid = int(tgt.get('spid')) if tgt is not None and tgt.get('spid') else None
                        if sid and sid not in seen:
                            seen.add(sid)
                            if sid in byid:
                                ids.append(sid)
                    if ids:
                        labels = []
                        for sid in ids:
                            r = byid[sid]
                            labels.append(r['txt'][:10] if r['txt']
                                          else '%.1f×%.1f' % (r['w'], r['h']))
                        cnt = {}
                        for lb in labels:
                            cnt[lb] = cnt.get(lb, 0) + 1
                        head = '、'.join('%s×%d' % (k, v) if v > 1 else k
                                         for k, v in list(cnt.items())[:5])
                        desc.append('%d 元素（%s）' % (len(ids), head))
        # 首组若是切页自动入场（auto 页），标注出来
        if desc:
            auto = tim is not None and tim.find('.//{%s}cond[@evt="onBegin"]' % P) is not None
            tag = '自动入场 ' if auto else ''
            groups[idx] = '点击击序 ' + tag + ' → '.join(desc)
        else:
            groups[idx] = '点击击序 无动画'
    z.close()
    return pages, groups, titles


def pick_port(prefer):
    for p in range(prefer, prefer + 40):
        with socket.socket() as s:
            try:
                s.bind(('127.0.0.1', p))
                return p
            except OSError:
                continue
    return prefer


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('--out', default=None)
    ap.add_argument('--port', type=int, default=8765)
    ap.add_argument('--no-render', action='store_true')
    ap.add_argument('--frames', action='store_true',
                    help='为有动画的页逐帧渲染（复用 make_frames 的 build_plan/render_frames），'
                         '帧图落 <out>/frames/，帧清单进 meta 供审阅页逐帧播放器使用')
    args = ap.parse_args()

    pptx = os.path.abspath(args.pptx)
    stem = os.path.splitext(os.path.basename(pptx))[0]
    out = args.out or os.path.join(os.path.dirname(pptx), stem + '_review')
    img_dir = os.path.join(out, 'img')
    os.makedirs(img_dir, exist_ok=True)

    # ---- 渲染 ----
    if not args.no_render:
        sys.path.insert(0, RENDER_WS)
        from render_all import render
        render(pptx, img_dir, IMG_W, int(IMG_W * 9 / 16))
    pngs = sorted(f for f in os.listdir(img_dir) if f.endswith('.png'))
    if not pngs:
        sys.exit('没有渲染图：%s' % img_dir)
    pages_n = len(pngs)

    from pptx import Presentation
    prs0 = Presentation(pptx)
    pages, groups, titles = parse_shapes_and_groups(pptx, prs0.slide_width / EMU)
    from PIL import Image
    with Image.open(os.path.join(img_dir, pngs[0])) as im:
        img_w, img_h = im.size

    # 形状 → 像素坐标（按幻灯片真实英寸尺寸换算，支持任意画幅）
    from pptx import Presentation
    prs = Presentation(pptx)
    shapes = {}
    sw_in = prs.slide_width / EMU
    sh_in = prs.slide_height / EMU
    kx, ky = img_w / sw_in, img_h / sh_in
    for idx, rows in pages.items():
        if idx > pages_n:
            continue
        shapes[idx] = [[r['id'], round(r['x'] * kx, 1), round(r['y'] * ky, 1),
                        round(r['w'] * kx, 1), round(r['h'] * ky, 1),
                        r['name'], r['txt'], r['tag'],
                        r['full'], r['sz'], r['mixed'], r['parent']] for r in rows]

    # ---- 帧准备（--frames）：有动画的页逐帧渲染供审阅页播放器；已有帧图且页数吻合则复用 ----
    frames_meta = {}
    if args.frames:
        sys.path.insert(0, KIT)
        import make_frames as MF
        fdir = os.path.join(out, 'frames')
        os.makedirs(fdir, exist_ok=True)
        plan, notes = MF.build_plan(pptx)
        if plan:
            done = True          # 已有帧图且每页帧数吻合则跳过重渲染（apply 重建走 --no-render 不重跑）
            for sp in plan:
                n_img = len([f for f in os.listdir(fdir)
                             if re.match(r'slide%02d_' % sp['no'], f) and f.endswith('.png')])
                if n_img != len(sp['frames']):
                    done = False
                    break
            if not done:
                for f in os.listdir(fdir):
                    if f.endswith('.png'):
                        os.remove(os.path.join(fdir, f))
                ok = MF.render_frames(pptx, fdir, plan)
                if not ok:
                    print('帧渲染失败（页面本体不受影响，帧播放器缺图）')
            for sp in plan:
                frames_meta[sp['no']] = [{'id': fr['id'], 'label': fr['label']}
                                         for fr in sp['frames']]

    data = {
        'v': int(time.time()),              # 图片 URL 版本号，避免浏览器拿旧渲染图
        'title': stem,
        'pptx': pptx,                       # 成品绝对路径（应用改动时回传校验）
        'px': kx, 'py': ky,                 # 每英寸像素数（x / y 方向）
        'sw': sw_in, 'sh': sh_in,           # 幻灯片英寸尺寸（导出坐标用）
        'digits': 2,
        'pages': [{'no': i, 'title': titles.get(i) or ('第 %d 页' % i),
                   'anim': groups.get(i, '无动画')} for i in range(1, pages_n + 1)],
        'shapes': shapes,
        'frames': frames_meta,              # {页码: [{id,label}...]}，空 = 该页无帧（整页常显或未跑 --frames）
        'key': re.sub(r'\W+', '_', stem)[:40],
    }

    tpl_path = os.path.join(KIT, 'review_template.html')
    tpl = open(tpl_path, encoding='utf-8').read()
    html = tpl.replace('/*__DATA__*/', json.dumps(data, ensure_ascii=False,
                                                  separators=(',', ':')))
    dst = os.path.join(out, 'index.html')
    open(dst, 'w', encoding='utf-8').write(html)

    # meta.json：供 serve_review.py（POST /apply 校验）与 apply_changes.py（写回+渲染尺寸）使用
    meta = {'pptx': pptx, 'key': data['key'], 'title': stem,
            'img_w': img_w, 'img_h': img_h, 'pages': pages_n,
            'generated': int(time.time())}
    open(os.path.join(out, 'meta.json'), 'w', encoding='utf-8').write(
        json.dumps(meta, ensure_ascii=False, indent=1))

    port = pick_port(args.port)
    print('审阅页: %s' % dst)
    print('形状数: %d 页 / %d 个' % (pages_n, sum(len(v) for v in shapes.values())))
    print('服务命令: python serve_review.py %d "%s"' % (port, out))
    print('访问地址: http://127.0.0.1:%d/' % port)


if __name__ == '__main__':
    main()
