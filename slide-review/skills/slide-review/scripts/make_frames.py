# -*- coding: utf-8 -*-
"""逐帧导出一页 PPT 的放映过程，供 AI 像人一样逐步通读演示。

粒度：一击一帧；击内 afterEffect 自动接续链（点一下后连续变化）自动展开子帧
（fK_sJ），纯 click/with 错落击维持一帧。原理：入场类动画不改变已呈现元素的
布局，"已呈现元素可见、未入场元素隐藏"是精确模拟；exit 动画在其所在帧隐藏；
强调/路径动画不改可见性，帧图即终态。

用法:
  python make_frames.py <pptx> [页码 ...]   # 页码=放映页序(1-based)，缺省全册
输出（<pptx同目录>/<stem>_frames/ 下）:
  slideNN_<帧id>.png    帧0=切页即见(auto组已现)，帧k=第k击后，fK_sJ=第k击自动接续子帧
  frames_manifest.md    每帧新出现的形状清单（AI 看帧图时对照）
  plan_full.json        解析明细（spid/形状名/帧序），留查
"""
import json, os, re, shutil, subprocess, sys, time, zipfile
from lxml import etree
from pptx import Presentation

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
def q(t): return "{%s}%s" % (P, t)

W, H = 1600, 900
NODETYPES = ('clickEffect', 'withEffect', 'afterEffect')

def preset_of(tgt):
    """spTgt 向上找最近带 presetClass 的祖先 cTn（行为级 cTn 不带 presetClass）。"""
    for anc in tgt.iterancestors():
        if anc.tag == q('cTn') and anc.get('presetClass'):
            return anc.get('presetClass')
    return None

def shape_info(sld):
    """顶层形状 spid -> (名称, 文字摘要)，供 manifest 对照。"""
    info = {}
    spTree = sld.find('.//' + q('cSld') + '/' + q('spTree'))
    if spTree is None: return info
    for ch in spTree:
        tag = etree.QName(ch).localname
        if tag in ('nvGrpSpPr', 'grpSpPr'): continue
        c = ch.find('.//' + q('cNvPr'))          # 文档序第一个即顶层 id（组合同理）
        if c is None: continue
        try: sid = int(c.get('id'))
        except (TypeError, ValueError): continue
        parts = []                               # 公式走 mc:Fallback 双份，剔掉 Fallback 分支
        def walk(el):
            for child in el:
                if child.tag == '{%s}Fallback' % MC: continue
                if child.tag in ('{%s}t' % A, '{%s}t' % M): parts.append(child.text or '')
                walk(child)
        walk(ch)
        info[sid] = (c.get('name') or tag, re.sub(r'\s+', '', ''.join(parts))[:22])
    return info

def click_effects(par):
    """击组内按文档序的效果序列 [(nodeType, [(spid, presetClass), ...]), ...]。
    效果 = 带 nodeType 的 cTn，其父 par 子树内的全部 spTgt（包装层 cTn 无 nodeType 不计）。"""
    effects = []
    for c in par.iter(q('cTn')):
        nt = c.get('nodeType')
        if nt not in NODETYPES: continue
        effs = []
        for t in c.getparent().iter(q('spTgt')):   # cTn 的父是效果 par
            try: sid = int(t.get('spid'))
            except (TypeError, ValueError): continue
            effs.append((sid, preset_of(t)))
        effects.append((nt, effs))
    return effects

def parse_timing(sld):
    """解析主序列击组 -> 有序帧序列。返回 dict(frames, clicks) 或 None（整页常显）。
    每帧 {id, label, show: [(spid,pc)], hide: [(spid,pc)]}；击内子帧切分规则：
    clickEffect/afterEffect 开新子帧，withEffect 并入当前子帧——纯 click/with 击
    切分后仅 1 子帧、维持 fK；含 afterEffect 链的击展开为 fK_s1..fK_sJ。"""
    timing = sld.find(q('timing'))
    if timing is None: return None
    main = None
    for seq in timing.iter(q('seq')):
        ctn = seq.find(q('cTn'))
        if ctn is not None and ctn.get('nodeType') == 'mainSeq':
            main = seq; break
    if main is None: return None
    lst = main.find(q('cTn') + '/' + q('childTnLst'))   # 击组在 seq/cTn/childTnLst 下
    if lst is None: return None
    pars = lst.findall(q('par'))
    if not pars: return None
    frames, ignored, skipped = [], {}, []
    seen_show, seen_hide = set(), set()                  # 跨击去重：同一形状取首次出现帧
    n_click = 0
    frames.append({'id': 'f0', 'label': '帧0（切页即见）', 'show': [], 'hide': []})   # 恒存在：初始态参照帧
    for par in pars:
        ctn = par.find(q('cTn'))
        conds = ctn.find(q('stCondLst')) if ctn is not None else None
        auto = False                            # onBegin 事件或数字 delay = 自动组 -> 帧0
        if conds is not None:
            for c in conds.findall(q('cond')):
                if c.get('evt') == 'onBegin' or (c.get('delay') not in (None, 'indefinite')):
                    auto = True
        if auto:
            base_id, base_label = 'f0', '帧0（切页即见）'
        else:
            n_click += 1
            base_id, base_label = 'f%d' % n_click, '第%d击后' % n_click
        effects = click_effects(par)
        if not effects:
            if not auto: skipped.append(n_click)   # 无效果节点的击（残缺 timing）
            continue
        # 子帧切分：with 并入当前，click/after 开新
        subs = []
        for nt, effs in effects:
            if nt == 'withEffect' and subs:
                subs[-1].append((nt, effs))
            else:
                subs.append([(nt, effs)])
        # 分流到 show/hide 后剔除无可见性变化的空子帧（纯强调/路径或重复登记的效果）
        sub_frames = []
        for sub in subs:
            show, hide = [], []
            for nt, effs in sub:
                for sid, pc in effs:
                    if pc == 'entr':
                        if sid not in seen_show: seen_show.add(sid); show.append((sid, pc))
                    elif pc == 'exit':
                        if sid not in seen_hide: seen_hide.add(sid); hide.append((sid, pc))
                    else:
                        ignored.setdefault(sid, pc)
            if show or hide: sub_frames.append({'show': show, 'hide': hide})
        if not sub_frames:
            if not auto: skipped.append(n_click)   # 仅强调/路径动画的击：无画面变化
            continue
        if auto:                                    # 自动组效果并入帧0（本机约定 auto ≤ 1 组）
            for sf in sub_frames:
                frames[0]['show'].extend(sf['show']); frames[0]['hide'].extend(sf['hide'])
            continue
        expand = len(sub_frames) > 1
        for j, sf in enumerate(sub_frames, 1):
            fid = '%s_s%d' % (base_id, j) if expand else base_id
            label = (base_label if j == 1 else '%s·自动接续%d' % (base_label, j - 1)) if expand else base_label
            frames.append({'id': fid, 'label': label, **sf})
    if not any(f['show'] or f['hide'] for f in frames):
        return None
    return {'frames': frames, 'clicks': n_click, 'ignored': ignored, 'skipped': skipped}

def _zip_slide_parts(z):
    """坏包（media CRC 损坏等 python-pptx 打不开的）兜底：按 presentation.xml 手工映射放映页序。"""
    R = "http://schemas.openxmlformats.org/package/2006/relationships"
    R0 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    rid2t = {r.get('Id'): r.get('Target') for r in rels.iter('{%s}Relationship' % R)}
    out = []
    for sld in etree.fromstring(z.read('ppt/presentation.xml')).iter(q('sldId')):
        t = rid2t.get(sld.get('{%s}id' % R0))
        if not t: continue
        t = t.lstrip('/')
        out.append(t if t.startswith('ppt/') else 'ppt/' + t.lstrip('./'))
    return out

def build_plan(src, pages=None):
    try:
        prs = Presentation(src)
        pages_el = [(i, s._element) for i, s in enumerate(prs.slides, 1)]
    except Exception:                         # 原稿常见坏 media（CRC 损坏）——zip 直读 slide XML 绕过
        z = zipfile.ZipFile(src)
        pages_el = [(i, etree.fromstring(z.read(part)))
                    for i, part in enumerate(_zip_slide_parts(z), 1)]
    plan, notes = [], []
    for i, sld in pages_el:
        if pages and i not in pages: continue
        t = parse_timing(sld)
        if not t:
            notes.append((i, '整页常显（无主序列击组），不产出帧图'))
            continue
        plan.append({'no': i, 'info': shape_info(sld), **t})
    return plan, notes

def _names(frame, info):
    out = []
    for sid, pc in frame['show']:
        nm, txt = info.get(sid, ('spid%s' % sid, ''))
        out.append('%s%s' % (nm, ('「%s」' % txt) if txt else ''))
    for sid, pc in frame['hide']:
        nm, txt = info.get(sid, ('spid%s' % sid, ''))
        out.append('〔淡出〕%s%s' % (nm, ('「%s」' % txt) if txt else ''))
    return out

def manifest_md(src, plan, notes):
    L = ['# 逐帧清单：%s' % os.path.basename(src), '']
    for n, msg in notes:
        L += ['## 第 %d 页' % n, '- %s' % msg, '']
    for sp in plan:
        L.append('## 第 %d 页（%d 击 + 自动组，共 %d 帧）' % (sp['no'], sp['clicks'], len(sp['frames'])))
        for fr in sp['frames']:
            got = _names(fr, sp['info'])
            if not got and fr['id'] == 'f0':
                got = ['（仅页面常显元素，无自动入场组）']
            L.append('- **%s**: %s' % (fr['label'], '；'.join(got) if got else '（无变化）'))
        for n in sp.get('skipped', []):
            L.append('- 注: 第 %d 击仅强调/路径动画，无画面变化，不产出帧图' % n)
        for sid, pc in sorted(sp['ignored'].items()):
            nm, txt = sp['info'].get(sid, ('spid%s' % sid, ''))
            if pc == 'path':
                L.append('- 注: %s（%s）仅路径动画，帧图显示其编辑态位置，非播放终态位置' % (nm, sid))
            else:
                L.append('- 注: %s（%s）仅 %s 动画，不改可见性，帧图即终态' % (nm, sid, pc or '未分类'))
        L.append('')
    return '\n'.join(L)

def render_frames(src, outdir, plan):
    """复制副本 -> COM 逐帧控制可见性导出 -> 清理。plan 写 plan.json 传给 pwsh。"""
    src, outdir = os.path.abspath(src), os.path.abspath(outdir)   # COM 不认相对路径
    tmp = os.path.join(outdir, '_frames_tmp.pptx')
    shutil.copy2(src, tmp)
    pj = []
    for sp in plan:
        init_hide = sorted({sid for fr in sp['frames'] for sid, _ in fr['show']})  # 入场前一律隐藏
        pj.append({'no': sp['no'], 'init_hide': init_hide,
                   'frames': [{'id': fr['id'],
                               'show': [sid for sid, _ in fr['show']],
                               'hide': [sid for sid, _ in fr['hide']]} for fr in sp['frames']]})
    plan_path = os.path.join(outdir, 'plan.json')
    with open(plan_path, 'w', encoding='utf-8') as fh:
        json.dump({'slides': pj}, fh, ensure_ascii=False)
    code = '''
$ErrorActionPreference='Stop'
$plan = Get-Content -Raw -Encoding UTF8 '%s' | ConvertFrom-Json
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open('%s', 0, 0, 0)
$miss = @{}
try {
  foreach ($sp in $plan.slides) {
    $s = $pres.Slides.Item([int]$sp.no)
    $map = @{}
    foreach ($sh in $s.Shapes) { $map[[int]$sh.Id] = $sh }
    $reg = { param($sid) if (-not $map.ContainsKey([int]$sid)) { $miss[[string]$sid + '@s' + $sp.no] = 1 } }
    foreach ($v in $sp.init_hide) { $sid = [int]$v; if ($map.ContainsKey($sid)) { $map[$sid].Visible = 0 } else { & $reg $sid } }
    $nn = "{0:d2}" -f [int]$sp.no
    foreach ($fr in $sp.frames) {
      foreach ($v in $fr.show) { $sid = [int]$v; if ($map.ContainsKey($sid)) { $map[$sid].Visible = -1 } else { & $reg $sid } }
      foreach ($v in $fr.hide) { $sid = [int]$v; if ($map.ContainsKey($sid)) { $map[$sid].Visible = 0 } else { & $reg $sid } }
      $s.Export((Join-Path '%s' ("slide" + $nn + "_" + $fr.id + ".png")), 'PNG', %d, %d)
    }
  }
  Write-Output ('OK pages=' + $plan.slides.Count + ' miss=' + $miss.Count)
  if ($miss.Count -gt 0) { Write-Output ('MISS ' + (($miss.Keys | Sort-Object) -join ',')) }
  $pres.Close()
} finally {
  if ($app.Presentations.Count -eq 0) { $app.Quit() }
}
''' % (plan_path, tmp, outdir, W, H)
    r = subprocess.run(['pwsh', '-NoProfile', '-Command', code], capture_output=True, text=True,
                       timeout=600, encoding='utf-8', errors='replace')
    out = ((r.stdout or '') + (r.stderr or '')).strip().replace('\r', '')
    subprocess.run(['pwsh', '-NoProfile', '-Command',
        "Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'wpp.exe') -and $_.CommandLine -match 'Embedding' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"],
        capture_output=True)
    time.sleep(1)
    for ln in out.splitlines(): print(' ', ln)
    try: os.remove(tmp)
    except OSError: pass
    return 'OK' in out

def main():
    src = sys.argv[1]
    pages = {int(a) for a in sys.argv[2:]} or None
    outdir = os.path.splitext(src)[0] + '_frames'
    os.makedirs(outdir, exist_ok=True)
    for f in os.listdir(outdir):                     # 清旧帧：重跑后帧号语义可能变，残留图会污染对照
        if f.endswith('.png'):
            os.remove(os.path.join(outdir, f))
    plan, notes = build_plan(src, pages)
    if not plan:
        print('没有可处理的动画页（全部整页常显）')
        return
    with open(os.path.join(outdir, 'plan_full.json'), 'w', encoding='utf-8') as fh:
        json.dump([{'no': sp['no'],
                    'info': {str(k): v for k, v in sp['info'].items()},
                    'clicks': sp['clicks'],
                    'frames': [{'id': fr['id'], 'label': fr['label'],
                                'show': [[sid, pc] for sid, pc in fr['show']],
                                'hide': [[sid, pc] for sid, pc in fr['hide']]} for fr in sp['frames']],
                    'ignored': {str(k): v for k, v in sp['ignored'].items()}} for sp in plan],
                   fh, ensure_ascii=False, indent=1)
    ok = render_frames(src, outdir, plan)
    with open(os.path.join(outdir, 'frames_manifest.md'), 'w', encoding='utf-8') as fh:
        fh.write(manifest_md(src, plan, notes))
    total = sum(len(sp['frames']) for sp in plan)
    pngs = len([f for f in os.listdir(outdir) if f.endswith('.png')])
    print('%s -> %s' % (os.path.basename(src), outdir))
    print('页: %s | 应出 %d 帧, 实出 %d 张 PNG | %s' % (
        ','.join(str(sp['no']) for sp in plan), total, pngs, '渲染OK' if ok else '渲染失败'))

if __name__ == '__main__':
    main()
