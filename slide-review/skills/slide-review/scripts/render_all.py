# -*- coding: utf-8 -*-
"""COM 全页渲染 + 每次清理 WPS 进程（subprocess 传参，避开 bash 转义坑）。"""
import subprocess, sys, os, time
def render(pptx, outdir, w=1600, h=900):
    os.makedirs(outdir, exist_ok=True)
    for f in os.listdir(outdir):
        if f.endswith('.png'): os.remove(os.path.join(outdir, f))
    code = f'''
$ErrorActionPreference='Stop'
$app = New-Object -ComObject PowerPoint.Application
try {{
  $pres = $app.Presentations.Open('{pptx}', -1, 0, 0)
  foreach ($s in $pres.Slides) {{
    $n = "{{0:d2}}" -f $s.SlideIndex
    $s.Export((Join-Path '{outdir}' ("slide"+$n+".png")), 'PNG', {w}, {h})
  }}
  Write-Output ('OK ' + $pres.Slides.Count)
  $pres.Close()
}} finally {{ if ($app.Presentations.Count -eq 0) {{ $app.Quit() }} }}'''
    r = subprocess.run(['pwsh','-NoProfile','-Command',code], capture_output=True, text=True, timeout=420, encoding='utf-8', errors='replace')
    out=((r.stdout or '')+(r.stderr or '')).strip().replace('\r','')
    subprocess.run(['pwsh','-NoProfile','-Command',
      "Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'wpp.exe') -and $_.CommandLine -match 'Embedding' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"], capture_output=True)
    time.sleep(1)
    print(pptx, '->', out.splitlines()[-1] if out else 'NO OUTPUT')
    return 'OK' in out
if __name__=='__main__':
    w = int(sys.argv[3]) if len(sys.argv) > 3 else 1600
    h = int(sys.argv[4]) if len(sys.argv) > 4 else 900
    ok1 = render(sys.argv[1], sys.argv[2], w, h)
    sys.exit(0 if ok1 else 1)


def render_pages(pptx, outdir, pages, w=1600, h=900):
    """只渲染指定页（1-based 页码集合），输出 slideNN.png；不清空 outdir。

    供审阅页预览使用（预览只需改动涉及页，全册渲染太慢）。"""
    os.makedirs(outdir, exist_ok=True)
    ps = sorted({int(p) for p in pages})
    if not ps:
        return True
    code = f'''
$ErrorActionPreference='Stop'
$want = @({','.join(str(p) for p in ps)})
$app = New-Object -ComObject PowerPoint.Application
try {{
  $pres = $app.Presentations.Open('{pptx}', -1, 0, 0)
  foreach ($s in $pres.Slides) {{
    if ($want -contains $s.SlideIndex) {{
      $n = "{{0:d2}}" -f $s.SlideIndex
      $s.Export((Join-Path '{outdir}' ("slide"+$n+".png")), 'PNG', {w}, {h})
    }}
  }}
  Write-Output ('OK ' + $want.Count)
  $pres.Close()
}} finally {{ if ($app.Presentations.Count -eq 0) {{ $app.Quit() }} }}'''
    r = subprocess.run(['pwsh','-NoProfile','-Command',code], capture_output=True, text=True, timeout=420, encoding='utf-8', errors='replace')
    out=((r.stdout or '')+(r.stderr or '')).strip().replace('\r','')
    subprocess.run(['pwsh','-NoProfile','-Command',
      "Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'wpp.exe') -and $_.CommandLine -match 'Embedding' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"], capture_output=True)
    time.sleep(1)
    print(pptx, '->', out.splitlines()[-1] if out else 'NO OUTPUT')
    return 'OK' in out
