# -*- coding: utf-8 -*-
"""增强 dump：每页顶层形状的 id/名/类型/prstGeom/位置/flip/文字，供写击序作业表用。"""
import sys, re, zipfile
from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P, "a": A}
src = sys.argv[1] if len(sys.argv) > 1 else r"E:\Zcode Data\ppt\1.4 加速度.pptx"
z = zipfile.ZipFile(src)
files = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
               key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))

EMU = 914400.0
for n in files:
    idx = int(re.search(r"slide(\d+)", n).group(1))
    root = etree.fromstring(z.read(n))
    spTree = root.find(".//p:cSld/p:spTree", NS)
    print(f"\n##### slide {idx} #####")
    if spTree is None:                      # 坏条目占位 / 结构异常
        print("  (无 spTree，跳过)")
        continue
    for ch in spTree:
        tag = etree.QName(ch).localname
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        c = ch.find(".//p:cNvPr", NS)
        sid, nm = c.get("id"), c.get("name")
        xfrm = ch.find(".//a:xfrm", NS)
        pos, flip = "", ""
        if xfrm is not None:
            off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
            if off is not None and ext is not None:
                pos = "(%.2f,%.2f) %.2fx%.2f" % (int(off.get("x")) / EMU, int(off.get("y")) / EMU,
                                                 int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU)
            fh, fv = xfrm.get("flipH"), xfrm.get("flipV")
            flip = ("H" if fh == "1" else "") + ("V" if fv == "1" else "")
            if xfrm.get("rot"):
                flip += " rot=%s" % (int(xfrm.get("rot")) / 60000)
        geom = ""
        pg = ch.find(".//a:prstGeom", NS)
        if pg is not None:
            geom = pg.get("prst")
        elif ch.find(".//a:custGeom", NS) is not None:
            geom = "custGeom"
        # 填充/线色
        fill = ""
        sf = ch.find(".//p:spPr/a:solidFill/a:srgbClr", NS)
        if sf is not None:
            fill = "fill#" + sf.get("val")
        ln = ch.find(".//p:spPr/a:ln/a:solidFill/a:srgbClr", NS)
        if ln is not None:
            fill += " ln#" + ln.get("val")
        tail = ""
        for e in ch.findall(".//a:tailEnd", NS):
            tail = "tailEnd=" + str(e.get("type"))
        if tag == "pic":
            tail += " PIC"
        t = re.sub(r"\s+", " ", "".join(ch.itertext())).strip()[:56]
        print(f"  id={sid:<4} {tag:<4} {geom:<14} {pos:<28} {flip:<8} {fill:<16} {tail} |{t}")
