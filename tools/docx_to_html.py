#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
절대 배치 DOCX(세로쓰기 텍스트박스) → HTML 변환기.

  python3 tools/docx_to_html.py 파일.docx                       # 같은 폴더에 파일.html
  python3 tools/docx_to_html.py 파일.docx -o 결과.html
  python3 tools/docx_to_html.py 파일.docx --font-file GenealogyRHSBKS=/path/font.ttf   # 글꼴을 HTML 안에 포함

- VML(v:rect/v:oval/v:roundrect/v:line/v:shape + v:textbox)과 DrawingML(wp:anchor + wps:wsp, pic:pic) 도형을
  페이지 기준 좌표(pt) 그대로 CSS 절대 배치로 옮긴다.
- 세로쓰기(layout-flow:vertical-ideographic, bodyPr vert="eaVert") → writing-mode: vertical-rl (라틴 문자는 눕힘)
- 구역(sectPr)마다 한 페이지. 인쇄하면 한 페이지가 한 장이 된다(@page, page-break).
- 글꼴 크기·굵기·기울임·색·형광·자간, 상자 안쪽 여백(inset), 줄 높이(고정/배수)를 옮긴다.
- 흐름 배치(인라인) 내용은 옮기지 않는다. 이 도구는 상자 좌표로만 이루어진 문서용이다.
"""
import sys, os, re, base64, argparse, zipfile, html, mimetypes
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
V = "urn:schemas-microsoft-com:vml"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ns = {"w": W, "v": V, "wp": WP, "a": A, "wps": WPS, "pic": PIC, "r": R}
EMU = 12700.0
DEFAULT_STACK = "AppleMyungjo, Batang, 'Noto Serif KR', serif"


def wv(el, name): return el.get(f"{{{W}}}{name}")
def style_dict(s): return {k.strip(): v.strip() for k, v in (kv.split(":", 1) for kv in (s or "").split(";") if ":" in kv)}
def pt(v, default=0.0):
    if v is None: return default
    v = str(v).strip().lower()
    m = re.match(r'(-?[\d.]+)\s*(pt|in|cm|mm|px|emu)?', v)
    if not m: return default
    n = float(m.group(1)); u = m.group(2) or "pt"
    return n * {"pt": 1, "in": 72, "cm": 72 / 2.54, "mm": 72 / 25.4, "px": 0.75, "emu": 1 / EMU}[u]
def color(v, default="#000"):
    if not v: return default
    v = v.strip()
    if v.startswith("#"): return v
    if re.fullmatch(r'[0-9A-Fa-f]{6}', v): return "#" + v
    return v  # 이름 색상


# ---------------------------------------------------------------- 단락/런 → HTML
def run_html(r, fonts_used):
    rpr = r.find("w:rPr", ns); css = []
    if rpr is not None:
        rf = rpr.find("w:rFonts", ns)
        fams = []
        if rf is not None:
            for k in ("eastAsia", "ascii", "hAnsi", "cs"):
                f = wv(rf, k)
                if f and f not in fams and not f.endswith("Theme"): fams.append(f)
        for f in fams: fonts_used.add(f)
        if fams: css.append("font-family:" + ",".join(f"'{f}'" for f in fams) + "," + DEFAULT_STACK)
        sz = rpr.find("w:sz", ns)
        if sz is not None: css.append(f"font-size:{int(wv(sz, 'val')) / 2:g}pt")
        if rpr.find("w:b", ns) is not None and wv(rpr.find("w:b", ns), "val") not in ("0", "false"): css.append("font-weight:bold")
        if rpr.find("w:i", ns) is not None and wv(rpr.find("w:i", ns), "val") not in ("0", "false"): css.append("font-style:italic")
        u = rpr.find("w:u", ns)
        if u is not None and wv(u, "val") not in (None, "none"): css.append("text-decoration:underline")
        c = rpr.find("w:color", ns)
        if c is not None and wv(c, "val") not in (None, "auto"): css.append("color:" + color(wv(c, "val")))
        hl = rpr.find("w:highlight", ns)
        if hl is not None: css.append("background:" + {"yellow": "#ff0", "green": "#0f0", "cyan": "#0ff", "magenta": "#f0f", "red": "#f00", "lightGray": "#ccc"}.get(wv(hl, "val"), wv(hl, "val")))
        sp = rpr.find("w:spacing", ns)
        if sp is not None and wv(sp, "val"): css.append(f"letter-spacing:{int(wv(sp, 'val')) / 20:g}pt")
        va = rpr.find("w:vertAlign", ns)
        if va is not None: css.append({"superscript": "vertical-align:super;font-size:smaller", "subscript": "vertical-align:sub;font-size:smaller"}.get(wv(va, "val"), ""))
    parts = []
    for ch in r:
        tag = etree.QName(ch).localname
        if tag == "t": parts.append(html.escape(ch.text or ""))
        elif tag == "tab": parts.append("\t")
        elif tag == "br": parts.append("<br>")
        elif tag == "sym": parts.append(html.escape(chr(int(wv(ch, "char"), 16))) if wv(ch, "char") else "")
    txt = "".join(parts)
    return f'<span style="{";".join(c for c in css if c)}">{txt}</span>' if css else f"<span>{txt}</span>"


def para_html(p, fonts_used):
    ppr = p.find("w:pPr", ns); css = ["margin:0"]
    if ppr is not None:
        sp = ppr.find("w:spacing", ns)
        if sp is not None:
            if wv(sp, "before"): css.append(f"margin-top:{int(wv(sp, 'before')) / 20:g}pt")
            if wv(sp, "after"): css.append(f"margin-bottom:{int(wv(sp, 'after')) / 20:g}pt")
            ln, rule = wv(sp, "line"), wv(sp, "lineRule") or "auto"
            if ln:
                if rule == "exact": css.append(f"line-height:{int(ln) / 20:g}pt")
                elif rule == "atLeast": css.append(f"line-height:{int(ln) / 20:g}pt")
                else: css.append("line-height:normal" if int(ln) == 240 else f"line-height:{int(ln) / 240:g}")
        ind = ppr.find("w:ind", ns)
        if ind is not None and wv(ind, "firstLine"): css.append(f"text-indent:{int(wv(ind, 'firstLine')) / 20:g}pt")
        jc = ppr.find("w:jc", ns)
        if jc is not None: css.append("text-align:" + {"center": "center", "right": "end", "end": "end", "both": "justify", "distribute": "justify"}.get(wv(jc, "val"), "start"))
    runs = "".join(run_html(r, fonts_used) for r in p.iter(f"{{{W}}}r"))
    return f'<p style="{";".join(css)}">{runs}</p>'


def textbox_html(content, fonts_used):
    return "".join(para_html(p, fonts_used) for p in content.findall("w:p", ns))


# ---------------------------------------------------------------- VML 도형
def vml_elements(container, fonts_used, warnings):
    out = []
    for el in container.iter(f"{{{V}}}rect", f"{{{V}}}oval", f"{{{V}}}roundrect", f"{{{V}}}line", f"{{{V}}}shape"):
        tag = etree.QName(el).localname; st = style_dict(el.get("style"))
        if st.get("mso-position-horizontal-relative", "page") not in ("page", "margin") or st.get("mso-position-vertical-relative", "page") not in ("page", "margin"):
            warnings.add("페이지 기준이 아닌 도형이 있어 위치가 어긋날 수 있습니다")
        z = st.get("z-index", "0")
        if tag == "line":
            x1, y1 = (pt(v) for v in (el.get("from") or "0,0").split(",")); x2, y2 = (pt(v) for v in (el.get("to") or "0,0").split(","))
            sw = pt(el.get("strokeweight"), 0.75)
            if abs(y1 - y2) < 0.01: out.append(f'<div style="left:{min(x1, x2):.3f}pt;top:{y1 - sw / 2:.3f}pt;width:{abs(x2 - x1):.3f}pt;height:{sw:.3f}pt;background:{color(el.get("strokecolor"))};z-index:{z}"></div>')
            elif abs(x1 - x2) < 0.01: out.append(f'<div style="left:{x1 - sw / 2:.3f}pt;top:{min(y1, y2):.3f}pt;width:{sw:.3f}pt;height:{abs(y2 - y1):.3f}pt;background:{color(el.get("strokecolor"))};z-index:{z}"></div>')
            else: warnings.add("대각선은 옮기지 않았습니다")
            continue
        L, T, Wd, Ht = pt(st.get("margin-left")), pt(st.get("margin-top")), pt(st.get("width")), pt(st.get("height"))
        filled = el.get("filled", "t") not in ("f", "false"); stroked = el.get("stroked", "t") not in ("f", "false")
        sw = pt(el.get("strokeweight"), 0.75) if stroked else 0
        css = [f"left:{L:.3f}pt", f"top:{T:.3f}pt", f"width:{Wd:.3f}pt", f"height:{Ht:.3f}pt", f"z-index:{z}", "box-sizing:border-box"]
        if tag == "oval": css.append("border-radius:50%")
        if tag == "roundrect": css.append("border-radius:" + (st.get("arcsize", "0.2").rstrip("f") and "10%"))
        tb = el.find("v:textbox", ns)
        if filled and tb is None: css.append("background:" + color(el.get("fillcolor"), "#fff"))
        elif filled and tb is not None and el.get("fillcolor"): css.append("background:" + color(el.get("fillcolor"), "#fff"))
        if stroked:
            if Ht < 1 and tb is None: css[3] = f"height:{sw:.3f}pt"; css[1] = f"top:{T - sw / 2 + Ht / 2:.3f}pt"; css.append("background:" + color(el.get("strokecolor")))
            elif Wd < 1 and tb is None: css[2] = f"width:{sw:.3f}pt"; css[0] = f"left:{L - sw / 2 + Wd / 2:.3f}pt"; css.append("background:" + color(el.get("strokecolor")))
            else: css.append(f"border:{sw:.3f}pt solid {color(el.get('strokecolor'))}")
        if tb is None:
            out.append(f'<div style="{";".join(css)}"></div>'); continue
        tst = style_dict(tb.get("style")); vert = tst.get("layout-flow", "").startswith("vertical")
        ins = [pt(x) for x in (tb.get("inset") or "0.1in,0.05in,0.1in,0.05in").split(",")]
        while len(ins) < 4: ins.append(ins[-1] if ins else 0)
        css.append(f"padding:{ins[1]:.3f}pt {ins[2]:.3f}pt {ins[3]:.3f}pt {ins[0]:.3f}pt")
        content = tb.find("w:txbxContent", ns)
        inner = textbox_html(content, fonts_used) if content is not None else ""
        out.append(f'<div class="{"tb v" if vert else "tb h"}" style="{";".join(css)}">{inner}</div>')
    return out


# ---------------------------------------------------------------- DrawingML 도형
def drawingml_elements(container, fonts_used, warnings, rels, zf):
    out = []
    for anc in container.iter(f"{{{WP}}}anchor"):
        ph, pv = anc.find("wp:positionH", ns), anc.find("wp:positionV", ns)
        if ph is None or pv is None or ph.get("relativeFrom") != "page" or pv.get("relativeFrom") != "page":
            warnings.add("페이지 기준이 아닌 DrawingML 도형이 있어 위치가 어긋날 수 있습니다")
        def off(p):
            o = p.find("wp:posOffset", ns) if p is not None else None
            return int(o.text) / EMU if o is not None and o.text else 0.0
        L, T = off(ph), off(pv); ext = anc.find("wp:extent", ns)
        Wd, Ht = int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU; z = anc.get("relativeHeight", "0")
        css = [f"left:{L:.3f}pt", f"top:{T:.3f}pt", f"width:{Wd:.3f}pt", f"height:{Ht:.3f}pt", f"z-index:{z}", "box-sizing:border-box"]
        wsp = anc.find(".//wps:wsp", ns); pic = anc.find(".//pic:pic", ns)
        if pic is not None:
            blip = pic.find(".//a:blip", ns); rid = blip.get(f"{{{R}}}embed") if blip is not None else None
            target = rels.get(rid)
            if target and ("word/" + target) in zf.namelist():
                data = zf.read("word/" + target); mt = mimetypes.guess_type(target)[0] or "image/png"
                out.append(f'<img style="{";".join(css)}" src="data:{mt};base64,{base64.b64encode(data).decode()}">')
            else: warnings.add("그림 파일을 찾지 못한 도형이 있습니다")
            continue
        if wsp is None: continue
        sppr = wsp.find("wps:spPr", ns); geom = sppr.find("a:prstGeom", ns) if sppr is not None else None
        if geom is not None and geom.get("prst") == "ellipse": css.append("border-radius:50%")
        fill = sppr.find("a:solidFill/a:srgbClr", ns) if sppr is not None else None
        if fill is not None: css.append("background:#" + fill.get("val"))
        ln = sppr.find("a:ln", ns) if sppr is not None else None
        if ln is not None and ln.find("a:noFill", ns) is None:
            lc = ln.find(".//a:srgbClr", ns); css.append(f"border:{int(ln.get('w', 9525)) / EMU:.3f}pt solid #{lc.get('val') if lc is not None else '000'}")
        txbx = wsp.find(".//w:txbxContent", ns); body = wsp.find("wps:bodyPr", ns)
        if txbx is None:
            out.append(f'<div style="{";".join(css)}"></div>'); continue
        vert = body is not None and body.get("vert") in ("eaVert", "vert", "vert270", "wordArtVert")
        ins = [int(body.get(k, d)) / EMU for k, d in (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720))] if body is not None else [7.2, 3.6, 7.2, 3.6]
        css.append(f"padding:{ins[1]:.3f}pt {ins[2]:.3f}pt {ins[3]:.3f}pt {ins[0]:.3f}pt")
        out.append(f'<div class="{"tb v" if vert else "tb h"}" style="{";".join(css)}">{textbox_html(txbx, fonts_used)}</div>')
    return out


# ---------------------------------------------------------------- 문서 → 페이지
def convert(src, dst, font_files):
    zf = zipfile.ZipFile(src); root = etree.fromstring(zf.read("word/document.xml")); body = root.find("w:body", ns)
    rels = {}
    if "word/_rels/document.xml.rels" in zf.namelist():
        for rel in etree.fromstring(zf.read("word/_rels/document.xml.rels")): rels[rel.get("Id")] = rel.get("Target")
    sections, cur = [], []
    for k in body:
        cur.append(k)
        if etree.QName(k).localname == "sectPr" or (etree.QName(k).localname == "p" and k.find("w:pPr/w:sectPr", ns) is not None):
            sections.append(cur); cur = []
    if cur: (sections[-1] if sections else sections.append(cur) or sections[-1]).extend(cur) if sections else sections.append(cur)
    fonts_used, warnings, pages = set(), set(), []
    for sec in sections:
        last = sec[-1]; sp = last if etree.QName(last).localname == "sectPr" else last.find("w:pPr/w:sectPr", ns)
        pg = sp.find("w:pgSz", ns) if sp is not None else None
        PW, PH = (int(wv(pg, "w")) / 20, int(wv(pg, "h")) / 20) if pg is not None else (595.3, 841.9)
        els = []
        for k in sec:
            els += vml_elements(k, fonts_used, warnings); els += drawingml_elements(k, fonts_used, warnings, rels, zf)
        pages.append((PW, PH, "".join(els)))
    faces = []
    for name, path in font_files.items():
        ext = os.path.splitext(path)[1].lower(); fmt = {".ttf": "truetype", ".otf": "opentype", ".woff": "woff", ".woff2": "woff2", ".ttc": "collection"}.get(ext, "truetype")
        data = base64.b64encode(open(path, "rb").read()).decode()
        faces.append(f'@font-face{{font-family:"{name}";src:url("data:font/{ext.lstrip(".")};base64,{data}") format("{fmt}")}}')
    PW0, PH0 = pages[0][0], pages[0][1]
    css = f'''{"".join(faces)}
html,body{{margin:0;background:#888}}
.page{{position:relative;background:#fff;overflow:hidden;margin:16pt auto;box-shadow:0 0 6pt rgba(0,0,0,.4)}}
.page>div,.page>img{{position:absolute;margin:0;padding:0}}
.tb{{overflow:visible}}
.tb p{{white-space:pre;margin:0}}
.v{{writing-mode:vertical-rl;text-orientation:mixed}}
.v p{{line-height:normal}}
@page{{size:{PW0:.2f}pt {PH0:.2f}pt;margin:0}}
@media print{{html,body{{background:#fff}}.page{{margin:0;box-shadow:none;page-break-after:always;break-after:page}}}}'''
    title = os.path.splitext(os.path.basename(src))[0]
    doc = [f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>{html.escape(title)}</title>',
           f'<!-- 원본: {html.escape(os.path.basename(src))} / 사용 글꼴: {", ".join(sorted(fonts_used))} -->', f'<style>{css}</style></head><body>']
    for i, (PW, PH, inner) in enumerate(pages, 1):
        doc.append(f'<div class="page" id="p{i}" style="width:{PW:.2f}pt;height:{PH:.2f}pt">{inner}</div>')
    doc.append('</body></html>')
    open(dst, "w", encoding="utf-8").write("\n".join(doc))
    ntext = sum(1 for _, _, inner in pages for _ in re.finditer(r'class="tb', inner))
    print(f"{dst}: pages={len(pages)}, text boxes={ntext}, fonts={sorted(fonts_used)}" + (f", missing embedded font for: {sorted(f for f in fonts_used if f not in font_files)}" if fonts_used - set(font_files) else ""))
    for w_ in sorted(warnings): print("  주의:", w_)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", nargs="+"); ap.add_argument("-o", "--out", default=None, help="출력 HTML (입력이 하나일 때)")
    ap.add_argument("--font-file", action="append", default=[], help="이름=경로 : 이 글꼴 파일을 HTML 안에 포함")
    a = ap.parse_args()
    fonts = dict(s.split("=", 1) for s in a.font_file)
    for f in a.docx:
        dst = a.out if (a.out and len(a.docx) == 1) else os.path.splitext(f)[0] + ".html"
        convert(f, dst, fonts)

if __name__ == "__main__":
    main()
