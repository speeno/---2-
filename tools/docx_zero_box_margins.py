#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX 안의 VML 텍스트박스(v:rect + v:textbox) 내부 여백을 모두 0으로 만든다.

  python3 tools/docx_zero_box_margins.py 파일.docx [--font AppleMyungjo] [--no-backup]

- inset 0, 단락 앞뒤 간격 0, 들여쓰기 0, 줄 높이 = 글자 크기(고정)
- 상자 크기를 글자 범위에 정확히 맞춘다: 세로쓰기는 너비 = 글자 크기, 높이 = 글자 진행 폭의 합(오른쪽 위 모서리 고정),
  가로쓰기는 너비 = 글자 폭의 합, 높이 = 글자 크기(왼쪽 위 모서리 고정)
- 줄바꿈 안 함(mso-wrap-style:none)을 지정해 반올림 오차로 마지막 글자가 다음 열로 넘어가는 일을 막는다
- 문서 기본 단락 간격(docDefaults)도 0으로 만든다
글자 진행 폭은 지정 글꼴의 실제 메트릭(세로 vmtx, 가로 hmtx)으로 계산한다.
"""
import sys, re, io, zipfile, glob, argparse, shutil
from PIL import ImageFont

def font_path(name):
    for d in ("/System/Library/Fonts/Supplemental", "/System/Library/Fonts", "/Library/Fonts", "~/Library/Fonts",
              "/Applications/Microsoft Word.app/Contents/Resources/DFonts"):
        for ext in ("ttf", "ttc", "otf"):
            hits = glob.glob(f"{d.replace('~', __import__('os').path.expanduser('~'))}/{name}*.{ext}")
            if hits: return hits[0]
    raise SystemExit(f"글꼴 파일을 찾지 못했습니다: {name}")

def metrics(name):
    path = font_path(name); pil = ImageFont.truetype(path, 1000)
    vadv = 1.0
    try:
        from fontTools.ttLib import TTFont
        f = TTFont(path, fontNumber=0); upm = f["head"].unitsPerEm; cmap = f.getBestCmap()
        if "vmtx" in f and ord("世") in cmap: vadv = f["vmtx"].metrics[cmap[ord("世")]][0] / upm
    except Exception: pass
    return path, vadv, (lambda ch: pil.getlength(ch) / 1000)

def process(src, font="AppleMyungjo", backup=True):
    path, vadv, hadv = metrics(font)
    z = zipfile.ZipFile(src); x = z.read("word/document.xml").decode("utf-8"); st = z.read("word/styles.xml").decode("utf-8")
    log = []
    def fix(m):
        rect, tb, body = m.group(1), m.group(2), m.group(3)
        style = re.search(r'style="([^"]*)"', rect).group(1)
        g = lambda k: float(re.search(k + r':([-\d.]+)pt', style).group(1))
        ml, mt, w, h = g("margin-left"), g("margin-top"), g("width"), g("height"); vert = "vertical" in tb
        szs = [int(v) for v in re.findall(r'<w:sz w:val="(\d+)"', body)]
        if not szs: return m.group(0)
        F = max(szs) / 2; adv = 0.0
        for r in re.findall(r'<w:r>(.*?)</w:r>', body, re.S):
            sz = re.search(r'<w:sz w:val="(\d+)"', r); s = (int(sz.group(1)) if sz else max(szs)) / 2
            t = "".join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', r))
            adv += sum(((hadv(c) if c.isascii() else vadv) * s) if vert else (hadv(c) * s) for c in t)
        if vert: nw, nh, nml, nmt = F, adv, ml + (w - F), mt          # 세로: 오른쪽 위 고정
        else:    nw, nh, nml, nmt = adv, F, ml, mt                     # 가로: 왼쪽 위 고정
        s2 = re.sub(r'margin-left:[-\d.]+pt', f'margin-left:{nml:.3f}pt', style)
        s2 = re.sub(r'margin-top:[-\d.]+pt', f'margin-top:{nmt:.3f}pt', s2)
        s2 = re.sub(r'width:[-\d.]+pt', f'width:{nw:.3f}pt', s2); s2 = re.sub(r'height:[-\d.]+pt', f'height:{nh:.3f}pt', s2)
        if "mso-wrap-style" not in s2: s2 = s2.rstrip(";") + ";mso-wrap-style:none"
        rect2 = rect.replace(f'style="{style}"', f'style="{s2}"')
        tb2 = re.sub(r'inset="[^"]*"', 'inset="0,0,0,0"', tb) if 'inset=' in tb else tb.replace("<v:textbox", '<v:textbox inset="0,0,0,0"', 1)
        body2 = re.sub(r'<w:spacing [^/]*/>', f'<w:spacing w:before="0" w:after="0" w:line="{int(round(F * 20))}" w:lineRule="exact"/>', body)
        body2 = re.sub(r'<w:ind [^/]*/>', '', body2).replace('<w:pPr>', '<w:pPr><w:ind w:left="0" w:right="0" w:firstLine="0"/>', 1) if '<w:pPr>' in body2 else body2
        log.append((re.sub(r'<[^>]+>', '', body), round(w - nw, 2), round(h - nh, 2)))
        return rect2 + tb2 + body2 + '</v:textbox></v:rect>'
    x2 = re.sub(r'(<v:rect\b[^>]*>)(<v:textbox[^>]*>)(<w:txbxContent>.*?</w:txbxContent>)</v:textbox></v:rect>', fix, x, flags=re.S)
    st2 = re.sub(r'<w:pPrDefault>.*?</w:pPrDefault>', '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>', st, flags=re.S)
    if backup: shutil.copy(src, re.sub(r'\.docx$', '.원본백업.docx', src)) if not glob.glob(re.sub(r'\.docx$', '.원본백업.docx', src)) else None
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for it in z.infolist():
            data = z.read(it.filename)
            if it.filename == "word/document.xml": data = x2.encode("utf-8")
            elif it.filename == "word/styles.xml": data = st2.encode("utf-8")
            out.writestr(it, data)
    z.close(); open(src, "wb").write(buf.getvalue())
    print(f"{src}: boxes={len(log)}, font={path}, vertical advance={vadv} em")
    print("  size reduction (pt) w/h, first 5:", log[:5])

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", nargs="+"); ap.add_argument("--font", default="AppleMyungjo"); ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args()
    for f in a.docx: process(f, a.font, not a.no_backup)
