#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 페이지 → DOCX 변환기 (족보 세로쓰기 전용). 페이지마다 DOCX 파일 하나를 만든다.

  python3 tools/pdf_page_to_docx.py "전주이씨 혜령군-2.pdf" 400            # 한 페이지
  python3 tools/pdf_page_to_docx.py "전주이씨 혜령군-1.pdf" 236 237 238    # 여러 페이지
  python3 tools/pdf_page_to_docx.py "전주이씨 혜령군-2.pdf" 300-320 --out docx/vol2

동작
- 텍스트 레이어가 있는 페이지: PDF 안의 글자 좌표를 그대로 세로쓰기 텍스트박스로 절대 배치한다.
- 이미지 글리프만 있는 페이지(1권 1~372쪽): 글리프별 OCR(PaddleOCR 한국어 + 중국어 모델)로 글자를 읽고,
  신뢰도 0.8 미만 글자는 노란 형광으로 표시한다.
- 격자선·테두리는 도형으로, 큰 그림은 이미지로 넣는다.
- 글자 사이의 빈 칸은 공백 문자가 아니라 자간(字間)으로 맞춘다. 폰트가 바뀌어도 위치가 유지된다.
- OCR 결과는 tools/glyph_cache.json 에 글리프 해시별로 저장된다. 항목에 "fix": "글자" 를 넣으면
  다음 변환부터 그 글자를 쓴다(같은 글리프가 나오는 모든 페이지에 적용).
"""
import sys, os, re, json, hashlib, argparse, statistics, tempfile
from xml.sax.saxutils import escape
import fitz
import numpy as np
from PIL import Image
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE_PATH = os.path.join(HERE, "glyph_cache.json")
EMU = 12700                      # EMU per pt
FONT = "Batang"                  # 바탕 (MS Office 동봉). 없으면 Word가 대체
SPACE_EM = 0.33                  # 바탕체 ASCII 공백의 진행 폭(em) — 세로쓰기 단어 사이에 사용
LOW_CONF = 0.8                   # 이 미만은 형광 표시
PIC_MIN_PT = 60                  # 이보다 큰 이미지는 글리프가 아니라 그림으로 취급
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
# 2권 RHSBSY 폰트의 깨진 ToUnicode 매핑 (페이지 번호 숫자와 〇). 800쪽 전수 대조로 확인.
DIGIT_MAP = {"뎔": "0", "뎠": "1", "뎡": "2", "뎨": "3", "뎬": "4", "도": "5", "독": "6", "돈": "7", "돋": "8", "돌": "9", "넴": "〇"}
# 세로쓰기 괄호 글리프를 가로로 눕혀 인식하면 열고 닫음이 뒤집혀 나온다.
PAREN_FIX = {"）": "（", "（": "）", ")": "（", "(": "）"}


def is_hangul(s): return len(s) == 1 and "가" <= s <= "힣"
def is_hanja(s): return len(s) == 1 and ("一" <= s <= "鿿" or "豈" <= s <= "﫿" or "㐀" <= s <= "䶿")
def emu(pt): return int(round(pt * EMU))


class Glyph:
    __slots__ = ("x0", "y0", "x1", "y1", "ch", "size", "conf")
    def __init__(s, rect, ch, size, conf=1.0):
        s.x0, s.y0, s.x1, s.y1 = rect.x0, rect.y0, rect.x1, rect.y1
        s.ch, s.size, s.conf = ch, size, conf
    xc = property(lambda s: (s.x0 + s.x1) / 2)
    yc = property(lambda s: (s.y0 + s.y1) / 2)


# ---------------------------------------------------------------- 글자 수집: 텍스트 레이어
def glyphs_from_text(page):
    M = page.rotation_matrix; out = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0: continue
        for l in b["lines"]:
            for s in l["spans"]:
                chars = [c for c in s["text"] if not c.isspace()]
                if not chars: continue
                r = fitz.Rect(s["bbox"]) * M; r.normalize(); n = len(chars)
                for i, c in enumerate(chars):
                    if n == 1: rr = r
                    elif r.height >= r.width: rr = fitz.Rect(r.x0, r.y0 + r.height * i / n, r.x1, r.y0 + r.height * (i + 1) / n)
                    else: rr = fitz.Rect(r.x0 + r.width * i / n, r.y0, r.x0 + r.width * (i + 1) / n, r.y1)
                    if s["font"] == "RHSBSY": c = DIGIT_MAP.get(c, c)
                    out.append(Glyph(rr, c, s["size"]))
    return out


# ---------------------------------------------------------------- 글자 수집: 이미지 글리프 + OCR
def recognize(tiles, cache):
    import logging; logging.disable(logging.WARNING)
    from paddleocr import TextRecognition
    rko = TextRecognition(model_name="korean_PP-OCRv5_mobile_rec")
    rzh = TextRecognition(model_name="PP-OCRv5_server_rec")
    paths = [fn for _, fn in tiles]
    ko = [(r["rec_text"], float(r["rec_score"])) for r in rko.predict(paths, batch_size=64)]
    zh = [(r["rec_text"], float(r["rec_score"])) for r in rzh.predict(paths, batch_size=64)]
    for (key, _), (kt, ks), (zt, zs) in zip(tiles, ko, zh):
        kt, zt = kt.strip(), zt.strip()
        if is_hangul(kt) and ks >= 0.5: ch, conf = kt, ks
        elif is_hanja(zt) and zs >= 0.5: ch, conf = zt, zs
        elif kt and ks >= zs: ch, conf = kt, ks
        elif zt: ch, conf = zt, zs
        else: ch, conf = "□", 0.0
        if len(ch) > 1: ch, conf = ch[0], min(conf, 0.5)      # 글리프 하나 = 글자 하나
        ch = PAREN_FIX.get(ch, ch)
        if ch.isascii() and ch.isalnum(): conf = min(conf, 0.5)  # 라틴 문자·숫자로 읽힌 글리프는 대개 오인식 → 검수 대상
        cache.setdefault(key, {}).update({"char": ch, "conf": round(conf, 3), "ko": [kt, round(ks, 3)], "zh": [zt, round(zs, 3)]})


def _tile(im, box, fn):
    """im 에서 box(px)를 잘라 흰 여백을 두르고 저장"""
    c = im.crop(box); pad = Image.new("L", (c.width + 48, c.height + 24), 255); pad.paste(c, (24, 12)); pad.save(fn)


def stack_segments(arr):
    """세로로 긴 잉크 덩어리를 빈 줄 기준으로 글자 단위 구간(높이 비율)으로 나눈다. 폭보다 훨씬 긴 경우에만 호출."""
    h, w = arr.shape; rows = arr.any(axis=1); bands = []; start = None
    for i, v in enumerate(rows):
        if v and start is None: start = i
        if not v and start is not None: bands.append([start, i]); start = None
    if start is not None: bands.append([start, h])
    bands = [b for b in bands if b[1] - b[0] >= 0.15 * w] or bands   # 이웃 글자의 부스러기(아주 얇은 띠)는 무시
    groups = []
    for b in bands:                                   # 한 글자 높이(≈폭) 안에 드는 띠는 같은 글자
        if groups and b[1] - groups[-1][0] <= 1.25 * w: groups[-1][1] = b[1]
        else: groups.append(list(b))
    merged = []
    for g in groups:                                  # 점 하나 크기의 조각은 이웃 글자에 합친다
        if merged and (g[1] - g[0] < 0.2 * w): merged[-1][1] = g[1]
        else: merged.append(g)
    groups = merged
    return [(0.0, 1.0)] if len(groups) <= 1 else [(g0 / h, g1 / h) for g0, g1 in groups]


def glyphs_from_images(doc, page, cache, pics):
    M = page.rotation_matrix; entries = []; tiles = []; queued = set()
    tmp = tempfile.mkdtemp(prefix="glyphs_")
    for info in page.get_image_info(xrefs=True):
        r = fitz.Rect(info["bbox"]) * M; r.normalize()
        if r.width < 0.5 or r.height < 0.5: continue
        if max(r.width, r.height) > PIC_MIN_PT: pics.append(r); continue
        R = r + (-0.5, -0.5, 0.5, 0.5); xref = info.get("xref") or 0; h = None; pm = None
        if xref > 0:
            try: h = hashlib.md5(doc.xref_stream_raw(xref)).hexdigest()
            except Exception: h = None
        if h is None:                                     # 인라인 이미지 등: 렌더링 픽셀로 해시
            pm = page.get_pixmap(clip=R, dpi=400, colorspace=fitz.csGRAY); h = "px" + hashlib.md5(pm.samples).hexdigest()
        meta = cache.get(h)
        if meta is None or "ink" not in meta:
            if pm is None: pm = page.get_pixmap(clip=R, dpi=400, colorspace=fitz.csGRAY)
            im = Image.frombytes("L", (pm.width, pm.height), pm.samples); arr = np.array(im) < 128
            ys, xs = np.where(arr); ink, segs = None, [(0.0, 1.0)]
            if len(xs):                                   # 비트맵 안의 실제 잉크 범위 (여백 제거)
                x0, y0, x1, y1 = int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1
                ink = [x0 / pm.width, y0 / pm.height, x1 / pm.width, y1 / pm.height]
                if (y1 - y0) >= 1.6 * (x1 - x0): segs = stack_segments(arr[y0:y1, x0:x1])   # 여러 글자가 한 비트맵에
            meta = cache[h] = {"ink": ink, "segs": segs}
            if len(segs) == 1:
                if h not in queued:
                    fn = os.path.join(tmp, h + ".png"); _tile(im, (0, 0, im.width, im.height), fn); tiles.append((h, fn)); queued.add(h)
            else:
                px0, py0, px1, py1 = ink[0] * pm.width, ink[1] * pm.height, ink[2] * pm.width, ink[3] * pm.height
                for i, (f0, f1) in enumerate(segs):
                    key = f"{h}#{i}"
                    if key in queued: continue
                    box = (max(0, int(px0) - 2), max(0, int(py0 + (py1 - py0) * f0) - 2), int(px1) + 2, int(py0 + (py1 - py0) * f1) + 2)
                    fn = os.path.join(tmp, f"{h}_{i}.png"); _tile(im, box, fn); tiles.append((key, fn)); queued.add(key)
        ink, segs = meta["ink"], meta.get("segs", [(0.0, 1.0)])
        gr = fitz.Rect(R.x0 + ink[0] * R.width, R.y0 + ink[1] * R.height, R.x0 + ink[2] * R.width, R.y0 + ink[3] * R.height) if ink else r
        if len(segs) == 1: entries.append((r, gr, h))
        else:
            for i, (f0, f1) in enumerate(segs):
                entries.append((r, fitz.Rect(gr.x0, gr.y0 + gr.height * f0, gr.x1, gr.y0 + gr.height * f1), f"{h}#{i}"))
    if tiles:
        print(f"  OCR: {len(tiles)} new glyph tiles ({len(entries) - len(tiles)} placements from cache)")
        recognize(tiles, cache)
    out = []
    for r, gr, key in entries:
        e = cache[key]; fix = e.get("fix")
        size = min(max(r.width, r.height), max(gr.width, gr.height) / 0.85)   # 잉크 크기로 글자 크기 추정
        out.append(Glyph(gr, fix or e.get("char", "□"), size, 1.0 if fix else e.get("conf", 0.0)))
    return out


# ---------------------------------------------------------------- 열/런 구성
def adv(ch, size):                       # 글자 하나가 진행 방향으로 차지하는 길이
    return size * 0.55 if (ch.isascii() and not ch.isspace()) else size

def _sz(g, med): return g.size if g.size >= 0.6 * med else med   # 작은 문장부호 글리프는 열 크기로

def make_run(gl, kind, med):
    """parts: [문자, 크기, 형광여부, 추가 자간]. 글자 사이 빈 칸은 자간으로 표현하므로 폰트가 달라도 위치가 유지된다."""
    parts = []; base = []; prev = None; est = 0.0
    for g in gl:
        size = _sz(g, med)
        if prev is not None:
            pitch = (g.yc - prev.yc) if kind == "v" else (g.xc - prev.xc)
            e = pitch - (adv(g.ch, size) + adv(prev.ch, _sz(prev, med))) / 2
            if e / med >= 0.6: parts.append([" ", med, False, e - SPACE_EM * med]); est += e    # 단어 사이: 공백 + 자간
            elif e / med > 0.3: parts[-1][3] += e; est += e                                     # 작은 틈: 앞 글자 자간
            else: base.append(e)
        parts.append([g.ch, size, g.conf < LOW_CONF, 0.0]); est += adv(g.ch, size)
        prev = g
    spacing = max(-2.0, min(6.0, statistics.median(base))) if base else 0.0
    for p in parts:
        if p[0] == " ": p[3] -= spacing
    est += max(spacing, 0) * len(gl)
    W = max(_sz(g, med) for g in gl)
    if kind == "v":
        axis = statistics.median(g.xc for g in gl); bw = W + 2
        y0 = gl[0].yc - _sz(gl[0], med) / 2 - 1; y1 = gl[-1].yc + _sz(gl[-1], med) / 2 + 1
        return dict(kind=kind, x0=axis - bw / 2, y0=y0, w=bw, h=max(y1 - y0, est * 1.08 + 3, W), line=bw, parts=parts, spacing=spacing)
    yc = statistics.median(g.yc for g in gl); bh = W * 1.25 + 2
    x0 = gl[0].xc - adv(gl[0].ch, _sz(gl[0], med)) / 2 - 1; x1 = gl[-1].xc + adv(gl[-1].ch, _sz(gl[-1], med)) / 2 + 1
    return dict(kind=kind, x0=x0, y0=yc - bh / 2, w=max(x1 - x0, est * 1.1 + 3), h=bh, line=bh, parts=parts, spacing=spacing)


def build_runs(glyphs, hrules=()):
    """hrules: [(y, x0, x1)] 가로 괘선. 괘선을 가로지르는 글자는 한 상자에 넣지 않는다."""
    runs = []; used = set()
    digits = sorted((g for g in glyphs if g.ch.isdigit()), key=lambda g: (round(g.yc / 4), g.xc))
    i = 0
    while i < len(digits):                                   # 가로로 놓인 숫자 = 페이지 번호
        grp = [digits[i]]; j = i + 1
        while j < len(digits) and abs(digits[j].yc - grp[-1].yc) < 0.3 * grp[-1].size and 0 < digits[j].xc - grp[-1].xc < 1.2 * grp[-1].size:
            grp.append(digits[j]); j += 1
        if len(grp) >= 2: runs.append(make_run(grp, "h", statistics.median(g.size for g in grp))); used.update(id(g) for g in grp)
        i = j
    rest = [g for g in glyphs if id(g) not in used]
    if not rest: return runs
    pmed = statistics.median(g.size for g in rest)
    normal = sorted((g for g in rest if g.size >= 0.5 * pmed), key=lambda g: g.xc)
    small = [g for g in rest if g.size < 0.5 * pmed]
    cols = []; cur = []
    for g in normal:                                         # x 중심이 가까운 글자끼리 한 열 (작은 쪽 크기 기준)
        if cur and g.xc - cur[-1].xc > 0.45 * min(g.size, cur[-1].size): cols.append(cur); cur = []
        cur.append(g)
    if cur: cols.append(cur)
    axes = [statistics.median(g.xc for g in c) for c in cols]
    for g in small:                                          # 작은 문장부호는 가장 가까운 열에 붙인다
        if cols:
            k = min(range(len(cols)), key=lambda k: abs(axes[k] - g.xc))
            if abs(axes[k] - g.xc) < 0.7 * statistics.median(x.size for x in cols[k]): cols[k].append(g); continue
        cols.append([g]); axes.append(g.xc)
    for col, axis in zip(cols, axes):
        col.sort(key=lambda g: g.yc)
        big = [g.size for g in col if g.size >= 0.5 * pmed]
        med = statistics.median(big) if big else statistics.median(g.size for g in col)
        seg = [col[0]]
        for g in col[1:]:
            prev = seg[-1]
            e = (g.yc - prev.yc) - (_sz(g, med) + _sz(prev, med)) / 2
            crossed = any(prev.yc < y < g.yc and x0 - 2 <= axis <= x1 + 2 for y, x0, x1 in hrules)
            if crossed or e >= 1.5 * med: runs.append(make_run(seg, "v", med)); seg = [g]
            else: seg.append(g)
        runs.append(make_run(seg, "v", med))
    return runs


# ---------------------------------------------------------------- DOCX XML
_id = [100]
def next_id(): _id[0] += 1; return _id[0]

def anchor(x, y, cx, cy, inner, z):
    i = next_id(); cx = max(cx, 0.3); cy = max(cy, 0.3)
    return (f'<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{z}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>{emu(x)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{emu(y)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{emu(cx)}" cy="{emu(cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{i}" name="obj{i}"/><wp:cNvGraphicFramePr/>{inner}</wp:anchor>')

WPS = '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
WPS_END = '</wps:wsp></a:graphicData></a:graphic>'

def shape_xml(cx, cy, fill, stroke_w):
    cx = max(cx, 0.3); cy = max(cy, 0.3)
    f = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' if fill else '<a:noFill/>'
    ln = f'<a:ln w="{emu(stroke_w)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' if stroke_w else '<a:ln><a:noFill/></a:ln>'
    return (WPS + '<wps:cNvSpPr/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(cx)}" cy="{emu(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{f}{ln}</wps:spPr>'
            '<wps:bodyPr/>' + WPS_END)

def textbox_xml(run):
    rs = []
    for text, size, hl, extra in run["parts"]:
        hp = int(round(size * 2)); spv = run["spacing"] + extra
        sp = f'<w:spacing w:val="{int(round(spv * 20))}"/>' if abs(spv) > 0.05 else ''
        h = '<w:highlight w:val="yellow"/>' if hl else ''
        rs.append(f'<w:r><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/>{sp}<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/>{h}</w:rPr>'
                  f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r>')
    p = f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{int(round(run["line"] * 20))}" w:lineRule="exact"/></w:pPr>{"".join(rs)}</w:p>'
    vert = "eaVert" if run["kind"] == "v" else "horz"
    return (WPS + '<wps:cNvSpPr txBox="1"/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(run["w"])}" cy="{emu(run["h"])}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
            f'<wps:txbx><w:txbxContent>{p}</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr rot="0" vert="{vert}" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>' + WPS_END)

def picture_xml(rId, cx, cy):
    i = next_id()
    return ('<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            f'<pic:nvPicPr><pic:cNvPr id="{i}" name="pic{i}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(cx)}" cy="{emu(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic>')


# ---------------------------------------------------------------- 페이지 변환
def convert_page(pdf, pno, outpath, cache):
    page = pdf[pno - 1]; W, H = page.rect.width, page.rect.height; M = page.rotation_matrix
    has_text = len(page.get_text().strip()) > 20
    pics = []
    glyphs = glyphs_from_text(page) if has_text else glyphs_from_images(pdf, page, cache, pics)
    drawings = page.get_drawings()
    hrules = []
    for dr in drawings:
        for it in dr["items"]:
            if it[0] == "l":
                a, b = fitz.Point(it[1]) * M, fitz.Point(it[2]) * M
                if abs(a.y - b.y) < 0.5: hrules.append((a.y, min(a.x, b.x), max(a.x, b.x)))
            elif it[0] == "re" and dr.get("fill") is not None and sum(dr["fill"]) / 3 < 0.5:
                rr = fitz.Rect(it[1]) * M; rr.normalize()
                if rr.height < 3 and rr.width > 20: hrules.append(((rr.y0 + rr.y1) / 2, rr.x0, rr.x1))
    runs = build_runs(glyphs, hrules) if glyphs else []
    d = Document(); sec = d.sections[0]
    sec.page_width = Pt(W); sec.page_height = Pt(H)
    for a in ("left_margin", "right_margin", "top_margin", "bottom_margin", "header_distance", "footer_distance"): setattr(sec, a, 0)
    para = d.paragraphs[0] if d.paragraphs else d.add_paragraph()
    z = [1]
    def add(inner_anchor_xml):
        para._p.append(parse_xml(f'<w:r {NS}><w:drawing>{inner_anchor_xml}</w:drawing></w:r>')); z[0] += 1
    nshape = 0
    for dr in drawings:                                   # 격자선·테두리
        fill, color, w = dr.get("fill"), dr.get("color"), dr.get("width") or 0
        for it in dr["items"]:
            if it[0] == "re":
                r = fitz.Rect(it[1]) * M; r.normalize()
                if fill is not None and sum(fill) / 3 < 0.5:
                    add(anchor(r.x0, r.y0, r.width, r.height, shape_xml(r.width, r.height, True, 0), z[0])); nshape += 1
                elif fill is None and color is not None and w > 0:
                    add(anchor(r.x0, r.y0, r.width, r.height, shape_xml(r.width, r.height, False, w), z[0])); nshape += 1
            elif it[0] == "l" and color is not None:
                a, b = fitz.Point(it[1]) * M, fitz.Point(it[2]) * M; t = max(w, 0.4)
                if abs(a.y - b.y) < 0.5: r = fitz.Rect(min(a.x, b.x), a.y - t / 2, max(a.x, b.x), a.y + t / 2)
                elif abs(a.x - b.x) < 0.5: r = fitz.Rect(a.x - t / 2, min(a.y, b.y), a.x + t / 2, max(a.y, b.y))
                else: continue
                add(anchor(r.x0, r.y0, r.width, r.height, shape_xml(r.width, r.height, True, 0), z[0])); nshape += 1
    for r in pics:                                        # 큰 그림
        fn = tempfile.mktemp(suffix=".png"); page.get_pixmap(clip=r, dpi=300).save(fn)
        rId, _ = d.part.get_or_add_image(fn)
        add(anchor(r.x0, r.y0, r.width, r.height, picture_xml(rId, r.width, r.height), z[0]))
    for run in runs:                                      # 글자
        if not any(t.strip() for t, _, _, _ in run["parts"]): continue
        add(anchor(run["x0"], run["y0"], run["w"], run["h"], textbox_xml(run), z[0]))
    d.save(outpath)
    low = sum(1 for g in glyphs if g.conf < LOW_CONF)
    print(f"  p{pno}: {'text' if has_text else 'OCR'} mode, glyphs={len(glyphs)}, text boxes={len(runs)}, shapes={nshape}, pictures={len(pics)}"
          + (f", low-confidence glyphs={low}" if not has_text else "") + f" -> {outpath}")


def parse_pages(specs, n):
    pages = []
    for s in specs:
        m = re.fullmatch(r"(\d+)-(\d+)", s)
        pages += list(range(int(m.group(1)), int(m.group(2)) + 1)) if m else [int(s)]
    return [p for p in pages if 1 <= p <= n]


def main():
    global FONT
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf"); ap.add_argument("pages", nargs="+", help="예: 236  또는 300-320")
    ap.add_argument("--out", default=None, help="출력 폴더 (기본: PDF 옆의 docx/)")
    ap.add_argument("--font", default=FONT, help=f"글꼴 이름 (기본: {FONT})")
    a = ap.parse_args()
    FONT = a.font
    pdf = fitz.open(a.pdf); base = os.path.splitext(os.path.basename(a.pdf))[0]
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.pdf)), "docx"); os.makedirs(out, exist_ok=True)
    cache = json.load(open(CACHE_PATH, encoding="utf-8")) if os.path.exists(CACHE_PATH) else {}
    for p in parse_pages(a.pages, len(pdf)):
        convert_page(pdf, p, os.path.join(out, f"{base}_p{p:04d}.docx"), cache)
        json.dump(cache, open(CACHE_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=0)

if __name__ == "__main__":
    main()
