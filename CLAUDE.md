# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this directory is

Not a codebase. It holds two typeset volumes of a Korean clan genealogy (족보): 全州李氏 惠寧君派譜 (Jeonju Yi clan, Prince Hyeryeong branch). There is no source code and no build, lint, or test tooling. Work here is document analysis and conversion (text extraction, OCR, DOCX/HWP re-typesetting). The user communicates in Korean.

| File | PDF pages | Size | Produced by | Text layer |
|---|---|---|---|---|
| `전주이씨 혜령군-1.pdf` (卷之一) | 472 | 165 MB | Adobe InDesign CS5.5 | **None** on pp. 1–372: every glyph is a separate 1-bit CCITT stencil image (about 400 per page), no fonts at all. Embedded fonts and extractable text only from p. 373 on (tree charts drawn with box-drawing characters). |
| `전주이씨 혜령군-2.pdf` | 800 | 8.7 MB | Founder PDF Library | Extractable on every page (subset-embedded CID TrueType with ToUnicode). |

Both volumes use 188 × 257 mm pages.

## Layout and encoding facts that affect extraction

- Traditional vertical writing (세로쓰기): columns run top-to-bottom and are read right-to-left. Vol. 1 mid-volume pages are prose (Hangul with Hanja in parentheses) with the running header 全州李氏惠寧君派譜 卷之一 along the right edge. Vol. 2 is the generation grid: six generation rows per page (e.g. 十六世 … 二十一世) labeled on the right edge, names in large type (about 20 pt) with small Hangul readings (about 9 pt) beside them, dates in Chinese numerals, cross-references like 見上六七.
- 790 of vol. 2's 800 pages, and most of vol. 1's text-bearing tail, carry `/Rotate 90`: content is stored landscape and rotated for display. PyMuPDF text coordinates come back in the unrotated space; apply `page.rotation_matrix` before comparing them with a rendered page.
- Vol. 2 has no line or column structure inside the PDF: every character is its own positioned text object (one-character spans). Columns and table cells have to be inferred geometrically from glyph bboxes. This is why generic PDF-to-DOCX converters emit one text box per character.
- Page-number digits in vol. 2 use the RHSBSY font, whose ToUnicode map is wrong: they extract as garbage Hangul syllables such as 뎨, 뎬, 뎡, 넴. Do not treat those as real text.
- Printed page numbers do not match PDF page indices (PDF p. 236 of vol. 1 is printed "222"; PDF p. 400 of vol. 2 is printed "342").
- Vol. 2 fonts (RHSBKS, RHSMKS, RHSBSY, RKMJKS) are proprietary Korean typesetting fonts, subset-embedded and not installed here. Closest installed substitutes for re-typesetting are `batang.ttc` (ships with Microsoft Office) and `AppleMyungjo.ttf`.

## Tools and commands

Poppler is in `/opt/homebrew/bin` (pdfinfo, pdftotext, pdfimages, pdffonts, pdftoppm, pdfseparate). `python3` imports PyMuPDF (`fitz`), `python-docx`, Pillow, numpy, PaddleOCR 3.3 (PP-OCRv5 models already cached under `~/.paddlex/official_models`: `korean_PP-OCRv5_mobile_rec`, `PP-OCRv5_server_rec`, `chinese_cht_PP-OCRv3_mobile_rec`), and EasyOCR. LibreOffice (`soffice`), tesseract (`kor`, `kor_vert`, `chi_tra_vert`, `jpn_vert`), pandoc, and the Xcode `swift` toolchain (for Apple Vision OCR scripts) are on PATH. Google Chrome can render HTML headlessly: `"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --screenshot=out.png --window-size=711,972 --allow-file-access-from-files file://…`. Microsoft Word and Hancom Office HWP are installed under /Applications.

Check for a text layer before extracting from a page:

```bash
pdftotext -f N -l N FILE - | wc -c   # near zero means image-only
pdffonts  -f N -l N FILE             # empty table means no text objects
```

Extract vol. 2 text. Default `pdftotext` ordering scrambles the vertical columns; `-raw` follows content-stream order, which walks the columns:

```bash
pdftotext -raw -f 400 -l 400 "전주이씨 혜령군-2.pdf" -
```

Render a page to look at it, then Read the PNG:

```bash
pdftoppm -r 110 -f N -l N -png FILE outprefix
```

Split out a single page for experiments:

```bash
pdfseparate -f N -l N FILE out.pdf
```

## Vol. 1 glyph geometry (the key to OCR)

- `page.get_image_info(xrefs=True)` returns every glyph placement with its bbox. Cluster by x-center (gap > 12 pt starts a new column), order columns right-to-left and glyphs top-to-bottom. The rightmost small column is the running header, not body text.
- Across pp. 1–372 there are 123,008 placements but only 8,388 unique bitmaps (2,886 cover 95% of uses). Labeling each unique bitmap once and decoding by lookup yields exact text for the whole volume; only the label set needs review.
- Mixed Hangul/Hanja in one OCR pass fails with every engine here. What worked (vol. 1 p. 236, tested 2026-09-09): render each glyph tile at 400 dpi with white padding, run PaddleOCR `TextRecognition` with both `korean_PP-OCRv5_mobile_rec` and `PP-OCRv5_server_rec`, keep the Korean answer if it is a Hangul syllable with score ≥ 0.5, else the Chinese answer if it is an ideograph. Result: Hanja recall 97%, Hangul 88%. Laying a column's tiles out as a horizontal strip and running Apple Vision (`ko-KR`) or PaddleOCR `korean` gives 90–94% on Hangul thanks to context, so the best pipeline is strip-level Korean for Hangul runs plus per-glyph Chinese for Hanja. Vertical parentheses come out swapped as ）（ and must be mapped back.

## Vol. 2 → HTML (works far better than DOCX)

Every character has a bbox; map it with `fitz.Rect(bbox) * page.rotation_matrix`, emit one absolutely positioned `<span>` per glyph (CSS `pt` units, page div 532.96 × 728.53 pt), and convert `page.get_drawings()` lines/rects into an inline SVG. Rendered in Chrome with AppleMyungjo as the substitute font this scored ink IoU 0.43 against the original render, versus 0.011 for LibreOffice's DOCX. Remaining gaps are font shape/metrics and the RHSBSY page-number glyphs.

## DOCX conversion (`tools/pdf_page_to_docx.py`)

`python3 tools/pdf_page_to_docx.py "<pdf>" 400` writes `docx/<pdf-stem>_p0400.docx`, one file per page. Page lists and ranges (`236 237`, `300-320`) work; `--out DIR` and `--font NAME` (default Batang) are optional. Pages with a text layer are placed from PDF glyph coordinates. Image-only pages go through per-glyph OCR: each glyph tile is recognized by both `korean_PP-OCRv5_mobile_rec` and `PP-OCRv5_server_rec`, the answer is chosen by script and confidence, glyphs under 0.8 confidence (and any read as Latin letters or digits) are highlighted yellow. Results are cached by glyph-bitmap hash in `tools/glyph_cache.json`; add `"fix": "字"` to an entry to override that glyph everywhere it recurs. Stacked bitmaps (the vol. 1 running header stores several characters per image) are split on blank rows.

Layout model: one absolutely positioned DrawingML text box (`wps:bodyPr vert="eaVert"`) per vertical run, split at grid rules and at gaps of 1.5 em or more; gaps inside a run are expressed as character spacing (`w:spacing`) plus an ASCII space for word gaps, never as U+3000 (LibreOffice shifts eaVert boxes containing U+3000 by tens of points). Grid lines become filled rectangles, big images become pictures, page is 188 × 257 mm with zero margins.

Verification status (2026-09-09): Word opens the files and reads every text box (AppleScript reported 53 shapes and correct text on vol. 2 p. 400), but Word's scripted `save as … format PDF` fails with error -1708 and `screencapture` needs Screen Recording permission, so Word's rendering is unverified. LibreOffice renders the layout correctly except that its vertical text advances by the font's line height rather than 1 em, so columns come out about 7% longer than the original. Ink IoU vs the original page (AppleMyungjo): vol. 2 p. 400 = 0.42, vol. 1 p. 236 = 0.21. Test pages: vol. 1 p. 236 has 442 glyphs, 89 flagged for review.

## Zeroing text-box margins in someone else's DOCX (`tools/docx_zero_box_margins.py`)

`output/` holds DOCX files the user produced elsewhere (VML `v:rect` + `v:textbox` boxes, `layout-flow:vertical-ideographic`, A4 page, AppleMyungjo). `python3 tools/docx_zero_box_margins.py FILE.docx` sets inset 0, paragraph spacing and indent 0, exact line height = font size, shrinks every box to its text extent (vertical: width = font size, height = sum of vertical advances from the font's vmtx, right-top corner fixed; horizontal: hmtx widths), adds `mso-wrap-style:none` so rounding can never wrap the last glyph into a hidden second column, zeroes docDefaults spacing, and writes a `*.원본백업.docx` copy first. Word's `overflowing` text-frame property is NOT a usable oracle: it flips 0→69→0 over time and settles at 0 even for boxes half the needed size. Word's text-frame `margin left/right/top/bottom` and shape `width/height` read back reliably via AppleScript and were used to confirm the result (0/0/0/0, 14.5 × 14.55 pt for a single 14.5 pt glyph).

## DOCX → HTML (`tools/docx_to_html.py`)

`python3 tools/docx_to_html.py FILE.docx [-o OUT.html] [--font-file NAME=PATH]` turns an absolutely positioned DOCX (VML `v:rect/v:oval/v:line` + `v:textbox`, or DrawingML `wp:anchor` + `wps:wsp`/`pic:pic`) into one self-contained HTML file: one `.page` div per section (size from `pgSz`), every shape as an absolutely positioned div in pt, vertical boxes as `writing-mode: vertical-rl; text-orientation: mixed` (matches Word's rotated Latin), insets as padding, exact/multiple line heights, font/size/bold/italic/color/highlight/letter-spacing per run, `@page` + page-break for printing. `--font-file` embeds a font as base64 so glyphs match on machines without it. Flow (inline) content is ignored by design. Generic exporters fail on these files: pandoc emits an empty file, LibreOffice's HTML export drops the vertical writing mode.

The user's `output/*_세로쓰기.docx` files reference fonts named `GenealogyRHSBKS`, `GenealogyRHSB01`, `GenealogyRHSBSY`, `GenealogyRHSMKS` (apparently extracted from the PDF's embedded fonts) that are not installed anywhere on this Mac, so Word and browsers both substitute; ask the user for the font files if glyph-identical output matters.

## Git

The directory became a git repository on 2026-09-09 (remote `origin` = https://github.com/speeno/---2-.git, branch `main`, nothing pushed yet). The two source PDFs are listed in `.gitignore` and were removed from every commit with `git filter-branch` because the 165 MB volume exceeds GitHub's 100 MB file limit. Pitfall: `filter-branch` deletes files it removed from history from the working tree too; the PDFs were restored byte-for-byte from `refs/backup/before-pdf-removal`, which still holds the pre-rewrite history (and the big blobs). Once the user confirms the push, that ref can be dropped with `git update-ref -d refs/backup/before-pdf-removal && git gc --prune=now`. Never `git push --mirror` while it exists. `tmp/` (the user's own experiments, including a vendored fontTools) and `.DS_Store` are tracked; nobody has asked to change that.

## Known dead ends (tested 2026-09-09)

- LibreOffice `--infilter="writer_pdf_import"` to DOCX: on vol. 1 it embeds the stencil glyphs as transparent PNG masks that render blank; on vol. 2 it emits one text box per character on an A4 page with the grid misplaced. Neither preserves layout.
- tesseract `kor_vert` at 300 dpi on whole vol. 1 pages reads most Hangul but garbles nearly all Hanja; the horizontal `kor` model is unusable on these pages, and even on re-laid horizontal strips tesseract skips whole columns (63% overall).
- Apple Vision with `zh-Hant` first, or PaddleOCR `chinese_cht`, on mixed strips: Hanja 30–70% but Hangul 0%. Vision with `ko-KR` first or PaddleOCR `korean`: Hangul 90–94% but Hanja 0%. Never run one language model over mixed script.
- pdf2docx is not installed and would not handle vertical per-glyph text anyway.
