import json,unicodedata
from collections import defaultdict
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from xml.sax.saxutils import escape
G=json.load(open('tmp/convert/geometry.json')); D=Document();s=D.sections[0]
s.page_width=Pt(G['width']);s.page_height=Pt(G['height']);s.top_margin=s.bottom_margin=s.left_margin=s.right_margin=Pt(0)
p=D.add_paragraph(); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=Pt(1)
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
i=0

def addbox(x,y,w,h,content,sz,vertical=True):
 global i
 i+=1
 direction='<w:textDirection w:val="tbRl"/>' if vertical else ''
 flow='layout-flow:vertical-ideographic;' if vertical else ''
 xml=f'''<w:r {NS}><w:pict><v:rect id="text{i}" style="position:absolute;margin-left:{x}pt;margin-top:{y}pt;width:{w}pt;height:{h}pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;z-index:{i}" filled="f" stroked="f"><v:textbox inset="0,0,0,0" style="{flow}"><w:txbxContent><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{round(sz*20)}" w:lineRule="exact"/>{direction}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="AppleMyungjo" w:hAnsi="AppleMyungjo" w:eastAsia="AppleMyungjo"/><w:sz w:val="{round(sz*2)}"/><w:szCs w:val="{round(sz*2)}"/><w:lang w:eastAsia="ko-KR"/></w:rPr><w:t xml:space="preserve">{escape(content)}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r>'''
 p._p.append(parse_xml(xml))

def line(x,y,w,h,weight):
 global i
 i+=1
 p._p.append(parse_xml(f'<w:r {NS}><w:pict><v:rect id="line{i}" style="position:absolute;margin-left:{x}pt;margin-top:{y}pt;width:{max(w,0.01)}pt;height:{max(h,0.01)}pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;z-index:0" filled="f" strokecolor="black" strokeweight="{weight}pt"/></w:pict></w:r>'))
for r in G['rects']:line(r['x0'],r['top'],r['width'],r['height'],r['linewidth'])
for l in G['lines']:line(l['x0'],l['top'],l['width'],l['height'],l['linewidth'])
groups=defaultdict(list)
for c in G['chars']:
 if c['top']<110 or not c['text'].strip(): continue
 key=(round(c['x0'],1),round(c['height'],1),int(c['top']//213.42))
 groups[key].append(c)
for k, chars in groups.items():
 chars.sort(key=lambda c:c['top'])
 # Split separate column blocks where the source leaves a gap.
 chunks=[]
 for c in chars:
  if not chunks or c['top']-chunks[-1][-1]['bottom']>3:chunks.append([])
  chunks[-1].append(c)
 for ch in chunks:
  text=''.join(unicodedata.normalize('NFKC',c['text']).replace('넴','〇') for c in ch)
  sz=ch[0]['height'];x=min(c['x0'] for c in ch);y=ch[0]['top'];height=ch[-1]['bottom']-y
  addbox(x,y,sz+1,height+3,text,sz,True)
addbox(458,91,35,20,'342',15.5,False)
D.save('output/전주이씨_혜령군2_400페이지_세로쓰기.docx')
print('text boxes',len(groups),'shapes',i)
