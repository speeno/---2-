import json,os,uuid,zipfile
from collections import defaultdict
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.enum.section import WD_SECTION_START
from xml.sax.saxutils import escape
from lxml import etree
P=json.load(open('tmp/last10/pages.json'));F=json.load(open('tmp/last10/fonts.json'));D=Document();i=0
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
fix={'넴':'〇','뎨':'3','뎬':'4','돈':'7','도':'5','독':'6','돋':'8','돌':'9','피':'勳'}
expected=[]
def box(x,y,w,h,text,sz,font='AppleMyungjo',vertical=True):
 global i
 if not text.strip():return
 if text in ['︵','︶']:
  x+=2.0;y+=1.5;sz=4.0;w=6;h=7
 i+=1;expected.append(text)
 direction='<w:textDirection w:val="tbRl"/>' if vertical else ''
 flow='layout-flow:vertical-ideographic;' if vertical else ''
 p._p.append(parse_xml(f'''<w:r {NS}><w:pict><v:rect id="text{i}" style="position:absolute;margin-left:{x:.3f}pt;margin-top:{y:.3f}pt;width:{w:.3f}pt;height:{h:.3f}pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;z-index:{i}" filled="f" stroked="f"><v:textbox inset="0,0,0,0" style="{flow}"><w:txbxContent><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{round(sz*20)}" w:lineRule="exact"/>{direction}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}"/><w:sz w:val="{round(sz*2)}"/><w:szCs w:val="{round(sz*2)}"/><w:lang w:eastAsia="ko-KR"/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r>'''))
def line(x,y,w,h,weight):
 global i
 i+=1
 p._p.append(parse_xml(f'<w:r {NS}><w:pict><v:rect id="line{i}" style="position:absolute;margin-left:{x}pt;margin-top:{y}pt;width:{max(w,.01)}pt;height:{max(h,.01)}pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;z-index:0" filled="f" strokecolor="black" strokeweight="{weight}pt"/></w:pict></w:r>'))
def col(x,y,text,sz=9,step=None,font='AppleMyungjo'):
 step=step or sz
 for j,ch in enumerate(text):box(x,y+j*step,sz+1,sz+3,ch,sz,font)
def colophon():
 for x,year,cy,month,act in [(398.6,'二〇一五','乙未','八','印刷'),(389.1,'二〇一八','戊戌','三','發行')]:
  col(x,206.8,'西紀'+year+'年');col(x,274.3,cy);col(x,301.3,month+'月');col(x,337.3,'日');col(x,368.8,act[0]);col(x,395.8,act[1])
 for x,title in [(371.1,'發行所'),(317.1,'發行人'),(273.1,'編輯人')]:col(x,369,title,step=13.5)
 for ch,y in zip('全州李氏惠寧君派宗親會',[421,436.9,452.7,468.5,492.5,508.4,524.2,540.1,564.1,580,595.8]):col(367.6,y,ch,16)
 for text,y in [('京畿道',437.8),('水原市',469.9),('京水大路',502.2),('四四六번길',543.9),('四二',594.6)]:col(353.6,y,text,9,9.4)
 col(344.6,483.5,'電話',9,8)
 col(335.1,477.2,'FAX',9,7)
 # Parentheses and separators remain upright as printed on the vertical columns.
 for x,text,ys in [(344.6,'︵〇三一︶二二一│二〇一七番',[503.7,508.6,516.9,525.2,532.6,537.5,545.8,554.1,561.5,570.5,578,587.1,595.4,603.7]),(335.1,'︵〇三一︶二二二│二〇三三番',[504.1,508.9,516.7,524.6,531.8,536.5,544.9,553.3,560.9,570.1,577.8,587,595.4,603.8])]:
  for ch,y in zip(text,ys):col(x,y,ch,8.3)
 for x,ys,name in [(314.6,[427.1,513.1,599.2],'李卿東'),(270.6,[427,512.2,597.4],'李連東')]:
  for y,ch in zip(ys,name):col(x,y,ch,14)
 for text,y in [('서울특별시',438.4),('강서구',488.6),('가로공원로',521),('七六가길',571.3)]:col(299.6,y,text,9,9)
 col(290.6,532.1,'八四의',9);col(290.6,559.1,'一七︵화곡동︶',9,8.1)
 for text,y,step in [('서울특별시',438.6,8.8),('구로구',484.8,8.6),('개봉로',513.4,8.6),('一一길',541.9,8),('三八의一二',571.3,8)]:col(255.6,y,text,9,step)
 col(228.4,428.8,'出版',9.1);col(237,428.8,'圖書',9.1)
 # Publisher's seal is recreated as a text character within two editable circles.
 global i
 for size in [17,14]:
  i+=1;p._p.append(parse_xml(f'<w:r {NS}><w:pict><v:oval id="seal{i}" style="position:absolute;margin-left:{237-(size/2)}pt;margin-top:{410.5-size/2}pt;width:{size}pt;height:{size}pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" filled="f" strokeweight=".5pt"/></w:pict></w:r>'))
 box(230.5,404.0,15,18,'想',13)
 for y,ch in zip([461.2,530,598.8],'回想社'):col(230.1,y,ch,14)
 col(209.6,464.6,'代表',9)
 for y,ch in zip([501.1,550.3,599.5],'朴炳浩'):col(207.1,y,ch,14)
 for text,y in [('大田廣域市',464.6),('東區',509.9),('宣化路',529.6),('一八六번길',557.9),('七',603.1)]:col(196.1,y,text,9,8.5)
 col(186.1,482.6,'電話',9);col(176.1,477.2,'FAX',9,7);col(166.1,464.6,'서울支社',9);col(156.1,477.2,'FAX',9,7)
 for x,text,ys in [(186.1,'︵〇四二︶二五三│九八八一〜三番',[504.2,509.6,515,520.4,523,526.6,534.2,541.9,548.6,557.2,564.8,572.4,580.1,586.9,595.4,603.1]),(176.1,'︵〇四二︶二五三│九八九一番',[504.2,509.6,515,520.4,523.1,526.7,536.2,545.7,554.4,564.9,574.4,584,593.6,603.1]),(166.1,'︵〇二︶七一八│九八八一番',[509.3,514.7,520.1,522.8,526.4,536,545.5,554.2,564.7,574.2,583.7,593.3,602.9]),(156.1,'︵〇二︶七一八│九八八二番',[509.2,514.6,520,522.7,526.2,535.8,545.4,554.2,564.6,574.2,583.8,593.4,603])]:
  for j,(ch,y) in enumerate(zip(text,ys)):col(x,y,ch,5.4 if j in ([1,2,3] if x>170 else [1,2]) else 8.5)
 col(146.1,545.6,'出版登錄사六號',9,9.6)

for idx,g in enumerate(P):
 s=D.sections[0] if idx==0 else D.add_section(WD_SECTION_START.NEW_PAGE)
 s.page_width=Pt(g['width']);s.page_height=Pt(g['height']);s.top_margin=s.bottom_margin=s.left_margin=s.right_margin=Pt(0)
 p=D.add_paragraph();p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=Pt(1)
 if g['number']==799:colophon();continue
 for r in g['rects']:
  if r['stroke']:line(r['x0'],r['top'],r['width'],r['height'],r['linewidth'])
 for l in g['lines']:
  x=l['x0'];width=l['width']
  if g['number']==797 and width>100:width-=115.8-x;x=115.8
  line(x,l['top'],width,l['height'],l['linewidth'])
 for c in g['chars']:
  if not c['text'].strip():continue
  fam=c['fontname'].split('+')[-1];text=fix.get(c['text'],c['text']) if fam in ['RHSBSY','RHSB01'] else c['text']
  sz=c['height'];box(c['x0'],c['top'],sz+1,sz+3,text,sz,F[fam]['family'],c['top']>110)
path='output/전주이씨_혜령군2_마지막10페이지_세로쓰기.docx';D.save(path)
# Embed the source fonts for portable editable text.
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';PK='http://schemas.openxmlformats.org/package/2006/relationships'
with zipfile.ZipFile(path) as z:data={n:z.read(n) for n in z.namelist()}
ft=etree.fromstring(data['word/fontTable.xml']);rels=etree.Element('{'+PK+'}Relationships',nsmap={None:PK})
for j,item in enumerate(F.values()):
 key=uuid.uuid4();mask=bytes.fromhex(key.hex)[::-1];blob=bytearray(open(item['path'],'rb').read())
 for k in range(32):blob[k]^=mask[k%16]
 target=f'fonts/source{j}.odttf';data['word/'+target]=bytes(blob)
 font=etree.SubElement(ft,'{'+W+'}font',{'{'+W+'}name':item['family']});etree.SubElement(font,'{'+W+'}embedRegular',{'{'+R+'}id':f'rFont{j}','{'+W+'}fontKey':'{'+str(key).upper()+'}'})
 etree.SubElement(rels,'{'+PK+'}Relationship',Id=f'rFont{j}',Type=R+'/font',Target=target)
data['word/fontTable.xml']=etree.tostring(ft,xml_declaration=True,encoding='UTF-8');data['word/_rels/fontTable.xml.rels']=etree.tostring(rels,xml_declaration=True,encoding='UTF-8')
ct=etree.fromstring(data['[Content_Types].xml']);etree.SubElement(ct,'{http://schemas.openxmlformats.org/package/2006/content-types}Default',Extension='odttf',ContentType='application/vnd.openxmlformats-officedocument.obfuscatedFont');data['[Content_Types].xml']=etree.tostring(ct)
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
 for n,b in data.items():z.writestr(n,b)
json.dump(expected,open('tmp/last10/expected.json','w'),ensure_ascii=False)
print(path,'characters',sum(map(len,expected)))
