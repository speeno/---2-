import sys,io,glob,re,json,os,unicodedata
sys.path.insert(0,'tmp/last10/deps')
from pypdf import PdfReader
from fontTools.ttLib import TTFont,newTable
from fontTools.ttLib.tables._c_m_a_p import CmapSubtable
os.makedirs('tmp/last10/fonts',exist_ok=True)
r=PdfReader(glob.glob('*-2.pdf')[0]); result={}
fix={'넴':'〇','뎨':'3','뎬':'4','돈':'7','도':'5','독':'6','돋':'8','돌':'9','피':'勳'}
for v in r.pages[790]['/Resources']['/Font'].values():
 f=v.get_object();desc=f['/DescendantFonts'][0].get_object()['/FontDescriptor']; family='Genealogy'+desc['/FontFamily']; font=TTFont(io.BytesIO(desc['/FontFile2'].get_data()));order=font.getGlyphOrder()
 cmap={}
 for block in re.findall(r'beginbfchar(.*?)endbfchar',f['/ToUnicode'].get_data().decode(),re.S):
  for a,b in re.findall(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>',block):
   ch=bytes.fromhex(b).decode('utf-16-be'); ch=fix.get(ch,ch) if ('SY' in family or '01' in family) else ch
   if len(ch)==1 and int(a,16)<len(order):cmap[ord(ch)]=order[int(a,16)]
 font['cmap']=newTable('cmap');font['cmap'].tableVersion=0;font['cmap'].tables=[]
 for platform,encoding in [(0,3),(3,1)]:
  tab=CmapSubtable.newSubtable(4);tab.platformID=platform;tab.platEncID=encoding;tab.language=0;tab.cmap=cmap; font['cmap'].tables.append(tab)
 for nid,val in [(1,family),(2,'Regular'),(3,family),(4,family),(6,family)]:
  font['name'].setName(val,nid,3,1,0x409)
 font['post']=newTable('post');font['post'].formatType=3; font['post'].italicAngle=0;font['post'].underlinePosition=-100;font['post'].underlineThickness=50;font['post'].isFixedPitch=0
 for a in ['minMemType42','maxMemType42','minMemType1','maxMemType1']:setattr(font['post'],a,0)
 path='tmp/last10/fonts/'+family+'.ttf';font.save(path)
 result[desc['/FontFamily']]={'family':family,'path':path,'ascent':desc['/Ascent'],'descent':desc['/Descent']}
json.dump(result,open('tmp/last10/fonts.json','w'))
print(result)
