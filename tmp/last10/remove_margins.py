from lxml import etree
import zipfile,shutil
from pathlib import Path
p=Path('output/전주이씨_혜령군2_마지막10페이지_세로쓰기.docx')
shutil.copy2(p,'tmp/last10/before_margin_fix.docx')
with zipfile.ZipFile(p) as z:data={n:z.read(n) for n in z.namelist()}
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';V='urn:schemas-microsoft-com:vml';O='urn:schemas-microsoft-com:office:office';A='http://schemas.openxmlformats.org/drawingml/2006/main'
ns={'w':W,'v':V,'a':A}
count=0
for name in list(data):
 if not(name.startswith('word/') and name.endswith('.xml')):continue
 root=etree.fromstring(data[name]);changed=False
 for tb in root.findall('.//v:textbox',ns):
  tb.set('inset','0pt,0pt,0pt,0pt');tb.set('{'+O+'}insetmode','custom');tb.getparent().set('{'+O+'}insetmode','custom')
  style=tb.get('style','').rstrip(';');tb.set('style',style+';mso-fit-shape-to-text:t;')
  count+=1;changed=True
 for b in root.findall('.//a:bodyPr',ns):
  for key in ['lIns','rIns','tIns','bIns']:b.set(key,'0')
  changed=True
 for para in root.findall('.//w:txbxContent/w:p',ns):
  pp=para.find('{'+W+'}pPr')
  if pp is None:pp=etree.Element('{'+W+'}pPr');para.insert(0,pp)
  for tag in ['ind','spacing','snapToGrid']:
   for e in pp.findall('{'+W+'}'+tag):pp.remove(e)
  etree.SubElement(pp,'{'+W+'}ind',{'{'+W+'}left':'0','{'+W+'}right':'0','{'+W+'}firstLine':'0','{'+W+'}hanging':'0'})
  etree.SubElement(pp,'{'+W+'}spacing',{'{'+W+'}before':'0','{'+W+'}after':'0','{'+W+'}beforeAutospacing':'0','{'+W+'}afterAutospacing':'0','{'+W+'}line':'240','{'+W+'}lineRule':'auto'})
  etree.SubElement(pp,'{'+W+'}snapToGrid',{'{'+W+'}val':'0'})
  changed=True
 if changed:data[name]=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
 for n,b in data.items():z.writestr(n,b)
print('Zeroed margins and removed fixed line clipping in',count,'text boxes.')
