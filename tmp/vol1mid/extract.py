import json,glob,os,hashlib,collections
from pypdf import PdfReader
from PIL import Image,ImageOps,ImageDraw
from concurrent.futures import ThreadPoolExecutor
import subprocess
r=PdfReader(glob.glob('*-1.pdf')[0]);I=json.load(open('tmp/vol1mid/images.json'));unique={};glyphs=[];pages=[]
os.makedirs('tmp/vol1mid/glyphs',exist_ok=True);os.makedirs('tmp/vol1mid/strips',exist_ok=True)
for j,ims in enumerate(I):
 p=r.pages[231+j];cache={};out=[];inline=0
 for info in ims:
  name='/'+info['name']
  if not info['name'].startswith('Im'):name=f'~{inline}~';inline+=1
  if name not in cache:
   image=p.images[name].image.convert('L');key=hashlib.sha256(image.tobytes()+str(image.size).encode()).hexdigest()
   if key not in unique:
    gid=len(unique);unique[key]=gid;image.save(f'tmp/vol1mid/glyphs/{gid}.png');glyphs.append({'id':gid,'size':image.size})
   cache[name]=unique[key]
  out.append(dict(info,gid=cache[name]))
 pages.append(out)
 print(j+232,len(out),len(unique),flush=True)
json.dump(pages,open('tmp/vol1mid/placements.json','w'));json.dump(glyphs,open('tmp/vol1mid/glyphs.json','w'))
strips=[]
for j,ims in enumerate(pages):
 groups=collections.defaultdict(list)
 for im in ims:
  if 82<im['x0']<510 and im['top']>110:groups[round(im['x0'],0)].append(im)
 for k,(x,items) in enumerate(sorted(groups.items(),reverse=True)):
  items.sort(key=lambda a:a['top']);n=len(items);strip=Image.new('L',(n*65+30,90),255)
  for z,it in enumerate(items):
   glyph=Image.open(f"tmp/vol1mid/glyphs/{it['gid']}.png").resize((60,60));strip.paste(glyph,(15+z*65,15))
  path=f'tmp/vol1mid/strips/{j+232}-{k}.png';strip.save(path);strips.append({'page':j+232,'col':k,'x':x,'items':items,'path':path})
json.dump(strips,open('tmp/vol1mid/strips.json','w'))
def ocr(s):
 res=subprocess.run(['tesseract',s['path'],'stdout','-l','kor+chi_tra','--psm','7'],capture_output=True,text=True)
 return dict(page=s['page'],col=s['col'],count=len(s['items']),text=res.stdout.strip())
with ThreadPoolExecutor(max_workers=6) as e:
 out=list(e.map(ocr,strips))
json.dump(out,open('tmp/vol1mid/ocr.json','w'),ensure_ascii=False,indent=2)
print('Finished',len(glyphs),'unique glyphs',len(strips),'columns',flush=True)
