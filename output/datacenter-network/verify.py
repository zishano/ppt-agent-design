from pathlib import Path
import json,hashlib,re,zipfile
import pymupdf as fitz
from PIL import Image
from pptx import Presentation
R=Path(__file__).resolve().parent;f=next((R/'exports').glob('*.pptx'));pdf=f.with_suffix('.pdf');p=Presentation(f);d=fitz.open(pdf)
assert len(p.slides)==len(d)==8
norm=lambda s:re.sub(r'\s+','',s)
missing=[];out=Image.new('RGB',(1600,1900),'#e2e8f0')
for i,s in enumerate(p.slides):
 pt=norm(d[i].get_text())
 for sh in s.shapes:
  if sh.has_text_frame and norm(sh.text) and norm(sh.text) not in pt:missing.append({'page':i+1,'text':sh.text})
 fn=R/'renders'/f'{i+1:02}.png';d[i].get_pixmap(matrix=fitz.Matrix(1.25,1.25)).save(fn);im=Image.open(fn);im.thumbnail((790,450));out.paste(im,((i%2)*800,(i//2)*475))
assert not missing,missing
out.save(R/'renders/contact.jpg',quality=90)
with zipfile.ZipFile(f) as z:
 assert z.testzip() is None
 assert not any(n.lower().endswith(('.gif','.mp4','.avi')) for n in z.namelist())
report={'slides':8,'pptx_sha256':hashlib.sha256(f.read_bytes()).hexdigest(),'pdf_sha256':hashlib.sha256(pdf.read_bytes()).hexdigest(),'package':'passed','text_coverage':'passed','editability':'native text and diagrams','static_only':True,'visual_review':'unchecked','renderer':'LibreOffice; not PowerPoint desktop','size_bytes':f.stat().st_size}
(R/'validation/automated.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));print(report)
