from pathlib import Path
import json,zipfile,hashlib
from PIL import Image
from pptx import Presentation
import fitz
r=Path(__file__).resolve().parent;f=next((r/'exports').glob('*.pptx'));d=fitz.open(next((r/'exports').glob('*.pdf')));p=Presentation(f)
c=Image.new('RGB',(1600,1350),'#E1E7F0')
for i,page in enumerate(d):
 page.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(r/'renders'/f'{i+1:02}.png');im=Image.open(r/'renders'/f'{i+1:02}.png');im.thumbnail((800,450));c.paste(im,(i%2*800,i//2*450))
c.save(r/'renders/contact.jpg',quality=92)
with zipfile.ZipFile(f) as z:
 bad=z.testzip();media=[n for n in z.namelist() if n.startswith('ppt/media/')];dynamic=[n for n in media if n.endswith(('.gif','.mp4','.webm'))]
assert len(p.slides)==len(d)==6 and bad is None and not dynamic
assert all('https://' in s.notes_slide.notes_text_frame.text for s in p.slides)
report={'slides':len(p.slides),'pdf_pages':len(d),'pptx_bytes':f.stat().st_size,'pptx_sha256':hashlib.sha256(f.read_bytes()).hexdigest(),'zip_error':bad,'media_count':len(media),'dynamic_media':dynamic,'source_notes':True,'editability':'raster main illustrations including text and paths; native source footer/notes','render_engine':'LibreOffice','text_check_limit':'image labels require visual review; PDF extraction does not validate raster text','visual_review':'see separate visual-review.json'}
(r/'validation/package-report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));print(json.dumps(report,ensure_ascii=False))
