from pathlib import Path
import json,hashlib,zipfile
from PIL import Image
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
R=Path(__file__).resolve().parent
sources={v['id']:v for v in json.loads((R.parent/'sources.json').read_text())}
page_specs=json.loads((R/'prompts.json').read_text())['pages']
p=Presentation();p.slide_width=Inches(13.333333);p.slide_height=Inches(7.5)
for spec in page_specs:
 n=spec['n'];im=Image.open(R/'assets'/f'{n}.png').convert('RGB');im.save(R/'assets'/f'{n}.jpg',quality=93,optimize=True,subsampling=0)
 s=p.slides.add_slide(p.slide_layouts[6]);w,h=im.size;available=Inches(7.20);k=min(p.slide_width/w,available/h);ww,hh=round(w*k),round(h*k)
 sh=s.shapes.add_picture(str(R/'assets'/f'{n}.jpg'),(p.slide_width-ww)//2,(available-hh)//2,width=ww,height=hh);sh.name=spec['title']
 tb=s.shapes.add_textbox(Inches(.18),Inches(7.28),Inches(12.95),Inches(.19));tf=tb.text_frame;tf.margin_top=tf.margin_bottom=0
 q=tf.paragraphs[0];q.text=f"{n}  来源："+' · '.join(spec['sources'])+'  |  资料核查 2026-11-17  |  静态机制示意；主图不可逐部件编辑';q.font.name='Noto Sans CJK SC';q.font.size=Pt(9);q.font.color.rgb=RGBColor.from_string('53637B')
 s.notes_slide.notes_text_frame.text=spec['title']+'\n'+'\n'.join(k+': '+sources[k]['url'] for k in spec['sources'])+'\n主图为完整静态生成式技术插画，经语义及视觉复核。概念设备非实机复原；数量、路径、缓冲格和几何不构成产品BOM或性能数据。GPU显存直达假定内存授权且平台支持。UET具体交付与乱序语义依版本/profile；规范能力不等于产品实现。图内文字不可逐项编辑。\n内容设计：\n'+spec['prompt']
f=R/'exports/数据中心网络_完整机制插画版_v6.pptx';p.save(f)
with zipfile.ZipFile(f) as z:
 assert z.testzip() is None
 assert len([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')])==6
 assert not any(n.endswith(('.gif','.mp4','.webm')) for n in z.namelist())
print(f.name, f.stat().st_size)
