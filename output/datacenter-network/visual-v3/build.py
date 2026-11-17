from pathlib import Path
import json,hashlib,zipfile
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from PIL import Image
R=Path(__file__).resolve().parent;E=R.parent
specs=[('01','四者关系',['S2','S6','S8','S11']),('02','RDMA数据路径',['S2','S6','S11']),('03','InfiniBand逐跳流控',['S2','S3']),('04','RoCE双反馈',['S6','S14']),('05','UET多路径与完成',['S4','S10','S11']),('06','图解对照',['S2','S6','S11','S14'])]
sources={x['id']:x['url'] for x in json.loads((E/'sources.json').read_text())};p=Presentation();p.slide_width=Inches(13.333333);p.slide_height=Inches(7.5)
for n,title,ids in specs:
 s=p.slides.add_slide(p.slide_layouts[6]);f=R/'assets'/f'{n}.jpg';w,h=Image.open(f).size;maxh=Inches(7.23);k=min(p.slide_width/w,maxh/h);ww,hh=int(w*k),int(h*k)
 s.shapes.add_picture(str(f),(p.slide_width-ww)//2,(maxh-hh)//2,width=ww,height=hh)
 tb=s.shapes.add_textbox(Inches(.2),Inches(7.27),Inches(12.9),Inches(.2));tf=tb.text_frame;tf.margin_top=tf.margin_bottom=0;q=tf.paragraphs[0];q.text=f'{n}  来源：'+ ' · '.join(ids)+'  |  2026-11-17核查  |  静态机制插画；设备非实机复原；功能与交付语义依具体profile';q.font.name='Noto Sans CJK SC';q.font.size=Pt(9);q.font.color.rgb=RGBColor.from_string('53637B')
 s.notes_slide.notes_text_frame.text=title+'\n'+'\n'.join(i+': '+sources[i] for i in ids)+'\n插画以已核查机制为依据，由图像模型生成并审阅。主图为静态栅格图片，不是逐部件可编辑图，也不是实机照片或机械CAD。设备型号、队列格数、包编号不构成实际BOM或测量。UE按版本/profile解释；乱序按偏移写入仅为一种示意，不宣称所有profile都强制重排交付。PFC示意优先级编号不是部署建议。'
 s.shapes[0].name=title
p.save(R/'exports/数据中心网络_机制插画版.pptx')
(R/'deck-map.json').write_text(json.dumps([{'page':int(n),'title':t,'source_ids':ss,'asset':'assets/'+n+'.jpg'} for n,t,ss in specs],ensure_ascii=False,indent=2))
f=R/'exports/数据中心网络_机制插画版.pptx'
with zipfile.ZipFile(f) as z:assert z.testzip() is None;assert not any(n.endswith(('.gif','.mp4')) for n in z.namelist())
print(f,round(f.stat().st_size/1e6,2),'MB')
