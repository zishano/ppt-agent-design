from pathlib import Path
import json,hashlib,re,zipfile
import fitz
from pptx import Presentation
R=Path(__file__).resolve().parent;f=R/'exports/AI集群通信_Scale-up与Scale-out.pptx';pdf=f.with_suffix('.pdf');p=Presentation(f);d=fitz.open(pdf)
assert len(p.slides)==len(d)==8
norm=lambda s:re.sub(r'\s+','',s)
missing=[]
for i,s in enumerate(p.slides):
 pt=norm(d[i].get_text())
 for sh in s.shapes:
  if sh.has_text_frame and norm(sh.text) and norm(sh.text) not in pt:missing.append({'page':i+1,'text':sh.text})
 d[i].get_pixmap(matrix=fitz.Matrix(1.2,1.2)).save(R/'renders'/f'{i+1:02}.png')
assert not missing,missing
with zipfile.ZipFile(f) as z:assert z.testzip() is None
sources=json.loads((R/'sources.json').read_text())
(R/'exports/资料来源.md').write_text('# AI 集群通信：资料来源\n\n核查日期：2026-11-16。图中端点数量除 GB200 NVL72 的 72 GPU 外均用于逻辑示意；没有引用实测性能数据。\n\n'+'\n'.join(f"- {x['id']}: {x['url']}" for x in sources)+'\n\nScale-up / Scale-out 的任务放置为设计解释，不是并行算法必须遵守的协议限制。实际拓扑与集合通信算法取决于硬件和软件实现。\n')
claims=[('C1','GB200 NVL72 connects 72 Blackwell GPUs in an NVLink domain',['S2'],'fact',[1,4,6]),('C2','AllReduce reduces and distributes full result to all ranks',['S1'],'fact',[2,3,7]),('C3','ReduceScatter followed by AllGather is equivalent to AllReduce',['S1'],'fact',[3]),('C4','All-to-all exchanges distinct destination chunks',['S1'],'fact',[2,3]),('C5','DDP aggregates gradients; distributed optimizer uses reduce-scatter and all-gather',['S5'],'fact',[2]),('C6','TP partitions layers; PP partitions stages; EP distributes experts',['S5'],'fact',[2,6]),('C7','InfiniBand supports multiple topologies and congestion/adaptive routing capabilities',['S3'],'fact',[5]),('C8','NCCL networking diagnosis includes InfiniBand and RoCE',['S8'],'fact',[1,5,6,8]),('C9','Keep frequent tightly coupled exchanges within a suitable domain when possible',['S2','S5'],'interpretation',[4,6,8]),('C10','Two-level reduce, representative allreduce, then broadcast yields global reduction',['S1'],'derived',[7])]
(R/'claims.json').write_text(json.dumps([{'id':i,'claim':c,'sources':src,'kind':kind,'pages':pg} for i,c,src,kind,pg in claims],ensure_ascii=False,indent=2))
report={'slides':8,'pptx_sha256':hashlib.sha256(f.read_bytes()).hexdigest(),'pdf_sha256':hashlib.sha256(pdf.read_bytes()).hexdigest(),'package':'passed','text_coverage':'passed','native_shapes_and_text':True,'static_only':True,'visual_review':'pending final changed-page review','renderer':'LibreOffice; not PowerPoint desktop','size_bytes':f.stat().st_size}
(R/'validation/report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));print(report)
