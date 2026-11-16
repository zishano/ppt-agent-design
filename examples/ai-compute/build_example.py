from pathlib import Path
import json
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE as SH,MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
import argparse
parser=argparse.ArgumentParser(description='Rebuild the six-slide AI-compute diagram example; not a generic title-to-PPT generator.')
parser.add_argument('--output-dir',type=Path,required=True)
parser.add_argument('--facts-file',type=Path,default=Path(__file__).resolve().parent/'research.facts.json')
args=parser.parse_args()
R=args.output_dir.resolve();(R/'A').mkdir(parents=True,exist_ok=True)
P=Presentation();P.slide_width=Inches(13.333333);P.slide_height=Inches(7.5)
C={'ink':'173A5E','blue':'377BE8','teal':'1E9D8A','purple':'8257C8','orange':'E49230','red':'D76270','gray':'6C7F97','line':'CED9E8','pale':'F3F7FC','white':'FFFFFF','lb':'EDF4FF','lt':'EBF7F2','lp':'F4EFFB','lo':'FFF4E8'}
def color(v):return RGBColor.from_string(C.get(v,v))
def rect(s,x,y,w,h,f='white',c=None,kind=SH.RECTANGLE):
 q=s.shapes.add_shape(kind,Inches(x/96),Inches(y/96),Inches(w/96),Inches(h/96));q.fill.solid();q.fill.fore_color.rgb=color(f)
 if c:q.line.color.rgb=color(c);q.line.width=Pt(.7)
 else:q.line.fill.background()
 return q

def txt(s,x,y,w,h,t,sz=20,c='ink',bold=False,center=False):
 q=s.shapes.add_textbox(Inches(x/96),Inches(y/96),Inches(w/96),Inches(h/96));f=q.text_frame;f.word_wrap=False;f.margin_left=f.margin_right=f.margin_top=f.margin_bottom=0
 for i,line in enumerate(t.split('\n')):
  p=f.paragraphs[0] if i==0 else f.add_paragraph();p.text=line;p.font.name='Noto Sans CJK SC';p.font.size=Pt(sz*.75);p.font.bold=bold;p.font.color.rgb=color(c)
  if center:p.alignment=PP_ALIGN.CENTER
 return q

def ln(s,x,y,X,Y,c='line',w=1.6,arrow=False,dash=False):
 q=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x/96),Inches(y/96),Inches(X/96),Inches(Y/96));q.line.color.rgb=color(c);q.line.width=Pt(w*.75)
 if arrow:
  e=OxmlElement('a:tailEnd');e.set('type','triangle');q._element.spPr.get_or_add_ln().append(e)
 if dash:
  from pptx.enum.dml import MSO_LINE_DASH_STYLE
  q.line.dash_style=MSO_LINE_DASH_STYLE.DASH
 return q

def wire(s,pts,c='blue',w=1.6,arrow=True,dash=False):
 for i in range(len(pts)-1):ln(s,*pts[i],*pts[i+1],c,w,arrow and i==len(pts)-2,dash)
def node(s,x,y,w,h,t,c='blue',f='white',sz=19):rect(s,x,y,w,h,f,c);txt(s,x+4,y+(h-sz*1.3)/2,w-8,h,t,sz,c,True,True)
def chip(s,x,y,w,h,c='teal',title='GPU'):
 rect(s,x,y,w,h,'white',c)
 for dx in range(16,int(w-8),19):ln(s,x+dx,y-5,x+dx,y,c);ln(s,x+dx,y+h,x+dx,y+h+5,c)
 for dy in range(16,int(h-8),19):ln(s,x-5,y+dy,x,y+dy,c);ln(s,x+w,y+dy,x+w+5,y+dy,c)
 txt(s,x+10,y+6,w-20,20,title,15,c,True)
def matrix(s,x,y,rows=3,cols=4,c='blue',cell=13,gap=3,sel=None):
 for j in range(rows):
  for k in range(cols):rect(s,x+k*(cell+gap),y+j*(cell+gap),cell,cell,c if sel is None or k in sel else 'pale')
def mini_doc(s,x,y,w=66,h=70,c='blue',lines=3):
 rect(s,x,y,w,h,'white',c)
 for j in range(lines):ln(s,x+9,y+15+j*14,x+w-9,y+15+j*14,c,1)
def band(s,x,y,w,h,title,c='blue',f='lb'):
 rect(s,x,y,w,h,f,c);txt(s,x+12,y+8,w-24,25,title,19,c,True)
def tokens(s,x,y,labels,c='purple',w=28,h=27,gap=6):
 for i,t in enumerate(labels):node(s,x+i*(w+gap),y,w,h,t,c,'white',14)
def phase(s,x,y,n,label,c):
 rect(s,x,y,25,25,c,kind=SH.OVAL);txt(s,x,y+1,25,25,str(n),15,'white',True,True);txt(s,x+34,y-1,450,30,label,20,c,True)
def page(n,title,sub,foot):
 s=P.slides.add_slide(P.slide_layouts[6]);rect(s,0,0,1280,720);rect(s,32,25,5,72,'blue');txt(s,48,23,1100,20,'AI COMPUTE / 计算机制图谱',12,'blue',True);txt(s,48,49,1180,48,title,32,'ink',True);txt(s,48,98,1180,28,sub,17,'gray');ln(s,48,132,1232,132,'line',1)
 txt(s,48,678,1120,28,foot,12,'gray');txt(s,1160,678,75,25,f'A  {n:02}/06',12,'blue');return s
# P01: complete task pipelines with inside states and branch semantics.
s=page(1,'AI 计算与传统数据中心计算的本质差异','关键差别：被改变的状态不同，完成任务的条件不同，依赖关系也不同。','传统以 OLTP 为例；AI 以同步训练和自回归 LLM 为例。结构示意，不代表性能比例。')
for x,c,f,title,k in [(48,'blue','lb','事务处理','状态提交'),(452,'teal','lt','模型训练','参数学习'),(856,'purple','lp','在线推理','序列生成')]:
 band(s,x,156,376,493,title,c,f);txt(s,x+256,167,105,24,k,15,c,False,True)
# transaction input mini code, internal branch and durable state
rect(s,68,205,336,71,'white','blue');txt(s,81,215,300,50,'BEGIN → 业务读写 → COMMIT',17,'blue');txt(s,81,246,300,22,'请求携带：对象 · 操作 · 条件',14,'gray')
ln(s,236,278,236,300,'blue',2,True)
node(s,137,309,198,42,'读取业务状态','blue');ln(s,236,352,236,371,'blue',2,True)
node(s,137,380,198,42,'执行业务操作','blue')
wire(s,[(137,401),(90,401),(90,494)],'red');txt(s,73,445,80,25,'失败',15,'red')
node(s,70,505,130,41,'整体回滚','red','white',18)
wire(s,[(335,401),(373,401),(373,494)],'blue');txt(s,330,445,80,25,'成功',15,'blue')
node(s,254,505,140,41,'持久化提交','blue')
wire(s,[(136,548),(136,572),(237,572)],'blue',1,False);wire(s,[(324,548),(324,572),(237,572)],'blue',1,False)
rect(s,154,589,164,39,'white','blue');txt(s,165,596,145,25,'完整业务结果',18,'blue',True,True)
# training
rect(s,472,205,336,71,'white','teal');tokens(s,487,223,['批 A','批 B','批 C'],'teal',82,32,14)
ln(s,640,279,640,300,'teal',2,True)
for x,t in [(476,'前向'),(590,'损失'),(704,'反向')]:node(s,x,318,99,46,t,'teal')
ln(s,578,341,586,341,'teal',1,True);ln(s,693,341,701,341,'teal',1,True)
for x in [483,593,703]:matrix(s,x,393,2,5,'teal',13,3)
wire(s,[(753,366),(753,384)],'teal');txt(s,487,437,290,25,'模型副本 → 梯度 → 同步',17,'teal',False,True)
node(s,527,492,226,46,'同步后更新权重','teal','white',20)
ln(s,640,541,640,565,'teal',2,True);node(s,538,585,204,42,'可部署模型','teal')
wire(s,[(753,515),(791,515),(791,293),(520,293),(520,310)],'orange',1.3);txt(s,664,296,104,19,'下一批',12,'orange')
# inference
rect(s,876,205,336,71,'white','purple');tokens(s,893,226,['提','示','词','…'],'purple',55,29,17)
ln(s,1044,279,1044,300,'purple',2,True)
node(s,925,311,238,43,'Prefill · 处理输入','purple');ln(s,1044,355,1044,375,'purple',2,True)
node(s,925,387,238,43,'Decode · 逐步生成','purple')
rect(s,879,470,148,71,'white','purple');matrix(s,895,487,2,5,'purple',12,3);txt(s,894,520,120,21,'共享模型权重',14,'purple')
rect(s,1052,470,155,71,'white','teal');tokens(s,1061,483,['K','V','…'],'teal',36,25,7);txt(s,1070,520,120,21,'请求 KV 状态',14,'teal')
wire(s,[(954,470),(954,442),(988,442),(988,432)],'purple');wire(s,[(1129,470),(1129,442),(1100,442),(1100,432)],'teal')
tokens(s,901,594,['t1','t2','t3','…'],'purple',62,32,16);wire(s,[(1163,409),(1220,409),(1220,566),(1044,566),(1044,588)],'purple')
# P02: DDP complete lane path, gradient tile and synch points.
s=page(2,'训练：每张卡算自己的数据，却共同更新同一模型','以同步数据并行为例：数据切分、前向/反向、梯度同步与参数更新连成闭环。','DDP 逻辑图；同步可与反向计算重叠。集合通信不是中央参数服务器，也不指定物理拓扑。')
phase(s,48,153,1,'数据进入各副本','blue');phase(s,422,153,2,'本地前向与反向','teal');phase(s,925,153,3,'梯度同步','purple')
for y,c,name in [(206,'blue','进程 A'),(335,'teal','进程 B'),(464,'purple','进程 C')]:
 rect(s,48,y,1043,108,'pale',c);txt(s,61,y+11,90,23,name,17,c,True)
 tokens(s,64,y+47,['样本','…'],c,38,31,6)
 ln(s,157,y+62,190,y+62,c,2,True);chip(s,203,y+18,150,76,c,'模型副本')
 matrix(s,214,y+47,2,4,c,12,3);txt(s,286,y+50,58,25,'权重',15,c)
 ln(s,363,y+62,398,y+62,c,2,True)
 node(s,409,y+38,100,43,'前向',c);ln(s,512,y+60,541,y+60,c,1.5,True)
 node(s,552,y+38,100,43,'损失',c);ln(s,655,y+60,684,y+60,c,1.5,True)
 node(s,695,y+38,100,43,'反向',c);ln(s,798,y+60,827,y+60,c,1.5,True)
 matrix(s,842,y+40,2,4,c,15,3);txt(s,930,y+44,100,26,'局部梯度',17,c)
 ln(s,1044,y+60,1120,y+60,c,2)
# bus and sync boundary
ln(s,1120,266,1120,524,'purple',3)
rect(s,1100,342,132,98,'lp','purple');txt(s,1110,355,112,28,'集合通信',21,'purple',True,True);txt(s,1109,391,115,35,'同步梯度',17,'purple',False,True)
wire(s,[(1166,443),(1166,601),(277,601)],'orange',2,False)
for y in [206,335,464]:wire(s,[(277,601),(183,601),(183,y+25),(201,y+25)],'orange',1.4)
rect(s,374,584,390,36,'white');txt(s,383,588,370,30,'同步后的梯度 → 各副本更新参数',19,'orange',True,True)
# compact math-like local detail, illustrative not numeric
rect(s,48,631,1184,30,'lb');txt(s,62,635,1150,25,'输入不同  /  模型结构相同  /  同步后的更新一致  /  下一批重复',16,'ink',False,True)
# P03: request scheduling and KV life cycle.
s=page(3,'推理：多条请求，共享计算，却各自保留上下文','将“输入处理—连续生成—缓存占用”放在同一张服务流程图里看。','以 vLLM 的 chunked prefill 为例；调度逻辑示意，不表示固定批大小、时延或缓存块容量。')
band(s,48,155,250,470,'① 请求与输入','blue','lb');band(s,326,155,510,470,'② 调度与模型执行','purple','lp');band(s,864,155,368,470,'③ 请求状态与输出','teal','lt')
for y,c,t in [(214,'blue','请求 A · 新输入'),(333,'purple','请求 B · 生成中'),(452,'orange','请求 C · 新输入')]:
 txt(s,64,y,220,25,t,17,c,True);tokens(s,64,y+36,(['t'] if c=='purple' else ['t1','t2','t3','…']),c,42,28,9)
 txt(s,64,y+73,215,23,'待处理 token' if c!='purple' else '下一个 token',14,'gray')
# input arrows attach scheduler (A chunks, B decode, C chunks)
for y in [264,383,502]:wire(s,[(287,y),(310,y),(310,280),(344,280)],'blue',1.2)
rect(s,350,216,461,112,'white','purple');txt(s,367,233,427,31,'调度器',23,'purple',True,True)
txt(s,371,269,415,26,'先放 Decode，再用余量接入 Prefill 分块',17,'purple',False,True)
txt(s,371,300,415,22,'p：输入分块 · d：生成步',14,'purple',False,True)
ln(s,580,329,580,354,'purple',2,True)
rect(s,350,365,461,79,'white','purple');tokens(s,363,385,['B:d','A:p','A:p','C:p','…'],'purple',74,35,13)
ln(s,580,445,580,470,'purple',2,True)
chip(s,350,485,461,115,'purple','同一次模型执行')
node(s,368,525,115,44,'读取权重','purple','white',17);node(s,520,525,115,44,'张量计算','purple','white',17);node(s,672,525,120,44,'更新 KV','teal','white',17)
ln(s,486,547,512,547,'purple',1.5,True);ln(s,638,547,665,547,'purple',1.5,True)
# KV memory and output rows
rect(s,884,211,328,189,'white','teal');txt(s,899,221,294,26,'KV cache：按请求保留',19,'teal',True)
for y,c,tag,n in [(266,'blue','A',4),(307,'purple','B',5),(348,'orange','C',3)]:
 txt(s,899,y+3,32,25,tag,18,c,True)
 for i in range(n):rect(s,938+i*46,y,35,25,c)
wire(s,[(800,550),(849,550),(849,307),(878,307)],'teal',1.5)
wire(s,[(884,351),(847,351),(847,574),(813,574)],'teal',1.5)
for y,c,tag,vals in [(443,'blue','A',['t1']),(498,'purple','B',['t1','t2','t3']),(553,'orange','C',['t1'])]:
 txt(s,884,y+3,33,25,tag,18,c,True);tokens(s,922,y,vals,c,56,29,13)
wire(s,[(580,605),(580,617),(1048,617),(1048,598)],'purple',1.5)
rect(s,48,638,1184,25,'white');txt(s,58,638,1150,24,'缓存不足 → 抢占 / 重算 → 端到端时延可能变差；延迟与吞吐要一起看',17,'red',False,True)
# P04: storage anatomy and ownership
s=page(4,'同样是“存数据”，三类负载保留的是不同状态','共享状态、局部状态和持久化状态分开看，才知道容量压力来自哪里。','定性内存剖面；块大小不代表容量。AI 可运行于 CPU，图中 GPU 仅表示一种部署选择。')
for x,c,f,t in [(48,'blue','lb','事务 / 业务状态'),(452,'teal','lt','训练 / 更新状态'),(856,'purple','lp','推理 / 请求状态')]:band(s,x,157,376,491,t,c,f)
# OLTP state path
node(s,72,211,328,43,'请求：读 / 写 / 条件','blue','white',19)
ln(s,236,257,236,282,'blue',2,True);chip(s,72,294,328,214,'blue','处理器与工作内存')
node(s,89,335,130,52,'业务运算','blue','white',18);node(s,251,335,130,52,'当前状态','blue','white',18);ln(s,224,361,245,361,'blue',1.5,True)
rect(s,89,410,292,72,'lb','blue');txt(s,103,421,260,25,'事务内：多步操作作为整体',17,'blue');txt(s,103,451,260,24,'成功提交 / 失败回滚',16,'blue')
ln(s,236,515,236,543,'blue',2,True);node(s,89,557,292,57,'持久化业务结果','blue','white',21)
# train memory rows
chip(s,476,211,328,344,'teal','训练设备：一个模型副本')
for y,t,c in [(256,'模型权重 W','teal'),(341,'前向中间结果','blue'),(426,'梯度 dW','purple')]:
 rect(s,494,y,292,66,'white',c);txt(s,505,y+9,145,24,t,17,c,True);matrix(s,655,y+12,2,5,c,15,3)
wire(s,[(642,324),(642,337)],'teal');wire(s,[(642,409),(642,422)],'blue')
wire(s,[(787,459),(821,459),(821,277),(788,277)],'orange');txt(s,480,584,320,30,'梯度同步后，更新权重',20,'teal',True,True)
# infer sharing
chip(s,880,211,328,344,'purple','推理设备：共享权重 + 请求缓存')
rect(s,899,254,290,68,'lp','purple');txt(s,911,265,136,26,'共享权重 W',18,'purple',True);matrix(s,1070,269,2,6,'purple',13,3)
for y,c,t,n in [(345,'blue','请求 A',4),(405,'teal','请求 B',5),(465,'orange','请求 C',3)]:
 txt(s,900,y+3,95,25,t,17,c,True)
 for k in range(n):rect(s,998+k*36,y,27,28,c)
 txt(s,900,y+28,80,19,'KV 状态',12,'gray')
node(s,911,577,265,38,'并发 / 上下文 → 占用变化','purple','white',18)
# P05 three parallel policies plus combined system picture
s=page(5,'“多加几张卡”之前，先分清三种并行','切张量、切层、复制模型：分担的对象不同，通信路径也不同。','并行方式为逻辑示意。DP 训练需梯度同步；DP 推理可由不同副本承接不同请求。')
for x,c,f,t in [(48,'blue','lb','张量并行 TP'),(452,'teal','lt','流水线并行 PP'),(856,'purple','lp','数据并行 DP')]:band(s,x,157,376,324,t,c,f)
# TP split weight columns plus devices
rect(s,161,208,152,77,'white','blue');matrix(s,172,220,3,4,'blue',15,4);matrix(s,251,220,3,2,'teal',15,4);ln(s,241,211,241,282,'orange',2)
wire(s,[(200,287),(200,315),(146,315),(146,337)],'blue');wire(s,[(273,287),(273,315),(325,315),(325,337)],'teal')
node(s,82,346,132,54,'GPU / 分片 A','blue','white',17);node(s,260,346,132,54,'GPU / 分片 B','teal','white',17);ln(s,219,372,253,372,'orange',2,True)
txt(s,89,430,302,30,'同一层内分担，结果需协作',18,'blue',False,True)
# PP layers stack and blocks
for y,t in [(211,'前段层'),(248,'后段层')]:node(s,573,y,132,31,t,'teal','white',16)
wire(s,[(639,283),(639,314),(546,314),(546,337)],'teal')
node(s,482,346,132,54,'GPU / 前段','teal','white',18);node(s,664,346,132,54,'GPU / 后段','teal','white',18);ln(s,619,372,658,372,'teal',2,True)
txt(s,488,430,302,30,'层间传递中间结果',18,'teal',False,True)
# DP replicas mini complete models
for x,t in [(884,'输入 A'),(1064,'输入 B')]:
 tokens(s,x+5,211,[t],'purple',111,31);ln(s,x+62,246,x+62,275,'purple',1.5,True)
 rect(s,x,286,132,115,'white','purple');txt(s,x+7,296,118,25,'完整模型副本',16,'purple',True,True);matrix(s,x+22,334,3,5,'purple',14,3)
txt(s,889,430,302,30,'不同输入，模型各自复制',18,'purple',False,True)
# System integration lower picture shows nodes and inside GPU membership
phase(s,48,508,4,'组合部署：节点内与节点间都有通信边界','orange')
for x,t in [(75,'节点 A'),(560,'节点 B')]:
 band(s,x,549,390,104,t,'orange','lo');node(s,x+101,580,116,44,'层 / 分片','teal','white',17);node(s,x+251,580,116,44,'层 / 分片','teal','white',17);ln(s,x+221,602,x+247,602,'teal',2,True)
ln(s,470,603,553,603,'orange',3,True);txt(s,984,565,228,65,'容量分担\n＋ 通信代价',23,'orange',True,True)
# P06 rich end-to-end decision trace with metric micrographs
s=page(6,'配置判断：沿着工作负载，一直追到交付结果','先识别图中的约束，再选择配置；不要让 GPU 型号代替系统分析。','定性诊断框架，不是性能预测。需固定任务、模型、质量、时延与成本口径后实测。')
phase(s,48,153,1,'工作负载','blue');phase(s,361,153,2,'约束位置','purple');phase(s,706,153,3,'可调手段','teal');phase(s,1024,153,4,'交付结果','orange')
for y,c,t,mid,opts,out in [
 (211,'blue','事务','状态与提交',['业务并发','数据访问'],'正确提交'),
 (344,'teal','训练','计算与同步',['批量 / 切分','并行策略'],'任务质量'),
 (477,'purple','推理','容量与时延',['KV 容量','请求调度'],'有效吞吐')]:
 rect(s,48,y,1184,110,'pale');node(s,65,y+26,161,53,t,c,'white',22)
 # visible sub-items, not decorative icons
 if t=='事务':mini_doc(s,242,y+20,62,71,c)
 elif t=='训练':matrix(s,241,y+26,3,4,c,14,4)
 else:tokens(s,237,y+39,['t','t','…'],c,22,25,4)
 ln(s,314,y+54,354,y+54,c,2,True);node(s,367,y+25,232,55,mid,c,'white',20)
 ln(s,607,y+54,673,y+54,c,2,True)
 node(s,686,y+16,261,34,opts[0],c,'white',17);node(s,686,y+60,261,34,opts[1],c,'white',17)
 wire(s,[(954,y+33),(982,y+33),(982,y+77),(954,y+77)],c,1.2,False);ln(s,982,y+54,1013,y+54,c,2,True)
 node(s,1027,y+25,183,55,out,c,'white',21)
wire(s,[(1120,590),(1120,626),(173,626),(173,593)],'orange',1.8)
rect(s,336,611,612,35,'white');txt(s,348,615,588,25,'同条件测量：质量 · 延迟 · 吞吐 · 功率 · 成本',18,'orange',True,True)
# Notes are complete source-backed explanations; page geometries are qualitative.
notes=[
'三个工作流按状态变化对照。事务中多步操作要满足原子性并持久化提交；失败回滚。训练从各自数据经过前向、损失、反向计算梯度，并经同步更新权重；示意格子表示状态，不表示真实维度。推理以已有模型权重处理提示词，再自回归生成输出；请求 KV 状态独立保留。传统以 OLTP 为例，不覆盖所有 HPC 或分析系统。',
'同步数据并行 DDP 中各进程持有模型副本，使用各自数据计算梯度。反向计算中的集合通信同步梯度，各副本依据同步后梯度更新参数。橙色回路表达各副本的更新，不是数据从一台中央参数服务器下发；同步可以与反向计算重叠。',
'以 vLLM chunked prefill 调度为例。图中的 p、d 分别表示 prefill 分块和 decode token，方格数量仅用于辨认请求。先调度 decode，在剩余 token 预算下调度 prefill，较长 prefill 可切块。请求共享模型计算和权重，但保留各自 KV 状态。空间不足时可抢占并重算，从而损害端到端时延。输出数量不表示性能。',
'事务图表示业务操作和持久化结果；训练图展示一个模型副本中的权重、前向中间结果和梯度，强调更新依赖；推理图展示共享模型权重与按请求保留的 KV 状态。并发和上下文长度改变占用。图是逻辑剖面，省略运行时与临时缓冲等，块面积不是容量；CPU 同样可运行 AI。',
'张量并行在同层内切分权重和计算，需要协作。流水线并行把模型层分到不同设备，传递中间结果。数据并行复制模型承接不同数据；训练时通常有梯度同步，推理副本之间可独立处理请求。下方只说明跨节点和节点内的组合边界，不指定某款设备和实际通信拓扑。',
'把事务、训练和推理沿着约束、配置手段与交付结果串起来。箭头表达分析路径，不承诺某项配置必然提高性能。比较结果要锁定任务、模型、质量、时延与成本范围，在同条件下观察容量、同步、吞吐、功率与成本。没有实测，不生成普适加速倍率或投资回报。']
facts=json.loads(args.facts_file.read_text());factids=[['F001','F002','F004','F006','F009','F013'],['F004','F005','F012'],['F006','F007','F008','F009','F015'],['F001','F002','F004','F009','F013','F018'],['F005','F010','F011','F012'],['F001','F004','F014','F015','F016','F017']]
for i,s in enumerate(P.slides):
 src=[f for f in facts['facts'] if f['fact_id'] in factids[i]]
 s.notes_slide.notes_text_frame.text=notes[i]+'\n\n'+'\n\n'.join(f['fact_id']+' '+f['source_title']+'\n'+f['source_url']+'\n'+f['source_locator'] for f in src)
 for sh in s.shapes:
  for el in sh._element.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}style'):sh._element.remove(el)
P.save(R/'A/A_分层机制图.pptx');(R/'A/notes.json').write_text(json.dumps(notes,ensure_ascii=False,indent=2));(R/'A/fact-map.json').write_text(json.dumps(factids,indent=2));print('A: six detailed system diagrams written')
