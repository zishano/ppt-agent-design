from pathlib import Path
import json,hashlib,zipfile
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE,MSO_CONNECTOR
from pptx.oxml.xmlchemy import OxmlElement
R=Path(__file__).resolve().parent;O=R/'exports';p=Presentation();p.slide_width=Inches(13.333333);p.slide_height=Inches(7.5)
C={'ink':'172C49','blue':'2164D9','green':'11856D','purple':'7855BF','orange':'D47A22','gray':'61738B','line':'C9D6E6','pale':'EEF4FC','white':'FFFFFF'}
U=lambda n:Inches(n/96)
def rgb(c):return RGBColor.from_string(C.get(c,c))
def box(s,x,y,w,h,fill='white',stroke='line',radius=False):
 sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,U(x),U(y),U(w),U(h));sh.fill.solid();sh.fill.fore_color.rgb=rgb(fill);sh.line.color.rgb=rgb(stroke);sh.line.width=Pt(.8)
 effect=OxmlElement('a:effectLst');sh._element.spPr.append(effect)
 if radius:sh.adjustments[0]=.12
 return sh

def text(s,x,y,w,h,t,size=22,color='ink',bold=False,align=None):
 sh=s.shapes.add_textbox(U(x),U(y),U(w),U(h));tf=sh.text_frame;tf.word_wrap=True;tf.margin_left=tf.margin_right=0;tf.margin_top=tf.margin_bottom=0
 for i,line in enumerate(t.split('\n')):
  q=tf.paragraphs[0] if i==0 else tf.add_paragraph();q.text=line;q.font.name='Noto Sans CJK SC';q.font.size=Pt(size*.75);q.font.bold=bold;q.font.color.rgb=rgb(color)
  if align is not None:q.alignment=align
 return sh

def line(s,x1,y1,x2,y2,color='blue',arrow=False,width=2):
 sh=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,U(x1),U(y1),U(x2),U(y2));sh.line.color.rgb=rgb(color);sh.line.width=Pt(width)
 if arrow:
  z=OxmlElement('a:tailEnd');z.set('type','triangle');sh._element.spPr.get_or_add_ln().append(z)
 return sh

def node(s,x,y,w,h,t,color='blue',size=20):
 box(s,x,y,w,h,'pale',color,True);text(s,x+6,y+8,w-12,h-10,t,size,color,True)
def gpu(s,x,y,label='GPU',color='blue',w=65,h=40):
 box(s,x,y,w,h,'white',color,True);text(s,x+5,y+(3 if h<30 else 7),w-10,h-3,label,13 if h<30 else 17,color,True)
 for i in range(4):
  line(s,x+9+i*(w-18)/3,y-4,x+9+i*(w-18)/3,y,color,False,.8)
  line(s,x+9+i*(w-18)/3,y+h,x+9+i*(w-18)/3,y+h+4,color,False,.8)

def tag(s,x,y,t,color='blue',w=200):node(s,x,y,w,34,t,color,18)
sources={k:(R/'research'/f'{k}.url').read_text() for k in ['S1','S2','S3','S5','S7','S8']}
meta=[]
def slide(title,sub,ids):
 s=p.slides.add_slide(p.slide_layouts[6]);box(s,0,0,1280,720,'white','white');box(s,30,27,6,46,'blue','blue');text(s,49,24,1190,56,title,34,'ink',True);text(s,49,85,1175,35,sub,19,'gray');line(s,32,662,1248,662,'line',False,1)
 n=len(p.slides);text(s,35,678,1165,27,'来源：'+' · '.join(ids)+'  |  资料核查：2026-11-16；拓扑均为逻辑示意，非布线图',13,'gray');text(s,1211,676,40,25,f'{n:02}',16,'blue',True)
 s.notes_slide.notes_text_frame.text='来源与范围\n'+'\n'.join(f'{i}: {sources[i]}' for i in ids)+'\n图为解释性逻辑拓扑，不表示具体产品布线、性能测量或部署规模。'
 meta.append({'page':n,'title':title,'sources':ids,'reading_order':'标题 → 主图 → 分层解释 → 结论','checks':['source_scope','topology','labels','render']});return s
# 1
s=slide('AI 集群需要两层互连：域内紧耦合，域间扩规模','同一个作业可以同时使用 Scale-up 和 Scale-out；边界由互连域决定，不由机箱或机柜外壳决定。',['S2','S3','S8'])
box(s,35,135,795,484,'white','line',True)
for x in [220,510]:node(s,x,157,150,40,'Spine 交换机','green',19)
for x in [115,605]:
 node(s,x,243,140,40,'Leaf 交换机','green',19)
 for z in [295,585]:line(s,x+70,243,z,197,'green',False,1.3)
for n,x in enumerate([64,252,440,628]):
 box(s,x,355,174,224,'pale','blue',True);text(s,x+12,367,150,26,f'Scale-up 域 {n+1}',18,'blue',True)
 node(s,x+37,306,100,31,'NIC / RDMA','green',14);line(s,x+87,306,185 if n<2 else 675,283,'green',False,1.4);line(s,x+87,337,x+87,405,'green',False,1.5)
 node(s,x+17,408,140,35,'域内交换互连','blue',17)
 for a in range(2):
  for b in range(2):
   gx=x+14+b*78;gy=468+a*59;line(s,gx+31,gy,x+87,443,'blue',False,1);gpu(s,gx,gy,f'GPU {a*2+b}',w=65,h=35)
text(s,57,591,750,23,'每域仅画 4 个 GPU 作示意；域大小取决于实际平台。',15,'gray')
for y,k,title,desc,col in [(150,'01','计算端','GPU 执行计算；HBM 保存本地数据。\n多卡协同需要显式通信。','ink'),(307,'02','Scale-up：域内','NVLink / NVSwitch 等紧耦合互连。\n典型承载频繁的张量与激活交换。','blue'),(468,'03','Scale-out：域间','通过 NIC 与交换网络连接多个域。\n典型采用 InfiniBand 或 Ethernet/RoCE。','green')]:
 tag(s,858,y,k+'  '+title,col,365);text(s,860,y+48,365,95,desc,20)
# 2
s=slide('先看通信需求：并行方式决定“谁在何时交换什么”','通信并非只有梯度同步：激活、参数分片与 MoE token 路由也会成为关键路径。',['S5','S1','S7'])
rows=[('TP 张量并行','同一层拆到多卡','层内部分结果合并；通信频繁','blue'),('DP 数据并行','模型副本处理不同批次','梯度聚合；分片优化器还需参数汇聚','green'),('PP 流水线并行','不同 GPU 持有不同层','阶段间传激活，反向传梯度','purple'),('EP 专家并行','专家分散在不同 GPU','Token dispatch / combine；常见 All-to-all','orange')]
for i,(a,b,c,col) in enumerate(rows):
 y=139+i*123;box(s,35,y,1210,109,'white','line',True);text(s,52,y+12,220,32,a,24,col,True);text(s,52,y+53,220,37,b,18,'gray')
 if i<2:
  for j in range(4):gpu(s,306+j*91,y+15,str(j),col,61,34)
  node(s,323,y+67,309,27,'局部结果合并' if i==0 else '梯度同步 / 分片交换',col,15)
  for j in range(4):line(s,337+j*91,y+49,337+j*91,y+67,col,True,1.1)
 elif i==2:
  for j in range(4):
   node(s,294+j*95,y+28,77,43,'层组 '+str(j+1),col,16)
   if j<3:line(s,371+j*95,y+48,389+j*95,y+48,col,True)
  text(s,307,y+78,370,20,'前向 → 激活     反向 ← 梯度',14,col)
 else:
  for j in range(3):node(s,296,y+10+j*28,80,23,'token '+str(j),col,13);node(s,584,y+10+j*28,80,23,'专家 '+str(j),col,13)
  for j in range(3):
   for k in range(3):line(s,376,y+22+j*28,584,y+22+k*28,col,False,.7)
 text(s,727,y+15,490,40,c,21,col,True);text(s,727,y+62,480,37,['关注：单次延迟、频率、通信重叠','关注：消息大小、有效带宽、慢节点','关注：跨阶段延迟、流水线气泡','关注：流量不均、热点、尾延迟'][i],18)
#3
s=slide('同样叫“通信”，集合操作的数据含义并不相同','下图数字只用于解释运算，不是模型参数、性能测试或业务数据。',['S1'])
for x,title,col in [(35,'AllReduce：每卡拿到相同总和','blue'),(651,'All-to-all：按目的地重新分发','purple')]:box(s,x,144,594,410,'white','line',True);tag(s,x+18,160,title,col,558)
for j in range(4):
 x=72+j*138;node(s,x,226,105,48,f'卡 {j}：{j+1}','blue',20);line(s,x+52,274,330,330,'blue',True,1)
node(s,206,331,250,48,'求和：1 + 2 + 3 + 4','blue',20)
for j in range(4):
 x=72+j*138;line(s,330,379,x+52,439,'blue',True,1);node(s,x,439,105,48,'结果：10','blue',20)
text(s,67,510,530,28,'含义：聚合数值，并把完整结果返回所有参与者。',17)
colors=['blue','green','purple','orange']
for j in range(4):
 text(s,674,221+j*62,55,25,'卡 '+str(j),16)
 for k in range(4):box(s,731+k*30,220+j*62,24,28,colors[k],colors[k])
 text(s,1155,221+j*62,60,25,'卡 '+str(j),16)
 for k in range(4):box(s,1021+k*30,220+j*62,24,28,colors[j],colors[j])
 for k in range(4):line(s,852,235+j*62,1014,235+k*62,colors[k],False,.8)
text(s,677,510,540,28,'含义：每卡向不同目的卡发送不同数据块。',17)
box(s,35,574,1210,57,'pale','line',True);text(s,54,589,1176,35,'ReduceScatter → AllGather 可实现 AllReduce：先各持一份归约分片，再汇齐全部分片。',23,'ink',True)
#4
s=slide('Scale-up：把一个互连域做强，让频繁协同留在域内','真实例子：NVIDIA GB200 NVL72 的 NVLink 域覆盖 72 个 Blackwell GPU；并非“只能机内”。',['S2','S5'])
box(s,35,143,700,480,'pale','blue',True);tag(s,55,158,'GB200 NVL72：72-GPU NVLink 域','blue',653)
for row in range(9):
 for col in range(8):gpu(s,63+col*80,215+row*35,str(row*8+col+1),'blue',61,25)
node(s,71,552,624,42,'NVLink Switch System：将域内 GPU 互连','blue',23)
text(s,64,601,653,20,'72 个逻辑端点；网格排列不代表实物托盘位置或交换芯片数量。',14,'gray')
for y,head,body in [(151,'01 连接对象','同一紧耦合域内的 GPU。\n域可跨服务器，并可扩展到整柜。'),(310,'02 典型任务','TP 的层内交换、紧密协同的模型分片。\nEP 能否留在域内取决于专家放置。'),(471,'03 设计边界','域容量、互连拓扑及通信库共同约束。\n不是自动得到一块透明共享显存。')]:
 tag(s,769,y,head,'blue',445);text(s,773,y+50,445,99,body,22)
#5
s=slide('Scale-out：把多个计算域连接成更大的 AI 集群','InfiniBand 与 Ethernet/RoCE 是常见选择；是否跑得快，还取决于拓扑、拥塞与通信实现。',['S3','S8','S5'])
box(s,35,143,775,477,'white','line',True)
for j in range(3):node(s,132+j*202,172,155,42,'Spine '+str(j+1),'green',21)
for i in range(4):
 x=73+i*182
 for j in range(3):line(s,x+65,310,209+j*202,214,'green',False,1.4)
 node(s,x,310,130,43,'Leaf '+str(i+1),'green',21);line(s,x+65,353,x+65,401,'green',False,1.5);node(s,x+11,401,108,36,'RDMA NIC','green',16);line(s,x+65,437,x+65,476,'green',False,1.5)
 box(s,x-7,476,145,98,'pale','blue',True);text(s,x+8,485,123,25,'计算域 '+str(i+1),18,'blue',True)
 for k in range(2):gpu(s,x+7+k*64,527,'GPU','blue',53,28)
text(s,61,591,730,21,'示例为两层 Clos 逻辑拓扑；交换机数、链路数与超售比均需按项目设计。',14,'gray')
for y,head,body in [(153,'01 端点注入能力','GPU → NIC 的路径与 NIC 总带宽\n决定计算域能向网络送出多少数据。'),(310,'02 交换容量与路径','Leaf 上联、路径分布、双向割带宽\n共同影响并发集合通信的吞吐。'),(467,'03 拥塞与慢尾','热点、拥塞控制与路由策略\n决定同步阶段是否被慢流拖住。')]:
 tag(s,840,y,head,'green',392);text(s,845,y+52,390,95,body,21)
#6
s=slide('区别看互连域与通信路径，而不是“快网”和“慢网”','两类网络共同服务一个作业；并行组可以跨域，但通信成本与最优放置会随之变化。',['S2','S3','S5','S8'])
node(s,244,145,470,45,'Scale-up：域内紧耦合','blue',25);node(s,742,145,502,45,'Scale-out：连接多个域','green',25)
data=[('边界','GPU 紧耦合互连域','NIC 与交换网络连接的多个域'),('典型路径','GPU ↔ 域内交换互连 ↔ GPU','GPU ↔ NIC ↔ 网络 ↔ NIC ↔ GPU'),('技术例子','NVLink / NVSwitch','InfiniBand；Ethernet + RoCE'),('典型压力','频繁、小颗粒的层内交换','大规模并发、吞吐与流量热点'),('优先放置','高频 TP；紧耦合分片','跨域 DP；按任务需求承载 PP / EP'),('要检查什么','域内拓扑、延迟、有效带宽','NIC 注入、超售、拥塞与尾延迟')]
for i,(key,a,b) in enumerate(data):
 y=211+i*65;box(s,35,y,1209,57,'pale' if i%2==0 else 'white','line',True);text(s,53,y+15,180,33,key,21,'ink',True);text(s,255,y+15,450,36,a,20,'blue');text(s,755,y+15,475,36,b,20,'green')
text(s,47,616,1190,28,'注意：TP / DP / PP / EP 与两类网络不是一一绑定关系；以上是常见放置思路，不是协议限制。',18,'gray')
#7
s=slide('两层网络如何合作？用一次跨域梯度求和串起来','为了说明层次关系，下图采用“域内归约 → 代表间归约 → 域内广播”；不代表 NCCL 固定算法。',['S1','S5'])
for d,x in enumerate([58,674]):
 box(s,x,148,548,424,'pale','blue',True);tag(s,x+18,165,'Scale-up 域 '+('A' if d==0 else 'B'),'blue',511)
 for j in range(4):gpu(s,x+25+j*130,236,'梯度 '+str(d*4+j+1),'blue',103,45);line(s,x+76+j*130,281,x+274,340,'blue',True,1.2)
 node(s,x+161,340,226,49,'代表：局部 → 全局','blue',21)
 for j in range(4):line(s,x+274,389,x+76+j*130,479,'blue',True,1);node(s,x+23+j*130,480,108,43,'全局结果','blue',17)
line(s,445,364,835,364,'green',True,3);line(s,835,370,445,370,'green',True,2);text(s,501,301,278,52,'代表之间交换并求和\nScale-out 网络',19,'green',True)
# cover line runs through own node? path starts node edge to opposite edge, fine
for x,t,col in [(45,'① 域内聚合：减少跨域参与端点','blue'),(450,'② 域间求和：连接两侧代表','green'),(855,'③ 域内广播：各卡得到总和','blue')]:node(s,x,594,380,42,t,col,18)
#8
s=slide('落地顺序：先确定通信图，再设计网络与并行组','目标是缩短训练步时或推理关键路径；端口标称速率不能直接替代应用效果。',['S5','S8','S1'])
heads=['01 识别通信','02 放置并行组','03 配置网络','04 实测验证'];bodies=['TP / DP / PP / EP\n消息大小 · 频率\n是否需要同步等待','高频交换优先域内\n容量不足时再跨域\n检查热点与专家偏斜','核对 NIC 注入能力\nLeaf / Spine 容量\n拥塞与冗余策略','集合通信微基准\n并发与拥塞测试\n最终验证应用指标']
for i in range(4):
 x=38+i*311;node(s,x,158,280,49,heads[i],['blue','purple','green','orange'][i],25);box(s,x,221,280,158,'pale','line',True);text(s,x+16,242,250,125,bodies[i],22)
 if i<3:line(s,x+280,183,x+305,183,'gray',True,2)
box(s,38,412,1204,205,'white','line',True);text(s,58,429,1155,36,'通信时间的简化模型：启动开销 + 数据传输 + 等待 / 拥塞',27,'ink',True)
for i,(h,b,col) in enumerate([('启动开销','小消息、多次同步时更敏感','purple'),('数据传输','消息量 ÷ 有效带宽；不是端口标称速率','blue'),('等待 / 拥塞','慢节点、热点与队列会放大尾延迟','orange')]):
 x=60+i*393;tag(s,x,487,h,col,365);text(s,x,536,365,65,b,21)
text(s,57,628,1180,23,'验收指标：训练步时、通信重叠、集群扩展效率；推理吞吐、首 token 延迟与逐 token 延迟。',17,'gray')
# save
O.mkdir(exist_ok=True);f=O/'AI集群通信_Scale-up与Scale-out.pptx';p.save(f)
(R/'deck-map.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2))
(R/'sources.json').write_text(json.dumps([{'id':k,'url':v,'accessed':'2026-11-16','local_excerpt':'research/'+k+'.txt','status':'read'} for k,v in sources.items()],ensure_ascii=False,indent=2))
(R/'brief.json').write_text(json.dumps({'title':'AI集群通信需求与Scale-up、Scale-out网络的区别','language':'zh-CN','aspect':'16:9','backend':'standalone python-pptx','pages':8,'diagram_detail_policy':'diagram-rich','audience':'技术与产业读者','scope':'通信需求、互连域、拓扑、并行放置与验证','illustrations':'逻辑示意，不虚构机器实物或性能数据','required':['图为主','区分Scale-up域与物理机柜边界','来源可追溯','静态、原生可编辑']},ensure_ascii=False,indent=2))
with zipfile.ZipFile(f) as z:assert z.testzip() is None;assert not any(n.endswith(('.gif','.mp4')) for n in z.namelist())
print(f, f.stat().st_size)
