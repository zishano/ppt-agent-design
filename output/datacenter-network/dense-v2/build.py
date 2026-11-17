from pathlib import Path
import json,hashlib,zipfile
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE,MSO_CONNECTOR
from pptx.oxml.xmlchemy import OxmlElement
R=Path(__file__).resolve().parent;E=R.parent;O=R/'exports';p=Presentation();p.slide_width=Inches(13.333333);p.slide_height=Inches(7.5)
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
from PIL import Image
from pptx.oxml.ns import qn
sources={x['id']:x['url'] for x in json.loads((E/'sources.json').read_text())};meta=[]
def page(title,sub,ids):
 s=p.slides.add_slide(p.slide_layouts[6]);box(s,0,0,1280,720,'white','white');box(s,27,22,6,44,'blue','blue');text(s,46,19,1204,48,title,33,'ink',True);text(s,46,76,1190,34,sub,19,'gray');line(s,28,670,1252,670,'line',False,1);text(s,30,682,1180,24,'来源：'+' · '.join(ids)+' | 核查 2026-11-17 · 图中路径、包序与缓冲格均为机制示意',13,'gray');text(s,1222,680,40,25,str(len(p.slides)),16,'blue',True)
 s.notes_slide.notes_text_frame.text='\n'.join(k+': '+sources[k] for k in ids)+'\n高信息密度修订版；机制图非实测数据。标准能力、实际产品实现与配置条件分别解释。'
 meta.append({'page':len(p.slides),'title':title,'sources':ids});return s
def panel(s,x,y,w,h,title,col='blue'):
 box(s,x,y,w,h,'white','line',True);box(s,x,y,w,34,'pale','line');text(s,x+10,y+6,w-20,27,title,20,col,True)
def note(s,x,y,w,t,col='gray',size=20):text(s,x,y,w,53,t,size,col)
def flow(s,xs,y,labels,col='blue',w=130):
 for i,x in enumerate(xs):
  node(s,x,y,w,42,labels[i],col,20)
  if i<len(xs)-1:line(s,x+w,y+21,xs[i+1],y+21,col,True,2)
def mini(s,x,y,w,h,t,col='blue',size=20):node(s,x,y,w,h,t,col,size)
#1
s=page('技术地图：先分清基础网络、RDMA 传输与完整互连体系','四者既有依赖关系，也有替代路径；不能当作四种并列网卡，更不能只按名称比较“快慢”。',['S2','S6','S8','S11','S14'])
xs=[30,344,658,972];ws=[298]*4;heads=['Ethernet｜网络基础','InfiniBand｜独立体系','RoCEv2｜RDMA 路径','Ultra Ethernet｜增强架构'];cols=['blue','orange','green','purple']
stacks=[['应用 / Sockets','TCP 或 UDP','IP','Ethernet MAC / PHY'],['RDMA / 消息应用','IB 传输层','IB 网络 / 链路','IB 物理互连'],['RDMA 应用','IB 传输语义','UDP / IP','Ethernet MAC / PHY'],['AI / HPC 应用','UET：RDMA 传输','承载与链路增强','Ethernet 基础']]
for x,h,c,st in zip(xs,heads,cols,stacks):
 panel(s,x,124,278,338,h,c)
 for j,label in enumerate(st):
  mini(s,x+15,173+j*63,248,43,label,c,21)
  if j<3:line(s,x+139,216+j*63,x+139,236+j*63,c,True,1.5)
rows=[('回答的问题',['帧如何通过链路传输','端到端互连如何协同','RDMA 如何经过 IP 网络','如何优化 AI/HPC 以太网']),('关键对象',['NIC · MAC/PHY · 交换网络','HCA · IB Switch · 管理体系','RNIC · UDP 4791 · 队列配置','UET 端点 · profile · 软件']),('不可混淆',['Ethernet 不自动保证 RDMA','IB 不是 Ethernet 上层协议','UDP 是封装，不负责可靠性','UET 不是 RoCE 的简单改名'])]
for i,(h,vals) in enumerate(rows):
 y=478+i*58
 for x,c,v in zip(xs,cols,vals):
  text(s,x+3,y,270,24,h,16,c,True);text(s,x+3,y+23,277,35,v,19)
text(s,33,651,1210,20,'阅读规则：先确认所在层级，再看端点能力、传输语义、网络配置和产品版本。',16,'gray')
#2
s=page('Ethernet + RDMA：同一拓扑下，数据路径与控制路径不同','RDMA = Remote Direct Memory Access（远程直接内存访问）；下面以已授权内存的一次 Write 为例。',['S2','S6','S11','S14'])
panel(s,30,122,714,285,'① 整体网络：拓扑只说明连接，不证明协议能力')
for x in [227,489]:mini(s,x,173,137,40,'Spine','blue',20)
for i,x in enumerate([58,229,400,571]):
 for xx in [295,557]:line(s,x+65,265,xx,213,'blue',False,1)
 mini(s,x,265,130,39,'Leaf '+str(i+1),'blue',19);line(s,x+65,304,x+65,331,'green',False,2);mini(s,x,331,130,39,'NIC → 节点','green',18)
text(s,47,381,680,24,'Leaf 接入节点；Spine 连接 Leaf。端口数量、超售及路径需另行设计。',17,'gray')
panel(s,761,122,489,285,'② 术语与职责：同一“网卡”不等于同一能力','green')
for y,h,b in [(169,'NIC / RNIC / HCA','网络接口 / RDMA 网卡 / IB 主机通道适配器'),(246,'内存注册与权限','确定可访问区域、访问方式及端点资源'),(323,'CPU / GPU 内存','GPU 显存直达还需设备、驱动和软件支持')]:
 text(s,777,y,453,30,h,23,'green',True);text(s,777,y+34,453,47,b,20)
panel(s,30,424,824,230,'③ 局部展开：CPU 准备资源，网卡执行搬运')
for x,h in [(47,'发送主机'),(565,'接收主机')]:
 text(s,x,468,260,25,h,18,'gray',True);mini(s,x,499,135,40,'CPU / 注册','gray',18)
 if x==47:mini(s,x,580,135,42,'授权内存','blue',20)
mini(s,210,580,115,42,'RDMA NIC','green',17);mini(s,565,580,115,42,'RDMA NIC','green',17)
# receiver memory shifts to right, CPU registration above
# hide previously placed receiver memory box with correct semantic NIC above it (same bounds)
mini(s,724,580,112,42,'授权内存','blue',19)
line(s,182,601,210,601,'blue',True,2);line(s,325,601,565,601,'green',True,2);text(s,382,565,151,28,'网络传输',19,'green',True);line(s,680,601,724,601,'blue',True,2)
line(s,114,539,267,558,'gray',False,1.5);line(s,267,558,267,580,'gray',True,1.5);line(s,632,539,622,580,'gray',True,1.5)
panel(s,871,424,379,230,'④ 判断边界','orange');text(s,886,474,349,165,'RDMA 不等于 CPU 完全退出。\n\n拓扑相同不代表端点互通。\n\n网卡速率不等于应用有效带宽。',22)
#3
s=page('InfiniBand：把 RDMA、交换 Fabric 与流控作为一套系统理解','HCA 处理端点通信；交换网络转发；逐跳反馈保护缓冲。无损机制不消除排队、故障或应用同步等待。',['S2','S3','S11'])
panel(s,30,122,817,280,'① 数据前进，流控反馈逐跳返回','orange')
flow(s,[53,258,463,668],195,['HCA A','IB Switch 1','IB Switch 2','HCA B'],'orange',152)
for a,b in [(258,205),(463,410),(668,615)]:line(s,a,240,a,279,'green',False,1.6);line(s,a,279,b,279,'green',False,1.6);line(s,b,279,b,240,'green',True,1.6)
text(s,52,307,766,37,'发送端 → 交换网络 → 接收端：承载消息及远端内存操作',22,'orange',True);text(s,52,350,766,37,'反馈针对相邻链路，不是接收应用给整个网络发“暂停”。',21,'green')
panel(s,865,122,385,280,'② 三个层级，各自负责什么','orange')
for y,h,b in [(169,'端点 HCA','执行通信操作与数据搬运'),(244,'交换 Fabric','转发与路径组织，连接端点'),(319,'运维体系','配置、监控、故障与兼容验证')]:text(s,881,y,353,28,h,23,'orange',True);text(s,881,y+32,353,40,b,20)
panel(s,30,419,518,235,'③ 缓冲局部：拥塞时为什么需要减速','green')
for j in range(9):box(s,49+j*39,482,28,35,'orange' if j<6 else 'white','orange')
text(s,51,532,469,88,'接收空间紧张 → 限制上游继续发送\n空间恢复 → 允许继续传输\n缓冲格只示意占用，不代表实际队列深度。',20)
panel(s,565,419,685,235,'④ 和 Ethernet / RoCE 对照时，避免三个误判')
for y,h,b in [(471,'不是 PFC 的别名','IB 使用自己的链路流控机制；不要套用以太网配置。'),(528,'无损仍会拥塞','可能不丢包，但仍会排队、反压，拖长同步任务尾延迟。'),(585,'别只看速率','还要检查端点、交换路径、故障恢复及应用通信模式。')]:text(s,581,y,165,32,h,21,'blue',True);text(s,754,y,477,53,b,20)
#4
s=page('RoCEv2：一页读懂封装、反馈回路与部署条件','RoCE = RDMA over Converged Ethernet；可靠交付由 RDMA 传输协议提供，UDP 只是封装层。',['S6','S14'])
for x,w,t,col in [(30,193,'Ethernet','blue'),(229,151,'IP','blue'),(386,234,'UDP · 4791','green'),(626,332,'RDMA 传输头','green'),(964,286,'数据 / 校验','purple')]:mini(s,x,123,w,45,t,col,23)
text(s,34,179,1202,28,'v1：二层 Ethernet 封装，EtherType 0x8915；v2：IP/UDP 封装，可路由，UDP 目的端口为 4791。',20,'gray')
panel(s,30,220,801,278,'① 数据向前传，ECN/CNP 与 PFC 分别反馈','green')
flow(s,[48,254,460,664],286,['发送 RNIC','交换机 A','交换机 B','接收 RNIC'],'green',146)
text(s,410,254,170,28,'拥塞点标记 ECN',18,'orange',True)
line(s,736,328,736,437,'green',False,2);line(s,736,437,121,437,'green',False,2);line(s,121,437,121,328,'green',True,2)
text(s,229,449,560,28,'CNP：接收端通知发送端，调整发送速率',20,'green',True)
line(s,460,328,460,377,'orange',False,2);line(s,460,377,400,377,'orange',False,2);line(s,400,377,400,328,'orange',True,2)
text(s,290,390,340,25,'PFC：相邻链路的指定优先级暂停',18,'orange',True)
panel(s,848,220,402,278,'② 术语与作用范围')
for y,h,b,c in [(265,'ECN｜显式拥塞通知','交换节点对经过的 IP 包作拥塞标记。','blue'),(340,'CNP｜拥塞通知包','接收端反馈源端；形成端到端调速闭环。','green'),(415,'PFC｜优先级流量控制','暂停相邻上游指定优先级，不是所有流量。','orange')]:text(s,861,y,373,29,h,20,c,True);text(s,861,y+30,373,43,b,19)
for x,w,h,t,c in [(30,394,'③ 端点条件','RoCE RNIC + 软件支持\n内存注册、权限和端点配置\n并非普通 NIC 开开关就具备 RDMA','green'),(441,394,'④ 交换网络条件','优先级与队列映射一致\nECN 阈值、反馈和缓冲需配合\nPFC 必须按部署策略配置','blue'),(852,398,'⑤ 常见误区','RoCE 存在有损 / 无损部署模式\nPFC 不是端到端拥塞控制的替代品\n“无损”也可能有反压和长尾','orange')]:panel(s,x,516,w,138,h,c);text(s,x+12,560,w-24,86,t,19)
#5
s=page('Ultra Ethernet：从 UET 传输到 UE 1.1 的分层增强','UEC 是组织；Ultra Ethernet 是规范架构；UET 是其传输层。标准定义的能力仍需产品与软件实现。',['S4','S10','S11'])
panel(s,30,122,642,296,'① 主机制：多路径 + 可靠交付 + 拥塞控制','purple')
for y in [205,274,343]:line(s,180,289,284,y+20,'purple',False,1.5);line(s,434,y+20,521,289,'purple',False,1.5);mini(s,284,y,150,40,'路径 '+str([205,274,343].index(y)+1),'purple',20)
mini(s,47,263,133,51,'UET 发送端','purple',19);mini(s,521,263,133,51,'UET 接收端','purple',19)
text(s,49,382,601,28,'包级分散、乱序交付、重传与完成语义须协同处理。',19,'gray')
panel(s,689,122,561,296,'② 分层理解：不能只更换交换机','purple')
for y,h,b in [(169,'软件 / 接口','应用提交操作；定义资源、内存与完成语义'),(238,'UET 传输','可靠交付、多路径、拥塞管理与安全等机制'),(307,'链路 / Fabric','匹配所选 profile 的承载、转发与流控能力')]:text(s,705,y,180,30,h,22,'purple',True);text(s,885,y,349,55,b,20)
text(s,705,383,525,27,'RoCE 端点不会因为“接上以太网”自动成为 UET 端点。',17,'gray')
panel(s,30,435,761,219,'③ UE 1.1：2026-10-12 发布，新增能力按 profile 看','purple')
for y,h,b in [(481,'In-Network Collectives','把部分集合通信能力纳入网络内执行机制'),(538,'可编程拥塞管理','PCM 等增强；不能用名称推断实际产品支持程度'),(595,'Unified Local profile','面向本地网络，包含 Scale-up 相关增强')]:text(s,45,y,279,35,h,20,'purple',True);text(s,327,y,446,45,b,20)
panel(s,808,435,442,219,'④ 三个必须保留的限定','orange');text(s,823,482,411,148,'规范有多个 profile，不能套一个配置。\n\n既涉及 best-effort，也涉及 lossless 场景。\n\n规范版本 ≠ 产品上市 ≠ 全功能互操作。',20)
#6
s=page('对照与排障：按技术对象选型，按通信路径验证','先对齐协议层级，再看端点、网络与应用证据；相同链路速率不代表相同有效吞吐或尾延迟。',['S2','S6','S10','S11','S14'])
xs=[212,471,730,989];cs=['blue','orange','green','purple']
for x,h,c in zip(xs,['Ethernet 基础','InfiniBand','RoCEv2','Ultra Ethernet'],cs):tag(s,x,124,h,c,252)
rows=[('技术定位',['MAC / PHY 网络基础','完整互连体系','Ethernet 上的 RDMA','以太网增强架构 / UET']),('端点要求',['NIC；上层功能另核对','IB HCA 与软件','RoCE RNIC 与软件','UET 实现及软件栈']),('传输路径',['可承载 TCP/IP 等','IB 传输 / 网络 / 链路','RDMA → UDP/IP → Eth','依 UE 版本 / profile']),('流控与拥塞',['取决于上层和配置','逐跳流控等机制','ECN/CNP；可配置 PFC','UET 拥塞管理与分层增强']),('主要边界',['Eth 不自动保证 RDMA','不能直插普通 Eth 端口','UDP ≠ RDMA 可靠性','支持 UET ≠ 支持所有功能'])]
for i,(h,vals) in enumerate(rows):
 y=172+i*55;box(s,30,y,1220,50,'pale' if i%2==0 else 'white','line',True);text(s,42,y+13,162,30,h,20,'ink',True)
 for x,v,c in zip(xs,vals,cs):text(s,x+5,y+10,242,39,v,18,c)
panel(s,30,461,759,194,'按症状往回查：不要只盯端口速率')
for y,h,b in [(507,'吞吐上不去','注入带宽 → 超售 / 路径 → 消息大小与通信重叠'),(556,'P99 延迟升高','热点 / 队列 → ECN 或暂停 → 重传 / 慢节点'),(605,'RDMA 无法通信','端点与软件 → 内存权限 → 地址 / 队列 / 网络配置')]:text(s,44,y,169,29,h,20,'blue',True);text(s,218,y,551,36,b,20)
panel(s,806,461,444,194,'交付前：三类证据缺一不可','green');text(s,822,510,410,126,'协议：版本、profile、端点功能清单\n\n互操作：NIC / HCA + 网络 + 软件组合\n\n业务：训练步时、尾延迟及故障场景',20)
# remove theme effects
for s in p.slides:
 for sh in s.shapes:
  st=sh._element.find(qn('p:style'))
  if st is not None:sh._element.remove(st)
f=O/'数据中心网络_高密度讲解版.pptx';p.save(f)
(R/'deck-map.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2));print(f,f.stat().st_size)
