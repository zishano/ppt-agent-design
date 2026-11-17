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
from PIL import Image
from pptx.oxml.ns import qn
sources={k:(R/'research'/f'{k}.url').read_text().strip() for k in ['S1','S2','S3','S4','S6','S8','S10','S11','S14']}
meta=[]
def slide(title,sub,ids):
 s=p.slides.add_slide(p.slide_layouts[6]);box(s,0,0,1280,720,'white','white');box(s,30,27,6,45,'blue','blue');text(s,49,25,1190,55,title,33,'ink',True);text(s,49,87,1175,40,sub,19,'gray');line(s,32,663,1248,663,'line',False,1)
 n=len(p.slides);text(s,35,679,1165,25,'来源：'+' · '.join(ids)+' | 核查：2026-11-17；拓扑、包序与队列为机制示意，不表示实测性能。',13,'gray');text(s,1211,677,40,25,f'{n:02}',16,'blue',True)
 s.notes_slide.notes_text_frame.text='来源\n'+'\n'.join(f'{i}: {sources[i]}' for i in ids)+'\n范围：示意拓扑不代表设备数量、布线或性能基准。UE以1.1规范及其profile区分能力；RoCE调优取决于实际端点、交换机与配置。'
 meta.append({'page':n,'title':title,'sources':ids,'reading_order':'整体关系 → 机制展开 → 边界','checks':['facts','protocol_roles','arrow_endpoints','text','actual_render']});return s

def switch(s,x,y,w=155,label='交换机',col='blue'):
 box(s,x,y,w,54,'pale',col,True);text(s,x+10,y+7,w-20,26,label,20,col,True)
 for j in range(8):box(s,x+12+j*(w-24)/8,y+38,9,7,'white',col)
def packet(s,x,y,t,col='blue',w=40):node(s,x,y,w,27,t,col,14)
def bracket(s,x,y,w,t,col):
 line(s,x,y,x+w,y,col,False,1.5);line(s,x,y,x,y-8,col,False,1.5);line(s,x+w,y,x+w,y-8,col,False,1.5);text(s,x,y+8,w,30,t,18,col,True)
#1
s=slide('四个名字不是四张平行的网：先把技术层级放对','Ethernet 提供网络基础；RoCE 与 UET 提供不同的 RDMA 传输路径；InfiniBand 是另一套互连体系。',['S2','S6','S8','S11'])
box(s,35,147,898,467,'pale','line',True)
cols=[(60,'Ethernet 常见业务栈','blue',['应用 / Sockets','TCP 或 UDP','IP','Ethernet MAC / PHY']), (348,'RoCEv2 路径','green',['RDMA 应用','IB 传输语义','UDP / IP','Ethernet MAC / PHY']), (636,'Ultra Ethernet 路径','purple',['AI / HPC 应用','UET / RDMA','IP 等承载能力¹','Ethernet 基础与增强'])]
for x,h,col,layers in cols:
 tag(s,x,164,h,col,270)
 for j,t in enumerate(layers):
  node(s,x+8,228+j*70,254,47,t,col,21)
  if j<3:line(s,x+135,275+j*70,x+135,298+j*70,col,True,1.5)
text(s,64,537,839,57,'共同依托以太网生态 ≠ 使用相同传输协议。\n¹ UE 1.1 增加本地网络 profile；这里展示典型可路由路径。',19,'gray')
box(s,958,147,286,467,'white','orange',True);tag(s,975,164,'InfiniBand','orange',252)
for j,t in enumerate(['RDMA / 消息应用','IB 传输层','IB 网络与链路','IB 物理互连']):
 node(s,981,228+j*70,240,47,t,'orange',21)
 if j<3:line(s,1101,275+j*70,1101,298+j*70,'orange',True,1.5)
text(s,977,538,249,64,'端点与交换网络\n采用 InfiniBand 体系。',20,'orange',True)
#2
s=slide('整体网络相似，端点能力和传输机制决定实际行为','Leaf–Spine 只描述拓扑；同样的拓扑图，不能证明协议兼容、RDMA 支持或应用吞吐。',['S2','S6','S11','S14'])
box(s,35,145,764,480,'white','line',True)
for i in range(3):switch(s,116+i*215,169,155,'Spine '+str(i+1),'blue')
for i in range(4):
 x=62+i*180
 for j in range(3):line(s,x+67,316,193+j*215,223,'blue',False,1.3)
 switch(s,x,316,136,'Leaf '+str(i+1),'blue');line(s,x+68,370,x+68,417,'green',False,2)
 node(s,x+10,417,116,39,'网卡 / HCA','green',17);line(s,x+68,456,x+68,493,'green',False,2)
 box(s,x-4,493,144,91,'pale','line',True);text(s,x+8,501,127,25,'计算节点 '+str(i+1),18,'ink',True);gpu(s,x+8,540,'CPU','gray',53,28);gpu(s,x+76,540,'GPU','blue',53,28)
text(s,59,599,715,22,'逻辑拓扑示意；协议选择需落实到所有端点、交换机及软件配置。',15,'gray')
for y,h,b,col in [(153,'01 应用与内存','通信库提交数据交换；\nCPU 内存与 GPU 显存路径须区分。','ink'),(311,'02 端点','普通 NIC、RoCE RNIC、IB HCA\n或 UET 端点，能力不自动等价。','green'),(469,'03 网络','转发、队列、拥塞反馈与管理\n必须匹配端点和实际业务。','blue')]:tag(s,834,y,h,col,398);text(s,839,y+50,392,104,b,22)
#3
s=slide('RDMA 优化的是数据通路：CPU 仍负责准备和控制','以一次 RDMA Write 为例：预先注册并授权内存后，网卡执行数据搬运；不意味着 CPU 完全不参与。',['S2','S6','S11'])
for x,h in [(45,'发送主机'),(790,'接收主机')]:
 box(s,x,151,445,405,'pale','line',True);tag(s,x+17,167,h,'blue',410)
 node(s,x+25,232,170,53,'应用 / CPU','gray',23);node(s,x+237,232,178,53,'注册与权限','gray',23);line(s,x+195,258,x+237,258,'gray',True,1.5)
 mx=x+28 if x==45 else x+237;nx=x+255 if x==45 else x+28
 node(s,mx,377,175,62,'已注册内存','blue',23);node(s,nx,377,160,62,'RDMA 网卡','green',23)
 line(s,x+326,285,x+326,330,'gray',False,1.4);line(s,x+326,330,nx+80,330,'gray',False,1.4);line(s,nx+80,330,nx+80,377,'gray',True,1.4)
 line(s,mx+175,406,nx,406,'blue',True,3) if x==45 else line(s,nx+160,406,mx,406,'blue',True,3)
 text(s,x+24,482,398,49,'控制路径：建立资源、配置访问权限\n数据路径：端点执行内存访问',19,'gray')
line(s,460,407,818,407,'green',True,3);node(s,540,373,200,67,'网络传输','green',25)
text(s,543,463,194,53,'IB / RoCE / UET\n取决于所选端点与协议',16,'green')
box(s,45,579,1190,58,'white','line',True);text(s,61,593,1159,34,'GPU 显存直达还需设备、驱动和软件支持；“支持 RDMA”本身不等于 GPU 数据路径已经打通。',22,'ink',True)
#4
s=slide('InfiniBand：RDMA、交换网络与逐跳流控协同工作','“无损”指正常拥塞管理下避免缓冲溢出丢包，不表示没有链路故障，也不等于没有拥塞等待。',['S2','S3'])
box(s,35,144,777,480,'white','line',True)
node(s,60,224,137,58,'HCA A','orange',25);switch(s,265,224,176,'IB Switch 1','orange');switch(s,500,224,176,'IB Switch 2','orange');node(s,698,224,91,58,'HCA B','orange',21)
for a,b in [(197,265),(441,500),(676,698)]:line(s,a,252,b,252,'orange',True,2.4)
for x1,x2 in [(265,197),(500,441),(698,676)]:
 line(s,x1,282,x1,316,'green',False,1.5);line(s,x1,316,x2,316,'green',False,2);line(s,x2,316,x2,282,'green',True,1.5)
text(s,79,340,686,33,'逐跳反馈：相邻链路按接收侧缓冲能力控制发送',22,'green',True)
text(s,64,174,712,32,'RDMA / 消息传输 → HCA → IB Fabric → HCA',22,'orange',True)
box(s,61,412,725,169,'pale','line',True);text(s,79,427,685,31,'局部放大：接收缓冲与可发送状态',23,'ink',True)
for j in range(10):box(s,88+j*39,483,29,34,'orange' if j<7 else 'white','orange')
text(s,503,479,251,62,'缓冲紧张 → 减少发送\n空间恢复 → 继续发送',20,'green',True)
text(s,85,536,655,25,'方格只说明缓冲占用，不代表队列深度或测量结果。',16,'gray')
for y,h,b in [(155,'01 原生 RDMA','IB 传输支持消息和远端内存操作，\n数据搬运由端点执行。'),(313,'02 逐跳控制','每跳保护接收侧缓冲；\n不能用以太网 PFC 名称替代。'),(471,'03 运维边界','端点、交换机与管理属于 IB 体系。\n仍需检查路由、故障与业务拥塞。')]:tag(s,843,y,h,'orange',388);text(s,848,y+48,380,104,b,21)
#5 RoCE
s=slide('RoCEv2：在 IP / UDP 上承载 RDMA，拥塞控制分两层','UDP 是封装层；可靠传输由 RDMA 传输协议负责。PFC 的暂停和 ECN/CNP 的反馈不是同一件事。',['S6','S14'])
for x,w,t,col in [(45,179,'Ethernet','blue'),(230,132,'IP','blue'),(368,202,'UDP · 4791','green'),(576,326,'RDMA 传输头','green'),(908,326,'数据 / 校验','purple')]:node(s,x,148,w,49,t,col,24)
text(s,49,210,1182,32,'RoCEv1 是二层封装；RoCEv2 加入 IP / UDP，可跨三层网络。4791 为 UDP 目的端口。',20,'gray')
box(s,35,258,1209,267,'white','line',True)
for x,label in [(62,'发送 RNIC'),(1062,'接收 RNIC')]:node(s,x,312,150,60,label,'green',23)
switch(s,335,312,180,'交换机 A','blue');switch(s,689,312,180,'交换机 B','blue')
for a,b in [(212,335),(515,689),(869,1062)]:line(s,a,340,b,340,'blue',True,2.5)
text(s,533,286,144,28,'ECN 标记',20,'orange',True);text(s,900,382,290,28,'收到拥塞标记',19,'green')
# end to end CNP return around bottom
line(s,1137,372,1137,467,'green',False,2);line(s,1137,467,137,467,'green',False,2);line(s,137,467,137,372,'green',True,2)
text(s,419,476,540,30,'ECN / CNP：接收端通知发送端调整速率',21,'green',True)
line(s,689,366,689,403,'orange',False,2);line(s,689,403,515,403,'orange',False,2.4);line(s,515,403,515,366,'orange',True,2);text(s,541,413,151,25,'PFC：逐跳暂停',17,'orange',True)
for x,h,b,col in [(35,'端到端：控制注入速率','拥塞点标记 ECN → 接收端反馈 CNP → 发送端调速。','green'),(650,'逐跳：保护选定优先级','PFC 暂停相邻上游的指定优先级；并非暂停所有业务。','orange')]:tag(s,x,545,h,col,594);text(s,x+8,592,575,49,b,20)
#6 UE
s=slide('Ultra Ethernet：以 UET 重构 AI/HPC 的以太网传输','当前官网版本为 UE 1.1；多路径、拥塞控制、可靠交付与软件语义需协同实现，不能只换一台交换机。',['S4','S10','S11'])
box(s,35,145,687,482,'pale','purple',True);tag(s,55,162,'整体：端点 + 多路径 Fabric + 端点','purple',648)
node(s,63,308,139,60,'UET 发送端','purple',20);node(s,553,308,139,60,'UET 接收端','purple',20)
for i,y in enumerate([235,333,431]):
 line(s,202,338,285,y+22,'purple',False,1.7);line(s,435,y+22,553,338,'purple',False,1.7);node(s,285,y,150,43,'路径 '+str(i+1),'purple',20);packet(s,230,y-7,str(i+1),'purple');packet(s,483,y-7,str(i+1),'purple')
text(s,62,517,629,49,'包级多路径与乱序交付能力由 UET 协同处理；\n重传、拥塞管理及完成语义仍由协议保障。',21)
text(s,64,589,620,22,'仅示意机制；不是 RoCE 封装，也不是实际路径或丢包测试。',15,'gray')
for y,h,b in [(151,'01 传输层：UET','面向 RDMA 的现代传输；\n结合多路径、可靠交付与拥塞管理。'),(302,'02 以太网与软件协同','规范涉及软件、传输、链路等层级。\n支持能力需核对具体 profile 与实现。'),(453,'03 UE 1.1：新增能力','2026-10-12 发布；包括网络内集合操作、\n可编程拥塞管理与 Unified Local profile。')]:tag(s,752,y,h,'purple',480);text(s,758,y+48,470,101,b,21)
text(s,758,607,468,30,'本地 profile 涉及 Scale-up；不能再一概只称 Scale-out。',16,'gray')
#7 matrix
s=slide('一张表对齐边界：网络基础、独立体系与 RDMA 路径','比较的是角色和部署要求，不是按名称给出固定的“速度排名”。',['S2','S6','S10','S11','S14'])
cols=[(222,241,'Ethernet 基础','blue'),(474,241,'InfiniBand','orange'),(726,241,'RoCEv2','green'),(978,265,'Ultra Ethernet','purple')]
for x,w,h,c in cols:tag(s,x,150,h,c,w)
rows=[('定义',['MAC / PHY 等基础','独立互连体系','RDMA over Ethernet','以太网增强架构 / UET']),('RDMA 支持',['不由基础本身保证','原生支持','需要 RoCE 端点','需要 UET 实现与软件']),('承载路径',['可承载 IP 业务','IB 传输 / 网络 / 链路','RDMA → UDP/IP → Eth','按 UE profile 定义']),('拥塞 / 流控',['由上层和配置决定','逐跳流控等机制','ECN/CNP；可配置 PFC','UET 拥塞管理；分 profile']),('兼容边界',['同属以太网不等于全兼容','不能直接接入普通 Eth 端口','交换网络与端点需匹配','RoCE 端点不自动变成 UET']),('验收重点',['速率、拓扑、业务 QoS','端点、路由、管理和故障','队列、反馈、丢包与尾延迟','版本、profile、产品互操作'])]
for i,(h,values) in enumerate(rows):
 y=208+i*65;box(s,35,y,1209,58,'pale' if i%2==0 else 'white','line',True);text(s,49,y+16,163,31,h,20,'ink',True)
 for (x,w,_,c),v in zip(cols,values):text(s,x+7,y+11,w-14,49,v,18,c)
text(s,48,620,1187,28,'RoCE 存在不同无损/有损部署配置；UE 也不能简单归类为“永远不需要流控”或“零丢包”。',18,'gray')
#8
s=slide('选型从应用通信图出发，再验证端点、网络与运维','同样的端口速率不保证同样的作业表现；用集合通信与真实应用测试闭环，而不是只测链路峰值。',['S2','S6','S11','S14'])
steps=[('01 工作负载','梯度同步 / All-to-all\n消息大小与并发数\n尾延迟与通信重叠','blue'),('02 端点与软件','NIC / HCA 支持什么\n通信库与内存路径\nGPU 显存访问条件','green'),('03 交换网络','上联与超售比例\n队列、ECN / PFC 策略\n路径与故障恢复','orange'),('04 系统验收','吞吐 / 训练步时\nP99 延迟与拥塞场景\n重传、暂停与 CPU 开销','purple')]
for i,(h,b,c) in enumerate(steps):
 x=35+i*311;tag(s,x,151,h,c,282);box(s,x,206,282,179,'pale','line',True);text(s,x+15,228,252,138,b,22)
 if i<3:line(s,x+283,169,x+305,169,'gray',True,2)
box(s,35,419,1210,202,'white','line',True);text(s,57,435,1167,35,'部署前，用三条证据确认“真的可用”',28,'ink',True)
for i,(h,b,c) in enumerate([('协议证据','版本、profile 与端点功能清单','blue'),('互操作证据','NIC + 交换机 + 软件组合测试','green'),('业务证据','拥塞和故障条件下的应用结果','purple')]):
 x=57+i*397;tag(s,x,492,h,c,371);text(s,x,544,365,60,b,22)
text(s,48,633,1177,22,'规范支持 ≠ 产品已实现；厂商峰值 ≠ 集群有效吞吐；网络无损 ≠ 没有拥塞。',18,'gray')
# remove default theme effects, retain authored fills/fonts
for s in p.slides:
 for sh in s.shapes:
  style=sh._element.find(qn('p:style'))
  if style is not None:sh._element.remove(style)
f=O/'数据中心网络_Ethernet_InfiniBand_RoCE_UltraEthernet.pptx';p.save(f)
(R/'deck-map.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2))
(R/'sources.json').write_text(json.dumps([{'id':k,'url':v,'accessed':'2026-11-17','local_excerpt':'research/'+k+'.txt','status':'read'} for k,v in sources.items()],ensure_ascii=False,indent=2))
print(f,f.stat().st_size)
