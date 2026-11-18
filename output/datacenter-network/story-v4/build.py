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
 s=p.slides.add_slide(p.slide_layouts[6]);box(s,0,0,1280,720,'white','white');box(s,28,26,6,40,'blue','blue');text(s,48,22,1200,49,title,32,'ink',True);text(s,48,82,1190,34,sub,19,'gray')
 line(s,30,676,1250,676,'line',False,1);text(s,32,686,1190,23,'来源：'+' · '.join(ids)+'  |  资料核查：2026-11-17  |  示意图非产品实物或实测数据',12,'gray');text(s,1220,683,40,28,str(len(p.slides)),17,'blue',True)
 s.notes_slide.notes_text_frame.text='\n'.join(k+': '+sources[k] for k in ids)
 meta.append({'page':len(p.slides),'title':title,'sources':ids});return s
def label(s,x,y,w,t,col='ink',size=20):text(s,x,y,w,48,t,size,col,True)
def panel(s,x,y,w,h,title):
 box(s,x,y,w,h,'pale','line',True);label(s,x+16,y+12,w-32,title,'blue',21)
def packet(s,x,y,t='K',w=42):
 box(s,x,y,w,30,'purple','purple',True);text(s,x+5,y+4,w-10,25,t,17,'white',True)
def switch(s,x,y,t):
 box(s,x,y,118,52,'pale','blue',True);text(s,x+10,y+7,100,24,t,17,'blue',True)
 for i in range(8):box(s,x+10+i*12,y+36,7,6,'blue','blue')
def hosts(s,endpoint):
 tag(s,34,133,'同一案例：GPU A → GPU B','purple',348)
 text(s,408,140,800,28,'紫色＝KV 数据    绿色＝反馈    灰色＝CPU 控制',17,'gray')
 for x,name in [(34,'服务器 A · 发送'),(944,'服务器 B · 接收')]:
  label(s,x+14,187,280,name,'ink',22);s.shapes.add_picture(str(R/'assets/server.png'),U(x),U(224),width=U(302),height=U(154))
  node(s,x+(4 if x==34 else 166),399,132,45,'GPU 显存','purple',20);node(s,x+(166 if x==34 else 4),399,132,45,endpoint,'blue',19)
 line(s,170,421,200,421,'purple',True,2);line(s,1080,421,1110,421,'purple',True,2)
 packet(s,65,459,'KV Cache',102);packet(s,1120,459,'KV Cache',102)
 text(s,36,641,1205,25,'案例假设：内存已注册与授权，平台支持 GPU 显存直达；图中服务器及分块数量均为机制示意。',15,'gray')
def bottom(s,items):
 for i,(h,b) in enumerate(items):
  x=34+i*410;panel(s,x,527,392,100,h);text(s,x+16,566,360,54,b,18)
# 1: relationships
s=page('Ethernet 与 IB 是两种网络体系；RoCE、UE 属于以太网路线','先回答“它们是什么关系”，再看数据如何传、拥塞如何处理。',['S2','S6','S8','S11'])
box(s,34,140,805,442,'pale','blue',True);label(s,55,155,760,'Ethernet｜以太网体系','blue',30);text(s,57,202,747,42,'解决基础互联：让设备通过以太网链路与交换网络通信。',22)
box(s,55,263,355,285,'white','blue',True);label(s,73,279,320,'RoCE｜以太网上的 RDMA','blue',24);text(s,74,326,315,62,'让网卡直接搬运\n已授权内存中的数据。',23);node(s,75,425,313,54,'沿用 Ethernet 网络','blue',23);text(s,76,495,310,31,'本例后续使用可路由的 RoCEv2',16,'gray')
box(s,430,263,386,285,'white','blue',True);label(s,447,279,350,'Ultra Ethernet｜增强规范','blue',24);text(s,448,325,351,59,'面向 AI / HPC 的\n传输、拥塞管理等增强。',23);box(s,449,411,346,112,'pale','blue',True);label(s,466,425,312,'UET｜其中的传输协议','blue',22);text(s,467,466,309,42,'负责端点间的数据传输与交付',19)
box(s,861,140,385,442,'white','blue',True);label(s,882,156,340,'InfiniBand｜IB 体系','blue',29);text(s,883,210,337,83,'将 RDMA、交换互连\n与流控作为一套体系。',24)
for y,t in [(325,'HCA：端点通信与搬运'),(389,'IB 交换网络：连接端点'),(453,'链路流控：保护接收缓冲')]:node(s,883,y,340,49,t,'blue',20)
label(s,58,607,1170,'读图结论：RoCE 与 UET 是不同的传输路线；UET 只是 Ultra Ethernet 的一部分。','ink',23)
#2
s=page('CPU 做准备，网卡把 KV Cache 搬到另一台服务器','RDMA：远程直接内存访问。先理解共同目标，再比较三种传输路线。',['S2','S6','S11'])
hosts(s,'RDMA 端点');switch(s,432,393,'交换节点');switch(s,693,393,'交换节点');line(s,332,421,432,421,'purple',True,3);line(s,550,421,693,421,'purple',True,3);line(s,811,421,948,421,'purple',True,3)
node(s,429,215,382,49,'① CPU 注册内存、授权、提交操作','gray',20);line(s,429,240,270,393,'gray',True,1.5)
label(s,436,294,390,'② 网卡执行数据搬运','purple',24)
for x,t in [(447,'K1'),(530,'K2'),(613,'K3')]:packet(s,x,339,t)
label(s,439,467,429,'③ 完成后，软件按协议交接使用','gray',19)
bottom(s,[('搬运的对象','同一块 KV Cache：从 A 的授权区域\n写入 B 的目标区域。'),('CPU 的职责','准备资源、提交工作、处理完成。\nRDMA 不等于完全不用 CPU。'),('接下来比较什么','IB / RoCEv2 / UET：\n端点、网络与反馈机制如何变化？')])
#3
s=page('下一跳还有多少缓冲空间，决定这一跳能发多少','InfiniBand：网卡换成 HCA，网络换成 IB；本页只放大逐跳链路流控。',['S2','S3'])
hosts(s,'IB HCA');switch(s,432,393,'IB 交换机');switch(s,693,393,'IB 交换机');
for a,b in [(332,432),(550,693),(811,948)]:line(s,a,421,b,421,'purple',True,3)
for a,b in [(432,332),(693,550),(948,811)]:
 line(s,a,445,a,481,'green',False,1.8);line(s,a,481,b,481,'green',False,1.8);line(s,b,481,b,445,'green',True,1.8)
label(s,434,208,443,'放大：交换机 B 的接收缓冲','blue',22)
for i in range(7):box(s,438+i*48,258,36,42,'purple' if i<5 else 'white','purple')
text(s,439,314,423,52,'已占用区域       可用空间 → 通告上游\n格数仅示意，不代表硬件队列深度。',18,'gray');line(s,710,366,751,393,'gray',False,1)
bottom(s,[('空间不足 → 限制发送','相邻发送方依据流控信息发数据，\n避免压垮下一跳接收缓冲。'),('空间释放 → 继续发送','消费数据后更新可用空间；\n反馈一跳一跳返回。'),('解决缓冲溢出，不消除拥塞','仍可能排队、反压和等待。\nIB 的流控不是 Ethernet PFC。')])
#4
s=page('RoCEv2 在以太网上做 RDMA，用拥塞反馈让源端降速','保留 Ethernet 网络，使用 RDMA 网卡；看清“拥塞点 → 接收端 → 发送端”的闭环。',['S6','S14'])
hosts(s,'RoCE RNIC');switch(s,432,393,'以太网交换机');switch(s,693,393,'以太网交换机')
for a,b in [(332,432),(550,693),(811,948)]:line(s,a,421,b,421,'purple',True,3)
label(s,424,207,420,'RoCEv2 数据包：UDP 是封装','blue',21)
for x,w,t in [(427,108,'Ethernet'),(540,52,'IP'),(597,109,'UDP 4791'),(711,121,'RDMA + K')]:node(s,x,251,w,42,t,'blue',16)
packet(s,727,326,'K + ECN',101);line(s,778,367,778,392,'purple',True,2);label(s,425,329,290,'① 队列拥塞 → 标记 ECN','green',18)
line(s,1014,446,1014,496,'green',False,2);line(s,1014,496,266,496,'green',False,2);line(s,266,496,266,444,'green',True,2)
text(s,423,463,620,28,'③ 源端降速  ←  CNP 拥塞通知  ←  ② 接收端发现标记',18,'green',True)
bottom(s,[('RDMA：网卡搬运数据','UDP 不负责可靠交付；\n可靠性由 RDMA 传输机制提供。'),('ECN / CNP：端到端降速','交换机标记拥塞，接收端发 CNP；\n源端据此调整发送速率。'),('PFC：另一种局部保护','可按优先级暂停相邻上游；\n是否启用取决于部署模式与配置。')])
#5
s=page('UET 可让数据分走多条路径，并可靠交付到目标内存','UET 是 Ultra Ethernet 的传输协议；这里展示其多路径与乱序处理能力。',['S10','S11'])
hosts(s,'UET 端点');switch(s,432,393,'以太网交换机');switch(s,693,393,'以太网交换机');line(s,332,421,432,421,'purple',True,3);line(s,811,421,948,421,'purple',True,3)
for y,t in [(242,'K1'),(305,'K2'),(368,'K3')]:
 line(s,550,415,579,y+15,'purple',False,1.7);line(s,579,y+15,666,y+15,'purple',True,1.7);line(s,666,y+15,693,415,'purple',False,1.7);packet(s,594,y,t)
text(s,427,198,429,28,'同一块数据 → 可选择多个可用路径',20,'purple',True)
line(s,1014,445,1014,498,'green',False,2);line(s,1014,498,267,498,'green',False,2);line(s,267,498,267,444,'green',True,2)
text(s,426,463,606,28,'确认 / 重传相关反馈 + 拥塞控制 → 保证交付并调节发送',18,'green',True)
bottom(s,[('发送：利用多条路径','分散流量，减少单一路径热点；\n图中三路不是固定产品拓扑。'),('接收：按目标位置放置','即使 K2 先到，也定位到对应区域；\n满足完成条件后再交给软件。'),('边界：看 profile 与实现','端点、网络和软件需支持所选能力。\nUE 1.1 规范能力 ≠ 所有产品已实现。')])
#6
s=page('记住：两种网络体系、一种 RDMA 方式、一套增强规范','回到最初的问题：它是什么、基于什么网络、解决什么问题？',['S2','S6','S8','S11','S14'])
xs=[34,250,579,829];ws=[216,329,250,417]
for x,w,t in zip(xs,ws,['名称','它是什么','基于什么网络','解决什么问题']):box(s,x,145,w,48,'blue','white');text(s,x+14,155,w-28,30,t,22,'white',True)
rows=[['Ethernet','基础网络体系','Ethernet','设备怎样通过链路与交换网络互联'],['InfiniBand','支持 RDMA 的完整互连体系','IB 原生网络','端点搬运、交换转发、逐跳流控\n如何协同工作'],['RoCE / RoCEv2','以太网上的 RDMA 方式','Ethernet\nv2 使用 IP / UDP 封装','在以太网承载远端内存访问；\n用拥塞控制与配置支撑稳定传输'],['Ultra Ethernet','面向 AI / HPC 的\n以太网增强规范','Ethernet','增强传输、拥塞管理及相关能力；\nUET 是其中的传输协议']]
for i,row in enumerate(rows):
 y=193+i*96
 for j,(x,w,t) in enumerate(zip(xs,ws,row)):
  box(s,x,y,w,96,'pale' if i%2==0 else 'white','line');text(s,x+14,y+19,w-28,72,t,21 if j==0 else 20,'blue' if j==0 else 'ink',j==0)
box(s,34,603,1212,49,'pale','blue',True);text(s,52,614,1170,30,'复述练习：我用什么网络？是否要 RDMA？选择 RoCE 还是 UET？端点与交换网络是否支持？',22,'blue',True)
for s in p.slides:
 for sh in s.shapes:
  st=sh._element.find(qn('p:style'))
  if st is not None:sh._element.remove(st)
O.mkdir(exist_ok=True);p.save(O/'数据中心网络_关系与机制版.pptx');(R/'deck-map.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2))
