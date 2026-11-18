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
# Common static illustrated scene; native data/control overlays.
def scene(s,endpoint,network):
 s.shapes.add_picture(str(R/'assets/network-scene.png'),U(34),U(151),width=U(1212),height=U(404))
 tag(s,34,128,'同一块 KV Cache：GPU A → GPU B','purple',412)
 text(s,680,136,565,25,'紫：数据　绿：反馈　灰：准备 / 控制',18,'gray')
 label(s,57,181,320,'服务器 A · 发送','ink',23);label(s,947,181,300,'服务器 B · 接收','ink',23)
 for x,t in [(78,'GPU 显存'),(948,'GPU 显存')]:
  node(s,x,280,124,38,t,'purple',18)
 for x in [257,1090]:node(s,x,407,140,38,endpoint,'blue',18)
 for x in [441,700]:node(s,x,401,141,38,network,'blue',18)
 # Separate serial segments: no shared bus or bypass around switches.
 line(s,139,318,139,470,'purple',False,2);line(s,139,470,327,470,'purple',False,2);line(s,327,470,327,445,'purple',True,2)
 for left,right,y in [(327,510,491),(510,770,470),(770,1150,491)]:
  line(s,left,445,left,y,'purple',False,2);line(s,left,y,right,y,'purple',False,2);line(s,right,y,right,445,'purple',True,2)
 line(s,1150,407,1010,318,'purple',True,2)
 packet(s,155,455,'KV',50);packet(s,1040,472,'KV',50)
 text(s,36,649,1210,24,'概念示意：设备外形、端口数和数据块数量非产品规格。假设内存已授权，平台支持 GPU 显存直达。',14,'gray')
def detail(s,x,w,title,body):
 panel(s,x,550,w,91,title);text(s,x+16,590,w-32,48,body,18)
def feedback(s,left,right,y=533):
 line(s,right,495,right,y,'green',False,2);line(s,right,y,left,y,'green',False,2);line(s,left,y,left,495,'green',True,2)
#2
s=page('CPU 做准备，网卡把 KV Cache 搬到另一台服务器','RDMA = 远程直接内存访问：先理解共同的数据搬运目标，再比较网络与传输机制。',['S2','S6','S11'])
scene(s,'RDMA 端点','交换网络')
node(s,454,213,354,44,'① 注册内存 / 授权 / 提交操作','gray',20)
line(s,454,235,327,405,'gray',True,1.8)
label(s,452,276,404,'② 网卡搬运已授权的数据','purple',23)
for x,t in [(468,'K1'),(554,'K2'),(640,'K3')]:packet(s,x,316,t,52)
text(s,438,508,480,27,'③ 传输完成后，软件按约定使用目标数据',19,'gray',True)
detail(s,34,392,'发送侧：CPU 仍然参与','准备资源和提交工作；\n不等于每个数据包都由 CPU 搬运。')
detail(s,444,392,'数据面：内存 → 网卡 → 网络','右侧端点写入授权目标区域；\nGPU 直达需要硬件、驱动与软件支持。')
detail(s,854,392,'后续比较同一件事','IB、RoCEv2 和 UET 怎样组织网络、\n控制发送，并完成这次搬运？')
#3
s=page('下一跳还有多少缓冲空间，决定这一跳能发多少','InfiniBand 是独立网络体系：HCA 执行端点通信，IB 交换网络使用自己的逐跳链路流控。',['S2','S3'])
scene(s,'IB HCA','IB 交换机')
label(s,449,212,400,'放大接收缓冲：有空间才继续发','blue',21)
for i in range(7):box(s,450+47*i,254,35,34,'purple' if i<5 else 'white','purple')
text(s,451,296,393,50,'紫色：已占用　空格：可用空间\n上游依据流控信息限制发送',18,'green',True)
line(s,787,348,770,399,'gray',True,1)
for l,r in [(327,510),(510,770),(770,1150)]:feedback(s,l,r)
text(s,451,497,427,24,'逐跳返回：可用空间 / 流控信息',18,'green',True)
detail(s,34,392,'缓冲紧张 → 上游等待','限制相邻发送方继续注入数据，\n避免接收缓冲溢出。')
detail(s,444,392,'释放空间 → 更新流控信息','每一跳分别保护下一跳缓冲；\n不是接收应用向全网发暂停。')
detail(s,854,392,'解决溢出，不消除拥塞','仍会排队、反压或等待。\nIB 链路流控不是 Ethernet PFC。')
#4
s=page('RoCEv2 在以太网上做 RDMA，用拥塞反馈让源端降速','Ethernet 提供基础互联；RoCEv2 使用 IP / UDP 封装，RDMA 端点负责数据搬运。',['S6','S14'])
scene(s,'RoCE RNIC','以太网交换机')
label(s,446,212,416,'① 拥塞点给数据包标记 ECN','green',21)
for x,w,t in [(441,86,'Ethernet'),(530,45,'IP'),(578,104,'UDP 4791'),(685,141,'RDMA + 数据')]:node(s,x,255,w,36,t,'blue',15)
packet(s,693,310,'K + ECN',113);line(s,749,343,770,399,'purple',True,1.6)
feedback(s,327,1150)
text(s,437,497,660,27,'③ 源端降速  ←  CNP 拥塞通知  ←  ② 接收端识别 ECN',18,'green',True)
detail(s,34,392,'UDP 封装 ≠ 可靠性来源','UDP 目的端口 4791；可靠交付\n由 RDMA 传输机制提供。')
detail(s,444,392,'ECN / CNP：端到端调速','交换机标记 → 接收端反馈 →\n发送端调整速率，减少拥塞。')
detail(s,854,392,'PFC：可选的局部保护','按优先级暂停相邻上游。\n是否启用取决于部署模式与配置。')
#5
s=page('UET 可让数据分走多条路径，并可靠交付到目标内存','Ultra Ethernet 是面向 AI / HPC 的增强规范；UET 是其中的传输协议，不是 RoCE 的别名。',['S10','S11'])
scene(s,'UET 端点','以太网交换机')
label(s,448,211,418,'局部展开：多路径分发同一块数据','purple',21)
for y,t in [(250,'K1'),(288,'K2'),(326,'K3')]:
 line(s,465,309,560,y+14,'purple',False,1.7);line(s,560,y+14,686,y+14,'purple',True,1.7);line(s,686,y+14,810,309,'purple',False,1.7);packet(s,597,y,t,50)
line(s,465,309,510,400,'gray',False,1);line(s,810,309,770,400,'gray',False,1)
feedback(s,327,1150)
text(s,438,497,625,27,'可靠交付反馈 / 重传机制 + 拥塞控制 → 调节发送',18,'green',True)
detail(s,34,392,'发送：减少单一路径热点','利用多个可用路径分散流量；\n示意路径数不是固定产品拓扑。')
detail(s,444,392,'接收：定位数据、满足完成条件','支持相应能力时，乱序数据按位置\n放置；完成后再交给软件使用。')
detail(s,854,392,'规范能力取决于实现与 profile','端点、网络、软件需要匹配支持。\nUE 1.1 不等于所有产品已实现。')
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
O.mkdir(exist_ok=True);p.save(O/'数据中心网络_机制插画版_v5.pptx');(R/'deck-map.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2))
