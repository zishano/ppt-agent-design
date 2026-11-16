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
sources={k:(R/'research'/f'{k}.url').read_text() for k in ['S1','S2','S3','S4','S5','S6','S7','S8']}
meta=[]
def slide(title,sub,ids):
 s=p.slides.add_slide(p.slide_layouts[6]);box(s,0,0,1280,720,'white','white');box(s,30,27,6,45,'blue','blue');text(s,49,25,1190,55,title,33,'ink',True);text(s,49,87,1175,40,sub,19,'gray');line(s,32,663,1248,663,'line',False,1)
 n=len(p.slides);text(s,35,679,1165,25,'来源：'+' · '.join(ids)+' | 核查：2026-11-16；逻辑图不代表实物布线，参数按指定版本解释。',13,'gray');text(s,1211,677,40,25,f'{n:02}',16,'blue',True)
 s.notes_slide.notes_text_frame.text='来源\n'+'\n'.join(f'{i}: {sources[i]}' for i in ids)+'\n所有拓扑均按文中平台或规范限定，不构成跨厂商性能测试。原厂图作为引用，其余图形与文字为原生可编辑元素。'
 meta.append({'page':n,'title':title,'sources':ids,'reading_order':'整体结构 → 分层详情 → 适用边界','checks':['facts','component_relations','counts','labels','render']});return s
def imagefit(s,path,x,y,w,h):
 iw,ih=Image.open(path).size;k=min(w/iw,h/ih);s.shapes.add_picture(str(path),U(x+(w-iw*k)/2),U(y+(h-ih*k)/2),width=U(iw*k),height=U(ih*k))
def pair(s,x,y,t1,t2,col='blue'):
 gpu(s,x,y,t1,col,95,48);gpu(s,x+300,y,t2,col,95,48);line(s,x+95,y+19,x+300,y+19,col,True,3);line(s,x+300,y+29,x+95,y+29,col,True,2)
# 1
s=slide('四个名字，三个层级：链路、交换组件与互连架构','NVLink 与 NVSwitch 是配合关系；UALink 是开放规范；Infinity Fabric 的范围不止 GPU 间连线。',['S1','S5','S7'])
box(s,35,145,746,470,'pale','line',True)
for y,title,col in [(164,'NVIDIA 路径','blue'),(309,'开放标准路径','green'),(454,'AMD 架构路径','purple')]:
 text(s,55,y,210,30,title,22,col,True)
 if y<400:
  gpu(s,63,y+58,'GPU A',col,106,47);gpu(s,645,y+58,'GPU B',col,106,47)
  node(s,323,y+56,176,51,'NVSwitch' if y==164 else 'UALink 交换组件',col,21 if y==164 else 18)
  line(s,169,y+78,323,y+78,col,False,3);line(s,499,y+78,645,y+78,col,False,3)
  text(s,188,y+45,122,25,'NVLink' if y==164 else 'UALink 链路',18,col,True);text(s,517,y+45,122,25,'NVLink' if y==164 else 'UALink 链路',18,col,True)
 else:
  node(s,63,y+58,190,49,'芯粒 / I/O Die',col,20);node(s,300,y+58,190,49,'封装 / GPU',col,20);node(s,538,y+58,215,49,'多 GPU 系统',col,20);line(s,253,y+82,300,y+82,col,False,3);line(s,490,y+82,538,y+82,col,False,3)
text(s,60,588,700,22,'图示表达技术角色；不同产品的接口、拓扑和一致性能力不能混用。',15,'gray')
for y,h,b,c in [(151,'01 链路：怎么传','NVLink 定义 NVIDIA GPU 互连；\nUALink 规范加速器及交换互连。','blue'),(307,'02 交换：怎么连更多端点','NVSwitch 是 NVLink 的交换组件。\n不能把它当成另一种独立链路。','green'),(465,'03 架构：覆盖哪些层级','Infinity Fabric 是 AMD 互连架构。\n必须落到具体芯片和系统分析。','purple')]:
 tag(s,816,y,h,c,418);text(s,821,y+48,412,98,b,21)
# 2
s=slide('NVLink 负责 GPU 间传输，NVSwitch 组织交换式互连','直接连接与交换式互连是不同拓扑；多个 GPU 可经交换组件通信，不意味着每对 GPU 都有独立物理线。',['S1'])
for x,h in [(35,'连接视角：一条 GPU ↔ GPU 通路'),(652,'系统视角：多个端点共享交换结构')]:box(s,x,146,592,382,'white','line',True);tag(s,x+16,161,h,'blue',560)
pair(s,108,279,'GPU A','GPU B');text(s,243,245,160,28,'NVLink 链路',22,'blue',True)
text(s,63,385,526,100,'传输数据与通信请求\n带宽与链路数量取决于 GPU 代际\n本图只解释直连关系，不指定产品拓扑',22)
for x,y,name in [(695,225,'GPU 0'),(1077,225,'GPU 1'),(695,421,'GPU 2'),(1077,421,'GPU 3')]:
 line(s,x+52,y+24,950,337,'blue',False,2);gpu(s,x,y,name,'blue',106,48)
node(s,839,309,223,55,'NVSwitch','blue',25)
text(s,689,489,530,25,'4 个端点仅作示意；交换芯片数量随平台变化。',16,'gray')
for x,h,b,col in [(35,'NVLink','接口 / 链路 / GPU 互连能力','blue'),(449,'NVSwitch','接入多个 NVLink，形成交换式域','green'),(863,'应用通信','仍需软件组织传输与集合操作','purple')]:
 tag(s,x,550,h,col,380);text(s,x+6,594,370,44,b,20)
#3 real hardware
s=slide('GB200 NVL72：把 NVLink 域扩展到一整柜 72 个 GPU','用真实整机说明层级：左侧是 NVIDIA 原厂产品图，右侧区分计算端、交换互连和域外网络。',['S1','S2'])
box(s,35,143,485,482,'ink','ink',True)
# native crop keeps only cabinet portion; original photo remains unmodified
ph=R/'research/gb200-nvl72.jpg';iw,ih=Image.open(ph).size
pic=s.shapes.add_picture(str(ph),U(158),U(151),width=U(215),height=U(452));pic.crop_left=.563;pic.crop_right=.339
text(s,61,601,429,22,'NVIDIA 原厂图；外观不代表所有 NVLink 平台。',14,'white')
for y,h in [(150,'01 计算端：72 GPU + 36 Grace CPU'),(309,'02 互连：NVLink 5 + NVLink Switch'),(469,'03 扩展：NIC 接入域外网络')]:tag(s,556,y,h,'blue' if y<450 else 'green',666)
for j in range(6):gpu(s,570+j*107,206,'GPU','blue',85,35)
text(s,573,253,635,35,'图中仅画 6 个端点；整柜为一个 72-GPU NVLink 域。',20)
node(s,573,361,260,44,'NVLink Switch System','blue',20);line(s,833,382,882,382,'blue',False,3);text(s,893,355,310,60,'1,800 GB/s / GPU\nNVLink 5 官方带宽口径',21,'blue',True)
node(s,575,523,185,42,'NVLink 域','blue',20);node(s,807,523,155,42,'NIC','green',20);node(s,1010,523,193,42,'Scale-out 网络','green',19);line(s,760,544,807,544,'green',False,2);line(s,962,544,1010,544,'green',False,2)
text(s,571,586,640,36,'域内互连与域外网络各有职责，不能把交换结构等同 NIC。',18,'gray')
#4 UALink
s=slide('UALink：开放的 Scale-up 互连规范，不是一台交换机','规范定义能力与互操作接口；具体设备、软件支持、认证与实际规模仍需逐产品验证。',['S7','S8'])
box(s,35,145,550,480,'pale','green',True);tag(s,55,161,'AI Pod：加速器 + 交换互连','green',509)
for i,(x,y) in enumerate([(72,220),(440,220),(72,415),(440,415)]):
 line(s,x+49,y+26,309,292,'green',False,2);gpu(s,x,y,'加速器','green',100,49)
node(s,170,263,278,58,'UALink 交换结构','green',25)
text(s,65,486,491,45,'200G / lane    ·    最多 1,024 加速器',25,'green',True);text(s,67,544,489,57,'以上为 200G 1.0 规范能力，\n不代表某台设备的吞吐或已部署规模。',19)
rows=[('Common 2.0','公共协议能力；引入 In-Network Compute'),('200G DL / PL 2.0','数据链路与物理层；与 Common 分拆'),('Chiplet 1.01','面向芯粒集成；与 UCIe 3.0 兼容'),('Manageability 1.0','系统控制与管理；标准化模型及 API')]
for i,(h,b) in enumerate(rows):
 y=150+i*112;tag(s,620,y,h,'green',610);text(s,625,y+47,603,51,b,22)
text(s,621,608,610,30,'官网还列有 128G DL / PL 1.0；规范版本须分层记录。',16,'gray')
#5 AMD
s=slide('Infinity Fabric：既连接芯粒，也可连接多个 GPU','MI300X 是具体例子：不能把这一代的直连拓扑或带宽，泛化为整个 Infinity Fabric 架构。',['S4','S5','S6'])
box(s,35,142,570,480,'white','line',True);imagefit(s,R/'research/mi300-node-level-arch.png',48,153,541,449)
text(s,58,606,530,20,'AMD 官方系统结构图：8 个 MI300X OAM + UBB。',14,'gray')
for y,h,c in [(150,'01 系统层：8 个 GPU，全互连','purple'),(310,'02 封装层：单个 MI300X 内部','purple'),(476,'03 主机层：PCIe 路径另行区分','orange')]:tag(s,636,y,h,c,593)
text(s,643,201,578,91,'每 GPU 用 7 条 GPU 间链路连接其余 7 个。\n图中红线：GPU 间 Infinity Fabric；\n不是中央交换芯片的星形拓扑。',21)
for j in range(8):gpu(s,644+j*71,364,'XCD','purple',59,34)
text(s,643,414,580,46,'8 XCD · 4 I/O Die · 8 HBM3 堆栈\n封装内部互连与 GPU 间链路不能合并计数。',20)
text(s,643,527,578,83,'图中黄线：PCIe Gen5 主机连接。\n图中浅蓝线：CPU 间 Infinity Fabric。\n同属一个系统，但链路对象与职责不同。',21)
#6 comparison
s=slide('横向比较先对齐对象：不要把链路、芯片与架构混排','这张表比较技术角色和边界，不是跨厂商性能排名。',['S1','S5','S6','S7'])
cols=[(238,235,'NVLink','blue'),(483,235,'NVSwitch','blue'),(728,245,'UALink','green'),(983,259,'Infinity Fabric','purple')]
for x,w,h,c in cols:tag(s,x,151,h,c,w)
rs=[('它是什么',['GPU 互连技术','NVLink 交换组件','开放互连规范','AMD 互连架构']),('连接/服务对象',['GPU 等端点','多个 NVLink 端点','加速器与交换结构','芯粒、GPU、CPU 等']),('范围看什么',['代际与平台','交换代际与系统','规范分层与实现','具体芯片与系统']),('本次真实例子',['GB200 的 NVLink 5','GB200 NVL72 交换域','200G 1.0 + 后续规范','MI300X 8-GPU 平台']),('不可直接推断',['共享显存自动透明','等同通用以太网交换机','所有厂商设备即插即用','任一代都用同一拓扑']),('搭配关系',['与 NVSwitch 配合','承载 NVLink 连接','由产品实现规范','随 AMD 平台演进'])]
for i,(h,vals) in enumerate(rs):
 y=208+i*65;box(s,35,y,1209,58,'pale' if i%2==0 else 'white','line',True);text(s,48,y+16,179,35,h,20,'ink',True)
 for (x,w,_,c),v in zip(cols,vals):text(s,x+7,y+12,w-14,47,v,18,c)
text(s,49,617,1180,28,'尤其注意：UALink 是开放标准，并不意味着它与 NVLink 或现有 Infinity Fabric 可以直接互插。',18,'gray')
#7 units
s=slide('带宽数字先看口径：per GPU、per link 与 per lane 不同','只比较同单位、同方向、同聚合范围的数据；峰值接口带宽也不等于 AllReduce 有效带宽。',['S1','S4','S5','S7'])
box(s,35,145,700,470,'pale','blue',True);tag(s,54,160,'同一官方表：NVLink 带宽 / GPU','blue',659)
for y,label,value,col in [(235,'NVLink 4 · Hopper',900,'blue'),(332,'NVLink 5 · Blackwell',1800,'blue'),(429,'NVLink 6 · Vera Rubin',3000,'purple')]:
 text(s,60,y-5,300,35,label,22,col,True);box(s,64,y+38,value/3000*470,28,col,col);text(s,552,y+31,154,45,f'{value:,} GB/s',22,col,True)
text(s,61,552,649,49,'按 NVIDIA 当前规格页列示；第六代标为初步规格。\n此处不将不同代际数据混入一个“通用 NVLink”参数。',18,'gray')
tag(s,770,153,'UALink 200G 1.0','green',458);text(s,780,209,445,49,'200G / lane',34,'green',True);text(s,780,266,443,70,'是单 lane 速率口径。\n还需确认 lane 数、方向及协议开销。',21)
tag(s,770,362,'MI300X 产品规格','purple',458);text(s,780,418,443,49,'128 GB/s / link',32,'purple',True);text(s,780,476,443,98,'产品页列 8 个 IF link；\n8-GPU 节点图用 7 条作 GPU 间互连。\n不能直接把“8 × 128”当成应用吞吐。',21)
#8
s=slide('读懂一套 GPU 互连，依次回答这四个问题','技术名称只给出起点；可用性能来自芯片能力、拓扑、软件与实际工作负载的共同作用。',['S1','S5','S6','S7'])
heads=['01 对象与版本','02 整体与部件','03 软件与语义','04 实测与边界'];bodies=['具体 GPU / 加速器型号\n链路与交换组件代际\n规范版本及实现状态','域内有多少端点\n直连还是交换式\n主机和域外路径在哪','可执行哪些通信操作\n运行时与通信库支持\n跨厂商互操作证据','实际消息大小和并行组\n有效带宽与尾延迟\n训练步时 / 推理延迟']
for i in range(4):
 x=35+i*311;tag(s,x,160,heads[i],['blue','green','purple','orange'][i],283);box(s,x,211,283,172,'pale','line',True);text(s,x+15,231,255,137,bodies[i],22)
 if i<3:line(s,x+283,178,x+305,178,'gray',True,2)
box(s,35,420,1215,200,'white','line',True)
text(s,57,440,1157,37,'把名字放回系统中，就不会混淆：',28,'ink',True)
for y,t,col in [(500,'NVIDIA：GPU — NVLink — NVSwitch — NVLink — GPU','blue'),(545,'UALink：开放规范 → 端点 / 交换组件 / 软件的具体实现','green'),(590,'AMD：Infinity Fabric → 具体芯粒、GPU 与系统拓扑','purple')]:text(s,64,y,1150,34,t,24,col,True)
# clean theme effects
for s in p.slides:
 for sh in s.shapes:
  style=sh._element.find(qn('p:style'))
  if style is not None:sh._element.remove(style)
O.mkdir(exist_ok=True);f=O/'GPU高速互连_NVLink_NVSwitch_UALink_InfinityFabric.pptx';p.save(f)
(R/'deck-map.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2))
(R/'sources.json').write_text(json.dumps([{'id':k,'url':v,'accessed':'2026-11-16','local_excerpt':'research/'+k+'.txt','status':'read'} for k,v in sources.items()],ensure_ascii=False,indent=2))
(R/'brief.json').write_text(json.dumps({'title':'GPU高速互连：NVLink、NVSwitch、UALink与Infinity Fabric','language':'zh-CN','aspect':'16:9','backend':'standalone python-pptx','pages':8,'diagram_detail_policy':'diagram-rich','audience':'技术与产业读者','scope':'技术层级、真实平台结构、规范版本与带宽口径','requirements':['原厂结构为据','左侧整体右侧详情用于硬件页','不把标准能力当部署','可编辑图形文字','静态小文件'],'project_root':'output/gpu-interconnect'},ensure_ascii=False,indent=2))
with zipfile.ZipFile(f) as z:assert z.testzip() is None
print(f,f.stat().st_size)
