# 数据中心网络：关系与机制版

6 页、16:9、静态 PPT。按用户四点反馈修订：概念关系优先、答案式标题、同一 KV Cache 案例贯穿、末页复述表。保留此前版本。

- [PPTX](exports/数据中心网络_关系与机制版.pptx)
- [PDF](exports/数据中心网络_关系与机制版.pdf)
- [全页预览](renders/contact.jpg)

## 阅读顺序
1. 两种网络体系；RoCE 和 Ultra Ethernet 在 Ethernet 路线中；UET 位于 UE 内部。
2. CPU 准备资源，网卡搬运同一块 KV Cache。
3. IB：逐跳缓冲流控。
4. RoCEv2：ECN 标记 → 接收端 CNP → 源端降速；UDP 为封装。
5. UET：多路径分发、可靠交付与拥塞反馈；能力取决于 profile 和实现。
6. 它是什么、基于什么网络、解决什么问题。

## 绘制与真实性
所有标题、关系框、机制路径、反馈线和对照表均为可编辑 PowerPoint 元素。仅通用服务器为 AI 生成的静态插画，不代表真实产品、BOM 或 SpaceX 设备。中间四页使用相同位置与语义颜色。示例假定内存已注册授权且支持 GPU 显存直达；数据分块、路径数量和缓冲格数都是机制示意。

复用上级目录已核查的官方资料（2026-11-17），具体 URL 在每页备注中，页脚列出来源 ID。UET 和 RoCE 属于不同传输路线，不能假设同一端点硬件可任意切换。UE 1.1 规范能力不等于所有产品已实现。

## 复现
使用含 python-pptx、Pillow 的 Python 运行 build.py；LibreOffice 导出 PDF，PyMuPDF 渲染。无视频或动画。validation/report.json 记录结构检查与渲染审查结果。
