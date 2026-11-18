# 数据中心网络 · 完整机制插画版 v6

按用户明确指定的 visual-v3 风格重新绘制，保留之前版本。主题为 Ethernet、InfiniBand、RoCE 与 Ultra Ethernet。

- [PPTX](exports/数据中心网络_完整机制插画版_v6.pptx)
- [PDF](exports/数据中心网络_完整机制插画版_v6.pdf)
- [六页预览](renders/contact.jpg)

## 画风与内容

整页完整技术插画：设备剖面、板卡、内存区域、数据块、队列、反馈回路与就地标注融合在主图内。局部状态图解释授权、缓冲变化、ECN/CNP、PFC、多路径与完成条件。不是设备贴图配原生文字卡片。

六页依次讲概念关系、CPU/网卡分工、IB 逐跳流控、RoCEv2 拥塞反馈、UET 多路径与可靠交付、三问对照。以同一块 KV Cache 从 GPU A 向 GPU B 的传输贯穿机制页，保留此前关于概念层级与复述的要求。

全部为静态图片；主图（包括图内文字与连线）不可逐部件编辑，只有来源页脚和备注为原生可编辑元素。已同步修订 mechanism-illustrated 风格规则的默认制作方式。无视频、GIF 或动画。

## 事实与复现

复用上级 sources.json 的官方证据，核查日期仍记为 2026-11-17，不声称本次重新联网核查。逐页 URL 在备注中，证据身份见 evidence-manifest.json。所有设备、缓冲格、路径数和数据块数为机制示意，不是实机规格。GPU 直达假定平台支持、内存已授权；UET 语义按 profile，规范能力不等于实际产品能力。

prompts.json 与 corrections.json 保存生成和修订指令；assets/ 包含审阅后的原始 PNG 和用于嵌入的压缩 JPEG。用 python-pptx、Pillow 运行 build.py；LibreOffice 导出 PDF，PyMuPDF 生成实际渲染。图内文字必须视觉检查，PDF 文本提取不能证明栅格标签正确。
