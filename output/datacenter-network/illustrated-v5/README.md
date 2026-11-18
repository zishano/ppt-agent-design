# 数据中心网络 · 机制插画版 v5

独立新修订，保留 visual-v3 和 story-v4。应用 mechanism-illustrated 风格：统一立体设备场景、原生可编辑数据路径/反馈、概念关系先导和末页对照。

- [PPTX](exports/数据中心网络_机制插画版_v5.pptx)
- [PDF](exports/数据中心网络_机制插画版_v5.pdf)
- [全页预览](renders/contact.jpg)

六页：概念关系、CPU/网卡分工、IB 逐跳流控、RoCEv2 拥塞反馈、UET 多路径与可靠交付、概念对照。共用 GPU A 向 GPU B 传 KV Cache 的例子。设备为通用示意，不是产品实机；端口、部件和路径数不构成规格。GPU 直达假定平台支持及内存已授权。UE 按 1.1 规范讨论，具体产品和 profile 能力需另行确认。

引用上级 sources.json 的已核查资料（2026-11-17），完整 URL 在 PPT 备注中；evidence-manifest.json 保留来源记录身份。本次未声称重新联网核查所有来源。场景为 imagegen 生成的静态 PNG，提示词保留在 assets/prompt.txt；仅设备插画是栅格图，标题、表格、关系图与机制连线均可编辑。不含动态媒体。

使用 python-pptx 运行 build.py；LibreOffice 导出 PDF。实际渲染检查、文本覆盖、ZIP 与静态媒体检查见 validation/report.json。PPT 约 1.72 MB。LibreOffice 结果不保证与所有 PowerPoint 版本完全一致。
