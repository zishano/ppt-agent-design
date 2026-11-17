# 数据中心网络：Ethernet、InfiniBand、RoCE与Ultra Ethernet

8页中文16:9，原生可编辑拓扑、协议栈、控制反馈与数据通路图。

- [PPT](exports/数据中心网络_Ethernet_InfiniBand_RoCE_UltraEthernet.pptx)
- [PDF](exports/数据中心网络_Ethernet_InfiniBand_RoCE_UltraEthernet.pdf)
- [资料来源与口径](exports/资料来源与口径.md)
- [总览预览](renders/contact.jpg)

构建：`build.py`（python-pptx、Pillow），渲染需LibreOffice或PowerPoint及中文字体；`verify.py`（PyMuPDF、Pillow、python-pptx）检查PPT包、文字覆盖并生成预览，不替代人工语义和视觉检查。`validation/review.json`记录实际审查及产物哈希。

research独立保存规范、摘录和抓取脚本，不嵌入PPT；规范文件较大，交付PPT仍保持轻量。重建前请保留旧轮次。

## 修订版本

- [机制插画版 visual-v3（推荐）](visual-v3/README.md)：6页，以静态技术插画呈现数据路径、队列和反馈机制。
- [高信息密度版 dense-v2](dense-v2/README.md)：6页，原生可编辑图文。
- 本目录 exports 保留最初8页版。
