# AI 集群通信需求与 Scale-up、Scale-out 网络的区别

8 页中文、16:9，以原生可编辑网络拓扑和通信过程图解释需求、边界、协同与验证。

- [PPT](exports/AI集群通信_Scale-up与Scale-out.pptx)
- [PDF](exports/AI集群通信_Scale-up与Scale-out.pdf)
- [资料来源](exports/资料来源.md)
- [页面总览](renders/contact-final.jpg)
- [验证记录](validation/report.json)

构建入口为 `build.py`，验证入口为 `verify.py`。脚本从自身位置解析项目路径，需要 python-pptx、Pillow、PyMuPDF；研究抓取另需 requests、beautifulsoup4。PDF 渲染需另行安装 LibreOffice 或 PowerPoint，并配置中文字体。

`verify.py` 只执行包、文本覆盖与预览生成检查，不代替语义和视觉审查。运行时会写入验证记录；重新构建应先复制到新的版本目录，保留此前交付及检查记录。图中端点数量通常为逻辑示意；具体事实来源见页备注与证据文件。
