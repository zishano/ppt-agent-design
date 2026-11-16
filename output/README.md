# PPT 项目输出目录

以后生成的 PPT 项目统一放在本目录：`output/<项目名>/`。用户明确指定其他路径时，以用户要求为准。

每个项目集中保存需求、调研证据、绘图代码、PPT/PDF、预览与验证记录：

```text
output/<项目名>/
  brief.json
  sources.json / claims.json / deck-map.json
  build.py
  research/
  exports/       # 对外提供的 PPT、PDF 和来源说明
  renders/       # 实际导出文件的页面预览
  validation/    # 检查记录与产物身份
```

- 同一主题的新轮次使用独立版本目录，保留已有交付。
- 脚本以自身位置解析项目根路径，避免绑定 `/workspace`。
- PPT 与必要证据可以纳入仓库；本地环境、凭据、渲染器配置和缓存不纳入版本控制。
- 文件移位后核对来源链接、脚本路径和交付文件哈希；迁移不等于重新生成或重新通过视觉验收。

## 当前项目

- [AI 集群通信与 Scale-up / Scale-out](ai-network-deck/README.md)：8 页中文、可编辑拓扑与通信示意。
