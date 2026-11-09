# 兼容矩阵：Word 版本 × WPS × Windows

## Microsoft Word 桌面版

| 版本 | 内部版本 | 状态 | 差异 |
|---|---|---|---|
| 365 / 2021 / 2019 / 2016 | 16.0 | ✅ 全功能 | 主力测试平台 |
| 2013 | 15.0 | ✅ 全功能 | UndoRecord/Hwnd 都有 |
| 2010 | 14.0 | ✅ 兼容模式 | 无 UndoRecord（撤销粒度退化为多条记录，首次启动提示一次）；Window.Hwnd 需 FindWindow("OpusApp") 兜底 |
| 2007 及更早 | 12.0- | ❌ | 右键菜单 XML（contextMenus 元素需 2010+） |

## WPS Office

| 项目 | 详情 |
|---|---|
| 支持版本 | 个人版/教育版 12.x（32 位进程） |
| 加载方式 | 读 Microsoft Office Word Addins 注册键（与 PDFMaker/YunOffice 同路） |
| 32 位 CLSID | 需 HKLM Wow6432Node（用户级 HKCU Wow6432Node 实测不被 Windows 解析） |
| 安装 | install-wps-elevated.ps1（UAC 一次性提权，可选步骤） |

### WPS 已知差异

| 功能 | 状态 | 说明 |
|---|---|---|
| 同类选择（格式匹配） | ✅ | 全链路实测可用 |
| 文本同类（字义匹配） | ✅ | 纯文本操作 |
| 批量格式 + 高亮 + 浮条 | ✅ | 完整 UI 体验 |
| 方案存/回放 | ✅ | |
| 公式转换 | ⚠️ 实验性 | 主管线（docx 注入）可用；OMath 对象缺 BuildUp → 线性片段兜底失败（保留原文本） |
| 功能区图标 | ✅ | Alpha DIB 路径稳定 |
| 主题跟随 | ✅ | PrintWindow 抓 WPS 窗口 + 回退默认 Light |
| Shading 强调色 | ⚠️ | WPS 无 Shading → 兜底青色高亮（wdTurquoise） |
| Font.HighlightColorIndex 读 | ❌ | 属性不存在（写/clear 可用）；产品用的是 Range 路径 |

### WPS 技术要点

- `app.Version` 自报 "12.0"（伪装 Word 2007 对象模型）
- NetOffice 集合枚举器抛 `FactoryException: Class does not exist: .Paragraph` → `WordCompat.*Of` 底层 IEnumVARIANT + 手工包装
- OMaths 索引器也走工厂包装 → 同样需枚举兜底
- 功能区自绘（KPromeMainWindow），UIA 看不见选项卡 → 验证 UI 用肉眼或问用户

## Windows 版本

| 版本 | 状态 | 注意事项 |
|---|---|---|
| Windows 11 | ✅ | |
| Windows 10 1903+ | ✅ | |
| Windows 10 更早 | ✅ | 需手动装 .NET 4.8（安装器自动检测） |
| Windows 8.1 | ✅ | 同上 |
| Windows 7 SP1 | ✅ | PS 2.0 无 $PSScriptRoot（安装脚本已垫）；GetDpiForWindow 不存在（GDI DpiX 兜底） |
| Windows 7 无 SP1 | ❌ | .NET 4.8 不支持 |
| Mac | ❌ | 平台限制 |
