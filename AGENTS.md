# AGENTS.md — TypeRight 开发规则与上下文

> 本文件面向 AI 助手（Claude/Copilot/Cursor 等）及新加入的开发者，定义本项目的开发规则、已知坑位与关键架构决策。

## 项目一句话

Word/WPS COM 加载项（.NET 4.8 + NetOffice），"指认式批量排版"：光标点样本段 → 格式签名匹配同类 → 浮条批量改格式 → 可存方案回放。

## 技术栈

| 组件 | 技术 |
|---|---|
| 框架 | .NET Framework 4.8（Win7 SP1+ 可装） |
| COM 互操作 | NetOffice 1.9.12（WordApi/OfficeApi/VBIDEApi） |
| UI | WPF（浮动条/进度窗/Toast）；Ribbon XML 自定义功能区 |
| 测试 | xUnit 45 个单元测试 + TypeRight.Acceptance 21 用例真机验收 |
| 发布 | GitHub Actions tag 触发 → package-release.ps1 → zip |

## 构建与部署

```powershell
# 单元测试
dotnet test tests\TypeRight.Core.Tests\TypeRight.Core.Tests.csproj

# 开发安装（构建→%LOCALAPPDATA%\TypeRight\Addin→HKCU 注册）
powershell -ExecutionPolicy Bypass -File tools\install.ps1

# 全量验收（21 用例，约 30-40 分钟）
tests\TypeRight.Acceptance\bin\Debug\TypeRight.Acceptance.exe

# 打安装包
powershell -File tools\package-release.ps1

# 发布：git tag v0.3.1 && git push origin v0.3.1
```

## 必须遵守的规则

### 1. PS1 含中文必须 UTF-8 BOM

PowerShell 5.1 把无 BOM 的 UTF-8 当 ANSI 读，中文路径/字符串全毁。

```
检查: python -c "open(f,'rb').read(3)==b'\xef\xbb\xbf'"
修复: python -c "d=open(f,'rb').read(); open(f,'wb').write(b'\xef\xbb\xbf'+d)"
```

### 2. 编辑 Ribbon.xml 后必须跑 XML 校验

Word 2610 对 customUI 校验极严但**零报错**——属性重复、`size="small"` 等都会导致功能区整块静默消失。

```powershell
python -c "import xml.etree.ElementTree as ET; ET.parse(r'src\TypeRight.Addin\Ribbon.xml'); print('OK')"
```

日志判据：`GetCustomUI → N 字符` 后有无 `Ribbon onLoad`（有 = 采纳；无 = 被拒）。

### 3. 验收套件必须单例运行

两条套件链并发 = 互杀 Word 进程 + 互写 tests/output。启动前确认无残留 WINWORD。

### 4. 杀 Word 前先问用户

部署/验证链不得默认 `Stop-Process WINWORD`。用户可能有未保存文档。

### 5. Bash heredoc 写含转义的代码用 Write/Edit 工具或 python 文件

Git Bash 的 heredoc 吃反斜杠和 `$`。用 Write 工具写 python 脚本再执行，或用 Write/Edit 工具直接改源文件。

### 6. 不要提交测试垃圾

`tests/output/*.docx`、`tests/output/*.log`、`~$*.docx` 是运行产物，不进 git。

## 架构速览

```
src/TypeRight.Core/          ← 无 UI 核心库
├── Signature/               ← 格式签名匹配引擎
│   ├── FormatSignature.cs   ← 5锚点模型（首/25%/50%/75%/尾 × 字体·字号·加粗·颜色）
│   ├── FormatSignatureReader.cs ← 从 Paragraph/Range 提取签名
│   ├── SimilarSelector.cs   ← 全文扫描匹配 + 高亮管理器 + 引导句感知
│   ├── TextMatcher.cs       ← 字义匹配（文本同类，2026-11 新增）
│   └── BatchFormatter.cs    ← 批量格式执行（幂等写入 + UndoRecord）
├── Formatting/              ← 公式转换
│   ├── FormulaConverter.cs  ← 三阶段流水线（规划→批量注入→O(1)终检）
│   └── WordFormatUtils.cs   ← 幂等写入工具
├── Formulas/                ← MathML → OMML 注入
│   └── MathmlInjector.cs    ← 全部公式一个 docx + 隐藏草稿一次装载 + 降序 FormattedText
├── Profiles/                ← 方案存回放
├── Settings.cs              ← 用户设置（settings.json）
└── WordCompat.cs            ← 宿主兼容层（老Word降级 + WPS集合枚举 + UndoRecord）

src/TypeRight.Addin/         ← COM 加载项
├── TypeRightAddin.cs        ← 入口：Ribbon XML / 事件订阅 / 自动化钩子
├── Controllers/
│   └── TypeRightController.cs ← 编排器：同类选择/文本同类/公式转换/方案/跳转
├── UI/
│   ├── FloatBar.xaml(.cs)   ← 浮动格式条（Fluent 换皮 2026-11）
│   ├── ConvertProgressWindow.cs ← 公式转换进度窗（三层反馈 + Esc 取消）
│   ├── ToastWindow.cs       ← 轻提示
│   ├── RibbonIcons.cs       ← WPF 矢量图标 → IPictureDisp (Alpha DIB)
│   └── OfficeTheme.cs       ← 宿主主题检测（PrintWindow 遮挡免疫）
└── Ribbon.xml               ← 功能区/右键菜单定义
```

## 关键设计决策（为什么这样）

| 决策 | 原因 |
|---|---|
| 签名用 5 锚点而非全段逐字符 | Font.Name* 有读取抖动（MS Mincho↔微软雅黑），5 锚点中段只比字号/加粗/颜色 |
| 引导句感知（混合段拆两单元） | 用户"行首标题句+正文"混合段需独立匹配 |
| 公式区感知（锚点跳过 OMML） | 公式区内 Font 读数是数学字体，污染签名 |
| 兼容枚举器（WordCompat.*Of） | NetOffice 枚举器在 WPS 内抛 FactoryException；但 Word 的 OMaths 枚举又是病态慢路径 → 索引器优先+枚举兜底 |
| Shading 而非 HighlightIndex 做强调色 | Shading 支持任意 RGB（柔和色板），高亮只有 16 色荧光 |
| GetAsyncKeyState 而非消息泵做 Esc | 转换在 UI 线程上跑，消息循环被占满 |

## 宿主兼容矩阵

| 宿主 | 状态 | 注意事项 |
|---|---|---|
| Word 2016~365 (16.0) | ✅ | 主力测试平台 |
| Word 2013 (15.0) | ✅ | UndoRecord/Hwnd 都有 |
| Word 2010 (14.0) | ✅ 兼容模式 | 无 UndoRecord（多条撤销）；Hwnd 需 FindWindow 兜底 |
| WPS Office | ✅ | 32 位进程需 HKLM Wow6432Node CLSID；OMath 缺 BuildUp（公式转换线性段兜底失败）；Font.HighlightColorIndex 读不存在 |
| Win7 SP1+ | ✅ | PS2.0 无 $PSScriptRoot（安装脚本已垫）；GetDpiForWindow 不存在（GDI DpiX 兜底） |
| Word 2007 及更早 | ❌ | 右键菜单 XML 不支持 |
| Mac / 网页版 | ❌ | 平台限制 |

## 更多文档

- [docs/architecture/](docs/architecture/) — 系统架构
- [docs/features/](docs/features/) — 功能规格
- [docs/compatibility/](docs/compatibility/) — 兼容详情
- [docs/troubleshooting/](docs/troubleshooting/) — 排障手册
- [docs/testing/](docs/testing/) — 测试策略
- [docs/release/](docs/release/) — 发布流程
- [DESIGN.md](DESIGN.md) — 设计系统（令牌/圆角/状态语义）
- [PRODUCT.md](PRODUCT.md) — 产品上下文
