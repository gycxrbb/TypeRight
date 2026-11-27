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
| 测试 | xUnit 315 个单元测试 + TypeRight.Acceptance 21 用例真机验收 |
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

# 发布：git tag v0.5.0 && git push origin v0.5.0
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
│   ├── FormatSignature.cs   ← 5锚点模型（首/25%/50%/75%/尾 × 字体·字号·加粗·颜色·字符类型）
│   ├── FormatSignatureReader.cs ← 从 Paragraph/Range 提取签名
│   ├── SimilarSelector.cs   ← 全文扫描匹配 + 高亮管理器 + 引导句感知
│   ├── TextMatcher.cs       ← 字义匹配（文本同类，2026-11 新增）
│   ├── CharClass.cs         ← 主导字符类型判定（中文/字母/数字/混合，2026-11 新增）
│   └── BatchFormatter.cs    ← 批量格式执行（幂等写入 + UndoRecord）
├── Formatting/              ← 公式转换
│   ├── FormulaConverter.cs  ← 三阶段流水线（规划→批量注入→O(1)终检）
│   └── WordFormatUtils.cs   ← 幂等写入工具
├── Formulas/                ← MathML → OMML 注入
│   └── MathmlInjector.cs    ← 全部公式一个 docx + 隐藏草稿一次装载 + 降序 FormattedText
├── Forms/                   ← 填表防崩套件（v0.4.0 新增）
│   ├── PhotoCellFitter.cs   ← 证件照严丝合缝入格（两步等比缩放+尺寸微调，零裁切，四边0.75pt内缩）
│   ├── CellTextBoxPlacer.cs ← PPT式悬浮填表框（表外同页就近锚定，保留原示例与底端签字，零抖动）
│   └── FormGeometry.cs      ← 单元格与纵向合并格极速几何测绘（扣除行距偏移 + 最少 Information 调用）
├── Sequences/               ← 序列重排（v0.5.0 新增）
│   └── SequenceRenumberer.cs ← 段首序号解析(9类) + 断1切流 + 只动断序 + 选区整流处理
├── Classification/          ← 段落角色规则表
│   └── RuleTable.cs         ← 角色/规则判定表
├── Roles/                   ← 段落角色模型
│   └── ParagraphRole.cs     ← 段落语义角色（标题/正文/引导句…）
├── Scanning/                ← 文档扫描
│   ├── DocumentScanner.cs   ← 全文单遍扫描
│   └── ParagraphFeatures.cs ← 段落特征提取
├── Profiles/                ← 方案存回放
├── Settings.cs              ← 用户设置（settings.json）
└── WordCompat.cs            ← 宿主兼容层（老Word降级 + WPS集合枚举 + UndoRecord）

src/TypeRight.Addin/         ← COM 加载项
├── TypeRightAddin.cs        ← 入口：Ribbon XML / 事件订阅 / 自动化钩子
├── Controllers/
│   └── TypeRightController.cs ← 编排器：同类选择/文本同类/公式转换/照片入格/悬浮填表框/序列重排/方案/跳转
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
| 序列重排"断 1 开新流 + 单条目不动" | 与"大题重新起号/行内序号续文"无法区分的场景一律保守不动；只修确定的断序（缺号/重号/乱序） |
| 序列重排只认纯逻辑层可测的标记 | 解析/切流/断序判定全部抽成无 Word 依赖纯函数（xUnit 直测），Word 层只剩枚举与写入 |
| 字号显示走 ChineseFontSizes 纯逻辑层 | 中文字号（四号/小四）映射表无 Word 依赖 → xUnit 直测；非标准档位（12.5）仍显示数值，不硬套中文名 |
| 未知字号（0/NaN）一律显示 — 而非 0 | 文本同类模式的 SampleSignature 是 default(struct)，SizePt 落在 0；字号恒为正，非正数即"读不到" |
| DPI 拿到真值才缓存 | PositionNear 在 Show() 之前调用时没有 HwndSource，缓存兜底 1.0 会让 125% 缩放的浮条定位到屏幕外（实测 Top=918/屏高 864） |
| GetPoint 返回负文档坐标即退回窗口居中 | 段落滚出视口时 GetPoint 吐 -31493 这类值；据此定位会把浮条丢到屏幕外 |
| 字体/字号框用 IsEditable ComboBox | 打字录入与点选档位两条路径都要留；`IsTextSearchEnabled=False` 否则每敲一键就自动选中候选并立刻改全文格式 |
| 字号非法输入不静默套默认值 | 打「乱码」时还原显示 + Toast；悄悄改成别的字号比报错更糟。范围 1~1638 磅，越界判非法 |
| 字体名不做"已安装"校验 | 主题字体（+中文正文）本来就不在 InstalledFontCollection 里，按存在性拒绝会挡掉合法场景 |
| 打字中 Esc ≠ 关会话 | 输入框内 Esc 是"放弃本次编辑"；EscHook 与 PreviewKeyDown 两处都要判 IsEditingText |
| 高亮染色移出点击路径（Background 队列） | 43 处命中 = 129 次 COM 往返 ≈ 550ms；同步做会让"同类选择"感觉沉。浮条先弹、黄标后补 |
| COM 加载项里用浮条自身的 Dispatcher | `Application.Current` 在 COM 加载项里通常为 null，用它会掉进同步回退，把 500ms 又塞回点击路径 |
| 主题检测不 force、PrintWindow 只抓窄带 | 整窗 PrintWindow 实测 164ms/次；每次浮条打开都 force 重测会让点击明显发顿 |
| 浮条必须显式声明 ClearType + LayoutRounding + SnapsToDevicePixels | 缺这三项时文字退化为灰度抗锯齿、1px 边框被拉成虚边（实测锐度差 3 倍）；`AllowsTransparency=True` 本身就会禁用子像素抗锯齿 |
| **不要试图改进程 DPI 感知** | Word.exe 是 System-DPI aware，awareness 在首个窗口创建时冻结。DLL manifest / SetProcessDpiAwareness* / SetThreadDpiAwarenessContext 三条路实测全部无效，见 docs/troubleshooting/performance.md |
| 字符类型维度用 CharClass 纯逻辑层，不靠 Font.NameAscii | 实测 Word 按内容自动归一化 NameAscii（纯中文段回落到中文字体），那是渲染回落不是作者意图 → 必须直接看字符本身。纯字符串运算，零 COM 成本 |
| **数字压过汉字 → 判数字段**（日期/数据类） | 实测 `2026年11月26日` 汉字3/数字8，与正文汉字114/数字49 的**字体字号完全相同**，字体维度区分不了。若按"有汉字即中文"，日期落款会被卷进正文一起改格式（考察报告实测 13→12 处） |
| **匹配整段、刷写按字符类型** | 用户口径"汉字归汉字、数字归数字"：匹配/计数仍是整段（不改模型），但高亮与改格式只作用于目标字符类型的连续片段。点汉字只动汉字，日期里的数字保持原样 |
| 目标类型取光标所在字符的类型 | `SelectSession.TargetClass`；落标点/空白时不过滤（退回旧行为），避免"点了逗号什么都选不到" |
| `FilterRuns` 遇到异类片段必须 flush | 回归坑：第一版只置 has=false 没把已累积片段存进结果 → 前面片段全丢，真机表现为"过滤后什么都没改" |
| 字符类型未知（None）时不阻止匹配 | 未知不该成为"不同类"的理由，否则旧调用点/读取抖动造成漏命中。区别于字号/加粗的"未知不相等"——那些是数值抖动，这里是这一维根本没采 |
| **方案规则必须按匹配条件合并**（`OpNormalizer`） | 一次会话里连着改字体→字号→加粗会产生若干条 match 完全相同的规则。回放是「每段只被首条命中规则处理」，不合并则后面几条命中 0 段，用户改的字号被静默吞掉（实测 3 条同样 match 的规则 → 只回放了第一次） |
| 方案/窗口文案一律"人话"，不暴露内部概念 | 用户面对方案只关心「什么样的段落 → 什么属性从 X 改成 Y」。禁止显示"规则 N 条"、`ConvertFormulas` 布尔值、"N 步批量操作"这类实现细节 |
| **存为方案与管理方案分窗** | 两个语境混一屏会让界面自相矛盾（标题写"存为方案"、列表里却是"应用到当前文档"）。`SaveProfileWindow` 只负责存，`ProfileManagerWindow` 只负责用和管理 |
| **所有独立窗口继承 `ThemedWindow`** | 颜色唯一真值 = `ThemePalette`；窗口底/文字必须 `SetResourceReference` 绑定，禁止 `Background = Brushes.White`，禁止 `new SolidColorBrush(P.Xxx)` 静态快照（不会跟随主题切换） |
| **`ApplyTheme` 内不得直接给 Background 赋值** | 直接赋值会顶掉构造函数里 `SetResourceReference` 建立的绑定，之后换主题永久失效（实测 ApplyTheme(Dark) 后底色仍是 #FAFAFA） |
| 资源令牌刷新必须 Remove 再 Add | 窗口可见后 `Resources[key] = b` 对同键**不触发** DynamicResource 重新求值，令牌值变了但所有引用处不动 |
| 主题验证必须用真实窗口类 | 自拼"等价"探针会骗自己：曾在探针里手动补 `win.Background` 与颠倒 Show/ApplyTheme 顺序，连续两次漏掉真 bug。回归断言见 `UiThemeTests` |

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
