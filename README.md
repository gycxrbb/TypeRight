# TypeRight — Word 智能排版插件

> 一键完成「识别段落语义角色 → 按排版方案批量重排 → 纯文本公式转公式对象」，全程可预览、一次 Ctrl+Z 完全撤销。

基于 PRD《Word智能排版插件_PRD.md》开发。当前覆盖里程碑 M1–M4 主体 + M5 的 LLM 客户端（默认关闭）。

## 快速开始

```powershell
# 安装（管理员：HKLM；非管理员：HKCU，均免 VSTO 运行时）
powershell -ExecutionPolicy Bypass -File tools\install.ps1

# 卸载
powershell -ExecutionPolicy Bypass -File tools\uninstall.ps1
```

安装后启动 Word → 功能区「智能排版」选项卡：

- **一键智能排版**：扫描全文 → 识别段落角色 → 预览确认（可逐段改角色）→ 应用
- **排版选区**：仅处理选中部分
- **方案管理**：编辑 / 导入 / 导出排版方案 JSON
- **设置**：LLM 增强识别（默认关闭，离线可用）、混排规则

首次建议：打开 `src\TypeRight.SmokeTest\out\smoke-result.docx`（或任意讲义/论文）→ 点「一键智能排版」。

诊断日志：`%LOCALAPPDATA%\TypeRight\addin-log.txt`；方案目录：`%APPDATA%\TypeRight\profiles\`。

## 技术形态（与 PRD 的差异说明）

| PRD 原选型 | 实际实现 | 原因 |
|---|---|---|
| VSTO Add-in | **NetOffice COM Add-in**（`NetOffice.WordApi.Tools.COMAddin`） | 本机未装 VSTO 工具链；PRD §2.1 已将 NetOffice 列为首选（免 VSTO 运行时/免 PIA/版本无关），与 Interop API 同构、后期可低成本切换 |
| WPF Task Pane | **NetOffice 原生 CTP 任务窗格**（WinForms UserControl + ElementHost 内嵌 WPF）+ 自动回落无模式窗口 | CTP 是 VSTO 之外唯一原生停靠窗格路径；NetOffice `Tools.COMAddin` 内建 `ICustomTaskPaneConsumer` 支持 |
| MML2OMML.XSL 兜底管线 | **已实现** `MathmlToOmml()`（XslCompiledTransform + Office 自带 XSL）+ `WrapLinearAsMtextMathml()`；主路径 BuildUp 失败时目前**无损保留原文**（PRD §6.2 的"失败不留残骸"） | BuildUp 覆盖了 UnicodeMath 全量；MathML 生成器属于 P2 导入路径，管线本体已就绪 |

## 架构

```
src/
├── TypeRight.Core/                  # 引擎（不依赖 UI）
│   ├── Roles/                       # 段落角色枚举（title/heading1..3/body/caption/listItem/tableText/formulaLine/skip + tableHeader）
│   ├── Scanning/                    # DocumentScanner（特征提取，含表格内段落）
│   ├── Classification/              # RuleTable（R1–R10 高置信）+ ClusterAnalyzer（R11 格式聚类）+ RoleClassifier（两段式编排）
│   ├── Formatting/                  # FormatApplier（UndoRecord 打包/幂等/保留强调）+ StyleMounter + FormulaConverter
│   ├── Formulas/                    # MathSegmenter（切分）+ UnicodeMathMapper（线性映射）— 纯逻辑可单测
│   ├── Mixing/                      # AutoSpacer（autocorrect/pangu 规则集移植：中西文空格 + 全半角标点）
│   ├── Profiles/                    # ProfileSchema + 校验器 + ProfileStore + 内置方案（嵌入资源）
│   └── Llm/                         # OpenAI 兼容客户端 + AmbiguityResolver（仅低置信段落，失败回落规则）
├── TypeRight.Addin/                 # COM 加载项：Ribbon + CTP 任务窗格 + 预览/设置/方案编辑 WPF 界面
└── TypeRight.SmokeTest/             # 真实 Word 端到端冒烟测试
tests/TypeRight.Core.Tests/          # 纯逻辑单测（50 个）
tools/install.ps1, uninstall.ps1     # per-user / per-machine 安装
```

## 关键设计

- **单次撤销**：`Application.UndoRecord.StartCustomRecord("TypeRight 智能排版")` 包裹全部改动（含样式定义修改）；执行异常时 `EndCustomRecord + Document.Undo()` 整体回退（冒烟测试已验证样式定义一并回退）。
- **幂等**：所有写入走 `SetIfChanged`（读-比-写）；公式区间跳过重转；同方案二次应用 0 改动（冒烟测试断言）。
- **保留强调**：只写字体族/字号与段落级属性，区间内局部加粗/颜色不触碰；挂样式段落用 `wdUndefined` 赋值清直接字体格式（首次运行自探测，失败自动回落直写目标值）。
- **字体分通道**：`eastAsiaFont → NameFarEast`，`latinFont → NameAscii + NameOther`，互不污染。
- **公式转换**：数学区之外的文本子区间才做字体操作（数学区由转换器管理 Cambria Math），片段降序处理避免偏移漂移；失败片段原样保留并计数。
- **角色识别**：R4 主标题用"正文聚类字号众数 +4pt"判定；R11 按（字号,加粗,颜色,对齐）签名聚类，占比最大簇为正文，大 1.5pt 以上的加粗簇按字号降序补位标题（每级限一簇）。

## 测试

```powershell
dotnet test TypeRight.sln                      # 50 个纯逻辑单测
dotnet run --project src\TypeRight.SmokeTest   # 真实 Word 端到端（识别/排版/撤销/幂等 23 项断言）
```

冒烟测试输出样例文档 `src\TypeRight.SmokeTest\out\smoke-result.docx` 供人工检查。

## 已知边界（P1 待办）

- 表格单元格内公式转换已启用，但公式字号取全局 `formula.sizePt`（与表格 9pt 文字不同号）
- 选区模式、图注域（题注交叉引用）、公式编号（书签+交叉引用）为 P2
- MML2OMML 兜底的 MathML 结构化生成器（矩阵/多级根号）待接入
- LLM 增强默认关闭；开启后 key 仅发往用户配置的 BaseURL
