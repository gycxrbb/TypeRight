# PRD：TypeRight —— Word 智能排版插件（VSTO）

> 项目名：TypeRight（"type + right"，排版一步到位）

> 版本 v1.0 ｜ 首发场景：论文 / 课程报告 / 讲义 ｜ 技术路线：VSTO（C# / .NET Framework 4.8）
> 本文档直接作为开发输入，交给编码 agent 实现。

---

## 1. 产品定位与核心痛点

**痛点**：绝大多数 Word 文档用「直接格式」（手动刷字体）而非样式写成，全文格式统一只能逐段手改；中西文字体规范（汉字宋体 / 西文 Times New Roman / 公式 Cambria Math 斜体）几乎无人手工做对。

**产品**：Word 本地插件，一键完成「识别段落语义角色 → 按排版方案批量重排 → 纯文本公式转公式对象」，全程可预览、可一键撤销。

**设计原则（优先级从高到低）**：
1. 确定性优先：规则引擎为主，可解释、可调试；LLM 仅作可选增强，缺省关闭、离线可用
2. 不破坏原则：任何情况下不损坏原文档结构（表格、图片、域、批注），失败必须整体回退
3. 预览先行：任何批量修改必须先展示「将发生什么」，用户确认后才执行

---

## 2. 技术架构

### 2.1 技术选型

| 项 | 选型 | 理由 |
|---|---|---|
| 插件框架 | VSTO Add-in（C#，.NET Framework 4.8） | 唯一支持 OMML 操作、自定义撤销单元、样式挂载的本地路线 |
| Office 互操作 | **首选 NetOfficeFw**（MIT，`github.com/NetOfficeFw/NetOffice`）；备选 Microsoft.Office.Interop.Word | 两者 API 完全同构、可互换；NetOffice 免 VSTO 运行时、免 PIA、免注册表，单 dll 复制即用，版本无关（Office 2010–365 通吃），显著降低安装分发摩擦。先用 Interop 开发也可，后期切换成本低 |
| UI | Ribbon XML（按钮组）+ WPF Task Pane（预览与设置） | VSTO 标准形态 |
| 方案存储 | JSON 文件（Newtonsoft.Json），`%APPDATA%\TypeRight\profiles\` | 可导入导出、可版本管理 |
| LLM 增强（可选） | OpenAI 兼容 HTTP 接口，用户自带 API key | 不绑定厂商 |
| 安装分发 | VSTO ClickOnce 或 MSI（后期再定）；若用 NetOffice 可简化为免安装包 | 首发可用 ClickOnce |

### 2.2 关键 API 能力备忘（实现时直接查这几条）

| 能力 | Word 对象模型 API |
|---|---|
| 中西文字体分离设置 | `Range.Font.NameFarEast`（中文）/ `Range.Font.NameAscii`、`NameOther`（西文） |
| 自定义撤销单元 | `Application.UndoRecord.StartCustomRecord("智能排版")` … `EndCustomRecord()`——整次排版打包成一次 Ctrl+Z |
| 线性公式转专业格式 | `Range.OMaths.Add(range)` 后加入线性格式文本（UnicodeMath，如 `x_1^2`、`(-b±√(b^2-4ac))/(2a)`），再 `OMath.BuildUp` 自动转为堆叠分式/根号/上下标。**不要手搓 OMML XML** |
| 挂载内置样式 | `Paragraph.Style = WdBuiltinStyle.wdStyleHeading1` 等 |
| 样式定义修改 | `Document.Styles["标题 1"].Font.NameFarEast = ...` |
| 段落遍历性能 | 禁用 `ScreenUpdating`、批量收集后统一写入；百页 < 10s |

### 2.3 模块划分

```
TypeRight/
├── Ribbon.xml / Ribbon.cs          # 功能区入口
├── Core/
│   ├── RoleClassifier.cs           # 角色识别引擎（规则）
│   ├── RuleTable.cs                # 规则优先级表（可数据驱动）
│   ├── FormatApplier.cs            # 排版执行器（含撤销单元打包）
│   ├── FormulaConverter.cs         # 纯文本公式 → OMath
│   ├── StyleMounter.cs             # 识别结果挂载 Word 样式
│   └── DocumentScanner.cs          # 段落/表格遍历与特征提取
├── Profiles/
│   ├── ProfileSchema.cs            # 方案数据结构
│   ├── ProfileStore.cs             # 方案读写/导入导出
│   └── BuiltIn/                    # 内置方案 JSON
├── Llm/
│   ├── LlmClient.cs                # OpenAI 兼容客户端
│   └── AmbiguityResolver.cs        # 仅对低置信段落调用
└── UI/
    ├── PreviewPane.xaml            # 预览确认（任务窗格）
    ├── ProfileEditor.xaml          # 方案编辑器（P1）
    └── SettingsPane.xaml           # LLM key 等设置
```

---

## 3. 核心数据结构：排版方案（Profile）JSON Schema

```json
{
  "schema_version": 1,
  "id": "academic-lecture",
  "name": "学术讲义/论文",
  "description": "首发内置方案：论文、课程报告、讲义通用",
  "roles": {
    "title": {
      "label": "文档主标题",
      "eastAsiaFont": "黑体",
      "latinFont": "Times New Roman",
      "sizePt": 16,
      "bold": true,
      "color": "#1A365D",
      "alignment": "center",
      "wordStyle": "Title",
      "spaceBeforePt": 12,
      "spaceAfterPt": 12
    },
    "heading1": {
      "label": "一级标题",
      "eastAsiaFont": "黑体",
      "latinFont": "Times New Roman",
      "sizePt": 15,
      "bold": true,
      "color": "#1A365D",
      "wordStyle": "Heading 1",
      "outlineLevel": 1
    },
    "heading2": {
      "label": "二级标题",
      "eastAsiaFont": "黑体",
      "latinFont": "Times New Roman",
      "sizePt": 13,
      "bold": true,
      "color": "#2B6CB0",
      "wordStyle": "Heading 2",
      "outlineLevel": 2
    },
    "heading3": {
      "label": "三级标题",
      "eastAsiaFont": "黑体",
      "latinFont": "Times New Roman",
      "sizePt": 12,
      "bold": true,
      "wordStyle": "Heading 3",
      "outlineLevel": 3
    },
    "body": {
      "label": "正文",
      "eastAsiaFont": "宋体",
      "latinFont": "Times New Roman",
      "sizePt": 10.5,
      "lineSpacing": 1.5,
      "firstLineIndentPt": 21,
      "wordStyle": "Normal"
    },
    "caption": {
      "label": "图注/表注",
      "eastAsiaFont": "宋体",
      "latinFont": "Times New Roman",
      "sizePt": 9,
      "color": "#595959",
      "alignment": "center",
      "wordStyle": "Caption"
    },
    "listItem": {
      "label": "列表项",
      "eastAsiaFont": "宋体",
      "latinFont": "Times New Roman",
      "sizePt": 10.5,
      "wordStyle": "List Paragraph"
    },
    "tableText": {
      "label": "表格文字",
      "eastAsiaFont": "宋体",
      "latinFont": "Times New Roman",
      "sizePt": 9
    }
  },
  "formula": {
    "convertPlainText": true,
    "mathFont": "Cambria Math",
    "sizePt": 10.5
  },
  "mixingRules": {
    "autoSpaceCjkLatin": false,
    "unifyPunctuation": "none",
    "digitsToLatinFont": true
  }
}
```

**混排规则实现说明**：`autoSpaceCjkLatin`（中西文间加空格）与 `unifyPunctuation`（全半角统一）不自行发明规则，移植 **huacnlee/autocorrect** 与 **vinta/pangu.js** 的成熟规则集为 C# 实现（核心为空格插入与标点映射的正则集，规模数十行），详见 §12。

**规则**：
- `eastAsiaFont` 只写 `Font.NameFarEast`，`latinFont` 只写 `Font.NameAscii` + `NameOther`，互不污染
- `wordStyle` 字段存在时，应用格式 = 先挂样式、再把字体属性写入样式定义（全文联动），而非逐段直接格式
- 方案文件校验失败（缺字段/字体名非法）时拒绝加载并提示，不 fallback 到半成品

---

## 4. 角色识别引擎（规则优先级表）

### 4.1 角色枚举

`title / heading1 / heading2 / heading3 / body / caption / listItem / tableText / formulaLine / skip`

### 4.2 识别流程（两段式）

**第一段：特征提取**——对每个段落（含表格单元格内段落）提取特征向量：

| 特征 | 来源 |
|---|---|
| 文本长度、是否为空 | Paragraph.Range.Text |
| 首行字号 / 加粗 / 颜色 / 对齐 | Range.Font 各属性 |
| 编号模式 | 正则：`^第[一二三四五六七八九十\d]+[章节模块]`、`^\d+\.\d+(\.\d+)?\s`、`^[一二三四五六七八九十]+、`、`^[①-⑳]`、`^【.*】` |
| 是否已有样式 | Paragraph.Style（已有 Heading 的直接信任） |
| 是否在表格内 / 表格行位置 | Range.Information |
| 数学字符密度 | 字母+数学符号（= + - ± √ ≤ ≥ ≠ ⟹ 上下标字符等）占比 |
| 图注模式 | `^(图|表|Figure|Table|【图)` 开头，或引用题注域 |

**第二段：判定**，严格按以下优先级，命中即停：

| 优先级 | 规则 | 判定为 |
|---|---|---|
| R1 | 空段 / 纯图片段 / 域代码段（目录、页码） | skip |
| R2 | 在表格内 → 表格首行且加粗 → 表头；其余 | tableText |
| R3 | 已挂 Heading 1/2/3 样式 | 对应 heading 级 |
| R4 | 全文第一个非空段 且 字号 ≥ 全文正文聚类字号 +4pt 且 加粗 | title |
| R5 | 匹配 `^第X章/第X部分/模块[一二三四五六七八九十]` 且 长度 ≤ 40 | heading1 |
| R6 | 匹配 `^\d+\.\d+\.\d+\s` 且 长度 ≤ 40 | heading3 |
| R7 | 匹配 `^\d+\.\d+\s` 且 长度 ≤ 40 | heading2 |
| R8 | 匹配图注模式 | caption |
| R9 | 带 Word 自动编号/项目符号（ListFormat 非空） | listItem |
| R10 | 独立成行 且 数学字符密度 > 0.7 且 含 `=` | formulaLine |
| R11 | **格式聚类兜底**：对未命中段落按（字号, 加粗, 颜色, 对齐）签名聚类；占比最大的簇 = body；比 body 字号大 ≥1.5pt 且加粗的簇，按字号降序依次补位 heading1/2/3（每级限一簇） | body / heading |
| R12 | 其余 | body |

### 4.3 置信度与 LLM 增强（可选）

- R1–R10 命中 = 高置信；R11/R12 = 低置信
- 设置中开启 LLM 后，仅把**低置信段落**的（序号, 前 50 字, 特征签名）批量发给模型，一次请求返回角色判定；失败/超时自动回落到规则结果
- LLM 配置项：BaseURL、API Key、模型名，存 `%APPDATA%\TypeRight\settings.json`，key 不回传任何服务器

---

## 5. 排版执行器

### 5.1 执行流程

```
扫描全文 → 角色判定 → 生成 PreviewModel（段落数统计 + 抽样展示）
  → 用户确认（可手动改个别段落角色）
  → UndoRecord.StartCustomRecord("智能排版")
  → 逐角色批量应用：
      a. 更新/创建对应 Word 样式定义（中西文字体、字号、颜色、间距）
      b. 段落挂载样式；清除与样式冲突的直接格式（只清字体类，保留局部加粗/颜色强调）
      c. digitsToLatinFont：正文类角色中西文/数字强制 latinFont
      d. convertPlainText：公式转换（见 §6）
  → EndCustomRecord()
  → 输出改动统计（标题 X 段 / 正文 Y 段 / 公式 Z 处 / 跳过 N 段）
```

### 5.2 硬性要求

1. **单次撤销**：StartCustomRecord/EndCustomRecord 包裹全部改动，Ctrl+Z 一次还原
2. **保留强调**：段落内局部加粗/变色（如加粗的关键词）不得被抹掉——只统一字体族与角色级属性
3. **幂等**：同一方案重复应用两次，第二次应 0 改动
4. **性能**：`Application.ScreenUpdating = false`；先收集 Range 引用再批量写；100 页文档 < 10s
5. **选区模式**：有选区时只处理选区内段落，UI 提供切换

---

## 6. 公式转换器（纯文本 → OMath）

### 6.1 识别目标

正文中的内联/独立公式，特征：字母 + 数学符号混合片段（`ax² + bx + c = 0`、`x₁ + x₂ = -b/a`、`√Δ / |a|`）。

### 6.2 转换策略（务必用线性输入法，不要手拼 OMML）

1. 用 §4 的数学片段切分规则（与角色引擎共用）在 Run 级切出公式片段
2. 片段文本 → UnicodeMath 线性串的映射：
   - Unicode 上下标字符 → `_`/`^` 语法（`x₁²` → `x_1^2`）
   - `√x` / `√(...)` / `√[...]` → `√(…)`（UnicodeMath 原生支持 `√`）
   - `a/b` 保持斜杠，BuildUp 自动堆叠
3. 将片段 Range 文本替换进 `OMaths.Add()` 新建的 OMath，`OMath.BuildUp` 转专业格式
4. 转换失败的片段**原样保留为文本**并计入统计，绝不留半转换残骸

### 6.2.1 兜底管线：MathML → OMML

`BuildUp` 无法处理的复杂片段（嵌套分式、矩阵、多层根号），走 Office 自带的 **`MML2OMML.XSL`**（位于 Office 安装目录，微软官方 XSLT）：
线性文本 → 生成 MathML → XSLT 转换 → OMML 注入。零第三方依赖；开源项目 formulyword 已验证该管线。实现时封装为 `FormulaConverter` 的 fallback 分支。

### 6.3 切分规则（与已验证实现一致的要点）

- 数学字符集：A-Za-z0-9、上下标字符、`√ Δ π ± · × ÷ = ≠ < > ≤ ≥ / | ( ) [ ] , . % + - * _ ⟹ ⟺ ↔ ∵ ∴ △`、空格
- 片段限定：必须含至少一个字母或数学专用符号，纯数字/百分比/小节号（如"6.1"）不转
- 排除白名单：`vs` 等英文缩略词
- 中文标点（，。（）：）天然作为片段边界

---

## 7. 交互设计

### 7.1 Ribbon

```
[智能排版] 选项卡
├── [一键智能排版]   大按钮：弹 PreviewPane
├── [排版选区]       仅处理选中部分
├── [方案管理]       ProfileEditor / 导入导出
└── [设置]           LLM key、混排规则开关
```

### 7.2 PreviewPane（任务窗格，WPF）

```
┌─ 智能排版预览 ─────────────────┐
│ 方案：[学术讲义/论文 ▾]        │
│ 识别结果：                      │
│  主标题 1 段 / 一级标题 7 段    │
│  二级标题 16 段 / 正文 76 段   │
│  图注 6 段 / 表格文字 11 表    │
│  公式片段 255 处（将转公式对象）│
│ ───────────────────────────── │
│ 逐段清单（可改角色）：          │
│  ¶3  "模块一：…"  [一级标题 ▾] │
│  ¶4  "1.1 一元…"  [二级标题 ▾] │
│  …（虚拟化列表，仅低置信段默认展开）│
│ ───────────────────────────── │
│  [✓] 同时挂载 Word 样式        │
│        [应用排版]  [取消]       │
└────────────────────────────────┘
```

- 低置信段落高亮标记，用户可在下拉框改角色
- 「同时挂载 Word 样式」默认勾选（治本选项）

---

## 8. 内置首发方案：学术讲义/论文

参数即 §3 示例 JSON。要点：
- 标题级：黑体系（中文）+ TNR（西文），一级深蓝、二级中蓝
- 正文：宋体 + TNR，小四（12pt）或五号（10.5pt）由方案定，1.5 倍行距，首行缩进 2 字符
- 公式：Cambria Math
- **公文模式（后续方案）**：不自行对照 GB/T 9704 抠参数，直接基于 **xkonglong/gw（小恐龙公文排版助手）** 源码适配——该项目 2026-10-21 起 MIT 开源，页面网格、仿宋、红头红线、断句重排均为现成实现
- 商务报告方案以同 Schema 扩展，不在首发范围

---

## 9. 非功能需求与验收标准

| 项 | 标准 |
|---|---|
| 兼容 | Windows + Word 2016 / 2019 / 2021 / Microsoft 365 桌面版 |
| 性能 | 100 页 / 300 公式文档，扫描+重排 < 10s |
| 稳定 | 转换公式失败率 < 2% 且失败片段无损保留原文 |
| 撤销 | 任意一次排版后单次 Ctrl+Z 完全还原 |
| 幂等 | 同方案二次应用 0 改动 |
| 离线 | 关闭 LLM 时全流程无网络请求 |
| 验收用例 | ① 纯直接格式堆叠的讲义 docx → 一键后排版合规且公式全转对象；② 已挂样式的论文 → 不破坏既有样式结构；③ 含 50+ 表格的报告 → 表格结构与图片零损坏 |

---

## 10. 开发里程碑

| 里程碑 | 内容 | 验收 |
|---|---|---|
| M1 骨架 | VSTO 工程 + Ribbon + Profile 读写 + FormatApplier（手动指定角色刷格式 + 撤销单元） | 能对选区按方案刷字体，Ctrl+Z 一次还原 |
| M2 识别引擎 | 特征提取 + R1–R12 规则表 + PreviewPane | 在 3 篇真实讲义/论文上角色识别准确率 ≥ 95% |
| M3 公式转换 | 切分 + UnicodeMath 映射 + BuildUp + 失败回退 | 测试文档 200+ 公式转换成功率 ≥ 98% |
| M4 样式挂载 + 方案编辑器 | 挂 Heading 样式、样式定义写入、方案 GUI 编辑与导入导出 | 排版后导航窗格出现完整大纲 |
| M5 LLM 增强 + 打磨 | 低置信段落 LLM 判定、设置页、安装包 | 弱格式文档识别率显著提升；离线模式全绿 |

---

## 11. 明确不做（首发范围外）

- Mac / WPS 支持（VSTO 天然不支持）
- PPT / Excel 排版
- 云端同步方案库
- 多人协作场景

---

## 12. 参考资料与开源复用清单

### 12.1 调研结论

现有同类工具全部为「模板创建 + 按钮刷格式」形态，**没有**任何产品实现本 PRD 的三个差异化能力：格式聚类智能识别、中西文字体按角色分离、纯文本公式批量转公式对象。项目不构成重复造轮子。

### 12.2 同类工具（可读源码 / 可参考）

| 项目 | 地址 | 许可证 | 参考点 |
|---|---|---|---|
| 小恐龙公文排版助手 | `github.com/xkonglong/gw` | MIT（2026-10-21 开源） | 公文 GB/T 9704 全套实现：页面网格、仿宋、红头红线、断句重排；支持 Word/WPS。**公文模式的直接源码底座** |
| word-all-in-one（办公助手） | `github.com/lun9090/word-all-in-one` | Apache-2.0 | 与本项目技术栈完全一致（VSTO+C#+Interop）：Ribbon 结构、COM 对象释放（try-finally + ReleaseComObject）、批量表格、性能优化写法。**工程脚手架参考** |
| Word快速排版 | `gitee.com/usdoc/doczs` | Apache-2.0 | 公文+论文一键排版的产品形态参考（代码质量存疑，仅看思路） |
| 论文抽屉 TAD | 商业闭源 | — | 不看代码；其功能清单（题注、交叉引用、GB/T 7714 参考文献）作为论文场景 P2 需求库 |
| FDscend_word_addin | `github.com/FDscend/fdscend_word_addin` | 开源 | 公式编号实现：书签 + 交叉引用 |
| Equation_and_Codebox | `github.com/LittleNewton/Equation_and_Codebox` | 开源 | 带编号公式插入的 VSTO 实现 |

### 12.3 可复用库（按模块映射）

| PRD 模块 | 库 / 资源 | 许可证 | 用法 |
|---|---|---|---|
| Office 互操作 | NetOfficeFw（`github.com/NetOfficeFw/NetOffice`） | MIT | 替代 PIA，免 VSTO 运行时/免注册，单 dll 分发 |
| 公式转换兜底 | MML2OMML.XSL（Office 自带） | 随 Office | MathML→OMML，见 §6.2.1；管线参考 formulyword 项目 |
| 中西文混排规则 | huacnlee/autocorrect、vinta/pangu.js | MIT / MIT | 规则集移植为 C# 正则；或调 autocorrect CLI，见 §3 |
| 方案 JSON | Newtonsoft.Json | MIT | 序列化 |
| 多文档批量（P2） | Open XML SDK | MIT（微软官方） | 不开 Word 进程批量改 docx |

### 12.4 开发动作要求

- 每个模块开工前先读对应参考仓库源码（§12.2 表格「参考点」列），禁止从零实现已有成熟写法的部分
- 引入任何第三方依赖前核对上表许可证（均为 MIT/Apache-2.0，可商用可修改，保留版权声明即可）
